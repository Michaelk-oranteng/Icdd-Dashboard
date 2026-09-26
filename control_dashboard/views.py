# control_dashboard/views.py

from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import login as auth_login, logout as auth_logout
from django.contrib.auth.models import User as DjangoUser
from django.http import JsonResponse, HttpResponse
from django.contrib import messages
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods, require_POST
from django.db import transaction
from django.db.models import Q, Count, Sum
from django.db.models.functions import Coalesce
from django.contrib.auth.decorators import login_required
import json
from datetime import datetime, timedelta, date
import base64
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from django.utils import timezone
import re
import logging

from .models import (
    UserProfile,
    Report,
    Checklist,
    ChecklistTask,
    ChecklistLog,
    ReportSubmission,
    ActivityLog,
    AdHocDeduction,
    Branch,
    Department,
    ReportDataField,
    ReportSubmissionField,
    ReportSchedule,
    ExceptionRecord,
)
from .forms import UserProfileForm

logger = logging.getLogger(__name__)


# ==================== CONSTANTS ====================

TRIAL_BALANCE_REPORT_TYPE = '__TRIAL_BALANCE__'
UPLOADED_STATUS = 'uploaded'


# ==================== HELPER FUNCTIONS ====================

def log_activity(user, activity_type, details, request=None):
    """Helper function to log user activities."""
    try:
        ip_address = request.META.get('REMOTE_ADDR') if request else None
        user_agent = request.META.get('HTTP_USER_AGENT', '') if request else ''
        ActivityLog.objects.create(
            user=user,
            activity_type=activity_type,
            details=details,
            ip_address=ip_address,
            user_agent=user_agent
        )
    except Exception as exc:
        logger.warning("log_activity failed: %s", exc)


def get_week_start(d):
    """Friday-based week start. Single source of truth."""
    weekday = d.weekday()
    if weekday >= 4:
        days_since_friday = weekday - 4
    else:
        days_since_friday = weekday + 3
    return d - timedelta(days=days_since_friday)


def count_expected_occurrences(checklist, period_start, period_end):
    """Number of expected occurrences for a checklist in a period."""
    freq = checklist.frequency
    if freq == 'daily':
        return checklist.get_expected_occurrences(period_start, period_end)
    if freq == 'weekly':
        days_diff = (period_end - period_start).days + 1
        return max(1, days_diff // 7)
    if freq in ('monthly', 'quarterly', 'bi-annual', 'annual', 'one-off'):
        count = 0
        m = period_start.month
        y = period_start.year
        while True:
            m_start = date(y, m, 1)
            if m == 12:
                m_end = date(y + 1, 1, 1) - timedelta(days=1)
            else:
                m_end = date(y, m + 1, 1) - timedelta(days=1)
            if m_start > period_end:
                break
            if m_start <= period_end and m_end >= period_start:
                count += 1
            if m == 12:
                m = 1
                y += 1
            else:
                m += 1
        return count
    return checklist.get_expected_occurrences(period_start, period_end)


def redirect_dashboard(user):
    """Redirect user to their dashboard based on role."""
    try:
        user_profile = UserProfile.objects.get(email=user.email)
        redirect_urls = {
            'admin': '/adminboard/admin/',
            'supervisor': '/adminboard/supervisor/',
            'member': '/adminboard/member/'
        }
        return redirect(redirect_urls.get(user_profile.role, '/adminboard/member/'))
    except UserProfile.DoesNotExist:
        return redirect('/adminboard/member/')


def get_or_create_django_user(email, username=None, full_name=None):
    """Get or create a Django User from email."""
    try:
        return DjangoUser.objects.get(email=email)
    except DjangoUser.DoesNotExist:
        if not username:
            username = email.split('@')[0]
            username = re.sub(r'[^a-zA-Z0-9._]', '', username).lower()

            original_username = username
            counter = 1
            while DjangoUser.objects.filter(username=username).exists():
                username = f"{original_username}{counter}"
                counter += 1

        django_user = DjangoUser.objects.create_user(username=username, email=email)
        django_user.set_unusable_password()
        django_user.save()

        if full_name:
            name_parts = full_name.split(' ', 1)
            django_user.first_name = name_parts[0]
            django_user.last_name = name_parts[1] if len(name_parts) > 1 else ''
            django_user.save()

        return django_user


def _add_frequency(dt, frequency):
    """
    Add exactly one interval of `frequency` to a timezone-aware datetime.

    Rules:
      - daily     : +1 calendar day, then skip Sat/Sun
      - weekly    : +7 days
      - monthly   : +1 month, same day, CLAMPED to last day if the day
                    doesn't exist in the target month (31 Jan → 28 Feb)
      - quarterly : +3 months, same-day-clamped
      - yearly    : +1 year, same-day-clamped (leap-safe)
      - one-off   : returns None (terminal — no next occurrence)
    """
    if frequency == 'one-off':
        return None

    if frequency == 'daily':
        nxt = dt + timedelta(days=1)
        while nxt.weekday() >= 5:  # 5 = Saturday, 6 = Sunday
            nxt += timedelta(days=1)
        return nxt

    if frequency == 'weekly':
        return dt + timedelta(days=7)

    if frequency == 'monthly':
        months_to_add = 1
    elif frequency == 'quarterly':
        months_to_add = 3
    elif frequency == 'yearly':
        months_to_add = 12
    else:
        return dt + timedelta(days=7)

    target_month_index = (dt.month - 1) + months_to_add
    target_year = dt.year + (target_month_index // 12)
    target_month = (target_month_index % 12) + 1

    if target_month == 12:
        last_day_of_target = (date(target_year + 1, 1, 1) - timedelta(days=1)).day
    else:
        last_day_of_target = (date(target_year, target_month + 1, 1) - timedelta(days=1)).day

    target_day = min(dt.day, last_day_of_target)

    return dt.replace(year=target_year, month=target_month, day=target_day)

def count_expected_report_occurrences(report, until=None):
    """
    Count how many recurrence periods a Report has produced from its
    anchor (Report.deadline_date) up to `until`.

    A period only counts once its deadline has passed. If the deadline
    is still in the future (relative to `until`), that period does not
    yet count — it hasn't been "expected" yet.

    A missed period still counts in the denominator (team performance
    penalises it by contributing 0 to the numerator). This function
    simply counts periods; the caller is responsible for how a missed
    period is weighted.
    """
    if until is None:
        until = timezone.now()

    if timezone.is_naive(until):
        until = timezone.make_aware(until)

    if not report.deadline_date:
        return 0

    anchor_time = report.deadline_time or datetime.strptime('23:59', '%H:%M').time()
    anchor_dt = timezone.make_aware(datetime.combine(report.deadline_date, anchor_time))

    frequency = report.frequency or 'one-off'

    if frequency == 'one-off':
        return 1 if anchor_dt <= until else 0

    count = 0
    period_deadline = anchor_dt
    safety = 0
    while safety < 10000:
        if period_deadline > until:
            break
        count += 1
        nxt = _add_frequency(period_deadline, frequency)
        if nxt is None:
            break
        period_deadline = nxt
        safety += 1

    return count


# ==================== LOGIN VIEWS ====================

def landing_page(request):
    if request.user.is_authenticated:
        return redirect_dashboard(request.user)
    return render(request, 'control_dashboard/index.html')


def logout_view(request):
    if request.user.is_authenticated:
        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            log_activity(
                user=user_profile,
                activity_type='logout',
                details=f'User {request.user.email} logged out',
                request=request
            )
        except UserProfile.DoesNotExist:
            pass

    auth_logout(request)
    messages.info(request, 'You have been logged out successfully.')
    return redirect('control_dashboard:landing_page')


# ==================== API - AUTHENTICATION ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_email_login(request):
    try:
        data = json.loads(request.body)
        login_input = data.get('email', '').strip()
        raw_input = data.get('raw_input', '').strip()

        if not login_input:
            return JsonResponse({'success': False, 'error': 'Username or email is required'}, status=400)

        logger.info("Login attempt for '%s'", login_input)

        user_profile = None

        try:
            user_profile = UserProfile.objects.get(email__iexact=login_input)
        except UserProfile.DoesNotExist:
            pass

        if not user_profile:
            username_to_check = login_input.split('@')[0] if '@' in login_input else login_input
            try:
                user_profile = UserProfile.objects.get(username__iexact=username_to_check)
            except UserProfile.DoesNotExist:
                pass

        if not user_profile and '@' not in login_input:
            for domain in ['@omnisbic.com.gh', '@omnisbic.com']:
                try:
                    user_profile = UserProfile.objects.get(email__iexact=login_input + domain)
                    break
                except UserProfile.DoesNotExist:
                    continue

        if not user_profile:
            return JsonResponse({
                'success': False,
                'error': f'No account found for "{raw_input or login_input}". Please check your username or contact your administrator.'
            }, status=404)

        if user_profile.status != 'active':
            return JsonResponse({'success': False, 'error': 'Your account is inactive. Please contact support.'}, status=403)

        django_user = None

        try:
            django_user = DjangoUser.objects.get(email=user_profile.email)
        except DjangoUser.DoesNotExist:
            pass

        if not django_user and user_profile.username:
            try:
                django_user = DjangoUser.objects.get(username=user_profile.username)
            except DjangoUser.DoesNotExist:
                pass

        if not django_user:
            django_username = user_profile.username or user_profile.email.split('@')[0]
            original_username = django_username
            counter = 1
            while DjangoUser.objects.filter(username=django_username).exists():
                django_username = f"{original_username}{counter}"
                counter += 1

            django_user = DjangoUser.objects.create_user(username=django_username, email=user_profile.email)
            django_user.set_unusable_password()
            django_user.save()

            if user_profile.full_name:
                name_parts = user_profile.full_name.split(' ', 1)
                django_user.first_name = name_parts[0]
                django_user.last_name = name_parts[1] if len(name_parts) > 1 else ''
                django_user.save()

        if django_user.email != user_profile.email:
            django_user.email = user_profile.email
            django_user.save()

        if user_profile.username and django_user.username != user_profile.username:
            if not DjangoUser.objects.filter(username=user_profile.username).exclude(id=django_user.id).exists():
                django_user.username = user_profile.username
                django_user.save()

        auth_login(request, django_user)

        log_activity(
            user=user_profile,
            activity_type='login',
            details=f'User {user_profile.email} logged in via {raw_input or login_input}',
            request=request
        )

        redirect_urls = {
            'admin': '/adminboard/admin/',
            'supervisor': '/adminboard/supervisor/',
            'member': '/adminboard/member/'
        }

        return JsonResponse({
            'success': True,
            'message': f'Welcome back, {user_profile.full_name}!',
            'redirect_url': redirect_urls.get(user_profile.role, '/adminboard/member/'),
            'user': {
                'id': user_profile.id,
                'email': user_profile.email,
                'username': user_profile.username,
                'full_name': user_profile.full_name,
                'role': user_profile.role,
                'position': user_profile.position
            }
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Login error")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def get_user_profile_api(request):
    if not request.user.is_authenticated:
        return JsonResponse({'authenticated': False}, status=401)

    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        return JsonResponse({
            'authenticated': True,
            'user': {
                'id': user_profile.id,
                'email': user_profile.email,
                'username': user_profile.username,
                'full_name': user_profile.full_name,
                'role': user_profile.role,
                'position': user_profile.position,
                'status': user_profile.status
            }
        })
    except UserProfile.DoesNotExist:
        return JsonResponse({
            'authenticated': True,
            'user': {
                'email': request.user.email,
                'username': request.user.username,
                'full_name': request.user.get_full_name() or request.user.username,
                'role': 'member',
                'position': 'member'
            }
        })


# ==================== ADMIN VIEWS ====================

@login_required
def admin_page(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    users = UserProfile.objects.all().order_by('full_name')
    positions = UserProfile.POSITION_CHOICES
    roles = UserProfile.ROLE_CHOICES
    statuses = UserProfile.STATUS_CHOICES
    branches = Branch.objects.filter(is_active=True).order_by('name')
    departments = Department.objects.filter(is_active=True).order_by('name')

    context = {
        'user_profile': user_profile,
        'users': users,
        'positions': positions,
        'roles': roles,
        'statuses': statuses,
        'branches': branches,
        'departments': departments,
    }

    return render(request, 'control_dashboard/adminboard.html', context)


# ==================== API - ADMIN USER MANAGEMENT ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_create_user(request):
    try:
        data = json.loads(request.body)
        email = data.get('email', '').strip()
        full_name = data.get('full_name', '').strip()
        position = data.get('position', 'member')
        role = data.get('role', 'member')
        status = data.get('status', 'active')
        department_id = data.get('department_id')
        branch_id = data.get('branch_id')

        if not email:
            return JsonResponse({'success': False, 'error': 'Email is required'}, status=400)
        if not full_name:
            return JsonResponse({'success': False, 'error': 'Full name is required'}, status=400)
        if UserProfile.objects.filter(email__iexact=email).exists():
            return JsonResponse({'success': False, 'error': 'A user with this email already exists'}, status=400)

        user = UserProfile.objects.create(
            email=email.lower(), full_name=full_name,
            position=position, role=role, status=status
        )

        if department_id:
            try:
                user.departments.add(Department.objects.get(id=department_id, is_active=True))
            except Department.DoesNotExist:
                pass

        if branch_id:
            try:
                user.branches.add(Branch.objects.get(id=branch_id, is_active=True))
            except Branch.DoesNotExist:
                pass

        user.save()

        get_or_create_django_user(email=email.lower(), username=user.username, full_name=full_name)

        log_activity(
            user=user,
            activity_type='user_created',
            details=f'User {email} was created with role {user.get_role_display()} (username: {user.username})',
            request=request
        )

        return JsonResponse({
            'success': True,
            'message': f'User created successfully. Username: {user.username}',
            'user_id': user.id,
            'username': user.username
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["PUT", "POST"])
def api_edit_user(request, user_id):
    try:
        user = get_object_or_404(UserProfile, id=user_id)
        data = json.loads(request.body)

        changes = []
        old_email = user.email
        old_username = user.username

        if 'email' in data:
            new_email = data['email'].strip().lower()
            if new_email and new_email != user.email:
                if UserProfile.objects.filter(email=new_email).exclude(id=user_id).exists():
                    return JsonResponse({'success': False, 'error': 'Email already in use'}, status=400)
                changes.append(f'Email changed from {user.email} to {new_email}')
                user.email = new_email

        if 'full_name' in data and data['full_name'].strip() != user.full_name:
            changes.append(f'Full name changed to {data["full_name"]}')
            user.full_name = data['full_name'].strip()

        if 'username' in data and data['username'].strip():
            new_username = data['username'].strip()
            if new_username != user.username:
                if UserProfile.objects.filter(username=new_username).exclude(id=user_id).exists():
                    return JsonResponse({'success': False, 'error': 'Username already in use'}, status=400)
                changes.append(f'Username changed from {user.username} to {new_username}')
                user.username = new_username

        if 'position' in data and data['position'] != user.position:
            old_pos = user.get_position_display()
            user.position = data['position']
            changes.append(f'Position changed from {old_pos} to {user.get_position_display()}')

        if 'role' in data and data['role'] != user.role:
            old_role = user.get_role_display()
            user.role = data['role']
            changes.append(f'Role changed from {old_role} to {user.get_role_display()}')

        if 'status' in data and data['status'] != user.status:
            old_status = user.get_status_display()
            user.status = data['status']
            changes.append(f'Status changed from {old_status} to {user.get_status_display()}')

        user.save()

        if 'department_ids' in data:
            department_ids = data['department_ids']
            if department_ids:
                user.departments.clear()
                departments = Department.objects.filter(id__in=department_ids, is_active=True)
                user.departments.add(*departments)
                changes.append(f'Departments updated to {len(departments)} departments')
            else:
                user.departments.clear()
                changes.append('Departments cleared')

        if 'branch_ids' in data:
            branch_ids = data['branch_ids']
            if branch_ids:
                user.branches.clear()
                branches = Branch.objects.filter(id__in=branch_ids, is_active=True)
                user.branches.add(*branches)
                changes.append(f'Branches updated to {len(branches)} branches')
            else:
                user.branches.clear()
                changes.append('Branches cleared')

        if 'department_id' in data and 'department_ids' not in data:
            department_id = data.get('department_id')
            if department_id:
                try:
                    department = Department.objects.get(id=department_id, is_active=True)
                    user.departments.clear()
                    user.departments.add(department)
                    changes.append(f'Department set to {department.name}')
                except Department.DoesNotExist:
                    pass
            else:
                user.departments.clear()
                changes.append('Departments cleared')

        if 'branch_id' in data and 'branch_ids' not in data:
            branch_id = data.get('branch_id')
            if branch_id:
                try:
                    branch = Branch.objects.get(id=branch_id, is_active=True)
                    user.branches.clear()
                    user.branches.add(branch)
                    changes.append(f'Branch set to {branch.name}')
                except Branch.DoesNotExist:
                    pass
            else:
                user.branches.clear()
                changes.append('Branches cleared')

        django_user = None
        if old_email:
            django_user = DjangoUser.objects.filter(email=old_email).first()
        if not django_user and old_username:
            django_user = DjangoUser.objects.filter(username=old_username).first()

        if django_user:
            if user.username and user.username != django_user.username:
                if not DjangoUser.objects.filter(username=user.username).exclude(id=django_user.id).exists():
                    django_user.username = user.username
            if user.email and user.email != django_user.email:
                django_user.email = user.email
            if user.full_name:
                name_parts = user.full_name.split(' ', 1)
                django_user.first_name = name_parts[0]
                django_user.last_name = name_parts[1] if len(name_parts) > 1 else ''
            django_user.save()
        else:
            get_or_create_django_user(email=user.email, username=user.username, full_name=user.full_name)

        if changes:
            log_activity(
                user=user,
                activity_type='user_updated',
                details=f'User {user.email} updated: ' + '; '.join(changes),
                request=request
            )

        return JsonResponse({'success': True, 'message': 'User updated successfully'})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error in api_edit_user")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["POST"])
def api_update_status(request, user_id):
    try:
        user = get_object_or_404(UserProfile, id=user_id)
        data = json.loads(request.body)
        new_status = data.get('status')

        if new_status not in ['active', 'inactive']:
            return JsonResponse({'success': False, 'error': 'Invalid status value'}, status=400)

        old_status = user.get_status_display()
        user.status = new_status
        user.save()

        log_activity(
            user=user,
            activity_type='user_updated',
            details=f'User {user.email} status changed from {old_status} to {user.get_status_display()}',
            request=request
        )

        return JsonResponse({'success': True, 'message': f'User status updated to {new_status}'})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_user(request, user_id):
    try:
        user = get_object_or_404(UserProfile, id=user_id)
        email = user.email
        username = user.username

        matched = DjangoUser.objects.filter(email=email)
        if not matched.exists() and username:
            matched = DjangoUser.objects.filter(username=username)
        for du in matched:
            du.delete()

        log_activity(
            user=user,
            activity_type='user_deleted',
            details=f'User {email} (username: {username}) was deleted',
            request=request
        )

        user.delete()

        return JsonResponse({'success': True, 'message': 'User deleted successfully'})

    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ==================== SUPERVISOR VIEWS ====================

@login_required
def supervisor_dashboard(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    today = timezone.now().date()

    # ============================================================
    # TOTAL EXCEPTIONS — fed from ExceptionRecord (matches Analytics)
    # ============================================================
    exceptions_qs = ExceptionRecord.objects.filter(
        report__isnull=False,
    ).exclude(
        report__report_type=TRIAL_BALANCE_REPORT_TYPE,
    )

    total_exceptions = exceptions_qs.count()
    today_exceptions = exceptions_qs.filter(created_at__date=today).count()

    # ============================================================
    # SUBMITTED REPORTS — still counts Report containers
    # (a "submitted report" is a container-level concept)
    # ============================================================
    submitted_qs = Report.objects.filter(status='submitted').exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    )
    submitted_reports_count = submitted_qs.count()

    # ============================================================
    # TEAM PERFORMANCE — unchanged logic
    # ============================================================
    team_members = UserProfile.objects.filter(role='member', status='active').order_by('full_name')

    max_submissions = (
        submitted_qs.values('created_by').annotate(c=Count('id'))
        .order_by('-c').values_list('c', flat=True).first()
    ) or 1

    team_performance = []
    for member in team_members:
        member_submitted = submitted_qs.filter(created_by=member).count()
        member_deductions = (
            AdHocDeduction.objects.filter(user=member)
            .aggregate(total=Sum('points')).get('total') or 0
        )
        final_score = max(0, member_submitted - member_deductions)
        percentage = int((final_score / max_submissions) * 100) if max_submissions > 0 else 0
        percentage = min(percentage, 100)

        if percentage >= 80:
            status = 'success'
        elif percentage >= 50:
            status = 'warning'
        else:
            status = 'danger'

        team_performance.append({
            'user': member,
            'submitted': member_submitted,
            'deductions': member_deductions,
            'final_score': final_score,
            'percentage': percentage,
            'status': status,
        })

    team_performance.sort(key=lambda x: x['percentage'], reverse=True)

    if team_performance:
        completion_rate = int(sum(m['percentage'] for m in team_performance) / len(team_performance))
    else:
        completion_rate = 0

    context = {
        'user_profile': user_profile,
        'today': timezone.now(),
        'total_exceptions': total_exceptions,
        'today_exceptions': today_exceptions,
        'submitted_reports_count': submitted_reports_count,
        'team_performance': team_performance,
        'completion_rate': completion_rate,
    }

    return render(request, 'control_dashboard/supervisorboard.html', context)

# ==================== REPORT CREATION & CENTER VIEWS ====================

@login_required
def report_creation(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    users = UserProfile.objects.filter(role='member', status='active').order_by('full_name')

    context = {
        'user_profile': user_profile,
        'users': users,
    }

    return render(request, 'control_dashboard/reportcreation.html', context)


@login_required
def report_center(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    reports = Report.objects.exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    ).exclude(
        status='submitted'
    ).order_by('-created_at')

    user_filter = request.GET.get('user', '')
    if user_filter and user_filter != 'all':
        reports = reports.filter(assigned_to__id=user_filter)

    users = UserProfile.objects.filter(role='member', status='active').order_by('full_name')

    context = {
        'user_profile': user_profile,
        'reports': reports,
        'users': users,
        'user_filter': user_filter,
    }

    return render(request, 'control_dashboard/reportcenter.html', context)


# ==================== API - REPORT MANAGEMENT ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_create_report(request):
    try:
        if not request.body:
            return JsonResponse({'success': False, 'error': 'Empty request payload'}, status=400)

        data = json.loads(request.body)

        report_type = data.get('report_type', '').strip()
        frequency = data.get('frequency', 'one-off')
        description = data.get('description', '').strip()
        deadline_date = data.get('deadline_date')
        deadline_time = data.get('deadline_time')
        assigned_users = data.get('assigned_users', [])
        is_assigned_to_all = data.get('is_assigned_to_all', False)

        if not report_type:
            return JsonResponse({'success': False, 'error': 'Report type is required'}, status=400)

        if report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({
                'success': False,
                'error': 'This report type is reserved for internal use.'
            }, status=400)

        try:
            created_by = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if deadline_date and str(deadline_date).strip():
            try:
                deadline_date = datetime.strptime(str(deadline_date).strip(), '%Y-%m-%d').date()
            except ValueError:
                return JsonResponse({'success': False, 'error': 'Invalid deadline_date format (YYYY-MM-DD)'}, status=400)
        else:
            deadline_date = None

        # Recurring reports need an anchor date so the deadline can recur.
        if frequency != 'one-off' and deadline_date is None:
            return JsonResponse({
                'success': False,
                'error': 'Recurring reports require a deadline date (it is the anchor for recurrence).'
            }, status=400)

        if deadline_time and str(deadline_time).strip():
            try:
                deadline_time = datetime.strptime(str(deadline_time).strip(), '%H:%M').time()
            except ValueError:
                return JsonResponse({'success': False, 'error': 'Invalid deadline_time format (HH:MM)'}, status=400)
        else:
            deadline_time = None

        report = Report.objects.create(
            report_type=report_type,
            frequency=frequency,
            description=description,
            deadline_date=deadline_date,
            deadline_time=deadline_time,
            is_assigned_to_all=is_assigned_to_all,
            created_by=created_by,
            status='assigned'
        )

        if is_assigned_to_all:
            report.assigned_to.set(UserProfile.objects.filter(role='member', status='active'))
        elif assigned_users:
            report.assigned_to.set(UserProfile.objects.filter(id__in=assigned_users, status='active'))

        log_activity(
            user=created_by,
            activity_type='report_created',
            details=f'Created report: {report_type}',
            request=request
        )

        return JsonResponse({
            'success': True,
            'message': 'Report created successfully',
            'report_id': report.id
        }, status=201)

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error creating report")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def api_get_report(request, report_id):
    try:
        report = get_object_or_404(Report, id=report_id)

        if report.report_type == TRIAL_BALANCE_REPORT_TYPE or report.status == UPLOADED_STATUS:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        return JsonResponse({
            'success': True,
            'report': {
                'id': report.id,
                'report_type': report.report_type,
                'frequency': report.frequency,
                'description': report.description,
                'status': report.status,
                'deadline_date': report.deadline_date.strftime('%Y-%m-%d') if report.deadline_date else '',
                'deadline_time': report.deadline_time.strftime('%H:%M') if report.deadline_time else '',
                'assigned_users': list(report.assigned_to.values_list('id', flat=True)),
                'is_assigned_to_all': report.is_assigned_to_all,
                'data': report.get_display_data(),
            }
        })

    except Exception as e:
        logger.exception("Error getting report")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["PUT", "POST"])
def api_edit_report(request, report_id):
    try:
        report = get_object_or_404(Report, id=report_id)

        if report.report_type == TRIAL_BALANCE_REPORT_TYPE or report.status == UPLOADED_STATUS:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if not request.body:
            return JsonResponse({'success': False, 'error': 'Empty request payload'}, status=400)

        data = json.loads(request.body)

        if 'report_type' in data and data['report_type'].strip() == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({
                'success': False,
                'error': 'This report type is reserved for internal use.'
            }, status=400)

        changes = []

        if 'report_type' in data and data['report_type'].strip() != report.report_type:
            changes.append(f'Type changed from {report.report_type} to {data["report_type"]}')
            report.report_type = data['report_type'].strip()

        if 'frequency' in data and data['frequency'] != report.frequency:
            old_freq = report.get_frequency_display()
            report.frequency = data['frequency']
            changes.append(f'Frequency changed from {old_freq} to {report.get_frequency_display()}')

        if 'status' in data and data['status'] != report.status:
            old_status = report.get_status_display()
            report.status = data['status']
            changes.append(f'Status changed from {old_status} to {report.get_status_display()}')

        if 'description' in data:
            report.description = data['description'].strip()

        if 'deadline_date' in data:
            val = data['deadline_date']
            if val and str(val).strip():
                try:
                    report.deadline_date = datetime.strptime(str(val).strip(), '%Y-%m-%d').date()
                except ValueError:
                    return JsonResponse({'success': False, 'error': 'Invalid deadline_date format (YYYY-MM-DD)'}, status=400)
            else:
                report.deadline_date = None

        if 'deadline_time' in data:
            val = data['deadline_time']
            if val and str(val).strip():
                try:
                    report.deadline_time = datetime.strptime(str(val).strip(), '%H:%M').time()
                except ValueError:
                    return JsonResponse({'success': False, 'error': 'Invalid deadline_time format (HH:MM)'}, status=400)
            else:
                report.deadline_time = None

        is_assigned_to_all = data.get('is_assigned_to_all', report.is_assigned_to_all)
        assigned_users = data.get('assigned_users', [])

        if is_assigned_to_all != report.is_assigned_to_all:
            report.is_assigned_to_all = is_assigned_to_all
            changes.append('Assignment mode toggled')

        if is_assigned_to_all:
            report.assigned_to.set(UserProfile.objects.filter(role='member', status='active'))
        elif 'assigned_users' in data:
            report.assigned_to.set(UserProfile.objects.filter(id__in=assigned_users, status='active'))
            changes.append('Assigned user profiles refreshed')

        report.save()

        if changes:
            log_activity(
                user=user_profile,
                activity_type='report_updated',
                details=f'Report {report.report_type} updated: ' + '; '.join(changes),
                request=request
            )

        return JsonResponse({'success': True, 'message': 'Report updated successfully'})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error editing report")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_report(request, report_id):
    try:
        report = get_object_or_404(Report, id=report_id)

        if report.report_type == TRIAL_BALANCE_REPORT_TYPE or report.status == UPLOADED_STATUS:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        report_type = report.report_type

        log_activity(
            user=report.created_by,
            activity_type='report_deleted',
            details=f'Report "{report_type}" was deleted',
            request=request
        )

        report.delete()

        return JsonResponse({'success': True, 'message': 'Report deleted successfully'})

    except Exception as e:
        logger.exception("Error deleting report")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ==================== CHECKLIST VIEWS ====================

@login_required
def checklist_builder(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    checklists = Checklist.objects.all().order_by('-created_at')
    users = UserProfile.objects.filter(status='active').order_by('full_name')
    branches = Branch.objects.filter(is_active=True).order_by('name')
    departments = Department.objects.filter(is_active=True).order_by('name')
    frequencies = Checklist.FREQUENCY_CHOICES
    assignments = Checklist.ASSIGNMENT_CHOICES

    context = {
        'user_profile': user_profile,
        'checklists': checklists,
        'users': users,
        'branches': branches,
        'departments': departments,
        'frequencies': frequencies,
        'assignments': assignments,
    }

    return render(request, 'control_dashboard/checklist.html', context)


@login_required
def checklist_list(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    checklists = Checklist.objects.all().prefetch_related('tasks', 'assigned_users').order_by('-created_at')
    users = UserProfile.objects.filter(status='active').order_by('full_name')

    context = {
        'user_profile': user_profile,
        'checklists': checklists,
        'users': users,
        'today': timezone.now(),
    }

    return render(request, 'control_dashboard/checklist_list.html', context)


# ==================== API - CHECKLIST MANAGEMENT ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_create_checklist(request):
    try:
        if not request.body:
            return JsonResponse({'success': False, 'error': 'Empty request body'}, status=400)

        data = json.loads(request.body)

        name = data.get('name', '').strip()
        description = data.get('description', '').strip()
        frequency = data.get('frequency', 'weekly')
        assignment_target = data.get('assignment_target', 'all')
        assigned_users_ids = data.get('assigned_users', [])
        assigned_branches_ids = data.get('assigned_branches', [])
        assigned_departments_ids = data.get('assigned_departments', [])
        tasks_data = data.get('tasks', [])

        if not name:
            return JsonResponse({'success': False, 'error': 'Activity name is required'}, status=400)
        if not tasks_data:
            return JsonResponse({'success': False, 'error': 'Please add at least one task'}, status=400)

        created_by = None
        if request.user and request.user.is_authenticated:
            created_by = UserProfile.objects.filter(email=request.user.email).first()
            if not created_by:
                created_by = UserProfile.objects.create(
                    email=request.user.email,
                    full_name=request.user.get_full_name() or request.user.username,
                    role='member',
                    position='member',
                    status='active'
                )

        if not created_by:
            return JsonResponse({'success': False, 'error': 'No active UserProfile records found to assign ownership'}, status=400)

        checklist = Checklist.objects.create(
            name=name,
            description=description,
            frequency=frequency,
            assignment_target=assignment_target,
            created_by=created_by,
            is_active=True
        )

        if assignment_target == 'specific' and assigned_users_ids:
            checklist.assigned_users.set(UserProfile.objects.filter(id__in=assigned_users_ids, status='active'))
        elif assignment_target == 'cc':
            checklist.assigned_users.set(UserProfile.objects.filter(position='cc', status='active'))
        elif assignment_target == 'hc':
            checklist.assigned_users.set(UserProfile.objects.filter(position='hc', status='active'))
        elif assignment_target == 'all':
            checklist.assigned_users.set(UserProfile.objects.filter(status='active'))

        if assigned_branches_ids:
            checklist.assigned_branches.set(Branch.objects.filter(id__in=assigned_branches_ids, is_active=True))

        if assigned_departments_ids:
            checklist.assigned_departments.set(Department.objects.filter(id__in=assigned_departments_ids, is_active=True))

        for index, task_item in enumerate(tasks_data):
            task_desc = task_item.get('description', '').strip()
            if task_desc:
                ChecklistTask.objects.create(checklist=checklist, description=task_desc, order=index)

        saved_branches = checklist.assigned_branches.all()
        saved_departments = checklist.assigned_departments.all()

        log_activity(
            user=created_by,
            activity_type='checklist_created',
            details=f'Created checklist "{name}" with {len(tasks_data)} tasks',
            request=request
        )

        return JsonResponse({
            'success': True,
            'message': 'Checklist created successfully',
            'checklist_id': checklist.id,
            'branches_saved': [b.name for b in saved_branches],
            'departments_saved': [d.name for d in saved_departments]
        }, status=201)

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Malformed or invalid JSON payload structure'}, status=400)
    except Exception as e:
        logger.exception("Error in api_create_checklist")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def api_get_checklist(request, checklist_id):
    try:
        checklist = get_object_or_404(Checklist, id=checklist_id)
        tasks = checklist.tasks.all().order_by('order')

        branch_names = [branch.name for branch in checklist.assigned_branches.all()]
        department_names = [dept.name for dept in checklist.assigned_departments.all()]

        return JsonResponse({
            'success': True,
            'checklist': {
                'id': checklist.id,
                'name': checklist.name,
                'description': checklist.description,
                'frequency': checklist.frequency,
                'assignment_target': checklist.assignment_target,
                'is_active': checklist.is_active,
                'assigned_users': list(checklist.assigned_users.values_list('id', flat=True)),
                'assigned_branches': list(checklist.assigned_branches.values_list('id', flat=True)),
                'assigned_branches_names': branch_names,
                'assigned_departments': list(checklist.assigned_departments.values_list('id', flat=True)),
                'assigned_departments_names': department_names,
                'tasks': [
                    {'id': task.id, 'description': task.description, 'order': task.order, 'is_completed': task.is_completed}
                    for task in tasks
                ]
            }
        })

    except Exception as e:
        logger.exception("Error in api_get_checklist")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["PUT"])
def api_edit_checklist(request, checklist_id):
    try:
        checklist = get_object_or_404(Checklist, id=checklist_id)
        if not request.body:
            return JsonResponse({'success': False, 'error': 'Empty patch payload parameters'}, status=400)

        data = json.loads(request.body)

        changes = []

        if 'name' in data:
            checklist.name = data['name'].strip()
            changes.append(f"Name updated to '{checklist.name}'")
        if 'frequency' in data:
            old_freq = checklist.get_frequency_display()
            checklist.frequency = data['frequency']
            changes.append(f"Frequency changed from {old_freq} to {checklist.get_frequency_display()}")
        if 'assignment_target' in data:
            old_target = checklist.get_assignment_display()
            checklist.assignment_target = data['assignment_target']
            changes.append(f"Assignment target changed from {old_target} to {checklist.get_assignment_display()}")

        checklist.save()

        assignment_target = data.get('assignment_target', checklist.assignment_target)
        if assignment_target == 'specific':
            assigned_users_ids = data.get('assigned_users', [])
            if assigned_users_ids:
                checklist.assigned_users.set(UserProfile.objects.filter(id__in=assigned_users_ids, status='active'))
                changes.append(f"Assigned {len(assigned_users_ids)} specific users")
            else:
                checklist.assigned_users.clear()
                changes.append("Cleared all user assignments")
        elif assignment_target == 'cc':
            checklist.assigned_users.set(UserProfile.objects.filter(position='cc', status='active'))
            changes.append("Assigned to Cluster Control users")
        elif assignment_target == 'hc':
            checklist.assigned_users.set(UserProfile.objects.filter(position='hc', status='active'))
            changes.append("Assigned to Head Office Control users")
        elif assignment_target == 'all':
            checklist.assigned_users.set(UserProfile.objects.filter(status='active'))
            changes.append("Assigned to all active users")

        assigned_branches_ids = data.get('assigned_branches', [])
        if assigned_branches_ids:
            branches = Branch.objects.filter(id__in=assigned_branches_ids, is_active=True)
            checklist.assigned_branches.set(branches)
            changes.append(f"Assigned to {len(branches)} branches")
        else:
            checklist.assigned_branches.clear()
            changes.append("Cleared all branch assignments")

        assigned_departments_ids = data.get('assigned_departments', [])
        if assigned_departments_ids:
            departments = Department.objects.filter(id__in=assigned_departments_ids, is_active=True)
            checklist.assigned_departments.set(departments)
            changes.append(f"Assigned to {len(departments)} departments")
        else:
            checklist.assigned_departments.clear()
            changes.append("Cleared all department assignments")

        if 'tasks' in data:
            incoming_tasks = data['tasks']
            existing_tasks = {t.id: t for t in checklist.tasks.all()}
            seen_ids = set()

            for index, task_item in enumerate(incoming_tasks):
                task_desc = task_item.get('description', '').strip()
                if not task_desc:
                    continue

                task_id = task_item.get('id')
                if task_id and task_id in existing_tasks:
                    task = existing_tasks[task_id]
                    task.description = task_desc
                    task.order = index
                    task.save()
                    seen_ids.add(task_id)
                else:
                    new_task = ChecklistTask.objects.create(checklist=checklist, description=task_desc, order=index)
                    seen_ids.add(new_task.id)

            for tid, task in existing_tasks.items():
                if tid not in seen_ids:
                    task.delete()

            changes.append(f"Updated {len(incoming_tasks)} tasks")

        if changes and request.user.is_authenticated:
            try:
                user_profile = UserProfile.objects.get(email=request.user.email)
                log_activity(
                    user=user_profile,
                    activity_type='checklist_updated',
                    details=f'Checklist "{checklist.name}" updated: ' + '; '.join(changes[:3]) + ('...' if len(changes) > 3 else ''),
                    request=request
                )
            except UserProfile.DoesNotExist:
                pass

        return JsonResponse({
            'success': True,
            'message': 'Checklist updated successfully',
            'changes': changes
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data payload parameters'}, status=400)
    except Exception as e:
        logger.exception("Error in api_edit_checklist")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_checklist(request, checklist_id):
    try:
        checklist = get_object_or_404(Checklist, id=checklist_id)
        checklist_name = checklist.name

        log_activity(
            user=checklist.created_by,
            activity_type='checklist_deleted',
            details=f'Checklist "{checklist_name}" was deleted',
            request=request
        )

        checklist.delete()

        return JsonResponse({'success': True, 'message': 'Checklist deleted successfully'})

    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ==================== MEMBER DASHBOARD VIEW ====================

# ==================== MEMBER DASHBOARD VIEW ====================

@login_required
def member_dashboard(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        user_profile = UserProfile.objects.create(
            email=request.user.email,
            full_name=request.user.get_full_name() or request.user.username,
            role='member',
            position='member',
            status='active'
        )

    today = timezone.now().date()
    week_start = get_week_start(today)
    week_end = week_start + timedelta(days=6)

    user_checklists = Checklist.objects.filter(
        is_active=True
    ).filter(
        Q(assigned_users=user_profile) |
        Q(assignment_target='all') |
        Q(assignment_target=user_profile.position)
    ).distinct()

    total_checklist_rows = 0
    frequency_counts = {}

    user_branch_ids = set(user_profile.branches.values_list('id', flat=True))
    user_dept_ids = set(user_profile.departments.values_list('id', flat=True))

    for checklist in user_checklists:
        branches = set(checklist.assigned_branches.values_list('id', flat=True))
        departments = set(checklist.assigned_departments.values_list('id', flat=True))

        visible_branches = branches & user_branch_ids
        visible_departments = departments & user_dept_ids

        if not visible_branches and not visible_departments:
            if not branches and not departments:
                total_checklist_rows += 1
        else:
            total_checklist_rows += len(visible_branches) + len(visible_departments)

        freq = checklist.frequency
        frequency_counts[freq] = frequency_counts.get(freq, 0) + 1

    total_tasks_completed = ChecklistLog.objects.filter(user=user_profile).count()

    month_start = today.replace(day=1)
    total_checklists_completed = 0

    user_reports = Report.objects.filter(created_by=user_profile).exclude(report_type=TRIAL_BALANCE_REPORT_TYPE)

    excel_entry_count = ExceptionRecord.objects.filter(
        report__created_by=user_profile,
    ).exclude(
        report__report_type=TRIAL_BALANCE_REPORT_TYPE,
    ).count()

    # Form-based reports (no Excel records) still count as captured exceptions
    non_excel_reports = user_reports.filter(
        exception_records__isnull=True,
        excel_headers='',
    ).count()

    total_exceptions_captured = excel_entry_count + non_excel_reports

    exceptions_this_week = Report.objects.filter(
        created_by=user_profile,
        created_at__date__gte=week_start,
        created_at__date__lte=week_end
    ).exclude(report_type=TRIAL_BALANCE_REPORT_TYPE).count()

    pending_qs = Report.objects.filter(
        Q(assigned_to=user_profile) | Q(is_assigned_to_all=True)
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).filter(
        created_at__date__gte=week_start,
        created_at__date__lte=week_end
    ).exclude(
        status__in=['completed', 'approved']
    ).distinct()

    pending_submissions = pending_qs.count()

    month_expected = 0
    month_actual = 0
    year_expected = 0
    year_actual = 0

    year_start = today.replace(month=1, day=1)
    year_end = today.replace(month=12, day=31)

    display_checklists_with_progress = []

    for checklist in user_checklists:
        branches = list(checklist.assigned_branches.all())
        departments = list(checklist.assigned_departments.all())

        visible_units = []
        for b in branches:
            if b.id in user_branch_ids:
                visible_units.append(('branch', b.id, b.name, b))
        for d in departments:
            if d.id in user_dept_ids:
                visible_units.append(('department', d.id, d.name, d))

        if not visible_units:
            if not branches and not departments:
                visible_units.append(('general', 0, 'General', None))
            else:
                continue

        checklist_month_expected = 0
        checklist_month_actual = 0
        checklist_year_expected = 0
        checklist_year_actual = 0

        for unit_type, unit_id, unit_name, unit_obj in visible_units:
            if unit_type == 'branch':
                unit_filter = {'branch_id': unit_id, 'department__isnull': True}
            elif unit_type == 'department':
                unit_filter = {'department_id': unit_id, 'branch__isnull': True}
            else:
                unit_filter = {'branch__isnull': True, 'department__isnull': True}

            m_exp = count_expected_occurrences(checklist, month_start, today)
            m_act = ChecklistLog.objects.filter(
                checklist=checklist, user=user_profile,
                log_date__gte=month_start, log_date__lte=today,
                **unit_filter
            ).values('log_date').distinct().count()

            y_exp = count_expected_occurrences(checklist, year_start, year_end)
            y_act = ChecklistLog.objects.filter(
                checklist=checklist, user=user_profile,
                log_date__gte=year_start, log_date__lte=year_end,
                **unit_filter
            ).values('log_date').distinct().count()

            checklist_month_expected += m_exp
            checklist_month_actual += m_act
            checklist_year_expected += y_exp
            checklist_year_actual += y_act

        month_expected += checklist_month_expected
        month_actual += checklist_month_actual
        year_expected += checklist_year_expected
        year_actual += checklist_year_actual

        month_progress = min(int(checklist_month_actual / checklist_month_expected * 100), 100) if checklist_month_expected > 0 else 0
        year_progress = min(int(checklist_year_actual / checklist_year_expected * 100), 100) if checklist_year_expected > 0 else 0

        if month_progress >= 100:
            status = 'completed'
            total_checklists_completed += 1
        elif month_progress >= 70:
            status = 'on_track'
        elif month_progress >= 40:
            status = 'in_progress'
        else:
            status = 'at_risk'

        next_due = checklist.get_next_due_date(user_profile) if hasattr(checklist, 'get_next_due_date') else None

        checklist.month_progress = month_progress
        checklist.year_progress = year_progress
        checklist.month_expected = checklist_month_expected
        checklist.month_actual = checklist_month_actual
        checklist.year_expected = checklist_year_expected
        checklist.year_actual = checklist_year_actual
        checklist.status = status
        checklist.next_due_date = next_due.strftime('%b %d, %Y') if next_due else None
        checklist.visible_units = visible_units

        display_checklists_with_progress.append(checklist)

    overall_month_progress = min(int(month_actual / month_expected * 100), 100) if month_expected > 0 else 0
    overall_quarter_progress = overall_month_progress
    overall_year_progress = min(int(year_actual / year_expected * 100), 100) if year_expected > 0 else 0

    freq_labels = {
        'daily': 'Daily', 'weekly': 'Weekly', 'monthly': 'Monthly',
        'quarterly': 'Quarterly', 'bi-annual': 'Bi-Annual', 'one-off': 'One-Off'
    }
    freq_parts = [f"{c} {freq_labels.get(f, f)}" for f, c in frequency_counts.items() if c > 0]
    frequency_summary = ', '.join(freq_parts) if freq_parts else 'No checklists'

    weekly_logs_count = ChecklistLog.objects.filter(
        user=user_profile,
        log_date__gte=week_start,
        log_date__lte=today
    ).count()

    weekly_expected = 0
    for checklist in user_checklists:
        branches = set(checklist.assigned_branches.values_list('id', flat=True))
        departments = set(checklist.assigned_departments.values_list('id', flat=True))
        visible_branches = branches & user_branch_ids
        visible_departments = departments & user_dept_ids
        unit_count = len(visible_branches) + len(visible_departments)
        if unit_count == 0 and not branches and not departments:
            unit_count = 1
        weekly_expected += count_expected_occurrences(checklist, week_start, today) * unit_count

    weekly_completion_rate = min(int(weekly_logs_count / weekly_expected * 100), 100) if weekly_expected > 0 else 0

    assigned_this_week = user_checklists.filter(
        created_at__date__gte=week_start,
        created_at__date__lte=week_end
    ).count()

    activity_logs = ActivityLog.objects.filter(user=user_profile).order_by('-created_at')[:10]

    total_logged_days = ChecklistLog.objects.filter(
        user=user_profile
    ).values('log_date').distinct().count()

    display_checklists_with_progress.sort(key=lambda x: x.month_progress)

    # ============================================================
    # IRREGULAR GL POSITIONS — date-aware
    # ------------------------------------------------------------
    # Default (on login): TODAY. If today has no trial balance yet,
    #                     fall back to the most recent available day.
    # With ?gl_date=YYYY-MM-DD: show that specific day.
    # ============================================================
    from .models import TrialBalanceEntry

    irregular_gls = []
    has_trial_balance = False
    available_tb_dates = []
    selectable_gl_dates = []
    selected_gl_date = None
    selected_gl_date_display = ''
    is_today_gl = False
    gl_date_param = request.GET.get('gl_date', '').strip()

    try:
        today_date = timezone.now().date()

        # All distinct dates that actually have trial-balance data
        # (most recent first, capped to keep the dropdown manageable)
        available_tb_dates = list(
            TrialBalanceEntry.objects
            .values_list('report_date', flat=True)
            .distinct()
            .order_by('-report_date')[:90]
        )

        # Resolve which date to display
        if gl_date_param:
            try:
                target_date = datetime.strptime(gl_date_param, '%Y-%m-%d').date()
            except ValueError:
                target_date = today_date
        else:
            # Fresh page load → default to TODAY.
            # If today has no TB data yet, fall back to the most recent day.
            target_date = today_date
            if not TrialBalanceEntry.objects.filter(report_date=today_date).exists():
                if available_tb_dates:
                    target_date = available_tb_dates[0]

        selected_gl_date = target_date
        selected_gl_date_display = target_date.strftime('%b %d, %Y')
        is_today_gl = (target_date == today_date)

        # Selectable dates for the dropdown:
        # today first (so it's always visible), then every other
        # available date, newest-first, without duplicates.
        selectable_gl_dates = [today_date]
        for d in available_tb_dates:
            if d != today_date:
                selectable_gl_dates.append(d)

        has_trial_balance = TrialBalanceEntry.objects.filter(
            report_date=target_date
        ).exists()

        # Branch / dept / role scoping (unchanged)
        code_to_name = dict(Branch.BRANCH_CODE_MAP)
        for b in Branch.objects.filter(is_active=True):
            if b.code:
                code_to_name[str(b.code).strip()] = b.name

        is_hc = (user_profile.position == 'hc')
        is_privileged = user_profile.role in ('admin', 'supervisor')

        user_branch_codes = set()
        for b in user_profile.branches.all():
            if b.code:
                user_branch_codes.add(str(b.code).strip())
            if b.name:
                for code, name in Branch.BRANCH_CODE_MAP.items():
                    if name.upper() == b.name.upper():
                        user_branch_codes.add(code)
                        break

        user_dept_names = {
            str(d.name).strip().upper()
            for d in user_profile.departments.all() if d.name
        }

        has_branch_restriction = len(user_branch_codes) > 0
        has_dept_restriction = len(user_dept_names) > 0

        entries_qs = TrialBalanceEntry.objects.filter(report_date=target_date)

        if not is_privileged:
            or_q = Q()
            if is_hc:
                or_q |= Q(branch_code='000')
            if not has_branch_restriction and not has_dept_restriction:
                or_q |= Q(pk__isnull=False)
            else:
                if has_branch_restriction:
                    or_q |= Q(branch_code__in=user_branch_codes)
                if has_dept_restriction:
                    dept_q = Q()
                    for dn in user_dept_names:
                        dept_q |= Q(category__iexact=dn)
                    or_q |= dept_q

            if or_q:
                entries_qs = entries_qs.filter(or_q)
            else:
                entries_qs = entries_qs.none()

        entries_qs = entries_qs.filter(
            Q(category__istartswith='ASSET') |
            Q(category__istartswith='EXPENSE') |
            Q(category__istartswith='LIABILIT') |
            Q(category__istartswith='INCOME')
        ).exclude(close_bal_lcy=0)

        for entry in entries_qs[:1000]:
            code = str(entry.branch_code or '').strip()
            category = str(entry.category or '').strip().upper()

            if category.startswith('ASSET') or category.startswith('EXPENSE'):
                expected_sign = 'negative'
            else:
                expected_sign = 'positive'

            try:
                bal = float(entry.close_bal_lcy or 0)
            except (ValueError, TypeError):
                bal = 0.0
            if bal == 0:
                continue

            actual_sign = 'positive' if bal > 0 else 'negative'
            if actual_sign == expected_sign:
                continue

            try:
                lcy_fmt = f'{float(entry.close_bal_lcy):,.2f}'
            except (ValueError, TypeError):
                lcy_fmt = str(entry.close_bal_lcy)
            try:
                fcy_fmt = f'{float(entry.close_bal_fcy):,.2f}'
            except (ValueError, TypeError):
                fcy_fmt = str(entry.close_bal_fcy)

            irregular_gls.append({
                'branch_code': entry.branch_code or '',
                'branch_name': code_to_name.get(code, code),
                'gl_code': entry.gl_code or '',
                'descr': entry.descr or '',
                'close_bal_lcy': lcy_fmt,
                'close_bal_fcy': fcy_fmt,
            })

            if len(irregular_gls) >= 100:
                break

    except Exception as e:
        logger.exception("Error computing irregular_gls in member_dashboard")
        irregular_gls = []
        has_trial_balance = False
        available_tb_dates = []
        selectable_gl_dates = []
        selected_gl_date = None
        selected_gl_date_display = ''
        is_today_gl = False

    context = {
        'user_profile': user_profile,
        'today': today,
        'week_start': week_start,
        'week_end': week_end,

        'all_checklists_count': total_checklist_rows,
        'total_checklists_completed': total_checklists_completed,
        'exceptions_this_week': exceptions_this_week,
        'total_exceptions_captured': total_exceptions_captured,
        'pending_submissions': pending_submissions,
        'assigned_this_week': assigned_this_week,

        'overall_month_progress': overall_month_progress,
        'overall_quarter_progress': overall_quarter_progress,
        'overall_year_progress': overall_year_progress,
        'weekly_completion_rate': weekly_completion_rate,
        'month_completed': month_actual,
        'month_total': month_expected,
        'year_completed': year_actual,
        'year_total': year_expected,
        'total_logged_days': total_logged_days,
        'frequency_summary': frequency_summary,

        'daily_checklists': display_checklists_with_progress[:5],
        'activity_logs': activity_logs,

        # --- Irregular GL panel (date-aware) ---
        'irregular_gls': irregular_gls,
        'has_trial_balance': has_trial_balance,
        'available_tb_dates': available_tb_dates,
        'selectable_gl_dates': selectable_gl_dates,
        'selected_gl_date': selected_gl_date,
        'selected_gl_date_display': selected_gl_date_display,
        'is_today_gl': is_today_gl,
        'gl_date_param': gl_date_param,

        'user_branches': user_profile.branches.all(),
        'user_departments': user_profile.departments.all(),

        'month_ring_offset': 100 - overall_month_progress,
        'quarter_ring_offset': 100 - overall_quarter_progress,
        'year_ring_offset': 100 - overall_year_progress,
        'weekly_ring_offset': 100 - weekly_completion_rate,
    }

    return render(request, 'control_dashboard/memberboard.html', context)

# ==================== SUBMIT REPORT VIEWS ====================

@login_required
def drafts_page(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    assigned_reports = Report.objects.filter(
        Q(assigned_to=user_profile) | Q(is_assigned_to_all=True)
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    ).filter(
        status__in=['assigned', 'in_progress']
    ).distinct().order_by('-created_at')

    report_data = []
    for report in assigned_reports:
        data_fields = report.data_fields.all().order_by('order')
        fields_dict = {field.field_name: field.field_value for field in data_fields}

        report_data.append({
            'report': report,
            'data_fields': data_fields,
            'fields_dict': fields_dict,
        })

    report_types = assigned_reports.values_list('report_type', flat=True).distinct()

    context = {
        'user_profile': user_profile,
        'today': timezone.now(),
        'report_types': list(report_types),
        'assigned_reports': assigned_reports,
        'report_data': report_data,
    }
    return render(request, 'control_dashboard/draft.html', context)

# control_dashboard/views.py — VERIFY/REPLACE reports_page

@login_required
def reports_page(request):
    """
    My Reports.
    
    Two distinct concepts are rendered:
    
      * submitted_reports — Report rows that the member actually sent
        via submit.html (statuses: submitted / completed / draft).
        These have corresponding SubmittedReportScore rows.
    
      * exception_uploads — Report rows created by draft.html Excel
        uploads (status='uploaded'). These are DATA CONTAINERS, not
        submissions. They are shown separately so the member can
        review / edit / delete the raw exception data they imported.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    # --- Real submitted reports (email sends) ---
    # These have statuses indicating they were actually submitted.
    submitted_reports = Report.objects.filter(
        created_by=user_profile
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS  # Exclude Excel data containers
    ).filter(
        Q(status='submitted') |
        Q(status='completed') |
        Q(status='draft')
    ).order_by('-created_at')

    # --- Excel exception data containers ---
    # These are the raw data imports from draft.html
    exception_uploads = Report.objects.filter(
    created_by=user_profile,
    status=UPLOADED_STATUS,
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).filter(
    Q(exception_records__isnull=False) | Q(data_fields__isnull=False)
    ).distinct().order_by('-created_at')

    # Get report types from actual submissions only
    report_types = submitted_reports.values_list('report_type', flat=True).distinct()
    branches = Branch.objects.filter(is_active=True).order_by('name')

    # Attach schedule info and display data
    for report in submitted_reports:
        schedule = ReportSchedule.objects.filter(report=report, is_active=True).first()
        report.has_schedule = bool(schedule)
        if schedule:
            report.next_due_date = schedule.next_due_date
            report.schedule_frequency = schedule.get_frequency_display()
        report._display_data = report.get_display_data()

    for report in exception_uploads:
        report._display_data = report.get_display_data()

    context = {
        'user_profile': user_profile,
        'today': timezone.now(),
        'reports': submitted_reports,           # Real submissions
        'exception_uploads': exception_uploads,  # Excel data containers
        'report_types': list(report_types),
        'branches': branches,
    }

    return render(request, 'control_dashboard/reports.html', context)

@login_required
def submit_page(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    # Only expose report types the member is actually assigned to
    # (directly, via "assigned to all", or via reports they created).
    # This guarantees every option has a real Report row behind it
    # with a real deadline — no more fabricated "Not set" rows.
    assigned_qs = Report.objects.filter(
        Q(assigned_to=user_profile) |
        Q(is_assigned_to_all=True) |
        Q(created_by=user_profile)
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    )

    report_types = list(
        assigned_qs.values_list('report_type', flat=True)
        .distinct().order_by('report_type')
    )

    all_users = UserProfile.objects.filter(status='active').order_by('full_name')

    assigned_exceptions = Report.objects.filter(
        Q(assigned_to=user_profile) |
        Q(created_by=user_profile) |
        Q(is_assigned_to_all=True)
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    ).filter(
        Q(status='submitted') | Q(status='draft') | Q(status='rejected')
    ).distinct().order_by('-updated_at')

    exceptions_data = []
    categories = set()

    for exception in assigned_exceptions:
        data_fields = exception.data_fields.all()
        exception_data = {}
        for field in data_fields:
            exception_data[field.field_name] = field.field_value

        unit = (
            exception_data.get('unit') or
            exception_data.get('branch') or
            exception_data.get('BRANCH') or
            exception_data.get('DEPARTMENT') or
            exception_data.get('BRANCH/UNIT') or
            exception_data.get('branch_unit') or
            'N/A'
        )

        if unit and unit != 'N/A':
            categories.add(unit)

        exceptions_data.append({
            'id': exception.id,
            'report_type': exception.report_type,
            'status': exception.status,
            'status_display': exception.get_status_display(),
            'data': exception_data,
            'created_at': exception.created_at.strftime('%b %d, %Y'),
            'unit': unit,
            'category': unit if unit != 'N/A' else 'uncategorized',
            'assigned_to': user_profile.email,
            'created_by': exception.created_by.email if exception.created_by else ''
        })

    categories_list = [{'id': cat, 'name': cat} for cat in categories if cat]
    exception_report_types = list(set([e['report_type'] for e in exceptions_data]))

    submission_data = request.session.pop('submission_data', None)

    context = {
        'user_profile': user_profile,
        'today': timezone.now(),
        'all_users': all_users,
        'report_types': exception_report_types or report_types,
        'assigned_report_types': exception_report_types,
        'categories_list': categories_list,
        'exceptions_json': json.dumps(exceptions_data, default=str),
        'exceptions': exceptions_data,
        'email_recipients': all_users,
        'user_data': {
            'name': user_profile.full_name,
            'email': user_profile.email,
            'username': user_profile.username,
            'department': user_profile.position or 'General'
        },
        'pre_filled_data': json.dumps(submission_data) if submission_data else '',
        'excel_data': json.dumps(submission_data.get('excel_data')) if submission_data and submission_data.get('excel_data') else '',
        'screenshot_data': json.dumps(submission_data.get('screenshot_data')) if submission_data and submission_data.get('screenshot_data') else '',
    }

    return render(request, 'control_dashboard/submit.html', context)


# ==================== API - IMPORT/EXPORT EXCEL ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_import_excel(request):
    try:
        data = json.loads(request.body)
        file_data = data.get('file_data', [])
        file_name = data.get('file_name', 'uploaded_file.xlsx')
        report_type = data.get('report_type', '')

        if not file_data or len(file_data) == 0:
            return JsonResponse({'success': False, 'error': 'No data found in the file'}, status=400)

        headers = file_data[0] if file_data else []
        rows = file_data[1:] if len(file_data) > 1 else []

        parsed_data = []
        for row in rows:
            if row and any(cell for cell in row):
                row_dict = {}
                for i, header in enumerate(headers):
                    if i < len(row):
                        row_dict[header] = row[i] if row[i] is not None else ''
                    else:
                        row_dict[header] = ''
                parsed_data.append(row_dict)

        request.session['excel_import_data'] = {
            'headers': headers,
            'data': parsed_data,
            'file_name': file_name,
            'report_type': report_type,
            'row_count': len(parsed_data)
        }

        return JsonResponse({
            'success': True,
            'headers': headers,
            'data': parsed_data,
            'row_count': len(parsed_data),
            'message': f'Successfully imported {len(parsed_data)} records'
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error importing Excel")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

# control_dashboard/views.py — REPLACE the api_save_imported_data function

@csrf_exempt
@require_http_methods(["POST"])
def api_save_imported_data(request):
    """
    Persist an Excel upload from draft.html as typed exception records.

    ONLY ExceptionRecord rows are created for Excel uploads.
    The Report container stores headers on `excel_headers` so the UI
    can render the correct columns.

    NO ReportDataField rows are created here — those are only for
    single-row form-based reports from submit.html.
    """
    try:
        data = json.loads(request.body)

        headers = data.get('headers', [])
        rows_data = data.get('data', [])
        report_type = data.get('report_type', '')
        file_name = data.get('file_name', 'uploaded.xlsx')

        if not headers or not rows_data:
            return JsonResponse({'success': False, 'error': 'No data to save. Please import an Excel file first.'}, status=400)

        if not report_type:
            return JsonResponse({'success': False, 'error': 'Please select a report type.'}, status=400)

        if report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'This report type is reserved for internal use.'}, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        validation = _validate_exception_headers(headers)

        if validation['status'] == 'error':
            return JsonResponse({
                'success': False,
                'error': (
                    'Column headers in row 1 do not match the standard exception '
                    'template. Matched {matched} of {total}. Missing: {missing}'
                ).format(
                    matched=validation['match_count'],
                    total=validation['total_canonical'],
                    missing=', '.join(validation['missing']),
                ),
                'validation': validation,
            }, status=400)

        # Remove any empty containers left over from previous failed uploads
        Report.objects.filter(
            created_by=user_profile,
            status=UPLOADED_STATUS,
            exception_records__isnull=True,
        ).delete()

        # Create the Report container — NO ReportDataField rows!
        report = Report.objects.create(
            report_type=report_type,
            frequency='one-off',
            description=f'Excel Import: {file_name}',
            status=UPLOADED_STATUS,
            created_by=user_profile,
            excel_headers=','.join(str(h) for h in headers),
        )

        # Build typed ExceptionRecord rows — this is the ONLY data storage
        typed_count = 0
        try:
            typed_count = _build_exception_records(
                report=report,
                headers=headers,
                rows_data=rows_data,
            )
        except Exception as exc:
            logger.exception("Failed to build typed ExceptionRecord rows: %s", exc)

        log_activity(
            user=user_profile,
            activity_type='report_submitted',
            details=(
                f'Uploaded {report_type} with {len(rows_data)} exception records '
                f'({typed_count} typed) from Excel'
            ),
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': (
                f'Successfully saved {len(rows_data)} records from {file_name} '
                f'({typed_count} typed as exceptions)'
            ),
            'report_id': report.id,
            'record_count': len(rows_data),
            'typed_record_count': typed_count,
            'validation': validation,
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error saving imported data")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@login_required
def member_checklist(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    today = timezone.now().date()

    branch_filter = request.GET.get('branch', 'all')
    department_filter = request.GET.get('department', 'all')
    month_filter = request.GET.get('month', 'all')
    year_filter = request.GET.get('year', 'all')
    frequency_filter = request.GET.get('frequency', 'all')
    quarter_filter = request.GET.get('quarter', 'all')

    user_branches = user_profile.branches.all().order_by('name')
    user_departments = user_profile.departments.all().order_by('name')

    user_branch_ids = set(user_branches.values_list('id', flat=True))
    user_department_ids = set(user_departments.values_list('id', flat=True))

    user_checklists = Checklist.objects.filter(
        is_active=True
    ).filter(
        Q(assigned_users=user_profile) |
        Q(assignment_target='all') |
        Q(assignment_target=user_profile.position)
    ).distinct()

    if branch_filter != 'all':
        try:
            user_checklists = user_checklists.filter(assigned_branches__id=int(branch_filter)).distinct()
        except ValueError:
            pass

    if department_filter != 'all':
        try:
            user_checklists = user_checklists.filter(assigned_departments__id=int(department_filter)).distinct()
        except ValueError:
            pass

    if frequency_filter != 'all':
        user_checklists = user_checklists.filter(frequency=frequency_filter)

    current_year = timezone.now().year
    years = list(range(current_year - 2, current_year + 1))

    display_month = today.month
    display_year = today.year

    if month_filter != 'all':
        try:
            display_month = int(month_filter) + 1
        except ValueError:
            pass

    if year_filter != 'all':
        try:
            display_year = int(year_filter)
        except ValueError:
            pass

    current_quarter = (display_month - 1) // 3 + 1

    if quarter_filter != 'all':
        try:
            current_quarter = int(quarter_filter)
        except ValueError:
            pass

    quarter_start_month = (current_quarter - 1) * 3 + 1
    quarter_end_month = current_quarter * 3

    quarter_start = date(display_year, quarter_start_month, 1)
    if quarter_end_month == 12:
        quarter_end = date(display_year + 1, 1, 1) - timedelta(days=1)
    else:
        quarter_end = date(display_year, quarter_end_month + 1, 1) - timedelta(days=1)

    month_start = date(display_year, display_month, 1)
    if display_month == 12:
        month_end = date(display_year + 1, 1, 1) - timedelta(days=1)
    else:
        month_end = date(display_year, display_month + 1, 1) - timedelta(days=1)

    year_start = date(display_year, 1, 1)
    year_end = date(display_year, 12, 31)

    def count_actual(checklist, unit_filter, period_start, period_end):
        return ChecklistLog.objects.filter(
            checklist=checklist,
            user=user_profile,
            log_date__gte=period_start,
            log_date__lte=period_end,
            **unit_filter
        ).values('log_date').distinct().count()

    checklist_data = []
    total_month_expected = 0
    total_month_actual = 0
    total_quarter_expected = 0
    total_quarter_actual = 0
    total_year_expected = 0
    total_year_actual = 0

    for checklist in user_checklists:
        tasks = list(checklist.tasks.all().order_by('order'))

        checklist_branches = checklist.assigned_branches.all()
        checklist_departments = checklist.assigned_departments.all()

        if branch_filter != 'all':
            try:
                aid = int(branch_filter)
                checklist_branches = checklist_branches.filter(id=aid) if aid in user_branch_ids else checklist_branches.none()
            except ValueError:
                checklist_branches = checklist_branches.none()
        else:
            checklist_branches = checklist_branches.filter(id__in=user_branch_ids)

        if department_filter != 'all':
            try:
                aid = int(department_filter)
                checklist_departments = checklist_departments.filter(id=aid) if aid in user_department_ids else checklist_departments.none()
            except ValueError:
                checklist_departments = checklist_departments.none()
        else:
            checklist_departments = checklist_departments.filter(id__in=user_department_ids)

        assign_units = (
            [('branch', b.id, b.name, b) for b in checklist_branches] +
            [('department', d.id, d.name, d) for d in checklist_departments]
        )

        if not assign_units:
            if checklist.assigned_branches.exists() or checklist.assigned_departments.exists():
                continue
            assign_units.append(('general', 0, 'General', None))

        next_due = checklist.get_next_due_date(user_profile) if hasattr(checklist, 'get_next_due_date') else None

        for unit_type, unit_id, unit_name, unit_obj in assign_units:
            if unit_type == 'branch':
                unit_filter = {'branch_id': unit_id, 'department__isnull': True}
            elif unit_type == 'department':
                unit_filter = {'department_id': unit_id, 'branch__isnull': True}
            else:
                unit_filter = {'branch__isnull': True, 'department__isnull': True}

            month_log_dates = list(
                ChecklistLog.objects.filter(
                    checklist=checklist,
                    user=user_profile,
                    log_date__gte=month_start,
                    log_date__lte=month_end,
                    **unit_filter
                ).values_list('log_date', flat=True)
            )
            log_dates_list = [d.strftime('%Y-%m-%d') for d in month_log_dates]

            month_expected = count_expected_occurrences(checklist, month_start, month_end)
            month_actual = count_actual(checklist, unit_filter, month_start, month_end)

            quarter_expected = count_expected_occurrences(checklist, quarter_start, quarter_end)
            quarter_actual = count_actual(checklist, unit_filter, quarter_start, quarter_end)

            year_expected = count_expected_occurrences(checklist, year_start, year_end)
            year_actual = count_actual(checklist, unit_filter, year_start, year_end)

            month_progress = min(int(month_actual / month_expected * 100), 100) if month_expected > 0 else 0
            quarter_progress = min(int(quarter_actual / quarter_expected * 100), 100) if quarter_expected > 0 else 0
            year_progress = min(int(year_actual / year_expected * 100), 100) if year_expected > 0 else 0

            if month_progress >= 100:
                status = 'completed'
            elif month_progress >= 70:
                status = 'on_track'
            elif month_progress >= 40:
                status = 'in_progress'
            else:
                status = 'at_risk'

            unit_display = (
                f"🏢 {unit_name}" if unit_type == 'branch' else
                f"🏛️ {unit_name}" if unit_type == 'department' else
                "📋 General"
            )
            row_key = f"{checklist.id}_{unit_type}_{unit_id}"

            checklist_data.append({
                'id': checklist.id,
                'row_key': row_key,
                'unit_type': unit_type,
                'unit_id': unit_id,
                'unit_name': unit_name,
                'unit_display': unit_display,
                'name': checklist.name,
                'frequency': checklist.frequency,
                'frequency_display': checklist.get_frequency_display(),
                'tasks': [{'description': t.description, 'id': t.id} for t in tasks],
                'logs': log_dates_list,
                'month_progress': month_progress,
                'quarter_progress': quarter_progress,
                'year_progress': year_progress,
                'month_expected': month_expected,
                'month_actual': month_actual,
                'quarter_expected': quarter_expected,
                'quarter_actual': quarter_actual,
                'year_expected': year_expected,
                'year_actual': year_actual,
                'status': status,
                'frequency_days': checklist.get_frequency_days(),
                'next_due_date': next_due.strftime('%b %d, %Y') if next_due else None,
                'can_log': True,
                'assigned_branches': list(checklist.assigned_branches.values_list('name', flat=True)),
                'assigned_departments': list(checklist.assigned_departments.values_list('name', flat=True)),
            })

            total_month_expected += month_expected
            total_month_actual += month_actual
            total_quarter_expected += quarter_expected
            total_quarter_actual += quarter_actual
            total_year_expected += year_expected
            total_year_actual += year_actual

    overall_month_progress = min(int(total_month_actual / total_month_expected * 100), 100) if total_month_expected > 0 else 0
    overall_quarter_progress = min(int(total_quarter_actual / total_quarter_expected * 100), 100) if total_quarter_expected > 0 else 0
    overall_year_progress = min(int(total_year_actual / total_year_expected * 100), 100) if total_year_expected > 0 else 0

    available_branches = user_branches.filter(checklist_assignments__in=user_checklists).distinct().order_by('name')
    available_departments = user_departments.filter(checklist_assignments__in=user_checklists).distinct().order_by('name')

    combined_filter_options = [{'type': 'all', 'id': 'all', 'display': 'All Branches/Departments', 'value': 'all'}]
    for branch in available_branches:
        combined_filter_options.append({
            'type': 'branch', 'id': branch.id,
            'display': f'🏢 {branch.name}', 'value': f'branch_{branch.id}'
        })
    for dept in available_departments:
        combined_filter_options.append({
            'type': 'department', 'id': dept.id,
            'display': f'🏛️ {dept.name}', 'value': f'department_{dept.id}'
        })

    active_combined_filter = 'all'
    if branch_filter != 'all':
        active_combined_filter = f'branch_{branch_filter}'
    elif department_filter != 'all':
        active_combined_filter = f'department_{department_filter}'

    selected_combined_display = 'All Branches/Departments'
    for option in combined_filter_options:
        if option['value'] == active_combined_filter:
            selected_combined_display = option['display']
            break

    available_frequencies = Checklist.FREQUENCY_CHOICES

    selected_frequency_display = 'All Frequencies'
    if frequency_filter != 'all':
        for fv, fl in Checklist.FREQUENCY_CHOICES:
            if fv == frequency_filter:
                selected_frequency_display = fl
                break

    month_names = [
        ('all', 'All Months'),
        ('0', 'January'), ('1', 'February'), ('2', 'March'), ('3', 'April'),
        ('4', 'May'), ('5', 'June'), ('6', 'July'), ('7', 'August'),
        ('8', 'September'), ('9', 'October'), ('10', 'November'), ('11', 'December'),
    ]

    quarter_options = [
        ('all', 'All Quarters'),
        ('1', 'Q1 (Jan - Mar)'),
        ('2', 'Q2 (Apr - Jun)'),
        ('3', 'Q3 (Jul - Sep)'),
        ('4', 'Q4 (Oct - Dec)'),
    ]

    quarter_names = {1: 'Q1', 2: 'Q2', 3: 'Q3', 4: 'Q4'}
    selected_quarter_display = 'All Quarters'
    if quarter_filter != 'all':
        try:
            q_num = int(quarter_filter)
            selected_quarter_display = quarter_names.get(q_num, f'Q{q_num}')
        except ValueError:
            pass

    context = {
        'user_profile': user_profile,
        'checklists': user_checklists,
        'checklist_data': checklist_data,
        'years': years,
        'overall_month_progress': overall_month_progress,
        'overall_quarter_progress': overall_quarter_progress,
        'overall_year_progress': overall_year_progress,
        'total_checklists': user_checklists.count(),
        'total_month_expected': total_month_expected,
        'total_month_actual': total_month_actual,
        'total_quarter_expected': total_quarter_expected,
        'total_quarter_actual': total_quarter_actual,
        'total_year_expected': total_year_expected,
        'total_year_actual': total_year_actual,
        'branch_filter': branch_filter,
        'department_filter': department_filter,
        'month_filter': month_filter,
        'year_filter': year_filter,
        'frequency_filter': frequency_filter,
        'quarter_filter': quarter_filter,
        'available_frequencies': available_frequencies,
        'month_names': month_names,
        'quarter_options': quarter_options,
        'selected_frequency_display': selected_frequency_display,
        'selected_quarter_display': selected_quarter_display,
        'today': today,
        'month_name': today.strftime('%B'),
        'year': today.year,
        'display_month': display_month,
        'display_year': display_year,
        'current_quarter': current_quarter,
        'combined_filter_options': combined_filter_options,
        'active_combined_filter': active_combined_filter,
        'selected_combined_display': selected_combined_display,
    }

    return render(request, 'control_dashboard/checklist-mem.html', context)


# ==================== API - CHECKLIST LOG ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_log_checklist(request):
    try:
        data = json.loads(request.body)

        checklist_id = data.get('checklist_id')
        log_date = data.get('log_date')
        action = data.get('action', 'log')
        unit_type = data.get('unit_type', 'general')
        unit_id = data.get('unit_id')

        if not checklist_id:
            return JsonResponse({'success': False, 'error': 'Checklist ID is required'}, status=400)

        if not log_date:
            return JsonResponse({'success': False, 'error': 'Log date is required'}, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        try:
            checklist = Checklist.objects.get(id=checklist_id, is_active=True)
        except Checklist.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'Checklist not found'}, status=404)

        is_assigned = (
            checklist.assigned_users.filter(id=user_profile.id).exists() or
            checklist.assignment_target == 'all' or
            checklist.assignment_target == user_profile.position
        )
        if not is_assigned:
            return JsonResponse({'success': False, 'error': 'You are not assigned to this checklist'}, status=403)

        try:
            date_obj = datetime.strptime(log_date, '%Y-%m-%d').date()
        except ValueError:
            return JsonResponse({'success': False, 'error': 'Invalid date format. Use YYYY-MM-DD'}, status=400)

        branch = None
        department = None

        if unit_type == 'branch' and unit_id:
            try:
                branch = Branch.objects.get(id=int(unit_id), is_active=True)
            except (ValueError, Branch.DoesNotExist):
                return JsonResponse({'success': False, 'error': 'Invalid branch'}, status=400)
        elif unit_type == 'department' and unit_id:
            try:
                department = Department.objects.get(id=int(unit_id), is_active=True)
            except (ValueError, Department.DoesNotExist):
                return JsonResponse({'success': False, 'error': 'Invalid department'}, status=400)

        if action == 'log':
            log_entry, created = ChecklistLog.objects.get_or_create(
                checklist=checklist,
                user=user_profile,
                branch=branch,
                department=department,
                log_date=date_obj
            )

            if created:
                return JsonResponse({'success': True, 'message': 'Checklist logged successfully', 'action': 'logged'})
            else:
                return JsonResponse({'success': True, 'message': 'Checklist already logged for this date', 'action': 'already_logged'})

        elif action == 'unlog':
            deleted_count, _ = ChecklistLog.objects.filter(
                checklist=checklist,
                user=user_profile,
                branch=branch,
                department=department,
                log_date=date_obj
            ).delete()

            if deleted_count > 0:
                return JsonResponse({'success': True, 'message': 'Checklist unlogged successfully', 'action': 'unlogged'})
            else:
                return JsonResponse({'success': False, 'error': 'No log found for this date'}, status=404)

        else:
            return JsonResponse({'success': False, 'error': 'Invalid action. Use "log" or "unlog"'}, status=400)

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def api_get_checklist_logs(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        logs = ChecklistLog.objects.filter(user=user_profile).select_related('checklist')

        log_data = []
        for log in logs:
            log_data.append({
                'checklist_id': log.checklist.id,
                'checklist_name': log.checklist.name,
                'log_date': log.log_date.strftime('%Y-%m-%d'),
                'logged_at': log.created_at.isoformat(),
            })

        return JsonResponse({'success': True, 'logs': log_data})

    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def api_get_checklist_stats(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)

        user_checklists = Checklist.objects.filter(
            is_active=True
        ).filter(
            Q(assigned_users=user_profile) |
            Q(assignment_target='all') |
            Q(assignment_target=user_profile.position)
        ).distinct()

        total_checklists = user_checklists.count()
        logs = ChecklistLog.objects.filter(user=user_profile)
        total_logs = logs.count()

        today = timezone.now().date()
        start_of_week = get_week_start(today)
        end_of_week = start_of_week + timedelta(days=6)

        weekly_logs = logs.filter(log_date__gte=start_of_week, log_date__lte=end_of_week).count()
        last_7_days = today - timedelta(days=7)
        recent_logs = logs.filter(log_date__gte=last_7_days).count()

        daily_checklists = user_checklists.filter(frequency='daily')
        daily_total = daily_checklists.count()

        completion_rate = 0
        if daily_total > 0:
            completed_days = 0
            for i in range(7):
                day = start_of_week + timedelta(days=i)
                if day > today:
                    break
                completed_for_day = logs.filter(log_date=day).count()
                if completed_for_day >= daily_total:
                    completed_days += 1
            completion_rate = int((completed_days / 7) * 100)

        return JsonResponse({
            'success': True,
            'stats': {
                'total_checklists': total_checklists,
                'total_logs': total_logs,
                'weekly_logs': weekly_logs,
                'recent_logs': recent_logs,
                'completion_rate': completion_rate,
                'daily_checklists': daily_total,
            }
        })

    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ==================== API - SAVE DRAFT ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_save_draft(request):
    try:
        data = json.loads(request.body)

        report_type = data.get('report_type', '').strip()
        template_name = data.get('template_name', '').strip()
        form_data = data.get('form_data', [])
        excel_data = data.get('excel_data', [])
        status = data.get('status', 'draft')

        if not report_type:
            return JsonResponse({'success': False, 'error': 'Report type is required'}, status=400)

        if report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'This report type is reserved for internal use.'}, status=400)

        if status == UPLOADED_STATUS:
            return JsonResponse({'success': False, 'error': 'This status is reserved for internal use.'}, status=400)

        if not form_data:
            return JsonResponse({'success': False, 'error': 'No form data provided'}, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        submission = ReportSubmission.objects.create(
            report_type=report_type,
            template_name=template_name,
            submitted_by=user_profile,
            status=status,
            data={'form_data': form_data, 'excel_data': excel_data},
        )

        # Mirror the draft into a Report row so it appears in "My Reports"
        # and can later be submitted via submit.html.
        report = Report.objects.create(
            report_type=report_type,
            frequency='one-off',
            description=template_name or f'Draft: {report_type}',
            status=status,
            created_by=user_profile,
            data={
                'submission_id': submission.id,
                'form_data': form_data,
                'excel_data': excel_data,
            },
        )

        return JsonResponse({
            'success': True,
            'message': 'Draft saved successfully',
            'submission_id': submission.id,
            'report_id': report.id
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error in api_save_draft")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ==================== API - DRAFT (GET, EDIT, DELETE) ====================

@csrf_exempt
@require_http_methods(["GET"])
def api_get_draft(request, report_id):
    try:
        lookup_id = int(report_id) if str(report_id).isdigit() else report_id
        report = get_object_or_404(Report, id=lookup_id)

        if report.report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User profile context not found'}, status=404)

        report_data = {}
        form_data = {}

        data_fields = report.data_fields.all()
        for field in data_fields:
            form_data[field.field_name] = field.field_value or ''

        excel_import = report.excel_imports.first()
        if excel_import:
            report_data['import_type'] = 'excel'
            report_data['headers'] = excel_import.get_headers_list()
            report_data['file_name'] = excel_import.file_name
            rows_data = []
            for row in excel_import.rows.all().order_by('row_index'):
                row_dict = {}
                for cell in row.cells.all():
                    row_dict[cell.column_name] = cell.value or ''
                rows_data.append(row_dict)
            report_data['data'] = rows_data
            report_data['row_count'] = len(rows_data)
        else:
            report_data['form_data'] = [form_data] if form_data else []

        response_data = {
            'success': True,
            'report': {
                'id': str(report.id),
                'report_type': report.report_type,
                'status': report.status,
                'data': report_data,
            }
        }
        return JsonResponse(response_data)
    except Exception as e:
        logger.exception("Error in api_get_draft")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["POST"])
def api_edit_draft(request, report_id):
    try:
        lookup_id = int(report_id) if str(report_id).isdigit() else report_id
        report = get_object_or_404(Report, id=lookup_id)

        if report.report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User profile context not found'}, status=404)

        data = json.loads(request.body)
        form_data = data.get('form_data', {})
        is_excel = data.get('is_excel', False)

        if is_excel and 'headers' in form_data and 'data' in form_data:
            headers = form_data.get('headers', [])
            rows_data = form_data.get('data', [])

            excel_import = report.excel_imports.first()
            if excel_import:
                excel_import.rows.all().delete()
                excel_import.headers = ','.join(headers)
                excel_import.row_count = len(rows_data)
                excel_import.save()
            else:
                excel_import = ReportExcelImport.objects.create(
                    report=report,
                    file_name=form_data.get('file_name', 'edited.xlsx'),
                    headers=','.join(headers),
                    row_count=len(rows_data)
                )

            for row_idx, row in enumerate(rows_data):
                excel_row = ReportExcelRow.objects.create(import_record=excel_import, row_index=row_idx)
                for i, value in enumerate(row):
                    if i < len(headers):
                        ReportExcelCell.objects.create(
                            row=excel_row,
                            column_name=headers[i] or f'Column_{i+1}',
                            value=str(value) if value is not None else ''
                        )

            for i, header in enumerate(headers):
                ReportDataField.objects.get_or_create(
                    report=report,
                    field_name=header or f'Column_{i+1}',
                    defaults={'field_value': '', 'field_type': 'text', 'order': i}
                )

            if report.status == 'assigned':
                report.status = 'in_progress'
            report.save()

            return JsonResponse({
                'success': True,
                'message': 'Excel report updated successfully',
                'data': {'headers': headers, 'rows': len(rows_data)}
            })

        else:
            for field_name, field_value in form_data.items():
                ReportDataField.objects.update_or_create(
                    report=report,
                    field_name=field_name,
                    defaults={'field_value': field_value, 'field_type': 'text'}
                )

            if report.status == 'assigned':
                report.status = 'in_progress'
            report.save()

            return JsonResponse({
                'success': True,
                'message': 'Report modifications saved successfully',
                'data': form_data
            })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON format received'}, status=400)
    except Exception as e:
        logger.exception("Exception during edit save")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_draft(request, report_id):
    try:
        lookup_id = int(report_id) if str(report_id).isdigit() else report_id
        report = get_object_or_404(Report, id=lookup_id)

        if report.report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User context profile not found'}, status=404)

        # Explicit cleanup — cascades would handle it, but be clear.
        ExceptionRecord.objects.filter(report=report).delete()
        report.data_fields.all().delete()
        report.excel_imports.all().delete()   # cascades to rows + cells
        report.delete()

        return JsonResponse({
            'success': True,
            'message': 'Draft and associated data permanently removed from database storage.'
        })

    except Exception as e:
        logger.exception("Error in api_delete_draft")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@csrf_exempt
@require_http_methods(["GET"])
def api_get_report_data(request):
    try:
        report_type = request.GET.get('report_type', '')

        if not report_type:
            return JsonResponse({'success': False, 'error': 'Report type is required'}, status=400)

        user_profile = UserProfile.objects.get(email=request.user.email)

        submissions = ReportSubmission.objects.filter(
            report_type=report_type,
            submitted_by=user_profile
        ).order_by('-submission_date')

        if submissions.exists():
            latest = submissions.first()
            form_data = (latest.data or {}).get('form_data', [])
            return JsonResponse({'success': True, 'data': form_data, 'submission_id': latest.id})
        else:
            return JsonResponse({'success': True, 'data': []})

    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ==================== MEMBER ACTIVITY LOGS ====================

@login_required
def member_activity_logs(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    start_date = request.GET.get('start_date', '')
    end_date = request.GET.get('end_date', '')
    activity_filter = request.GET.get('activity', '')

    queryset = ActivityLog.objects.filter(user=user_profile)

    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__gte=start_date_obj)
        except ValueError:
            pass

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__lte=end_date_obj)
        except ValueError:
            pass

    if activity_filter and activity_filter != 'all':
        queryset = queryset.filter(activity_type=activity_filter)

    activity_logs = queryset.order_by('-created_at')[:100]

    today = timezone.now().date()
    start_of_week = get_week_start(today)
    start_of_month = today.replace(day=1)

    today_logs = ActivityLog.objects.filter(user=user_profile, created_at__date=today)
    week_logs = ActivityLog.objects.filter(user=user_profile, created_at__date__gte=start_of_week, created_at__date__lte=today)
    month_logs = ActivityLog.objects.filter(user=user_profile, created_at__date__gte=start_of_month, created_at__date__lte=today)

    last_activity = ActivityLog.objects.filter(user=user_profile).order_by('-created_at').first()
    last_activity_display = last_activity.created_at.strftime('%b %d, %Y %H:%M') if last_activity else None

    activity_types = ActivityLog.ACTIVITY_TYPES

    context = {
        'user_profile': user_profile,
        'activity_logs': activity_logs,
        'activity_types': activity_types,
        'activity_filter': activity_filter,
        'start_date': start_date,
        'end_date': end_date,
        'today_logs': today_logs,
        'week_logs': week_logs,
        'month_logs': month_logs,
        'last_activity': last_activity_display,
        'today': today,
    }

    return render(request, 'control_dashboard/activity-mem.html', context)


# ==================== API - EXPORT LOGS ====================

@csrf_exempt
@require_http_methods(["GET"])
def api_export_logs(request):
    try:
        UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

    user_filter = request.GET.get('user', 'all')
    start_date = request.GET.get('start_date', '')
    end_date = request.GET.get('end_date', '')
    activity_filter = request.GET.get('activity', '')

    queryset = ActivityLog.objects.all().select_related('user')

    if user_filter != 'all':
        try:
            queryset = queryset.filter(user_id=int(user_filter))
        except ValueError:
            pass

    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__gte=start_date_obj)
        except ValueError:
            pass

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__lte=end_date_obj)
        except ValueError:
            pass

    if activity_filter and activity_filter != 'all':
        queryset = queryset.filter(activity_type=activity_filter)

    logs = queryset.order_by('-created_at')

    import csv

    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = f'attachment; filename="activity_logs_{timezone.now().strftime("%Y%m%d")}.csv"'

    writer = csv.writer(response)
    writer.writerow(['Date & Time', 'User', 'Activity Type', 'Details', 'IP Address'])

    for log in logs:
        writer.writerow([
            log.created_at.strftime('%Y-%m-%d %H:%M:%S'),
            log.user.full_name or log.user.email,
            log.get_activity_type_display(),
            log.details or '',
            log.ip_address or ''
        ])

    return response


# ==================== SUPERVISOR VIEWS (CONTINUED) ====================

# ==================== SUPERVISOR VIEWS (CONTINUED) ====================

@login_required
def team_performance(request):
    """
    Team Performance.

    Scoring model (from AdHocDeduction):
      • Total Credits = Σ points_added
      • Total Debits  = Σ points
      • Net Score     = Credits − Debits  (can be negative)
      • Performance % = clamp(Net Score, 0, 100)

    The performance percentage is the NET of credits minus debits,
    capped at 100% so it can never exceed the ceiling.
    Excel uploads (status='uploaded') are EXCLUDED from the task
    denominator.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    today = timezone.now()

    team_members = UserProfile.objects.filter(
        role='member', status='active'
    ).order_by('full_name')

    team_data = []
    total_members = team_members.count()
    grand_total_tasks = 0
    grand_total_completed = 0
    grand_total_debits = 0
    grand_total_credits = 0
    sum_of_percentages = 0

    for member in team_members:
        # ------------------------------------------------------------
        # 1. Total Tasks (denominator) — expected occurrences across
        #    every Report assigned to this member (excluding uploads
        #    and the internal trial balance container).
        # ------------------------------------------------------------
        assigned_reports = Report.objects.filter(
            Q(assigned_to=member) | Q(is_assigned_to_all=True)
        ).exclude(
            report_type=TRIAL_BALANCE_REPORT_TYPE
        ).exclude(
            status=UPLOADED_STATUS
        ).distinct()

        member_total_tasks = 0
        for report in assigned_reports:
            member_total_tasks += count_expected_report_occurrences(report, until=today)

        # ------------------------------------------------------------
        # 2. Total Debits (deductions) and Total Credits (bonuses)
        #    pulled live from control_dashboard_adhocdeduction.
        # ------------------------------------------------------------
        member_debits = (
            AdHocDeduction.objects
            .filter(user=member)
            .aggregate(total=Sum('points'))
            .get('total') or 0
        )
        member_credits = (
            AdHocDeduction.objects
            .filter(user=member)
            .aggregate(total=Sum('points_added'))
            .get('total') or 0
        )

        # ------------------------------------------------------------
        # 3. Net Score = Credits − Debits
        #    The difference is CAPPED at 100% and floored at 0%.
        # ------------------------------------------------------------
        net_score = member_credits - member_debits
        raw_percentage = net_score              # signed, for the "Net" column
        percentage = max(0, min(100, net_score))  # 0 ≤ performance ≤ 100

        if percentage >= 90:
            status = 'success'
            status_text = 'Outstanding'
        elif percentage >= 70:
            status = 'success'
            status_text = 'Excellent'
        elif percentage >= 50:
            status = 'warning'
            status_text = 'In Progress'
        elif percentage > 0:
            status = 'warning'
            status_text = 'Building Up'
        else:
            status = 'danger'
            status_text = 'Needs Attention'

        team_data.append({
            'user': member,
            'total_tasks': member_total_tasks,
            'completed': 0,
            'debits': member_debits,
            'credits': member_credits,
            'net_score': net_score,
            'raw_percentage': raw_percentage,
            'percentage': percentage,
            'percentage_bar': percentage,
            'status': status,
            'status_text': status_text,
        })

        grand_total_tasks += member_total_tasks
        grand_total_completed += 0
        grand_total_debits += member_debits
        grand_total_credits += member_credits
        sum_of_percentages += percentage

    team_data.sort(key=lambda x: x['percentage'], reverse=True)

    overall_completion = (
        min(100, int(sum_of_percentages / total_members)) if total_members else 0
    )

    context = {
        'user_profile': user_profile,
        'team_data': team_data,
        'total_members': total_members,
        'total_tasks': grand_total_tasks,
        'total_completed': grand_total_completed,
        'total_debits': grand_total_debits,
        'total_credits': grand_total_credits,
        'overall_completion': overall_completion,
    }

    return render(request, 'control_dashboard/team.html', context)


@csrf_exempt
@require_http_methods(["GET"])
def api_team_performance_live(request):
    """
    Lightweight JSON endpoint for real-time team performance refreshes.

    Returns the same scoring model used by `team_performance`:
      Net = Credits − Debits,  Performance % = clamp(Net, 0, 100).
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

    today = timezone.now()
    team_members = UserProfile.objects.filter(role='member', status='active').order_by('full_name')

    members_payload = []
    for member in team_members:
        member_debits = (
            AdHocDeduction.objects
            .filter(user=member)
            .aggregate(total=Sum('points'))
            .get('total') or 0
        )
        member_credits = (
            AdHocDeduction.objects
            .filter(user=member)
            .aggregate(total=Sum('points_added'))
            .get('total') or 0
        )

        net_score = member_credits - member_debits
        percentage = max(0, min(100, net_score))

        members_payload.append({
            'id': member.id,
            'full_name': member.full_name or member.email,
            'debits': member_debits,
            'credits': member_credits,
            'net_score': net_score,
            'percentage': percentage,
        })

    return JsonResponse({'success': True, 'members': members_payload})

@login_required
def submitted_reports(request):
    """
    Submitted Reports (supervisor view).
    
    Now shows Excel uploads (status='uploaded') since email
    submissions have been removed.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    user_filter = request.GET.get('user', 'all')
    category_filter = request.GET.get('category', 'all')
    start_date = request.GET.get('start_date', '')
    end_date = request.GET.get('end_date', '')

    # Show Excel uploads (data containers)
    uploads = Report.objects.exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).filter(
        status=UPLOADED_STATUS
    ).order_by('-created_at')

    if user_filter != 'all':
        try:
            uploads = uploads.filter(created_by_id=int(user_filter))
        except ValueError:
            pass

    if category_filter != 'all':
        uploads = uploads.filter(report_type=category_filter)

    if start_date:
        try:
            start_d = datetime.strptime(start_date, '%Y-%m-%d').date()
            uploads = uploads.filter(created_at__date__gte=start_d)
        except ValueError:
            pass

    if end_date:
        try:
            end_d = datetime.strptime(end_date, '%Y-%m-%d').date()
            uploads = uploads.filter(created_at__date__lte=end_d)
        except ValueError:
            pass

    users = UserProfile.objects.filter(role='member', status='active').order_by('full_name')
    categories = Report.objects.exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).filter(
        status=UPLOADED_STATUS
    ).values_list('report_type', flat=True).distinct()

    report_data = []
    for report in uploads:
        record_count = report.exception_records.count()
        report_data.append({
            'report': report,
            'created_by': report.created_by,
            'report_type': report.report_type,
            'status': report.status,
            'status_display': report.get_status_display(),
            'submitted_at': report.created_at,
            'record_count': record_count,
        })

    context = {
        'user_profile': user_profile,
        'report_data': report_data,
        'users': users,
        'categories': categories,
        'user_filter': user_filter,
        'category_filter': category_filter,
        'start_date': start_date,
        'end_date': end_date,
        'total_reports': len(report_data),
    }

    return render(request, 'control_dashboard/submitted.html', context)

@login_required
def ad_hoc_scorecard(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    today = timezone.now().date()

    user_filter = request.GET.get('user', 'all')
    users = UserProfile.objects.filter(role='member', status='active').order_by('full_name')

    try:
        deductions = AdHocDeduction.objects.select_related('user', 'created_by').all().order_by('-created_at')

        if user_filter != 'all':
            try:
                deductions = deductions.filter(user_id=int(user_filter))
            except ValueError:
                pass

        deduction_data = []
        total_points = 0
        total_points_added = 0

        for d in deductions:
            total_points += d.points
            total_points_added += (d.points_added or 0)
            deduction_data.append({
                'id': d.id,
                'user_name': d.user.full_name,
                'user_email': d.user.email,
                'task_description': d.task_description,
                'points': d.points,
                'points_added': d.points_added or 0,
                'reason': d.reason,
                'created_at': d.created_at,
                'badge_class': d.get_badge_class(),
            })
    except Exception as e:
        logger.exception("Error loading deductions")
        deduction_data = []
        total_points = 0
        total_points_added = 0

    points_added_today = (
        AdHocDeduction.objects
        .filter(created_at__date=today)
        .aggregate(total=Sum('points_added'))
        .get('total') or 0
    )

    context = {
        'user_profile': user_profile,
        'users': users,
        'user_filter': user_filter,
        'deductions': deduction_data,
        'total_entries': len(deduction_data),
        'total_points': total_points,
        'total_points_added': total_points_added,
        'points_added_today': points_added_today,
    }

    return render(request, 'control_dashboard/ad-hoc.html', context)


@csrf_exempt
@require_http_methods(["POST"])
def api_create_ad_hoc_deduction(request):
    try:
        data = json.loads(request.body)

        user_id = data.get('user_id')
        task_description = data.get('task_description', '').strip()
        points = data.get('points', 0)
        points_added = data.get('points_added', 0)
        reason = data.get('reason', '').strip()

        if not user_id:
            return JsonResponse({'success': False, 'error': 'User is required'}, status=400)

        if not task_description:
            return JsonResponse({'success': False, 'error': 'Task description is required'}, status=400)

        try:
            points = int(points) if str(points).strip() else 0
        except (ValueError, TypeError):
            return JsonResponse({'success': False, 'error': 'Invalid points value'}, status=400)

        try:
            points_added = int(points_added) if str(points_added).strip() else 0
        except (ValueError, TypeError):
            return JsonResponse({'success': False, 'error': 'Invalid points_added value'}, status=400)

        if points < 0 or points > 100:
            return JsonResponse({'success': False, 'error': 'Deducted points must be between 0 and 100'}, status=400)

        if points_added < 0 or points_added > 100:
            return JsonResponse({'success': False, 'error': 'Added points must be between 0 and 100'}, status=400)

        if points == 0 and points_added == 0:
            return JsonResponse({'success': False, 'error': 'Enter either deducted or added points'}, status=400)

        try:
            user = UserProfile.objects.get(id=user_id)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        try:
            created_by = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'Creator not found'}, status=404)

        deduction = AdHocDeduction.objects.create(
            user=user,
            task_description=task_description,
            points=points,
            points_added=points_added,
            reason=reason,
            created_by=created_by
        )

        log_activity(
            user=created_by,
            activity_type='deduction_created',
            details=(f'Created score for {user.full_name} — '
                     f'deducted {points}, added {points_added} — {task_description}'),
            request=request
        )

        return JsonResponse({'success': True, 'message': 'Entry saved successfully', 'deduction_id': deduction.id})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error in api_create_ad_hoc_deduction")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["POST"])
def api_update_ad_hoc_deduction(request, deduction_id):
    try:
        deduction = get_object_or_404(AdHocDeduction, id=deduction_id)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if deduction.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        data = json.loads(request.body)

        if 'task_description' in data:
            deduction.task_description = data['task_description'].strip()

        if 'points' in data:
            try:
                points = int(data['points'])
                if points < 0 or points > 100:
                    return JsonResponse({'success': False, 'error': 'Deducted points must be between 0 and 100'}, status=400)
                deduction.points = points
            except (ValueError, TypeError):
                return JsonResponse({'success': False, 'error': 'Invalid points value'}, status=400)

        if 'points_added' in data:
            try:
                points_added = int(data['points_added'])
                if points_added < 0 or points_added > 100:
                    return JsonResponse({'success': False, 'error': 'Added points must be between 0 and 100'}, status=400)
                deduction.points_added = points_added
            except (ValueError, TypeError):
                return JsonResponse({'success': False, 'error': 'Invalid points_added value'}, status=400)

        if 'reason' in data:
            deduction.reason = data['reason'].strip()

        deduction.save()

        log_activity(
            user=user_profile,
            activity_type='deduction_updated',
            details=f'Updated score for {deduction.user.full_name} — {deduction.task_description}',
            request=request
        )

        return JsonResponse({'success': True, 'message': 'Entry updated successfully'})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error in api_update_ad_hoc_deduction")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_ad_hoc_deduction(request, deduction_id):
    try:
        deduction = get_object_or_404(AdHocDeduction, id=deduction_id)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if deduction.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        log_activity(
            user=user_profile,
            activity_type='deduction_deleted',
            details=f'Deleted deduction for {deduction.user.full_name} - {deduction.task_description}',
            request=request
        )

        deduction.delete()

        return JsonResponse({'success': True, 'message': 'Deduction deleted successfully'})

    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

# control_dashboard/views.py — VERIFY/REPLACE logged_exceptions

@login_required
def logged_exceptions(request):
    """
    Logged Exceptions (supervisor view).

    Shows Excel data containers (status='uploaded') that have at least
    one typed ExceptionRecord attached.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    type_filter = request.GET.get('type', 'all')
    status_filter = request.GET.get('status', 'all')
    start_date = request.GET.get('start_date', '')
    end_date = request.GET.get('end_date', '')

    # Only show Excel uploads that actually carry exception records
    exceptions = Report.objects.exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).filter(
        status=UPLOADED_STATUS,
        exception_records__isnull=False,
    ).distinct().order_by('-created_at')

    if type_filter != 'all':
        exceptions = exceptions.filter(report_type=type_filter)

    if status_filter != 'all':
        exceptions = exceptions.filter(status=status_filter)

    if start_date:
        try:
            start_d = datetime.strptime(start_date, '%Y-%m-%d').date()
            exceptions = exceptions.filter(created_at__date__gte=start_d)
        except ValueError:
            pass

    if end_date:
        try:
            end_d = datetime.strptime(end_date, '%Y-%m-%d').date()
            exceptions = exceptions.filter(created_at__date__lte=end_d)
        except ValueError:
            pass

    report_types = Report.objects.exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).filter(
        status=UPLOADED_STATUS,
        exception_records__isnull=False,
    ).values_list('report_type', flat=True).distinct()

    context = {
        'user_profile': user_profile,
        'exceptions': exceptions,
        'report_types': report_types,
        'type_filter': type_filter,
        'status_filter': status_filter,
        'start_date': start_date,
        'end_date': end_date,
    }

    return render(request, 'control_dashboard/logged.html', context)
@login_required
def supervisor_checklist(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    today = timezone.now().date()
    month_start = today.replace(day=1)
    year_start = today.replace(month=1, day=1)

    all_checklists = Checklist.objects.filter(is_active=True)

    team_members = (
        UserProfile.objects
        .filter(role='member', status='active')
        .prefetch_related('branches', 'departments')
        .order_by('full_name')
    )

    user_completion_data = []
    total_users = team_members.count()
    total_completed = 0
    total_assigned = 0
    sum_of_rates = 0

    for member in team_members:
        member_checklists = all_checklists.filter(
            Q(assigned_users=member) |
            Q(assignment_target='all') |
            Q(assignment_target=member.position)
        ).distinct()

        member_branch_ids = set(member.branches.values_list('id', flat=True))
        member_dept_ids = set(member.departments.values_list('id', flat=True))

        member_expected_units = 0
        member_completed_units = 0
        member_ytd_expected = 0
        member_ytd_completed = 0

        per_branch = {}
        per_department = {}

        for checklist in member_checklists:
            ck_branches = checklist.assigned_branches.filter(id__in=member_branch_ids)
            ck_departments = checklist.assigned_departments.filter(id__in=member_dept_ids)

            if not checklist.assigned_branches.exists() and not checklist.assigned_departments.exists():
                exp = count_expected_occurrences(checklist, month_start, today)
                act = ChecklistLog.objects.filter(
                    checklist=checklist, user=member,
                    log_date__gte=month_start, log_date__lte=today,
                    branch__isnull=True, department__isnull=True,
                ).values('log_date').distinct().count()

                member_expected_units += exp
                member_completed_units += min(act, exp) if exp else 0
                member_ytd_expected += count_expected_occurrences(checklist, year_start, today)
                member_ytd_completed += min(
                    ChecklistLog.objects.filter(
                        checklist=checklist, user=member,
                        log_date__gte=year_start, log_date__lte=today,
                        branch__isnull=True, department__isnull=True,
                    ).values('log_date').distinct().count(),
                    count_expected_occurrences(checklist, year_start, today) or 0,
                )
                continue

            for branch in ck_branches:
                exp = count_expected_occurrences(checklist, month_start, today)
                act = ChecklistLog.objects.filter(
                    checklist=checklist, user=member,
                    log_date__gte=month_start, log_date__lte=today,
                    branch_id=branch.id, department__isnull=True,
                ).values('log_date').distinct().count()

                act = min(act, exp) if exp else 0
                member_expected_units += exp
                member_completed_units += act

                row = per_branch.setdefault(branch.id, {'id': branch.id, 'name': branch.name, 'expected': 0, 'actual': 0})
                row['expected'] += exp
                row['actual'] += act

                y_exp = count_expected_occurrences(checklist, year_start, today)
                y_act = ChecklistLog.objects.filter(
                    checklist=checklist, user=member,
                    log_date__gte=year_start, log_date__lte=today,
                    branch_id=branch.id, department__isnull=True,
                ).values('log_date').distinct().count()
                member_ytd_expected += y_exp
                member_ytd_completed += min(y_act, y_exp) if y_exp else 0

            for dept in ck_departments:
                exp = count_expected_occurrences(checklist, month_start, today)
                act = ChecklistLog.objects.filter(
                    checklist=checklist, user=member,
                    log_date__gte=month_start, log_date__lte=today,
                    department_id=dept.id, branch__isnull=True,
                ).values('log_date').distinct().count()

                act = min(act, exp) if exp else 0
                member_expected_units += exp
                member_completed_units += act

                row = per_department.setdefault(dept.id, {'id': dept.id, 'name': dept.name, 'expected': 0, 'actual': 0})
                row['expected'] += exp
                row['actual'] += act

                y_exp = count_expected_occurrences(checklist, year_start, today)
                y_act = ChecklistLog.objects.filter(
                    checklist=checklist, user=member,
                    log_date__gte=year_start, log_date__lte=today,
                    department_id=dept.id, branch__isnull=True,
                ).values('log_date').distinct().count()
                member_ytd_expected += y_exp
                member_ytd_completed += min(y_act, y_exp) if y_exp else 0

        completion_rate = int(member_completed_units / member_expected_units * 100) if member_expected_units else 0
        completion_rate = min(completion_rate, 100)
        ytd_rate = int(member_ytd_completed / member_ytd_expected * 100) if member_ytd_expected else 0
        ytd_rate = min(ytd_rate, 100)

        for row in per_branch.values():
            row['progress'] = int(row['actual'] / row['expected'] * 100) if row['expected'] else 0
            row['progress'] = min(row['progress'], 100)
        for row in per_department.values():
            row['progress'] = int(row['actual'] / row['expected'] * 100) if row['expected'] else 0
            row['progress'] = min(row['progress'], 100)

        branch_rows = sorted(per_branch.values(), key=lambda r: r['name'].lower())
        dept_rows = sorted(per_department.values(), key=lambda r: r['name'].lower())

        if completion_rate >= 90:
            status = 'success'
            status_text = 'Excellent'
        elif completion_rate >= 70:
            status = 'warning'
            status_text = 'Good'
        elif completion_rate >= 40:
            status = 'warning'
            status_text = 'In Progress'
        else:
            status = 'danger'
            status_text = 'Needs Improvement'

        total_completed += member_completed_units
        total_assigned += member_expected_units
        sum_of_rates += completion_rate

        user_completion_data.append({
            'user': member,
            'total_checklists': member_expected_units,
            'completed_checklists': member_completed_units,
            'completion_rate': completion_rate,
            'ytd_completion_rate': ytd_rate,
            'branch_count': len(branch_rows),
            'department_count': len(dept_rows),
            'branches': branch_rows,
            'departments': dept_rows,
            'status': status,
            'status_text': status_text,
        })

    avg_completion = int(sum_of_rates / total_users) if total_users else 0

    user_completion_data.sort(key=lambda x: x['completion_rate'], reverse=True)

    context = {
        'user_profile': user_profile,
        'user_completion_data': user_completion_data,
        'total_users': total_users,
        'avg_completion': avg_completion,
        'total_checklists': all_checklists.count(),
        'total_completed': total_completed,
        'total_assigned': total_assigned,
    }

    return render(request, 'control_dashboard/checklist-sup.html', context)


@csrf_exempt
@require_http_methods(["GET"])
def api_checklist_detail(request, user_id):
    try:
        try:
            user = UserProfile.objects.get(id=user_id)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        today = timezone.now().date()
        month_start = today.replace(day=1)
        year_start = today.replace(month=1, day=1)

        member_checklists = Checklist.objects.filter(
            is_active=True
        ).filter(
            Q(assigned_users=user) |
            Q(assignment_target='all') |
            Q(assignment_target=user.position)
        ).distinct()

        member_branch_ids = set(user.branches.values_list('id', flat=True))
        member_dept_ids = set(user.departments.values_list('id', flat=True))

        overall_expected = 0
        overall_actual = 0
        overall_ytd_expected = 0
        overall_ytd_actual = 0

        branches = {}
        departments = {}

        for checklist in member_checklists:
            ck_branches = checklist.assigned_branches.filter(id__in=member_branch_ids)
            ck_departments = checklist.assigned_departments.filter(id__in=member_dept_ids)

            if not checklist.assigned_branches.exists() and not checklist.assigned_departments.exists():
                exp = count_expected_occurrences(checklist, month_start, today)
                act = ChecklistLog.objects.filter(
                    checklist=checklist, user=user,
                    log_date__gte=month_start, log_date__lte=today,
                    branch__isnull=True, department__isnull=True,
                ).values('log_date').distinct().count()
                act = min(act, exp) if exp else 0
                overall_expected += exp
                overall_actual += act

                y_exp = count_expected_occurrences(checklist, year_start, today)
                y_act = ChecklistLog.objects.filter(
                    checklist=checklist, user=user,
                    log_date__gte=year_start, log_date__lte=today,
                    branch__isnull=True, department__isnull=True,
                ).values('log_date').distinct().count()
                overall_ytd_expected += y_exp
                overall_ytd_actual += min(y_act, y_exp) if y_exp else 0
                continue

            for branch in ck_branches:
                exp = count_expected_occurrences(checklist, month_start, today)
                act = ChecklistLog.objects.filter(
                    checklist=checklist, user=user,
                    log_date__gte=month_start, log_date__lte=today,
                    branch_id=branch.id, department__isnull=True,
                ).values('log_date').distinct().count()
                act = min(act, exp) if exp else 0
                overall_expected += exp
                overall_actual += act

                row = branches.setdefault(branch.id, {'id': branch.id, 'name': branch.name, 'code': branch.code or '', 'expected': 0, 'actual': 0})
                row['expected'] += exp
                row['actual'] += act

                y_exp = count_expected_occurrences(checklist, year_start, today)
                y_act = ChecklistLog.objects.filter(
                    checklist=checklist, user=user,
                    log_date__gte=year_start, log_date__lte=today,
                    branch_id=branch.id, department__isnull=True,
                ).values('log_date').distinct().count()
                overall_ytd_expected += y_exp
                overall_ytd_actual += min(y_act, y_exp) if y_exp else 0

            for dept in ck_departments:
                exp = count_expected_occurrences(checklist, month_start, today)
                act = ChecklistLog.objects.filter(
                    checklist=checklist, user=user,
                    log_date__gte=month_start, log_date__lte=today,
                    department_id=dept.id, branch__isnull=True,
                ).values('log_date').distinct().count()
                act = min(act, exp) if exp else 0
                overall_expected += exp
                overall_actual += act

                row = departments.setdefault(dept.id, {'id': dept.id, 'name': dept.name, 'code': dept.code or '', 'expected': 0, 'actual': 0})
                row['expected'] += exp
                row['actual'] += act

                y_exp = count_expected_occurrences(checklist, year_start, today)
                y_act = ChecklistLog.objects.filter(
                    checklist=checklist, user=user,
                    log_date__gte=year_start, log_date__lte=today,
                    department_id=dept.id, branch__isnull=True,
                ).values('log_date').distinct().count()
                overall_ytd_expected += y_exp
                overall_ytd_actual += min(y_act, y_exp) if y_exp else 0

        for row in branches.values():
            row['progress'] = int(row['actual'] / row['expected'] * 100) if row['expected'] else 0
            row['progress'] = min(row['progress'], 100)
        for row in departments.values():
            row['progress'] = int(row['actual'] / row['expected'] * 100) if row['expected'] else 0
            row['progress'] = min(row['progress'], 100)

        checklist_details = []
        for checklist in member_checklists:
            tasks = checklist.tasks.all().order_by('order')
            total_tasks = tasks.count()
            is_completed_today = ChecklistLog.objects.filter(checklist=checklist, user=user, log_date=today).exists()

            task_list = []
            completed_tasks = 0
            for task in tasks:
                task_list.append({'description': task.description, 'is_completed': is_completed_today})
                if is_completed_today:
                    completed_tasks += 1

            rate = int(completed_tasks / total_tasks * 100) if total_tasks else 0

            checklist_details.append({
                'id': checklist.id,
                'name': checklist.name,
                'frequency': checklist.get_frequency_display(),
                'total_tasks': total_tasks,
                'completed_tasks': completed_tasks,
                'task_completion_rate': rate,
                'is_completed': is_completed_today,
                'tasks': task_list,
            })

        overall_rate = int(overall_actual / overall_expected * 100) if overall_expected else 0
        overall_rate = min(overall_rate, 100)
        ytd_rate = int(overall_ytd_actual / overall_ytd_expected * 100) if overall_ytd_expected else 0
        ytd_rate = min(ytd_rate, 100)

        return JsonResponse({
            'success': True,
            'user': {
                'id': user.id,
                'full_name': user.full_name,
                'email': user.email,
                'username': user.username,
                'position': user.get_position_display() or 'Member',
                'completion_rate': overall_rate,
                'ytd_completion_rate': ytd_rate,
                'month_expected': overall_expected,
                'month_actual': overall_actual,
                'ytd_expected': overall_ytd_expected,
                'ytd_actual': overall_ytd_actual,
                'branches': sorted(branches.values(), key=lambda r: r['name'].lower()),
                'departments': sorted(departments.values(), key=lambda r: r['name'].lower()),
                'checklist_details': checklist_details,
            }
        })

    except Exception as e:
        logger.exception("Error in api_checklist_detail")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@login_required
def supervisor_activity_logs(request):
    """
    Activity Logs (supervisor view).

    Supervisors see ALL users' activity logs — no role scoping,
    no per-user filtering by default. The filter dropdown lets
    them narrow to a single user if desired.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    # ------------------------------------------------------------
    # Query params
    # ------------------------------------------------------------
    user_filter     = request.GET.get('user', 'all').strip()
    start_date      = request.GET.get('start_date', '').strip()
    end_date        = request.GET.get('end_date', '').strip()
    activity_filter = request.GET.get('activity', '').strip()

    # ------------------------------------------------------------
    # Base queryset — ALL users' activity logs
    # ------------------------------------------------------------
    queryset = ActivityLog.objects.select_related('user').all()

    # Optional: scope to a specific user
    if user_filter and user_filter != 'all':
        try:
            queryset = queryset.filter(user_id=int(user_filter))
        except ValueError:
            pass

    # Optional: date range
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__gte=start_date_obj)
        except ValueError:
            pass

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__lte=end_date_obj)
        except ValueError:
            pass

    # Optional: activity type
    if activity_filter and activity_filter != 'all':
        queryset = queryset.filter(activity_type=activity_filter)

    # ------------------------------------------------------------
    # Stats for the header (computed BEFORE slicing)
    # ------------------------------------------------------------
    total_logs    = queryset.count()
    unique_users  = queryset.values('user_id').distinct().count()

    today = timezone.now().date()
    today_logs = queryset.filter(created_at__date=today).count()

    # ------------------------------------------------------------
    # Final result — newest first, capped at 200 rows
    # ------------------------------------------------------------
    activity_logs = queryset.order_by('-created_at')[:200]

    # ------------------------------------------------------------
    # User list for the filter dropdown
    # Show every user that has ever logged activity, PLUS any
    # active members. This avoids a supervisor not being able to
    # filter by a user who was recently deactivated.
    # ------------------------------------------------------------
    user_ids_with_logs = (
        ActivityLog.objects
        .values_list('user_id', flat=True)
        .distinct()
    )
    users = (
        UserProfile.objects
        .filter(Q(id__in=user_ids_with_logs) | Q(status='active'))
        .distinct()
        .order_by('full_name')
    )

    activity_types = ActivityLog.ACTIVITY_TYPES

    context = {
        'user_profile': user_profile,
        'activity_logs': activity_logs,
        'users': users,
        'activity_types': activity_types,
        'user_filter': user_filter,
        'activity_filter': activity_filter,
        'start_date': start_date,
        'end_date': end_date,
        # Stats
        'total_logs': total_logs,
        'unique_users': unique_users,
        'today_logs': today_logs,
        'today': today,
    }

    return render(request, 'control_dashboard/activity-sup.html', context)

@login_required
def activity_logs(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    user_filter = request.GET.get('user', 'all')
    start_date = request.GET.get('start_date', '')
    end_date = request.GET.get('end_date', '')
    activity_filter = request.GET.get('activity', '')

    queryset = ActivityLog.objects.all().select_related('user')

    if user_filter != 'all':
        try:
            queryset = queryset.filter(user_id=int(user_filter))
        except ValueError:
            pass

    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__gte=start_date_obj)
        except ValueError:
            pass

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, '%Y-%m-%d').date()
            queryset = queryset.filter(created_at__date__lte=end_date_obj)
        except ValueError:
            pass

    if activity_filter and activity_filter != 'all':
        queryset = queryset.filter(activity_type=activity_filter)

    unique_users = queryset.values('user_id').distinct().count()
    today = timezone.now().date()
    today_logs = queryset.filter(created_at__date=today).count()

    activity_breakdown = queryset.values('activity_type').annotate(count=Count('id')).order_by('-count')

    user_activity_summary = queryset.values(
        'user__full_name', 'user__email', 'user__role'
    ).annotate(total_activities=Count('id')).order_by('-total_activities')[:10]

    last_activity_log = queryset.order_by('-created_at').first()
    last_activity = last_activity_log.created_at.strftime('%b %d, %Y %H:%M') if last_activity_log else None

    activity_logs = queryset.order_by('-created_at')[:200]

    users = UserProfile.objects.filter(status='active').order_by('full_name')
    activity_types = ActivityLog.ACTIVITY_TYPES

    context = {
        'user_profile': user_profile,
        'activity_logs': activity_logs,
        'users': users,
        'activity_types': activity_types,
        'user_filter': user_filter,
        'activity_filter': activity_filter,
        'start_date': start_date,
        'end_date': end_date,
        'unique_users': unique_users,
        'today_logs': today_logs,
        'last_activity': last_activity,
        'activity_breakdown': activity_breakdown,
        'user_activity_summary': user_activity_summary,
    }

    return render(request, 'control_dashboard/activity.html', context)


# ==================== ANALYTICS DASHBOARD ====================

# ==================== ANALYTICS DASHBOARD ====================

@login_required
def analytics_dashboard(request):
    """
    Analytics Dashboard — feeds directly from ExceptionRecord.

    Role scoping:
      admin / supervisor → see ALL exception records
      member             → see only records in Reports they created

    Filters (query params, all optional):
      branch, month, year, start_date, end_date, category, status

    Template dispatch:
      admin / supervisor → analytics-sup.html
      member             → analytics.html
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    if user_profile.role not in ('admin', 'supervisor', 'member'):
        messages.error(request, 'You do not have permission to access this page.')
        return redirect_dashboard(request.user)

    from decimal import Decimal

    today = timezone.now().date()
    is_privileged = user_profile.role in ('admin', 'supervisor')

    # ============================================================
    # BASE QUERYSET — all exception records the user can see
    # ============================================================
    base_qs = ExceptionRecord.objects.filter(
        report__isnull=False,
    ).exclude(
        report__report_type=TRIAL_BALANCE_REPORT_TYPE,
    )

    if not is_privileged:
        base_qs = base_qs.filter(report__created_by=user_profile)

    # ============================================================
    # FILTERS
    # ============================================================
    branch_filter   = request.GET.get('branch', 'all').strip()
    month_filter    = request.GET.get('month', 'all').strip()
    year_filter     = request.GET.get('year', 'all').strip()
    start_date_str  = request.GET.get('start_date', '').strip()
    end_date_str    = request.GET.get('end_date', '').strip()
    category_filter = request.GET.get('category', 'all').strip()
    status_filter   = request.GET.get('status', 'all').strip()

    filtered_qs = base_qs

    if branch_filter and branch_filter != 'all':
        filtered_qs = filtered_qs.filter(branch_unit=branch_filter)

    if category_filter and category_filter != 'all':
        filtered_qs = filtered_qs.filter(category=category_filter)

    if status_filter and status_filter != 'all':
        filtered_qs = filtered_qs.filter(status=status_filter)

    if year_filter and year_filter != 'all':
        try:
            y = int(year_filter)
            filtered_qs = filtered_qs.filter(date_noted__year=y)
        except ValueError:
            pass

    if month_filter and month_filter != 'all':
        try:
            m = int(month_filter)  # 1-12 expected from the front-end
            filtered_qs = filtered_qs.filter(date_noted__month=m)
        except ValueError:
            pass

    if start_date_str:
        try:
            start_d = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            filtered_qs = filtered_qs.filter(date_noted__gte=start_d)
        except ValueError:
            pass

    if end_date_str:
        try:
            end_d = datetime.strptime(end_date_str, '%Y-%m-%d').date()
            filtered_qs = filtered_qs.filter(date_noted__lte=end_d)
        except ValueError:
            pass

    # ============================================================
    # KPI 1 — TOTAL EXCEPTIONS
    # ============================================================
    total_exceptions = filtered_qs.count()

    # ============================================================
    # KPI 2 / 3 — OPEN vs CLOSED
    # Treat 'closed' and 'resolved' as CLOSED; everything else as OPEN
    # ============================================================
    OPEN_STATUSES   = ['open', 'in_progress', 'pending', 'overdue', 'rejected']
    CLOSED_STATUSES = ['closed', 'resolved']

    open_exceptions   = filtered_qs.filter(status__in=OPEN_STATUSES).count()
    closed_exceptions = filtered_qs.filter(status__in=CLOSED_STATUSES).count()

    # ============================================================
    # KPI 4 — RESOLUTION RATE
    # ============================================================
    resolution_pct = (
        int(round((closed_exceptions / total_exceptions) * 100))
        if total_exceptions else 0
    )

    # ============================================================
    # KPI 5 — COST SAVED
    # ============================================================
    cost_saved_total = (
        filtered_qs
        .aggregate(total=Sum('income_cost_saved'))
        .get('total') or Decimal('0')
    )

    # ============================================================
    # PIE 1 — TOP 3 BRANCHES / DEPARTMENTS
    # ============================================================
    branch_counter = (
        filtered_qs
        .exclude(branch_unit='')
        .values('branch_unit')
        .annotate(count=Count('id'))
        .order_by('-count')[:3]
    )
    top_branch_labels = [b['branch_unit'] for b in branch_counter]
    top_branch_values = [b['count'] for b in branch_counter]

    # ============================================================
    # PIE 2 — TOP 3 CATEGORIES
    # ============================================================
    category_counter = (
        filtered_qs
        .exclude(category='')
        .values('category')
        .annotate(count=Count('id'))
        .order_by('-count')[:3]
    )
    top_category_labels = [c['category'] for c in category_counter]
    top_category_values = [c['count'] for c in category_counter]

    # ============================================================
    # OPEN EXCEPTIONS TABLE — full detail, open items only
    # ============================================================
    open_exceptions_list = (
        filtered_qs
        .filter(status__in=OPEN_STATUSES)
        .order_by('-date_noted', '-created_at')[:100]
    )

    # ============================================================
    # FILTER DROPDOWN DATA
    # Built from the UNFILTERED base so users can always switch
    # ============================================================
    available_branches = list(
        base_qs
        .exclude(branch_unit='')
        .values_list('branch_unit', flat=True)
        .distinct()
        .order_by('branch_unit')
    )

    available_categories = list(
        base_qs
        .exclude(category='')
        .values_list('category', flat=True)
        .distinct()
        .order_by('category')
    )

    available_years = sorted(
        {
            d.year
            for d in base_qs
            .exclude(date_noted__isnull=True)
            .values_list('date_noted', flat=True)
        },
        reverse=True,
    )

    month_names = [
        (1, 'January'), (2, 'February'), (3, 'March'), (4, 'April'),
        (5, 'May'), (6, 'June'), (7, 'July'), (8, 'August'),
        (9, 'September'), (10, 'October'), (11, 'November'), (12, 'December'),
    ]

    available_statuses = [
        ('open', 'Open'),
        ('in_progress', 'In Progress'),
        ('pending', 'Pending'),
        ('overdue', 'Overdue'),
        ('closed', 'Closed'),
        ('resolved', 'Resolved'),
    ]

    # ============================================================
    # CONTEXT
    # ============================================================
    context = {
        'user_profile': user_profile,
        'today': today,
        'is_privileged': is_privileged,

        # --- KPI cards ---
        'total_exceptions': total_exceptions,
        'open_exceptions': open_exceptions,
        'closed_exceptions': closed_exceptions,
        'resolution_pct': resolution_pct,
        'cost_saved_total': f'{cost_saved_total:,.2f}',

        # --- Pie charts ---
        'top_branch_labels': json.dumps(top_branch_labels),
        'top_branch_values': json.dumps(top_branch_values),
        'top_category_labels': json.dumps(top_category_labels),
        'top_category_values': json.dumps(top_category_values),
        'open_closed_labels': json.dumps(['Open', 'Closed']),
        'open_closed_values': json.dumps([open_exceptions, closed_exceptions]),

        # --- Open table ---
        'open_exceptions_list': open_exceptions_list,

        # --- Filter options + current values ---
        'available_branches': available_branches,
        'available_categories': available_categories,
        'available_years': available_years,
        'available_statuses': available_statuses,
        'month_names': month_names,
        'branch_filter': branch_filter,
        'category_filter': category_filter,
        'status_filter': status_filter,
        'year_filter': year_filter,
        'month_filter': month_filter,
        'start_date': start_date_str,
        'end_date': end_date_str,
    }

    # ------------------------------------------------------------
    # Template dispatch — supervisors/admins get the supervisor
    # skin; members get the standard member skin. Both read the
    # exact same context, so any future field additions only need
    # to be made once.
    # ------------------------------------------------------------
    template = (
        'control_dashboard/analytics-sup.html'
        if is_privileged
        else 'control_dashboard/analytics.html'
    )

    return render(request, template, context)
# ==================== SUBMIT SELECTED REPORTS ====================

@login_required
@csrf_exempt
def submit_selected_reports(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    if request.method == 'POST':
        try:
            selected_ids = json.loads(request.POST.get('selected_reports', '[]'))
            if not selected_ids:
                messages.error(request, 'No reports selected.')
                return redirect('control_dashboard:reports_page')

            reports = Report.objects.filter(
                id__in=selected_ids,
                created_by=user_profile
            ).exclude(report_type=TRIAL_BALANCE_REPORT_TYPE)

            if not reports.exists():
                messages.error(request, 'Selected reports not found.')
                return redirect('control_dashboard:reports_page')

            excel_data = generate_excel_from_reports(reports)
            screenshot_data = generate_exception_screenshot(reports)

            request.session['submission_data'] = {
                'reports': [{'id': r.id, 'report_type': r.report_type} for r in reports],
                'excel_data': excel_data,
                'screenshot_data': screenshot_data,
                'report_count': reports.count()
            }

            return redirect('control_dashboard:submit_page')

        except Exception as e:
            logger.exception("Error submitting reports")
            messages.error(request, f'Error preparing submission: {str(e)}')
            return redirect('control_dashboard:reports_page')

    return redirect('control_dashboard:submit_page')


def generate_excel_from_reports(reports):
    wb = openpyxl.Workbook()

    for idx, report in enumerate(reports):
        ws = wb.create_sheet(title=f"Report_{idx+1}_{report.report_type[:20]}")

        display_data = report.get_display_data()
        headers = display_data.get('headers', [])
        rows = display_data.get('rows', [])

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")

        for row_idx, row in enumerate(rows, 2):
            if isinstance(row, dict):
                for col_idx, header in enumerate(headers, 1):
                    value = row.get(header, '')
                    ws.cell(row=row_idx, column=col_idx, value=value)
            elif isinstance(row, list):
                for col_idx, value in enumerate(row, 1):
                    if col_idx <= len(headers):
                        ws.cell(row=row_idx, column=col_idx, value=value)

        for col in range(1, len(headers) + 1):
            column_letter = get_column_letter(col)
            max_length = 15
            for row in range(1, min(ws.max_row + 1, 50)):
                cell_value = ws.cell(row=row, column=col).value
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            ws.column_dimensions[column_letter].width = min(max_length + 2, 50)

    wb.remove(wb['Sheet'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return {
        'file_name': f'Exception_Report_{timezone.now().strftime("%Y%m%d")}.xlsx',
        'file_data': base64.b64encode(output.getvalue()).decode('utf-8'),
        'report_count': reports.count()
    }


def generate_exception_screenshot(reports):
    import html

    html_content = """
    <html>
    <head>
        <style>
            body { font-family: Arial, sans-serif; padding: 20px; background: white; }
            .report-container { margin-bottom: 30px; border: 1px solid #e5e7eb; border-radius: 8px; padding: 15px; }
            .report-title { font-size: 16px; font-weight: 600; color: #1a2332; margin-bottom: 10px; border-bottom: 2px solid #0066cc; padding-bottom: 8px; }
            table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 10px; }
            th { background: #0066cc; color: white; padding: 8px 12px; text-align: left; font-weight: 600; }
            td { padding: 6px 12px; border-bottom: 1px solid #e5e7eb; }
            tr:hover { background: #f8fafc; }
            .badge { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 10px; font-weight: 600; }
            .badge-success { background: #d1fae5; color: #065f46; }
            .badge-warning { background: #fef3c7; color: #92400e; }
            .badge-danger { background: #fee2e2; color: #991b1b; }
            .badge-secondary { background: #f3f4f6; color: #6b7280; }
            .excel-note { background: #f0f7ff; padding: 10px 15px; border-radius: 6px; border: 1px solid #0066cc; margin-top: 10px; font-size: 13px; color: #1a2332; }
        </style>
    </head>
    <body>
    """

    for report in reports:
        display_data = report.get_display_data()
        headers = display_data.get('headers', [])
        rows = display_data.get('rows', [])
        is_excel = display_data.get('is_excel', False)

        html_content += f"""
        <div class="report-container">
            <div class="report-title">
                📊 {html.escape(report.report_type)}
                <span style="font-size:12px;font-weight:400;color:#6b7280;margin-left:10px;">
                    ({len(rows)} records)
                </span>
            </div>
        """

        if rows and headers:
            html_content += "<table>"
            html_content += "<thead><tr>"
            for header in headers:
                html_content += f"<th>{html.escape(str(header))}</th>"
            html_content += "</tr></thead>"

            html_content += "<tbody>"
            for row in rows[:20]:
                html_content += "<tr>"
                if isinstance(row, dict):
                    for header in headers:
                        value = row.get(header, '')
                        if 'status' in header.lower():
                            status_class = get_status_class(value)
                            html_content += f"<td><span class='badge {status_class}'>{html.escape(str(value))}</span></td>"
                        else:
                            html_content += f"<td>{html.escape(str(value))}</td>"
                elif isinstance(row, list):
                    for value in row[:len(headers)]:
                        html_content += f"<td>{html.escape(str(value))}</td>"
                html_content += "</tr>"

            if len(rows) > 20:
                html_content += f"<tr><td colspan='{len(headers)}' style='text-align:center;color:#6b7280;font-style:italic;'>... and {len(rows) - 20} more rows</td></tr>"

            html_content += "</tbody></table>"
        else:
            html_content += "<p style='color:#6b7280;font-style:italic;'>No data available</p>"

        if is_excel:
            html_content += """
            <div class="excel-note">
                📎 Excel file attached with all data.
            </div>
            """

        html_content += "</div>"

    html_content += """
    </body>
    </html>
    """

    return {'html': html_content, 'report_count': reports.count()}


def get_status_class(status):
    status_lower = str(status).lower() if status else ''
    if status_lower in ['open', 'pending', 'in progress']:
        return 'badge-warning'
    elif status_lower in ['closed', 'resolved', 'approved', 'completed']:
        return 'badge-success'
    elif status_lower in ['rejected', 'cancelled']:
        return 'badge-danger'
    else:
        return 'badge-secondary'


# ==================== DAILY TRIAL BALANCE ====================

@login_required
def daily_trial_balance(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    context = {
        'user_profile': user_profile,
        'today': timezone.now(),
    }
    return render(request, 'control_dashboard/trialbalance.html', context)


@csrf_exempt
@require_http_methods(["POST"])
def api_parse_trial_balance(request):
    try:
        headers = []
        rows_data = []
        file_name = 'trial_balance.xlsx'

        if request.FILES.get('file'):
            uploaded = request.FILES['file']
            file_name = uploaded.name or file_name
            try:
                wb = openpyxl.load_workbook(io.BytesIO(uploaded.read()), data_only=True)
                ws = wb.active
                all_rows = []
                for row in ws.iter_rows(values_only=True):
                    if row and any(cell is not None and str(cell).strip() != '' for cell in row):
                        all_rows.append([cell if cell is not None else '' for cell in row])
                if not all_rows:
                    return JsonResponse({'success': False, 'error': 'The uploaded file appears to be empty.'}, status=400)
                headers = [str(h).strip() if h is not None else f'Column_{i+1}' for i, h in enumerate(all_rows[0])]
                rows_data = all_rows[1:]
            except Exception as parse_err:
                logger.exception("Excel parse error")
                return JsonResponse({'success': False, 'error': f'Could not parse Excel file: {str(parse_err)}'}, status=400)

        else:
            data = json.loads(request.body)
            file_name = data.get('file_name', file_name)
            file_b64 = data.get('file_data')
            rows_from_client = data.get('rows')

            if rows_from_client and isinstance(rows_from_client, list):
                if len(rows_from_client) > 0:
                    headers = rows_from_client[0]
                    rows_data = rows_from_client[1:]

            elif file_b64:
                try:
                    if ',' in file_b64:
                        file_b64 = file_b64.split(',', 1)[1]
                    raw = base64.b64decode(file_b64)
                    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
                    ws = wb.active
                    all_rows = []
                    for row in ws.iter_rows(values_only=True):
                        if row and any(cell is not None and str(cell).strip() != '' for cell in row):
                            all_rows.append([cell if cell is not None else '' for cell in row])
                    if not all_rows:
                        return JsonResponse({'success': False, 'error': 'The uploaded file appears to be empty.'}, status=400)
                    headers = [str(h).strip() if h is not None else f'Column_{i+1}' for i, h in enumerate(all_rows[0])]
                    rows_data = all_rows[1:]
                except Exception as parse_err:
                    logger.exception("Excel parse error")
                    return JsonResponse({'success': False, 'error': f'Could not parse Excel file: {str(parse_err)}'}, status=400)
            else:
                return JsonResponse({'success': False, 'error': 'No file data provided.'}, status=400)

        if not headers:
            return JsonResponse({'success': False, 'error': 'No header row found in file.'}, status=400)

        parsed_rows = []
        for r in rows_data:
            row_dict = {}
            for i, h in enumerate(headers):
                row_dict[h] = r[i] if i < len(r) else ''
            parsed_rows.append(row_dict)

        def _to_float(v):
            try:
                if v is None or v == '':
                    return 0.0
                s = str(v).replace(',', '').replace('GHS', '').strip()
                if s in ('-', '--'):
                    return 0.0
                return float(s)
            except (ValueError, TypeError):
                return 0.0

        debit_col = None
        credit_col = None
        for h in headers:
            hu = str(h).upper()
            if debit_col is None and ('DEBIT' in hu or hu.startswith('DR_BAL') or hu == 'DR'):
                debit_col = h
            if credit_col is None and ('CREDIT' in hu or hu.startswith('CR_BAL') or hu == 'CR'):
                credit_col = h

        total_debit = sum(_to_float(r.get(debit_col, 0)) for r in parsed_rows) if debit_col else 0.0
        total_credit = sum(_to_float(r.get(credit_col, 0)) for r in parsed_rows) if credit_col else 0.0
        variance = total_debit - total_credit

        return JsonResponse({
            'success': True,
            'file_name': file_name,
            'headers': headers,
            'rows': parsed_rows,
            'row_count': len(parsed_rows),
            'summary': {
                'total_rows': len(parsed_rows),
                'total_debit': round(total_debit, 2),
                'total_credit': round(total_credit, 2),
                'variance': round(variance, 2),
                'debit_column': debit_col,
                'credit_column': credit_col,
            }
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data.'}, status=400)
    except Exception as e:
        logger.exception("Error in api_parse_trial_balance")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["POST"])
def api_submit_trial_balance(request):
    try:
        data = json.loads(request.body)
        file_name = data.get('file_name', 'trial_balance.xlsx')
        headers = data.get('headers', [])
        rows_data = data.get('rows', [])
        report_date_str = data.get('report_date', '')

        if not headers or not rows_data:
            return JsonResponse({'success': False, 'error': 'No data to submit. Please upload a file first.'}, status=400)

        if not report_date_str:
            return JsonResponse({'success': False, 'error': 'Please select a report date before submitting.'}, status=400)

        try:
            report_date = datetime.strptime(report_date_str, '%Y-%m-%d').date()
        except ValueError:
            return JsonResponse({'success': False, 'error': 'Invalid report date format. Use YYYY-MM-DD.'}, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

        from .models import TrialBalanceEntry

        existing_any = TrialBalanceEntry.objects.filter(report_date=report_date).exists()
        if existing_any:
            anchor_report_ids = list(
                TrialBalanceEntry.objects.filter(report_date=report_date)
                .values_list('report_id', flat=True).distinct()
            )
            TrialBalanceEntry.objects.filter(report_date=report_date).delete()
            Report.objects.filter(
                id__in=anchor_report_ids,
                report_type=TRIAL_BALANCE_REPORT_TYPE
            ).delete()

        report = Report.objects.create(
            report_type=TRIAL_BALANCE_REPORT_TYPE,
            frequency='daily',
            description=f'Trial Balance for {report_date.isoformat()} ({file_name})',
            status='submitted',
            created_by=user_profile,
        )

        norm_map = {h: str(h).strip().upper().replace(' ', '_') for h in headers}

        def get_field(row, key):
            for orig_h, norm_h in norm_map.items():
                if norm_h == key:
                    return row.get(orig_h, '')
            return ''

        def to_decimal(v):
            from decimal import Decimal, InvalidOperation
            if v is None or v == '':
                return Decimal('0')
            try:
                s = str(v).replace(',', '').replace('GHS', '').strip()
                if s in ('-', '--', ''):
                    return Decimal('0')
                return Decimal(s)
            except (InvalidOperation, ValueError, TypeError):
                return Decimal('0')

        entries = []
        for idx, row in enumerate(rows_data):
            if not isinstance(row, dict):
                continue

            entry = TrialBalanceEntry(
                report=report,
                uploaded_by=user_profile,
                report_date=report_date,
                file_name=file_name,
                row_index=idx,
                branch_code=str(get_field(row, 'BRANCH_CODE'))[:50],
                today=str(get_field(row, 'TODAY'))[:50],
                category=str(get_field(row, 'CATEGORY'))[:100],
                parent_gl=str(get_field(row, 'PARENT_GL'))[:50],
                gl_code=str(get_field(row, 'GL_CODE'))[:50],
                descr=str(get_field(row, 'DESCR'))[:500],
                ccy=str(get_field(row, 'CCY'))[:10],
                open_bal_lcy=to_decimal(get_field(row, 'OPEN_BAL_LCY')),
                dr_bal_lcy=to_decimal(get_field(row, 'DR_BAL_LCY')),
                cr_bal_lcy=to_decimal(get_field(row, 'CR_BAL_LCY')),
                close_bal_lcy=to_decimal(get_field(row, 'CLOSE_BAL_LCY')),
                open_bal_fcy=to_decimal(get_field(row, 'OPEN_BAL_FCY')),
                dr_bal_fcy=to_decimal(get_field(row, 'DR_BAL_FCY')),
                cr_bal_fcy=to_decimal(get_field(row, 'CR_BAL_FCY')),
                close_bal_fcy=to_decimal(get_field(row, 'CLOSE_BAL_FCY')),
                gl_status=str(get_field(row, 'GL_STATUS'))[:50],
            )
            entries.append(entry)

        TrialBalanceEntry.objects.bulk_create(entries, batch_size=500)

        log_activity(
            user=user_profile,
            activity_type='report_submitted',
            details=f'Submitted Daily Trial Balance for {report_date} with {len(entries)} entries',
            request=request
        )

        return JsonResponse({
            'success': True,
            'message': f'Successfully saved {len(entries)} trial balance entries for {report_date}.',
            'report_id': report.id,
            'report_date': report_date.isoformat(),
            'record_count': len(entries)
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data.'}, status=400)
    except Exception as e:
        logger.exception("Error in api_submit_trial_balance")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@login_required
def api_download_trial_balance_template(request):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Trial Balance'

    headers = [
        'Account Code', 'Account Name', 'Branch',
        'Debit (GHS)', 'Credit (GHS)', 'Balance (GHS)', 'Notes',
    ]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0066CC', end_color='0066CC', fill_type='solid')
        cell.alignment = Alignment(horizontal='center')

    sample = ['1000', 'Cash on Hand', 'Head Office', 0, 0, 0, 'Sample entry']
    for col, value in enumerate(sample, 1):
        ws.cell(row=2, column=col, value=value)

    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 20

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="Daily_Trial_Balance_Template_{timezone.now().strftime("%Y%m%d")}.xlsx"'
    return response


# ==================== CONSOLIDATED REPORTS ====================

@login_required
def consolidated_reports(request):
    """
    Consolidated Reports — cross-user view of exception data.

    Access rule:
        The logged-in user may open this page ONLY if they have at
        least one Report assigned to them whose report_type contains
        the word "consolidated" (case-insensitive).

    Data rule:
        The consolidated report is a VIEW over the underlying
        exception data, not a data container of its own. When the
        report name implies a scope:

          • "Head Office" / "Headoffice" / "Head-Office"
                → aggregate EVERY non-TB exception record uploaded by
                  users with position='hc' (across all their reports)
          • "Cluster"
                → same, but for position='cc'

        When no scope can be derived from the name, fall back to
        exact report_type matching.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    # ------------------------------------------------------------------
    # Determine which consolidated report types this user may access
    # ------------------------------------------------------------------
    assigned_consolidated_qs = Report.objects.filter(
        Q(assigned_to=user_profile) |
        Q(is_assigned_to_all=True) |
        Q(created_by=user_profile)
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    ).filter(
        report_type__icontains='consolidated'
    )

    available_report_types = sorted(
        {rt for rt in assigned_consolidated_qs.values_list('report_type', flat=True) if rt}
    )

    has_access = len(available_report_types) > 0

    # ------------------------------------------------------------------
    # Resolve the selected report type
    # ------------------------------------------------------------------
    selected_report_type = (request.GET.get('report_type') or '').strip()
    if not selected_report_type or selected_report_type == 'all':
        if available_report_types:
            selected_report_type = available_report_types[0]

    if selected_report_type not in available_report_types:
        selected_report_type = available_report_types[0] if available_report_types else ''

    # ------------------------------------------------------------------
    # Derive uploader scope from the report name
    # ------------------------------------------------------------------
    def _derive_position_scope(report_type):
        name = (report_type or '').lower()
        if 'head office' in name or 'headoffice' in name or 'head-office' in name:
            return 'hc'
        if 'cluster' in name:
            return 'cc'
        return None

    uploader_scope = _derive_position_scope(selected_report_type)

    # ------------------------------------------------------------------
    # Build the preview across ALL matching exception records
    # ------------------------------------------------------------------
    exception_rows = []
    total_count = 0
    total_cost = 0
    distinct_uploaders = 0

    if has_access and selected_report_type:
        er_qs = ExceptionRecord.objects.filter(
            report__isnull=False,
        ).exclude(
            report__report_type=TRIAL_BALANCE_REPORT_TYPE,
        )

        if uploader_scope:
            # CONSOLIDATED VIEW: aggregate from every report
            # uploaded by users of the scoped position.
            er_qs = er_qs.filter(
                report__created_by__position=uploader_scope,
            )
        else:
            # No scope derivable → exact report_type match only
            er_qs = er_qs.filter(
                report__report_type=selected_report_type,
            )

        er_qs = er_qs.select_related(
            'report', 'report__created_by'
        ).order_by('-report__created_at', 'source_row_index')

        total_count = er_qs.count()

        cost_agg = er_qs.aggregate(total=Sum('income_cost_saved'))
        try:
            total_cost = float(cost_agg.get('total') or 0)
        except (TypeError, ValueError):
            total_cost = 0.0

        distinct_uploaders = er_qs.values('report__created_by').distinct().count()

        for rec in er_qs[:100]:
            exception_rows.append({
                'serial_number': rec.serial_number,
                'branch_unit': rec.branch_unit or '—',
                'exception': rec.exception or '—',
                'date_noted': rec.date_noted,
                'target_closure_date': rec.target_closure_date,
                'category': rec.category or '—',
                'responsible_officer': rec.responsible_officer or '—',
                'supervisor': rec.supervisor or '—',
                'status': rec.get_status_display() or rec.status_raw or '—',
                'income_cost_saved': rec.income_cost_saved,
                'source_report': (rec.report.report_type if rec.report else '—'),
            })

    context = {
        'user_profile': user_profile,
        'today': timezone.now(),
        'has_access': has_access,
        'available_report_types': available_report_types,
        'selected_report_type': selected_report_type,
        'uploader_scope': uploader_scope,
        'exception_rows': exception_rows,
        'total_count': total_count,
        'total_cost': total_cost,
        'distinct_uploaders': distinct_uploaders,
    }
    return render(request, 'control_dashboard/consolidated.html', context)

@csrf_exempt
@require_http_methods(["GET"])
def api_consolidated_filter_options(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

    report_types = list(
        Report.objects
        .filter(created_by=user_profile)
        .exclude(report_type=TRIAL_BALANCE_REPORT_TYPE)
        .exclude(status=UPLOADED_STATUS)
        .values_list('report_type', flat=True)
        .distinct()
        .order_by('report_type')
    )

    return JsonResponse({'success': True, 'report_types': report_types})

@login_required
@require_http_methods(["GET"])
def generate_consolidated_excel(request):
    """
    Export a consolidated report as Excel.

    Uses the same aggregation logic as `consolidated_reports`:
    when the report name implies Head Office / Cluster, ALL
    exception records from that scope are pulled in — not just
    rows whose report_type matches the consolidated name.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        messages.error(request, 'User profile not found.')
        return redirect('control_dashboard:consolidated_reports')

    # ------------------------------------------------------------------
    # Access check
    # ------------------------------------------------------------------
    assigned_consolidated_qs = Report.objects.filter(
        Q(assigned_to=user_profile) |
        Q(is_assigned_to_all=True) |
        Q(created_by=user_profile)
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    ).filter(
        report_type__icontains='consolidated'
    )

    allowed_report_types = {
        rt for rt in assigned_consolidated_qs.values_list('report_type', flat=True) if rt
    }

    if not allowed_report_types:
        messages.error(request, 'You do not have access to consolidated reports.')
        return redirect('control_dashboard:consolidated_reports')

    report_type_filter = (request.GET.get('report_type') or '').strip()
    start_date_str = (request.GET.get('start_date') or '').strip()
    end_date_str = (request.GET.get('end_date') or '').strip()

    if not report_type_filter or report_type_filter == 'all':
        report_type_filter = sorted(allowed_report_types)[0]

    if report_type_filter not in allowed_report_types:
        messages.error(request, f'You do not have access to "{report_type_filter}".')
        return redirect('control_dashboard:consolidated_reports')

    # ------------------------------------------------------------------
    # Derive position scope
    # ------------------------------------------------------------------
    def _derive_position_scope(report_type):
        name = (report_type or '').lower()
        if 'head office' in name or 'headoffice' in name or 'head-office' in name:
            return 'hc'
        if 'cluster' in name:
            return 'cc'
        return None

    uploader_scope = _derive_position_scope(report_type_filter)

    # ------------------------------------------------------------------
    # Query — same aggregation logic as the preview view
    # ------------------------------------------------------------------
    er_qs = ExceptionRecord.objects.filter(
        report__isnull=False,
    ).exclude(
        report__report_type=TRIAL_BALANCE_REPORT_TYPE,
    )

    if uploader_scope:
        er_qs = er_qs.filter(
            report__created_by__position=uploader_scope,
        )
    else:
        er_qs = er_qs.filter(
            report__report_type=report_type_filter,
        )

    er_qs = er_qs.select_related(
        'report', 'report__created_by'
    ).order_by('-report__created_at', 'source_row_index')

    if start_date_str:
        try:
            start_d = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            er_qs = er_qs.filter(date_noted__gte=start_d)
        except ValueError:
            pass

    if end_date_str:
        try:
            end_d = datetime.strptime(end_date_str, '%Y-%m-%d').date()
            er_qs = er_qs.filter(date_noted__lte=end_d)
        except ValueError:
            pass

    records = list(er_qs)

    if not records:
        messages.warning(request, f'No exception records found for "{report_type_filter}".')
        return redirect('control_dashboard:consolidated_reports')

    # ------------------------------------------------------------------
    # Workbook
    # ------------------------------------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    safe_title = re.sub(r'[\[\]\:\*\?\/\\]', '_', report_type_filter)[:31] or 'Consolidated'
    ws.title = safe_title

    headers = [
        'S/N', 'Branch / Unit', 'Exception',
        'Date Noted', 'Target Closure', 'Category',
        'Responsible Officer', 'Supervisor', 'Status',
        "Auditee's Response", 'Remarks',
        'Income / Cost Saved',
        'Source Report',
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0066CC', end_color='0066CC', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    row_idx = 2
    for rec in records:
        ws.cell(row=row_idx, column=1,  value=rec.serial_number)
        ws.cell(row=row_idx, column=2,  value=rec.branch_unit or '')
        ws.cell(row=row_idx, column=3,  value=rec.exception or '')
        ws.cell(row=row_idx, column=4,  value=rec.date_noted.strftime('%Y-%m-%d') if rec.date_noted else (rec.date_noted_raw or ''))
        ws.cell(row=row_idx, column=5,  value=rec.target_closure_date.strftime('%Y-%m-%d') if rec.target_closure_date else (rec.target_closure_date_raw or ''))
        ws.cell(row=row_idx, column=6,  value=rec.category or '')
        ws.cell(row=row_idx, column=7,  value=rec.responsible_officer or '')
        ws.cell(row=row_idx, column=8,  value=rec.supervisor or '')
        ws.cell(row=row_idx, column=9,  value=rec.get_status_display() or rec.status_raw or '')
        ws.cell(row=row_idx, column=10, value=rec.auditee_response or '')
        ws.cell(row=row_idx, column=11, value=rec.remarks or '')
        try:
            ws.cell(row=row_idx, column=12, value=float(rec.income_cost_saved) if rec.income_cost_saved is not None else 0)
        except (TypeError, ValueError):
            ws.cell(row=row_idx, column=12, value=0)
        ws.cell(row=row_idx, column=13, value=(rec.report.report_type if rec.report else ''))
        row_idx += 1

    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        max_len = 12
        for r in range(1, min(ws.max_row + 1, 80)):
            v = ws.cell(row=r, column=col).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = min(max_len + 2, 45)

    ws.freeze_panes = 'A2'

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    safe_filename_part = re.sub(r'[^A-Za-z0-9_-]', '_', report_type_filter)[:40]
    filename = f'Consolidated_{safe_filename_part}_{timezone.now().strftime("%Y%m%d_%H%M")}.xlsx'

    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response

# ==================== TRIAL BALANCE READ APIs ====================

@csrf_exempt
@require_http_methods(["GET"])
def api_get_trial_balance_by_date(request):
    try:
        date_str = request.GET.get('date', '').strip()

        if not date_str:
            return JsonResponse({'success': False, 'error': 'Date parameter is required.'}, status=400)

        try:
            report_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            return JsonResponse({'success': False, 'error': 'Invalid date format. Use YYYY-MM-DD.'}, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

        from .models import TrialBalanceEntry

        code_to_name = dict(Branch.BRANCH_CODE_MAP)
        for b in Branch.objects.filter(is_active=True):
            if b.code:
                code_to_name[str(b.code).strip()] = b.name

        entries = list(
            TrialBalanceEntry.objects
            .filter(report_date=report_date)
            .select_related('uploaded_by')
            .order_by('row_index')
        )

        if not entries:
            return JsonResponse({
                'success': True,
                'found': False,
                'report_date': report_date.isoformat(),
                'message': f'No trial balance has been uploaded for {report_date.isoformat()} yet.'
            })

        user_branches = list(user_profile.branches.all())
        user_departments = list(user_profile.departments.all())

        allowed_branch_codes = set()
        for b in user_branches:
            if b.code:
                allowed_branch_codes.add(str(b.code).strip())
            if b.name:
                for code, name in Branch.BRANCH_CODE_MAP.items():
                    if name.upper() == b.name.upper():
                        allowed_branch_codes.add(code)
                        break

        allowed_branch_names = {str(b.name).strip().upper() for b in user_branches if b.name}
        allowed_department_names = {str(d.name).strip().upper() for d in user_departments if d.name}

        has_branch_restriction = len(allowed_branch_codes) > 0 or len(allowed_branch_names) > 0
        has_department_restriction = len(allowed_department_names) > 0

        is_privileged = user_profile.role in ('admin', 'supervisor')
        is_hc_staff = (user_profile.position == 'hc')

        def passes_assignment(entry):
            if is_privileged:
                return True

            code = str(entry.branch_code or '').strip()
            name = code_to_name.get(code, '').upper()

            if is_hc_staff and code == '000':
                return True

            if not has_branch_restriction and not has_department_restriction:
                return True

            if has_branch_restriction:
                if code in allowed_branch_codes:
                    return True
                if name and name in allowed_branch_names:
                    return True

            if has_department_restriction:
                cat = str(entry.category or '').strip().upper()
                if cat in allowed_department_names:
                    return True

            return False

        visible_entries = [e for e in entries if passes_assignment(e)]

        if not visible_entries:
            return JsonResponse({
                'success': True,
                'found': False,
                'report_date': report_date.isoformat(),
                'message': (f'No entries visible for your assigned branches/departments '
                            f'on {report_date.isoformat()}.')
            })

        headers = [
            'BRANCH_CODE', 'BRANCH_NAME', 'TODAY', 'CATEGORY', 'PARENT_GL',
            'GL_CODE', 'DESCR', 'CCY',
            'OPEN_BAL_LCY', 'DR_BAL_LCY', 'CR_BAL_LCY', 'CLOSE_BAL_LCY',
            'OPEN_BAL_FCY', 'DR_BAL_FCY', 'CR_BAL_FCY', 'CLOSE_BAL_FCY',
            'GL_STATUS',
        ]

        rows = []
        total_debit = 0.0
        total_credit = 0.0

        for entry in visible_entries:
            raw_code = str(entry.branch_code or '').strip()
            branch_name = code_to_name.get(raw_code, raw_code)

            rows.append({
                'BRANCH_CODE': entry.branch_code,
                'BRANCH_NAME': branch_name,
                'TODAY': entry.today,
                'CATEGORY': entry.category,
                'PARENT_GL': entry.parent_gl,
                'GL_CODE': entry.gl_code,
                'DESCR': entry.descr,
                'CCY': entry.ccy,
                'OPEN_BAL_LCY': str(entry.open_bal_lcy),
                'DR_BAL_LCY': str(entry.dr_bal_lcy),
                'CR_BAL_LCY': str(entry.cr_bal_lcy),
                'CLOSE_BAL_LCY': str(entry.close_bal_lcy),
                'OPEN_BAL_FCY': str(entry.open_bal_fcy),
                'DR_BAL_FCY': str(entry.dr_bal_fcy),
                'CR_BAL_FCY': str(entry.cr_bal_fcy),
                'CLOSE_BAL_FCY': str(entry.close_bal_fcy),
                'GL_STATUS': entry.gl_status,
            })

            try:
                total_debit += float(entry.dr_bal_lcy)
            except (ValueError, TypeError):
                pass
            try:
                total_credit += float(entry.cr_bal_lcy)
            except (ValueError, TypeError):
                pass

        first_entry = visible_entries[0]
        uploader = getattr(first_entry, 'uploaded_by', None)
        uploaded_by_info = None
        if uploader:
            uploaded_by_info = {
                'id': uploader.id,
                'full_name': uploader.full_name or uploader.email,
                'email': uploader.email,
            }

        if is_hc_staff and not is_privileged:
            combined_allowed = set(allowed_branch_codes)
            combined_allowed.add('000')
            allowed_branches_list = sorted(combined_allowed)
        else:
            allowed_branches_list = sorted(allowed_branch_codes) if has_branch_restriction else None

        allowed_departments_list = sorted(allowed_department_names) if has_department_restriction else None

        return JsonResponse({
            'success': True,
            'found': True,
            'report_date': report_date.isoformat(),
            'file_name': first_entry.file_name or 'trial_balance.xlsx',
            'headers': headers,
            'rows': rows,
            'row_count': len(rows),
            'summary': {
                'total_rows': len(rows),
                'total_debit': round(total_debit, 2),
                'total_credit': round(total_credit, 2),
                'variance': round(total_debit - total_credit, 2),
                'debit_column': 'DR_BAL_LCY',
                'credit_column': 'CR_BAL_LCY',
            },
            'allowed_branches': allowed_branches_list,
            'allowed_departments': allowed_departments_list,
            'uploaded_by': uploaded_by_info,
            'uploaded_at': first_entry.created_at.isoformat() if first_entry.created_at else None,
        })

    except Exception as e:
        logger.exception("Error in api_get_trial_balance_by_date")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def api_check_trial_balance_day(request):
    try:
        date_str = request.GET.get('date', '').strip()

        if not date_str:
            return JsonResponse({'success': False, 'error': 'Date parameter is required.'}, status=400)

        try:
            report_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            return JsonResponse({'success': False, 'error': 'Invalid date format. Use YYYY-MM-DD.'}, status=400)

        from .models import TrialBalanceEntry

        first_entry = (
            TrialBalanceEntry.objects
            .filter(report_date=report_date)
            .select_related('uploaded_by')
            .order_by('created_at')
            .first()
        )

        if not first_entry:
            return JsonResponse({
                'success': True,
                'locked': False,
                'report_date': report_date.isoformat(),
                'uploaded_by': None,
                'uploaded_at': None,
                'row_count': 0,
            })

        row_count = TrialBalanceEntry.objects.filter(report_date=report_date).count()
        uploader = first_entry.uploaded_by

        return JsonResponse({
            'success': True,
            'locked': True,
            'report_date': report_date.isoformat(),
            'uploaded_by': {
                'id': uploader.id,
                'full_name': uploader.full_name or uploader.email,
                'email': uploader.email,
            },
            'uploaded_at': first_entry.created_at.isoformat() if first_entry.created_at else None,
            'row_count': row_count,
        })

    except Exception as e:
        logger.exception("Error in api_check_trial_balance_day")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["POST"])
def api_update_report_score(request):
    """
    Supervisor deduction override on a SINGLE submission.

    The frontend now sends `score_id` (the SubmittedReportScore row),
    so each leg of the journey can be scored independently.
    `report_id` is still accepted for backward compatibility.
    """
    try:
        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if user_profile.role not in ('supervisor', 'admin'):
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        data = json.loads(request.body)
        score_id = data.get('score_id')
        report_id = data.get('report_id')      # legacy fallback
        deduction = data.get('deduction', 0)

        try:
            deduction = int(deduction)
        except (ValueError, TypeError):
            return JsonResponse({'success': False, 'error': 'Invalid deduction value'}, status=400)

        if deduction < 0 or deduction > 100:
            return JsonResponse({'success': False, 'error': 'Deduction must be between 0 and 100'}, status=400)

        # ------------------------------------------------------------
        # Resolve the target submission row.
        # ------------------------------------------------------------
        if score_id:
            score_obj = get_object_or_404(SubmittedReportScore, id=score_id)
        elif report_id:
            # Legacy: fall back to latest submission for this report.
            score_obj = (
                SubmittedReportScore.objects
                .filter(report_id=report_id)
                .order_by('-sent_at', '-created_at')
                .first()
            )
            if score_obj is None:
                return JsonResponse({'success': False, 'error': 'No submission found for this report'}, status=404)
        else:
            return JsonResponse({'success': False, 'error': 'score_id is required'}, status=400)

        report = score_obj.report
        if report.report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        # Apply the deduction.
        if deduction == 0:
            score_obj.manual_score = None
            score_obj.override_reason = ''
            action = 'cleared'
        else:
            score_obj.manual_score = max(0, score_obj.auto_score - deduction)
            action = f'set (deduction {deduction}%)'

        score_obj.save()

        # NOTE: snapshot_final_score is intentionally NOT updated here.
        # Team performance reads the frozen snapshot so historical
        # metrics don't shift when a supervisor edits an old submission.

        log_activity(
            user=user_profile,
            activity_type='score_updated',
            details=f'Deduction {action} on submission #{score_obj.id} ({report.report_type})',
            request=request
        )

        return JsonResponse({
            'success': True,
            'message': f'Score {action}',
            'score_id': score_obj.id,
            'report_id': report.id,
            'auto_score': score_obj.auto_score,
            'manual_score': score_obj.manual_score,
            'final_score': score_obj.final_score,
            'deduction': deduction,
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error in api_update_report_score")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

# ============================================================
# EXCEPTION TEMPLATE — canonical headers + aliases
# ============================================================
# These are the 12 canonical headers every exception report uses.
EXCEPTION_TEMPLATE_HEADERS = [
    'S/N',
    'BRANCH/UNIT',
    'EXCEPTION',
    'DATE EXCEPTION WAS NOTED',
    'TARGET DATE FOR CLOSURE',
    'CATEGORY OF EXCEPTION',
    'RESPONSIBLE OFFICER',
    'SUPERVISOR',
    "AUDITEE'S RESPONSE",
    'REMARKS',
    'STATUS',
    'INCOME/COST SAVED',
]

# Alias map: canonical header → list of accepted variants
EXCEPTION_HEADER_ALIASES = {
    'S/N': ['S/N', 'SN', 'SERIAL NUMBER', 'SERIAL NO', '#'],
    'BRANCH/UNIT': ['BRANCH/UNIT', 'BRANCH', 'UNIT'],
    'EXCEPTION': ['EXCEPTION', 'EXCEPTIONS', 'FINDING', 'FINDINGS'],
    'DATE EXCEPTION WAS NOTED': [
        'DATE EXCEPTION WAS NOTED', 'DATE EXCEPTION NOTED',
        'DATE NOTED', 'DATE OBSERVED',
    ],
    'TARGET DATE FOR CLOSURE': [
        'TARGET DATE FOR CLOSURE', 'TARGET CLOSURE DATE',
        'TARGET DATE', 'DUE DATE',
    ],
    'CATEGORY OF EXCEPTION': [
        'CATEGORY OF EXCEPTION', 'CATEGORY', 'EXCEPTION CATEGORY',
    ],
    'RESPONSIBLE OFFICER': [
        'RESPONSIBLE OFFICER', 'RESPONSIBLE', 'RESPONSIBLE STAFF', 'OFFICER',
    ],
    'SUPERVISOR': ['SUPERVISOR', 'SUPERVISORS'],
    "AUDITEE'S RESPONSE": [
        "AUDITEE'S RESPONSE", 'AUDITEE RESPONSE', 'RESPONSE',
        'MANAGEMENT RESPONSE',
    ],
    'REMARKS': ['REMARKS', 'REMARK', 'COMMENTS', 'COMMENT', 'NOTES'],
    'STATUS': ['STATUS', 'STATE'],
    'INCOME/COST SAVED': [
        'INCOME/COST SAVED', 'INCOME COST SAVED', 'COST SAVED',
        'INCOME SAVED', 'SAVINGS', 'AMOUNT SAVED',
    ],
}

# Canonical header → ExceptionRecord model field name
EXCEPTION_HEADER_TO_FIELD = {
    'S/N': 'serial_number',
    'BRANCH/UNIT': 'branch_unit',
    'EXCEPTION': 'exception',
    'DATE EXCEPTION WAS NOTED': 'date_noted',
    'TARGET DATE FOR CLOSURE': 'target_closure_date',
    'CATEGORY OF EXCEPTION': 'category',
    'RESPONSIBLE OFFICER': 'responsible_officer',
    'SUPERVISOR': 'supervisor',
    "AUDITEE'S RESPONSE": 'auditee_response',
    'REMARKS': 'remarks',
    'STATUS': 'status',
    'INCOME/COST SAVED': 'income_cost_saved',
}


def _normalize_header(h):
    """Uppercase, collapse whitespace, normalize quotes, strip."""
    import re
    s = str(h or '')
    s = s.replace('\u2019', "'").replace('\u2018', "'")   # curly → straight
    s = s.replace('\u201c', '"').replace('\u201d', '"')
    s = re.sub(r'\s+', ' ', s)
    return s.strip().upper()


# Build a normalized lookup: normalized_alias → canonical
_NORMALIZED_ALIAS_LOOKUP = {}
for _canonical, _aliases in EXCEPTION_HEADER_ALIASES.items():
    for _alias in _aliases:
        _NORMALIZED_ALIAS_LOOKUP[_normalize_header(_alias)] = _canonical


def _match_canonical_header(user_header):
    """Return the canonical template header for a user-supplied header, or None."""
    return _NORMALIZED_ALIAS_LOOKUP.get(_normalize_header(user_header))


def _validate_exception_headers(user_headers):
    """
    Return a dict describing how well user_headers matches the template.
    Same shape as the frontend validator (for consistency).
    """
    matched = []
    unmatched = []
    seen_canonicals = set()

    for idx, h in enumerate(user_headers):
        canonical = _match_canonical_header(h)
        if canonical:
            if canonical not in seen_canonicals:
                seen_canonicals.add(canonical)
                matched.append({'index': idx, 'original': h, 'canonical': canonical})
            else:
                unmatched.append({'index': idx, 'original': h})
        elif h is not None and str(h).strip():
            unmatched.append({'index': idx, 'original': h})

    missing = [c for c in EXCEPTION_TEMPLATE_HEADERS if c not in seen_canonicals]
    match_count = len(matched)
    total = len(EXCEPTION_TEMPLATE_HEADERS)

    if match_count == total and not unmatched:
        status = 'ok'
    elif match_count >= int(total * 0.75):
        status = 'warning'
    else:
        status = 'error'

    return {
        'status': status,
        'matched': matched,
        'missing': missing,
        'unmatched': unmatched,
        'match_count': match_count,
        'total_canonical': total,
    }


def _parse_date_safe(raw):
    """Try common date formats; return (date_obj, original_string)."""
    raw_str = str(raw or '').strip()
    if not raw_str:
        return None, ''
    # Handle Excel serial dates (int/float)
    try:
        if isinstance(raw, (int, float)) and 10000 < float(raw) < 60000:
            from datetime import timedelta
            base = date(1899, 12, 30)
            return base + timedelta(days=int(float(raw))), raw_str
    except (ValueError, TypeError):
        pass

    for fmt in (
        '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y',
        '%d-%b-%Y', '%d %b %Y', '%Y/%m/%d', '%d.%m.%Y',
        '%d-%B-%Y', '%d %B %Y',
    ):
        try:
            return datetime.strptime(raw_str, fmt).date(), raw_str
        except ValueError:
            continue
    return None, raw_str


def _parse_decimal_safe(raw):
    """Strip currency symbols/commas; return (Decimal, original_string)."""
    from decimal import Decimal, InvalidOperation
    raw_str = str(raw or '').strip()
    if raw_str in ('', '-', '--', 'N/A', 'n/a', 'na', 'nil'):
        return Decimal('0'), raw_str
    cleaned = (raw_str.replace(',', '')
                       .replace('GHS', '')
                       .replace('$', '')
                       .replace('₵', '')
                       .strip())
    # Handle parentheses as negative: (1,234.56) → -1234.56
    is_negative = cleaned.startswith('(') and cleaned.endswith(')')
    if is_negative:
        cleaned = cleaned[1:-1].strip()
    try:
        val = Decimal(cleaned)
        return (-val if is_negative else val), raw_str
    except (InvalidOperation, ValueError):
        return Decimal('0'), raw_str


def _parse_status_safe(raw):
    """Map free-text status to our enum + keep original."""
    raw_str = str(raw or '').strip()
    key = raw_str.lower()
    if key in ('open', ''):
        return 'open', raw_str
    if 'progress' in key or 'ongoing' in key:
        return 'in_progress', raw_str
    if key in ('closed', 'close', 'done'):
        return 'closed', raw_str
    if 'resolv' in key:
        return 'resolved', raw_str
    if 'pend' in key:
        return 'pending', raw_str
    if 'reject' in key:
        return 'rejected', raw_str
    if 'overdue' in key or 'over due' in key:
        return 'overdue', raw_str
    return 'open', raw_str

def _build_exception_records(report, headers, rows_data):
    """
    Parse Excel rows into typed ExceptionRecord rows.

    Only ExceptionRecord rows are created — no raw cell mirror.
    Returns the number of records created.
    """
    header_to_field = {}
    for idx, h in enumerate(headers):
        canonical = _match_canonical_header(h)
        if canonical:
            field_name = EXCEPTION_HEADER_TO_FIELD.get(canonical)
            if field_name:
                header_to_field[idx] = field_name

    records_to_create = []
    for row_idx, row in enumerate(rows_data):
        extracted = {}
        for col_idx, field_name in header_to_field.items():
            if col_idx < len(row):
                extracted[field_name] = row[col_idx]

        if not any(str(v or '').strip() for v in extracted.values()):
            continue

        sn_raw = extracted.get('serial_number', row_idx + 1)
        try:
            serial_number = int(str(sn_raw).strip() or (row_idx + 1))
        except (ValueError, TypeError):
            serial_number = row_idx + 1

        date_noted, date_noted_raw = _parse_date_safe(extracted.get('date_noted'))
        target_date, target_date_raw = _parse_date_safe(
            extracted.get('target_closure_date')
        )
        cost_saved, cost_saved_raw = _parse_decimal_safe(
            extracted.get('income_cost_saved')
        )
        status_val, status_raw = _parse_status_safe(extracted.get('status'))

        records_to_create.append(ExceptionRecord(
            report=report,
            serial_number=serial_number,
            branch_unit=str(extracted.get('branch_unit', '') or '').strip()[:200],
            exception=str(extracted.get('exception', '') or '').strip(),
            date_noted=date_noted,
            date_noted_raw=date_noted_raw[:50],
            target_closure_date=target_date,
            target_closure_date_raw=target_date_raw[:50],
            category=str(extracted.get('category', '') or '').strip()[:200],
            responsible_officer=str(
                extracted.get('responsible_officer', '') or ''
            ).strip()[:200],
            supervisor=str(extracted.get('supervisor', '') or '').strip()[:200],
            auditee_response=str(
                extracted.get('auditee_response', '') or ''
            ).strip(),
            remarks=str(extracted.get('remarks', '') or '').strip(),
            status=status_val,
            status_raw=status_raw[:100],
            income_cost_saved=cost_saved,
            income_cost_saved_raw=cost_saved_raw[:50],
            source_row_index=row_idx,
        ))

    if records_to_create:
        ExceptionRecord.objects.bulk_create(records_to_create, batch_size=500)

    return len(records_to_create)

# ============================================================
# REPLACE api_save_imported_data with this version
# ============================================================

@csrf_exempt
@require_http_methods(["POST"])
def api_save_imported_data(request):
    """
    Persist an Excel upload from draft.html as typed exception records.

    Only ExceptionRecord rows are created. The Report container keeps
    the original headers on `excel_headers` so the UI can render the
    correct columns.
    """
    try:
        data = json.loads(request.body)

        headers = data.get('headers', [])
        rows_data = data.get('data', [])
        report_type = data.get('report_type', '')
        file_name = data.get('file_name', 'uploaded.xlsx')

        if not headers or not rows_data:
            return JsonResponse({'success': False, 'error': 'No data to save. Please import an Excel file first.'}, status=400)

        if not report_type:
            return JsonResponse({'success': False, 'error': 'Please select a report type.'}, status=400)

        if report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'This report type is reserved for internal use.'}, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        validation = _validate_exception_headers(headers)

        if validation['status'] == 'error':
            return JsonResponse({
                'success': False,
                'error': (
                    'Column headers in row 1 do not match the standard exception '
                    'template. Matched {matched} of {total}. Missing: {missing}'
                ).format(
                    matched=validation['match_count'],
                    total=validation['total_canonical'],
                    missing=', '.join(validation['missing']),
                ),
                'validation': validation,
            }, status=400)

        # Remove any empty containers left over from previous failed uploads
        Report.objects.filter(
            created_by=user_profile,
            status=UPLOADED_STATUS,
            exception_records__isnull=True,
            data_fields__isnull=True,
        ).delete()

        report = Report.objects.create(
            report_type=report_type,
            frequency='one-off',
            description=f'Excel Import: {file_name}',
            status=UPLOADED_STATUS,
            created_by=user_profile,
            excel_headers=','.join(str(h) for h in headers),
        )

        # Keep ReportDataField rows for any legacy consumers (single-row
        # form reports). Excel uploads never read from them.
        for i, header in enumerate(headers):
            ReportDataField.objects.create(
                report=report,
                field_name=header or f'Column_{i+1}',
                field_value='',
                field_type='text',
                order=i,
            )

        typed_count = 0
        try:
            typed_count = _build_exception_records(
                report=report,
                headers=headers,
                rows_data=rows_data,
            )
        except Exception as exc:
            logger.exception("Failed to build typed ExceptionRecord rows: %s", exc)

        log_activity(
            user=user_profile,
            activity_type='report_submitted',
            details=(
                f'Uploaded {report_type} with {len(rows_data)} exception records '
                f'({typed_count} typed) from Excel'
            ),
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': (
                f'Successfully saved {len(rows_data)} records from {file_name} '
                f'({typed_count} typed as exceptions)'
            ),
            'report_id': report.id,
            'record_count': len(rows_data),
            'typed_record_count': typed_count,
            'validation': validation,
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error saving imported data")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

# ==================== API - EXCEL ROW (SINGLE) EDIT / DELETE ====================

@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_excel_row(request, row_id):
    """
    Delete a SINGLE ExceptionRecord row.
    The URL parameter is now the ExceptionRecord.id, not ReportExcelRow.id.
    """
    try:
        record = get_object_or_404(ExceptionRecord, id=row_id)
        report = record.report

        if report is None or report.report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        record.delete()

        remaining = report.exception_records.count()

        log_activity(
            user=user_profile,
            activity_type='draft_deleted',
            details=f'Deleted 1 exception row from report #{report.id} ({report.report_type})',
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': 'Row deleted successfully.',
            'remaining_rows': remaining,
        })

    except Exception as e:
        logger.exception("Error in api_delete_excel_row")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@csrf_exempt
@require_http_methods(["POST"])
def api_edit_excel_row(request, row_id):
    """
    Edit a SINGLE ExceptionRecord row.
    Body: { "cells": { "COLUMN_NAME": "new value", ... } }
    """
    try:
        record = get_object_or_404(ExceptionRecord, id=row_id)
        report = record.report

        if report is None or report.report_type == TRIAL_BALANCE_REPORT_TYPE:
            return JsonResponse({'success': False, 'error': 'Not found'}, status=404)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if report.created_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        data = json.loads(request.body)
        cells = data.get('cells', {})

        if not isinstance(cells, dict):
            return JsonResponse({'success': False, 'error': 'cells must be an object'}, status=400)

        for column_name, raw_val in cells.items():
            canonical = _match_canonical_header(column_name)
            field = EXCEPTION_HEADER_TO_FIELD.get(canonical) if canonical else None
            if not field:
                continue

            if field == 'serial_number':
                try:
                    record.serial_number = int(str(raw_val).strip())
                except (ValueError, TypeError):
                    pass
            elif field == 'date_noted':
                d, d_raw = _parse_date_safe(raw_val)
                record.date_noted = d
                record.date_noted_raw = d_raw[:50]
            elif field == 'target_closure_date':
                d, d_raw = _parse_date_safe(raw_val)
                record.target_closure_date = d
                record.target_closure_date_raw = d_raw[:50]
            elif field == 'income_cost_saved':
                dec, dec_raw = _parse_decimal_safe(raw_val)
                record.income_cost_saved = dec
                record.income_cost_saved_raw = dec_raw[:50]
            elif field == 'status':
                s_val, s_raw = _parse_status_safe(raw_val)
                record.status = s_val
                record.status_raw = s_raw[:100]
            else:
                setattr(record, field, str(raw_val or '').strip())

        record.save()

        log_activity(
            user=user_profile,
            activity_type='report_updated',
            details=f'Edited 1 exception row (#{record.id}) in report #{report.id}',
            request=request,
        )

        return JsonResponse({'success': True, 'message': 'Row updated successfully.'})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error in api_edit_excel_row")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@csrf_exempt
@require_http_methods(["GET"])
def api_supervisor_top_performers_live(request):
    """
    Lightweight JSON endpoint for the Supervisor Dashboard's
    'Top Performers' widget.

    Returns the top 5 members ranked by the same score model used
    in the initial render:
        final_score = max(0, submitted − deductions)
        percentage  = min(100, final_score / max_submissions × 100)

    Also returns the summary aggregates so the KPI header cards can
    refresh in the same round-trip.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

    today = timezone.now().date()

    # -------- Summary counters (mirrors supervisor_dashboard) --------
    exceptions_qs = ExceptionRecord.objects.filter(
        report__isnull=False,
    ).exclude(
        report__report_type=TRIAL_BALANCE_REPORT_TYPE,
    )
    total_exceptions = exceptions_qs.count()
    today_exceptions = exceptions_qs.filter(created_at__date=today).count()

    submitted_qs = Report.objects.filter(status='submitted').exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    )
    submitted_reports_count = submitted_qs.count()

    team_members = (
        UserProfile.objects
        .filter(role='member', status='active')
        .order_by('full_name')
    )

    max_submissions = (
        submitted_qs.values('created_by')
        .annotate(c=Count('id'))
        .order_by('-c')
        .values_list('c', flat=True)
        .first()
    ) or 1

    performers = []
    sum_of_percentages = 0
    total_members = 0

    for member in team_members:
        member_submitted = submitted_qs.filter(created_by=member).count()
        member_deductions = (
            AdHocDeduction.objects
            .filter(user=member)
            .aggregate(total=Sum('points'))
            .get('total') or 0
        )
        final_score = max(0, member_submitted - member_deductions)
        percentage = int((final_score / max_submissions) * 100) if max_submissions > 0 else 0
        percentage = min(percentage, 100)

        if percentage >= 80:
            status = 'success'
            status_icon = '🌟'
            status_text = 'Excellent'
        elif percentage >= 50:
            status = 'warning'
            status_icon = '📈'
            status_text = 'Good'
        else:
            status = 'danger'
            status_icon = '⚠️'
            status_text = 'Needs Attention'

        performers.append({
            'id': member.id,
            'full_name': member.full_name or member.email,
            'submitted': member_submitted,
            'deductions': member_deductions,
            'final_score': final_score,
            'percentage': percentage,
            'status': status,
            'status_text': status_text,
            'status_icon': status_icon,
        })

        total_members += 1
        sum_of_percentages += percentage

    performers.sort(key=lambda x: x['percentage'], reverse=True)
    top_performers = performers[:5]

    completion_rate = (
        int(sum_of_percentages / total_members) if total_members else 0
    )

    return JsonResponse({
        'success': True,
        'top_performers': top_performers,
        'summary': {
            'total_exceptions': total_exceptions,
            'today_exceptions': today_exceptions,
            'submitted_reports_count': submitted_reports_count,
            'completion_rate': completion_rate,
            'team_size': total_members,
        },
    })