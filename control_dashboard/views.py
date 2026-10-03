from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import login as auth_login, logout as auth_logout
from django.contrib.auth.models import User as DjangoUser
from django.utils import timezone
from django.http import JsonResponse, HttpResponse
from django.contrib import messages
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods, require_POST
from django.db import transaction
from django.core.mail import EmailMessage, EmailMultiAlternatives
from django.conf import settings
from django.db.models import Q, Count, Sum
from django.db.models.functions import Coalesce
from django.contrib.auth.decorators import login_required
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from PIL import Image
import json
from datetime import datetime, timedelta, date
import base64
import io
import os, logging
import uuid
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from html import escape as html_escape
import re

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
    ExceptionUpload,
    TrialBalanceUpload,
    TrialBalanceEntry,
    SentEmail,
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

def _initials_from_name(full_name):
    """Return 1–2 uppercase initials from a name or email."""
    if not full_name:
        return ''
    s = str(full_name).strip()
    if not s:
        return ''
    # If it looks like an email, use the local part before the @
    if '@' in s:
        s = s.split('@', 1)[0]
    # Split on spaces, dots, underscores, hyphens
    parts = re.split(r'[\s._\-]+', s)
    parts = [p for p in parts if p]
    if not parts:
        return ''
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[1][0]).upper()

def get_week_start(d):
    """Friday-based week start. Single source of truth."""
    weekday = d.weekday()
    if weekday >= 4:
        days_since_friday = weekday - 4
    else:
        days_since_friday = weekday + 3
    return d - timedelta(days=days_since_friday)

def count_expected_occurrences(checklist, period_start, period_end):
    """
    Number of expected occurrences for a checklist in a period.

    Single source of truth for the expected-occurrence denominator.
    MUST agree with:
      • _add_frequency()                 (Report recurrence stepping)
      • api_log_checklist()'s period    (one-log-per-period rule)
      • get_week_start()                (Friday-anchored weeks)

    Rules per frequency:
      daily      → delegate to the model (weekday-only)
      weekly     → one per Friday-anchored (Fri–Thu) week overlapping
                   the period
      monthly    → one per calendar month overlapping the period
      quarterly  → one per calendar quarter overlapping the period
      bi-annual  → one per Jan–Jun / Jul–Dec block overlapping
      annual     → one per calendar year overlapping
      one-off    → 1 if the period contains the checklist's creation
                   date, else 0
    """
    freq = checklist.frequency

    if freq == 'daily':
        return checklist.get_expected_occurrences(period_start, period_end)

    if freq == 'one-off':
        created = getattr(checklist, 'created_at', None)
        if created is None:
            return 1 if period_start <= period_end else 0
        created_date = created.date() if hasattr(created, 'date') else created
        return 1 if period_start <= created_date <= period_end else 0

    if freq == 'weekly':
        # Count Friday-anchored weeks that overlap [period_start, period_end].
        count = 0
        cursor = get_week_start(period_start)
        safety = 0
        while cursor <= period_end and safety < 10000:
            week_end = cursor + timedelta(days=6)
            if week_end >= period_start and cursor <= period_end:
                count += 1
            cursor += timedelta(days=7)
            safety += 1
        return count

    # Monthly / quarterly / bi-annual / annual — step by month-blocks
    if freq == 'monthly':
        block_months = 1
    elif freq == 'quarterly':
        block_months = 3
    elif freq == 'bi-annual':
        block_months = 6
    elif freq == 'annual':
        block_months = 12
    else:
        # Unknown frequency — safest fallback
        return checklist.get_expected_occurrences(period_start, period_end)

    count = 0
    safety = 0

    # Anchor on the first day of the block containing period_start.
    cursor = period_start.replace(day=1)
    if block_months == 3:
        q_start_month = ((cursor.month - 1) // 3) * 3 + 1
        cursor = cursor.replace(month=q_start_month)
    elif block_months == 6:
        h_start_month = 1 if cursor.month <= 6 else 7
        cursor = cursor.replace(month=h_start_month)
    elif block_months == 12:
        cursor = cursor.replace(month=1)

    while cursor <= period_end and safety < 10000:
        # Compute block_end inclusive
        end_index = (cursor.month - 1) + block_months
        end_year  = cursor.year + (end_index // 12)
        end_month = (end_index % 12)
        if end_month == 0:
            end_month = 12
            end_year -= 1
        if end_month == 12:
            block_end = date(end_year + 1, 1, 1) - timedelta(days=1)
        else:
            block_end = date(end_year, end_month + 1, 1) - timedelta(days=1)

        if block_end >= period_start and cursor <= period_end:
            count += 1

        # Advance cursor to the first day of the next block
        next_index = (cursor.month - 1) + block_months
        next_year  = cursor.year + (next_index // 12)
        next_month = (next_index % 12) + 1
        cursor = date(next_year, next_month, 1)

        safety += 1

    return count

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

def _build_member_scorecard(member, today=None):
    """
    Compute a member's scorecard from their SentEmail history.

    Rules (locked with the business):
      • Each expected occurrence of an assigned Report is worth up to 100 pts.
      • Sent on time (sent_at <= deadline)                  → 100 pts (minus deduction)
      • Sent within 3 days AFTER the deadline               → 40 pts  (minus deduction)
      • Sent more than 3 days after the deadline            → 0 pts
      • No email, deadline passed > 3 days ago              → 0 pts
      • No email, deadline passed but still in grace (<=3d) → 0 pts for now,
        but the occurrence stays in the denominator so the score has room
        to grow if they send within the window.
      • No email, deadline still in the future              → not counted yet.

    Emails are matched to occurrences in deadline order (oldest first),
    so skipping a period cannot be "papered over" by sending extra
    emails later.

    Bonus emails (more emails than expected occurrences for a report_type)
    still add credits to the numerator.

    AdHoc adjustments move the numerator independently.

    Returns a dict with all fields the team.html template needs.
    """
    if today is None:
        today = timezone.now()
    if timezone.is_naive(today):
        today = timezone.make_aware(today)

    GRACE_DAYS = 3

    # ── 1. Assigned reports (excluding internal + uploads) ─────────
    assigned_reports = (
        Report.objects
        .filter(Q(assigned_to=member) | Q(is_assigned_to_all=True))
        .exclude(report_type=TRIAL_BALANCE_REPORT_TYPE)
        .exclude(status=UPLOADED_STATUS)
        .distinct()
    )

    # ── 2. Emails sent by this member ──────────────────────────────
    emails = list(
        SentEmail.objects
        .filter(sender=member)
        .order_by('sent_at')
    )

    # Bucket emails by report_type (used for the count-based matching)
    emails_by_type = {}
    for email in emails:
        emails_by_type.setdefault(email.report_type, []).append(email)

    total_expected   = 0
    total_earned     = 0
    total_deducted   = 0   # supervisor deductions applied on earned emails

    # ── 3. Score each report the member is assigned to ─────────────
    for report in assigned_reports:
        # Build the deadline list for this report from its anchor
        # up to now, respecting the frequency.
        occurrences = _generate_occurrences(report, until=today)
        if not occurrences:
            continue

        rtype_emails = emails_by_type.get(report.report_type, [])

        # Match emails to occurrences oldest-first.
        # emails are already sorted ascending; occurrences are too.
        used_email_ids = set()

        for idx, deadline in enumerate(occurrences):
            # Only count occurrences whose deadline has arrived.
            if deadline > today:
                continue

            total_expected += 1

            # Find the next unused email for this report_type
            matched_email = None
            for e in rtype_emails:
                if e.id in used_email_ids:
                    continue
                matched_email = e
                break

            if matched_email is None:
                # No email → 0 points for this occurrence.
                # Still counted in the denominator (either missed or
                # still in grace — either way earns nothing right now).
                continue

            used_email_ids.add(matched_email.id)

            # Score based on lateness
            delta = matched_email.sent_at - deadline
            seconds_late = delta.total_seconds()

            if seconds_late <= 0:
                earned = 100
            elif seconds_late <= GRACE_DAYS * 86400:
                earned = 40
            else:
                earned = 0

            # Supervisor deduction applies to earned points, floored at 0
            deduction = matched_email.manual_deduction or 0
            earned = max(0, earned - deduction)

            total_earned   += earned
            total_deducted += deduction

        # Bonus emails beyond expected occurrences → +100 each (minus deduction)
        bonus_emails = [e for e in rtype_emails if e.id not in used_email_ids]
        for e in bonus_emails:
            deduction = e.manual_deduction or 0
            total_earned   += max(0, 100 - deduction)
            total_deducted += deduction

    # ── 4. AdHoc adjustments ───────────────────────────────────────
    adhoc_debits = (
        AdHocDeduction.objects
        .filter(user=member)
        .aggregate(total=Sum('points'))
        .get('total') or 0
    )
    adhoc_credits = (
        AdHocDeduction.objects
        .filter(user=member)
        .aggregate(total=Sum('points_added'))
        .get('total') or 0
    )

    # ── 5. Totals ──────────────────────────────────────────────────
    credits = total_earned + adhoc_credits
    debits  = (total_expected * 100 - total_earned) + adhoc_debits
    net_score = credits - debits

    email_count = len(emails)

    if total_expected > 0:
        percentage = max(0, min(100, int(round(net_score / (total_expected * 100) * 100))))
    else:
        percentage = 0

    # ── 6. Status label ────────────────────────────────────────────
    if total_expected == 0:
        status = 'danger'
        status_text = 'No Tasks'
    elif percentage >= 90:
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

    return {
        'user':            member,
        'email_count':     email_count,
        'total_tasks':     total_expected,      # kept for template compat
        'expected':        total_expected,
        'debits':          debits,
        'credits':         credits,
        'net_score':       net_score,
        'percentage':      percentage,
        'percentage_bar':  percentage,
        'status':          status,
        'status_text':     status_text,
    }


def _generate_occurrences(report, until=None):
    """
    Walk a Report forward from its anchor (deadline_date + deadline_time)
    using its frequency, and return the list of deadline datetimes that
    are anchored to it.

    The list is sorted ascending. Each entry is a timezone-aware datetime
    representing "the moment the submission becomes late".

    For one-off reports, this returns a single entry (its anchor) if the
    anchor exists.
    """
    if until is None:
        until = timezone.now()
    if timezone.is_naive(until):
        until = timezone.make_aware(until)

    if not report.deadline_date:
        return []

    anchor_time = (
        report.deadline_time
        or datetime.strptime('23:59', '%H:%M').time()
    )
    anchor_dt = timezone.make_aware(
        datetime.combine(report.deadline_date, anchor_time)
    )

    frequency = report.frequency or 'one-off'

    if frequency == 'one-off':
        return [anchor_dt]

    # Recurring — walk forward until we pass `until`
    occurrences = []
    current = anchor_dt
    safety = 0
    while current <= until and safety < 2000:
        occurrences.append(current)
        nxt = _add_frequency(current, frequency)
        if nxt is None:
            break
        current = nxt
        safety += 1

    return occurrences

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
        # ---- Admin role guard ----
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        data = json.loads(request.body)
        email = data.get('email', '').strip().lower()
        username = data.get('username', '').strip()
        full_name = data.get('full_name', '').strip()
        position = data.get('position', 'member')
        role = data.get('role', 'member')
        status = data.get('status', 'active')
        department_id = data.get('department_id')
        branch_id = data.get('branch_id')

        # ---- Required field validation ----
        if not email:
            return JsonResponse({'success': False, 'error': 'Email is required'}, status=400)
        if not username:
            return JsonResponse({'success': False, 'error': 'Username is required'}, status=400)
        if not full_name:
            return JsonResponse({'success': False, 'error': 'Full name is required'}, status=400)

        # ---- Username format check ----
        if not re.match(r'^[a-zA-Z0-9._-]{3,150}$', username):
            return JsonResponse({
                'success': False,
                'error': 'Username must be 3–150 chars, letters/digits/._- only.',
            }, status=400)

        # ---- Uniqueness ----
        if UserProfile.objects.filter(email__iexact=email).exists():
            return JsonResponse({'success': False, 'error': 'A user with this email already exists'}, status=400)
        if UserProfile.objects.filter(username__iexact=username).exists():
            return JsonResponse({'success': False, 'error': 'This username is already taken'}, status=400)

        user = UserProfile.objects.create(
            email=email,
            username=username,
            full_name=full_name,
            position=position,
            role=role,
            status=status,
        )

        # ---------- Assign departments ----------
        dept_ids = []
        if isinstance(data.get('department_ids'), list):
            dept_ids = [int(x) for x in data['department_ids'] if str(x).strip().isdigit()]
        elif department_id:
            dept_ids = [int(department_id)]

        if dept_ids:
            depts = Department.objects.filter(id__in=dept_ids, is_active=True)
            if depts.exists():
                user.departments.set(depts)

        # ---------- Assign branches ----------
        br_ids = []
        if isinstance(data.get('branch_ids'), list):
            br_ids = [int(x) for x in data['branch_ids'] if str(x).strip().isdigit()]
        elif branch_id:
            br_ids = [int(branch_id)]

        if br_ids:
            branches_qs = Branch.objects.filter(id__in=br_ids, is_active=True)
            if branches_qs.exists():
                user.branches.set(branches_qs)

        user.save()

        log_activity(
            user=caller,
            activity_type='user_created',
            details=f'User {email} was created with role {user.get_role_display()} (username: {user.username})',
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': f'User created successfully.',
            'user_id': user.id,
            'username': user.username,
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception:
        logger.exception("Error in api_create_user")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)

@csrf_exempt
@require_http_methods(["PUT", "POST"])
def api_edit_user(request, user_id):
    try:
        # ---- Admin role guard ----
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

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
        # ---- Admin role guard ----
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

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


# ═══ CHANGED ═══ Added admin role guard + self-deletion guard

@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_user(request, user_id):
    try:
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        user = get_object_or_404(UserProfile, id=user_id)

        if user.id == caller.id:
            return JsonResponse({'success': False, 'error': 'You cannot delete your own account'}, status=400)

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

    # ---- Total exceptions ----
    exceptions_qs = ExceptionRecord.objects.filter(
        upload__isnull=False,
    )
    total_exceptions = exceptions_qs.count()
    today_exceptions = exceptions_qs.filter(created_at__date=today).count()

    # ---- Submitted reports count ----
    submitted_qs = Report.objects.filter(status='submitted').exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    )
    submitted_reports_count = submitted_qs.count()

    # ---- Team performance — RANKED BY NET SCORE ----
    team_members = (
        UserProfile.objects
        .filter(role='member', status='active')
        .order_by('full_name')
    )

    team_performance = []
    for member in team_members:
        member_submitted = submitted_qs.filter(created_by=member).count()

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

        if percentage >= 90:
            status = 'success'
        elif percentage >= 70:
            status = 'success'
        elif percentage >= 50:
            status = 'warning'
        elif percentage > 0:
            status = 'warning'
        else:
            status = 'danger'

        team_performance.append({
            'user': member,
            'submitted': member_submitted,
            'deductions': member_debits,
            'credits': member_credits,
            'net_score': net_score,
            'final_score': net_score,
            'percentage': percentage,
            'status': status,
        })

    team_performance.sort(
        key=lambda x: (x['net_score'], x['submitted']),
        reverse=True,
    )

    completion_rate = (
        int(sum(m['percentage'] for m in team_performance) / len(team_performance))
        if team_performance else 0
    )

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
    """
    Checklist Builder — create-only page.

    Fields:
      • Name
      • Tasks
      • Frequency
      • Assigned Branches (multi-select)
      • Assigned Departments (multi-select)
      • Create button
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    context = {
        'user_profile': user_profile,
        'branches':    Branch.objects.filter(is_active=True).order_by('name'),
        'departments': Department.objects.filter(is_active=True).order_by('name'),
        'frequencies': Checklist.FREQUENCY_CHOICES,
    }
    return render(request, 'control_dashboard/checklist.html', context)


@login_required
def checklist_list(request):
    """Checklist List — table + edit modal + delete."""
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    checklists = (
        Checklist.objects
        .prefetch_related('tasks', 'assigned_branches', 'assigned_departments')
        .order_by('-created_at')
    )

    context = {
        'user_profile': user_profile,
        'checklists':   checklists,
        'branches':     Branch.objects.filter(is_active=True).order_by('name'),
        'departments':  Department.objects.filter(is_active=True).order_by('name'),
        'frequencies':  Checklist.FREQUENCY_CHOICES,
        'today':        timezone.now(),
    }
    return render(request, 'control_dashboard/checklist_list.html', context)


# ==================== API - CHECKLIST MANAGEMENT ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_create_checklist(request):
    """
    Create a checklist.

    Payload:
    {
        "name":                 "...",
        "description":          "..." (optional),
        "frequency":            "daily" | "weekly" | ...,
        "tasks":                [{"description": "..."}, ...],
        "assigned_branches":    [1, 2, 3],
        "assigned_departments": [4, 5]
    }
    """
    try:
        if not request.body:
            return JsonResponse({'success': False, 'error': 'Empty request body.'}, status=400)

        data = json.loads(request.body)

        name                     = (data.get('name') or '').strip()
        description              = (data.get('description') or '').strip()
        frequency                = (data.get('frequency') or 'weekly').strip()
        tasks_data               = data.get('tasks') or []
        assigned_branches_ids    = data.get('assigned_branches') or []
        assigned_departments_ids = data.get('assigned_departments') or []

        # ── Validation ──
        if not name:
            return JsonResponse({'success': False, 'error': 'Checklist name is required.'}, status=400)

        valid_freq = {v for v, _ in Checklist.FREQUENCY_CHOICES}
        if frequency not in valid_freq:
            return JsonResponse({'success': False, 'error': 'Invalid frequency.'}, status=400)

        if not tasks_data:
            return JsonResponse({'success': False, 'error': 'Add at least one task.'}, status=400)

        if not assigned_branches_ids and not assigned_departments_ids:
            return JsonResponse({
                'success': False,
                'error': 'Assign at least one branch or department.',
            }, status=400)

        # ── Caller ──
        if not request.user.is_authenticated:
            return JsonResponse({'success': False, 'error': 'Authentication required.'}, status=401)

        created_by = UserProfile.objects.filter(email=request.user.email).first()
        if not created_by:
            return JsonResponse({'success': False, 'error': 'Admin profile not found.'}, status=404)
        if created_by.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied.'}, status=403)

        # ── Create ──
        checklist = Checklist.objects.create(
            name=name,
            description=description,
            frequency=frequency,
            is_active=True,
            created_by=created_by,
        )

        if assigned_branches_ids:
            checklist.assigned_branches.set(
                Branch.objects.filter(id__in=assigned_branches_ids, is_active=True)
            )

        if assigned_departments_ids:
            checklist.assigned_departments.set(
                Department.objects.filter(id__in=assigned_departments_ids, is_active=True)
            )

        for index, task_item in enumerate(tasks_data):
            task_desc = (task_item.get('description') or '').strip()
            if task_desc:
                ChecklistTask.objects.create(
                    checklist=checklist,
                    description=task_desc,
                    order=index,
                )

        log_activity(
            user=created_by,
            activity_type='checklist_created',
            details=f'Created checklist "{name}" with {len(tasks_data)} tasks ({frequency})',
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': 'Checklist created successfully.',
            'checklist_id': checklist.id,
        }, status=201)

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON payload.'}, status=400)
    except Exception:
        logger.exception("Error in api_create_checklist")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)

@csrf_exempt
@require_http_methods(["GET"])
def api_get_checklist(request, checklist_id):
    try:
        c = get_object_or_404(Checklist, id=checklist_id)
        return JsonResponse({
            'success': True,
            'checklist': {
                'id': c.id,
                'name': c.name,
                'description': c.description,
                'frequency': c.frequency,
                'is_active': c.is_active,
                'assigned_branches':    list(c.assigned_branches.values_list('id', flat=True)),
                'assigned_departments': list(c.assigned_departments.values_list('id', flat=True)),
                'tasks': [
                    {'id': t.id, 'description': t.description, 'order': t.order}
                    for t in c.tasks.all().order_by('order')
                ],
            }
        })
    except Exception:
        logger.exception("Error in api_get_checklist")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)

@csrf_exempt
@require_http_methods(["PUT", "POST"])
def api_edit_checklist(request, checklist_id):
    """
    Edit from the Checklist List page. Accepts any subset of:
        name, description, frequency, is_active,
        assigned_branches, assigned_departments, tasks
    """
    try:
        checklist = get_object_or_404(Checklist, id=checklist_id)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

        if user_profile.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied.'}, status=403)

        if not request.body:
            return JsonResponse({'success': False, 'error': 'Empty payload.'}, status=400)

        data = json.loads(request.body)
        changes = []

        if 'name' in data:
            new_name = (data['name'] or '').strip()
            if not new_name:
                return JsonResponse({'success': False, 'error': 'Name cannot be empty.'}, status=400)
            if new_name != checklist.name:
                checklist.name = new_name
                changes.append(f'Name → "{new_name}"')

        if 'description' in data:
            checklist.description = (data['description'] or '').strip()

        if 'frequency' in data:
            freq = data['frequency']
            valid = {v for v, _ in Checklist.FREQUENCY_CHOICES}
            if freq in valid and freq != checklist.frequency:
                changes.append(f'Frequency → {freq}')
                checklist.frequency = freq

        if 'is_active' in data:
            checklist.is_active = bool(data['is_active'])
            changes.append(f'Active → {checklist.is_active}')

        checklist.save()

        if 'assigned_branches' in data:
            ids = [int(x) for x in (data['assigned_branches'] or []) if str(x).isdigit()]
            checklist.assigned_branches.set(
                Branch.objects.filter(id__in=ids, is_active=True)
            )
            changes.append(f'{len(ids)} branch(es)')

        if 'assigned_departments' in data:
            ids = [int(x) for x in (data['assigned_departments'] or []) if str(x).isdigit()]
            checklist.assigned_departments.set(
                Department.objects.filter(id__in=ids, is_active=True)
            )
            changes.append(f'{len(ids)} department(s)')

        if 'tasks' in data:
            incoming = data['tasks'] or []
            existing = {t.id: t for t in checklist.tasks.all()}
            seen_ids = set()

            for idx, item in enumerate(incoming):
                desc = (item.get('description') or '').strip()
                if not desc:
                    continue
                tid = item.get('id')
                if tid and tid in existing:
                    t = existing[tid]
                    t.description = desc
                    t.order = idx
                    t.save()
                    seen_ids.add(tid)
                else:
                    new_t = ChecklistTask.objects.create(
                        checklist=checklist, description=desc, order=idx,
                    )
                    seen_ids.add(new_t.id)

            for tid, t in existing.items():
                if tid not in seen_ids:
                    t.delete()

            changes.append(f'{len(incoming)} task(s)')

        if changes:
            log_activity(
                user=user_profile,
                activity_type='checklist_updated',
                details=f'Checklist "{checklist.name}" updated: ' + '; '.join(changes),
                request=request,
            )

        return JsonResponse({'success': True, 'message': 'Checklist updated successfully.'})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON.'}, status=400)
    except Exception:
        logger.exception("Error in api_edit_checklist")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)

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

    # ================================================================
    # BRANCH / DEPARTMENT FILTER — scopes ONLY the Irregular GL panel
    # ================================================================
    unit_filter_raw = request.GET.get('unit', 'all').strip()

    selected_unit_type = 'all'
    selected_unit_id = None
    selected_unit_display = 'All Branches & Departments'

    if unit_filter_raw.startswith('branch_'):
        try:
            selected_unit_id = int(unit_filter_raw.split('_', 1)[1])
            selected_unit_type = 'branch'
        except (ValueError, IndexError):
            selected_unit_id = None
            selected_unit_type = 'all'
    elif unit_filter_raw.startswith('dept_'):
        try:
            selected_unit_id = int(unit_filter_raw.split('_', 1)[1])
            selected_unit_type = 'department'
        except (ValueError, IndexError):
            selected_unit_id = None
            selected_unit_type = 'all'

    unit_filter_options = [{'value': 'all', 'display': 'All Branches & Departments'}]
    for b in user_profile.branches.all().order_by('name'):
        unit_filter_options.append({
            'value': f'branch_{b.id}',
            'display': f'🏢 {b.name}',
        })
    for d in user_profile.departments.all().order_by('name'):
        unit_filter_options.append({
            'value': f'dept_{d.id}',
            'display': f'🏛️ {d.name}',
        })

    if selected_unit_type != 'all' and selected_unit_id is not None:
        for opt in unit_filter_options:
            if opt['value'] == unit_filter_raw:
                selected_unit_display = opt['display']
                break

    user_branch_ids = set(user_profile.branches.values_list('id', flat=True))
    user_dept_ids = set(user_profile.departments.values_list('id', flat=True))

    user_checklists = Checklist.objects.filter(
        is_active=True
    ).filter(
        Q(assigned_branches__in=user_branch_ids) |
        Q(assigned_departments__in=user_dept_ids)
    ).distinct()

    total_checklist_rows = 0
    frequency_counts = {}

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

    # Exceptions captured = total typed exception rows the member has
    # uploaded (across all their ExceptionUpload containers).
    total_exceptions_captured = ExceptionRecord.objects.filter(
        upload__uploaded_by=user_profile,
    ).count()

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

    # ============================================================
    # PERIOD WINDOWS — month, quarter, year
    # ============================================================
    month_expected = 0
    month_actual = 0
    quarter_expected = 0
    quarter_actual = 0
    year_expected = 0
    year_actual = 0

    year_start = today.replace(month=1, day=1)
    year_end = today.replace(month=12, day=31)

    current_quarter = (today.month - 1) // 3 + 1
    q_start_month = (current_quarter - 1) * 3 + 1
    q_end_month   = current_quarter * 3

    quarter_start = today.replace(month=q_start_month, day=1)
    if q_end_month == 12:
        quarter_end = today.replace(month=12, day=31)
    else:
        quarter_end = today.replace(
            month=q_end_month + 1, day=1
        ) - timedelta(days=1)

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
        checklist_quarter_expected = 0
        checklist_quarter_actual = 0
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

            q_exp = count_expected_occurrences(checklist, quarter_start, quarter_end)
            q_act = ChecklistLog.objects.filter(
                checklist=checklist, user=user_profile,
                log_date__gte=quarter_start, log_date__lte=quarter_end,
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
            checklist_quarter_expected += q_exp
            checklist_quarter_actual += q_act
            checklist_year_expected += y_exp
            checklist_year_actual += y_act

        month_expected += checklist_month_expected
        month_actual += checklist_month_actual
        quarter_expected += checklist_quarter_expected
        quarter_actual += checklist_quarter_actual
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
        checklist.quarter_expected = checklist_quarter_expected
        checklist.quarter_actual = checklist_quarter_actual
        checklist.year_expected = checklist_year_expected
        checklist.year_actual = checklist_year_actual
        checklist.status = status
        checklist.next_due_date = next_due.strftime('%b %d, %Y') if next_due else None
        checklist.visible_units = visible_units

        display_checklists_with_progress.append(checklist)

    overall_month_progress   = min(int(month_actual   / month_expected   * 100), 100) if month_expected   > 0 else 0
    overall_quarter_progress = min(int(quarter_actual / quarter_expected * 100), 100) if quarter_expected > 0 else 0
    overall_year_progress    = min(int(year_actual    / year_expected    * 100), 100) if year_expected    > 0 else 0

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
    # IRREGULAR GL POSITIONS — date-aware + unit filter
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

        available_tb_dates = list(
            TrialBalanceEntry.objects
            .values_list('report_date', flat=True)
            .distinct()
            .order_by('-report_date')[:90]
        )

        if gl_date_param:
            try:
                target_date = datetime.strptime(gl_date_param, '%Y-%m-%d').date()
            except ValueError:
                target_date = today_date
        else:
            target_date = today_date
            if not TrialBalanceEntry.objects.filter(report_date=today_date).exists():
                if available_tb_dates:
                    target_date = available_tb_dates[0]

        selected_gl_date = target_date
        selected_gl_date_display = target_date.strftime('%b %d, %Y')
        is_today_gl = (target_date == today_date)

        selectable_gl_dates = [today_date]
        for d in available_tb_dates:
            if d != today_date:
                selectable_gl_dates.append(d)

        has_trial_balance = TrialBalanceEntry.objects.filter(
            report_date=target_date
        ).exists()

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

        if selected_unit_type == 'branch' and selected_unit_id:
            try:
                chosen_branch = Branch.objects.get(id=selected_unit_id)
            except Branch.DoesNotExist:
                chosen_branch = None

            if chosen_branch is None:
                entries_qs = entries_qs.none()
            else:
                chosen_code = (chosen_branch.code or '').strip()
                if chosen_code:
                    entries_qs = entries_qs.filter(branch_code=chosen_code)
                else:
                    matching_codes = [
                        code for code, name in Branch.BRANCH_CODE_MAP.items()
                        if name.strip().upper() == chosen_branch.name.strip().upper()
                    ]
                    if matching_codes:
                        entries_qs = entries_qs.filter(branch_code__in=matching_codes)
                    else:
                        entries_qs = entries_qs.none()

        elif selected_unit_type == 'department' and selected_unit_id:
            try:
                chosen_dept = Department.objects.get(id=selected_unit_id)
                entries_qs = entries_qs.filter(category__iexact=chosen_dept.name)
            except Department.DoesNotExist:
                entries_qs = entries_qs.none()

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

        'overall_month_progress':   overall_month_progress,
        'overall_quarter_progress': overall_quarter_progress,
        'overall_year_progress':    overall_year_progress,
        'weekly_completion_rate':   weekly_completion_rate,
        'month_completed':   month_actual,
        'month_total':       month_expected,
        'quarter_completed': quarter_actual,
        'quarter_total':     quarter_expected,
        'year_completed':    year_actual,
        'year_total':        year_expected,
        'total_logged_days': total_logged_days,
        'frequency_summary': frequency_summary,

        'daily_checklists': display_checklists_with_progress[:5],
        'activity_logs': activity_logs,

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

        'unit_filter_options':   unit_filter_options,
        'active_unit_filter':    unit_filter_raw,
        'selected_unit_display': selected_unit_display,
        'selected_unit_type':    selected_unit_type,

        'month_ring_offset':   100 - overall_month_progress,
        'quarter_ring_offset': 100 - overall_quarter_progress,
        'year_ring_offset':    100 - overall_year_progress,
        'weekly_ring_offset':  100 - weekly_completion_rate,
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

# ============================================================
# MY REPORTS — display builder for Excel uploads
# ============================================================
CANONICAL_EXCEPTION_HEADERS = [
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


def _build_exception_display(upload):
    """
    Build the {is_excel, headers, rows} display dict for an ExceptionUpload.

    Header strings and row keys come from the SAME literals, so
    `row|get_item:header` can never drift out of sync.
    """
    records = upload.exception_records.all().order_by('source_row_index', 'serial_number')

    rows = []
    for rec in records:
        rows.append({
            'row_id': rec.id,
            'S/N': rec.serial_number if rec.serial_number is not None else '',
            'BRANCH/UNIT': rec.branch_unit or '',
            'EXCEPTION': rec.exception or '',
            'DATE EXCEPTION WAS NOTED': (
                rec.date_noted.strftime('%Y-%m-%d')
                if rec.date_noted
                else (rec.date_noted_raw or '')
            ),
            'TARGET DATE FOR CLOSURE': (
                rec.target_closure_date.strftime('%Y-%m-%d')
                if rec.target_closure_date
                else (rec.target_closure_date_raw or '')
            ),
            'CATEGORY OF EXCEPTION': rec.category or '',
            'RESPONSIBLE OFFICER': rec.responsible_officer or '',
            'SUPERVISOR': rec.supervisor or '',
            "AUDITEE'S RESPONSE": rec.auditee_response or '',
            'REMARKS': rec.remarks or '',
            'STATUS': rec.get_status_display() or rec.status_raw or '',
            'INCOME/COST SAVED': (
                f'{rec.income_cost_saved:,.2f}'
                if rec.income_cost_saved is not None else ''
            ),
        })

    return {
        'is_excel': True,
        'headers': CANONICAL_EXCEPTION_HEADERS,
        'rows': rows,
    }


@login_required
def reports_page(request):
    """
    My Reports.

    Two distinct groups:
      • submitted_reports  — Report rows the member actually sent
      • exception_uploads  — ExceptionUpload rows (Excel data)
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    # ------------------------------------------------------------
    # Submitted reports — still on Report (recurring, submitted, draft)
    # ------------------------------------------------------------
    submitted_reports = Report.objects.filter(
        created_by=user_profile
    ).exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    ).filter(
        Q(status='submitted') |
        Q(status='completed') |
        Q(status='draft')
    ).order_by('-created_at')

    # ------------------------------------------------------------
    # Exception uploads — now from ExceptionUpload
    # ------------------------------------------------------------
    exception_uploads = (
        ExceptionUpload.objects
        .filter(uploaded_by=user_profile)
        .filter(row_count__gt=0)
        .filter(exception_records__isnull=False)
        .distinct()
        .order_by('-created_at')
    )

    # Attach the display dict
    for upload in exception_uploads:
        upload.display_data = _build_exception_display(upload)

    # ------------------------------------------------------------
    # Filter dropdown — union of both groups, with counts
    # ------------------------------------------------------------
    from collections import defaultdict

    type_counts = defaultdict(int)
    for rt in submitted_reports.values_list('report_type', flat=True):
        if rt:
            type_counts[rt] += 1
    for rt in exception_uploads.values_list('report_type', flat=True):
        if rt:
            type_counts[rt] += 1

    report_type_options = sorted(
        (
            {'value': rt, 'label': rt, 'count': cnt}
            for rt, cnt in type_counts.items()
        ),
        key=lambda o: o['label'].lower(),
    )

    report_types = [o['value'] for o in report_type_options]

    branches = Branch.objects.filter(is_active=True).order_by('name')

    # Attach schedule + display data for submitted reports
    for report in submitted_reports:
        schedule = ReportSchedule.objects.filter(report=report, is_active=True).first()
        report.has_schedule = bool(schedule)
        if schedule:
            report.next_due_date = schedule.next_due_date
            report.schedule_frequency = schedule.get_frequency_display()
        report._display_data = report.get_display_data()

    context = {
        'user_profile':         user_profile,
        'today':                timezone.now(),
        'reports':              submitted_reports,
        'exception_uploads':    exception_uploads,
        'report_types':         report_types,
        'report_type_options':  report_type_options,
        'branches':             branches,
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

    Creates one ExceptionUpload row + N ExceptionRecord rows.
    No Report rows are touched.
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

        # Create the dedicated upload container
        upload = ExceptionUpload.objects.create(
            uploaded_by=user_profile,
            report_type=report_type,
            file_name=file_name,
            excel_headers=','.join(str(h) for h in headers),
            row_count=0,
        )

        typed_count = 0
        try:
            typed_count = _build_exception_records(
                upload=upload,
                headers=headers,
                rows_data=rows_data,
            )
        except Exception as exc:
            logger.exception("Failed to build typed ExceptionRecord rows: %s", exc)

        upload.row_count = typed_count
        upload.save(update_fields=['row_count'])

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
            'upload_id': upload.id,
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

    # -----------------------------------------------------------------
    # Filter params
    # -----------------------------------------------------------------
    branch_filter     = request.GET.get('branch', 'all')
    department_filter = request.GET.get('department', 'all')
    month_filter      = request.GET.get('month', 'all')       # 0-based index
    year_filter       = request.GET.get('year', 'all')
    frequency_filter  = request.GET.get('frequency', 'all')
    quarter_filter    = request.GET.get('quarter', 'all')

    user_branches     = user_profile.branches.all().order_by('name')
    user_departments  = user_profile.departments.all().order_by('name')

    user_branch_ids    = set(user_branches.values_list('id', flat=True))
    user_department_ids = set(user_departments.values_list('id', flat=True))

    # -----------------------------------------------------------------
    # Base checklist queryset
    # -----------------------------------------------------------------
    user_checklists = Checklist.objects.filter(
        is_active=True
    ).filter(
        Q(assigned_branches__in=user_branch_ids) |
        Q(assigned_departments__in=user_department_ids)
    ).distinct()

    if branch_filter != 'all':
        try:
            user_checklists = user_checklists.filter(
                assigned_branches__id=int(branch_filter)
            ).distinct()
        except (ValueError, TypeError):
            pass

    if department_filter != 'all':
        try:
            user_checklists = user_checklists.filter(
                assigned_departments__id=int(department_filter)
            ).distinct()
        except (ValueError, TypeError):
            pass

    if frequency_filter != 'all':
        user_checklists = user_checklists.filter(frequency=frequency_filter)

    # =================================================================
    # Resolve the "active period"
    # Priority: explicit month → quarter (first month of quarter) → today
    # =================================================================
    display_year = today.year
    if year_filter != 'all':
        try:
            display_year = int(year_filter)
        except (ValueError, TypeError):
            display_year = today.year

    display_month = today.month
    if month_filter != 'all':
        try:
            display_month = int(month_filter) + 1   # UI sends 0-based
        except (ValueError, TypeError):
            display_month = today.month
    elif quarter_filter != 'all':
        try:
            q = int(quarter_filter)
            if 1 <= q <= 4:
                display_month = (q - 1) * 3 + 1
        except (ValueError, TypeError):
            display_month = today.month

    display_month = max(1, min(12, display_month))
    display_year  = max(2000, min(2100, display_year))

    current_quarter = (display_month - 1) // 3 + 1

    # Selected label used in the KPI subtitle
    selected_month_label = date(display_year, display_month, 1).strftime('%B')

    # -----------------------------------------------------------------
    # Period windows derived from the active month
    # -----------------------------------------------------------------
    month_start = date(display_year, display_month, 1)
    if display_month == 12:
        month_end = date(display_year + 1, 1, 1) - timedelta(days=1)
    else:
        month_end = date(display_year, display_month + 1, 1) - timedelta(days=1)

    quarter_start_month = (current_quarter - 1) * 3 + 1
    quarter_end_month   = current_quarter * 3

    quarter_start = date(display_year, quarter_start_month, 1)
    if quarter_end_month == 12:
        quarter_end = date(display_year + 1, 1, 1) - timedelta(days=1)
    else:
        quarter_end = date(display_year, quarter_end_month + 1, 1) - timedelta(days=1)

    year_start = date(display_year, 1, 1)
    year_end   = date(display_year, 12, 31)

    def count_actual(checklist, unit_filter, period_start, period_end):
        return ChecklistLog.objects.filter(
            checklist=checklist,
            user=user_profile,
            log_date__gte=period_start,
            log_date__lte=period_end,
            **unit_filter
        ).values('log_date').distinct().count()

    # -----------------------------------------------------------------
    # Build the row data
    # -----------------------------------------------------------------
    checklist_data = []
    total_month_expected = 0
    total_month_actual = 0
    total_quarter_expected = 0
    total_quarter_actual = 0
    total_year_expected = 0
    total_year_actual = 0

    for checklist in user_checklists:
        tasks = list(checklist.tasks.all().order_by('order'))

        checklist_branches    = checklist.assigned_branches.all()
        checklist_departments = checklist.assigned_departments.all()

        if branch_filter != 'all':
            try:
                aid = int(branch_filter)
                checklist_branches = (
                    checklist_branches.filter(id=aid)
                    if aid in user_branch_ids
                    else checklist_branches.none()
                )
            except (ValueError, TypeError):
                checklist_branches = checklist_branches.none()
        else:
            checklist_branches = checklist_branches.filter(id__in=user_branch_ids)

        if department_filter != 'all':
            try:
                aid = int(department_filter)
                checklist_departments = (
                    checklist_departments.filter(id=aid)
                    if aid in user_department_ids
                    else checklist_departments.none()
                )
            except (ValueError, TypeError):
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

        next_due = (
            checklist.get_next_due_date(user_profile)
            if hasattr(checklist, 'get_next_due_date') else None
        )

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
            month_actual   = count_actual(checklist, unit_filter, month_start, month_end)

            quarter_expected = count_expected_occurrences(checklist, quarter_start, quarter_end)
            quarter_actual   = count_actual(checklist, unit_filter, quarter_start, quarter_end)

            year_expected = count_expected_occurrences(checklist, year_start, year_end)
            year_actual   = count_actual(checklist, unit_filter, year_start, year_end)

            month_progress   = min(int(month_actual / month_expected * 100), 100) if month_expected > 0 else 0
            quarter_progress = min(int(quarter_actual / quarter_expected * 100), 100) if quarter_expected > 0 else 0
            year_progress    = min(int(year_actual / year_expected * 100), 100) if year_expected > 0 else 0

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

            total_month_expected   += month_expected
            total_month_actual     += month_actual
            total_quarter_expected += quarter_expected
            total_quarter_actual   += quarter_actual
            total_year_expected    += year_expected
            total_year_actual      += year_actual

    overall_month_progress   = min(int(total_month_actual / total_month_expected * 100), 100) if total_month_expected > 0 else 0
    overall_quarter_progress = min(int(total_quarter_actual / total_quarter_expected * 100), 100) if total_quarter_expected > 0 else 0
    overall_year_progress    = min(int(total_year_actual / total_year_expected * 100), 100) if total_year_expected > 0 else 0

    # -----------------------------------------------------------------
    # Filter dropdown options
    # -----------------------------------------------------------------
    available_branches = (
        user_branches
        .filter(checklist_assignments__in=user_checklists)
        .distinct().order_by('name')
    )
    available_departments = (
        user_departments
        .filter(checklist_assignments__in=user_checklists)
        .distinct().order_by('name')
    )

    combined_filter_options = [
        {'type': 'all', 'id': 'all', 'display': 'All Branches/Departments', 'value': 'all'}
    ]
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
        except (ValueError, TypeError):
            pass

    context = {
        'user_profile': user_profile,
        'checklists': user_checklists,
        'checklist_data': checklist_data,
        'years': list(range(timezone.now().year - 2, timezone.now().year + 1)),

        # KPI values, precomputed server-side
        'overall_month_progress':   overall_month_progress,
        'overall_quarter_progress': overall_quarter_progress,
        'overall_year_progress':    overall_year_progress,
        'total_month_actual':       total_month_actual,
        'total_month_expected':     total_month_expected,
        'total_quarter_actual':     total_quarter_actual,
        'total_quarter_expected':   total_quarter_expected,
        'total_year_actual':        total_year_actual,
        'total_year_expected':      total_year_expected,

        # Active period
        'display_month':         display_month,
        'display_year':          display_year,
        'current_quarter':       current_quarter,
        'selected_month_label':  selected_month_label,

        # Filters and labels
        'branch_filter':      branch_filter,
        'department_filter':  department_filter,
        'month_filter':       month_filter,
        'year_filter':        year_filter,
        'frequency_filter':   frequency_filter,
        'quarter_filter':     quarter_filter,

        'available_frequencies':       available_frequencies,
        'month_names':                 month_names,
        'quarter_options':             quarter_options,
        'selected_frequency_display':  selected_frequency_display,
        'selected_quarter_display':    selected_quarter_display,

        'today':      today,
        'month_name': selected_month_label,
        'year':       display_year,

        'combined_filter_options': combined_filter_options,
        'active_combined_filter':  active_combined_filter,
        'selected_combined_display': selected_combined_display,
    }

    return render(request, 'control_dashboard/checklist-mem.html', context)

# ==================== API - CHECKLIST LOG ====================

@csrf_exempt
@require_http_methods(["POST"])
def api_log_checklist(request):
    """
    Log or unlog a checklist for a given date.

    ONE-LOG-PER-PERIOD RULES:
      daily      → one log per calendar day (unchanged)
      weekly     → one log per Fri–Thu week
      monthly    → one log per calendar month
      quarterly  → one log per calendar quarter
      bi-annual  → one log per Jan–Jun or Jul–Dec block
      annual     → one log per calendar year

    Any attempt to log a second day inside the same period is
    rejected with HTTP 409 and a helpful `already_logged_on` date.
    """
    try:
        data = json.loads(request.body)

        checklist_id = data.get('checklist_id')
        log_date     = data.get('log_date')
        action       = data.get('action', 'log')
        unit_type    = data.get('unit_type', 'general')
        unit_id      = data.get('unit_id')

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

        user_branch_ids = set(user_profile.branches.values_list('id', flat=True))
        user_dept_ids   = set(user_profile.departments.values_list('id', flat=True))

        is_assigned = (
            checklist.assigned_branches.filter(id__in=user_branch_ids).exists() or
            checklist.assigned_departments.filter(id__in=user_dept_ids).exists()
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

        # ------------------------------------------------------------------
        # EVERY log for this (checklist, user, unit) tuple — used to detect
        # whether any other day in the same period has already been logged.
        # ------------------------------------------------------------------
        base_qs = ChecklistLog.objects.filter(
            checklist=checklist,
            user=user_profile,
            branch=branch,
            department=department,
        )

        freq = checklist.frequency or 'daily'

        # ============================================================
        # Period window for the *candidate* date
        # ============================================================
        period_start = None
        period_end   = None

        if freq == 'weekly':
            # Friday-based week (single source of truth)
            week_start = get_week_start(date_obj)
            period_start = week_start
            period_end   = week_start + timedelta(days=6)

        elif freq == 'monthly':
            period_start = date(date_obj.year, date_obj.month, 1)
            if date_obj.month == 12:
                period_end = date(date_obj.year + 1, 1, 1) - timedelta(days=1)
            else:
                period_end = date(date_obj.year, date_obj.month + 1, 1) - timedelta(days=1)

        elif freq == 'quarterly':
            q_start_month = ((date_obj.month - 1) // 3) * 3 + 1
            q_end_month   = q_start_month + 2
            period_start = date(date_obj.year, q_start_month, 1)
            if q_end_month == 12:
                period_end = date(date_obj.year + 1, 1, 1) - timedelta(days=1)
            else:
                period_end = date(date_obj.year, q_end_month + 1, 1) - timedelta(days=1)

        elif freq == 'bi-annual':
            if date_obj.month <= 6:
                period_start = date(date_obj.year, 1, 1)
                period_end   = date(date_obj.year, 6, 30)
            else:
                period_start = date(date_obj.year, 7, 1)
                period_end   = date(date_obj.year, 12, 31)

        elif freq == 'annual':
            period_start = date(date_obj.year, 1, 1)
            period_end   = date(date_obj.year, 12, 31)

        # ============================================================
        # ACTION: log
        # ============================================================
        if action == 'log':
            # For non-daily frequencies, if any OTHER day in the same
            # period is already logged, reject with 409.
            if period_start is not None and period_end is not None:
                existing_in_period = base_qs.filter(
                    log_date__gte=period_start,
                    log_date__lte=period_end,
                ).exclude(log_date=date_obj).first()

                if existing_in_period is not None:
                    return JsonResponse({
                        'success': False,
                        'error': (
                            f'This is a {freq} activity. '
                            f'It has already been logged for this period '
                            f'on {existing_in_period.log_date.isoformat()}.'
                        ),
                        'already_logged_on': existing_in_period.log_date.isoformat(),
                        'period': freq,
                        'period_start': period_start.isoformat(),
                        'period_end': period_end.isoformat(),
                    }, status=409)

            log_entry, created = ChecklistLog.objects.get_or_create(
                checklist=checklist,
                user=user_profile,
                branch=branch,
                department=department,
                log_date=date_obj,
            )

            if created:
                return JsonResponse({
                    'success': True,
                    'message': 'Checklist logged successfully',
                    'action': 'logged',
                    'period': freq,
                })
            else:
                return JsonResponse({
                    'success': True,
                    'message': 'Checklist already logged for this date',
                    'action': 'already_logged',
                    'period': freq,
                })

        # ============================================================
        # ACTION: unlog
        # ============================================================
        elif action == 'unlog':
            deleted_count, _ = base_qs.filter(log_date=date_obj).delete()

            if deleted_count > 0:
                return JsonResponse({
                    'success': True,
                    'message': 'Checklist unlogged successfully',
                    'action': 'unlogged',
                })
            else:
                return JsonResponse({
                    'success': False,
                    'error': 'No log found for this date',
                }, status=404)

        else:
            return JsonResponse({'success': False, 'error': 'Invalid action. Use "log" or "unlog"'}, status=400)

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data'}, status=400)
    except Exception as e:
        logger.exception("Error in api_log_checklist")
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

@csrf_exempt
@require_http_methods(["GET"])
def api_get_checklist_stats(request):
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)

        user_branch_ids = user_profile.branches.values_list('id', flat=True)
        user_dept_ids   = user_profile.departments.values_list('id', flat=True)

        user_checklists = Checklist.objects.filter(
            is_active=True
        ).filter(
            Q(assigned_branches__in=user_branch_ids) |
            Q(assigned_departments__in=user_dept_ids)
        ).distinct()

        total_checklists = user_checklists.count()
        logs = ChecklistLog.objects.filter(user=user_profile)
        total_logs = logs.count()

        today = timezone.now().date()
        start_of_week = get_week_start(today)
        end_of_week = start_of_week + timedelta(days=6)

        weekly_logs = logs.filter(log_date__gte=start_of_week, log_date__lte=end_of_week).count()
        recent_logs = logs.filter(log_date__gte=today - timedelta(days=7)).count()

        daily_checklists = user_checklists.filter(frequency='daily')
        daily_total = daily_checklists.count()

        completion_rate = 0
        if daily_total > 0:
            completed_days = 0
            for i in range(7):
                day = start_of_week + timedelta(days=i)
                if day > today:
                    break
                if logs.filter(log_date=day).count() >= daily_total:
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
    Team Performance — score derived from SentEmail history.

    For every expected occurrence of an assigned Report, the member
    can earn up to 100 points (see _build_member_scorecard). Missed
    deadlines, late emails past the 3-day grace window, and supervisor
    deductions all pull the score down. AdHoc adjustments move it too.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    today = timezone.now()

    team_members = (
        UserProfile.objects
        .filter(role='member', status='active')
        .order_by('full_name')
    )

    team_data = []
    total_members = 0
    grand_total_debits = 0
    grand_total_credits = 0
    sum_of_percentages = 0

    for member in team_members:
        total_members += 1
        card = _build_member_scorecard(member, today=today)
        team_data.append(card)

        grand_total_debits  += card['debits']
        grand_total_credits += card['credits']
        sum_of_percentages  += card['percentage']

    team_data.sort(
        key=lambda x: (x['percentage'], x['net_score'], x['email_count']),
        reverse=True,
    )

    overall_completion = (
        min(100, int(round(sum_of_percentages / total_members)))
        if total_members else 0
    )

    context = {
        'user_profile': user_profile,
        'team_data': team_data,
        'total_members': total_members,
        'total_debits': grand_total_debits,
        'total_credits': grand_total_credits,
        'overall_completion': overall_completion,
    }

    return render(request, 'control_dashboard/team.html', context)

@csrf_exempt
@require_http_methods(["GET"])
def api_team_performance_live(request):
    """
    Live JSON endpoint for the Team Performance auto-refresh.

    Returns the same scorecard computed by _build_member_scorecard,
    one entry per active member.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

    today = timezone.now()

    team_members = (
        UserProfile.objects
        .filter(role='member', status='active')
        .order_by('full_name')
    )

    members_payload = []
    for member in team_members:
        card = _build_member_scorecard(member, today=today)
        members_payload.append({
            'id':              member.id,
            'full_name':       member.full_name or member.email,
            'email_count':     card['email_count'],
            'debits':          card['debits'],
            'credits':         card['credits'],
            'net_score':       card['net_score'],
            'percentage':      card['percentage'],
        })

    return JsonResponse({'success': True, 'members': members_payload})

@login_required
def submitted_reports(request):
    """
    Submitted Reports (supervisor view).

    Each row = ONE email that was sent (a SentEmail record).
    Filters: user, category/report_type, date range on sent_at.

    For each email we attach:
      • the sender (from SentEmail.sender)
      • the matching Report (for the deadline) — matched by report_type
      • lateness = sent_at vs (deadline_date + deadline_time)
      • deduction from SentEmail.manual_deduction
      • final_score = 100 - deduction
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    user_filter     = request.GET.get('user', 'all')
    category_filter = request.GET.get('category', 'all')
    start_date      = request.GET.get('start_date', '')
    end_date        = request.GET.get('end_date', '')

    # ── Base queryset: every email that was sent ─────────────
    emails_qs = (
        SentEmail.objects
        .select_related('sender')
        .order_by('-sent_at')
    )

    if user_filter != 'all':
        try:
            emails_qs = emails_qs.filter(sender_id=int(user_filter))
        except ValueError:
            pass

    if category_filter != 'all':
        emails_qs = emails_qs.filter(report_type=category_filter)

    if start_date:
        try:
            start_d = datetime.strptime(start_date, '%Y-%m-%d').date()
            emails_qs = emails_qs.filter(sent_at__date__gte=start_d)
        except ValueError:
            pass

    if end_date:
        try:
            end_d = datetime.strptime(end_date, '%Y-%m-%d').date()
            emails_qs = emails_qs.filter(sent_at__date__lte=end_d)
        except ValueError:
            pass

    # ── Build a report_type → Report map for deadline lookup ─
    # If multiple Reports share a report_type, prefer the most
    # recently created (which is usually the active one).
    report_type_list = (
        emails_qs
        .values_list('report_type', flat=True)
        .distinct()
    )
    report_lookup = {}
    for rt in report_type_list:
        rpt = (
            Report.objects
            .filter(report_type=rt)
            .exclude(report_type=TRIAL_BALANCE_REPORT_TYPE)
            .order_by('-created_at')
            .first()
        )
        if rpt:
            report_lookup[rt] = rpt

    # ── Build the row data ───────────────────────────────────
    report_data = []
    for email in emails_qs:
        matching_report = report_lookup.get(email.report_type)

        # Compute the deadline as an aware datetime, if we have one
        deadline_at = None
        if matching_report and matching_report.deadline_date:
            deadline_time = (
                matching_report.deadline_time
                or datetime.strptime('23:59', '%H:%M').time()
            )
            deadline_at = timezone.make_aware(
                datetime.combine(matching_report.deadline_date, deadline_time)
            )

        # Compute lateness
        minutes_late = 0
        if deadline_at and email.sent_at:
            delta = email.sent_at - deadline_at
            if delta.total_seconds() > 0:
                minutes_late = int(delta.total_seconds() // 60)

        report_data.append({
            'email': email,
            'created_by': email.sender,
            'report_type': email.report_type,
            'submitted_at': email.sent_at,
            'deadline_at': deadline_at,
            'minutes_late': minutes_late,
            'report': matching_report,
            'deduction': email.manual_deduction or 0,
            'final_score': email.final_score,
            'badge_class': email.badge_class,
            'is_overridden': email.is_overridden,
            'recipient_count': email.recipient_count,
        })

    users = (
        UserProfile.objects
        .filter(role='member', status='active')
        .order_by('full_name')
    )

    # Category dropdown — every distinct report_type we've emailed
    categories = (
        SentEmail.objects
        .values_list('report_type', flat=True)
        .distinct()
        .order_by('report_type')
    )

    context = {
        'user_profile':     user_profile,
        'report_data':      report_data,
        'users':            users,
        'categories':       categories,
        'user_filter':      user_filter,
        'category_filter':  category_filter,
        'start_date':       start_date,
        'end_date':         end_date,
        'total_reports':    len(report_data),
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
    Logged Exceptions (supervisor view) — now reads from ExceptionUpload.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    type_filter = request.GET.get('type', 'all')
    start_date = request.GET.get('start_date', '')
    end_date = request.GET.get('end_date', '')

    exceptions = ExceptionUpload.objects.filter(
        row_count__gt=0,
    ).order_by('-created_at')

    if type_filter != 'all':
        exceptions = exceptions.filter(report_type=type_filter)

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

    report_types = ExceptionUpload.objects.filter(
        row_count__gt=0,
    ).values_list('report_type', flat=True).distinct()

    context = {
        'user_profile': user_profile,
        'exceptions': exceptions,
        'report_types': report_types,
        'type_filter': type_filter,
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
        member_branch_ids = set(member.branches.values_list('id', flat=True))
        member_dept_ids = set(member.departments.values_list('id', flat=True))

        member_checklists = all_checklists.filter(
            Q(assigned_users=member) |
            Q(assignment_target='all') |
            Q(assignment_target=member.position) |
            Q(assigned_departments__in=member_dept_ids) |
            Q(assigned_branches__in=member_branch_ids)
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

        member_branch_ids = set(user.branches.values_list('id', flat=True))
        member_dept_ids = set(user.departments.values_list('id', flat=True))

        member_checklists = Checklist.objects.filter(
            is_active=True
        ).filter(
            Q(assigned_users=user) |
            Q(assignment_target='all') |
            Q(assignment_target=user.position) |
            Q(assigned_departments__in=member_dept_ids) |
            Q(assigned_branches__in=member_branch_ids)
        ).distinct()

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
    All Submitted Exceptions (supervisor view).

    Shows every exception submitted by every user, grouped by
    uploader. Admins / supervisors see everyone; members see only
    their own.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    if user_profile.role not in ('admin', 'supervisor', 'member'):
        messages.error(request, 'You do not have permission to access this page.')
        return redirect_dashboard(request.user)

    from decimal import Decimal
    from collections import OrderedDict

    today = timezone.now().date()
    is_privileged = user_profile.role in ('admin', 'supervisor')

    base_qs = ExceptionRecord.objects.filter(upload__isnull=False)

    # Members see only their own uploads.
    if not is_privileged:
        base_qs = base_qs.filter(upload__uploaded_by=user_profile)

    # ── Filters ────────────────────────────────────────────────
    uploader_filter = request.GET.get('uploader', 'all').strip()
    branch_filter   = request.GET.get('branch', 'all').strip()
    category_filter = request.GET.get('category', 'all').strip()
    status_filter   = request.GET.get('status', 'all').strip()
    start_date_str  = request.GET.get('start_date', '').strip()
    end_date_str    = request.GET.get('end_date', '').strip()

    filtered_qs = base_qs

    if uploader_filter and uploader_filter != 'all':
        try:
            filtered_qs = filtered_qs.filter(upload__uploaded_by_id=int(uploader_filter))
        except ValueError:
            pass

    if branch_filter and branch_filter != 'all':
        filtered_qs = filtered_qs.filter(branch_unit=branch_filter)

    if category_filter and category_filter != 'all':
        filtered_qs = filtered_qs.filter(category=category_filter)

    if status_filter and status_filter != 'all':
        filtered_qs = filtered_qs.filter(status=status_filter)

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

    # ── Group by uploader ──────────────────────────────────────
    filtered_qs = (
        filtered_qs
        .select_related('upload', 'upload__uploaded_by')
        .order_by('upload__uploaded_by__full_name', 'upload__created_at', 'source_row_index')
    )

    grouped = OrderedDict()
    for rec in filtered_qs:
        uploader = getattr(rec.upload, 'uploaded_by', None) if rec.upload else None
        if not uploader:
            continue
        key = uploader.id
        if key not in grouped:
            grouped[key] = {
                'uploader': uploader,
                'count': 0,
                'rows': [],
            }
        grouped[key]['count'] += 1
        grouped[key]['rows'].append(rec)

    uploader_groups = list(grouped.values())

    # ── Filter dropdown options ────────────────────────────────
    available_uploaders = (
        UserProfile.objects
        .filter(exception_uploads__isnull=False)
        .distinct()
        .order_by('full_name')
    )

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

    available_statuses = [
        ('open', 'Open'),
        ('in_progress', 'In Progress'),
        ('pending', 'Pending'),
        ('overdue', 'Overdue'),
        ('closed', 'Closed'),
        ('resolved', 'Resolved'),
    ]

    context = {
        'user_profile': user_profile,
        'today': today,
        'is_privileged': is_privileged,

        'uploader_groups': uploader_groups,

        'available_uploaders': available_uploaders,
        'available_branches': available_branches,
        'available_categories': available_categories,
        'available_statuses': available_statuses,

        'uploader_filter': uploader_filter,
        'branch_filter': branch_filter,
        'category_filter': category_filter,
        'status_filter': status_filter,
        'start_date': start_date_str,
        'end_date': end_date_str,
    }

    template = (
        'control_dashboard/report-sup.html'
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

        # One trial balance per day: wipe any existing upload for that date
        existing = TrialBalanceUpload.objects.filter(report_date=report_date).first()
        if existing:
            existing.delete()

        upload = TrialBalanceUpload.objects.create(
            uploaded_by=user_profile,
            report_date=report_date,
            file_name=file_name,
            row_count=0,
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
            entries.append(TrialBalanceEntry(
                upload=upload,                                       # <-- CHANGED
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
            ))

        TrialBalanceEntry.objects.bulk_create(entries, batch_size=500)

        upload.row_count = len(entries)
        upload.save(update_fields=['row_count'])

        log_activity(
            user=user_profile,
            activity_type='report_submitted',
            details=f'Submitted Daily Trial Balance for {report_date} with {len(entries)} entries',
            request=request
        )

        return JsonResponse({
            'success': True,
            'message': f'Successfully saved {len(entries)} trial balance entries for {report_date}.',
            'upload_id': upload.id,
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
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

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

    selected_report_type = (request.GET.get('report_type') or '').strip()
    if not selected_report_type or selected_report_type == 'all':
        if available_report_types:
            selected_report_type = available_report_types[0]

    if selected_report_type not in available_report_types:
        selected_report_type = available_report_types[0] if available_report_types else ''

    def _derive_position_scope(report_type):
        name = (report_type or '').lower()
        if 'head office' in name or 'headoffice' in name or 'head-office' in name:
            return 'hc'
        if 'cluster' in name:
            return 'cc'
        return None

    uploader_scope = _derive_position_scope(selected_report_type)

    exception_rows = []
    total_count = 0
    total_cost = 0
    distinct_uploaders = 0

    if has_access and selected_report_type:
        er_qs = ExceptionRecord.objects.filter(
            upload__isnull=False,
        )

        if uploader_scope:
            er_qs = er_qs.filter(
                upload__uploaded_by__position=uploader_scope,
            )
        else:
            er_qs = er_qs.filter(
                upload__report_type=selected_report_type,
            )

        er_qs = er_qs.select_related(
            'upload', 'upload__uploaded_by'
        ).order_by('-upload__created_at', 'source_row_index')

        total_count = er_qs.count()

        cost_agg = er_qs.aggregate(total=Sum('income_cost_saved'))
        try:
            total_cost = float(cost_agg.get('total') or 0)
        except (TypeError, ValueError):
            total_cost = 0.0

        distinct_uploaders = er_qs.values('upload__uploaded_by').distinct().count()

        for rec in er_qs[:100]:
            # ---- Resolve the report's uploader ----
            uploader = getattr(rec.upload, 'uploaded_by', None) if rec.upload else None
            if uploader:
                uploader_name = uploader.full_name or uploader.email or ''
                uploader_initials = _initials_from_name(uploader_name)
            else:
                uploader_name = ''
                uploader_initials = ''

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
                'source_report': (rec.upload.report_type if rec.upload else '—'),
                # ---- New uploader fields for the template ----
                'uploaded_by_name': uploader_name,
                'uploaded_by_initials': uploader_initials,
                'uploaded_at': rec.upload.created_at if rec.upload else None,
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
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        messages.error(request, 'User profile not found.')
        return redirect('control_dashboard:consolidated_reports')

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

    def _derive_position_scope(report_type):
        name = (report_type or '').lower()
        if 'head office' in name or 'headoffice' in name or 'head-office' in name:
            return 'hc'
        if 'cluster' in name:
            return 'cc'
        return None

    uploader_scope = _derive_position_scope(report_type_filter)

    er_qs = ExceptionRecord.objects.filter(
        upload__isnull=False,
    )

    if uploader_scope:
        er_qs = er_qs.filter(
            upload__uploaded_by__position=uploader_scope,
        )
    else:
        er_qs = er_qs.filter(
            upload__report_type=report_type_filter,
        )

    er_qs = er_qs.select_related(
        'upload', 'upload__uploaded_by'
    ).order_by('-upload__created_at', 'source_row_index')

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
        ws.cell(row=row_idx, column=13, value=(rec.upload.report_type if rec.upload else ''))
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

                # ------------------------------------------------------------
        # Whitelist of 8 columns to expose to the frontend.
        # Keep this in sync with the exception side view.
        # ------------------------------------------------------------
        headers = [
            'BRANCH_CODE',
            'CATEGORY',
            'GL_CODE',
            'DESCR',
            'CCY',
            'CLOSE_BAL_LCY',
            'CLOSE_BAL_FCY',
            'GL_STATUS',
        ]

        rows = []
        total_debit = 0.0
        total_credit = 0.0

        for entry in visible_entries:
            # -------- Filter: only active GLs --------
            status_raw = str(entry.gl_status or '').strip().lower()
            if status_raw != 'active':
                continue

            rows.append({
                'BRANCH_CODE': entry.branch_code or '',
                'CATEGORY': entry.category or '',
                'GL_CODE': entry.gl_code or '',
                'DESCR': entry.descr or '',
                'CCY': entry.ccy or '',
                'CLOSE_BAL_LCY': str(entry.close_bal_lcy),
                'CLOSE_BAL_FCY': str(entry.close_bal_fcy),
                'GL_STATUS': entry.gl_status or '',
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

        upload = (
            TrialBalanceUpload.objects
            .filter(report_date=report_date)
            .select_related('uploaded_by')
            .first()
        )

        if not upload:
            return JsonResponse({
                'success': True,
                'locked': False,
                'report_date': report_date.isoformat(),
                'uploaded_by': None,
                'uploaded_at': None,
                'row_count': 0,
            })

        return JsonResponse({
            'success': True,
            'locked': True,
            'report_date': report_date.isoformat(),
            'uploaded_by': {
                'id': upload.uploaded_by.id,
                'full_name': upload.uploaded_by.full_name or upload.uploaded_by.email,
                'email': upload.uploaded_by.email,
            },
            'uploaded_at': upload.created_at.isoformat() if upload.created_at else None,
            'row_count': upload.row_count,
        })

    except Exception as e:
        logger.exception("Error in api_check_trial_balance_day")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

# ============================================================
# COMPARISON APIs — previous-day & date-range side-by-side
# ============================================================

COMPARISON_HEADERS = [
    'BRANCH_CODE',
    'CATEGORY',
    'GL_CODE',
    'DESCR',
    'CCY',
    'CLOSE_BAL_LCY',
    'CLOSE_BAL_FCY',
    'GL_STATUS',
]


def _serialize_tb_entry(entry):
    """Serialize a single TrialBalanceEntry for side-by-side comparison."""
    return {
        'BRANCH_CODE': entry.branch_code or '',
        'CATEGORY': entry.category or '',
        'GL_CODE': entry.gl_code or '',
        'DESCR': entry.descr or '',
        'CCY': entry.ccy or '',
        'CLOSE_BAL_LCY': str(entry.close_bal_lcy),
        'CLOSE_BAL_FCY': str(entry.close_bal_fcy),
        'GL_STATUS': entry.gl_status or '',
    }


def _visible_entries_for_date(user_profile, report_date):
    """
    Return TrialBalanceEntry queryset for a date, respecting the
    same branch/department filtering rules as api_get_trial_balance_by_date.
    Only returns ACTIVE GLS.
    """
    from .models import TrialBalanceEntry

    code_to_name = dict(Branch.BRANCH_CODE_MAP)
    for b in Branch.objects.filter(is_active=True):
        if b.code:
            code_to_name[str(b.code).strip()] = b.name

    entries = list(
        TrialBalanceEntry.objects
        .filter(report_date=report_date)
        .order_by('row_index')
    )

    if not entries:
        return []

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

    def passes(entry):
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

    visible = [e for e in entries if passes(e)]

    # Only active GLs
    visible = [e for e in visible if str(e.gl_status or '').strip().lower() == 'active']

    return visible


@csrf_exempt
@require_http_methods(["GET"])
def api_compare_trial_balance_previous(request):
    """
    Compare a trial balance against the previous upload.

    Query params:
        date    (required)  — YYYY-MM-DD of the "current" day

    Response:
        {
          "success": true,
          "current": { "date": "...", "found": bool, "rows": [...] },
          "previous": { "date": "...", "found": bool, "rows": [...] },
          "merged": [
             { "key": "BRANCH|GL_CODE", "BRANCH_CODE": ..., ...,
               "CURRENT_LCY": ..., "PREVIOUS_LCY": ..., "DELTA_LCY": ...,
               "CURRENT_FCY": ..., "PREVIOUS_FCY": ..., "DELTA_FCY": ... }
          ]
        }
    """
    try:
        date_str = (request.GET.get('date') or '').strip()
        if not date_str:
            return JsonResponse({'success': False, 'error': 'date parameter is required.'}, status=400)

        try:
            current_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            return JsonResponse({'success': False, 'error': 'Invalid date format. Use YYYY-MM-DD.'}, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

        from .models import TrialBalanceEntry

        # ---- Find the previous upload before current_date ----
        prev_upload = (
            TrialBalanceUpload.objects
            .filter(report_date__lt=current_date)
            .order_by('-report_date')
            .first()
        )
        previous_date = prev_upload.report_date if prev_upload else None

        current_visible = _visible_entries_for_date(user_profile, current_date)
        previous_visible = (
            _visible_entries_for_date(user_profile, previous_date)
            if previous_date else []
        )

        # ---- Build lookups keyed by (BRANCH_CODE, GL_CODE) ----
        def to_key(e):
            return f"{str(e.branch_code or '').strip()}|{str(e.gl_code or '').strip()}"

        current_map = {to_key(e): e for e in current_visible}
        previous_map = {to_key(e): e for e in previous_visible}

        all_keys = set(current_map.keys()) | set(previous_map.keys())

        merged = []
        for key in sorted(all_keys):
            cur = current_map.get(key)
            prev = previous_map.get(key)
            base = cur or prev  # for metadata

            def _dec(v):
                try:
                    return float(v or 0)
                except (TypeError, ValueError):
                    return 0.0

            cur_lcy = _dec(cur.close_bal_lcy) if cur else 0.0
            prev_lcy = _dec(prev.close_bal_lcy) if prev else 0.0
            cur_fcy = _dec(cur.close_bal_fcy) if cur else 0.0
            prev_fcy = _dec(prev.close_bal_fcy) if prev else 0.0

            merged.append({
                'key': key,
                'BRANCH_CODE': base.branch_code or '',
                'CATEGORY': base.category or '',
                'GL_CODE': base.gl_code or '',
                'DESCR': base.descr or '',
                'CCY': base.ccy or '',
                'GL_STATUS': base.gl_status or '',
                'CURRENT_LCY': cur_lcy,
                'PREVIOUS_LCY': prev_lcy,
                'DELTA_LCY': cur_lcy - prev_lcy,
                'CURRENT_FCY': cur_fcy,
                'PREVIOUS_FCY': prev_fcy,
                'DELTA_FCY': cur_fcy - prev_fcy,
                'IN_CURRENT':  cur is not None,
                'IN_PREVIOUS': prev is not None,
            })

        return JsonResponse({
            'success': True,
            'current': {
                'date': current_date.isoformat(),
                'found': len(current_visible) > 0,
                'row_count': len(current_visible),
                'rows': [_serialize_tb_entry(e) for e in current_visible],
            },
            'previous': {
                'date': previous_date.isoformat() if previous_date else None,
                'found': len(previous_visible) > 0,
                'row_count': len(previous_visible),
                'rows': [_serialize_tb_entry(e) for e in previous_visible],
            },
            'merged': merged,
            'headers': COMPARISON_HEADERS,
        })

    except Exception as e:
        logger.exception("Error in api_compare_trial_balance_previous")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@csrf_exempt
@require_http_methods(["GET"])
def api_compare_trial_balance_range(request):
    """
    Compare a trial balance across a date range.

    Query params:
        start   (required)  — YYYY-MM-DD
        end     (required)  — YYYY-MM-DD (inclusive)

    Response:
        {
          "success": true,
          "dates":   ["2026-09-20", "2026-09-21", ...],
          "merged":  [
            {
              "key": "BRANCH|GL_CODE",
              "BRANCH_CODE": ..., "GL_CODE": ..., "DESCR": ...,
              "balances": {
                 "2026-09-20": {"lcy": 1234.5, "fcy": 100.0},
                 "2026-09-21": {"lcy": 1300.0, "fcy": 105.0}
              }
            }
          ]
        }
    """
    try:
        start_str = (request.GET.get('start') or '').strip()
        end_str   = (request.GET.get('end') or '').strip()

        if not start_str or not end_str:
            return JsonResponse({'success': False, 'error': 'start and end parameters are required.'}, status=400)

        try:
            start_date = datetime.strptime(start_str, '%Y-%m-%d').date()
            end_date   = datetime.strptime(end_str,   '%Y-%m-%d').date()
        except ValueError:
            return JsonResponse({'success': False, 'error': 'Invalid date format. Use YYYY-MM-DD.'}, status=400)

        if end_date < start_date:
            start_date, end_date = end_date, start_date

        # Cap range to avoid runaway queries
        if (end_date - start_date).days > 90:
            return JsonResponse({
                'success': False,
                'error': 'Date range cannot exceed 90 days.',
            }, status=400)

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

        # ---- Collect all dates in the range that have a TB upload ----
        upload_dates = list(
            TrialBalanceUpload.objects
            .filter(report_date__gte=start_date, report_date__lte=end_date)
            .values_list('report_date', flat=True)
            .distinct()
            .order_by('report_date')
        )

        if not upload_dates:
            return JsonResponse({
                'success': True,
                'dates': [],
                'merged': [],
                'message': 'No trial balance uploads in that range.',
            })

        # ---- Fetch each day, build a per-day map keyed by BRANCH|GL_CODE ----
        per_day = {}  # { date_iso: { key: entry } }
        metadata = {} # { key: { BRANCH_CODE, CATEGORY, GL_CODE, DESCR, CCY } }

        for d in upload_dates:
            entries = _visible_entries_for_date(user_profile, d)
            day_map = {}
            for e in entries:
                k = f"{str(e.branch_code or '').strip()}|{str(e.gl_code or '').strip()}"
                day_map[k] = e
                if k not in metadata:
                    metadata[k] = {
                        'BRANCH_CODE': e.branch_code or '',
                        'CATEGORY': e.category or '',
                        'GL_CODE': e.gl_code or '',
                        'DESCR': e.descr or '',
                        'CCY': e.ccy or '',
                    }
            per_day[d.isoformat()] = day_map

        date_isos = [d.isoformat() for d in upload_dates]

        # ---- Merge across days ----
        all_keys = set()
        for day_map in per_day.values():
            all_keys.update(day_map.keys())

        merged = []
        for k in sorted(all_keys):
            meta = metadata.get(k, {})
            balances = {}
            for d_iso in date_isos:
                entry = per_day.get(d_iso, {}).get(k)
                if entry:
                    try:
                        lcy = float(entry.close_bal_lcy or 0)
                    except (TypeError, ValueError):
                        lcy = 0.0
                    try:
                        fcy = float(entry.close_bal_fcy or 0)
                    except (TypeError, ValueError):
                        fcy = 0.0
                    balances[d_iso] = {'lcy': lcy, 'fcy': fcy}
                else:
                    balances[d_iso] = None

            merged.append({
                'key': k,
                **meta,
                'balances': balances,
            })

        return JsonResponse({
            'success': True,
            'dates': date_isos,
            'merged': merged,
            'headers': COMPARISON_HEADERS,
        })

    except Exception as e:
        logger.exception("Error in api_compare_trial_balance_range")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@csrf_exempt
@require_http_methods(["POST"])
def api_update_report_score(request):
    """
    Supervisor deduction override on a SINGLE sent email.

    Payload:
    {
        "email_id": 123,
        "deduction": 15
    }

    Sets SentEmail.manual_deduction. final_score is derived
    as 100 - deduction.
    """
    try:
        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if user_profile.role not in ('supervisor', 'admin'):
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        data = json.loads(request.body)
        email_id  = data.get('email_id')
        deduction = data.get('deduction', 0)
        reason    = (data.get('reason') or '').strip()

        try:
            deduction = int(deduction)
        except (ValueError, TypeError):
            return JsonResponse({'success': False, 'error': 'Invalid deduction value'}, status=400)

        if deduction < 0 or deduction > 100:
            return JsonResponse({'success': False, 'error': 'Deduction must be between 0 and 100'}, status=400)

        if not email_id:
            return JsonResponse({'success': False, 'error': 'email_id is required'}, status=400)

        email_obj = get_object_or_404(SentEmail, id=email_id)

        email_obj.manual_deduction = deduction
        email_obj.override_reason  = reason
        email_obj.scored_by        = user_profile
        email_obj.scored_at        = timezone.now()
        email_obj.save(update_fields=[
            'manual_deduction', 'override_reason', 'scored_by', 'scored_at',
        ])

        log_activity(
            user=user_profile,
            activity_type='score_updated',
            details=(
                f'Deduction {deduction} on email #{email_obj.id} '
                f'({email_obj.report_type}) — final score {email_obj.final_score}%'
            ),
            request=request,
        )

        return JsonResponse({
            'success':      True,
            'message':      f'Deduction set to {deduction}',
            'email_id':     email_obj.id,
            'deduction':    email_obj.manual_deduction,
            'final_score':  email_obj.final_score,
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


def _build_exception_records(upload, headers, rows_data):
    """
    Parse Excel rows into typed ExceptionRecord rows.

    `upload` is an ExceptionUpload instance (the container).
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
        target_date, target_date_raw = _parse_date_safe(extracted.get('target_closure_date'))
        cost_saved, cost_saved_raw = _parse_decimal_safe(extracted.get('income_cost_saved'))
        status_val, status_raw = _parse_status_safe(extracted.get('status'))

        records_to_create.append(ExceptionRecord(
            upload=upload,                                        # <-- CHANGED
            serial_number=serial_number,
            branch_unit=str(extracted.get('branch_unit', '') or '').strip()[:200],
            exception=str(extracted.get('exception', '') or '').strip(),
            date_noted=date_noted,
            date_noted_raw=date_noted_raw[:50],
            target_closure_date=target_date,
            target_closure_date_raw=target_date_raw[:50],
            category=str(extracted.get('category', '') or '').strip()[:200],
            responsible_officer=str(extracted.get('responsible_officer', '') or '').strip()[:200],
            supervisor=str(extracted.get('supervisor', '') or '').strip()[:200],
            auditee_response=str(extracted.get('auditee_response', '') or '').strip(),
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


# ==================== API - EXCEL ROW (SINGLE) EDIT / DELETE ====================

@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_excel_row(request, row_id):
    """Delete a SINGLE ExceptionRecord row."""
    try:
        record = get_object_or_404(ExceptionRecord, id=row_id)
        upload = record.upload

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if upload.uploaded_by != user_profile and user_profile.role != 'admin':
                return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        record.delete()

        remaining = upload.exception_records.count()
        upload.row_count = remaining
        upload.save(update_fields=['row_count'])

        log_activity(
            user=user_profile,
            activity_type='draft_deleted',
            details=f'Deleted 1 exception row from upload #{upload.id} ({upload.report_type})',
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
    """Edit a SINGLE ExceptionRecord row."""
    try:
        record = get_object_or_404(ExceptionRecord, id=row_id)
        upload = record.upload

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
            if upload.uploaded_by != user_profile and user_profile.role != 'admin':
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
            details=f'Edited 1 exception row (#{record.id}) in upload #{upload.id}',
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
    Live JSON endpoint for the Supervisor Dashboard.

    Top 5 performers are ranked by the SAME scorecard used on the
    Team Performance page (SentEmail-derived + AdHoc adjustments +
    deadline penalty for missed/late submissions).

    Also returns the summary numbers shown on the KPI cards so the
    dashboard can refresh them in place.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role not in ('supervisor', 'admin'):
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)
    except UserProfile.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

    today_dt   = timezone.now()
    today_date = today_dt.date()

    # ── Summary counters ────────────────────────────────────────
    exceptions_qs = ExceptionRecord.objects.filter(upload__isnull=False)
    total_exceptions = exceptions_qs.count()
    today_exceptions = exceptions_qs.filter(created_at__date=today_date).count()

    submitted_qs = (
        Report.objects
        .filter(status='submitted')
        .exclude(report_type=TRIAL_BALANCE_REPORT_TYPE)
    )
    submitted_reports_count = submitted_qs.count()

    # ── Score every active member via the shared scorecard ─────
    team_members = (
        UserProfile.objects
        .filter(role='member', status='active')
        .order_by('full_name')
    )

    performers = []
    sum_of_percentages = 0
    total_members = 0

    for member in team_members:
        card = _build_member_scorecard(member, today=today_dt)

        # Number of emails this member has actually sent — this is
        # what the dashboard's "Submitted" column has always shown.
        email_count = card['email_count']

        performers.append({
            'id':          member.id,
            'full_name':   member.full_name or member.email,
            'submitted':   email_count,
            'percentage':  card['percentage'],
            'net_score':   card['net_score'],
            'debits':      card['debits'],
            'credits':     card['credits'],
            'status':      card['status'],          # 'success' | 'warning' | 'danger'
            'status_text': card['status_text'],
            # Status icons used by the small dashboard pill
            'status_icon': (
                '🌟' if card['percentage'] >= 70
                else '📈' if card['percentage'] > 0
                else '⚠️'
            ),
        })

        sum_of_percentages += card['percentage']
        total_members += 1

    # Sort by the same key the team page uses so the two views never
    # disagree about who is top performer.
    performers.sort(
        key=lambda x: (x['percentage'], x['net_score'], x['submitted']),
        reverse=True,
    )
    top_performers = performers[:5]

    completion_rate = (
        int(round(sum_of_percentages / total_members))
        if total_members else 0
    )

    return JsonResponse({
        'success': True,
        'top_performers': top_performers,
        'summary': {
            'total_exceptions':        total_exceptions,
            'today_exceptions':        today_exceptions,
            'submitted_reports_count': submitted_reports_count,
            'completion_rate':         completion_rate,
            'team_size':               total_members,
        },
    })

# ═══ NEW ═══ Avatar upload / delete endpoints

@login_required
@require_http_methods(["POST"])
def api_upload_avatar(request, user_id):
    """
    Upload or replace a user's profile picture.

    Multipart/form-data with an 'avatar' file field.
    Admin-only. Max 2 MB. Image mime type required.
    """
    try:
        # ---- Admin role guard ----
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        target = get_object_or_404(UserProfile, id=user_id)

        uploaded = request.FILES.get('avatar')
        if not uploaded:
            return JsonResponse({'success': False, 'error': 'No file uploaded'}, status=400)

        if uploaded.size > 2 * 1024 * 1024:
            return JsonResponse({'success': False, 'error': 'Image must be smaller than 2 MB'}, status=400)

        content_type = uploaded.content_type or ''
        if not content_type.startswith('image/'):
            return JsonResponse({'success': False, 'error': 'File must be an image'}, status=400)

        # Delete old file first
        if target.avatar:
            try:
                target.avatar.delete(save=False)
            except Exception:
                pass

        target.avatar = uploaded
        target.save(update_fields=['avatar', 'updated_at'])

        log_activity(
            user=caller,
            activity_type='user_updated',
            details=f'Updated avatar for {target.email}',
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': 'Avatar updated successfully',
            'avatar_url': target.avatar_url,
        })

    except Exception:
        logger.exception("Error in api_upload_avatar")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)


@login_required
@require_http_methods(["DELETE"])
def api_delete_avatar(request, user_id):
    """Remove a user's uploaded avatar. Falls back to Gravatar."""
    try:
        # ---- Admin role guard ----
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        target = get_object_or_404(UserProfile, id=user_id)

        if target.avatar:
            try:
                target.avatar.delete(save=False)
            except Exception:
                pass
            target.avatar = None
            target.save(update_fields=['avatar', 'updated_at'])

        log_activity(
            user=caller,
            activity_type='user_updated',
            details=f'Removed avatar for {target.email}',
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': 'Avatar removed',
            'avatar_url': target.avatar_url,
        })

    except Exception:
        logger.exception("Error in api_delete_avatar")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)

# ==================== ADMIN — USER LIST ====================

@login_required
def admin_user_list(request):
    """
    User List page.

    Renders every UserProfile with:
      - avatar (uploaded or Gravatar fallback)
      - full name, email, position, role, status
      - activate / deactivate toggle
      - edit (modal form → api_edit_user)

    Only admins may access this page.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    users = (
        UserProfile.objects
        .prefetch_related('branches', 'departments')
        .order_by('full_name')
    )

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
    return render(request, 'control_dashboard/userlist.html', context)

# ==================== ADMIN — BRANCH / DEPARTMENT MANAGEMENT ====================

@login_required
def admin_units(request):
    """
    Branch / Department management page.

    Renders a single page with:
      - Add form (name, type, assignment scope, specific users)
      - Table of every saved Branch and Department
      - Edit / Delete actions on each row

    Only admins may access this page.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
        if user_profile.role != 'admin':
            messages.error(request, 'You do not have permission to access this page.')
            return redirect_dashboard(request.user)
    except UserProfile.DoesNotExist:
        return redirect_dashboard(request.user)

    branches = Branch.objects.all().order_by('name')
    departments = Department.objects.all().order_by('name')

    # Users eligible for "specific" assignment — every active user
    assignable_users = (
        UserProfile.objects
        .filter(status='active')
        .order_by('full_name')
    )

    context = {
        'user_profile': user_profile,
        'branches': branches,
        'departments': departments,
        'assignable_users': assignable_users,
    }
    return render(request, 'control_dashboard/branch_department.html', context)


@csrf_exempt
@require_http_methods(["POST"])
def api_create_unit(request):
    """
    Create a Branch or Department.

    Body:
    {
        "unit_type":   "branch" | "department",
        "name":        "Accra Main",
        "code":        "ACC" (optional, branch only),
        "assignment":  "cc" | "hc" | "specific",
        "user_ids":    [1, 2, 3]   (only when assignment == "specific"),
        "description": "..."        (optional)
    }
    """
    try:
        # ---- Admin role guard ----
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        data = json.loads(request.body)

        unit_type   = (data.get('unit_type') or '').strip().lower()
        name        = (data.get('name') or '').strip()
        code        = (data.get('code') or '').strip()
        assignment  = (data.get('assignment') or 'cc').strip().lower()
        user_ids    = data.get('user_ids') or []
        description = (data.get('description') or '').strip()

        # ---- Validation ----
        if unit_type not in ('branch', 'department'):
            return JsonResponse({'success': False, 'error': 'Unit type must be "branch" or "department".'}, status=400)

        if not name:
            return JsonResponse({'success': False, 'error': 'Name is required.'}, status=400)

        if assignment not in ('cc', 'hc', 'specific'):
            return JsonResponse({'success': False, 'error': 'Invalid assignment scope.'}, status=400)

        if assignment == 'specific' and not user_ids:
            return JsonResponse({'success': False, 'error': 'Select at least one user for specific assignment.'}, status=400)

        # ---- Duplicate check ----
        Model = Branch if unit_type == 'branch' else Department
        if Model.objects.filter(name__iexact=name).exists():
            return JsonResponse({
                'success': False,
                'error': f'A {unit_type} named "{name}" already exists.'
            }, status=400)

        # ---- Create ----
        create_kwargs = {
            'name': name,
            'description': description,
            'is_active': True,
        }
        if unit_type == 'branch' and code:
            create_kwargs['code'] = code

        unit = Model.objects.create(**create_kwargs)

        # ---- Attach users for "specific" assignment ----
        if assignment == 'specific':
            users_qs = UserProfile.objects.filter(id__in=user_ids, status='active')
            if hasattr(unit, 'assigned_users'):
                unit.assigned_users.set(users_qs)

        log_activity(
            user=caller,
            activity_type='unit_created',
            details=f'Created {unit_type} "{name}" (assignment: {assignment})',
            request=request,
        )

        return JsonResponse({
            'success': True,
            'message': f'{unit_type.capitalize()} created successfully.',
            'unit_id': unit.id,
            'unit_type': unit_type,
        }, status=201)

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data.'}, status=400)
    except Exception:
        logger.exception("Error in api_create_unit")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)


@csrf_exempt
@require_http_methods(["POST"])
def api_edit_unit(request, unit_type, unit_id):
    """
    Edit a Branch or Department.

    URL: /adminboard/api/units/<branch|department>/<id>/edit/
    """
    try:
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        if unit_type not in ('branch', 'department'):
            return JsonResponse({'success': False, 'error': 'Invalid unit type.'}, status=400)

        Model = Branch if unit_type == 'branch' else Department
        unit  = get_object_or_404(Model, id=unit_id)

        data = json.loads(request.body)

        changes = []

        if 'name' in data:
            new_name = (data['name'] or '').strip()
            if not new_name:
                return JsonResponse({'success': False, 'error': 'Name cannot be empty.'}, status=400)
            if new_name.lower() != unit.name.lower():
                if Model.objects.filter(name__iexact=new_name).exclude(id=unit.id).exists():
                    return JsonResponse({'success': False, 'error': f'Another {unit_type} already uses that name.'}, status=400)
                changes.append(f'Name: "{unit.name}" → "{new_name}"')
                unit.name = new_name

        if 'description' in data:
            unit.description = (data['description'] or '').strip()

        if unit_type == 'branch' and 'code' in data:
            unit.code = (data['code'] or '').strip()[:20]

        if 'is_active' in data:
            unit.is_active = bool(data['is_active'])

        unit.save()

        # ---- Update specific-user assignment ----
        assignment = (data.get('assignment') or '').strip().lower()
        if assignment == 'specific' and 'user_ids' in data:
            users_qs = UserProfile.objects.filter(id__in=data['user_ids'], status='active')
            if hasattr(unit, 'assigned_users'):
                unit.assigned_users.set(users_qs)
                changes.append(f'Users assigned: {users_qs.count()}')

        if changes:
            log_activity(
                user=caller,
                activity_type='unit_updated',
                details=f'{unit_type.capitalize()} "{unit.name}" updated: ' + '; '.join(changes),
                request=request,
            )

        return JsonResponse({'success': True, 'message': f'{unit_type.capitalize()} updated successfully.'})

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data.'}, status=400)
    except Exception:
        logger.exception("Error in api_edit_unit")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)


@csrf_exempt
@require_http_methods(["DELETE"])
def api_delete_unit(request, unit_type, unit_id):
    """Delete a Branch or Department."""
    try:
        try:
            caller = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found'}, status=404)

        if caller.role != 'admin':
            return JsonResponse({'success': False, 'error': 'Permission denied'}, status=403)

        if unit_type not in ('branch', 'department'):
            return JsonResponse({'success': False, 'error': 'Invalid unit type.'}, status=400)

        Model = Branch if unit_type == 'branch' else Department
        unit  = get_object_or_404(Model, id=unit_id)

        unit_name = unit.name
        unit.delete()

        log_activity(
            user=caller,
            activity_type='unit_deleted',
            details=f'Deleted {unit_type} "{unit_name}"',
            request=request,
        )

        return JsonResponse({'success': True, 'message': f'{unit_type.capitalize()} deleted successfully.'})

    except Exception:
        logger.exception("Error in api_delete_unit")
        return JsonResponse({'success': False, 'error': 'An internal error occurred.'}, status=500)










def _get_email_report_type_options(user_profile, is_privileged=False):
    """
    Report type options for the Email compose page.

    Members  → only report types they're assigned to, PLUS the
               consolidated report types they can access.
    Admins / supervisors → every non-TB, non-uploaded report type.
    """
    base_qs = Report.objects.exclude(
        report_type=TRIAL_BALANCE_REPORT_TYPE
    ).exclude(
        status=UPLOADED_STATUS
    )

    if is_privileged:
        qs = base_qs
    else:
        qs = base_qs.filter(
            Q(assigned_to=user_profile) |
            Q(is_assigned_to_all=True) |
            Q(created_by=user_profile)
        )

    types = sorted(
        {rt for rt in qs.values_list('report_type', flat=True) if rt}
    )

    # Always allow the "consolidated reports" that the user can see,
    # even if they weren't directly assigned to the underlying reports.
    consolidated_types = sorted(
        {rt for rt in Report.objects.filter(
            Q(assigned_to=user_profile) |
            Q(is_assigned_to_all=True) |
            Q(created_by=user_profile)
        ).exclude(
            report_type=TRIAL_BALANCE_REPORT_TYPE
        ).filter(
            report_type__icontains='consolidated'
        ).values_list('report_type', flat=True) if rt}
    )

    combined = sorted(set(types) | set(consolidated_types))

    # Return as list of dicts so the template can render option labels
    return [{'value': rt, 'label': rt} for rt in combined]


@login_required
def email_page(request):
    """
    Email compose page.

    To, CC, Report Type (DB), Header, Body, Send.
    """
    try:
        user_profile = UserProfile.objects.get(email=request.user.email)
    except UserProfile.DoesNotExist:
        return redirect('control_dashboard:member_dashboard')

    is_privileged = user_profile.role in ('admin', 'supervisor')

    # Recipients: every active user, excluding the sender themselves
    # (they can still manually type their own address if they want).
    recipients = (
        UserProfile.objects
        .filter(status='active')
        .exclude(id=user_profile.id)
        .order_by('full_name')
    )

    report_type_options = _get_email_report_type_options(
        user_profile, is_privileged=is_privileged
    )

    # If a ?report_type=<X> is passed (e.g. from the report page),
    # pre-select it in the dropdown.
    preselect_report_type = (request.GET.get('report_type') or '').strip()

    context = {
        'user_profile': user_profile,
        'today': timezone.now(),
        'recipients': recipients,
        'report_type_options': report_type_options,
        'preselect_report_type': preselect_report_type,
        'is_privileged': is_privileged,
    }
    return render(request, 'control_dashboard/email.html', context)


# ── Canonical exception headers (used in both email HTML and Excel) ──
EXCEPTION_CANONICAL_HEADERS = [
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


def _build_consolidated_email_html(report_type, header_line, date_from, date_to, rows):
    """
    Build the HTML body — ONE table, one header, all rows flowing through.
    Uses inline styles so it survives Gmail/Outlook/Apple Mail.
    """
    # Header row
    th_style = (
        'background:#1a3a6b;color:#ffffff;padding:8px 6px;'
        'border:1px solid #1a3a6b;text-align:left;font-weight:700;'
        'font-size:10px;letter-spacing:0.4px;text-transform:uppercase;'
        'vertical-align:middle;'
    )
    td_style = (
        'padding:6px;border:1px solid #d0d7e2;font-size:11px;'
        'color:#1a2332;vertical-align:top;'
    )

    header_cells = ''.join(
        f'<th style="{th_style}">{html_escape(h)}</th>'
        for h in EXCEPTION_CANONICAL_HEADERS
    )

    body_rows = []
    for idx, r in enumerate(rows, start=1):
        cells = [
            idx,                          # renumbered S/N
            r['branch_unit'],
            r['exception'],
            r['date_noted'],
            r['target_closure_date'],
            r['category'],
            r['responsible_officer'],
            r['supervisor'],
            r['auditee_response'],
            r['remarks'],
            r['status'],
            r['income_cost_saved'],
        ]
        tds = ''.join(
            f'<td style="{td_style}">{html_escape(str(c if c not in (None, "") else ""))}</td>'
            for c in cells
        )
        body_rows.append(f'<tr>{tds}</tr>')

    # Range label
    from_label = date_from.strftime('%b %d, %Y') if date_from else 'the beginning'
    to_label   = date_to.strftime('%b %d, %Y')   if date_to   else 'today'

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f5f7fb;font-family:Arial,Helvetica,sans-serif;">
  <div style="max-width:1200px;margin:0 auto;padding:24px;">

    <div style="background:#ffffff;border:1px solid #e2e6ef;border-radius:10px;padding:24px;">

      <h2 style="margin:0 0 6px 0;font-size:16px;color:#0f1b33;">
        {html_escape(report_type)}
      </h2>
      <p style="margin:0 0 18px 0;font-size:12px;color:#6b7280;">
        Range: <strong>{from_label}</strong> → <strong>{to_label}</strong>
        &nbsp;·&nbsp; {len(rows)} exception{'s' if len(rows) != 1 else ''}
      </p>

      <div style="overflow-x:auto;-webkit-overflow-scrolling:touch;">
        <table cellpadding="0" cellspacing="0" border="0"
               style="border-collapse:collapse;width:100%;min-width:1100px;">
          <thead><tr>{header_cells}</tr></thead>
          <tbody>{''.join(body_rows)}</tbody>
        </table>
      </div>

    </div>

    <p style="margin:18px 0 0 0;font-size:11px;color:#9aa3b2;text-align:center;">
      Sent automatically from the Exception Reporting System.
    </p>

  </div>
</body>
</html>"""
    return html


def _build_consolidated_email_excel(report_type, date_from, date_to, rows):
    """
    Build the Excel workbook — ONE sheet, ONE header, all rows flowing through.
    Matches the uploaded source format (Calibri 11, dark-blue header,
    white bold text, wrapped, thin borders, auto column widths).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Exceptions'

    # ---- Header row ----
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1A3A6B', end_color='1A3A6B', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(border_style='thin', color='9AA3B2')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(EXCEPTION_CANONICAL_HEADERS, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = header_align
        c.border = border

    # ---- Data rows ----
    body_font = Font(name='Calibri', size=11)
    body_align_top = Alignment(vertical='top', wrap_text=True)

    for idx, r in enumerate(rows, start=1):
        values = [
            idx,
            r['branch_unit'],
            r['exception'],
            r['date_noted'],
            r['target_closure_date'],
            r['category'],
            r['responsible_officer'],
            r['supervisor'],
            r['auditee_response'],
            r['remarks'],
            r['status'],
            r['income_cost_saved'],
        ]
        for col, v in enumerate(values, start=1):
            c = ws.cell(row=idx + 1, column=col, value=v if v not in (None, '') else '')
            c.font = body_font
            c.alignment = body_align_top
            c.border = border

    # ---- Column widths (approximations of the source sheet) ----
    widths = [6, 16, 55, 12, 14, 16, 16, 14, 30, 20, 12, 16]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(EXCEPTION_CANONICAL_HEADERS))}{ws.max_row}'

    # Save to bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


@csrf_exempt
@require_http_methods(["POST"])
def api_send_email(request):
    try:
        if not request.body:
            return JsonResponse({'success': False, 'error': 'Empty request body.'}, status=400)

        data = json.loads(request.body)

        to_list       = data.get('to') or []
        cc_list       = data.get('cc') or []
        report_type   = (data.get('report_type') or '').strip()
        date_from_str = (data.get('date_from') or '').strip()
        date_to_str   = (data.get('date_to')   or '').strip()
        header        = (data.get('header') or '').strip()
        body          = (data.get('body') or '').strip()

        # ---- Validation ----
        if not isinstance(to_list, list) or not to_list:
            return JsonResponse({'success': False, 'error': 'At least one "To" recipient is required.'}, status=400)
        if not report_type:
            return JsonResponse({'success': False, 'error': 'Report type is required.'}, status=400)
        if not header:
            return JsonResponse({'success': False, 'error': 'Email header is required.'}, status=400)
        if not body:
            return JsonResponse({'success': False, 'error': 'Email body is required.'}, status=400)

        # ---- Parse optional date range ----
        date_from = None
        date_to   = None
        if date_from_str:
            try:
                date_from = datetime.strptime(date_from_str, '%Y-%m-%d').date()
            except ValueError:
                return JsonResponse({'success': False, 'error': 'Invalid "From" date.'}, status=400)
        if date_to_str:
            try:
                date_to = datetime.strptime(date_to_str, '%Y-%m-%d').date()
            except ValueError:
                return JsonResponse({'success': False, 'error': 'Invalid "To" date.'}, status=400)
        if date_from and date_to and date_from > date_to:
            return JsonResponse({'success': False, 'error': '"From" date must not be after "To" date.'}, status=400)

        # ---- Normalize recipients ----
        def _clean(lst):
            out = []
            for addr in lst:
                a = str(addr or '').strip()
                if a and a not in out:
                    out.append(a)
            return out

        to_list = _clean(to_list)
        cc_list = _clean(cc_list)

        email_re = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')
        bad = [a for a in (to_list + cc_list) if not email_re.match(a)]
        if bad:
            return JsonResponse({
                'success': False,
                'error': f'Invalid email address(es): {", ".join(bad)}',
            }, status=400)

        # ---- Caller ----
        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

        # ================================================================
        # Collect exception rows.
        #
        # If the sender ticked specific rows, only those are sent.
        # If nothing was ticked, we fall back to the full range —
        # every exception this user uploaded for the chosen report type.
        # ================================================================
        selected_ids = data.get('selected_ids') or []

        if selected_ids:
            # Sender hand-picked rows — enforce they belong to this user
            records_qs = (
                ExceptionRecord.objects
                .filter(
                    id__in=selected_ids,
                    upload__uploaded_by=user_profile,
                )
                .select_related('upload')
                .order_by('upload__created_at', 'source_row_index', 'serial_number')
            )
        else:
            # Nothing selected → send everything in range
            uploads_qs = ExceptionUpload.objects.filter(
                uploaded_by=user_profile,
                report_type=report_type,
            )
            if date_from:
                uploads_qs = uploads_qs.filter(created_at__date__gte=date_from)
            if date_to:
                uploads_qs = uploads_qs.filter(created_at__date__lte=date_to)
            uploads_qs = uploads_qs.order_by('created_at')

            records_qs = (
                ExceptionRecord.objects
                .filter(upload__in=uploads_qs)
                .select_related('upload')
                .order_by('upload__created_at', 'source_row_index', 'serial_number')
            )

        def _fmt_date(d, raw):
            if d:
                return d.strftime('%d-%b-%y')
            return raw or ''

        def _fmt_money(v):
            try:
                return f'{float(v):,.2f}' if v is not None else ''
            except (TypeError, ValueError):
                return ''

        rows = []
        for rec in records_qs:
            rows.append({
                'branch_unit':         rec.branch_unit or '',
                'exception':           rec.exception or '',
                'date_noted':          _fmt_date(rec.date_noted, rec.date_noted_raw),
                'target_closure_date': _fmt_date(rec.target_closure_date, rec.target_closure_date_raw),
                'category':            rec.category or '',
                'responsible_officer': rec.responsible_officer or '',
                'supervisor':          rec.supervisor or '',
                'auditee_response':    rec.auditee_response or '',
                'remarks':             rec.remarks or '',
                'status':              rec.get_status_display() or rec.status_raw or '',
                'income_cost_saved':   _fmt_money(rec.income_cost_saved),
            })

        # ================================================================
        # Build HTML + Excel
        # ================================================================
        html_body = _build_consolidated_email_html(
            report_type, header, date_from, date_to, rows,
        ) if rows else None

        excel_bytes = _build_consolidated_email_excel(
            report_type, date_from, date_to, rows,
        ) if rows else None

        # ---- Compose message ----
        full_subject = f'[{report_type}] {header}'
        from_email   = getattr(settings, 'DEFAULT_FROM_EMAIL', None) or user_profile.email

        # Plain-text fallback for clients that don't render HTML
        from_label = date_from.strftime('%b %d, %Y') if date_from else 'the beginning'
        to_label   = date_to.strftime('%b %d, %Y')   if date_to   else 'today'
        plain_body = (
            f'{body}\n\n'
            f'---\n'
            f'Report type: {report_type}\n'
            f'Range: {from_label} → {to_label}\n'
            f'Exceptions included: {len(rows)}\n'
            f'---\n\n'
            f'The full table is in the attached Excel workbook.'
        )

        msg = EmailMultiAlternatives(
            subject=full_subject,
            body=plain_body,
            from_email=from_email,
            to=to_list,
            cc=cc_list,
        )

        if html_body:
            msg.attach_alternative(html_body, 'text/html')

        if excel_bytes:
            safe_rt = re.sub(r'[^A-Za-z0-9_-]', '_', report_type)[:40]
            file_from = date_from.strftime('%Y%m%d') if date_from else 'start'
            file_to   = date_to.strftime('%Y%m%d')   if date_to   else 'today'
            filename  = f'{safe_rt}_{file_from}_{file_to}.xlsx'
            msg.attach(
                filename,
                excel_bytes,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

        # ---- Send ----
        send_error = None
        try:
            msg.send(fail_silently=False)
            delivery_status = 'sent'
        except Exception as exc:
            logger.exception("SMTP send failed")
            send_error = str(exc)
            delivery_status = 'failed'

        # ---- Persist audit row (records sender, time, recipients) ----
        try:
            SentEmail.objects.create(
                sender=user_profile,
                report_type=report_type,
                subject=full_subject,
                body=body,
                to_addresses=', '.join(to_list),
                cc_addresses=', '.join(cc_list),
                status=delivery_status,
                error_message=send_error or '',
            )
        except Exception:
            logger.exception("Failed to persist SentEmail audit row")

        log_activity(
            user=user_profile,
            activity_type='email_sent',
            details=(
                f'Sent email "{full_subject}" to {", ".join(to_list)}'
                + (f' (cc: {", ".join(cc_list)})' if cc_list else '')
                + (f' — range {date_from_str or "…"} → {date_to_str or "…"}'
                   f' ({len(rows)} exceptions)' if (date_from or date_to) else '')
                + (f' — SMTP error: {send_error}' if send_error else '')
            ),
            request=request,
        )

        if send_error:
            return JsonResponse({
                'success': False,
                'error': f'Email saved but could not be sent: {send_error}',
            }, status=500)

        return JsonResponse({
            'success': True,
            'message': (
                f'Email sent to {len(to_list)} recipient(s)'
                + (f' and {len(cc_list)} cc' if cc_list else '')
                + (f' — {len(rows)} exception(s) attached.' if rows else '.')
            ),
        })

    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON data.'}, status=400)
    except Exception as e:
        logger.exception("Error in api_send_email")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@csrf_exempt
@require_http_methods(["GET"])
def api_email_preview_exceptions(request):
    """
    Return the current user's own exception rows for a given
    report type and upload-date range, so the compose page can
    offer a picker and a live preview.
    """
    try:
        report_type = (request.GET.get('report_type') or '').strip()
        date_from_s = (request.GET.get('date_from') or '').strip()
        date_to_s   = (request.GET.get('date_to')   or '').strip()

        if not report_type:
            return JsonResponse({'success': True, 'rows': []})

        try:
            user_profile = UserProfile.objects.get(email=request.user.email)
        except UserProfile.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'User not found.'}, status=404)

        date_from = None
        date_to   = None
        if date_from_s:
            try: date_from = datetime.strptime(date_from_s, '%Y-%m-%d').date()
            except ValueError: pass
        if date_to_s:
            try: date_to = datetime.strptime(date_to_s, '%Y-%m-%d').date()
            except ValueError: pass

        uploads_qs = ExceptionUpload.objects.filter(
            uploaded_by=user_profile,
            report_type=report_type,
        )
        if date_from:
            uploads_qs = uploads_qs.filter(created_at__date__gte=date_from)
        if date_to:
            uploads_qs = uploads_qs.filter(created_at__date__lte=date_to)

        uploads_qs = uploads_qs.order_by('created_at')

        records_qs = (
            ExceptionRecord.objects
            .filter(upload__in=uploads_qs)
            .select_related('upload')
            .order_by('upload__created_at', 'source_row_index', 'serial_number')[:500]
        )

        def _fmt_date(d, raw):
            if d: return d.strftime('%d-%b-%y')
            return raw or ''

        def _fmt_money(v):
            try:    return f'{float(v):,.2f}' if v is not None else ''
            except (TypeError, ValueError): return ''

        rows = []
        for rec in records_qs:
            rows.append({
                'id':                  rec.id,
                'serial_number':       rec.serial_number,
                'branch_unit':         rec.branch_unit or '',
                'exception':           rec.exception or '',
                'date_noted':          _fmt_date(rec.date_noted, rec.date_noted_raw),
                'target_closure_date': _fmt_date(rec.target_closure_date, rec.target_closure_date_raw),
                'category':            rec.category or '',
                'responsible_officer': rec.responsible_officer or '',
                'supervisor':          rec.supervisor or '',
                'auditee_response':    rec.auditee_response or '',
                'remarks':             rec.remarks or '',
                'status':              rec.get_status_display() or rec.status_raw or '',
                'income_cost_saved':   _fmt_money(rec.income_cost_saved),
            })

        return JsonResponse({'success': True, 'count': len(rows), 'rows': rows})

    except Exception as e:
        logger.exception("Error in api_email_preview_exceptions")
        return JsonResponse({'success': False, 'error': str(e)}, status=500)