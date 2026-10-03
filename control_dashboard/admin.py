# control_dashboard/admin.py

from django.contrib import admin
from django.contrib import messages
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User
from django.utils.html import format_html
from django.utils import timezone

from import_export import resources, fields
from import_export.admin import ImportExportModelAdmin
from import_export.formats.base_formats import CSV

import re

from .models import (
    # Core users
    UserProfile,
    Branch,
    Department,

    # Reports
    Report,
    ReportDataField,
    ReportSubmission,
    ReportSubmissionField,
    ReportSchedule,

    # Checklists
    Checklist,
    ChecklistTask,
    ChecklistLog,
    ChecklistChangeRequest,

    # Uploads
    ExceptionUpload,
    ExceptionRecord,
    TrialBalanceUpload,
    TrialBalanceEntry,

    # Scoring & audit
    SentEmail,
    AdHocDeduction,
    ActivityLog,
)


# ==================== CHECKLIST RESOURCE ====================

class ChecklistResource(resources.ModelResource):
    """Resource for importing Checklist data from CSV"""

    name = fields.Field(attribute='name', column_name='ACTIVITY')
    description = fields.Field(attribute='description', column_name='TASK / DESCRIPTION')
    frequency = fields.Field(attribute='frequency', column_name='FREQUENCY')
    assignment_target = fields.Field(attribute='assignment_target', column_name='ASSIGNMENT TARGET')

    class Meta:
        model = Checklist
        fields = ('name', 'description', 'frequency', 'assignment_target')
        import_id_fields = ('name',)
        skip_unchanged = True
        report_skipped = False

    def before_import_row(self, row, **kwargs):
        """Clean and prepare data before import"""

        if 'ACTIVITY' in row:
            row['ACTIVITY'] = str(row['ACTIVITY']).strip()
            row['ACTIVITY'] = ' '.join(row['ACTIVITY'].split())

        if 'TASK / DESCRIPTION' in row:
            desc = str(row['TASK / DESCRIPTION'])
            desc = desc.replace('•', '').strip()
            desc = desc.replace('\n', ' ').strip()
            desc = desc.replace('\r', ' ').strip()
            desc = desc.replace('"', '').strip()
            desc = ' '.join(desc.split())
            row['TASK / DESCRIPTION'] = desc

        if 'FREQUENCY' in row:
            freq = str(row['FREQUENCY']).strip().lower()
            freq_map = {
                'daily': 'daily',
                'weekly': 'weekly',
                'monthly': 'monthly',
                'quarterly': 'quarterly',
                'bi-annual': 'bi-annual',
                'bi annual': 'bi-annual',
                'biannual': 'bi-annual',
                'annual': 'annual',
                'yearly': 'annual',
            }
            row['FREQUENCY'] = freq_map.get(freq, 'weekly')

        if 'ASSIGNMENT TARGET' in row:
            target = str(row['ASSIGNMENT TARGET']).strip().lower()
            target_map = {
                'cluster control': 'cc',
                'cc': 'cc',
                'head office control': 'hc',
                'hc': 'hc',
                'all users': 'all',
                'all': 'all',
                'specific': 'specific',
            }
            row['ASSIGNMENT TARGET'] = target_map.get(target, 'all')

        return row

    def after_import_row(self, row, row_result, **kwargs):
        """After each row is imported, create tasks and assign users"""
        if row_result.errors:
            return

        try:
            checklist = Checklist.objects.get(name=row['ACTIVITY'])
        except Checklist.DoesNotExist:
            print(f"❌ Checklist not found: {row.get('ACTIVITY')}")
            return
        except Exception as e:
            print(f"❌ Error: {e}")
            return

        try:
            # 1. Set assignment target
            assignment_target = row.get('ASSIGNMENT TARGET', 'all')
            checklist.assignment_target = assignment_target
            checklist.is_active = True
            checklist.save()

            # 2. Create tasks from the description
            description = row.get('TASK / DESCRIPTION', '')
            if description:
                tasks = self.split_into_tasks(description)

                ChecklistTask.objects.filter(checklist=checklist).delete()

                for order, task_text in enumerate(tasks, start=1):
                    ChecklistTask.objects.create(
                        checklist=checklist,
                        description=task_text.strip()[:500],
                        order=order,
                        is_completed=False,
                    )
                print(f"✅ Created {len(tasks)} tasks for: {checklist.name}")

            # 3. Assign users based on target
            #    NOTE: 'cc' and 'hc' are POSITION values, not ROLE values.
            if assignment_target == 'cc':
                cc_users = UserProfile.objects.filter(position='cc')
                if cc_users.exists():
                    checklist.assigned_users.set(cc_users)
                    print(f"✅ Assigned {cc_users.count()} CC users to: {checklist.name}")
                else:
                    print(f"⚠️ No CC users found for: {checklist.name}")

            elif assignment_target == 'hc':
                hc_users = UserProfile.objects.filter(position='hc')
                if hc_users.exists():
                    checklist.assigned_users.set(hc_users)
                    print(f"✅ Assigned {hc_users.count()} HC users to: {checklist.name}")
                else:
                    print(f"⚠️ No HC users found for: {checklist.name}")

            elif assignment_target == 'all':
                all_users = UserProfile.objects.filter(status='active')
                checklist.assigned_users.set(all_users)
                print(f"✅ Assigned {all_users.count()} users to: {checklist.name}")

        except Exception as e:
            print(f"❌ Error while finalizing {row.get('ACTIVITY')}: {e}")

    def split_into_tasks(self, text):
        """Split description into individual tasks"""
        if '•' in text:
            return [t.strip() for t in text.split('•') if t.strip()]

        if re.search(r'\d+\.', text):
            tasks = re.split(r'\d+\.\s*', text)
            return [t.strip() for t in tasks if t.strip()]

        if '.' in text:
            return [t.strip() + '.' for t in text.split('.') if t.strip()]

        return [text]


# ==================== INLINES ====================

class ChecklistTaskInline(admin.TabularInline):
    model = ChecklistTask
    extra = 0
    fields = ['description', 'order', 'is_completed']
    ordering = ['order']
    show_change_link = True


class ReportDataFieldInline(admin.TabularInline):
    model = ReportDataField
    extra = 0
    fields = ['field_name', 'field_value', 'field_type', 'order']
    ordering = ['order']


class ReportScheduleInline(admin.TabularInline):
    model = ReportSchedule
    extra = 0
    fields = ['frequency', 'start_date', 'end_date', 'due_time',
              'last_submitted', 'next_due_date', 'is_active']
    ordering = ['next_due_date']


class ReportSubmissionFieldInline(admin.TabularInline):
    model = ReportSubmissionField
    extra = 0
    fields = ['field_key', 'field_value', 'field_type', 'field_label', 'order']
    ordering = ['order']


class ExceptionRecordInline(admin.TabularInline):
    model = ExceptionRecord
    extra = 0
    fields = ['serial_number', 'branch_unit', 'exception', 'status', 'income_cost_saved']
    ordering = ['source_row_index']
    show_change_link = True
    can_delete = True


class TrialBalanceEntryInline(admin.TabularInline):
    model = TrialBalanceEntry
    extra = 0
    fields = ['row_index', 'branch_code', 'gl_code', 'descr',
              'close_bal_lcy', 'close_bal_fcy', 'gl_status']
    ordering = ['row_index']
    show_change_link = True
    can_delete = True


# ==================== CHECKLIST ADMIN ====================

@admin.register(Checklist)
class ChecklistAdmin(ImportExportModelAdmin):
    resource_class = ChecklistResource

    list_display = (
        'name', 'get_frequency_display', 'get_assignment_display',
        'is_active', 'created_at',
    )
    list_filter = ('frequency', 'assignment_target', 'is_active')
    search_fields = ('name', 'description')
    readonly_fields = ('created_at', 'updated_at')
    filter_horizontal = ('assigned_users', 'assigned_branches', 'assigned_departments')
    fields = (
        'name', 'description', 'frequency',
        'assignment_target', 'assigned_users',
        'assigned_branches', 'assigned_departments',
        'is_active', 'created_by',
        'created_at', 'updated_at',
    )
    inlines = [ChecklistTaskInline]

    def get_import_formats(self):
        return [CSV]

    def process_import(self, request, *args, **kwargs):
        try:
            result = super().process_import(request, *args, **kwargs)
            count = Checklist.objects.count()
            tasks_count = ChecklistTask.objects.count()
            messages.success(
                request,
                f"✅ Import completed! {count} checklists and {tasks_count} tasks created."
            )
            return result
        except Exception as e:
            messages.error(request, f"❌ Import failed: {e}")
            raise


@admin.register(ChecklistTask)
class ChecklistTaskAdmin(admin.ModelAdmin):
    list_display = ('checklist', 'description', 'order', 'is_completed')
    list_filter = ('checklist', 'is_completed')
    search_fields = ('description',)


@admin.register(ChecklistLog)
class ChecklistLogAdmin(admin.ModelAdmin):
    list_display = ('checklist', 'user', 'branch', 'department', 'log_date', 'created_at')
    list_filter = ('checklist', 'user', 'branch', 'department')
    search_fields = ('checklist__name', 'user__full_name')
    date_hierarchy = 'log_date'


# ==================== USER PROFILE INLINE ====================

class UserProfileInline(admin.StackedInline):
    model = UserProfile
    can_delete = False
    verbose_name_plural = 'Profile'
    fk_name = 'user'
    fields = (
        'email', 'full_name', 'position', 'role', 'status',
        'branches', 'departments', 'avatar',
    )
    filter_horizontal = ('branches', 'departments')


class CustomUserAdmin(UserAdmin):
    inlines = (UserProfileInline,)
    list_display = ('username', 'email', 'first_name', 'last_name', 'is_staff')
    search_fields = ('username', 'email', 'first_name', 'last_name')


admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)


# ==================== USER PROFILE ADMIN ====================

@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):
    list_display = ('avatar_preview', 'full_name', 'email', 'username',
                    'position', 'role', 'status')
    list_filter = ('position', 'role', 'status', 'branches', 'departments')
    search_fields = ('full_name', 'email', 'username')
    readonly_fields = ('created_at', 'updated_at', 'avatar_preview_large')
    filter_horizontal = ('branches', 'departments')

    fieldsets = (
        ('Identity', {
            'fields': ('user', 'email', 'username', 'full_name'),
        }),
        ('Roles & Permissions', {
            'fields': ('role', 'position', 'status'),
        }),
        ('Assignments', {
            'fields': ('branches', 'departments'),
        }),
        ('Profile Picture', {
            'fields': ('avatar', 'avatar_preview_large'),
            'description': (
                'Upload a square image (recommended: 400×400 px, under 200 KB). '
                'If left blank, the user will get an auto-generated Gravatar '
                'based on their email.'
            ),
        }),
        ('Metadata', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',),
        }),
    )

    def avatar_preview(self, obj):
        if obj.avatar and getattr(obj.avatar, 'url', None):
            try:
                return format_html(
                    '<img src="{}" style="width:32px;height:32px;'
                    'border-radius:50%;object-fit:cover;'
                    'border:1px solid #ddd;vertical-align:middle;" />',
                    obj.avatar.url
                )
            except ValueError:
                pass

        return format_html(
            '<span style="display:inline-grid;place-items:center;'
            'width:32px;height:32px;border-radius:50%;'
            'background:#D4AF37;color:#001F3D;'
            'font-weight:700;font-size:12px;'
            'vertical-align:middle;">{}</span>',
            obj.initials or '?'
        )
    avatar_preview.short_description = 'Avatar'

    def avatar_preview_large(self, obj):
        if obj.avatar and getattr(obj.avatar, 'url', None):
            try:
                return format_html(
                    '<img src="{}" style="width:120px;height:120px;'
                    'border-radius:50%;object-fit:cover;'
                    'border:2px solid #D4AF37;" />',
                    obj.avatar.url
                )
            except ValueError:
                pass

        return format_html(
            '<div style="display:inline-grid;place-items:center;'
            'width:120px;height:120px;border-radius:50%;'
            'background:#D4AF37;color:#001F3D;'
            'font-weight:700;font-size:36px;letter-spacing:2px;">{}</div>'
            '<p style="margin-top:8px;color:#6B7280;font-size:12px;">'
            'No custom avatar — Gravatar will be used.</p>',
            obj.initials or '?'
        )
    avatar_preview_large.short_description = 'Current avatar'


# ==================== BRANCH / DEPARTMENT ====================

@admin.register(Branch)
class BranchAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'assignment_scope', 'is_active', 'created_at')
    list_filter = ('assignment_scope', 'is_active')
    search_fields = ('name', 'code')
    filter_horizontal = ('assigned_users',)
    readonly_fields = ('created_at', 'updated_at')


@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'assignment_scope', 'is_active', 'created_at')
    list_filter = ('assignment_scope', 'is_active')
    search_fields = ('name', 'code')
    filter_horizontal = ('assigned_users',)
    readonly_fields = ('created_at', 'updated_at')


# ==================== REPORT ADMIN ====================

@admin.register(Report)
class ReportAdmin(admin.ModelAdmin):
    list_display = (
        'report_type', 'frequency', 'status',
        'deadline_date', 'deadline_time',
        'is_assigned_to_all', 'created_by', 'created_at',
    )
    list_filter = ('frequency', 'status', 'is_assigned_to_all')
    search_fields = ('report_type', 'description')
    filter_horizontal = ('assigned_to',)
    readonly_fields = ('created_at', 'updated_at')
    inlines = [ReportDataFieldInline, ReportScheduleInline]
    date_hierarchy = 'created_at'


@admin.register(ReportDataField)
class ReportDataFieldAdmin(admin.ModelAdmin):
    list_display = ('report', 'field_name', 'field_type', 'order')
    list_filter = ('field_type',)
    search_fields = ('report__report_type', 'field_name')


@admin.register(ReportSubmission)
class ReportSubmissionAdmin(admin.ModelAdmin):
    list_display = ('report_type', 'submitted_by', 'status', 'submission_date')
    list_filter = ('status', 'report_type')
    search_fields = ('report_type', 'submitted_by__full_name', 'submitted_by__email')
    readonly_fields = ('submission_date', 'updated_at')
    inlines = [ReportSubmissionFieldInline]
    date_hierarchy = 'submission_date'


@admin.register(ReportSubmissionField)
class ReportSubmissionFieldAdmin(admin.ModelAdmin):
    list_display = ('submission', 'field_key', 'field_type', 'order')
    list_filter = ('field_type',)
    search_fields = ('field_key', 'submission__report_type')


@admin.register(ReportSchedule)
class ReportScheduleAdmin(admin.ModelAdmin):
    list_display = ('report', 'frequency', 'start_date', 'next_due_date', 'is_active')
    list_filter = ('frequency', 'is_active')
    search_fields = ('report__report_type',)


# ==================== EXCEPTION UPLOADS ====================

@admin.register(ExceptionUpload)
class ExceptionUploadAdmin(admin.ModelAdmin):
    list_display = ('report_type', 'uploaded_by', 'file_name', 'row_count', 'created_at')
    list_filter = ('report_type',)
    search_fields = ('report_type', 'file_name', 'uploaded_by__full_name')
    readonly_fields = ('created_at', 'updated_at')
    inlines = [ExceptionRecordInline]
    date_hierarchy = 'created_at'


@admin.register(ExceptionRecord)
class ExceptionRecordAdmin(admin.ModelAdmin):
    list_display = (
        'serial_number', 'branch_unit', 'category',
        'status', 'responsible_officer',
        'target_closure_date', 'income_cost_saved',
        'upload',
    )
    list_filter = ('status', 'category', 'branch_unit')
    search_fields = (
        'exception', 'branch_unit', 'responsible_officer',
        'supervisor', 'remarks',
    )
    readonly_fields = ('created_at', 'updated_at', 'is_overdue', 'days_until_target')
    date_hierarchy = 'date_noted'

    fieldsets = (
        ('Container', {
            'fields': ('upload',),
        }),
        ('Identity', {
            'fields': ('serial_number', 'branch_unit', 'category'),
        }),
        ('Exception details', {
            'fields': ('exception', 'date_noted', 'date_noted_raw',
                       'target_closure_date', 'target_closure_date_raw'),
        }),
        ('People', {
            'fields': ('responsible_officer', 'supervisor'),
        }),
        ('Resolution', {
            'fields': ('auditee_response', 'remarks',
                       'status', 'status_raw',
                       'income_cost_saved', 'income_cost_saved_raw'),
        }),
        ('Meta', {
            'fields': ('source_row_index',
                       'created_at', 'updated_at',
                       'is_overdue', 'days_until_target'),
            'classes': ('collapse',),
        }),
    )


# ==================== TRIAL BALANCE UPLOADS ====================

@admin.register(TrialBalanceUpload)
class TrialBalanceUploadAdmin(admin.ModelAdmin):
    list_display = ('report_date', 'uploaded_by', 'file_name', 'row_count', 'created_at')
    list_filter = ('report_date',)
    search_fields = ('file_name', 'uploaded_by__full_name')
    readonly_fields = ('created_at', 'updated_at')
    inlines = [TrialBalanceEntryInline]
    date_hierarchy = 'report_date'


@admin.register(TrialBalanceEntry)
class TrialBalanceEntryAdmin(admin.ModelAdmin):
    list_display = (
        'report_date', 'branch_code', 'category', 'gl_code',
        'descr', 'ccy', 'close_bal_lcy', 'close_bal_fcy', 'gl_status',
    )
    list_filter = ('report_date', 'gl_status', 'ccy', 'branch_code')
    search_fields = ('gl_code', 'descr', 'branch_code')
    readonly_fields = ('created_at',)
    date_hierarchy = 'report_date'


# ==================== SENT EMAIL ====================

@admin.register(SentEmail)
class SentEmailAdmin(admin.ModelAdmin):
    list_display = (
        'subject', 'sender', 'report_type',
        'status', 'recipient_count', 'final_score',
        'sent_at',
    )
    list_filter = ('status', 'report_type')
    search_fields = ('subject', 'report_type', 'sender__full_name',
                     'to_addresses', 'cc_addresses')
    readonly_fields = ('sent_at', 'final_score', 'badge_class', 'is_overridden')
    date_hierarchy = 'sent_at'

    fieldsets = (
        ('Audit', {
            'fields': ('sender', 'report_type', 'subject', 'body'),
        }),
        ('Recipients', {
            'fields': ('to_addresses', 'cc_addresses'),
        }),
        ('Delivery', {
            'fields': ('status', 'error_message', 'sent_at'),
        }),
        ('Supervisor scoring', {
            'fields': ('manual_deduction', 'override_reason',
                       'scored_by', 'scored_at',
                       'final_score', 'badge_class', 'is_overridden'),
        }),
    )


# ==================== AD-HOC SCORING ====================

@admin.register(AdHocDeduction)
class AdHocDeductionAdmin(admin.ModelAdmin):
    list_display = ('user', 'task_description', 'points', 'points_added',
                    'created_by', 'created_at')
    list_filter = ('points', 'points_added')
    search_fields = ('user__full_name', 'user__email', 'task_description', 'reason')
    readonly_fields = ('created_at', 'updated_at')
    date_hierarchy = 'created_at'


# ==================== ACTIVITY LOG ====================

@admin.register(ActivityLog)
class ActivityLogAdmin(admin.ModelAdmin):
    list_display = ('user', 'activity_type', 'short_details', 'ip_address', 'created_at')
    list_filter = ('activity_type',)
    search_fields = ('user__full_name', 'user__email', 'details')
    readonly_fields = ('user', 'activity_type', 'details',
                       'ip_address', 'user_agent', 'created_at')
    date_hierarchy = 'created_at'

    def short_details(self, obj):
        text = obj.details or ''
        return text[:80] + ('…' if len(text) > 80 else '')
    short_details.short_description = 'Details'

    def has_add_permission(self, request):
        # ActivityLog rows are created by the app, never by hand.
        return False


# ==================== CHECKLIST CHANGE REQUESTS ====================

@admin.register(ChecklistChangeRequest)
class ChecklistChangeRequestAdmin(admin.ModelAdmin):
    list_display = (
        'id', 'request_type', 'status',
        'requester', 'target_checklist',
        'reviewed_by', 'created_at',
    )
    list_filter = ('request_type', 'status')
    search_fields = ('requester__full_name', 'requester__email',
                     'target_checklist__name', 'justification')
    readonly_fields = ('created_at', 'updated_at',
                       'reviewed_at', 'created_checklist')
    date_hierarchy = 'created_at'

    fieldsets = (
        ('Request', {
            'fields': ('requester', 'request_type', 'status',
                       'target_checklist', 'justification'),
        }),
        ('Payload', {
            'fields': ('payload', 'current_snapshot'),
        }),
        ('Review', {
            'fields': ('reviewed_by', 'reviewed_at', 'review_note',
                       'created_checklist'),
        }),
        ('Meta', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',),
        }),
    )