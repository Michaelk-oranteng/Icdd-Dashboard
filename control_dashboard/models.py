# control_dashboard/models.py

from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from django.db.models.signals import pre_save, post_save
from django.dispatch import receiver
import re
from datetime import date, timedelta
import hashlib
from urllib.parse import urlencode


# ============================================================
# CONSTANTS — single source of truth for status values
# ============================================================

# The Report.status value that marks a Report row as an
# Excel data container (created by draft.html). These are
# NOT submissions — they are raw exception data.
#
# DEPRECATED: Excel uploads now live on `ExceptionUpload`.
# Kept here so legacy code and migrations still resolve.
UPLOADED_STATUS = 'uploaded'

# Statuses that count as "actually submitted by a member".
# Used by Report.is_submitted — must exist before any class
# references it.
SUBMITTED_STATUSES = ('submitted', 'completed', 'approved')


class UserProfile(models.Model):
    """
    User Profile model for storing user information.
    """
    ROLE_CHOICES = [
        ('admin', 'Admin'),
        ('supervisor', 'Supervisor'),
        ('member', 'Member'),
    ]

    POSITION_CHOICES = [
        ('hc', 'Headoffice Control'),
        ('cc', 'Cluster Control'),
    ]

    STATUS_CHOICES = [
        ('active', 'Active'),
        ('inactive', 'Inactive'),
    ]

    # Link to Django's built-in User model
    user = models.OneToOneField(
        User,
        on_delete=models.CASCADE,
        related_name='profile',
        null=True,
        blank=True
    )

    email = models.EmailField(unique=True)
    full_name = models.CharField(max_length=200)
    username = models.CharField(max_length=150, unique=True, blank=True, null=True)
    position = models.CharField(max_length=50, choices=POSITION_CHOICES, default='member')
    role = models.CharField(max_length=50, choices=ROLE_CHOICES, default='member')
    status = models.CharField(max_length=50, choices=STATUS_CHOICES, default='active')

    # ============================================
    # PROFILE PICTURE
    # Admin-uploaded photo. If empty, avatar_url
    # falls back to a Gravatar identicon.
    # ============================================
    avatar = models.ImageField(
        upload_to='avatars/',
        null=True,
        blank=True,
        help_text='Admin-uploaded profile picture (optional).',
    )

    # ============================================
    # DEPARTMENT AND BRANCH ASSIGNMENTS (Many-to-Many)
    # ============================================
    departments = models.ManyToManyField(
        'Department',
        blank=True,
        related_name='user_profiles'
    )
    branches = models.ManyToManyField(
        'Branch',
        blank=True,
        related_name='user_profiles'
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.full_name or self.email

    def generate_username_from_email(self):
        """
        Generate username from email address.
        Example: john.doe@company.com -> john.doe
        """
        if not self.email:
            return None

        # Remove everything after @
        username = self.email.split('@')[0]

        # Remove any special characters except dot and underscore
        username = re.sub(r'[^a-zA-Z0-9._]', '', username)

        # Convert to lowercase
        username = username.lower()

        # Handle edge cases
        if not username:
            username = f"user_{self.id}" if self.id else "user_temp"

        # Make it unique if it already exists
        original_username = username
        counter = 1
        while UserProfile.objects.filter(username=username).exclude(id=self.id).exists():
            username = f"{original_username}{counter}"
            counter += 1

        return username

    def create_django_user(self, password=None):
        """
        Create a Django User from this profile.
        """
        if self.user:
            return self.user

        # Generate username if not set
        if not self.username:
            self.username = self.generate_username_from_email()

        # Create Django user
        if password:
            user = User.objects.create_user(
                username=self.username,
                email=self.email,
                password=password,
            )
        else:
            user = User.objects.create_user(
                username=self.username,
                email=self.email,
            )
            user.set_unusable_password()
            user.save()

        # Set full name
        name_parts = self.full_name.split(' ', 1)
        user.first_name = name_parts[0]
        user.last_name = name_parts[1] if len(name_parts) > 1 else ''
        user.save()

        self.user = user
        self.save()

        return user

    def get_department_names(self):
        """Get comma-separated department names."""
        return ', '.join([d.name for d in self.departments.all()])

    def get_branch_names(self):
        """Get comma-separated branch names."""
        return ', '.join([b.name for b in self.branches.all()])

    # ============================================================
    # AVATAR & NAME HELPERS
    # ============================================================
    @property
    def avatar_url(self):
        """
        Return a profile picture URL for this user.

        Priority order:
          1. Admin-uploaded photo (self.avatar)
          2. Gravatar identicon derived from the email
        """
        # 1. Admin-uploaded image
        if self.avatar and hasattr(self.avatar, 'url'):
            try:
                return self.avatar.url
            except (ValueError, AttributeError):
                pass

        # 2. Gravatar fallback
        email = (self.email or '').strip().lower().encode('utf-8')
        digest = hashlib.md5(email).hexdigest()
        params = urlencode({
            'd': 'identicon',   # identicon | mp | retro | robohash | wavatar | monsterid
            's': '200',
            'r': 'pg',
        })
        return f'https://www.gravatar.com/avatar/{digest}?{params}'

    @property
    def initials(self):
        """
        First letters of the full name — e.g. 'Michael Koranteng' → 'MK'.
        Used as a text-based fallback when the image fails to load.
        """
        name = (self.full_name or self.email or '?').strip()
        parts = [p for p in name.split() if p]
        if len(parts) >= 2:
            return (parts[0][0] + parts[-1][0]).upper()
        return name[:2].upper()

    @property
    def first_name(self):
        """
        First name only — e.g. 'Michael Koranteng' → 'Michael'.
        Falls back to the email prefix if full_name is empty.
        """
        if not self.full_name:
            if self.email and '@' in self.email:
                return self.email.split('@')[0]
            return 'User'
        parts = [p for p in self.full_name.strip().split() if p]
        return parts[0] if parts else 'User'

    @property
    def last_name(self):
        """
        Last name only — everything after the first space.
        'Michael Koranteng' → 'Koranteng'
        'Mary Jane Watson'  → 'Jane Watson'
        """
        if not self.full_name:
            return ''
        parts = [p for p in self.full_name.strip().split() if p]
        if len(parts) < 2:
            return ''
        return ' '.join(parts[1:])

    class Meta:
        db_table = 'user_profiles'
        ordering = ['full_name']


# ============================================
# BRANCH AND DEPARTMENT MODELS
# ============================================

class Branch(models.Model):
    """
    Bank branch.

    `assignment_scope` controls who can see this branch:
      - 'cc'       → visible to all Cluster Control staff
      - 'hc'       → visible to all Head Office Control staff
      - 'specific' → visible only to users in `assigned_users`
    """

    ASSIGNMENT_SCOPE_CHOICES = [
        ('cc', 'Cluster Control'),
        ('hc', 'Head Office Control'),
        ('specific', 'Specific Users'),
    ]

    BRANCH_CODE_MAP = {
        # ... unchanged ...
    }

    BRANCH_CODE_CHOICES = [
        (code, f"{code} — {name}")
        for code, name in sorted(BRANCH_CODE_MAP.items())
    ]

    name = models.CharField(max_length=100, unique=True)
    code = models.CharField(max_length=20, unique=True, blank=True, null=True)
    description = models.TextField(blank=True, default='')

    assignment_scope = models.CharField(
        max_length=20,
        choices=ASSIGNMENT_SCOPE_CHOICES,
        default='cc',
        help_text="Who can see this branch in their dashboard.",
    )
    assigned_users = models.ManyToManyField(
        'UserProfile',
        blank=True,
        related_name='scoped_branches',
        help_text="Only used when assignment_scope='specific'.",
    )

    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

    class Meta:
        db_table = 'branches'
        ordering = ['name']
        verbose_name_plural = 'Branches'


class Department(models.Model):
    """Bank department — same assignment semantics as Branch."""

    ASSIGNMENT_SCOPE_CHOICES = [
        ('cc', 'Cluster Control'),
        ('hc', 'Head Office Control'),
        ('specific', 'Specific Users'),
    ]

    name = models.CharField(max_length=100, unique=True)
    code = models.CharField(max_length=20, unique=True, blank=True, null=True)
    description = models.TextField(blank=True, default='')

    assignment_scope = models.CharField(
        max_length=20,
        choices=ASSIGNMENT_SCOPE_CHOICES,
        default='cc',
        help_text="Who can see this department in their dashboard.",
    )
    assigned_users = models.ManyToManyField(
        'UserProfile',
        blank=True,
        related_name='scoped_departments',
        help_text="Only used when assignment_scope='specific'.",
    )

    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

    class Meta:
        db_table = 'departments'
        ordering = ['name']
        verbose_name_plural = 'Departments'


# ============================================
# SIGNALS: Auto-generate username and create user
# ============================================

@receiver(pre_save, sender=UserProfile)
def auto_generate_username_fallback(sender, instance, **kwargs):
    """
    Fallback only — generate a username when one wasn't supplied.

    The admin Create User form and the API always send a username
    explicitly, so this signal only fires for legacy code paths
    (e.g. UserProfile.objects.create(...) without a username).
    """
    if not instance.username and instance.email:
        instance.username = instance.generate_username_from_email()


@receiver(post_save, sender=UserProfile)
def create_user_for_profile(sender, instance, created, **kwargs):
    """Automatically create Django User when UserProfile is created."""
    if created and not instance.user:
        try:
            instance.create_django_user()
        except Exception as e:
            print(f"Error creating user for {instance.email}: {e}")


# ============================================
# REPORT MODELS
# ============================================

class Report(models.Model):
    """
    Model for storing reports.

    After the Trial Balance / Exception upload split, `Report`
    is used ONLY for:

    1. ADMIN-CREATED RECURRING REPORTS
       (assigned / in_progress / submitted / completed /
        approved / rejected / draft)
       Created via report_creation.html. Have a frequency and a
       deadline_date anchor.

    2. CONSOLIDATED REPORT DEFINITIONS
       Report rows whose report_type contains 'consolidated'.
       These are assignment containers, not data.

    Excel exception uploads now live on `ExceptionUpload`.
    Trial balance uploads now live on `TrialBalanceUpload`.
    """
    FREQUENCY_CHOICES = [
        ('one-off', 'One Off'),
        ('daily', 'Daily'),
        ('weekly', 'Weekly'),
        ('monthly', 'Monthly'),
        ('quarterly', 'Quarterly'),
        ('yearly', 'Yearly'),
    ]

    STATUS_CHOICES = [
        ('assigned', 'Assigned'),
        ('in_progress', 'In Progress'),
        ('submitted', 'Submitted'),
        ('completed', 'Completed'),
        ('approved', 'Approved'),
        ('rejected', 'Rejected'),
        ('draft', 'Draft'),
        # Legacy — kept so any old rows still resolve their display value.
        ('uploaded', 'Uploaded (Legacy — do not use)'),
    ]

    report_type = models.CharField(max_length=200)
    frequency = models.CharField(max_length=50, choices=FREQUENCY_CHOICES, default='one-off')
    description = models.TextField(blank=True)
    deadline_date = models.DateField(null=True, blank=True)
    deadline_time = models.TimeField(null=True, blank=True)
    assigned_to = models.ManyToManyField('UserProfile', blank=True, related_name='assigned_reports')
    is_assigned_to_all = models.BooleanField(default=False)
    status = models.CharField(max_length=50, choices=STATUS_CHOICES, default='assigned')
    created_by = models.ForeignKey('UserProfile', on_delete=models.CASCADE, related_name='created_reports')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    # Legacy field — only used by old Report rows that were uploads.
    # New uploads no longer touch this.
    excel_headers = models.TextField(blank=True, default='')

    def __str__(self):
        return self.report_type

    def get_frequency_display(self):
        return dict(self.FREQUENCY_CHOICES).get(self.frequency, self.frequency)

    def get_status_display(self):
        return dict(self.STATUS_CHOICES).get(self.status, self.status)

    @property
    def is_exception_upload(self):
        """Legacy: only True for old Report rows that were uploads."""
        return self.status == UPLOADED_STATUS

    @property
    def is_submitted(self):
        return self.status in SUBMITTED_STATUSES

    def get_excel_headers_list(self):
        if not self.excel_headers:
            return []
        return [h.strip() for h in self.excel_headers.split(',') if h.strip()]

    class Meta:
        db_table = 'reports'
        ordering = ['-created_at']

    def get_display_data(self):
        """
        Legacy display builder for old Report rows that were uploads.

        New uploads are rendered from ExceptionUpload, not from here.
        """
        display_data = {
            'rows': [],
            'headers': [],
            'is_excel': False,
            'row_count': 0,
        }

        if self.status == UPLOADED_STATUS:
            headers = self.get_excel_headers_list()

            if not headers:
                headers = [
                    'S/N', 'BRANCH/UNIT', 'EXCEPTION',
                    'DATE EXCEPTION WAS NOTED', 'TARGET DATE FOR CLOSURE',
                    'CATEGORY OF EXCEPTION', 'RESPONSIBLE OFFICER',
                    'SUPERVISOR', "AUDITEE'S RESPONSE", 'REMARKS',
                    'STATUS', 'INCOME/COST SAVED',
                ]

            display_data['is_excel'] = True
            display_data['headers'] = headers

            # NOTE: legacy rows still point at Report via ExceptionRecord.report.
            # Once the data migration runs, that FK will be removed from
            # ExceptionRecord and this block will be dead code.
            for rec in self.exception_records.order_by('source_row_index'):
                full_row = {
                    'row_id': rec.id,
                    'S/N': rec.serial_number,
                    'BRANCH/UNIT': rec.branch_unit,
                    'EXCEPTION': rec.exception,
                    'DATE EXCEPTION WAS NOTED': rec.date_noted_raw or (str(rec.date_noted) if rec.date_noted else ''),
                    'TARGET DATE FOR CLOSURE': rec.target_closure_date_raw or (str(rec.target_closure_date) if rec.target_closure_date else ''),
                    'CATEGORY OF EXCEPTION': rec.category,
                    'RESPONSIBLE OFFICER': rec.responsible_officer,
                    'SUPERVISOR': rec.supervisor,
                    "AUDITEE'S RESPONSE": rec.auditee_response,
                    'REMARKS': rec.remarks,
                    'STATUS': rec.status_raw or rec.get_status_display(),
                    'INCOME/COST SAVED': rec.income_cost_saved_raw or str(rec.income_cost_saved),
                }
                row_view = {h: full_row.get(h, '') for h in headers}
                row_view['row_id'] = rec.id
                display_data['rows'].append(row_view)

            display_data['row_count'] = len(display_data['rows'])
            return display_data

        data_fields = self.data_fields.all()
        if data_fields.exists():
            row_dict = {}
            for field in data_fields:
                field_name = field.field_name
                field_value = field.field_value or ''
                if field_name.startswith('Column_') and not field_value:
                    continue
                row_dict[field_name] = field_value

            if row_dict:
                display_data['rows'].append(row_dict)
                display_data['headers'] = list(row_dict.keys())
                display_data['row_count'] = 1
            return display_data

        display_data['headers'] = ['Branch/Unit', 'Date', 'Observation', 'Responsible Staff', 'Status']
        display_data['row_count'] = 0
        return display_data


# ============================================
# TRIAL BALANCE UPLOAD — dedicated model
# ============================================

class TrialBalanceUpload(models.Model):
    """
    One row = one uploaded Trial Balance Excel file for a given date.

    Replaces the old pattern where each upload created a Report
    row with report_type='__TRIAL_BALANCE__'. This table now owns
    the header metadata; the rows of the file live on
    `TrialBalanceEntry`, which FKs back to this model.
    """

    uploaded_by = models.ForeignKey(
        UserProfile,
        on_delete=models.CASCADE,
        related_name='trial_balance_uploads',
        help_text='The user who uploaded this file.',
    )
    report_date = models.DateField(
        db_index=True,
        help_text='The trading day this trial balance applies to.',
    )
    file_name = models.CharField(max_length=255, blank=True, default='')
    row_count = models.IntegerField(
        default=0,
        help_text='Number of data rows stored for this upload.',
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'trial_balance_uploads'
        ordering = ['-report_date', '-created_at']
        # One trial balance per calendar day, globally.
        constraints = [
            models.UniqueConstraint(
                fields=['report_date'],
                name='unique_trial_balance_per_day',
            ),
        ]
        indexes = [
            models.Index(fields=['report_date']),
            models.Index(fields=['uploaded_by', 'report_date']),
        ]

    def __str__(self):
        return f'{self.report_date} — {self.file_name or "(no file)"} — {self.row_count} rows'


# ============================================
# EXCEPTION UPLOAD — dedicated model
# ============================================

class ExceptionUpload(models.Model):
    """
    One row = one Excel exception file uploaded from draft.html.

    Replaces the old pattern where each upload created a Report
    row with status='uploaded'. This table now owns the header
    metadata; the typed rows of the file live on
    `ExceptionRecord`, which FKs back to this model.
    """

    uploaded_by = models.ForeignKey(
        UserProfile,
        on_delete=models.CASCADE,
        related_name='exception_uploads',
        help_text='The user who uploaded this file.',
    )
    report_type = models.CharField(
        max_length=200,
        db_index=True,
        help_text="User-chosen report type (e.g. 'Weekly Exceptions Report').",
    )
    file_name = models.CharField(max_length=255, blank=True, default='')
    excel_headers = models.TextField(
        blank=True,
        default='',
        help_text='Comma-separated original Excel headers, preserved for rendering.',
    )
    row_count = models.IntegerField(
        default=0,
        help_text='Number of exception rows stored for this upload.',
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'exception_uploads'
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['report_type']),
            models.Index(fields=['uploaded_by', 'created_at']),
        ]

    def __str__(self):
        return f'{self.report_type} — {self.file_name or "(no file)"} — {self.row_count} rows'

    # ------------------------------------------------------------
    # Header helpers
    # ------------------------------------------------------------
    def get_excel_headers_list(self):
        if not self.excel_headers:
            return []
        return [h.strip() for h in self.excel_headers.split(',') if h.strip()]


# ============================================
# REPORT SUBMISSION MODELS
# ============================================

class ReportDataField(models.Model):
    """Model for storing report data fields (replaces JSON)."""
    FIELD_TYPES = [
        ('text', 'Text'),
        ('number', 'Number'),
        ('date', 'Date'),
        ('email', 'Email'),
        ('url', 'URL'),
        ('textarea', 'Text Area'),
    ]

    report = models.ForeignKey('Report', on_delete=models.CASCADE, related_name='data_fields')
    field_name = models.CharField(max_length=255)
    field_value = models.TextField(blank=True, null=True)
    field_type = models.CharField(max_length=50, choices=FIELD_TYPES, default='text')
    order = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'report_data_fields'
        ordering = ['order']
        unique_together = ['report', 'field_name']

    def __str__(self):
        return f"{self.report.report_type} - {self.field_name}"


class ReportSubmission(models.Model):
    """Model for storing member report submissions."""
    STATUS_CHOICES = [
        ('submitted', 'Submitted'),
    ]

    report_type = models.CharField(max_length=200)
    submitted_by = models.ForeignKey('UserProfile', on_delete=models.CASCADE, related_name='report_submissions')
    submission_date = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    status = models.CharField(max_length=50, choices=STATUS_CHOICES, default='submitted')
    notes = models.TextField(blank=True)

    def __str__(self):
        return f"{self.report_type} - {self.submitted_by.full_name} - {self.submission_date.strftime('%Y-%m-%d')}"

    class Meta:
        db_table = 'report_submissions'
        ordering = ['-submission_date']


class ReportSubmissionField(models.Model):
    """Model for storing submission form data (replaces JSON)."""
    FIELD_TYPES = [
        ('text', 'Text'),
        ('number', 'Number'),
        ('date', 'Date'),
        ('email', 'Email'),
        ('url', 'URL'),
        ('textarea', 'Text Area'),
        ('select', 'Select'),
        ('checkbox', 'Checkbox'),
        ('radio', 'Radio'),
    ]

    submission = models.ForeignKey('ReportSubmission', on_delete=models.CASCADE, related_name='fields')
    field_key = models.CharField(max_length=255)
    field_value = models.TextField(blank=True, null=True)
    field_type = models.CharField(max_length=50, choices=FIELD_TYPES, default='text')
    field_label = models.CharField(max_length=255, blank=True)
    order = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'report_submission_fields'
        ordering = ['order']
        unique_together = ['submission', 'field_key']

    def __str__(self):
        return f"{self.submission.report_type} - {self.field_key}"


# ============================================
# CHECKLIST MODELS
# ============================================

class Checklist(models.Model):
    """Model for storing checklists/activities."""
    FREQUENCY_CHOICES = [
        ('daily', 'Daily'),
        ('weekly', 'Weekly'),
        ('monthly', 'Monthly'),
        ('quarterly', 'Quarterly'),
        ('bi-annual', 'Bi-Annual'),
        ('annual', 'Annual'),
    ]

    ASSIGNMENT_CHOICES = [
        ('all', 'All Users'),
        ('cc', 'Cluster Control'),
        ('hc', 'Head Office Control'),
        ('specific', 'Specific Users'),
    ]

    name = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    frequency = models.CharField(max_length=50, choices=FREQUENCY_CHOICES, default='weekly')
    assignment_target = models.CharField(max_length=50, choices=ASSIGNMENT_CHOICES, default='all')
    assigned_users = models.ManyToManyField('UserProfile', blank=True, related_name='assigned_checklists')
    is_active = models.BooleanField(default=True)
    created_by = models.ForeignKey('UserProfile', on_delete=models.SET_NULL, null=True, related_name='created_checklists')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    assigned_branches = models.ManyToManyField('Branch', blank=True, related_name='checklist_assignments')
    assigned_departments = models.ManyToManyField('Department', blank=True, related_name='checklist_assignments')

    def __str__(self):
        return self.name

    def get_frequency_display(self):
        return dict(self.FREQUENCY_CHOICES).get(self.frequency, self.frequency)

    def get_assignment_display(self):
        return dict(self.ASSIGNMENT_CHOICES).get(self.assignment_target, self.assignment_target)

    def get_assigned_users_display(self):
        if self.assignment_target == 'all':
            return 'All Users'
        elif self.assignment_target == 'cc':
            return 'Cluster Control'
        elif self.assignment_target == 'hc':
            return 'Head Office Control'
        else:
            users = self.assigned_users.all()
            if users:
                return ', '.join([user.full_name for user in users])
            return 'No users assigned'

    def get_frequency_days(self):
        frequency_map = {
            'daily': 1,
            'weekly': 7,
            'monthly': 30,
            'quarterly': 91,
            'bi-annual': 182,
            'annual': 365,
        }
        return frequency_map.get(self.frequency, 7)

    def get_expected_occurrences(self, start_date=None, end_date=None):
        if not start_date:
            start_date = timezone.now().date().replace(month=1, day=1)
        if not end_date:
            end_date = timezone.now().date()

        if end_date < start_date:
            start_date, end_date = end_date, start_date

        if self.frequency == 'annual':
            years = end_date.year - start_date.year + 1
            return years

        if self.frequency == 'daily':
            current = start_date
            count = 0
            while current <= end_date:
                if current.weekday() < 5:
                    count += 1
                current += timedelta(days=1)
            return count

        if self.frequency == 'monthly':
            months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) + 1
            return months

        if self.frequency == 'weekly':
            days_diff = (end_date - start_date).days + 1
            weeks = days_diff // 7
            return max(1, weeks)

        if self.frequency == 'quarterly':
            days_diff = (end_date - start_date).days + 1
            quarters = days_diff // 91
            return max(1, quarters)

        if self.frequency == 'bi-annual':
            days_diff = (end_date - start_date).days + 1
            half_years = days_diff // 182
            return max(1, half_years)

        days_diff = (end_date - start_date).days + 1
        frequency_days = self.get_frequency_days()
        return max(1, days_diff // frequency_days)

    def get_completion_percentage(self, user_profile, start_date=None, end_date=None):
        if not start_date:
            start_date = timezone.now().date().replace(month=1, day=1)
        if not end_date:
            end_date = timezone.now().date()

        if end_date < start_date:
            start_date, end_date = end_date, start_date

        expected = self.get_expected_occurrences(start_date, end_date)
        if expected == 0:
            return 0

        actual = ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile,
            log_date__gte=start_date,
            log_date__lte=end_date
        ).count()

        percentage = int((actual / expected) * 100)
        return min(percentage, 100)

    def get_monthly_completion(self, user_profile, month=None, year=None):
        today = timezone.now().date()
        if month is None:
            month = today.month
        if year is None:
            year = today.year

        start_date = date(year, month, 1)
        if month == 12:
            end_date = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(year, month + 1, 1) - timedelta(days=1)

        return self.get_completion_percentage(user_profile, start_date, end_date)

    def get_year_to_date_completion(self, user_profile):
        today = timezone.now().date()
        start_date = date(today.year, 1, 1)
        return self.get_completion_percentage(user_profile, start_date, today)

    def get_monthly_expected(self, user_profile, month=None, year=None):
        today = timezone.now().date()
        if month is None:
            month = today.month
        if year is None:
            year = today.year

        start_date = date(year, month, 1)
        if month == 12:
            end_date = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(year, month + 1, 1) - timedelta(days=1)

        return self.get_expected_occurrences(start_date, end_date)

    def get_monthly_actual(self, user_profile, month=None, year=None):
        today = timezone.now().date()
        if month is None:
            month = today.month
        if year is None:
            year = today.year

        start_date = date(year, month, 1)
        if month == 12:
            end_date = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(year, month + 1, 1) - timedelta(days=1)

        return ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile,
            log_date__gte=start_date,
            log_date__lte=end_date
        ).count()

    def get_year_to_date_expected(self, user_profile):
        today = timezone.now().date()
        start_date = date(today.year, 1, 1)
        return self.get_expected_occurrences(start_date, today)

    def get_year_to_date_actual(self, user_profile):
        today = timezone.now().date()
        start_date = date(today.year, 1, 1)
        return ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile,
            log_date__gte=start_date,
            log_date__lte=today
        ).count()

    def get_next_due_date(self, user_profile):
        if self.frequency == 'one-off' or self.frequency == 'annual':
            return None

        last_log = ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile
        ).order_by('-log_date').first()

        if not last_log:
            today = timezone.now().date()
            next_date = today
            if self.frequency == 'daily':
                while next_date.weekday() >= 5:
                    next_date += timedelta(days=1)
            return next_date

        if self.frequency == 'daily':
            next_date = last_log.log_date + timedelta(days=1)
            while next_date.weekday() >= 5:
                next_date += timedelta(days=1)
        elif self.frequency == 'weekly':
            next_date = last_log.log_date + timedelta(days=7)
        elif self.frequency == 'monthly':
            if last_log.log_date.month == 12:
                next_date = date(last_log.log_date.year + 1, 1, last_log.log_date.day)
            else:
                next_date = date(last_log.log_date.year, last_log.log_date.month + 1, last_log.log_date.day)
        elif self.frequency == 'quarterly':
            next_date = last_log.log_date + timedelta(days=91)
        elif self.frequency == 'bi-annual':
            next_date = last_log.log_date + timedelta(days=182)
        else:
            days = self.get_frequency_days()
            next_date = last_log.log_date + timedelta(days=days)

        if next_date < timezone.now().date():
            next_date = timezone.now().date()
            if self.frequency == 'daily':
                while next_date.weekday() >= 5:
                    next_date += timedelta(days=1)

        return next_date

    def get_actual_completion_count(self, user_profile, start_date=None, end_date=None):
        if not start_date:
            start_date = timezone.now().date().replace(month=1, day=1)
        if not end_date:
            end_date = timezone.now().date()

        if end_date < start_date:
            start_date, end_date = end_date, start_date

        return ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile,
            log_date__gte=start_date,
            log_date__lte=end_date
        ).count()

    class Meta:
        db_table = 'checklists'
        ordering = ['name']


class ReportSchedule(models.Model):
    """Model for storing report schedules and predicting next due dates."""

    FREQUENCY_CHOICES = [
        ('daily', 'Daily'),
        ('weekly', 'Weekly'),
        ('monthly', 'Monthly'),
        ('quarterly', 'Quarterly'),
        ('yearly', 'Yearly'),
        ('one-off', 'One-off'),
    ]

    report = models.ForeignKey('Report', on_delete=models.CASCADE, related_name='schedules')
    frequency = models.CharField(max_length=50, choices=FREQUENCY_CHOICES, default='weekly')
    start_date = models.DateField()
    end_date = models.DateField(null=True, blank=True)
    due_time = models.TimeField(null=True, blank=True)
    last_submitted = models.DateTimeField(null=True, blank=True)
    next_due_date = models.DateField(null=True, blank=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'report_schedules'
        ordering = ['next_due_date']

    def calculate_next_due_date(self):
        if not self.last_submitted:
            return self.start_date

        last_date = self.last_submitted.date()

        frequency_map = {
            'daily': 1,
            'weekly': 7,
            'monthly': 30,
            'quarterly': 91,
            'yearly': 365,
            'one-off': None,
        }

        days = frequency_map.get(self.frequency)
        if not days:
            return None

        next_date = last_date + timedelta(days=days)

        if self.frequency == 'daily':
            while next_date.weekday() >= 5:
                next_date += timedelta(days=1)

        if self.end_date and next_date > self.end_date:
            return None

        return next_date

    def save(self, *args, **kwargs):
        if not self.next_due_date:
            self.next_due_date = self.calculate_next_due_date()
        super().save(*args, **kwargs)

    def get_frequency_display(self):
        return dict(self.FREQUENCY_CHOICES).get(self.frequency, self.frequency)


class ChecklistTask(models.Model):
    """Model for storing checklist tasks."""
    checklist = models.ForeignKey(Checklist, on_delete=models.CASCADE, related_name='tasks')
    description = models.CharField(max_length=500)
    order = models.IntegerField(default=0)
    is_completed = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.checklist.name} - {self.description[:50]}"

    class Meta:
        db_table = 'checklist_tasks'
        ordering = ['order']


class ChecklistLog(models.Model):
    checklist = models.ForeignKey(
        Checklist, on_delete=models.CASCADE, related_name='logs'
    )
    user = models.ForeignKey(
        UserProfile, on_delete=models.CASCADE, related_name='checklist_logs'
    )
    log_date = models.DateField()
    branch = models.ForeignKey(
        Branch, on_delete=models.CASCADE,
        null=True, blank=True, related_name='checklist_logs'
    )
    department = models.ForeignKey(
        Department, on_delete=models.CASCADE,
        null=True, blank=True, related_name='checklist_logs'
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['checklist', 'user', 'branch', 'department', 'log_date'],
                name='unique_checklist_log_per_unit_per_day'
            )
        ]

    def __str__(self):
        unit = self.branch.name if self.branch else (
            self.department.name if self.department else 'General'
        )
        return f"{self.checklist.name} — {unit} — {self.log_date}"


class ActivityLog(models.Model):
    """Model to track user activities across the application."""
    ACTIVITY_TYPES = (
        ('login', 'Login'),
        ('logout', 'Logout'),
        ('report_created', 'Report Created'),
        ('report_updated', 'Report Updated'),
        ('report_deleted', 'Report Deleted'),
        ('report_submitted', 'Report Submitted'),
        ('report_approved', 'Report Approved'),
        ('report_rejected', 'Report Rejected'),
        ('checklist_completed', 'Checklist Completed'),
        ('user_created', 'User Created'),
        ('user_updated', 'User Updated'),
        ('user_deleted', 'User Deleted'),
        ('email_sent', 'Email Sent'),
        ('draft_saved', 'Draft Saved'),
        ('draft_deleted', 'Draft Deleted'),
        ('deduction_created', 'Deduction Created'),
        ('deduction_updated', 'Deduction Updated'),
        ('deduction_deleted', 'Deduction Deleted'),
        ('score_updated', 'Score Updated'),
    )

    ACTIVITY_ICONS = {
        'login': 'fa-sign-in-alt',
        'logout': 'fa-sign-out-alt',
        'report_created': 'fa-file-alt',
        'report_updated': 'fa-edit',
        'report_deleted': 'fa-trash',
        'report_submitted': 'fa-paper-plane',
        'report_approved': 'fa-check-circle',
        'report_rejected': 'fa-times-circle',
        'checklist_completed': 'fa-check-double',
        'user_created': 'fa-user-plus',
        'user_updated': 'fa-user-edit',
        'user_deleted': 'fa-user-minus',
        'email_sent': 'fa-envelope',
        'draft_saved': 'fa-save',
        'draft_deleted': 'fa-trash-alt',
        'deduction_created': 'fa-minus-circle',
        'deduction_updated': 'fa-edit',
        'deduction_deleted': 'fa-trash',
        'score_updated': 'fa-star',
    }

    user = models.ForeignKey(
        'UserProfile',
        on_delete=models.CASCADE,
        related_name='activity_logs'
    )
    activity_type = models.CharField(
        max_length=50,
        choices=ACTIVITY_TYPES
    )
    details = models.TextField(blank=True, null=True)
    ip_address = models.GenericIPAddressField(blank=True, null=True)
    user_agent = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Activity Log'
        verbose_name_plural = 'Activity Logs'
        indexes = [
            models.Index(fields=['user', 'created_at']),
            models.Index(fields=['activity_type']),
        ]

    def __str__(self):
        return f"{self.user.full_name} - {self.get_activity_type_display()} - {self.created_at.strftime('%Y-%m-%d %H:%M')}"

    def get_activity_icon(self):
        return self.ACTIVITY_ICONS.get(self.activity_type, 'fa-circle')


class AdHocDeduction(models.Model):
    """Ad-hoc scoring event for a team member."""
    user = models.ForeignKey(
        'UserProfile',
        on_delete=models.CASCADE,
        related_name='ad_hoc_deductions'
    )
    task_description = models.CharField(max_length=255)
    points = models.IntegerField(default=0)
    points_added = models.IntegerField(default=0)
    reason = models.TextField(blank=True, null=True)
    created_by = models.ForeignKey(
        'UserProfile',
        on_delete=models.SET_NULL,
        null=True,
        related_name='created_deductions'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Ad-Hoc Score Entry'
        verbose_name_plural = 'Ad-Hoc Score Entries'
        indexes = [
            models.Index(fields=['user', 'created_at']),
            models.Index(fields=['points']),
            models.Index(fields=['points_added']),
        ]

    def __str__(self):
        parts = []
        if self.points:
            parts.append(f"-{self.points}")
        if self.points_added:
            parts.append(f"+{self.points_added}")
        score_str = " / ".join(parts) if parts else "0"
        return f"{self.user.full_name} — {score_str} — {self.task_description[:30]}"

    def get_points_display(self):
        parts = []
        if self.points:
            parts.append(f"-{self.points}%")
        if self.points_added:
            parts.append(f"+{self.points_added}%")
        return " ".join(parts) if parts else "0%"

    def get_badge_class(self):
        if self.points >= 20:
            return 'high'
        if self.points_added >= 20:
            return 'high-positive'
        if self.points >= 10:
            return 'medium'
        if self.points_added >= 10:
            return 'medium-positive'
        if self.points_added:
            return 'low-positive'
        return 'low'


# ============================================
# TRIAL BALANCE ENTRY — now FK to TrialBalanceUpload
# ============================================

class TrialBalanceEntry(models.Model):
    """
    One row of an uploaded Daily Trial Balance file.

    The header metadata lives on `TrialBalanceUpload`; this row
    belongs to exactly one upload.
    """
    upload = models.ForeignKey(
        TrialBalanceUpload,
        on_delete=models.CASCADE,
        related_name='entries',
        null=True, blank=True,
        help_text='The upload this row belongs to.',
    )
    uploaded_by = models.ForeignKey(
        UserProfile,
        on_delete=models.CASCADE,
        related_name='trial_balance_entries'
    )
    report_date = models.DateField(db_index=True)
    file_name = models.CharField(max_length=255, blank=True, default='')

    branch_code = models.CharField(max_length=50, blank=True, default='')
    today = models.CharField(max_length=50, blank=True, default='')
    category = models.CharField(max_length=100, blank=True, default='')
    parent_gl = models.CharField(max_length=50, blank=True, default='')
    gl_code = models.CharField(max_length=50, blank=True, default='')
    descr = models.CharField(max_length=500, blank=True, default='')
    ccy = models.CharField(max_length=10, blank=True, default='')

    open_bal_lcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)
    dr_bal_lcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)
    cr_bal_lcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)
    close_bal_lcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)

    open_bal_fcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)
    dr_bal_fcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)
    cr_bal_fcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)
    close_bal_fcy = models.DecimalField(max_digits=18, decimal_places=2, default=0)

    gl_status = models.CharField(max_length=50, blank=True, default='')

    created_at = models.DateTimeField(auto_now_add=True)
    row_index = models.IntegerField(default=0)

    class Meta:
        db_table = 'trial_balance_entries'
        ordering = ['report_date', 'row_index']
        indexes = [
            models.Index(fields=['report_date', 'uploaded_by']),
            models.Index(fields=['upload', 'row_index']),
        ]

    def __str__(self):
        return f"{self.report_date} | {self.gl_code} | {self.descr[:40]}"


# ============================================
# EXCEPTION RECORD — now FK to ExceptionUpload
# ============================================

class ExceptionRecord(models.Model):
    """
    A single typed exception row from any exception report.

    The header metadata lives on `ExceptionUpload`; this row
    belongs to exactly one upload.
    """
    STATUS_CHOICES = [
        ('open', 'Open'),
        ('in_progress', 'In Progress'),
        ('closed', 'Closed'),
        ('resolved', 'Resolved'),
        ('pending', 'Pending'),
        ('rejected', 'Rejected'),
        ('overdue', 'Overdue'),
    ]

    upload = models.ForeignKey(
        ExceptionUpload,
        on_delete=models.CASCADE,
        related_name='exception_records',
        null=True, blank=True,
        help_text='The upload this row belongs to.',
    )

    serial_number = models.IntegerField(default=0)

    branch_unit = models.CharField(
        max_length=200, blank=True, default='', db_index=True
    )

    exception = models.TextField(blank=True, default='')

    date_noted = models.DateField(null=True, blank=True, db_index=True)
    date_noted_raw = models.CharField(max_length=50, blank=True, default='')

    target_closure_date = models.DateField(null=True, blank=True, db_index=True)
    target_closure_date_raw = models.CharField(max_length=50, blank=True, default='')

    category = models.CharField(max_length=200, blank=True, default='', db_index=True)

    responsible_officer = models.CharField(max_length=200, blank=True, default='', db_index=True)

    supervisor = models.CharField(max_length=200, blank=True, default='')

    auditee_response = models.TextField(blank=True, default='')

    remarks = models.TextField(blank=True, default='')

    status = models.CharField(
        max_length=50, choices=STATUS_CHOICES, default='open', db_index=True
    )
    status_raw = models.CharField(max_length=100, blank=True, default='')

    income_cost_saved = models.DecimalField(
        max_digits=18, decimal_places=2, default=0, db_index=True
    )
    income_cost_saved_raw = models.CharField(max_length=50, blank=True, default='')

    source_row_index = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'exception_records'
        ordering = ['upload', 'serial_number', 'source_row_index']
        indexes = [
            models.Index(fields=['upload', 'status']),
            models.Index(fields=['branch_unit', 'status']),
            models.Index(fields=['responsible_officer']),
            models.Index(fields=['target_closure_date']),
            models.Index(fields=['category']),
        ]

    def __str__(self):
        return f"#{self.serial_number} {self.branch_unit} — {self.exception[:50]}"

    @property
    def is_overdue(self):
        if self.status in ('closed', 'resolved'):
            return False
        if not self.target_closure_date:
            return False
        return self.target_closure_date < timezone.now().date()

    @property
    def days_until_target(self):
        if not self.target_closure_date:
            return None
        return (self.target_closure_date - timezone.now().date()).days


# ============================================
# SENT EMAIL AUDIT
# ============================================

class SentEmail(models.Model):
    """
    Audit record for every email sent from the compose page.
    ...
    """

    DELIVERY_STATUS = [
        ('sent', 'Sent'),
        ('failed', 'Failed'),
    ]

    # Who sent it
    sender = models.ForeignKey(
        'UserProfile',
        on_delete=models.SET_NULL,
        null=True,
        related_name='sent_emails',
    )

    # What they sent
    report_type = models.CharField(max_length=200, db_index=True)
    subject = models.CharField(max_length=255)
    body = models.TextField()

    # Comma-separated lists kept as text so we don't need a
    # join table for what is essentially a snapshot record.
    to_addresses = models.TextField(
        help_text='Comma-separated list of To: addresses at send time.',
    )
    cc_addresses = models.TextField(
        blank=True, default='',
        help_text='Comma-separated list of CC: addresses at send time.',
    )

    # Delivery bookkeeping
    status = models.CharField(
        max_length=20,
        choices=DELIVERY_STATUS,
        default='sent',
        db_index=True,
    )
    error_message = models.TextField(blank=True, default='')

    # ── SUPERVISOR SCORING ────────────────────────────────────
    # Points deducted by a supervisor from this submission's
    # final score. 0 = no deduction.
    manual_deduction = models.IntegerField(
        default=0,
        help_text='Points deducted by a supervisor (0–100).',
    )
    override_reason = models.TextField(
        blank=True, default='',
        help_text='Optional note explaining the deduction.',
    )
    scored_by = models.ForeignKey(
        'UserProfile',
        on_delete=models.SET_NULL,
        null=True, blank=True,
        related_name='scored_emails',
        help_text='Supervisor who last edited the deduction.',
    )
    scored_at = models.DateTimeField(
        null=True, blank=True,
        help_text='When the deduction was last edited.',
    )

    @property
    def final_score(self):
        return max(0, min(100, 100 - (self.manual_deduction or 0)))

    @property
    def badge_class(self):
        s = self.final_score
        if s >= 80: return 'success'
        if s >= 50: return 'warning'
        return 'danger'

    @property
    def is_overridden(self):
        return (self.manual_deduction or 0) > 0
    # ──────────────────────────────────────────────────────────

    # Timestamps
    sent_at = models.DateTimeField(auto_now_add=True, db_index=True)

    class Meta:
        db_table = 'sent_emails'
        ordering = ['-sent_at']
        indexes = [
            models.Index(fields=['sender', 'sent_at']),
            models.Index(fields=['report_type', 'sent_at']),
            models.Index(fields=['status', 'sent_at']),
        ]

    def __str__(self):
        return f'{self.subject} → {self.to_addresses[:60]} ({self.status})'
        

    # ---------- Helpers ----------
    @property
    def to_list(self):
        return [a.strip() for a in (self.to_addresses or '').split(',') if a.strip()]

    @property
    def cc_list(self):
        return [a.strip() for a in (self.cc_addresses or '').split(',') if a.strip()]

    @property
    def recipient_count(self):
        return len(self.to_list) + len(self.cc_list)

    # ---------- Scoring helpers ----------
    @property
    def final_score(self):
        """
        Auto score is 100 minus the manual deduction.
        Floored at 0, capped at 100.
        """
        return max(0, min(100, 100 - (self.manual_deduction or 0)))

    @property
    def badge_class(self):
        s = self.final_score
        if s >= 80: return 'success'
        if s >= 50: return 'warning'
        return 'danger'

    @property
    def is_overridden(self):
        return (self.manual_deduction or 0) > 0

# ============================================
# CHECKLIST CHANGE REQUESTS
# ============================================

class ChecklistChangeRequest(models.Model):
    """
    A member-initiated request to create, modify, delete, or extend
    a checklist. Supervisors approve or reject; on approval the
    payload is applied to the real Checklist/ChecklistTask rows.
    """

    class RequestType(models.TextChoices):
        CREATE   = 'create',   'Create new checklist'
        MODIFY   = 'modify',   'Modify existing checklist'
        DELETE   = 'delete',   'Delete checklist'
        ADD_TASK = 'add_task', 'Add task to checklist'

    class Status(models.TextChoices):
        PENDING   = 'pending',   'Pending review'
        APPROVED  = 'approved',  'Approved'
        REJECTED  = 'rejected',  'Rejected'
        CANCELLED = 'cancelled', 'Cancelled by requester'

    requester = models.ForeignKey(
        'UserProfile',
        on_delete=models.CASCADE,
        related_name='checklist_requests',
    )
    request_type = models.CharField(max_length=16, choices=RequestType.choices)
    status = models.CharField(
        max_length=12,
        choices=Status.choices,
        default=Status.PENDING,
        db_index=True,
    )

    target_checklist = models.ForeignKey(
        'Checklist',
        on_delete=models.CASCADE,
        null=True, blank=True,
        related_name='change_requests',
        help_text='Null for CREATE requests.',
    )

    # Snapshot of the checklist as it was when the request was made.
    # Lets a supervisor see the diff even if the checklist changed
    # in the meantime.
    current_snapshot = models.JSONField(default=dict, blank=True)

    # Proposed state. Shape depends on request_type:
    #   CREATE   → {name, description, frequency, tasks:[{description}],
    #               assigned_branches:[id], assigned_departments:[id]}
    #   MODIFY   → any subset of the above
    #   DELETE   → {} (empty)
    #   ADD_TASK → {description}
    payload = models.JSONField(default=dict)

    justification = models.TextField(blank=True)

    reviewed_by = models.ForeignKey(
        'UserProfile',
        on_delete=models.SET_NULL,
        null=True, blank=True,
        related_name='reviewed_checklist_requests',
    )
    reviewed_at = models.DateTimeField(null=True, blank=True)
    review_note = models.TextField(blank=True)

    # Populated after a CREATE request is approved.
    created_checklist = models.ForeignKey(
        'Checklist',
        on_delete=models.SET_NULL,
        null=True, blank=True,
        related_name='originating_request',
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'checklist_change_requests'
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['status', '-created_at']),
            models.Index(fields=['requester', '-created_at']),
            models.Index(fields=['target_checklist', 'status']),
        ]

    def __str__(self):
        return f'{self.get_request_type_display()} — {self.requester} ({self.status})'

    @property
    def is_pending(self):
        return self.status == self.Status.PENDING

    # ------------------------------------------------------------------
    # Serialization — used by the API views
    # ------------------------------------------------------------------
    def to_dict(self):
        return {
            'id': self.id,
            'request_type': self.request_type,
            'request_type_display': self.get_request_type_display(),
            'status': self.status,
            'status_display': self.get_status_display(),
            'requester': {
                'id': self.requester_id,
                'name': self.requester.full_name or self.requester.email,
                'email': self.requester.email,
                'avatar_url': self.requester.avatar_url,
            },
            'target_checklist': (
                {'id': self.target_checklist_id, 'name': self.target_checklist.name}
                if self.target_checklist else None
            ),
            'current_snapshot': self.current_snapshot,
            'payload': self.payload,
            'justification': self.justification,
            'reviewed_by': (
                {
                    'id': self.reviewed_by_id,
                    'name': self.reviewed_by.full_name or self.reviewed_by.email,
                }
                if self.reviewed_by else None
            ),
            'reviewed_at': self.reviewed_at.isoformat() if self.reviewed_at else None,
            'review_note': self.review_note,
            'created_checklist_id': self.created_checklist_id,
            'created_at': self.created_at.isoformat(),
            'updated_at': self.updated_at.isoformat(),
        }

    # ------------------------------------------------------------------
    # Review actions
    # ------------------------------------------------------------------
    def approve(self, supervisor, note=''):
        from django.db import transaction

        if self.status != self.Status.PENDING:
            raise ValueError(f'Cannot approve a {self.status} request.')

        with transaction.atomic():
            result = self._apply()
            self.status = self.Status.APPROVED
            self.reviewed_by = supervisor
            self.reviewed_at = timezone.now()
            self.review_note = note
            if result is not None and self.request_type == self.RequestType.CREATE:
                self.created_checklist = result
            self.save(update_fields=[
                'status', 'reviewed_by', 'reviewed_at',
                'review_note', 'created_checklist', 'updated_at',
            ])

        # Log to ActivityLog — fires after transaction commits
        ActivityLog.objects.create(
            user=supervisor,
            activity_type='score_updated',   # reuse existing choice
            details=(
                f'Approved {self.get_request_type_display()} request '
                f'#{self.id} from {self.requester.full_name}.'
                + (f' Note: {note}' if note else '')
            ),
        )

        return self

    def reject(self, supervisor, note=''):
        if self.status != self.Status.PENDING:
            raise ValueError(f'Cannot reject a {self.status} request.')

        self.status = self.Status.REJECTED
        self.reviewed_by = supervisor
        self.reviewed_at = timezone.now()
        self.review_note = note
        self.save(update_fields=[
            'status', 'reviewed_by', 'reviewed_at', 'review_note', 'updated_at',
        ])

        ActivityLog.objects.create(
            user=supervisor,
            activity_type='score_updated',
            details=(
                f'Rejected {self.get_request_type_display()} request '
                f'#{self.id} from {self.requester.full_name}.'
                + (f' Reason: {note}' if note else '')
            ),
        )

        return self

    def cancel(self):
        if self.status != self.Status.PENDING:
            raise ValueError('Only pending requests can be cancelled.')
        self.status = self.Status.CANCELLED
        self.save(update_fields=['status', 'updated_at'])
        return self

    # ------------------------------------------------------------------
    # Payload appliers
    # ------------------------------------------------------------------
    def _apply(self):
        handler = {
            self.RequestType.CREATE:   self._apply_create,
            self.RequestType.MODIFY:   self._apply_modify,
            self.RequestType.DELETE:   self._apply_delete,
            self.RequestType.ADD_TASK: self._apply_add_task,
        }.get(self.request_type)

        if not handler:
            raise ValueError(f'Unknown request_type: {self.request_type}')
        return handler()

    def _apply_create(self):
        from django.db import transaction

        p = self.payload
        with transaction.atomic():
            checklist = Checklist.objects.create(
                name=p['name'],
                description=p.get('description', ''),
                frequency=p['frequency'],
                assignment_target='specific',
                created_by=self.requester,
            )
            for i, t in enumerate(p.get('tasks', [])):
                ChecklistTask.objects.create(
                    checklist=checklist,
                    description=t['description'][:500],
                    order=i,
                )
            if p.get('assigned_branches'):
                checklist.assigned_branches.set(p['assigned_branches'])
            if p.get('assigned_departments'):
                checklist.assigned_departments.set(p['assigned_departments'])
        return checklist

    def _apply_modify(self):
        from django.db import transaction

        if not self.target_checklist:
            raise ValueError('Modify request has no target checklist.')

        p = self.payload
        c = self.target_checklist

        with transaction.atomic():
            if 'name' in p:        c.name = p['name']
            if 'description' in p: c.description = p['description']
            if 'frequency' in p:   c.frequency = p['frequency']
            c.save()

            if 'tasks' in p:
                c.tasks.all().delete()
                for i, t in enumerate(p['tasks']):
                    ChecklistTask.objects.create(
                        checklist=c,
                        description=t['description'][:500],
                        order=i,
                    )

            if 'assigned_branches' in p:
                c.assigned_branches.set(p['assigned_branches'])
            if 'assigned_departments' in p:
                c.assigned_departments.set(p['assigned_departments'])

        return c

    def _apply_delete(self):
        if not self.target_checklist:
            raise ValueError('Delete request has no target checklist.')
        c = self.target_checklist
        # Soft-delete so historical ChecklistLog rows remain meaningful.
        c.is_active = False
        c.save(update_fields=['is_active', 'updated_at'])
        return c

    def _apply_add_task(self):
        from django.db import transaction

        if not self.target_checklist:
            raise ValueError('Add-task request has no target checklist.')

        p = self.payload
        c = self.target_checklist
        with transaction.atomic():
            next_order = c.tasks.count()
            ChecklistTask.objects.create(
                checklist=c,
                description=p['description'][:500],
                order=next_order,
            )
        return c


# ============================================
# SIGNAL — snapshot checklist before Modify/Delete requests
# ============================================

def snapshot_checklist(checklist):
    """Return a serializable snapshot of a Checklist and its tasks."""
    return {
        'id': checklist.id,
        'name': checklist.name,
        'description': checklist.description,
        'frequency': checklist.frequency,
        'is_active': checklist.is_active,
        'assigned_branches': list(
            checklist.assigned_branches.values_list('id', flat=True)
        ),
        'assigned_departments': list(
            checklist.assigned_departments.values_list('id', flat=True)
        ),
        'tasks': [
            {'id': t.id, 'description': t.description, 'order': t.order}
            for t in checklist.tasks.all()
        ],
    }