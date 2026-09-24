# control_dashboard/models.py

from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from django.db.models.signals import pre_save, post_save
from django.dispatch import receiver
import re
from datetime import date, timedelta


# ============================================================
# CONSTANTS — single source of truth for status values
# ============================================================

# The Report.status value that marks a Report row as an
# Excel data container (created by draft.html). These are
# NOT submissions — they are raw exception data.
UPLOADED_STATUS = 'uploaded'


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
        user = User.objects.create_user(
            username=self.username,
            email=self.email,
            password=password or 'defaultpassword123'
        )
        
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
    
    class Meta:
        db_table = 'user_profiles'
        ordering = ['full_name']


# ============================================
# BRANCH AND DEPARTMENT MODELS
# ============================================

class Branch(models.Model):
    """
    Model for storing bank branches.

    Also carries a static BRANCH_CODE_MAP that maps the codes found
    in the trial balance BRANCH_CODE column (e.g. '001', '101') to
    human-readable branch names.
    """

    # ============================================
    # BRANCH CODE → NAME MAPPING
    # Codes come from the trial balance BRANCH_CODE column.
    # Edit this dictionary to add/rename branches.
    # ============================================
    BRANCH_CODE_MAP = {
        '001': 'CMU - ACCRA',
        '000': 'HEAD OFFICE',
        '101': 'AIRPORT',
        '102': 'ABOSSEY OKAI',
        '103': 'ASHIAMAN',
        '104': 'ABELENKPE',
        '105': 'DOME',
        '107': 'TEMA HARBOUR',
        '108': 'KOKOMLEMLE',
        '109': 'MADINA MARKET',
        '110': 'TUDU',
        '111': 'SPINTEX BASKET',
        '112': 'OSU',
        '113': 'WEIJA',
        '114': 'EAST LEGON',
        '115': 'TEMA COMM 1',
        '117': 'DANSOMAN',
        '118': 'ODORKOR',
        '119': 'ASHALEY BOTWE',
        '120': 'ADABRAKA',
        '121': 'ACCRA CENTRAL',
        '122': 'SPINTEX MANET',
        '123': 'NIMA',
        '124': 'NIA',
        '125': 'MADINA ESTATE',
        '126': 'ACHIMOTA',
        '127': 'TEMA COMM 11',
        '201': 'KOFORIDUA',
        '301': 'KASOA',
        '401': 'TAKORADI MARKET',
        '402': 'TARKWA',
        '403': 'TAKORADI LIBERATION',
        '404': 'TAKORADI - CMU',
        '601': 'MANHYIA',
        '602': 'ADUM PREMPEH',
        '603': 'KEJETIA',
        '604': 'ANLOGA',
        '605': 'KRONUM',
        '606': 'ADUM ADDO KUFUOR',
        '607': 'AHODWO',
        '608': 'KUMASI - CMU',
        '609': 'KNUST',
        '701': 'TECHIMAN',
        '702': 'SUNYANI',
        '801': 'TAMALE',
    }

    # Pre-built list of (code, display) tuples for dropdowns
    BRANCH_CODE_CHOICES = [
        (code, f"{code} — {name}")
        for code, name in sorted(BRANCH_CODE_MAP.items())
    ]

    name = models.CharField(max_length=100, unique=True)
    code = models.CharField(max_length=20, unique=True, blank=True, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

    # ============================================
    # STATIC HELPERS — resolve a raw trial-balance
    # code (e.g. '001') to a human-readable name.
    # ============================================
    @classmethod
    def code_to_name(cls, code):
        """Return the display name for a given branch code, or the code itself if unknown."""
        if code is None:
            return ''
        key = str(code).strip()
        return cls.BRANCH_CODE_MAP.get(key, key)

    @classmethod
    def code_to_display(cls, code):
        """Return 'CODE — Name' for a given code, or just the code if unknown."""
        if code is None:
            return ''
        key = str(code).strip()
        name = cls.BRANCH_CODE_MAP.get(key)
        return f"{key} — {name}" if name else key

    @classmethod
    def all_code_choices(cls):
        """Return a fresh list of (code, 'CODE — Name') tuples for dropdowns."""
        return [
            (code, f"{code} — {name}")
            for code, name in sorted(cls.BRANCH_CODE_MAP.items())
        ]

    class Meta:
        db_table = 'branches'
        ordering = ['name']
        verbose_name_plural = 'Branches'


class Department(models.Model):
    """Model for storing bank departments."""
    name = models.CharField(max_length=100, unique=True)
    code = models.CharField(max_length=20, unique=True, blank=True, null=True)
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
def auto_generate_username(sender, instance, **kwargs):
    """
    Automatically generate username before saving.
    """
    if not instance.username and instance.email:
        instance.username = instance.generate_username_from_email()


@receiver(post_save, sender=UserProfile)
def create_user_for_profile(sender, instance, created, **kwargs):
    """
    Automatically create Django User when UserProfile is created.
    """
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

    ═══════════════════════════════════════════════════════════════
    TWO DISTINCT KINDS OF REPORT ROWS EXIST IN THIS TABLE
    ═══════════════════════════════════════════════════════════════

    1. ADMIN-CREATED REPORTS (status: assigned / in_progress /
       submitted / completed / approved / rejected / draft)

       Created by admin via report_creation.html.
       These have a frequency and a deadline_date anchor.
       The member submits them via submit.html → each send
       appends a SubmittedReportScore row.

    2. EXCEL DATA CONTAINERS (status: 'uploaded')

       Created by draft.html → api_save_imported_data.
       These hold TYPED exception rows (ExceptionRecord).
       They are NOT submissions. They exist only to feed
       Analytics and Logged Exceptions.

       Headers for the upload are stored on `excel_headers` as a
       comma-separated string, so the UI can render the correct
       columns without needing a separate "import" table.

    Helper properties below make the distinction explicit.
    ═══════════════════════════════════════════════════════════════
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
        ('uploaded', 'Uploaded (Exception Data)'),   # ← Excel data container
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

    # Comma-separated list of the original Excel column headers for
    # this upload. Used by get_display_data() to render the table.
    # Only populated when status='uploaded'.
    excel_headers = models.TextField(blank=True, default='')
    
    def __str__(self):
        return self.report_type
    
    def get_frequency_display(self):
        return dict(self.FREQUENCY_CHOICES).get(self.frequency, self.frequency)
    
    def get_status_display(self):
        return dict(self.STATUS_CHOICES).get(self.status, self.status)
    
    # ============================================================
    # CLASSIFICATION HELPERS
    # ============================================================
    
    @property
    def is_exception_upload(self):
        """
        True if this Report is an Excel data container.
        Excel uploads are created by draft.html and are NOT submissions.
        """
        return self.status == UPLOADED_STATUS
    
    @property
    def is_submitted(self):
        """
        True if this Report has been submitted via email.
        Note: this checks the Report.status, not individual submissions.
        """
        return self.status in SUBMITTED_STATUSES
    
    def get_excel_headers_list(self):
        """Return excel_headers as a clean list of strings."""
        if not self.excel_headers:
            return []
        return [h.strip() for h in self.excel_headers.split(',') if h.strip()]
    
    class Meta:
        db_table = 'reports'
        ordering = ['-created_at']
    
    def get_display_data(self):
        """
        Return {headers, rows, is_excel, row_count} for the template.

        Source of truth:
          • Excel containers (status='uploaded') → ExceptionRecord rows
          • Form-based reports → ReportDataField rows
        """
        display_data = {
            'rows': [],
            'headers': [],
            'is_excel': False,
            'row_count': 0,
        }

        # ---------- Excel container: read from ExceptionRecord ----------
        if self.status == UPLOADED_STATUS:
            headers = self.get_excel_headers_list()

            if not headers:
                # Fall back to canonical headers if somehow not stored
                headers = [
                    'S/N', 'BRANCH/UNIT', 'EXCEPTION',
                    'DATE EXCEPTION WAS NOTED', 'TARGET DATE FOR CLOSURE',
                    'CATEGORY OF EXCEPTION', 'RESPONSIBLE OFFICER',
                    'SUPERVISOR', "AUDITEE'S RESPONSE", 'REMARKS',
                    'STATUS', 'INCOME/COST SAVED',
                ]

            display_data['is_excel'] = True
            display_data['headers'] = headers

            for rec in self.exception_records.order_by('source_row_index'):
                full_row = {
                    'row_id': rec.id,   # ← ExceptionRecord.id used by edit/delete
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
                # Project only the columns this upload actually had
                row_view = {h: full_row.get(h, '') for h in headers}
                row_view['row_id'] = rec.id
                display_data['rows'].append(row_view)

            display_data['row_count'] = len(display_data['rows'])
            return display_data

        # ---------- Form-based report: read from ReportDataField ----------
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

        # ---------- Nothing found ----------
        display_data['headers'] = ['Branch/Unit', 'Date', 'Observation', 'Responsible Staff', 'Status']
        display_data['row_count'] = 0
        return display_data


class ReportDataField(models.Model):
    """
    Model for storing report data fields (replaces JSON).
    """
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


# ============================================
# REPORT SUBMISSION MODELS
# ============================================

class ReportSubmission(models.Model):
    """
    Model for storing member report submissions.
    """
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
    """
    Model for storing submission form data (replaces JSON).
    """
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
    """
    Model for storing checklists/activities.
    """
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
        """Get a display string for assigned users."""
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
        """Return the number of days between each occurrence based on frequency."""
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
        """
        Calculate expected occurrences for a date range based on frequency.
        For daily frequency, only counts weekdays (Monday-Friday).
        For monthly frequency, counts months.
        For weekly frequency, counts weeks.
        """
        from datetime import timedelta
        
        if not start_date:
            start_date = timezone.now().date().replace(month=1, day=1)
        if not end_date:
            end_date = timezone.now().date()
        
        # If end_date is before start_date, swap them
        if end_date < start_date:
            start_date, end_date = end_date, start_date
        
        # For annual frequency, count the number of years
        if self.frequency == 'annual':
            years = end_date.year - start_date.year + 1
            return years
        
        # For daily frequency, count only weekdays (Monday to Friday)
        if self.frequency == 'daily':
            current = start_date
            count = 0
            while current <= end_date:
                # Monday = 0, Sunday = 6, so weekdays are 0-4
                if current.weekday() < 5:  # Monday to Friday
                    count += 1
                current += timedelta(days=1)
            return count
        
        # For monthly frequency, count the number of months in the range
        if self.frequency == 'monthly':
            months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) + 1
            return months
        
        # For weekly frequency, count the number of weeks in the range
        if self.frequency == 'weekly':
            days_diff = (end_date - start_date).days + 1
            weeks = days_diff // 7
            return max(1, weeks)
        
        # For quarterly frequency, count the number of quarters
        if self.frequency == 'quarterly':
            days_diff = (end_date - start_date).days + 1
            quarters = days_diff // 91
            return max(1, quarters)
        
        # For bi-annual frequency, count the number of half-years
        if self.frequency == 'bi-annual':
            days_diff = (end_date - start_date).days + 1
            half_years = days_diff // 182
            return max(1, half_years)
        
        # For other frequencies, calculate based on days difference
        days_diff = (end_date - start_date).days + 1
        frequency_days = self.get_frequency_days()
        return max(1, days_diff // frequency_days)
    
    def get_completion_percentage(self, user_profile, start_date=None, end_date=None):
        """
        Calculate completion percentage for a user based on frequency.
        Capped at 100%.
        """
        if not start_date:
            start_date = timezone.now().date().replace(month=1, day=1)
        if not end_date:
            end_date = timezone.now().date()
        
        # If end_date is before start_date, swap them
        if end_date < start_date:
            start_date, end_date = end_date, start_date
        
        expected = self.get_expected_occurrences(start_date, end_date)
        if expected == 0:
            return 0
        
        # Get actual logs
        actual = ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile,
            log_date__gte=start_date,
            log_date__lte=end_date
        ).count()
        
        # Calculate percentage and cap at 100%
        percentage = int((actual / expected) * 100)
        return min(percentage, 100)
    
    def get_monthly_completion(self, user_profile, month=None, year=None):
        """Calculate completion percentage for a specific month."""
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
        """Calculate year-to-date completion percentage."""
        today = timezone.now().date()
        start_date = date(today.year, 1, 1)
        return self.get_completion_percentage(user_profile, start_date, today)
    
    def get_monthly_expected(self, user_profile, month=None, year=None):
        """Get the expected number of occurrences for a specific month."""
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
        """Get the actual number of completions for a specific month."""
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
        """Get the expected number of occurrences for year-to-date."""
        today = timezone.now().date()
        start_date = date(today.year, 1, 1)
        return self.get_expected_occurrences(start_date, today)
    
    def get_year_to_date_actual(self, user_profile):
        """Get the actual number of completions for year-to-date."""
        today = timezone.now().date()
        start_date = date(today.year, 1, 1)
        return ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile,
            log_date__gte=start_date,
            log_date__lte=today
        ).count()
    
    def get_next_due_date(self, user_profile):
        """
        Calculate the next due date based on frequency and last completion.
        For daily, skips weekends.
        """
        from datetime import timedelta
        
        if self.frequency == 'one-off' or self.frequency == 'annual':
            return None
        
        # Get the last log for this checklist and user
        last_log = ChecklistLog.objects.filter(
            checklist=self,
            user=user_profile
        ).order_by('-log_date').first()
        
        if not last_log:
            # If no logs, next due is today (or next weekday for daily)
            today = timezone.now().date()
            next_date = today
            if self.frequency == 'daily':
                while next_date.weekday() >= 5:  # Skip weekends
                    next_date += timedelta(days=1)
            return next_date
        
        # Calculate next due date based on frequency
        if self.frequency == 'daily':
            next_date = last_log.log_date + timedelta(days=1)
            # Skip weekends
            while next_date.weekday() >= 5:
                next_date += timedelta(days=1)
        elif self.frequency == 'weekly':
            next_date = last_log.log_date + timedelta(days=7)
        elif self.frequency == 'monthly':
            # Add one month
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
        
        # If next date is in the past, set to today (or next weekday)
        if next_date < timezone.now().date():
            next_date = timezone.now().date()
            if self.frequency == 'daily':
                while next_date.weekday() >= 5:
                    next_date += timedelta(days=1)
        
        return next_date
    
    def get_actual_completion_count(self, user_profile, start_date=None, end_date=None):
        """Get the actual number of completed occurrences."""
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
        """Calculate the next due date based on frequency."""
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
        
        # For daily, skip weekends if needed
        if self.frequency == 'daily':
            while next_date.weekday() >= 5:  # Saturday or Sunday
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
    """
    Model for storing checklist tasks.
    """
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
        # Each (checklist, user, unit, date) is unique
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
    """
    Model to track user activities across the application.
    """
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
    details = models.TextField(
        blank=True,
        null=True,
        help_text="Detailed description of the activity"
    )
    ip_address = models.GenericIPAddressField(
        blank=True,
        null=True,
        help_text="IP address of the user"
    )
    user_agent = models.TextField(
        blank=True,
        null=True,
        help_text="User agent string"
    )
    created_at = models.DateTimeField(
        auto_now_add=True
    )
    
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
        """Get the FontAwesome icon class for this activity type."""
        return self.ACTIVITY_ICONS.get(self.activity_type, 'fa-circle')


class AdHocDeduction(models.Model):
    """
    Ad-hoc scoring event for a team member.
    Each row can record points DEDUCTED, points ADDED, or both.
    """
    user = models.ForeignKey(
        'UserProfile',
        on_delete=models.CASCADE,
        related_name='ad_hoc_deductions',
        help_text="The team member this score entry applies to"
    )
    task_description = models.CharField(
        max_length=255,
        help_text="Description of the task or report"
    )
    points = models.IntegerField(
        default=0,
        help_text="Number of points DEDUCTED (0-100)"
    )
    points_added = models.IntegerField(
        default=0,
        help_text="Number of points ADDED/AWARDED (0-100)"
    )
    reason = models.TextField(
        blank=True,
        null=True,
        help_text="Reason for the score change"
    )
    created_by = models.ForeignKey(
        'UserProfile',
        on_delete=models.SET_NULL,
        null=True,
        related_name='created_deductions',
        help_text="Supervisor who created the entry"
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
        """Human-readable single-string summary."""
        parts = []
        if self.points:
            parts.append(f"-{self.points}%")
        if self.points_added:
            parts.append(f"+{self.points_added}%")
        return " ".join(parts) if parts else "0%"

    def get_badge_class(self):
        """Badge class driven by whichever value is set."""
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


class TrialBalanceEntry(models.Model):
    """
    One row of an uploaded Daily Trial Balance file.
    Each upload creates many entries, all sharing the same report_date.
    """
    # Link back to the parent report (optional — useful for grouping)
    report = models.ForeignKey(
        Report,
        on_delete=models.CASCADE,
        related_name='trial_balance_entries',
        null=True, blank=True
    )
    
    # Who uploaded it
    uploaded_by = models.ForeignKey(
        UserProfile,
        on_delete=models.CASCADE,
        related_name='trial_balance_entries'
    )
    
    # The reporting date (user-selected)
    report_date = models.DateField(db_index=True)
    
    # Original file
    file_name = models.CharField(max_length=255, blank=True, default='')
    
    # ============================================
    # The 16 columns from the Excel file
    # ============================================
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
    
    # Meta
    created_at = models.DateTimeField(auto_now_add=True)
    row_index = models.IntegerField(default=0)  # preserves original file order
    
    class Meta:
        ordering = ['report_date', 'row_index']
        indexes = [
            models.Index(fields=['report_date', 'uploaded_by']),
        ]
    
    def __str__(self):
        return f"{self.report_date} | {self.gl_code} | {self.descr[:40]}"


# ============================================================
# SUBMITTED REPORT SCORE
# ============================================================
# One row per email submission. This is the ONLY model that
# represents "a report was actually submitted". Excel uploads
# never create rows here.
# ============================================================


class ExceptionRecord(models.Model):
    """
    A single typed exception row from any exception report.

    This is the CANONICAL, queryable representation of an exception.
    It maps 1:1 to the headers used across all exception reports:

        S/N
        BRANCH/UNIT
        EXCEPTION
        DATE EXCEPTION WAS NOTED
        TARGET DATE FOR CLOSURE
        CATEGORY OF EXCEPTION
        RESPONSIBLE OFFICER
        SUPERVISOR
        AUDITEE'S RESPONSE
        REMARKS
        STATUS
        INCOME/COST SAVED

    Each ExceptionRecord belongs DIRECTLY to a Report (the Excel data
    container). There is no intermediate import or batch layer — the
    Report itself carries the original headers via `excel_headers`.
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

    # ------------------------------------------------------------
    # Parent linkage
    # ------------------------------------------------------------
    report = models.ForeignKey(
        'Report',
        on_delete=models.CASCADE,
        related_name='exception_records',
        null=True, blank=True,
        help_text="The Report (Excel data container) this record belongs to",
    )

    # ------------------------------------------------------------
    # Column: S/N
    # ------------------------------------------------------------
    serial_number = models.IntegerField(
        default=0,
        help_text="S/N from the original report",
    )

    # ------------------------------------------------------------
    # Column: BRANCH/UNIT
    # ------------------------------------------------------------
    branch_unit = models.CharField(
        max_length=200,
        blank=True,
        default='',
        db_index=True,
        help_text="Branch or unit the exception belongs to",
    )

    # ------------------------------------------------------------
    # Column: EXCEPTION
    # ------------------------------------------------------------
    exception = models.TextField(
        blank=True,
        default='',
        help_text="Description of the exception / finding",
    )

    # ------------------------------------------------------------
    # Column: DATE EXCEPTION WAS NOTED
    # ------------------------------------------------------------
    date_noted = models.DateField(
        null=True,
        blank=True,
        db_index=True,
        help_text="Date the exception was first observed",
    )
    date_noted_raw = models.CharField(
        max_length=50,
        blank=True,
        default='',
        help_text="Original string, in case parsing failed",
    )

    # ------------------------------------------------------------
    # Column: TARGET DATE FOR CLOSURE
    # ------------------------------------------------------------
    target_closure_date = models.DateField(
        null=True,
        blank=True,
        db_index=True,
        help_text="Date by which the exception should be closed",
    )
    target_closure_date_raw = models.CharField(
        max_length=50,
        blank=True,
        default='',
        help_text="Original string, in case parsing failed",
    )

    # ------------------------------------------------------------
    # Column: CATEGORY OF EXCEPTION
    # ------------------------------------------------------------
    category = models.CharField(
        max_length=200,
        blank=True,
        default='',
        db_index=True,
        help_text="Category / classification of the exception",
    )

    # ------------------------------------------------------------
    # Column: RESPONSIBLE OFFICER
    # ------------------------------------------------------------
    responsible_officer = models.CharField(
        max_length=200,
        blank=True,
        default='',
        db_index=True,
        help_text="Officer responsible for resolving the exception",
    )

    # ------------------------------------------------------------
    # Column: SUPERVISOR
    # ------------------------------------------------------------
    supervisor = models.CharField(
        max_length=200,
        blank=True,
        default='',
        help_text="Supervisor overseeing the exception",
    )

    # ------------------------------------------------------------
    # Column: AUDITEE'S RESPONSE
    # ------------------------------------------------------------
    auditee_response = models.TextField(
        blank=True,
        default='',
        help_text="Response from the auditee / responsible party",
    )

    # ------------------------------------------------------------
    # Column: REMARKS
    # ------------------------------------------------------------
    remarks = models.TextField(
        blank=True,
        default='',
        help_text="Additional remarks",
    )

    # ------------------------------------------------------------
    # Column: STATUS
    # ------------------------------------------------------------
    status = models.CharField(
        max_length=50,
        choices=STATUS_CHOICES,
        default='open',
        db_index=True,
        help_text="Current status of the exception",
    )
    status_raw = models.CharField(
        max_length=100,
        blank=True,
        default='',
        help_text="Original status string as it appeared in the file",
    )

    # ------------------------------------------------------------
    # Column: INCOME/COST SAVED
    # ------------------------------------------------------------
    income_cost_saved = models.DecimalField(
        max_digits=18,
        decimal_places=2,
        default=0,
        db_index=True,
        help_text="Monetary value of income recovered or cost saved",
    )
    income_cost_saved_raw = models.CharField(
        max_length=50,
        blank=True,
        default='',
        help_text="Original string, in case parsing failed",
    )

    # ------------------------------------------------------------
    # Bookkeeping
    # ------------------------------------------------------------
    source_row_index = models.IntegerField(
        default=0,
        help_text="Row index in the original Excel file (for traceability)",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'exception_records'
        ordering = ['report', 'serial_number', 'source_row_index']
        indexes = [
            models.Index(fields=['report', 'status']),
            models.Index(fields=['branch_unit', 'status']),
            models.Index(fields=['responsible_officer']),
            models.Index(fields=['target_closure_date']),
            models.Index(fields=['category']),
        ]

    def __str__(self):
        return f"#{self.serial_number} {self.branch_unit} — {self.exception[:50]}"

    # ------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------
    @property
    def is_overdue(self):
        """True if the target closure date has passed and the exception isn't closed."""
        if self.status in ('closed', 'resolved'):
            return False
        if not self.target_closure_date:
            return False
        return self.target_closure_date < timezone.now().date()

    @property
    def days_until_target(self):
        """Days remaining until target closure date. Negative if overdue."""
        if not self.target_closure_date:
            return None
        return (self.target_closure_date - timezone.now().date()).days