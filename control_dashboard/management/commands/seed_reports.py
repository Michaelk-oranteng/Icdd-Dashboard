"""
Seed the Report table with the standard exception reporting schedule.

Usage
-----
    python manage.py seed_reports
    python manage.py seed_reports --creator admin@example.com
    python manage.py seed_reports --dry-run

Idempotent — any report_type that already exists (case-insensitive) is
skipped, so it is safe to re-run repeatedly. The two reports you have
already configured (Head Office Consolidated Exceptions Report and
Weekly Exceptions Report) are also explicitly skipped.
"""

import calendar
from datetime import date, datetime, timedelta

from django.core.management.base import BaseCommand
from django.utils import timezone

from control_dashboard.models import Report, UserProfile


# Reports already set up by the admin — skip them
SKIP_REPORT_TYPES = {
    'head office consolidated exceptions report',
    'weekly exceptions report',
}


# (report_type, frequency, deadline_time, deadline_rule, responsible, backup)
#
# deadline_rule:
#     ('dom', N)                -> N-th of each month
#     ('weekday', 0..6)         -> Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6
#     ('lwd',)                  -> last working day of the month
#     ('first_working_day',)    -> first working day of the month
REPORTS_TO_SEED = [
    ('EOM Presentation (PPT) on Work Done and Key Findings',
     'monthly', '16:45', ('lwd',), 'All', None),

    ('CMU-Takoradi GL Proof Review',
     'monthly', '14:00', ('dom', 10), 'Amanda Aba Bonney', 'Baffour Kyei Donkor'),

    ('CMU- Kumasi GL Proof Review',
     'monthly', '14:00', ('dom', 10), 'Baffour Kyei Donkor', 'Paul Dakora'),

    ('Investigation Schedule',
     'monthly', '14:00', ('dom', 25), 'Christiana Maku Blaboe', 'Godfred Aryee'),

    ('FINOPs- GL Proof review',
     'monthly', '14:00', ('dom', 10), 'Christiana Maku Blaboe', 'Kwabena Aboagye-Abrokwah'),

    ('Finance- GL Proof review',
     'monthly', '14:00', ('dom', 10), 'Christiana Maku Blaboe', 'Kwabena Aboagye-Abrokwah'),

    ('Debit Reversal',
     'monthly', '14:00', ('dom', 25), 'Christiana Maku Blaboe', 'Roberta Arhin-Andoh'),

    ('TROPs GL Proof Review',
     'monthly', '14:00', ('dom', 10), 'Daniel DelaliAtsuvi', 'Kwabena Aboagye-Abrokwah'),

    ('CMU -Accra GL Proof Review',
     'monthly', '14:00', ('dom', 10), 'Daniel Delali Atsuvi', 'Kwabena Aboagye-Abrokwah'),

    ('Head Office Workplan',
     'monthly', '14:00', ('dom', 25), 'Derrick Selorm Dzeku', 'Godfred Aryee'),

    ('Cash Imbalance Report',
     'monthly', '14:00', ('first_working_day',), 'Elvis Boafo', 'Baffour Kyei Donkor'),

    ('Branch Workplan',
     'monthly', '14:00', ('dom', 25), 'Frank Okyere Amoako', 'Naomi Ennin'),

    ('Recovery/Credit GL Proof Review',
     'monthly', '14:00', ('dom', 10), 'Godfred Aryee', 'Kwabena Aboagye-Abrokwah'),

    ('Cost Saved',
     'monthly', '14:00', ('dom', 25), 'Godfred Aryee', 'Peter Appiah Gyimah'),

    ('Trade GL Proof Review',
     'monthly', '14:00', ('dom', 10), 'Kwabena Aboagye-Abrokwah', 'Michael Akye Mensah'),

    ('E-Business GL Proof review',
     'monthly', '14:00', ('dom', 10), 'Michael Ofosu Koranteng', 'Michael Akye Mensah'),

    ('Report Score Sheet',
     'monthly', None, ('lwd',), 'Naana Odoso Akrofi-Addo', None),

    ('Branch GL Proof Report review Consolidation',
     'monthly', '14:00', ('dom', 10), 'Noella Afia Konadu Agyapong', 'Elvis Boafo'),

    ('Security Sweep Report',
     'monthly', '14:00', ('dom', 25), 'Paul Dakora', 'Habiba Tamimu'),

    ('Clearing GL Proof Review',
     'monthly', '14:00', ('dom', 10), 'Peter Appiah Gyimah', 'Daniel Delali Atsuvi'),

    ('Balance Confirmation',
     'monthly', '14:00', ('dom', 25), 'Peter Appiah Gyimah', 'Daniel Delali Atsuvi'),

    ('Branch Consolidated Exceptions Report',
     'weekly', '15:00', ('weekday', 3), 'Roberta Arhin-Andoh', 'Paul Dakora'),

    ('Statement Verification',
     'monthly', '14:00', ('dom', 25), 'Roberta Arhin-Andoh', 'Kwabena Aboagye-Abrokwah'),

    ('Cash Count Report',
     'monthly', '14:00', ('dom', 25), 'Sandra Nana Konadu Ofe', 'Noella Afia Konadu Agyapong'),

    ('Report on Customer Details wrongly Captured and Details not captured in CBA',
     'monthly', '14:00', ('dom', 25), 'Stella Adjetey', 'Ebenezer Tawiah'),
]


# ------------------------------------------------------------------
# Date helpers
# ------------------------------------------------------------------

def _last_working_day_of_month(d):
    """Last weekday (Mon-Fri) of the month containing `d`."""
    last = d.replace(day=calendar.monthrange(d.year, d.month)[1])
    while last.weekday() >= 5:
        last -= timedelta(days=1)
    return last


def _first_working_day_of_month(d):
    """First weekday (Mon-Fri) of the month containing `d`."""
    first = d.replace(day=1)
    while first.weekday() >= 5:
        first += timedelta(days=1)
    return first


def _next_day_of_month(target_day, from_date):
    """Next date >= from_date whose day == target_day."""
    try:
        candidate = from_date.replace(day=target_day)
        if candidate >= from_date:
            return candidate
    except ValueError:
        pass

    if from_date.month == 12:
        year, month = from_date.year + 1, 1
    else:
        year, month = from_date.year, from_date.month + 1
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, min(target_day, last_day))


def _next_weekday(target_weekday, from_date):
    """Next date >= from_date that falls on target_weekday (0=Mon..6=Sun)."""
    days_ahead = (target_weekday - from_date.weekday()) % 7
    return from_date + timedelta(days=days_ahead)


def _compute_anchor(rule, today):
    """Return a reasonable future anchor date for `rule` starting from today."""
    kind = rule[0]

    if kind == 'dom':
        return _next_day_of_month(rule[1], today)

    if kind == 'weekday':
        return _next_weekday(rule[1], today)

    if kind == 'lwd':
        candidate = _last_working_day_of_month(today)
        if candidate < today:
            # roll into next month
            if today.month == 12:
                nxt = today.replace(year=today.year + 1, month=1, day=1)
            else:
                nxt = today.replace(month=today.month + 1, day=1)
            candidate = _last_working_day_of_month(nxt)
        return candidate

    if kind == 'first_working_day':
        candidate = _first_working_day_of_month(today)
        if candidate < today:
            if today.month == 12:
                nxt = today.replace(year=today.year + 1, month=1, day=1)
            else:
                nxt = today.replace(month=today.month + 1, day=1)
            candidate = _first_working_day_of_month(nxt)
        return candidate

    return today


# ------------------------------------------------------------------
# User matching
# ------------------------------------------------------------------

def _normalize(s):
    """Lower-case, strip spaces / hyphens / periods."""
    return (
        str(s or '')
        .lower()
        .replace(' ', '')
        .replace('-', '')
        .replace('.', '')
        .replace(',', '')
    )


def _build_user_index():
    """Return {normalized_full_name: UserProfile}."""
    index = {}
    for u in UserProfile.objects.all():
        key = _normalize(u.full_name)
        if key and key not in index:
            index[key] = u
    return index


def _find_user(index, name):
    """Best-effort lookup by full name."""
    if not name:
        return None
    key = _normalize(name)
    if key in index:
        return index[key]

    # Fallback: substring match on normalized names
    for k, u in index.items():
        if key in k or k in key:
            return u
    return None


# ------------------------------------------------------------------
# Command
# ------------------------------------------------------------------

class Command(BaseCommand):
    help = (
        'Seed the Report table with the standard exception reporting schedule. '
        'Idempotent — existing reports (by name, case-insensitive) are skipped.'
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--creator',
            help='Email of the UserProfile to use as creator for the new reports. '
                 'Defaults to the first admin found.',
        )
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Preview what would be created without writing anything.',
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        creator_email = options.get('creator')

        # ------------------------------------------------------------------
        # Resolve creator
        # ------------------------------------------------------------------
        creator = None
        if creator_email:
            creator = UserProfile.objects.filter(email__iexact=creator_email).first()
            if not creator:
                self.stderr.write(self.style.ERROR(
                    f'No UserProfile found for --creator "{creator_email}".'
                ))
                return

        if not creator:
            creator = UserProfile.objects.filter(role='admin').order_by('id').first()

        if not creator:
            self.stderr.write(self.style.ERROR(
                'No admin UserProfile exists and no --creator was given. '
                'Create an admin user first, or pass --creator <email>.'
            ))
            return

        self.stdout.write(f'Using creator: {creator.full_name or creator.email}')

        today = timezone.now().date()
        user_index = _build_user_index()

        created_count = 0
        skipped_count = 0
        missing_officers = []

        for (name, frequency, deadline_time_str, rule, responsible, backup) in REPORTS_TO_SEED:

            # ---- Skip explicitly excluded reports ----
            if name.strip().lower() in SKIP_REPORT_TYPES:
                self.stdout.write(self.style.WARNING(
                    f'  SKIP  (already handled by admin): {name}'
                ))
                skipped_count += 1
                continue

            # ---- Skip if a report with this exact name already exists ----
            if Report.objects.filter(report_type__iexact=name).exists():
                self.stdout.write(self.style.WARNING(
                    f'  SKIP  (already exists): {name}'
                ))
                skipped_count += 1
                continue

            # ---- Compute anchor + deadline_time ----
            anchor_date = _compute_anchor(rule, today)
            deadline_time = None
            if deadline_time_str:
                try:
                    deadline_time = datetime.strptime(deadline_time_str, '%H:%M').time()
                except ValueError:
                    self.stderr.write(f'    ! Invalid time "{deadline_time_str}" for {name}')

            # ---- Resolve officers ----
            assigned_users = []
            is_assigned_to_all = False

            if responsible and responsible.strip().lower() == 'all':
                is_assigned_to_all = True
            else:
                primary = _find_user(user_index, responsible)
                if primary:
                    assigned_users.append(primary)
                elif responsible:
                    missing_officers.append((name, responsible))

            if backup:
                back_user = _find_user(user_index, backup)
                if back_user and back_user not in assigned_users:
                    assigned_users.append(back_user)
                elif backup:
                    missing_officers.append((name, backup))

            # ---- Preview / write ----
            preview = (
                f'  CREATE  {name}\n'
                f'          frequency={frequency}  anchor={anchor_date}  '
                f'time={deadline_time}  all={is_assigned_to_all}  '
                f'users={[u.full_name for u in assigned_users]}'
            )

            if dry_run:
                self.stdout.write(preview)
                created_count += 1
                continue

            report = Report.objects.create(
                report_type=name,
                frequency=frequency,
                description=f'Auto-seeded: {name}',
                deadline_date=anchor_date,
                deadline_time=deadline_time,
                is_assigned_to_all=is_assigned_to_all,
                created_by=creator,
                status='assigned',
            )

            if is_assigned_to_all:
                report.assigned_to.set(
                    UserProfile.objects.filter(role='member', status='active')
                )
            elif assigned_users:
                report.assigned_to.set(assigned_users)

            self.stdout.write(self.style.SUCCESS(preview))
            created_count += 1

        # ------------------------------------------------------------------
        # Summary
        # ------------------------------------------------------------------
        self.stdout.write('')
        self.stdout.write(self.style.SUCCESS(
            f'Done. Created {created_count}, skipped {skipped_count}.'
        ))

        if missing_officers:
            self.stdout.write('')
            self.stdout.write(self.style.WARNING(
                'Some officers were not found in UserProfile — the reports were still '
                'created, but no one is assigned to those slots. Add them from the '
                'Admin dashboard or create the missing user profiles:'
            ))
            for (rep, who) in missing_officers:
                self.stdout.write(f'  • {rep} → missing "{who}"')