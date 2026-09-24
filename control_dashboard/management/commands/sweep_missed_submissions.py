from django.core.management.base import BaseCommand
from django.utils import timezone
from datetime import datetime, timedelta

from control_dashboard.models import (
    ReportSchedule,
    SubmittedReportScore,
)


class Command(BaseCommand):
    help = "Record score-0 rows for schedules whose period closed without a submission."

    def handle(self, *args, **options):
        now = timezone.now()
        swept = 0

        for schedule in ReportSchedule.objects.filter(is_active=True):
            # If the due date already passed
            if not schedule.next_due_date or schedule.next_due_date > now.date():
                continue

            report = schedule.report

            # Was there a score row submitted on/after the due date?
            already = SubmittedReportScore.objects.filter(
                report=report,
                sent_at__date__gte=schedule.next_due_date,
            ).exists()

            if already:
                continue

            # Check if a zero-score placeholder already exists for this period.
            placeholder = SubmittedReportScore.objects.filter(
                report=report,
                sent_at__isnull=True,
                created_at__date=schedule.next_due_date,
            ).exists()

            if placeholder:
                continue

            # Create a "missed" placeholder → score 0.
            SubmittedReportScore.objects.create(
                report=report,
                sent_at=None,
                had_attachment=False,
                deadline_at=timezone.make_aware(
                    datetime.combine(
                        schedule.next_due_date,
                        schedule.due_time or datetime.strptime('23:59', '%H:%M').time(),
                    )
                ),
            )
            swept += 1

        self.stdout.write(self.style.SUCCESS(f"Swept {swept} missed submission(s)."))