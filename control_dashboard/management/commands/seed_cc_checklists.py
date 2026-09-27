"""
Seed the Cluster Control (CC) activity checklists.

Creates one Checklist per activity from the Operations Control
checklist, attaches its tasks, and assigns each one to every
active user with position='cc' (Cluster Control).

Idempotent: re-running updates existing checklists instead of
duplicating them.

Usage:
    python manage.py seed_cc_checklists              # create / update
    python manage.py seed_cc_checklists --dry-run    # preview only
    python manage.py seed_cc_checklists --reset      # delete all CC-seeded then recreate
"""

from django.core.management.base import BaseCommand
from django.db import transaction
from django.utils import timezone

from control_dashboard.models import Checklist, ChecklistTask, UserProfile


# ============================================================
# MASTER DATA — every activity, its frequency, and its tasks
# ============================================================
CHECKLIST_DATA = [
    {
        'activity': 'Cash count - Vault and ATM',
        'frequency': 'weekly',
        'tasks': [
            'Count the physical cash at vault and compare physical cash with system balance and vault summary sheet.',
            'Investigate if there is an overage or shortage.',
            'Request for evidence of approval if branch vault limit is exceeded.',
            'For ATM, count and compare physical cash in the cassette with system balance and ATM counters.',
            'Request for evidence of approval if branch ATM limit is exceeded for week day, weekend and holidays.',
        ],
    },
    {
        'activity': 'Vault Management',
        'frequency': 'weekly',
        'tasks': [
            'Confirm and observe the use of dual control and appropriate use of key and combination by authorized persons.',
            'Check for the presence of foreign materials in the vault. Non-cash related items must be recorded in a register with the details.',
            'Review of Vault attendance register.',
            'Ensure vault is well illuminated.',
            'Inspect for cleanliness.',
        ],
    },
    {
        'activity': 'Key Register',
        'frequency': 'weekly',
        'tasks': [
            'Confirm if the names of custodians and their backups as well as the keys in their possession are documented.',
            'Confirm if the details in the register have been updated with the current key holders.',
        ],
    },
    {
        'activity': 'Tellers Call Over',
        'frequency': 'daily',
        'tasks': [
            'Review all vouchers processed to detect and correct errors.',
            'Confirm if payment instruments have been verified and the right stamps (paid, received and signature verification) are appended as well as the signature of the processing officer on the stamp.',
            'Check if cheque is postdated/outdated, amount in words and figures are the same and alterations signed against.',
            'Confirm if cheque payment above GHS5,000 and deposit above GHS150,000 has been signed by the BSS as evidence of review.',
            'Confirm if mobile money withdrawal above GHS5,000 and mobile money deposit above GHS100,000 has been signed by the BSS as evidence of review.',
            'Confirm indemnity in the CBA for transactional request sent via email, SMS and WhatsApp.',
            'Confirm if transactions are posted into the correct account.',
        ],
    },
    {
        'activity': 'ATM Management',
        'frequency': 'weekly',
        'tasks': [
            'Use of dual control and appropriate use of key and combination by authorized persons.',
            'Inspect for cleanliness.',
            'Review of ATM attendance register.',
            'Ensure room is well illuminated.',
        ],
    },
    {
        'activity': 'Vault Summary Sheet',
        'frequency': 'weekly',
        'tasks': [
            'Review the vault summary sheet to confirm if the denominations match with the denominations in retail and physical cash. Investigate if there is any disparity.',
            'Review vault GL to confirm balances.',
            'Review Cash In Transit form.',
            'Confirm if the sheet has been duly completed and signed by the BSS and BM.',
            'Sheets should be filed neatly on daily basis.',
        ],
    },
    {
        'activity': 'Cash Open and Cash Close Slips',
        'frequency': 'weekly',
        'tasks': [
            'Review to confirm amount received/sent to vault tallies with the information on the cash open/close slips.',
            'Review to confirm if all amount received/sent to vault per transaction during the day has a supporting document (cash open/close slips) to reflect the same amount in the GL.',
            'Review the cash open and close slip to confirm if the amount in figures matches with the amount in words.',
            'Confirm if the sheets have been duly completed and signed by the Teller and BSS.',
        ],
    },
    {
        'activity': 'ATM Reconciliation Reports',
        'frequency': 'weekly',
        'tasks': [
            'Confirm if ATM reconciliation reports are prepared daily and filed.',
            'Confirm if all outstanding differences are identified and the accuracy of the transactions captured as difference in the report by reviewing the journal/ATM GL.',
            'Confirm if cash taken from the cassette is credited to the tellers till before replenishment.',
        ],
    },
    {
        'activity': 'Account Open',
        'frequency': 'daily',
        'tasks': [
            'Request for account opening packs or documentation.',
            'Review for completeness. Ensure all information captured in the account pack have been accurately captured in the CBA.',
            'Confirm the availability of required documentation (refer to checklist at the back of account pack) and any other regulatory requirement as per the nature of business.',
            'Confirm initial deposit per tariff guide (Request for approval to open zero balance account in exceptional circumstances).',
            'Confirm if all products/services selected by the customer have been requested.',
        ],
    },
    {
        'activity': 'Request for Statement',
        'frequency': 'weekly',
        'tasks': [
            'Evidence of request on file.',
            'Completion and review of forms (application of all necessary stamps, date, customer signature, number of pages).',
            'Availability of funds to facilitate a successful debit of customer\'s account.',
            'Review customer\'s account and commission GL to confirm if applicable charges have been applied.',
            'Customer request is recorded in register and signed by all parties.',
            'Authority note and Ghana card verification should be sighted on file for statement received by third party.',
            'CCO should ensure that all statement requests via mail are recorded in the register.',
        ],
    },
    {
        'activity': 'Request for Counter Cheque',
        'frequency': 'weekly',
        'tasks': [
            'Evidence of request on file and cheque book.',
            'Completion and review of forms (application of all necessary stamps).',
            'Check if counter cheque is serially issued.',
            'Destroyed counter cheque should be cancelled and recorded in the register.',
            'Review customer\'s account and commission GL to confirm if applicable charges have been applied.',
            'Customer request is recorded in register and signed by all parties.',
        ],
    },
    {
        'activity': 'Request for Cheque Book',
        'frequency': 'weekly',
        'tasks': [
            'Evidence of request on file and receipts of orders.',
            'Completion and review of forms (application of all necessary stamps).',
            'Review customer\'s account and commission GL to confirm if applicable charges have been applied.',
        ],
    },
    {
        'activity': 'Request for ATM Card',
        'frequency': 'weekly',
        'tasks': [
            'Request for ATM request file and acknowledgement receipt.',
            'Completion and review of forms (application of all necessary stamps).',
            'Review customer\'s account and commission GL to confirm if applicable charges have been applied.',
            'Customer request is recorded in register and signed by all parties.',
        ],
    },
    {
        'activity': 'Dormant Account Activation',
        'frequency': 'weekly',
        'tasks': [
            'Evidence of reactivation form on file.',
            'Confirm the capturing of updated information in the CBA.',
            'Review customer account to confirm deposit for the reactivation.',
            'Confirm if customer details have been captured in the register.',
        ],
    },
    {
        'activity': 'Stop Cheque Requests',
        'frequency': 'monthly',
        'tasks': [
            'Evidence of request on file.',
            'Confirm if cheque has been stopped in the CBA and applicable charge taken.',
            'Review customer\'s account and commission GL to confirm if applicable charges have been applied.',
            'Customer request is recorded in register and signed by all parties.',
        ],
    },
    {
        'activity': 'Account Closure',
        'frequency': 'monthly',
        'tasks': [
            'Evidence of request and approval.',
            'Retrieval of cheque book.',
            'Ensure that unused cheque numbers are cancelled.',
            'Confirm if account has been closed and all cheque book and cards retrieved.',
            'Confirm account of all deceased customers with Letters of Administration on file are closed.',
            'Customer details should be captured in the register.',
        ],
    },
    {
        'activity': 'Cheque Book Review',
        'frequency': 'weekly',
        'tasks': [
            'Inspection of cheque book stock.',
            'Spool cheque book request report from OBI.',
            'Confirm the application of cheque book commission charge on request on file.',
            'Approval should be sighted for cheque book not charged.',
            'Overaged books (1 year and above) are treated per the Operational Policy.',
            'Reconcile physical cheque book receipts of orders and acknowledgement advice.',
            'Confirm if cheque book has been received in the CBA.',
            'Confirm if cheque book has been delivered in the CBA if it is in the possession of the customer.',
            'Cheque Acknowledgement advice and register.',
        ],
    },
    {
        'activity': 'ATM Cards/Pins',
        'frequency': 'weekly',
        'tasks': [
            'Inspection of ATM card stock by comparing it with Indigo platform or receipts capturing the card details and the number sent to the branch by E-Business.',
            'Confirm all requested ATM cards are recorded in the register.',
            'Ensure that the PAN number has been masked in the register.',
            'Verify customer\'s signature on the ATM issuance receipts to confirm if it is consistent with signature in CBA.',
            'Confirm that card delivery form sign-off by the officer and customers is filed for all third party collection by RMs/ROs.',
        ],
    },
    {
        'activity': 'Captured ATM Cards',
        'frequency': 'monthly',
        'tasks': [
            'All captured cards must be recorded in the register and treated per bank\'s policy.',
            'Overaged captured ATM Cards (6 months and above) are treated per BOG guideline.',
        ],
    },
    {
        'activity': 'E-Products Request File',
        'frequency': 'weekly',
        'tasks': [
            'Evidence of request on file (Mobile and Internet Banking, SMS, Ghana Pay, POS Merchant, Freedom, E-Zwich etc.).',
            'Treatment of request in the CBA.',
            'Review to confirm applicable charge.',
            'Verify customer\'s signature to confirm if it is consistent with signature in CBA.',
        ],
    },
    {
        'activity': 'Fixed Deposit/Term Deposit',
        'frequency': 'weekly',
        'tasks': [
            'Evidence of request on file.',
            'Verify customer\'s signature to confirm if it is consistent with signature in CBA.',
            'Spool report to compare with request.',
            'Review to confirm if the FD has been booked with the right amount and rate.',
            'Confirm if disposal instructions on maturity have been applied.',
            'Ensure certificates are attached to requests.',
            'There must be evidence of approval if FD has been booked above board rate.',
            'Request is recorded in register and signed by all parties.',
        ],
    },
    {
        'activity': 'Liquidation of Fixed Deposit',
        'frequency': 'monthly',
        'tasks': [
            'Evidence of request on file.',
            'Verify customer\'s signature to confirm if it is consistent with signature in CBA.',
            'Ensure that the interest is paid in line with the Bank\'s policy on premature liquidation of FDs.',
            'Confirm if the penal charges have been applied.',
            'Request is recorded in register and signed by all parties.',
        ],
    },
    {
        'activity': 'Treasury Bill',
        'frequency': 'weekly',
        'tasks': [
            'Check for evidence of request on file.',
            'Verify customer\'s signature to confirm if it is consistent with signature in CBA.',
            'Confirm if T-Bill has been booked and customer\'s account debited with the right amount.',
            'Confirm if disposal instructions on maturity have been applied.',
            'Ensure certificates are attached to requests.',
        ],
    },
    {
        'activity': 'Payment Order',
        'frequency': 'weekly',
        'tasks': [
            'Check for evidence of request on file.',
            'Check for completeness of request and verify customer\'s signature to confirm if it is consistent with signature in CBA.',
            'Confirm if customer\'s account has been debited with PO amount and commission taken per the approved tariff guide.',
            'Check for evidence of approval in cases where the maximum amount is exceeded i.e. GHS30,000.',
            'Ensure that the right persons have signed on the PO leaflet and stub.',
            'Review PO GL to confirm balance.',
        ],
    },
    {
        'activity': 'Standing Order',
        'frequency': 'monthly',
        'tasks': [
            'Check for evidence of request on file.',
            'Verify customer\'s signature to confirm if it is consistent with signature in CBA.',
            'Ensure that the right entries are passed by ensuring that start date, amount, tenor and end date are captured accurately.',
            'Evidence of discontinuation of standing order should be sighted on file.',
        ],
    },
    {
        'activity': 'Returned Cheques',
        'frequency': 'daily',
        'tasks': [
            'All returned cheques must be recorded in the register.',
            'Spool settlement sheet/report from XNETT to aid in review.',
        ],
    },
    {
        'activity': 'Back Office Call Over',
        'frequency': 'daily',
        'tasks': [
            'Review all vouchers processed to detect and correct errors.',
            'Confirm if payment instruments have been verified and the right stamps (received and signature verification) are appended as well as the signature of the processing officer on the stamp.',
            'Check if cheque is postdated/outdated, amount in words and figures are the same and alterations signed against.',
            'Confirm indemnity in the CBA for transactional request sent via email, SMS and WhatsApp.',
        ],
    },
    {
        'activity': 'Branch Registers',
        'frequency': 'weekly',
        'tasks': [
            'Ensure that all mandatory registers are kept and updated in line with the bank\'s approved template.',
            'Confirm evidence of weekly review of registers by BSS.',
        ],
    },
    {
        'activity': 'Display of Regulatory Notices (TV)',
        'frequency': 'monthly',
        'tasks': [
            'Ensure that all regulatory licenses and certificates are displayed and renewed per BOG directives.',
        ],
    },
    {
        'activity': 'GL Review',
        'frequency': 'daily',
        'tasks': [
            'Review the branch\'s GLs daily by using the trial balance as guide.',
            'Check if transactions were posted to the appropriate GLs.',
            'Confirm/investigate GLs with balances that are supposed to be zeroed by EOD.',
            'Investigate all income GLs with negative balances.',
            'Investigate vault with positive balances.',
            'Check if transactions processed into the expense GL are within branch limit or with approvals.',
        ],
    },
    {
        'activity': 'Monthly GL Proof',
        'frequency': 'monthly',
        'tasks': [
            'Review the GL proofs submitted by Reconciliation department.',
            'Check for correctness and accuracy of GL balances captured in the proof report as at close of the month by comparing with OBI.',
            'Check for discrepancies/unexplained outstanding entries on proof.',
            'Check for correct preparation date captured in the proof.',
            'Respond with exceptions if any for correction and update of proof report.',
        ],
    },
    {
        'activity': 'Physical Security',
        'frequency': 'weekly',
        'tasks': [
            'Ensure that all physical security items are in place and functioning (CCTV, panic alarms, smoke detectors, biometric devices, fire alarm panel, fire extinguisher).',
        ],
    },
    {
        'activity': 'CCTV',
        'frequency': 'weekly',
        'tasks': [
            'Confirm if cameras are covering all sensitive areas and recording live footages.',
            'Confirm if display monitor is working.',
            'Confirm if CCTV can playback a minimum of 90 days.',
        ],
    },
    {
        'activity': 'Police & Security',
        'frequency': 'monthly',
        'tasks': [
            'Review to confirm attendance of Police and Security Guards (Name, date, time in/out and signature).',
        ],
    },
    {
        'activity': 'Branch Ambience',
        'frequency': 'weekly',
        'tasks': [
            'Inspect branch for cleanliness of walls, building and branding defects.',
            'Inspect security post for cleanliness.',
        ],
    },
    {
        'activity': "Driver's Log Book",
        'frequency': 'monthly',
        'tasks': [
            "Inspect driver's log book to confirm if it is being populated with details of trips for the day.",
        ],
    },
    {
        'activity': 'ATM Cleanliness',
        'frequency': 'weekly',
        'tasks': [
            'Check if ATM machines are dust-free at all times.',
            'Check if card slot and cash dispenser areas are clean and free from dust.',
            'Ensure that ATM keypads are regularly cleaned.',
            'Check if floor areas within the ATM space are free from dust and debris.',
            'Check if glass doors are clean and free from stains.',
            'Check if dustbins are available and emptied regularly.',
            'Check if surroundings of the ATM are clean and free from litter.',
        ],
    },
]


class Command(BaseCommand):
    help = 'Seed / assign the CC (Cluster Control) activity checklists.'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run', action='store_true',
            help='Preview what would be created without saving.',
        )
        parser.add_argument(
            '--reset', action='store_true',
            help='Delete all existing CC checklists first, then recreate.',
        )

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        reset = options['reset']

        # ------------------------------------------------------------
        # Resolve the target user set — every active CC user
        # ------------------------------------------------------------
        cc_users = UserProfile.objects.filter(
            position='cc', status='active'
        ).order_by('full_name')

        cc_count = cc_users.count()

        if cc_count == 0:
            self.stdout.write(self.style.WARNING(
                '⚠️  No active users with position="cc" found. '
                'Nothing will be assigned.'
            ))

        self.stdout.write(self.style.MIGRATE_HEADING(
            f'\nSeeding {len(CHECKLIST_DATA)} activity checklists'
        ))
        self.stdout.write(f'Target users (position=cc, active): {cc_count}')
        if dry_run:
            self.stdout.write(self.style.WARNING('DRY RUN — no changes will be saved.\n'))
        if reset and not dry_run:
            self.stdout.write(self.style.WARNING('RESET — deleting existing CC checklists first.\n'))
        else:
            self.stdout.write('')

        with transaction.atomic():
            # ------------------------------------------------------------
            # Optional: wipe existing CC-seeded checklists
            # ------------------------------------------------------------
            if reset and not dry_run:
                existing = Checklist.objects.filter(assignment_target='cc')
                n = existing.count()
                existing.delete()
                self.stdout.write(self.style.WARNING(f'Deleted {n} existing CC checklist(s).\n'))

            created = updated = 0
            task_total = 0

            for idx, item in enumerate(CHECKLIST_DATA, 1):
                activity = item['activity']
                frequency = item['frequency']
                tasks = item['tasks']

                # Find or create the checklist
                checklist, was_created = Checklist.objects.get_or_create(
                    name=activity,
                    defaults={
                        'description': f'CC activity checklist — {frequency}',
                        'frequency': frequency,
                        'assignment_target': 'cc',
                        'is_active': True,
                    },
                )

                if was_created:
                    created += 1
                    self.stdout.write(f'  [{idx:02d}] ✅ Created  {activity}  ({frequency}, {len(tasks)} tasks)')
                else:
                    updated += 1
                    # Keep assignment_target='cc' even if it existed as 'all'
                    checklist.assignment_target = 'cc'
                    checklist.frequency = frequency
                    checklist.is_active = True
                    self.stdout.write(f'  [{idx:02d}] ♻️  Updated  {activity}  ({frequency}, {len(tasks)} tasks)')

                if dry_run:
                    continue

                checklist.save()

                # ------------------------------------------------------------
                # Assign every active CC user
                # ------------------------------------------------------------
                if cc_count:
                    checklist.assigned_users.set(cc_users)

                # ------------------------------------------------------------
                # Rebuild tasks (clear then recreate — simplest + safe)
                # ------------------------------------------------------------
                ChecklistTask.objects.filter(checklist=checklist).delete()
                for order, task_desc in enumerate(tasks):
                    ChecklistTask.objects.create(
                        checklist=checklist,
                        description=task_desc,
                        order=order,
                        is_completed=False,
                    )
                task_total += len(tasks)

        # ------------------------------------------------------------
        # Summary
        # ------------------------------------------------------------
        self.stdout.write('')
        self.stdout.write(self.style.SUCCESS('─' * 60))
        self.stdout.write(self.style.SUCCESS('  DONE'))
        self.stdout.write(self.style.SUCCESS('─' * 60))
        self.stdout.write(f'  Checklists created:  {created}')
        self.stdout.write(f'  Checklists updated:  {updated}')
        self.stdout.write(f'  Tasks written:       {task_total}')
        self.stdout.write(f'  Users assigned to:   {cc_count}')
        if dry_run:
            self.stdout.write(self.style.WARNING('  (DRY RUN — nothing saved)'))
        self.stdout.write('')

        if dry_run:
            self.stdout.write(self.style.MIGRATE_HEADING(
                'Run again WITHOUT --dry-run to actually save.\n'
            ))