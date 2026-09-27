"""
Seed Head Office department activity checklists.

Every "team" column value becomes a Department. Every activity
becomes a Checklist scoped to that Department and assigned to
every active user inside it.

Idempotent — safe to re-run. Tasks are rebuilt from the source
on every run.

Usage:
    python manage.py seed_ho_departments --dry-run
    python manage.py seed_ho_departments
    python manage.py seed_ho_departments --team "IT CONTROL"
    python manage.py seed_ho_departments --reset
"""

from django.core.management.base import BaseCommand
from django.db import transaction

from control_dashboard.models import (
    Checklist, ChecklistTask, Department, UserProfile,
)


# ============================================================
# FREQUENCY NORMALIZER — model only accepts these values
# ============================================================
VALID_FREQ = {'daily', 'weekly', 'monthly', 'quarterly', 'bi-annual', 'annual'}

FREQ_ALIAS = {
    'bi-weekly': 'weekly', 'biweekly': 'weekly', 'bi weekly': 'weekly',
    'fortnightly': 'weekly', 'annually': 'annual', 'yearly': 'annual',
    'bi-annually': 'bi-annual', 'biannual': 'bi-annual',
    'bi annual': 'bi-annual', 'half-yearly': 'bi-annual',
    'one-off': 'monthly', 'one off': 'monthly', 'oneoff': 'monthly',
}


def norm_freq(raw):
    key = (raw or '').strip().lower()
    key = FREQ_ALIAS.get(key, key)
    return key if key in VALID_FREQ else 'monthly'


# ============================================================
# MASTER DATA
# Format per team: list of (activity, frequency, [tasks])
# ============================================================
TEAM_DATA = {
    'IT CONTROL': [
        ('User Access and Privileges Management Reviews', 'monthly', [
            'Validate user roles vs job functions',
            'Review access changes and approvals',
            'Check for excessive privileges',
            'Monitor privileged user activities',
            'Ensure timely revocation of access',
        ]),
        ('Data Backup processes, Backup to Tapes, Offsite Movement, Retrieval, Testing & Restoration', 'quarterly', [
            'Verify backup schedules and success rates',
            'Review offsite storage and media handling',
            'Test backup restoration periodically',
            'Validate encryption and data protection',
            'Check backup logs and failures',
        ]),
        ('Database Consistency & Integrity Reviews', 'bi-annual', [
            'Validate data integrity checks and controls',
            'Review DB errors and corruption logs',
            'Check reconciliation processes',
            'Monitor DB performance and anomalies',
            'Ensure audit logging is enabled',
        ]),
        ('Disaster Recovery (DR) Replication Review', 'annual', [
            'Monitor replication status and lag',
            'Review replication logs and errors',
            'Validate failover readiness',
            'Check replication security controls',
            'Ensure completeness of replicated data',
        ]),
        ('Income Assurance', 'monthly', [
            'Review income assurance activities',
        ]),
        ('IT Change and Configuration Management reviews', 'monthly', [
            'Review change requests and approvals',
            'Validate testing and rollback plans',
            'Check unauthorized changes',
        ]),
        ('IT Asset Management and Inventory reviews', 'quarterly', [
            'Verify asset register accuracy',
            'Track asset ownership and location',
            'Review asset lifecycle (procurement–disposal)',
            'Check compliance with policies',
            'Identify missing or untagged assets',
        ]),
        ('Network & Security Review', 'bi-annual', [
            'Review network configurations and segmentation',
            'Validate security controls and policies',
            'Check firewall and access rules',
            'Monitor network anomalies',
            'Assess vulnerability exposure',
        ]),
        ('SIEM Review (Server & Application Logs)', 'annual', [
            'Ensure log integration (OS + DB + apps)',
            'Review alerts and incident handling',
            'Validate use cases and correlation rules',
            'Check log retention and integrity',
            'Identify monitoring gaps',
        ]),
        ('Branch Infrastructure Review', 'annual', [
            'Inspect branch IT setup and controls',
            'Verify connectivity and uptime',
            'Review endpoint security compliance',
            'Check hardware standards',
            'Identify infrastructure risks',
        ]),
        ('Physical Access & Environmental Control (Data Center and DR Site)', 'quarterly', [
            'Review access to data center/DR site',
            'Validate CCTV and monitoring systems',
            'Check fire suppression and cooling systems',
            'Inspect access logs and controls',
            'Ensure environmental compliance',
        ]),
        ('System Health Checks (Core Banking, E-Banking, ATM Switch, DB, Servers, Network)', 'bi-annual', [
            'Monitor uptime and performance',
            'Review system logs and errors',
            'Validate capacity and utilization',
            'Check service availability',
            'Identify recurring issues',
        ]),
        ('Incident and Response Reviews', 'bi-annual', [
            'Review incident logs and classification',
            'Validate response timelines',
            'Check root cause analysis',
            'Ensure proper escalation',
            'Track closure and lessons learned',
        ]),
        ('VPN Access Review', 'quarterly', [
            'Validate authorized VPN users',
            'Review access logs and usage',
            'Check MFA implementation',
            'Identify unusual access patterns',
            'Remove inactive users',
        ]),
        ('Vulnerability & Patch Management', 'quarterly', [
            'Review vulnerability scan results',
            'Track patch deployment status',
            'Validate critical patch timelines',
            'Identify unpatched systems',
            'Ensure remediation tracking',
        ]),
        ('Card Management Review', 'bi-annual', [
            'Review card issuance and controls',
            'Validate transaction monitoring',
            'Check fraud detection mechanisms',
            'Ensure compliance with standards',
            'Investigate anomalies',
        ]),
        ('DR Simulation Exercise Validation', 'annual', [
            'Review DR test plans and execution',
            'Validate recovery time (RTO/RPO)',
            'Check system failover success',
            'Identify gaps and improvements',
            'Ensure documentation of results',
        ]),
        ('IT SLA Compliance Review', 'bi-annual', [
            'Monitor SLA performance metrics',
            'Review service delivery reports',
            'Identify breaches and penalties',
            'Validate vendor performance',
            'Ensure corrective actions',
        ]),
        ('Firewall, IDS/IPS & EDR Alerts Review', 'bi-annual', [
            'Analyze security alerts and incidents',
            'Validate rule effectiveness',
            'Identify false positives/negatives',
            'Ensure timely response',
            'Tune detection mechanisms',
        ]),
        ('Projects and SDLC Review', 'annual', [
            'Review project governance and approvals',
            'Validate security in SDLC',
            'Check testing and sign-offs',
            'Monitor change control',
            'Ensure documentation completeness',
        ]),
        ('API/Endpoint Security Review', 'bi-annual', [
            'Validate API authentication and authorization',
            'Review endpoint protection controls',
            'Check for exposed services',
            'Monitor API traffic and logs',
            'Identify vulnerabilities',
        ]),
        ('Review of End of Month (EOM) processes, activities, error logs and their resolutions', 'monthly', [
            'Review end-of-month activities',
            'Validate reconciliations and reports',
            'Analyze error logs',
            'Confirm issue resolution',
            'Ensure completeness of processing',
        ]),
        ('Review of End of Year (EOY) processes, activities, error logs and their resolutions', 'annual', [
            'Validate year-end processing activities',
            'Review financial/data reconciliations',
            'Check error logs and resolutions',
            'Ensure compliance requirements met',
            'Confirm audit readiness',
        ]),
        ('Third-Party / Vendor Risk Management', 'bi-annual', [
            'Assess vendor security posture through due diligence and risk assessments',
            'Review third-party access to systems and enforce least privilege',
            'Evaluate contracts and SLAs for security, compliance, and data protection clauses',
            'Monitor vendor performance, incidents, and compliance with agreed controls',
            'Ensure periodic reassessment and offboarding procedures for third parties are in place',
        ]),
        ('Database Review', 'bi-annual', [
            'Review database access and privileges',
            'Monitor database logs and user activities',
            'Validate security configurations and controls',
            'Assess backup, recovery, and data integrity',
            'Track and verify database changes and audit trails',
        ]),
        ('Storage & SAN/NAS Review', 'monthly', [
            'Storage utilization, RAID, replication, capacity, performance, alerts and access controls',
        ]),
        ('Virtualization Review', 'monthly', [
            'VMware/Hyper-V configuration, administrator access, VM inventory, snapshots, patching, HA/DRS and security',
        ]),
        ('Technology Resilience Review', 'monthly', [
            'Single points of failure, HA, clustering, redundancy, failover and resilience of critical systems',
        ]),
        ('Privileged Access Management (PAM) Review', 'monthly', [
            'Domain Admin, DBA, root, application admin accounts, privileged session monitoring, emergency/break-glass accounts',
        ]),
        ('Service Account Review', 'monthly', [
            'Review all service accounts',
        ]),
    ],

    'ICDD': [
        ('VISA STATEMENT VERIFICATION', 'daily', [
            'Receive request via corporate email for Visa Verification (Embassy or accredited Institution)',
            'Capture the details of the request on the VISA statement schedule (date, name, account number, embassy name and indicate the status of the request Yes/No); thus consistency with bank details.',
            'Review all attached documentation and confirm consistency with details in the CBA/OBI',
            'In an instance where it is required the Bank confirms that customer is aware of their statement being used for VISA purposes, send a mail to the RM and copy Head of SBU, BM, BSS with ICDD Head Office Control mailing group.',
            'Follow up with RM/RO if no responses is received after 3hours and escalate to the Head of SBU/Branch.',
            'Draft the initial response based on the queries from the embassy and confirmation done in the CBA/OBI',
            'Send the draft copy to ICDD checker via email for further review',
            'Respond accordingly via mail from the embassy or corporate institution using the existing email and subject after the checkers response/comment.',
            'Send the request from the embassy to CPU to take the charge per the prevailing Bank tariff guide',
            "In an instance where customer's account balance is insufficient, send a mail to the RM/RO/Branch/SBU with ICDD Head office Control in copy",
            'Follow up and ensure that the account is funded for the charge to be taken.',
            'Review the schedule from time to time and ensure that all charges are taken.',
        ]),
    ],

    'TROPs': [
        ('Call over FDs Booked', 'weekly', [
            'Request for FDs request forms/instructions processed by the TROPs officer',
            'Review the form for completeness and confirm signature in line with account mandate',
            'Confirm indemnity in the CBA for request sent via email, SMS and WhatsApp.',
            'Spool Term Deposit reports from OBI',
            'Compare details captured on the form to details captured in OBI (processing date, maturity date, amount, rate, tenor, rollover options)',
            'Sight approval if the booked rate is above the board rate',
            'Investigate if there any discrepancies',
        ]),
        ('Call over FDs Liquidated (Retail Customers)', 'monthly', [
            "Review customer's instruction to pre-liquidate FD (confirm the customer's signature in line with account mandate)",
            'Confirm indemnity in the CBA for transactional request sent via email, SMS and WhatsApp.',
            'Review computation to confirm if the applicable penal charges were applied and the right interest and principal credited to the customer based on the number of days the investment has travelled.',
        ]),
        ('Call over FDs Liquidated (Corporate Customers)', 'weekly', [
            "Since customers' accounts are automatically credited upon maturity, one has to confirm if the withholding tax i.e. 25% of the accrued interest has been charged to the customer.",
            'Investigate if there are discrepancies',
        ]),
        ('Booking of T-Bills', 'monthly', [
            'Request for T-Bill request forms/instructions processed by the TROPs officer',
            'Review the form for completeness and confirm signature in line with account mandate',
            'Confirm indemnity in the CBA for transactional request sent via email, SMS and WhatsApp.',
            'Compare details captured on the form to details captured on the certificate.',
            'Confirm if the account has been debited.',
        ]),
        ('Pre-Liquidation of T-Bills', 'monthly', [
            "Review customer's instruction to pre-liquidate T-Bill (confirm the customer's instruction by verifying signature in line with account mandate or confirming with indemnity)",
            'Confirm indemnity in the CBA for transactional request sent via email, SMS and WhatsApp.',
            'Review computation to confirm if the customer has been credited with the right principal and applicable interest.',
        ]),
        ('Review of Deposits Placed/Taken', 'monthly', [
            'Request for deal slip and review for completeness (check for approval from Market Risk and Treasury in line with the approval matrix)',
            'Spool callover of deposits placed/taken',
            'Compare details indicated on deal slip to details captured in the OBI report (check details such as counterparty, principal, interest, value date, maturity date, tenor, rate, whether funds have been fully secured with T-Bills/other investments)',
        ]),
        ('FX Sale/Purchase (Bank)', 'monthly', [
            'Request for deal slip and request to purchase or sell FX',
            'Confirm the rate agreed on the request form with the rate indicated on the deal slip',
            'Investigate if there are discrepancies',
        ]),
        ('FX Sale/Purchase (Customer)', 'monthly', [
            "Request for deal slip and customer's request to purchase or sell FX",
            "Confirm the customer's instruction by verifying signature in line with account mandate)",
            'Confirm indemnity in the CBA for transactional request sent via email, SMS and WhatsApp.',
            'Confirm the rate agreed on the request form with the rate indicated on the deal slip',
            'Investigate if there are discrepancies',
        ]),
    ],

    'TREASURY': [
        ('Deal Room', 'monthly', [
            'Check that biometric access devices are installed at the deal room.',
            'Confirm only authorised staff has access to the deal room.',
            'Inspect the set up and use of approved systems and applications - computer system, communication equipment (recorded), TV set.',
            'Confirm that dealers have the required certification to perform their functions (ACI)',
        ]),
        ('INTERBANK LENDING/PLACEMENT', 'weekly', [
            'Check whether the amount to lend is within the dealers limit.',
            'Confirm if there is approval for interbank lending transaction.',
            'Check that approved rate has been applied.',
            'Confirm deal and settlement with the counter party.',
            'Check for collateral supporting the amount lending.',
        ]),
        ('INTERBANK TAKING/BORROWING', 'weekly', [
            'Confirm if the amount to be borrowed is within the dealer\u2019s limit.',
            'Check if deal has been authorised',
            'Confirm if the borrowing is secured with collateral',
        ]),
        ('TEST OF FOREIGN FINANCIAL EXPOSURE', 'weekly', [
            'Check Banks transactions/investment with foreign bank does not breach the BOG regulatory limit (Foreign Sovereign and Bank Exposures: Capped at 20% of Net Owned Funds (NOF) for banks and 15% for SDIs).',
            'Check for approvals from BOG in instances of breach',
        ]),
        ('TREASURY AUCTION /REQUEST', 'monthly', [
            'Check that customers request has been uploaded in Treasury T.Bill request folder.',
            'Treasury officer downloads the uploaded request and completes an excel schedule with the request details',
            'The schedule completed is forwarded the Banks primary dealer for auction.',
            'Following the results of the auction, TROPS is advised with the details of the purchase.',
            'For Unsuccessful auction and wrong account details, branches are advised accordingly.',
        ]),
    ],

    'TRANSPORT': [
        ('Review of list of vehicles/motor bike and locations', 'monthly', [
            'Confirm whether the master list has been updated to reflect all vehicles/motor bikes held by the bank and their locations',
            'Confirm the existence of vehicle spare keys held by the department and how they are managed.',
        ]),
        ('Review of Drivers License', 'monthly', [
            'Inspect the list of drivers',
            'Confirm if all drivers have copies of their licences in place and on file.',
            'For new drivers, confirm whether drivers have had their license for a minimum of three (3) years.',
            'Confirm the expiry date of licences.',
            'Review the schedule monitoring the expiry date of drivers licence',
        ]),
        ('Review of vehicle registration, insurance and road worthy', 'monthly', [
            'Confirm the existence and expiry of the insurance policy for each vehicle',
            'Confirm the existence and expiry of road worthy of vehicles held by the bank.',
            'Confirm whether the road worthy or insurance policy have been renewed per their expiry date',
            'Third party or comprehensive',
        ]),
        ('Scheduled Maintenance for Vehicles', 'monthly', [
            'Confirm whether the schedule for vehicle maintenance has been updated',
            'Confirm whether vehicle maintenance was done as per the scheduled time',
            'NB: Maintenance is done usually every three months / per the time scheduled for maintenance / or when the vehicle has covered 5000 mileage',
        ]),
        ('Insurance claim', 'monthly', [
            'Confirm the existence of an incidence report on file',
            'Confirm the existence of a mail or memo sent to the insurance company for claim',
            'Confirm whether a schedule is maintained for pre-finance payment made on vehicle repairs.',
            'Compare the GL (Sundry creditors - Insurance Claim) with the maintained schedule to confirm whether all insurance claims have been received.',
            'Review the monthly GL proof for the Sundry creditors - Insurance claim GL.',
        ]),
        ('Sinking fund', 'monthly', [
            'Confirm whether a schedule has been maintained for payment made using the sinking fund',
            'Confirm the existence of the incident report on file requiring payment from the GL',
            'Ensure reports have been properly filed',
        ]),
        ('Access Control Review/CCTV', 'monthly', [
            'Review the active access to confirm that all exited staff have been deactivated.',
            'Review the active access list to confirm whether staff have the right access and branch.',
            'Confirm if access to the NVR has been created for BSS and BM.',
            'Review the CCTV footages at head office to ensure all sensitive areas are properly captured/monitored.',
        ]),
        ('Private Security guards coordination', 'monthly', [
            'Confirm the existence of a monthly report indicating that security guards report at their respective branches and efficient with the discharge of duties.',
            'Confirm whether payment were made for security personnel who failed to report to work during the month.',
        ]),
        ('Review of the courier services', 'weekly', [
            'Confirm whether incoming envelope mails have been time stamped',
            'Confirm whether the department keeps a copy of received waybill on items received by the branch.',
            'Review waybills to ensure full completion of details or required fields',
            'Review the dispatch log book and ensure the log book is duly updated',
        ]),
        ('Review of travels', 'monthly', [
            'Confirm whether all document in relation to a particular travel has been properly filed.',
            'Confirm whether hotel bookings are done with approved hotels',
            'Confirm whether the required approvals for travels exist and are on file',
        ]),
        ('SLA review', 'monthly', [
            "Confirm the existence of SLA with vendors. Outdated SLA's should be escalated for revision.",
        ]),
        ('Speed limit of drivers', 'monthly', [
            'Confirm that drivers are within the required speed limit',
        ]),
        ('Confirmation of Tracking devices on vehicles', 'monthly', [
            'Confirm whether the tracking devices are working on Vehicles',
        ]),
    ],

    'TRADE': [
        ('GL REVIEW', 'daily', [
            'Spool GLs to identify any mispostings and wrong GL balances',
            'Escalate any discrepancy identified',
        ]),
        ('CALLOVER', 'monthly', [
            'Spool callover report on OBIE for Trade Officers',
            'For each transaction in the callover report, request for documentation and validate',
            'Escalate any discrepancy identified',
        ]),
        ('LETTERS OF CREDIT', 'monthly', [
            'Commercial Invoice',
            'Bill of Lading',
            'Insurance Certificate',
            'Packing List',
        ]),
    ],

    'SERVICE QUALITY': [
        ('CONTACT CENTRE', 'weekly', [
            'Review Enquiries/Requests/Complaints Received via Phone from External Customers',
            'Review Enquiries via Social Media (Facebook, Twitter, Instagram, Website, etc)',
            'Review Card-Blocking Request/Enquiry/Complaints Report',
            'Review Complaints Received via Emails from Customers',
            'Review of Branch Call-Back Process & Post-Visit Customer Feedback',
        ]),
        ('SLA REPORTS', 'monthly', ['Review monthly SLA report']),
        ('STATUTORY', 'monthly', ['Review quarterly report to BoG']),
        ('SOLAR WIND', 'weekly', ['Weekly downtime report']),
        ('SOCIAL MEDIA REPORT', 'monthly', ['Review Social Media report']),
        ('CONTACT CENTRE REPORT', 'monthly', ['Review contact centre call report']),
        ('BRANCH ROVING EYE REPORT', 'monthly', ['Check for branch roving eye report']),
    ],

    'REMITTANCE': [
        ('Review of Remittance accesses', 'monthly', [
            'Review remittance access (RIA, Western Union, MoneyGram, and Unity Link) for all branches and ensure that each branch has access to all remittance services.',
        ]),
        ('Review of Remote transactions', 'monthly', [
            'Review the remittance forms to confirm they have been fully completed and that a copy of the customer\u2019s Ghana Card is attached.',
            'Verify that the information on the form is consistent with the transaction receipt.',
            'Confirm that the Ghana Card has been verified.',
            'Enquire why the request was not processed at the branch, especially during weekdays.',
        ]),
    ],

    'RECOVERY': [
        ('Recovery Targets & Monitoring Reports', 'monthly', [
            'Confirm targets have been assigned to all Recovery Officers.',
            'Check the existence of performance monitoring reports.',
            'Check if Reports on the status of defaulting loans have been prepared.',
            'Confirm level of compliance with recovery target by reviewing call report.',
            'Confirm customers account has been credited with the collections made.',
        ]),
        ('COLLECTIONS/DEBT RECOVERY', 'monthly', [
            'Check whether a loan collection schedule exists.',
            'Confirm amounts collected have been credited to the Omni recovery account or the respective customer account.',
            "Confirm respective amounts collected have been transferred from the customers' accounts and income recognised by Finance department accordingly.",
        ]),
    ],

    'RECON UNIT': [
        ('Review of Recon Report', 'weekly', [
            'Request for Nostro and Central Bank account statement via email for review the Recon team.',
            'Match the balances per Bank statement with balances that have been captured in the recon report',
            'Confirm all outstanding items in the recon report with the mirror statement and nostro statement.',
            'Confirm all closing balances in the recon report by checking balances in CBA.',
        ]),
        ('GL PROOF REPORT', 'monthly', [
            'Confirm if the Recon unit is reviewing the GL proof report submitted by the branches and Head Office Department.',
            'Check for timely submission of the GL proof report to ICDD',
            'Check if exceptions identified by ICDD has been addressed and the GL proof report updated.',
        ]),
        ('Review of DOT.GOV ACCOUNT', 'monthly', [
            'Request for Dot.Gov Portal statement from the recon unit via mail and spool Dot.Gov account statement from CBA.',
            'Check for outstanding entries which are yet to be settled and ensure they match the balances on the account.',
            'Raise exceptions for long outstanding transactions and discrepancies if any.',
            'Confirm if collections received have been transferred to the agencies (GRA, SSNIT) within the agreed timelines.',
        ]),
        ('Review of AWA and Passion GL', 'monthly', [
            'Request for Dot.Gov Portal statement from the recon unit via mail and spool Dot.Gov account statement from CBA.',
            'Check for outstanding entries which are yet to be settled and ensure they match the balances on the account.',
            'Raise exceptions for long outstanding transactions and discrepancies if any.',
            'Confirm if collections received have been transferred to the agencies (GRA, SSNIT) within the agreed timelines.',
        ]),
    ],

    'LEGAL': [
        ('MORTGAGE REGISTRATION/PERFECTION', 'monthly', [
            'Confirm the receipt of the mortgage deed by the department.',
            'Confirm search has been conducted on the property and report filed.',
            'Verify there is no interest registered on the property.',
            'Check customer has been debited with the mortgage registration fee.',
            'Confirm payment for stamping on the property has been paid.',
            'Inspect the land document (indenture) has been stamped',
        ]),
        ('MORTGAGE DISCHARGE', 'monthly', [
            'Confirm the customer has no exposure with the Bank',
            "Record customers' details in the discharge register",
            'Customer signs off the deed register as evidence of receipt of mortgage document',
        ]),
        ('DRAFTING AND REVIEW OF AGREEMENT', 'monthly', [
            'Confirm if SLAs received are recorded',
            'Confirm if SLA is reviewed with the approved TAT',
        ]),
        ('STAFF AUTO LOAN (Documentation)', 'monthly', [
            'Check that spare keys are kept for all staff auto loans',
            'Check that spare keys are labelled',
            'Confirm that original car documents have been submitted',
        ]),
    ],

    'HCM': [
        ('AUTO LOANS', 'monthly', [
            'Review checklist to ensure completeness of auto loan requirements',
            'Confirm submission of Spare keys',
            'Confirm authenticity of Spare keys',
        ]),
        ('EXITED STAFF', 'monthly', [
            'Confirm if exited staff accesses have been disabled.',
            'Review to confirm accurate computation & payment of compensation for exited staff',
            'Review staff indebtedness & a plan for repayment.',
            'Review of reclassification of exited staff account',
            'Review signed exit clearance form with handing over notes',
            'Confirm if all customer accounts (RM) have been retagged.',
        ]),
        ('TRAINING', 'monthly', [
            'Review of Training Schedule for Staff',
            'Confirm adherence to Staff training schedule',
            'Confirmation of participation by staff on training ie Evaluation form, attendance list',
            'Check for submission of training reports',
        ]),
        ('RECRUITMENTS', 'monthly', [
            'Review checklist for new recruits',
        ]),
        ('STAFF LEAVE PORTAL', 'monthly', [
            'Check for submission of leave plan by departments/branches',
            'Check adherence to leave plan',
        ]),
        ('STAFF ALLOWANCE PAYMENT', 'monthly', [
            'Review fuel allowance schedule to confirm accuracy & completeness',
            'Review Phone credit schedule to confirm accuracy & completeness',
        ]),
        ('EMPLOYEE RELATIONS', 'monthly', [
            'Review of Contract between the Bank and Food vendors to validate adherence by both parties',
            'Periodic Certification of Food Vendors',
        ]),
        ('REDEPLOYMENT HANDING OVER NOTES', 'monthly', [
            'Check for submission of handing over notes by redeployed staff',
        ]),
        ('REASSIGNING OF EXITED ACCOUNT PORTFOLIO', 'monthly', [
            'Review reassignment of exited account portfolio',
        ]),
    ],

    'GENERAL SERVICES': [
        ('Store Room Management', 'monthly', [
            'Count the physical stock held by the department and compare with the schedule kept by the department.',
            'Any difference should be addressed',
            'Confirm whether the store room is accessed only by authorized persons',
            'Confirm the existence of a register for attendance',
        ]),
        ('Review of approved vendors', 'monthly', [
            'Confirm existence of a schedule of approved vendors',
            'Confirm the existence of the SLAs with the vendors and any breach should be escalated.',
        ]),
        ('Review of approvals by the procurement committee', 'monthly', [
            'Review the list of procurement requests sent to the committee for approval.',
            'Review the report of approved procurement and whether approvals have been signed by all stakeholders',
        ]),
        ('SIM Packages/MIFI', 'monthly', [
            'Confirm list of staff on sim package (ensure staff are on the required grade to enjoy the package)',
            'Confirm whether exited staff have been taken off the schedule.',
            'Request for active sim packages from telcos.',
            'Confirm department with MIFI allocation',
            'Confirm use of active sim cards by the respective departments/individuals',
        ]),
        ('Logistics Requisition', 'monthly', [
            'Review the request sent by the requesting branch or unit. This should be signed by the initiator & head of department of unit',
            'The request should be duly completed.',
        ]),
        ('Vendor evaluation Review', 'monthly', [
            'Confirm whether vendors are evaluated per agreed policy',
            'Confirm evidence of evaluation done',
            'Confirm whether feedback has been relayed to stakeholders',
            'Escalation for management decision',
        ]),
        ('Distributions of water and dailies', 'monthly', [
            'Confirm whether water and dailies are being distributed to the right personnel',
            'Confirm whether a file is kept as evidence of receipt by stakeholders',
        ]),
        ('Handing over', 'monthly', [
            'Confirm whether the department keeps a handing over file and it is duly updated',
        ]),
    ],

    'FINOPS': [
        ('Call Over of transaction vouchers', 'daily', [
            'Review all vouchers processed to detect and correct errors.',
            'Confirm if payment vouchers have been authorized in accordance to the expense management approval limit.',
            'Confirm whether the budget has been completed',
            'Confirm if all needed document have been attached to the payment vouchers',
            'Confirm if approvals for cash advance request have been adhered to',
            'Confirm receipt for cash advance retirement',
            'Confirm if transactions are posted into the correct account/GLs',
            'Investigate if there are any discrepancies',
        ]),
        ('Payment Order Review', 'weekly', [
            'Check for evidence of request on file.',
            'Check for completeness of request',
            "Confirm cheque is being issued chronologically and it's being signed by the authorized signatories",
            'Ensure the register is duly completed and updated at all times',
        ]),
        ('GL review', 'daily', [
            'Review the Head office GLs daily by using the trial balance as guide.',
            'Check if transactions were posted to the appropriate GLs with narrations properly captured.',
            'Confirm/investigate transit GLs with balances at close of business',
            'Confirm/investigate GLs with an irregular GL position',
        ]),
        ('Monthly GL proof', 'monthly', [
            'Review the GL proofs submitted by Reconciliation department',
            'Check for correctness and accuracy of GL balances captured in the proof report as at close of the month by comparing with OBI',
            'Check for discrepancies/unexplained outstanding entries on proof.',
            'Check for correct preparation date captured in the proof.',
            'Respond with exceptions if any for correction and update of proof report.',
        ]),
        ('Cash advance review', 'weekly', [
            'Review the schedule to confirm staff who have breached the retiring period (10 working days)',
            'Confirm whether staff who breach have been debited',
        ]),
        ('Handing over', 'monthly', [
            'Confirm whether the department keeps a handing over file and it is duly updated',
        ]),
    ],

    'FINANCE': [
        ('Prepayment Schedule', 'monthly', [
            'Review the schedule prepared by the department for correctness by comparing the closing balance on the schedule against the GL balance.',
            'Download the prepayment GL for the month and confirm whether all items in the GL have been captured on the schedule and amortized.',
            'Escalate all the prepayment that has not been captured in the schedule for amortization',
        ]),
        ('Accrual Schedule', 'monthly', [
            'Review the schedule prepared by the department for correctness by comparing the closing balance on the schedule against the GL balance.',
            'Balances must agree.',
            'Download the accrual GLs for the month and confirm whether all items on the GL have been captured on the schedule',
            'Escalate all accruals or provisions that has not been captured on the schedule.',
            'All GLs with negative balance should be investigated.',
        ]),
        ('Management Accounting (MA) Schedule', 'monthly', [
            'Review the schedule prepared by the department for correctness by comparing the closing balance on the schedule against the GL balance.',
            'Balances must agree.',
            'Download the accrual GLs for the month and confirm whether all items on the GL have been captured on the schedule',
            'Escalate all accruals or provisions that has not been captured on the schedule.',
            'All GLs with negative balance should be investigated.',
        ]),
        ('Fixed Asset Register', 'monthly', [
            'Download the fixed asset acquisition GL for the month',
            'Compare details on the fixed asset acquisition GL to the fixed asset schedule prepared by the department to ensure all assets have been captured',
            'Confirm whether assets are being depreciated',
            'Escalate all issues',
        ]),
        ('Regulatory Returns Review', 'monthly', [
            'Download the trial balance as at end of month',
            'Calculate the banks total balance or position for all GLs by summing up closing balances of the GLs on the schedule.',
            'Review the schedule prepared by the department for correctness by comparing the balances on the schedule and with re-computed balances.',
            'Address all differences identified',
        ]),
        ('Budget Monitoring Report Review', 'monthly', [
            'Confirm whether yearly budget have been signed off and by the right authorizers per the SOP.',
            'Compare the actual expenses to the budgeted one and any deviation should be addressed and escalated.',
            'Where there are deviations confirm whether there are approvals or supplementary budget in place.',
        ]),
        ('GL Creation Report Review', 'monthly', [
            'Confirm the existence of a request to create a GL with the report from OBIE',
            'The request should be duly completed by all stakeholders',
            'Confirm whether or not the GL has been created.',
        ]),
        ('Review of Approval of Asset held for sale', 'monthly', [
            'Confirm the list of asset held for sale',
            'Request for approval for Asset Held for Sale for more than one year',
        ]),
        ('GL review', 'daily', [
            'Check if transactions were posted to the appropriate GLs.',
            'Confirm/investigate transit GLs with balances at close of business',
            'Confirm/investigate GLs with a wrong GL position',
        ]),
        ('Call Over of transaction vouchers', 'weekly', [
            'Review all vouchers processed to detect and correct errors.',
            'Confirm if payment vouchers have been authorized in accordance to the expense management approval limit.',
            'Confirm if all needed document have been attached to the payment vouchers',
            'Investigate if there are any discrepancies',
        ]),
        ('Debit reversal', 'monthly', [
            "Confirm the existence of a request form to reverse debit on customer's account",
            'Confirm whether reversal has been done',
            'Confirm whether the right amount was debited.',
            'Investigate if there are any discrepancies',
        ]),
    ],

    'HEAD OFFICE': [
        ('Expense Review (Requests <= GHS20,000)', 'daily', [
            'Capture details of expense on expense review schedule/tracker.',
            'Check expense schedule to confirm whether PV has been represented.',
            'Check whether PV has been signed off by the initiator and the Head of the requesting department.',
            'Check whether all fields on the PV form have been completed.',
            'Check whether amount quoted/invoice is in line with existing SLA.',
            'Check whether relevant approvals have been obtained',
            'Check whether receipt for cash advance corresponds/add up to cash advanced. Unused funds should be deposited into the branch transit account.',
            'Confirm whether amount in words corresponds to amount in figures.',
            'Recompute taxes/tariffs indicated on invoices.',
            'Check for job completion for services provided.',
            'When necessary, escalate identified issues with the relevant departments.',
            'After all requirements are met, stamp to indicate it has been supported.',
            'Update the expense review schedule to reflect the supported PV/cash advance form',
            'Indicate on the schedule when and who the PV was dispatched through',
        ]),
        ('Expense Review', 'daily', [
            'Check whether budget schedule has been completed.',
            'Check whether invoices/memos attached support the amount and purpose being paid for.',
            'Check whether amount quoted/invoice is in line with existing SLA.',
            'Check whether relevant approvals have been obtained',
            'Check whether receipt for cash advance corresponds/add up to cash advanced. Unused funds should be deposited into the branch transit account.',
            'Escalate identified issues with the relevant departments.',
        ]),
    ],

    'CREDIT': [
        ('Loan Application', 'weekly', [
            'Confirm receipt of loan requests including board resolutions if its a company account.',
            'Confirm checks with the credit reference bureau on a prospective borrower.',
            'Confirm receipt of: Surveyors report, valuation report, insurance on property, property rate, EPA, Fire cert, BOP, Submission of property inspection report.',
            'Check for security cover on the facility.',
            'Confirm collateral status with Legal department for landed properties.',
            'Check in CBA to confirm cash collateral has been booked as Fixed deposit/Lien',
            'Confirm corporate guarantees have been verified.',
        ]),
        ('Revenue assurance', 'weekly', [
            'Confirm appropriate fees on the facility has been charged.',
        ]),
        ('Pre-disbursement', 'weekly', [
            'Check that disbursement condition has been met.',
            'Where conditions have been deferred, check for approval and deferral timelines',
        ]),
        ('Post Disbursement condition', 'weekly', [
            'Check for monitoring of post disbursement conditions.',
            'Check for receipt of reports on monitoring activities.',
            'Confirm the achievement turnover requirements',
        ]),
        ('Draw down conditions', 'weekly', [
            'Confirm drawdown conditions has been met',
        ]),
        ('GUARANTEES ISSUED', 'monthly', [
            'Check that the issue and monitoring fees on Guarantee issued have been charged.',
            'Check that supporting documents on the guarantee has been filed.',
        ]),
    ],

    'CORPORATE COMMUNICATIONS': [
        ('Stock Count', 'monthly', [
            'Request for the list of merchandise from the Corporate Communications Department.',
            'The list should include the opening stock',
            'Compare records indicated in the list to the physical count of items.',
            'Review to confirm if there are records of distribution lists for merchandise given out as a second level check.',
            'Investigate if there are any discrepancies.',
        ]),
        ('Review of Information on Bank Website', 'monthly', [
            'Go through all information indicated on the website to confirm accuracy',
            "Review information to confirm if identified information is accurate and aligns with bank's policy",
        ]),
        ('Review of Branded Souvenirs GLs', 'monthly', [
            'Capture the GL number (Branded Souvenirs) in the OBIE portal',
            'Spool the GL for the required period',
            "Compare GL with department's receipts issued to staff",
            'Identify and communicate any identified discrepancy to department staff',
        ]),
        ('Review of evidence of Quarterly Inspection of Directional Signs/Billboards', 'monthly', [
            'Request for the list of directional signs and billboards from Corporate Communications Department',
            'Review list provided by the Department.',
            'Request for report to confirm the quarterly visitation by the Department to inspect directional signs.',
            'Send list to branches to confirm the status of directional signs within their jurisdiction.',
            'Investigate if there are any discrepancies.',
        ]),
        ('Review of Reports (Events, Vendor, Sponsorship, CSR)', 'monthly', [
            'Request for reports on events, vendor, sponsorship, CSR',
            'Review to confirm availability and execution of assigned responsibilities',
            'Discuss if there are discrepancies',
        ]),
        ('Sighting of Evidence of Weekly Ambience Reports', 'weekly', [
            'Request for report from the weekly ambience portal',
            'Review to confirm if branches have submitted response',
            'Confirm if issues identified have been resolved',
        ]),
        ('Vendor List', 'monthly', [
            'Request for vendor list',
            'Review details i.e contact details of vendors',
            'Compare vendor list with expenses to confirm if all vendors have been captured as such.',
        ]),
    ],

    'CLEARING': [
        ('GL Review', 'daily', [
            'Check the position of the GL to confirm whether it has regular/irregular balances.',
            'Investigate reasons for GLs in irregular positions.',
        ]),
        ('Express Clearing Charges Review', 'daily', [
            "Confirm whether all express charges in line with the bank's tariff guide have been effected.",
            'Confirm whether express cheque charges have been effected on time.',
            'Confirm whether accounts with insufficient funds have been liened with the express charge.',
            'Obtain approval(s) from the clearing team for reversals effected.',
            'Obtain concession approval(s) from the clearing team for customers not charged.',
        ]),
        ('Monthly GL proof Review', 'monthly', [
            'Review the GL proofs submitted by Reconciliation department',
            'Check for correctness and accuracy of GL balances captured in the proof report as at close of the month by comparing with OBI',
            'Check for discrepancies/unexplained outstanding entries on proof.',
            'Check for correct preparation date captured in the proof.',
            'Respond with exceptions if any for correction and update of proof report.',
        ]),
    ],

    'CPU': [
        ('Review of Head Office Remittance GLs (RIA, Tranfast, Western Union, Moneygram, Unity Link)', 'daily', [
            'Check for outstanding settlement and follow up with remittance business team to ensure all delayed settlements are received from MTOs.',
            'Check if remittance transactions are settled per the approved SLA.',
            'Check that the correct narrations are captured in the GL',
            'Ensure that correct entries are passed regarding the settlement received: DR Corresponding Bank A/C (USD Aktif or GHIB); CR MTOs Head Office settlement GL (USD); DR MTOs Settlement GL (USD); CR Head Office settlement GL (GHS)',
        ]),
        ('Review of Salaries (For Internal & External Customers)', 'monthly', [
            'Confirm receipt of salaries schedule',
            'Check for timely processing of salaries by comparing the date of receipt and date of processing.',
            'Check that the request have been stamped as evidence of receipt',
            'Check that the request has been stamped with the signature verification stamp as evidence of signature verification',
            'Confirm that the signatures are consistent with the mandate in the CBA',
            'Ensure that all salary requests are properly filed',
            'Confirm that the amount debited from the ordering account is consistent with the instruction.',
            'Entries to be passed: Dr BOG in-house/Ordering customer account; Cr Beneficiary account/salary transit account for external accounts. Where charges apply, ensure that all charges are taken per the tariff guide.',
            "Check whether the commission for external accounts have been debited from the customer's account",
        ]),
        ('Review of Insurance Premium Collection (Enterprise, Milife, Hollard, StarLife)', 'monthly', [
            'Check if mandate from the insurance company is consistent with the CBA for the first time premium deduction',
            'Confirm receipt of premium deduction file for the month from the Insurance Company.',
            'Check for correctness of the data (customer name corresponds to the correct account number)',
            'Entries to be passed: Debit customer; Credit Insurance sundry creditors GL; Debit Insurance sundry creditors GL; Credit the insurance companies with the collection',
        ]),
        ('Other miscellaneous transactions (VISA Confirmation Charge, Balance Confirmation Charge, Any Other Charges)', 'daily', [
            "Confirm receipt of advice and supporting documents from Internal Audit, Internal Control, Branch Support or from the SBUs via email for processing",
            'Confirm customer email address is consistent with email on the indemnity form uploaded to the CBA',
            'Check for correct posting of the transactions',
            'Entries passed: Correct customer account is debited; correct commission GL credited',
            'Check if commission has been taken on the total premium amount collected for the month. Rate is per agreement with Insurance company (0.5-3%)',
            'Entries passed: Dr sundry creditors; Cr Insurance commission GL',
        ]),
        ('Callover', 'weekly', [
            'Check if callover report has been signed off by all stakeholders (maker, checker, Head of Department)',
            'Review all vouchers processed to detect and correct errors (wrong entries, wrong narrations, wrong amount etc)',
            'Confirm if instruments have been stamped and verified with the signature verification stamp and the processor received stamp.',
            'Confirm indemnity in the CBA for transactions sent via email.',
        ]),
        ('Monthly GL proof', 'monthly', [
            'Review the GL proofs submitted by Reconciliation department for CPU',
            'Check for correctness and accuracy of GL balances captured in the proof report as at close of the month by comparing with OBI report',
            'Check if details of outstanding transactions have been listed in the proof',
            'Check for discrepancies/unexplained outstanding entries on proof.',
            'Respond with exceptions if any for correction and update of proof report.',
        ]),
        ('GL review', 'daily', [
            "Review GLs used by the CPU on daily basis for transactions postings (trial balance)",
            'Check if transactions were posted to the appropriate GLs with the correct narrations',
            "Investigate all income GLs with negative balances",
            'Request for approvals for income reversal transactions in line with the expense policy.',
        ]),
    ],

    'CMU': [
        ('Cash Count', 'weekly', [
            'Check if the dual control process i.e. the required custodians and backups is being adhered to.',
            'Confirm from the vault key register to confirm the custodians/backups with the combination and key.',
            'Review to confirm if custodians record their name in the vault attendance register.',
            'Proceed to count the physical cash (denominations) of all the currencies (GHS, USD, GBP, XOF) at vault.',
            'Compare physical cash with system balance and vault summary sheet to confirm if they are balanced.',
            'Investigate if there is an overage or shortage.',
        ]),
        ('Cash Collection Charges', 'monthly', [
            'Request for cash collection schedule and SLAs from CMU.',
            'Compare with cash pick register to confirm the count of cash picks per customer.',
            "Check from the customer's account statement to confirm if they have been charged the right amount.",
            'If the customer has not been debited/charged the wrong amount; a mail should be sent to CMU to notify them.',
            'Review to confirm if the department has a tracker to identify customers who were not debited due to insufficient funds.',
            'Follow up to ensure that the exceptions identified are resolved.',
        ]),
        ('Review of Security Protocols (CCTV, Intruder Alert)', 'monthly', [
            'Visit the server room',
            'Record your details in the server room register',
            "Review register to confirm if there have been visits in line with the department's SOP",
            'Check DVR to confirm if footages are recording',
            'Check if DVR can record for more than 90 days',
            'Liaise with Physical Security Unit to confirm functionality of the Intruder Alert',
        ]),
        ('Review of Cash Deposits/Withdrawals at BOG', 'monthly', [
            'Review request form i.e. deposit/withdrawal',
            'Check for completeness i.e. appropriate signatories, amount in figures/words',
            'Confirm if the form has been duly authenticated',
            'Check GL to confirm if the deposit/withdrawal has been passed',
        ]),
    ],

    'BRANCH SUPPORT': [
        ('Reporting of missing instrument & Communication of missing instrument', 'monthly', [
            'Check register for letters on missing instrument from other Banks and OmniBSIC Bank.',
            'Check for prompt notification to branches on missing instruments.',
            'Confirm if notification was sent to other Banks on OmniBSIC missing instrument',
            'For OmniBSIC missing instruments, confirm if the instrument number is cancelled in the CBA and a new payment order issued.',
        ]),
        ('SPARE KEY MANAGEMENT', 'monthly', [
            'Confirm receipt of spare key memo from branches.',
            'Check if all spare keys are well labelled.',
            'Check if spare keys are recorded in the spare key register.',
            'Check if spare key register is duly completed.',
            'Confirm with Branch Control Officer if keys have been lodged at specified branches/zones',
        ]),
        ('Review of vault custodians sent to Branch Support', 'monthly', [
            'Confirm the existence of a folder or file containing details of vault key custodians, their backups from Branches as well as other keys in their possession.',
            'Check for evidence of update by Branch Support.',
        ]),
        ('Cheque book requisition', 'weekly', [
            'Review excel file generated from OBI (cheque book request for upload) for correctness and accuracy of requests submitted.',
            'Confirm if all details meet the request criteria; Leaflet for Corporate books are 50 or 100; personal books are 25 or 50 and savings booklet are 25 in ranges.',
            'Confirm sequence generated from cheque book application (Omni upload APP).',
            'Check if batch has been sent to the printing Company (camlotchequeordering@camlotprint.com) via corporate email with branch support in copy.',
            'Review the details received from the printing company and ensure the details are consistent with the details sent for printing.',
        ]),
        ('Receipt and Dispatch of cheque books', 'weekly', [
            'Confirm if the batch received has been recorded in the register',
            'Check if cheque request status has been changed to "generated" in the CBA.',
            'Reconcile cheque book dispatch forms sent to the branches with branch chequebook request to identify any inconsistency.',
            'Confirm if the Officer incharge of dispatch has signed the cheque book dispatch register.',
            'Check if Cheque books in possession of Branch Support has been kept in a fire proof cabinet under lock and key.',
            'Check for proof of signed off receipt of cheque book dispatch forms from branches.',
        ]),
        ('Purchasing & Reconciliation of the basestock', 'monthly', [
            'Confirm existence of memo approving the purchase of basestock.',
            'Confirm whether the basestock invoice from printing house is consistent with the approved memo',
            'Request for any outstanding base stock from the printing company.',
            'Reconcile basestock level from the printing company with the utilization report to confirm reorder level (2,500)',
            'Check if new base stock has been recorded in the register.',
            'Request for the consolidated cheque book report for the period and compare it with the customised cheque book report received from the printing company.',
            'Confirm whether cheque reconciliation schedule has been signed off by respective officers.',
        ]),
        ('Confirm the status of exited staff', 'monthly', [
            'To check in the CBA after HCM has notified ICDD about exit staff and ensure the account class has been changed from Staff account to current account',
        ]),
        ('Review Saturday Banking schedule', 'weekly', [
            'Review the Saturday banking schedule and ensure that only vault custodians and their designated backups open the vault.',
        ]),
    ],

    'CENTRAL ACCOUNT': [
        ('ACCOUNT OPEN', 'weekly', [
            'Request for account opening packs or documentation.',
            'Review for completeness. Ensure all information captured in the account pack have been accurately captured in the CBA.',
            'Confirm the availability of required documentation (refer to checklist at the back of account pack) and any other regulatory requirement as per the nature of business.',
            'Confirm initial deposit per tariff guide (Request for approval to open zero balance account in exceptional circumstances).',
            'Confirm if all products/services selected by the customer have been requested.',
        ]),
    ],

    'SBUs': [
        ('Balance Confirmation', 'daily', [
            'Review of balance confirmation requests sent by auditors.',
            'Review request to confirm if the signature has been verified and cross-check from the CBA.',
            'Prepare the draft and send to the RM/RO for confirmation',
            'After confirmation by the RM/RO, an instruction is sent to the Central Processing Unit for the required charges to be effected.',
            'The letter is printed, signed and forwarded to the auditors via mail or by address.',
        ]),
    ],

    'AMA&D': [
        ('ASSET MANAGEMENT', 'weekly', [
            'Maintain an accurate and up-to-date fixed asset register for all bank assets.',
            'Track asset location and custody.',
            'Conduct periodic physical verification of assets and reconcile results with the fixed asset register.',
        ]),
        ('ASSET TAGGING', 'monthly', [
            'Confirm if all assets have been tagged with unique identification numbers.',
            'Ensure asset tags are durable, tamper-resistant, and clearly visible.',
            'Ensure asset tags reflect the correct asset category and ownership.',
        ]),
        ('Review of auction result', 'monthly', [
            'Review auction results to ensure compliance with approved auction procedures and policies.',
            'Verify that auctioned assets match approved disposal lists and asset registers',
            'Review bid result via mail and auction reports for accuracy and completeness.',
            'Review documentation supporting the auction',
        ]),
    ],

    'ICDD - KORANTENG': [
        ('GL Proofs Review', 'monthly', [
            'Reconcile balances in the reports against OBI balances.',
            'Identify and escalate overaged, unrecognized or unexplained entries.',
            'Review adjustments or forced balancing entries passed by the schedule officer or department.',
            'Confirm whether the daily GL proofs & reconciliation reports are submitted on time.',
        ]),
        ('GL Review', 'daily', [
            'Conduct daily review of GL balances with the trial balance to unravel irregularities.',
            'Verify that all transactions posted to the GL have adequate source documents.',
            'Validate postings into GLs and identify any mis-postings.',
            'Verify the GL structure aligns with accounting standards and reporting requirements.',
            'Report findings and escalate any unresolved issues for further review.',
        ]),
        ('CALL OVER REPORT', 'weekly', [
            'Conduct a daily callover of all transactions processed for the previous day.',
        ]),
        ('Access Controls', 'monthly', [
            'Review user access lists for systems/platforms.',
            'Request for all users enrolled on the platform',
            'Check for users that are not supposed to be enrolled',
            'Check access type for users enrolled',
            'Ensure access is based on the principle of least privilege.',
            'Pay special attention to privileged/superuser accounts.',
        ]),
        ('Quality Assurance on Application acquisition or Development', 'weekly', [
            'Participate and review the acquisition and development of applications/systems and other related E-Business projects',
            "Participate in pre-deployment UATs to ensure adherence to bank's information security standards",
            'Provide update on the status of ongoing projects.',
        ]),
        ('Card Issuance and Maintenance Fees', 'monthly', [
            'Check if fees have been charged for all cards issued.',
            'Validate that fees charged align with approved tariff.',
            'Perform sample testing of issued cards vs fees posted.',
            'Identify unauthorized waivers or incorrect charges.',
            'Confirm charges are applied accurately and consistently.',
            'Review exceptions, reversals, and complaints.',
            'Check for waivers or exemptions applied and their approvals.',
        ]),
        ('Credit Card Contracts', 'monthly', [
            'Verify the total number of cards issued.',
            'Check cards issuance records against approved request forms.',
            'Verify the OD amount maintained on the card against the credit card account.',
            'Verify existence of signed agreements for all issued credit cards.',
            'Confirm if requirements (collateral) of the contracts have been met.',
            'Check for cancelled cards that have contracts still running.',
        ]),
        ('Chargeback Review', 'monthly', [
            'Verify adherence to scheme timelines and procedures (GhIPSS, Visa and MasterCard).',
            'Review supporting documentation for chargebacks.',
            'Track outstanding and unresolved chargebacks.',
            'Identify control gaps leading to financial losses and escalate.',
            'Verify that chargebacks are filed with the correct reason codes and that supporting documentation is complete.',
        ]),
        ('Validation of E-Business Reports', 'daily', [
            'Validate reports sent by E-Business to ensure reports are accurate.',
        ]),
        ('Card Load Review (Prepaid and Credit)', 'daily', [
            'Check settlement report from NI on card loads for the day.',
            'Match the report with the card load file that was sent to NI by the cards team.',
            'Identify card loads that were not part of the file which was sent to NI.',
        ]),
        ('Unclaimed Items GL Review', 'monthly', [
            'Review all entries passed into the unclaimed items GL and review all approvals granted IFO debit transactions',
        ]),
    ],

    'ICDD - MICHAEL': [
        ('POS Chargeback Review', 'monthly', [
            'Reconcile chargeback entries against dispute resolution documentation.',
            'Investigate reasons for chargebacks and identify preventable patterns.',
            'Follow up with POS team regarding chargeback claims.',
            'Ensure chargeback decisions align with scheme rules and timelines.',
            'Monitor outstanding chargebacks and track the resolution.',
        ]),
        ('POS Inventory Review', 'monthly', [
            'Reconcile physical POS terminals with deployment and inventory records.',
            'Verify procurement and disposal procedures for POS terminals.',
            'Check proper storage and tracking of POS stock awaiting deployment.',
        ]),
        ('Merchant Category Code Review', 'monthly', [
            'Spool records of all POS merchants that have been set up in Postilion.',
            'Check Merchant MCC assigned on Postilion against merchant nature of business stated in the Merchant onboarding application form.',
            'Confirm if the right MCC has been assigned using the Visa Merchant Categorization report as guide.',
            'Report any misalignment of MCC for corrective actions to be taken.',
            'Record findings in the exception report.',
        ]),
        ('Mobile App Security Answer & Freedom Pin Reset', 'monthly', [
            'Spool report for pin resets that was done for customers/users.',
            'For each treated pin reset, request for evidence of requests from customers.',
        ]),
        ('Push And Pull Setup Review', 'monthly', [
            'Spool for records of MTN And Telecel PAP setup performed by bank admins.',
            'For each setup, request for evidence of requests from customers',
        ]),
        ('Access Controls', 'monthly', [
            'Review user access lists for systems/platforms.',
            'Request for all users enrolled on the platform',
            'Check for users that are not supposed to be enrolled',
            'Check access type for users enrolled',
            'Ensure access is based on the principle of least privilege.',
            'Pay special attention to privileged/superuser accounts.',
        ]),
        ('POS Deployment', 'monthly', [
            'Obtain the records of total number of POS terminals deployed.',
            'Check terminal deployment records against merchant agreements.',
            'Confirm deduction of setup fee.',
            'Confirm merchant details (KYC compliance) for each deployed POS.',
            'Validate transaction volumes and ensure the proper application of merchant commission rates.',
            'Verify if there is an approved concession memo for merchant commission rates.',
        ]),
        ('POS Merchant Performance Review', 'monthly', [
            'Review transaction volume per merchant to identify high-performing and low-performing merchants.',
            'Analyze chargeback trends and their impact on merchant performance.',
            'Generate performance reports and rank merchants based on transaction value and volume',
            'Recommend improvements or terminations for underperforming merchant',
        ]),
        ('Quality Assurance on Application acquisition or Development', 'weekly', [
            'Participate and review the acquisition and development of applications/systems and other related E-Business projects',
            "Participate in pre-deployment UATs to ensure adherence to bank's information security standards",
            'Provide update on the status of ongoing projects.',
        ]),
        ('Telcos Wallet Balance Review', 'daily', [
            'Login to the Telco portal to confirm the daily balances',
            'Check corresponding GL balance for Telco on OBI',
            'Record balances in the Telco Wallet Balances Report',
            'Report imbalances noted via mail to responsible unit to obtain reason for the imbalance',
            'Capture in Exceptions Report if necessary',
        ]),
        ('Events Collection Review', 'monthly', [
            'Log in to the collections portal and spool the report for the review period',
            'Confirm signed agreement exist for all collections done',
            'Confirm amount credited is consistent with amount captured on the collections portal',
            'Confirm appropriate commission has been charged',
            'Capture in Exceptions Report if necessary',
        ]),
        ('Adherence to Card Limits and Issuance Directive', 'monthly', [
            'Test compliance with approved limits and issuance criteria.',
            'Identify and investigate overrides or unauthorized limit assignments.',
            'Ensure approvals are properly documented.',
            'Verify requests are supported by appropriate documentation.',
            'Confirm approval hierarchy is adhered to.',
            'Check that system updates reflect approved limits only.',
        ]),
        ('POS Transaction Review', 'daily', [
            'Maintain a schedule of POS transactions daily',
            'Review merchant activity based on transaction history',
            'Identify instances of a single card performing multiple transactions, high transaction amounts for a single card.',
            'Send mails for proof of payment to Alternate Channels.',
        ]),
        ('Accounts With Negative Balances', 'daily', [
            'Spool report from OBIE on Accounts with Negative Balances as a result of transactions processed to identify all overdrawn accounts.',
            'Escalate overdrawn accounts to the department that processed the transactions',
            'Escalate to IT if the transactions were processed by the system.',
        ]),
    ],

    'FACILITIES': [
        ('MAINTENANCE & REPAIRS', 'monthly', [
            'Review scheduled maintenance for Fire panels',
            'Review scheduled maintenance for Smoke Detectors',
            'Review scheduled maintenance for Panic alarms',
            'Review scheduled maintenance for Fire extinguishers',
            'Review work order approvals and documentation for repairs',
            'Confirm painting of branch according to painting schedule',
        ]),
        ('CERTIFICATIONS', 'monthly', [
            'Review Fire Certificates bankwide',
            'Review Business Operating Permit bankwide',
        ]),
        ('PROJECTS', 'monthly', [
            'Check for approval memos projects',
            'Review of Lease agreement & Lease payment',
            'Project completion report',
        ]),
    ],
}


class Command(BaseCommand):
    help = 'Seed Head Office activity checklists grouped by Department.'

    def add_arguments(self, parser):
        parser.add_argument('--team', default=None,
                            help='Seed only this team (case-insensitive).')
        parser.add_argument('--dry-run', action='store_true')
        parser.add_argument('--reset', action='store_true')

    def handle(self, *args, **options):
        dry_run = options['dry_run']
        reset = options['reset']
        team_filter = (options['team'] or '').strip().upper() or None

        teams = sorted(TEAM_DATA.keys())
        if team_filter:
            teams = [t for t in teams if t.upper() == team_filter]
            if not teams:
                self.stdout.write(self.style.ERROR(
                    f'No team matches "{team_filter}".'
                ))
                return

        total_activities = sum(len(TEAM_DATA[t]) for t in teams)

        self.stdout.write(self.style.MIGRATE_HEADING(
            f'\nSeeding {total_activities} department checklists '
            f'across {len(teams)} team(s)'
        ))
        if dry_run:
            self.stdout.write(self.style.WARNING('DRY RUN — nothing will be saved.'))
        self.stdout.write('')

        created = updated = task_total = 0
        dept_map = {}
        empty_depts = []

        with transaction.atomic():
            if reset and not dry_run:
                existing = Checklist.objects.filter(
                    assigned_departments__isnull=False
                ).distinct()
                n = existing.count()
                existing.delete()
                self.stdout.write(self.style.WARNING(
                    f'Deleted {n} existing department-scoped checklist(s).\n'
                ))

            counter = 0
            for team in teams:
                dept, _ = Department.objects.get_or_create(
                    name=team,
                    defaults={'is_active': True},
                )
                dept_map[team] = dept

                dept_users = UserProfile.objects.filter(
                    departments=dept, status='active'
                )
                if dept_users.count() == 0:
                    empty_depts.append(team)

                for activity, freq, tasks in TEAM_DATA[team]:
                    counter += 1
                    full_name = f'[{team}] {activity}'[:200]
                    freq = norm_freq(freq)

                    checklist, was_created = Checklist.objects.get_or_create(
                        name=full_name,
                        defaults={
                            'description': f'{team} — Head Office activity',
                            'frequency': freq,
                            'assignment_target': 'specific',
                            'is_active': True,
                        },
                    )

                    if was_created:
                        created += 1
                        mark = '✅ Created'
                    else:
                        updated += 1
                        mark = '♻️  Updated'

                    self.stdout.write(
                        f'  [{counter:03d}] {mark}  {full_name}  '
                        f'({freq}, {len(tasks)} tasks)'
                    )

                    if dry_run:
                        continue

                    checklist.frequency = freq
                    checklist.assignment_target = 'specific'
                    checklist.is_active = True
                    checklist.save()

                    checklist.assigned_departments.set([dept])
                    checklist.assigned_users.set(dept_users)

                    ChecklistTask.objects.filter(checklist=checklist).delete()
                    for order, desc in enumerate(tasks):
                        ChecklistTask.objects.create(
                            checklist=checklist,
                            description=desc[:500],
                            order=order,
                            is_completed=False,
                        )
                    task_total += len(tasks)

        # ---- Summary ----
        self.stdout.write('')
        self.stdout.write(self.style.SUCCESS('─' * 70))
        self.stdout.write(self.style.SUCCESS('  DONE'))
        self.stdout.write(self.style.SUCCESS('─' * 70))
        self.stdout.write(f'  Checklists created: {created}')
        self.stdout.write(f'  Checklists updated: {updated}')
        self.stdout.write(f'  Tasks written:      {task_total}')
        self.stdout.write(f'  Departments:        {len(teams)}')

        if empty_depts:
            self.stdout.write('')
            self.stdout.write(self.style.WARNING(
                '⚠️  These departments have NO active users yet:'
            ))
            for t in sorted(empty_depts):
                self.stdout.write(f'    • {t}')
            self.stdout.write(self.style.WARNING(
                '   Assign users in /admin/control_dashboard/userprofile/ '
                'so the checklists appear on their Activity Checklist page.'
            ))

        if dry_run:
            self.stdout.write(self.style.WARNING('\n(DRY RUN — nothing saved)'))