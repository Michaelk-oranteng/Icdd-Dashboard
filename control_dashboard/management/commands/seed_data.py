# control_dashboard/management/commands/seed_data.py
from django.core.management.base import BaseCommand
from control_dashboard.models import Branch, Department

class Command(BaseCommand):
    help = 'Seed branches and departments data into the database'

    def handle(self, *args, **options):
        # Branch data
        branches = [
            'ATOMIC', 'DANSOMAN', 'DOME', 'EAST LEGON', 'KASOA', 'KEJETIA',
            'KNUST', 'KOFORIDUA', 'KOKOMLEMLE', 'KRONUM', 'LABONE', 'MADINA ESTATE',
            'MANHYIA', 'NIMA', 'NORTH INDUSTRIAL AREA', 'ODORKOR', 'OSU',
            'SPINTEX BASKET', 'SPINTEX MANET', 'SUNYANI', 'TAKORADI HARBOUR',
            'TAKORADI MARKET CIRCLE', 'TAMALE', 'TARKWA', 'TECHIMAN',
            'TEMA COMMUNITY 1', 'TEMA COMMUNITY 11', 'TEMA EAST', 'TEMA HARBOUR', 'WEIJA'
        ]

        # Department data
        departments = [
            'GENERAL SERVICES & PROCUREMENT', 'BRANCH SUPPORT', 'CLEARING', 'FINOPS',
            'CORPORATE COMMS', 'CREDIT', 'RISK', 'E-BUSINESS', 'FACILITIES MANAGEMENT',
            'FINANCE', 'CPU', 'HCM', 'LEGAL', 'RECONCILIATION', 'SERVICE QUALITY',
            'TRADE SERVICES', 'CMU', 'TREASURY', 'REMITTANCE', 'IT',
            'TRANSPORT & LOGISTICS', 'RECOVERY', 'CORPORATE GROUP', 'PBB',
            'CENTRALIZED ACCOUNT OPENING', 'TROPS'
        ]

        self.stdout.write(self.style.SUCCESS('=' * 60))
        self.stdout.write(self.style.SUCCESS('SEEDING BRANCHES AND DEPARTMENTS'))
        self.stdout.write(self.style.SUCCESS('=' * 60))

        # Create branches
        self.stdout.write('\n📌 Creating branches...')
        branches_created = 0
        branches_existing = 0

        for branch_name in branches:
            branch, created = Branch.objects.get_or_create(
                name=branch_name,
                defaults={'is_active': True}
            )
            if created:
                branches_created += 1
                self.stdout.write(self.style.SUCCESS(f'  ✅ Created: {branch_name}'))
            else:
                branches_existing += 1
                self.stdout.write(f'  ⏭️  Already exists: {branch_name}')

        # Create departments
        self.stdout.write('\n📌 Creating departments...')
        departments_created = 0
        departments_existing = 0

        for dept_name in departments:
            dept, created = Department.objects.get_or_create(
                name=dept_name,
                defaults={'is_active': True}
            )
            if created:
                departments_created += 1
                self.stdout.write(self.style.SUCCESS(f'  ✅ Created: {dept_name}'))
            else:
                departments_existing += 1
                self.stdout.write(f'  ⏭️  Already exists: {dept_name}')

        self.stdout.write('\n' + '=' * 60)
        self.stdout.write(self.style.SUCCESS('✅ SEEDING COMPLETE!'))
        self.stdout.write('=' * 60)
        self.stdout.write(f'📊 Total Branches: {Branch.objects.count()}')
        self.stdout.write(f'   - Newly created: {branches_created}')
        self.stdout.write(f'   - Already existed: {branches_existing}')
        self.stdout.write(f'\n📊 Total Departments: {Department.objects.count()}')
        self.stdout.write(f'   - Newly created: {departments_created}')
        self.stdout.write(f'   - Already existed: {departments_existing}')
        self.stdout.write('=' * 60)

        # Show all branches
        self.stdout.write('\n📋 All Branches:')
        for branch in Branch.objects.all().order_by('name'):
            self.stdout.write(f'  • {branch.name}')

        self.stdout.write('\n📋 All Departments:')
        for dept in Department.objects.all().order_by('name'):
            self.stdout.write(f'  • {dept.name}')