# control_dashboard/management/commands/import_checklists.py

import csv
import re
from django.core.management.base import BaseCommand
from django.db import transaction
from control_dashboard.models import Checklist, ChecklistTask, UserProfile


class Command(BaseCommand):
    help = 'Import checklists from CSV file with tasks'

    def add_arguments(self, parser):
        parser.add_argument('csv_file', type=str, help='Path to CSV file')

    def handle(self, *args, **options):
        csv_file = options['csv_file']
        
        self.stdout.write(f"📂 Reading CSV file: {csv_file}")
        
        # Get Cluster Control users (role='cc')
        cc_users = UserProfile.objects.filter(role='cc')
        if not cc_users.exists():
            self.stdout.write(self.style.WARNING("⚠️ No Cluster Control users found. Using all users with position='cc'"))
            # Try position field instead
            cc_users = UserProfile.objects.filter(position__icontains='cc')
            if not cc_users.exists():
                self.stdout.write(self.style.WARNING("⚠️ No users found. Please create some users first."))
                return
        
        self.stdout.write(f"✅ Found {cc_users.count()} Cluster Control users")
        for user in cc_users:
            self.stdout.write(f"   - {user.full_name} ({user.email})")
        
        success_count = 0
        error_count = 0
        task_count = 0
        
        with transaction.atomic():
            with open(csv_file, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                
                for row_num, row in enumerate(reader, start=2):
                    try:
                        # Get data
                        activity = row['ACTIVITY'].strip()
                        description = row['TASK / DESCRIPTION'].strip()
                        frequency = row['FREQUENCY'].strip().lower()
                        assignment = row['ASSIGNMENT TARGET'].strip().lower()
                        
                        if not activity:
                            self.stdout.write(self.style.WARNING(f"⚠️ Skipping empty row {row_num}"))
                            continue
                        
                        # Clean description - remove bullet points
                        description = description.replace('•', '').strip()
                        description = ' '.join(description.split())
                        
                        # Map frequency
                        freq_map = {
                            'daily': 'daily',
                            'weekly': 'weekly',
                            'monthly': 'monthly',
                            'quarterly': 'quarterly',
                        }
                        frequency = freq_map.get(frequency, 'weekly')
                        
                        # Map assignment target
                        target_map = {
                            'cluster control': 'cc',
                            'cc': 'cc',
                            'head office control': 'hc',
                            'hc': 'hc',
                            'all users': 'all',
                            'all': 'all',
                        }
                        assignment_target = target_map.get(assignment, 'all')
                        
                        # Create or update checklist
                        checklist, created = Checklist.objects.update_or_create(
                            name=activity,
                            defaults={
                                'description': description[:500],
                                'frequency': frequency,
                                'assignment_target': assignment_target,
                            }
                        )
                        
                        # Split description into tasks
                        # Split by bullet points (if any remain)
                        if '•' in description:
                            tasks = [t.strip() for t in description.split('•') if t.strip()]
                        else:
                            # Split by periods for sentences
                            tasks = [t.strip() + '.' for t in description.split('.') if t.strip()]
                        
                        # If no tasks found, use the whole description
                        if not tasks:
                            tasks = [description]
                        
                        # Clear existing tasks
                        ChecklistTask.objects.filter(checklist=checklist).delete()
                        
                        # Create new tasks
                        for order, task_text in enumerate(tasks, start=1):
                            ChecklistTask.objects.create(
                                checklist=checklist,
                                description=task_text[:500],
                                order=order,
                                is_completed=False
                            )
                            task_count += 1
                        
                        # Assign users based on target
                        if assignment_target == 'cc':
                            checklist.assigned_users.set(cc_users)
                        elif assignment_target == 'all':
                            # Get all users (using status='active' if available, or all users)
                            all_users = UserProfile.objects.all()
                            checklist.assigned_users.set(all_users)
                        elif assignment_target == 'hc':
                            hc_users = UserProfile.objects.filter(role='hc')
                            if hc_users.exists():
                                checklist.assigned_users.set(hc_users)
                        
                        success_count += 1
                        status = "✅ Created" if created else "🔄 Updated"
                        self.stdout.write(f"{status}: {activity} ({len(tasks)} tasks)")
                        
                    except KeyError as e:
                        error_count += 1
                        self.stdout.write(self.style.ERROR(f"❌ Row {row_num}: Missing column {e}"))
                    except Exception as e:
                        error_count += 1
                        self.stdout.write(self.style.ERROR(f"❌ Row {row_num}: {str(e)}"))
                        # Don't raise - continue with next row
        
        self.stdout.write(self.style.SUCCESS(f"\n🎉 Import complete!"))
        self.stdout.write(self.style.SUCCESS(f"   - {success_count} checklists imported"))
        self.stdout.write(self.style.SUCCESS(f"   - {task_count} tasks created"))
        self.stdout.write(self.style.SUCCESS(f"   - {error_count} errors"))
        
        # Verify the import
        total_checklists = Checklist.objects.count()
        total_tasks = ChecklistTask.objects.count()
        self.stdout.write(self.style.SUCCESS(f"\n📊 Database totals:"))
        self.stdout.write(self.style.SUCCESS(f"   - Total checklists: {total_checklists}"))
        self.stdout.write(self.style.SUCCESS(f"   - Total tasks: {total_tasks}"))