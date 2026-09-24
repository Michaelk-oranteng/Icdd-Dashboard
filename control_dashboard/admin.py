# control_dashboard/admin.py

from django.contrib import admin
from django.contrib import messages
from .models import UserProfile, Checklist, ChecklistTask, ChecklistLog
from import_export import resources, fields
from import_export.admin import ImportExportModelAdmin
from import_export.formats.base_formats import CSV
import re


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
        
        # Clean activity name
        if 'ACTIVITY' in row:
            row['ACTIVITY'] = str(row['ACTIVITY']).strip()
            row['ACTIVITY'] = ' '.join(row['ACTIVITY'].split())
        
        # Clean description - keep the full text
        if 'TASK / DESCRIPTION' in row:
            desc = str(row['TASK / DESCRIPTION'])
            desc = desc.replace('•', '').strip()
            desc = desc.replace('\n', ' ').strip()
            desc = desc.replace('\r', ' ').strip()
            desc = desc.replace('"', '').strip()
            desc = ' '.join(desc.split())
            row['TASK / DESCRIPTION'] = desc
        
        # Map frequency values
        if 'FREQUENCY' in row:
            freq = str(row['FREQUENCY']).strip().lower()
            freq_map = {
                'daily': 'daily',
                'weekly': 'weekly',
                'monthly': 'monthly',
                'quarterly': 'quarterly',
            }
            row['FREQUENCY'] = freq_map.get(freq, 'weekly')
        
        # Map assignment target values
        if 'ASSIGNMENT TARGET' in row:
            target = str(row['ASSIGNMENT TARGET']).strip().lower()
            target_map = {
                'cluster control': 'cc',
                'cc': 'cc',
                'head office control': 'hc',
                'hc': 'hc',
                'all users': 'all',
                'all': 'all',
            }
            row['ASSIGNMENT TARGET'] = target_map.get(target, 'all')
        
        return row
    
    def after_import_row(self, row, row_result, **kwargs):
        """After each row is imported, create tasks and assign users"""
        if not row_result.errors:
            try:
                checklist = Checklist.objects.get(name=row['ACTIVITY'])
                
                # 1. Set assignment target
                assignment_target = row.get('ASSIGNMENT TARGET', 'all')
                checklist.assignment_target = assignment_target
                checklist.is_active = True
                checklist.save()
                
                # 2. Create tasks from the description
                description = row.get('TASK / DESCRIPTION', '')
                if description:
                    # Split into individual tasks
                    tasks = self.split_into_tasks(description)
                    
                    # Clear existing tasks
                    ChecklistTask.objects.filter(checklist=checklist).delete()
                    
                    # Create new tasks
                    for order, task_text in enumerate(tasks, start=1):
                        ChecklistTask.objects.create(
                            checklist=checklist,
                            description=task_text.strip(),
                            order=order,
                            is_completed=False
                        )
                    print(f"✅ Created {len(tasks)} tasks for: {checklist.name}")
                
                # 3. Assign users based on target
                if assignment_target == 'cc':
                    cc_users = UserProfile.objects.filter(
                        role='cc'
                    ) | UserProfile.objects.filter(
                        position__icontains='Cluster Control'
                    )
                    if cc_users.exists():
                        checklist.assigned_users.set(cc_users)
                        print(f"✅ Assigned {cc_users.count()} CC users to: {checklist.name}")
                    else:
                        print(f"⚠️ No CC users found for: {checklist.name}")
                elif assignment_target == 'all':
                    all_users = UserProfile.objects.filter(is_active=True)
                    checklist.assigned_users.set(all_users)
                    print(f"✅ Assigned {all_users.count()} users to: {checklist.name}")
                
            except Checklist.DoesNotExist:
                print(f"❌ Checklist not found: {row.get('ACTIVITY')}")
            except Exception as e:
                print(f"❌ Error: {e}")
    
    def split_into_tasks(self, text):
        """Split description into individual tasks"""
        # Try splitting by bullet points
        if '•' in text:
            tasks = [t.strip() for t in text.split('•') if t.strip()]
            return tasks
        
        # Try splitting by numbered items
        if re.search(r'\d+\.', text):
            tasks = re.split(r'\d+\.\s*', text)
            tasks = [t.strip() for t in tasks if t.strip()]
            return tasks
        
        # Try splitting by periods (sentences)
        if '.' in text:
            tasks = [t.strip() + '.' for t in text.split('.') if t.strip()]
            return tasks
        
        # If nothing works, return as single task
        return [text]


# ==================== CHECKLIST TASK INLINE (Define BEFORE ChecklistAdmin) ====================

class ChecklistTaskInline(admin.TabularInline):
    """Inline admin for ChecklistTask - shows tasks inside Checklist"""
    model = ChecklistTask
    extra = 0
    fields = ['description', 'order', 'is_completed']
    ordering = ['order']
    readonly_fields = ['created_at', 'updated_at']


# ==================== CHECKLIST ADMIN ====================

@admin.register(Checklist)
class ChecklistAdmin(ImportExportModelAdmin):
    resource_class = ChecklistResource
    
    list_display = ['name', 'get_frequency_display', 'get_assignment_display', 'is_active', 'created_at']
    list_filter = ['frequency', 'assignment_target', 'is_active']
    search_fields = ['name', 'description']
    readonly_fields = ['created_at', 'updated_at']
    fields = ['name', 'description', 'frequency', 'assignment_target', 'assigned_users', 'is_active']
    inlines = [ChecklistTaskInline]  # Now this is defined!
    
    def get_import_formats(self):
        from import_export.formats.base_formats import CSV
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
            messages.error(request, f"❌ Import failed: {str(e)}")
            raise


# ==================== CHECKLIST TASK ADMIN ====================

@admin.register(ChecklistTask)
class ChecklistTaskAdmin(admin.ModelAdmin):
    list_display = ['checklist', 'description', 'order', 'is_completed']
    list_filter = ['checklist', 'is_completed']
    search_fields = ['description']


# ==================== CHECKLIST LOG ADMIN ====================

@admin.register(ChecklistLog)
class ChecklistLogAdmin(admin.ModelAdmin):
    list_display = ['checklist', 'user', 'log_date', 'created_at']
    list_filter = ['checklist', 'user']
    search_fields = ['checklist__name', 'user__full_name']


# control_dashboard/admin.py

from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User
from .models import UserProfile

class UserProfileInline(admin.StackedInline):
    model = UserProfile
    can_delete = False
    verbose_name_plural = 'Profile'

class CustomUserAdmin(UserAdmin):
    inlines = (UserProfileInline,)
    list_display = ('username', 'email', 'first_name', 'last_name', 'is_staff')
    search_fields = ('username', 'email', 'first_name', 'last_name')

admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)

@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'email', 'username', 'position', 'role', 'status')
    search_fields = ('full_name', 'email', 'username')
    list_filter = ('position', 'role', 'status')
    readonly_fields = ('username', 'created_at', 'updated_at')