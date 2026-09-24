from django.core.management.base import BaseCommand
from django.contrib.auth.models import User as DjangoUser
from control_dashboard.models import UserProfile
import re

class Command(BaseCommand):
    help = 'Sync UserProfile with Django User for authentication'

    def handle(self, *args, **options):
        self.stdout.write("\n=== SYNCING USERS ===\n")
        
        profiles = UserProfile.objects.all()
        total = profiles.count()
        created = 0
        updated = 0
        
        self.stdout.write(f"Found {total} UserProfiles\n")
        
        for profile in profiles:
            # Try to find Django user by email
            django_user = None
            try:
                django_user = DjangoUser.objects.get(email=profile.email)
            except DjangoUser.DoesNotExist:
                pass
            
            # If not found by email, try by username
            if not django_user and profile.username:
                try:
                    django_user = DjangoUser.objects.get(username=profile.username)
                except DjangoUser.DoesNotExist:
                    pass
            
            if django_user:
                # Update existing Django user
                needs_update = False
                
                if profile.username and django_user.username != profile.username:
                    # Check if username is available
                    if not DjangoUser.objects.filter(username=profile.username).exclude(id=django_user.id).exists():
                        django_user.username = profile.username
                        needs_update = True
                        self.stdout.write(f"  🔄 Updating username: {django_user.username} -> {profile.username}")
                
                if django_user.email != profile.email:
                    django_user.email = profile.email
                    needs_update = True
                    self.stdout.write(f"  🔄 Updating email: {django_user.email} -> {profile.email}")
                
                if needs_update:
                    django_user.save()
                    updated += 1
                    self.stdout.write(f"  ✅ Updated Django user: {django_user.username}")
                else:
                    self.stdout.write(f"  ℹ️ Django user already synced: {django_user.username}")
            else:
                # Create new Django user
                username = profile.username or profile.email.split('@')[0]
                
                # Make sure username is unique
                original_username = username
                counter = 1
                while DjangoUser.objects.filter(username=username).exists():
                    username = f"{original_username}{counter}"
                    counter += 1
                
                django_user = DjangoUser.objects.create_user(
                    username=username,
                    email=profile.email,
                    password='defaultpassword123'
                )
                
                # Set full name
                if profile.full_name:
                    name_parts = profile.full_name.split(' ', 1)
                    django_user.first_name = name_parts[0]
                    django_user.last_name = name_parts[1] if len(name_parts) > 1 else ''
                    django_user.save()
                
                created += 1
                self.stdout.write(f"  ✅ Created Django user: {username} for {profile.email}")
        
        self.stdout.write("\n" + "=" * 50)
        self.stdout.write(self.style.SUCCESS(f"✅ Sync complete!"))
        self.stdout.write(f"   Created: {created}")
        self.stdout.write(f"   Updated: {updated}")
        self.stdout.write(f"   Total: {total}")
        
        # Show all Django users
        self.stdout.write("\n=== ALL DJANGO USERS ===")
        self.stdout.write("-" * 70)
        self.stdout.write(f"{'ID':<5} {'Username':<25} {'Email':<35}")
        self.stdout.write("-" * 70)
        
        for du in DjangoUser.objects.all().order_by('id'):
            self.stdout.write(
                f"{du.id:<5} {du.username:<25} {du.email:<35}"
            )
        
        self.stdout.write("-" * 70)