# control_dashboard/forms.py

from django import forms
from .models import UserProfile


class UserProfileForm(forms.ModelForm):
    """Form for editing existing users."""
    
    class Meta:
        model = UserProfile
        fields = ['email', 'full_name', 'position', 'role', 'status']
        widgets = {
            'email': forms.EmailInput(attrs={'class': 'form-control', 'placeholder': 'Enter email address...'}),
            'full_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Enter full name...'}),
            'position': forms.Select(attrs={'class': 'form-control'}),
            'role': forms.Select(attrs={'class': 'form-control'}),
            'status': forms.Select(attrs={'class': 'form-control'}),
        }
    
    def clean_email(self):
        email = self.cleaned_data.get('email')
        if email:
            # Check for existing email except current instance
            existing = UserProfile.objects.filter(email=email)
            if self.instance.pk:
                existing = existing.exclude(pk=self.instance.pk)
            if existing.exists():
                raise forms.ValidationError('A user with this email already exists.')
        return email