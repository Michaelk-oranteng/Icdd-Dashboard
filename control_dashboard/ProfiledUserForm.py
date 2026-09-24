from django import forms
from .models import ProfiledUser

class ProfiledUserForm(forms.ModelForm):
    class Meta:
        model = ProfiledUser
        fields = [
            'email', 'full_name', 'position', 'role', 'status',
            'branch_code', 'department', 'employee_id'
        ]
        widgets = {
            'email': forms.EmailInput(attrs={'class': 'form-control', 'placeholder': 'Enter email address'}),
            'full_name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Enter full name'}),
            'position': forms.Select(attrs={'class': 'form-control'}),
            'role': forms.Select(attrs={'class': 'form-control'}),
            'status': forms.Select(attrs={'class': 'form-control'}),
            'branch_code': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Branch code'}),
            'department': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Department'}),
            'employee_id': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Employee ID'}),
        }