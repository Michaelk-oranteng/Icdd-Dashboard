from django import forms
from .models import UserProfile


class UserProfileForm(forms.ModelForm):
    """Form for creating and editing users."""

    class Meta:
        model = UserProfile
        fields = [
            'email',
            'username',
            'full_name',
            'avatar',
            'position',
            'role',
            'status',
        ]
        widgets = {
            'email': forms.EmailInput(attrs={
                'class': 'form-control',
                'placeholder': 'Enter email address...',
            }),
            'username': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Enter username...',
                'autocomplete': 'off',
            }),
            'full_name': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Enter full name...',
            }),
            'avatar': forms.ClearableFileInput(attrs={
                'class': 'form-control',
                'accept': 'image/*',
            }),
            'position': forms.Select(attrs={'class': 'form-control'}),
            'role': forms.Select(attrs={'class': 'form-control'}),
            'status': forms.Select(attrs={'class': 'form-control'}),
        }

    def clean_email(self):
        email = self.cleaned_data.get('email')
        if email:
            existing = UserProfile.objects.filter(email__iexact=email)
            if self.instance.pk:
                existing = existing.exclude(pk=self.instance.pk)
            if existing.exists():
                raise forms.ValidationError('A user with this email already exists.')
        return email.lower()

    def clean_username(self):
        username = self.cleaned_data.get('username')
        if not username:
            raise forms.ValidationError('Username is required.')
        username = username.strip()
        existing = UserProfile.objects.filter(username__iexact=username)
        if self.instance.pk:
            existing = existing.exclude(pk=self.instance.pk)
        if existing.exists():
            raise forms.ValidationError('This username is already taken.')
        return username

    def clean_avatar(self):
        avatar = self.cleaned_data.get('avatar')
        if avatar and hasattr(avatar, 'size'):
            # 2 MB limit
            if avatar.size > 2 * 1024 * 1024:
                raise forms.ValidationError('Image must be smaller than 2 MB.')
            # Basic content-type check
            content_type = getattr(avatar, 'content_type', '')
            if content_type and not content_type.startswith('image/'):
                raise forms.ValidationError('File must be an image.')
        return avatar