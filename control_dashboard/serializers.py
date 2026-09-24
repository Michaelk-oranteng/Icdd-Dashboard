from rest_framework import serializers
from .models import User

class UserSerializer(serializers.ModelSerializer):
    class Meta:
        model = User
        fields = ['id', 'email', 'full_name', 'position', 'role', 'status', 'created_at']
    
    def validate_email(self, value):
        if User.objects.filter(email=value).exists():
            raise serializers.ValidationError("A user with this email already exists.")
        return value

class UserStatusSerializer(serializers.Serializer):
    status = serializers.ChoiceField(choices=User.STATUS_CHOICES)