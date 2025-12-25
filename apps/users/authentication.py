from rest_framework_simplejwt.authentication import JWTAuthentication
from rest_framework_simplejwt.exceptions import InvalidToken
from apps.users.models import User


class CustomJWTAuthentication(JWTAuthentication):
    def get_user(self, validated_token):
        try:
            user_id = validated_token.get('user_id')
            user = User.objects.get(id=user_id, valid=True)
            return user
        except User.DoesNotExist:
            raise InvalidToken('User not found')
