from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import AllowAny
from rest_framework.response import Response
from rest_framework_simplejwt.tokens import RefreshToken
from django.contrib.auth.hashers import make_password, check_password
from .models import User
from .serializers import UserRegistrationSerializer, UserLoginSerializer


@api_view(['POST'])
@permission_classes([AllowAny])
def register(request):
    serializer = UserRegistrationSerializer(data=request.data)
    if serializer.is_valid():
        user_data = serializer.validated_data
        user_data['password'] = make_password(user_data['password'])
        user = User.objects.create(**user_data)
        
        refresh = RefreshToken.for_user(user)
        refresh['user_id'] = user.id
        
        return Response({
            'message': 'User registered successfully',
            'user_id': user.id,
            'access_token': str(refresh.access_token),
            'refresh_token': str(refresh)
        }, status=status.HTTP_201_CREATED)
    
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
@permission_classes([AllowAny])
def login(request):
    from decouple import config
    from apps.organizations.models import Organization
    
    serializer = UserLoginSerializer(data=request.data)
    if serializer.is_valid():
        username = serializer.validated_data['username']
        password = serializer.validated_data['password']
        org_name = serializer.validated_data.get('org_name')
        
        # Get org_name from env if not provided
        if not org_name:
            org_name = config('DEFAULT_ORG', default='default')
        
        try:
            # Get organization
            org = Organization.objects.get(org_name=org_name)
            
            # Get user with matching username, org, and valid status
            user = User.objects.get(username=username, org=org, valid=True)
            
            if check_password(password, user.password):
                refresh = RefreshToken.for_user(user)
                refresh['user_id'] = user.id
                
                return Response({
                    'message': 'Login successful',
                    'user_id': user.id,
                    'access_token': str(refresh.access_token),
                    'refresh_token': str(refresh)
                }, status=status.HTTP_200_OK)
            else:
                return Response({'error': 'Invalid credentials'}, status=status.HTTP_401_UNAUTHORIZED)
        except (User.DoesNotExist, Organization.DoesNotExist):
            return Response({'error': 'Invalid credentials'}, status=status.HTTP_401_UNAUTHORIZED)
    
    return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)