from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import AllowAny, IsAuthenticated
from rest_framework.response import Response
from rest_framework_simplejwt.tokens import RefreshToken
from django.contrib.auth.hashers import make_password, check_password
from django.db.models import Q
from .models import User
from .serializers import UserRegistrationSerializer, UserLoginSerializer


@api_view(['POST'])
@permission_classes([AllowAny])
def register(request):
    serializer = UserRegistrationSerializer(data=request.data)
    if serializer.is_valid():
        user_data = serializer.validated_data
        user_data['password'] = make_password(user_data['password'])
        
        # Set created_by to the user being created (self-reference)
        user = User(**user_data)
        user.save()
        user.created_by = user
        user.save()
        
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


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def logout(request):
    device_token = request.data.get('device_token')
    
    if device_token:
        from apps.notifications.models import UserDevice
        UserDevice.objects.filter(user=request.user, device_token=device_token).delete()
    
    return Response({'message': 'Logout successful'}, status=status.HTTP_200_OK)


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def search_users(request):
    search = request.GET.get('search', '').strip()
    
    if not search:
        return Response({'error': 'search parameter is required'}, status=status.HTTP_400_BAD_REQUEST)
    
    users = User.objects.filter(
        Q(username__icontains=search) | Q(name__icontains=search),
        valid=True
    ).values('id', 'username', 'name', 'org__org_name', 'role__role_name')[:20]
    
    return Response({'users': list(users), 'count': len(users)})