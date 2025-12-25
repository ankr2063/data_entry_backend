from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from django.utils import timezone
from .models import UserDevice, Notification
from .serializers import UserDeviceSerializer, NotificationSerializer
from .services import NotificationService


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def register_device(request):
    """Register user's device token for push notifications"""
    user = request.user
    device_token = request.data.get('device_token')
    device_type = request.data.get('device_type', 'web')
    
    if not device_token:
        return Response({'error': 'device_token is required'}, status=status.HTTP_400_BAD_REQUEST)
    
    device, created = UserDevice.objects.get_or_create(
        user=user,
        device_token=device_token,
        defaults={'device_type': device_type}
    )
    
    serializer = UserDeviceSerializer(device)
    return Response({
        'message': 'Device registered successfully',
        'device': serializer.data
    }, status=status.HTTP_201_CREATED if created else status.HTTP_200_OK)


@api_view(['DELETE'])
@permission_classes([IsAuthenticated])
def unregister_device(request):
    """Unregister user's device token"""
    user = request.user
    device_token = request.data.get('device_token')
    
    if not device_token:
        return Response({'error': 'device_token is required'}, status=status.HTTP_400_BAD_REQUEST)
    
    deleted_count, _ = UserDevice.objects.filter(user=user, device_token=device_token).delete()
    
    if deleted_count > 0:
        return Response({'message': 'Device unregistered successfully'})
    else:
        return Response({'error': 'Device not found'}, status=status.HTTP_404_NOT_FOUND)


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def get_notifications(request):
    """Get user's notifications"""
    user = request.user
    notifications = Notification.objects.filter(user=user).order_by('-sent_at')
    serializer = NotificationSerializer(notifications, many=True)
    
    unread_count = notifications.filter(read=False).count()
    
    return Response({
        'notifications': serializer.data,
        'unread_count': unread_count,
        'count': len(serializer.data)
    })


@api_view(['PUT'])
@permission_classes([IsAuthenticated])
def mark_as_read(request, notification_id):
    """Mark notification as read"""
    user = request.user
    
    try:
        notification = Notification.objects.get(id=notification_id, user=user)
        notification.read = True
        notification.read_at = timezone.now()
        notification.save()
        
        return Response({'message': 'Notification marked as read'})
    except Notification.DoesNotExist:
        return Response({'error': 'Notification not found'}, status=status.HTTP_404_NOT_FOUND)


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def send_notification(request):
    """Send notification to a user"""
    user_id = request.data.get('user_id')
    title = request.data.get('title')
    body = request.data.get('body')
    data = request.data.get('data', {})
    
    if not all([user_id, title, body]):
        return Response(
            {'error': 'user_id, title, and body are required'},
            status=status.HTTP_400_BAD_REQUEST
        )
    
    result = NotificationService.send_notification(user_id, title, body, data)
    
    if result['success']:
        return Response(result, status=status.HTTP_200_OK)
    else:
        return Response(result, status=status.HTTP_400_BAD_REQUEST)
