from firebase_admin import credentials, messaging, initialize_app
from decouple import config
from .models import UserDevice, Notification
import json


class NotificationService:
    _initialized = False
    
    @classmethod
    def initialize_firebase(cls):
        if not cls._initialized:
            try:
                cred_dict = {
                    "type": "service_account",
                    "project_id": config('FIREBASE_PROJECT_ID'),
                    "private_key": config('FIREBASE_PRIVATE_KEY').replace('\\n', '\n'),
                    "client_email": config('FIREBASE_CLIENT_EMAIL'),
                }
                cred = credentials.Certificate(cred_dict)
                initialize_app(cred)
                cls._initialized = True
            except Exception as e:
                print(f"Firebase initialization error: {e}")
    
    @classmethod
    def send_notification(cls, user_id, title, body, data=None):
        """Send notification to a user"""
        cls.initialize_firebase()
        
        # Get user's device tokens
        devices = UserDevice.objects.filter(user_id=user_id)
        
        if not devices.exists():
            return {'success': False, 'message': 'No devices registered'}
        
        # Store notification in database
        notification = Notification.objects.create(
            user_id=user_id,
            title=title,
            body=body,
            data=data
        )
        
        # Send to all user devices
        tokens = [device.device_token for device in devices]
        
        message = messaging.MulticastMessage(
            notification=messaging.Notification(
                title=title,
                body=body
            ),
            data=data or {},
            tokens=tokens
        )
        
        try:
            response = messaging.send_multicast(message)
            return {
                'success': True,
                'notification_id': notification.id,
                'success_count': response.success_count,
                'failure_count': response.failure_count
            }
        except Exception as e:
            return {'success': False, 'message': str(e)}
