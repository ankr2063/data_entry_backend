from django.urls import path
from . import views

urlpatterns = [
    path('register-device/', views.register_device, name='register_device'),
    path('unregister-device/', views.unregister_device, name='unregister_device'),
    path('', views.get_notifications, name='get_notifications'),
    path('<int:notification_id>/read/', views.mark_as_read, name='mark_as_read'),
    path('send/', views.send_notification, name='send_notification'),
]
