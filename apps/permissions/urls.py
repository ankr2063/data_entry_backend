from django.urls import path
from . import views

urlpatterns = [
    path('', views.permission_list_create, name='permission_list_create'),
    path('<int:pk>/', views.permission_detail, name='permission_detail'),
    path('roles/', views.role_list_create, name='role_list_create'),
    path('roles/<int:pk>/', views.role_detail, name='role_detail'),
]
