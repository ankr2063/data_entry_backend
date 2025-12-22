from django.urls import path
from . import views

urlpatterns = [
    path('', views.get_forms_list, name='get_forms_list'),
    path('create/', views.create_form_from_sharepoint, name='create_form_from_sharepoint'),
    path('update/', views.update_form_from_sharepoint, name='update_form_from_sharepoint'),
    path('<int:form_id>/metadata/<str:metadata_type>/', views.get_form_metadata, name='get_form_metadata'),
]