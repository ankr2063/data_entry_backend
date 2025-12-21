from django.urls import path
from . import views

urlpatterns = [
    path('extract-metadata/', views.extract_sharepoint_metadata, name='extract_sharepoint_metadata'),
    path('<int:form_id>/metadata/', views.get_form_metadata, name='get_form_metadata'),
]