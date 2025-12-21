from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from django.db import transaction
from .models import Form, FormDisplayVersion, FormEntryVersion
from .serializers import SharePointMetadataSerializer, FormSerializer
from .services import SharePointService


from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from django.db import transaction
from .models import Form, FormDisplayVersion, FormEntryVersion
from .serializers import SharePointMetadataSerializer, FormSerializer
from .services import SharePointService
import json


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def create_form_from_sharepoint(request):
    """Create new form from SharePoint URL"""
    try:
        form_name = request.data.get('form_name')
        sharepoint_url = request.data.get('sharepoint_url')
        
        if not form_name or not sharepoint_url:
            return Response(
                {'error': 'form_name and sharepoint_url are required'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        sharepoint_service = SharePointService()
        
        # Process SharePoint URL with new flow
        result = sharepoint_service.create_new_form(
            sharepoint_url,
            form_name,
            created_by='system',  # Replace with actual user from request
            updated_by='system'   # Replace with actual user from request
        )
        
        return Response({
            'message': 'Form created successfully from SharePoint',
            'form_id': result['form_id'],
            'form_name': result['form_name'],
            'display_version': result['display_version'],
            'entry_version': result['entry_version'],
            'worksheets': {
                'display': result['display_sheet'],
                'entry': result['entry_sheet']
            }
        }, status=status.HTTP_201_CREATED)
        
    except Exception as e:
        return Response(
            {'error': f'Failed to create form: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def update_form_from_sharepoint(request):
    """Update existing form from SharePoint URL"""
    try:
        form_id = request.data.get('form_id')
        sharepoint_url = request.data.get('sharepoint_url')
        
        if not form_id or not sharepoint_url:
            return Response(
                {'error': 'form_id and sharepoint_url are required'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        sharepoint_service = SharePointService()
        
        # Update existing form
        result = sharepoint_service.update_existing_form(
            form_id,
            sharepoint_url,
            updated_by='system'  # Replace with actual user from request
        )
        
        return Response({
            'message': 'Form updated successfully from SharePoint',
            'form_id': result['form_id'],
            'form_name': result['form_name'],
            'display_version': result['display_version'],
            'entry_version': result['entry_version'],
            'versions_updated': result['versions_updated'],
            'worksheets': {
                'display': result['display_sheet'],
                'entry': result['entry_sheet']
            }
        }, status=status.HTTP_200_OK)
        
    except Form.DoesNotExist:
        return Response(
            {'error': 'Form not found'}, 
            status=status.HTTP_404_NOT_FOUND
        )
    except Exception as e:
        return Response(
            {'error': f'Failed to update form: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def get_forms_list(request):
    """Get list of all forms"""
    try:
        forms = Form.objects.all().order_by('-created_at')
        serializer = FormSerializer(forms, many=True)
        return Response({
            'forms': serializer.data,
            'count': forms.count()
        })
        
    except Exception as e:
        return Response(
            {'error': f'Failed to get forms list: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def get_form_metadata(request, form_id):
    """Get form with display and entry versions"""
    try:
        form = Form.objects.get(id=form_id)
        
        # Get latest versions
        display_version = FormDisplayVersion.objects.filter(form=form).order_by('-form_version').first()
        entry_version = FormEntryVersion.objects.filter(form=form).order_by('-form_version').first()
        
        return Response({
            'form': FormSerializer(form).data,
            'display_version': {
                'id': display_version.id if display_version else None,
                'version': display_version.form_version if display_version else None,
                'approved': display_version.approved if display_version else None,
                'data': display_version.form_display_json if display_version else None
            },
            'entry_version': {
                'id': entry_version.id if entry_version else None,
                'version': entry_version.form_version if entry_version else None,
                'approved': entry_version.approved if entry_version else None,
                'data': entry_version.form_entry_json if entry_version else None
            }
        })
        
    except Form.DoesNotExist:
        return Response(
            {'error': 'Form not found'}, 
            status=status.HTTP_404_NOT_FOUND
        )
    except Exception as e:
        return Response(
            {'error': f'Failed to get form data: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )