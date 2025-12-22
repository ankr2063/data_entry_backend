from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from django.db import transaction
from .models import Form, FormDisplayVersion, FormEntryVersion, FormData, FormDataHistory
from .serializers import SharePointMetadataSerializer, FormSerializer
from .services import SharePointService
import json


@api_view(['POST'])
#@permission_classes([IsAuthenticated])
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
#@permission_classes([IsAuthenticated])
def update_form_from_sharepoint(request):
    """Update existing form from SharePoint URL"""
    try:
        form_id = request.data.get('form_id')
        
        if not form_id :
            return Response(
                {'error': 'form_id is required'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        sharepoint_service = SharePointService()
        
        # Update existing form
        result = sharepoint_service.update_existing_form(
            form_id,
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
#@permission_classes([IsAuthenticated])
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
#@permission_classes([IsAuthenticated])
def get_form_metadata(request, form_id, metadata_type):
    """Get form metadata based on type: 'entry', 'display', or 'both'"""
    try:
        if metadata_type not in ['entry', 'display', 'both']:
            return Response(
                {'error': 'metadata_type must be one of: entry, display, both'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        form = Form.objects.get(id=form_id)
        response_data = {'form': FormSerializer(form).data}
        
        if metadata_type in ['entry', 'both']:
            entry_version = FormEntryVersion.objects.filter(form=form).order_by('-form_version').first()
            response_data['entry_data'] = entry_version.form_entry_json if entry_version else []
            response_data['entry_version'] = {
                'id': entry_version.id if entry_version else None,
                'version': entry_version.form_version if entry_version else None,
                'approved': entry_version.approved if entry_version else None
            }
        
        if metadata_type in ['display', 'both']:
            display_version = FormDisplayVersion.objects.filter(form=form).order_by('-form_version').first()
            response_data['display_data'] = display_version.form_display_json if display_version else {}
            response_data['display_version'] = {
                'id': display_version.id if display_version else None,
                'version': display_version.form_version if display_version else None,
                'approved': display_version.approved if display_version else None
            }
        
        return Response(response_data)
        
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


@api_view(['POST'])
#@permission_classes([IsAuthenticated])
def save_form_data(request):
    """Save form data submission from frontend"""
    try:
        form_id = request.data.get('form_id')
        form_values = request.data.get('form_values')
        
        if not form_id or not form_values:
            return Response(
                {'error': 'form_id and form_values are required'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        form = Form.objects.get(id=form_id)
        
        # Get latest versions
        entry_version = FormEntryVersion.objects.filter(form=form).order_by('-form_version').first()
        
        if not entry_version:
            return Response(
                {'error': 'Form versions not found'}, 
                status=status.HTTP_404_NOT_FOUND
            )
        
        # Create form data entry
        form_data = FormData.objects.create(
            form=form,
            form_entry_version=entry_version,
            form_values_json=form_values,
            created_by='system',  # Replace with actual user
            updated_by='system'
        )
        
        # Get next version number for history
        last_history = FormDataHistory.objects.filter(form=form).order_by('-version').first()
        next_version = (last_history.version + 1) if last_history else 1
        
        # Save to history
        FormDataHistory.objects.create(
            form=form,
            form_entry_version=entry_version,
            form_values_json=form_values,
            version=next_version,
            created_by='system',
            updated_by='system'
        )
        
        return Response({
            'message': 'Form data saved successfully',
            'form_data_id': form_data.id,
            'form_id': form.id,
            'entry_version': entry_version.form_version
        }, status=status.HTTP_201_CREATED)
        
    except Form.DoesNotExist:
        return Response(
            {'error': 'Form not found'}, 
            status=status.HTTP_404_NOT_FOUND
        )
    except Exception as e:
        return Response(
            {'error': f'Failed to save form data: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )