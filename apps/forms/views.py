from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from django.db import transaction
from .models import Form, FormDisplayVersion, FormEntryVersion
from .serializers import SharePointMetadataSerializer, FormSerializer
from .services import SharePointService


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def extract_sharepoint_metadata(request):
    """Process SharePoint URL with display and entry worksheets"""
    try:
        serializer = SharePointMetadataSerializer(data=request.data)
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
        
        data = serializer.validated_data
        sharepoint_service = SharePointService()
        
        # Process SharePoint URL with new flow
        result = sharepoint_service.process_sharepoint_url(
            data['sharepoint_url'],
            created_by='system',  # Replace with actual user from request
            updated_by='system'   # Replace with actual user from request
        )
        
        return Response({
            'message': 'SharePoint worksheets processed successfully',
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
            {'error': f'Failed to process SharePoint data: {str(e)}'}, 
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