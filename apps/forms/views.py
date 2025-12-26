from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from django.db import transaction
from .models import Form, FormDisplayVersion, FormEntryVersion, FormData, FormDataHistory, UserFormAccess
from apps.permissions.models import Role
from .serializers import SharePointMetadataSerializer, FormSerializer
from .services import SharePointService
import json


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def create_form_from_sharepoint(request):
    """Create new form from SharePoint URL"""
    try:
        user = request.user
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
            created_by=user,
            updated_by=user,
            custom_scripts=request.data.get('custom_scripts', [])
        )
        
        # Grant admin access to creator
        admin_role = Role.objects.get(role_name='Form Admin')
        UserFormAccess.objects.create(
            user=user,
            form_id=result['form_id'],
            role=admin_role,
            created_by=user
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
        
        if not form_id :
            return Response(
                {'error': 'form_id is required'}, 
                status=status.HTTP_400_BAD_REQUEST
            )
        
        sharepoint_service = SharePointService()
        
        # Update custom_scripts if provided
        custom_scripts = request.data.get('custom_scripts')
        if custom_scripts is not None:
            form = Form.objects.get(id=form_id)
            form.custom_scripts = custom_scripts
            form.updated_by = request.user
            form.save()
        
        # Update existing form
        result = sharepoint_service.update_existing_form(
            form_id,
            updated_by=request.user
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
        print(e)
        return Response(
            {'error': f'Failed to update form: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def get_forms_list(request):
    """Get list of forms user has access to"""
    try:
        user = request.user
        
        # Get forms where user has access
        user_form_access = UserFormAccess.objects.filter(user=user).select_related('form')
        accessible_forms = [access.form for access in user_form_access]
        
        serializer = FormSerializer(accessible_forms, many=True)
        return Response({
            'forms': serializer.data,
            'count': len(accessible_forms)
        })
        
    except Exception as e:
        return Response(
            {'error': f'Failed to get forms list: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['GET'])
@permission_classes([IsAuthenticated])
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
        print(e)
        return Response(
            {'error': f'Failed to get form data: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def save_form_data(request):
    """Save form data submission from frontend"""
    try:
        import base64
        import os
        from pathlib import Path
        
        user = request.user
        form_id = request.data.get('form_id')
        form_values = request.data.get('form_values')
        attachments = request.data.get('attachments', {})
        
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
            user=user,
            created_by=user,
            updated_by=user
        )
        
        # Save attachments
        attachment_urls = {}
        for field_id, files in attachments.items():
            attachment_urls[field_id] = []
            for idx, file_data in enumerate(files, 1):
                filename = file_data.get('filename')
                content = file_data.get('content')
                
                if filename and content:
                    # Create directory structure: /userUploads/formId/entryId/fieldId/
                    upload_dir = Path('userUploads') / str(form_id) / str(form_data.id) / str(field_id)
                    upload_dir.mkdir(parents=True, exist_ok=True)
                    
                    # Save file
                    file_path = upload_dir / filename
                    with open(file_path, 'wb') as f:
                        f.write(base64.b64decode(content))
                    
                    # Store relative URL
                    attachment_urls[field_id].append(str(file_path))
        
        # Get next version number for history
        last_history = FormDataHistory.objects.filter(form=form, user=user).order_by('-version').first()
        next_version = (last_history.version + 1) if last_history else 1
        
        # Save to history
        FormDataHistory.objects.create(
            form=form,
            form_entry_version=entry_version,
            form_values_json=form_values,
            version=next_version,
            user=user,
            created_by=user,
            updated_by=user
        )
        
        return Response({
            'message': 'Form data saved successfully',
            'form_data_id': form_data.id,
            'form_id': form.id,
            'entry_version': entry_version.form_version,
            'attachments': attachment_urls
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


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def get_form_entries(request, form_id):
    """Get all form data entries for a specific form filtered by user"""
    try:
        user = request.user
        form = Form.objects.get(id=form_id)
        form_entries = FormData.objects.filter(form=form, user=user).order_by('-id')
        
        # Get latest entry version to extract column names
        latest_entry_version = FormEntryVersion.objects.filter(form=form).order_by('-form_version').first()
        
        # Create columns dictionary from entry JSON (id -> name mapping)
        columns = {}
        if latest_entry_version and latest_entry_version.form_entry_json:
            for item in latest_entry_version.form_entry_json:
                if 'id' in item and 'name' in item:
                    columns[str(item['id'])] = item['name']
        
        # Format entries data
        entries_data = []
        for entry in form_entries:
            # Get attachments for this entry
            import os
            from pathlib import Path
            
            attachments = {}
            upload_dir = Path('userUploads') / str(form.id) / str(entry.id)
            
            if upload_dir.exists():
                for field_dir in upload_dir.iterdir():
                    if field_dir.is_dir():
                        field_id = field_dir.name
                        attachments[field_id] = []
                        for file_path in field_dir.iterdir():
                            if file_path.is_file():
                                attachments[field_id].append(str(file_path))
            
            entries_data.append({
                'id': entry.id,
                'values': entry.form_values_json,
                'attachments': attachments,
                'created_by': entry.created_by.id if entry.created_by else None,
                'created_at': entry.created_at,
                'updated_by': entry.updated_by.id if entry.updated_by else None,
                'updated_at': entry.updated_at
            })
        
        return Response({
            'form_id': form.id,
            'form_name': form.form_name,
            'columns': columns,
            'entries': entries_data,
            'count': len(entries_data)
        })
        
    except Form.DoesNotExist:
        return Response(
            {'error': 'Form not found'}, 
            status=status.HTTP_404_NOT_FOUND
        )
    except Exception as e:
        return Response(
            {'error': f'Failed to get form entries: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def get_filled_display_data(request, form_data_id):
    """Get display data with values filled from form data"""
    try:
        form_data = FormData.objects.get(id=form_data_id)
        form = form_data.form
        
        # Get latest display version
        display_version = FormDisplayVersion.objects.filter(form=form).order_by('-form_version').first()
        
        if not display_version:
            return Response(
                {'error': 'Display version not found'}, 
                status=status.HTTP_404_NOT_FOUND
            )
        
        display_data = display_version.form_display_json.copy()
        form_values = form_data.form_values_json
        
        # Parse form_values if it's a string
        if isinstance(form_values, str):
            import json
            form_values = json.loads(form_values)
        
        # Fill values in cells
        for cell in display_data.get('cells', []):
            cell_value = cell.get('value', '')
            if cell_value and '<pa' in str(cell_value):
                # Extract ID from pattern like <pa_1>
                import re
                match = re.search(r'<pa_(\d+)>', str(cell_value))
                if match:
                    field_id = match.group(1)
                    if field_id in form_values:
                        cell['value'] = form_values[field_id]
                        cell['display_value'] = str(form_values[field_id])
        
        return Response({
            'form_data_id': form_data_id,
            'form_id': form.id,
            'form_name': form.form_name,
            'display_data': display_data
        })
        
    except FormData.DoesNotExist:
        return Response(
            {'error': 'Form data not found'}, 
            status=status.HTTP_404_NOT_FOUND
        )
    except Exception as e:
        return Response(
            {'error': f'Failed to get filled display data: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )