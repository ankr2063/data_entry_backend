from rest_framework import status
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from django.db import transaction
from .models import Form, WorksheetMetadata, CellMetadata
from .serializers import SharePointMetadataSerializer, FormSerializer
from .services import SharePointService


@api_view(['POST'])
@permission_classes([IsAuthenticated])
def extract_sharepoint_metadata(request):
    """Extract metadata from SharePoint Excel file and save to database"""
    try:
        serializer = SharePointMetadataSerializer(data=request.data)
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
        
        data = serializer.validated_data
        sharepoint_service = SharePointService()
        
        # Get metadata from SharePoint
        metadata = sharepoint_service.get_worksheet_complete_metadata(
            data['sharepoint_url'], 
            data.get('worksheet_name')
        )
        
        # Save to database
        with transaction.atomic():
            # Create or get form
            form, created = Form.objects.get_or_create(
                form_name=data['form_name'],
                defaults={
                    'source': 'sharepoint',
                    'url': data['sharepoint_url'],
                    'created_by': 'system'  # Replace with actual user
                }
            )
            
            # Create worksheet metadata
            worksheet = WorksheetMetadata.objects.create(
                form=form,
                worksheet_name=metadata['worksheet_name'],
                row_count=metadata['dimensions']['rows'],
                column_count=metadata['dimensions']['columns'],
                sharepoint_url=data['sharepoint_url'],
                raw_data=metadata['raw_values'],
                created_by='system'  # Replace with actual user
            )
            
            # Create cell metadata
            cell_objects = []
            for cell in metadata['cells']:
                cell_objects.append(CellMetadata(
                    worksheet=worksheet,
                    cell_address=cell['address'],
                    row_index=cell['row'],
                    column_index=cell['column'],
                    value=str(cell.get('value', '')),
                    formula=cell.get('formula', ''),
                    data_type=cell.get('data_type', 'unknown'),
                    created_by='system'  # Replace with actual user
                ))
            
            CellMetadata.objects.bulk_create(cell_objects)
        
        return Response({
            'message': 'SharePoint metadata extracted and saved successfully',
            'form_id': form.id,
            'worksheet_id': worksheet.id,
            'cells_count': len(cell_objects),
            'metadata': {
                'worksheet_name': metadata['worksheet_name'],
                'dimensions': metadata['dimensions']
            }
        }, status=status.HTTP_201_CREATED)
        
    except Exception as e:
        return Response(
            {'error': f'Failed to extract metadata: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )


@api_view(['GET'])
@permission_classes([IsAuthenticated])
def get_form_metadata(request, form_id):
    """Get saved form metadata from database"""
    try:
        form = Form.objects.prefetch_related('worksheets__cells').get(id=form_id)
        serializer = FormSerializer(form)
        return Response(serializer.data)
        
    except Form.DoesNotExist:
        return Response(
            {'error': 'Form not found'}, 
            status=status.HTTP_404_NOT_FOUND
        )
    except Exception as e:
        return Response(
            {'error': f'Failed to get form metadata: {str(e)}'}, 
            status=status.HTTP_400_BAD_REQUEST
        )