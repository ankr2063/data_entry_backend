from rest_framework import serializers
from .models import Form, WorksheetMetadata, CellMetadata


class SharePointMetadataSerializer(serializers.Serializer):
    sharepoint_url = serializers.URLField()
    worksheet_name = serializers.CharField(required=False, allow_blank=True)
    form_name = serializers.CharField(max_length=255)


class CellMetadataSerializer(serializers.ModelSerializer):
    class Meta:
        model = CellMetadata
        fields = '__all__'
        read_only_fields = ('id', 'created_at', 'updated_at')


class WorksheetMetadataSerializer(serializers.ModelSerializer):
    cells = CellMetadataSerializer(many=True, read_only=True)
    
    class Meta:
        model = WorksheetMetadata
        fields = '__all__'
        read_only_fields = ('id', 'created_at', 'updated_at')


class FormSerializer(serializers.ModelSerializer):
    worksheets = WorksheetMetadataSerializer(many=True, read_only=True)
    
    class Meta:
        model = Form
        fields = '__all__'
        read_only_fields = ('id', 'created_at', 'updated_at')