from rest_framework import serializers
from .models import Form


class SharePointMetadataSerializer(serializers.Serializer):
    sharepoint_url = serializers.URLField()


class FormSerializer(serializers.ModelSerializer):
    class Meta:
        model = Form
        fields = '__all__'
        read_only_fields = ('id', 'created_at', 'updated_at')