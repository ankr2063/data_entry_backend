from rest_framework import serializers
from .models import Permission, Role, RolePermission


class PermissionSerializer(serializers.ModelSerializer):
    class Meta:
        model = Permission
        fields = ['id', 'permission_name', 'description', 'created_by', 'created_at', 'updated_by', 'updated_at']
        read_only_fields = ['id', 'created_at', 'updated_at']


class RoleSerializer(serializers.ModelSerializer):
    permissions = PermissionSerializer(many=True, read_only=True)
    permission_ids = serializers.ListField(child=serializers.IntegerField(), write_only=True, required=False)
    
    class Meta:
        model = Role
        fields = ['id', 'role_name', 'description', 'permissions', 'permission_ids', 'created_by', 'created_at', 'updated_by', 'updated_at']
        read_only_fields = ['id', 'created_at', 'updated_at']
    
    def create(self, validated_data):
        permission_ids = validated_data.pop('permission_ids', [])
        role = Role.objects.create(**validated_data)
        
        for perm_id in permission_ids:
            RolePermission.objects.create(
                role=role,
                permission_id=perm_id,
                created_by=validated_data.get('created_by', 'system')
            )
        
        return role
    
    def update(self, instance, validated_data):
        permission_ids = validated_data.pop('permission_ids', None)
        
        for attr, value in validated_data.items():
            setattr(instance, attr, value)
        instance.save()
        
        if permission_ids is not None:
            RolePermission.objects.filter(role=instance).delete()
            for perm_id in permission_ids:
                RolePermission.objects.create(
                    role=instance,
                    permission_id=perm_id,
                    created_by=validated_data.get('updated_by', 'system')
                )
        
        return instance
