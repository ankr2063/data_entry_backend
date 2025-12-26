from django.db import models
from apps.organizations.models import Organization


class Permission(models.Model):
    id = models.AutoField(primary_key=True)
    org = models.ForeignKey(Organization, on_delete=models.CASCADE, db_column='org_id')
    permission_name = models.CharField(max_length=255)
    description = models.TextField(null=True, blank=True)
    created_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='created_permissions', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='updated_permissions', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'permissions'
        unique_together = ('org', 'permission_name')

    def __str__(self):
        return self.permission_name


class Role(models.Model):
    id = models.AutoField(primary_key=True)
    org = models.ForeignKey(Organization, on_delete=models.CASCADE, db_column='org_id')
    role_name = models.CharField(max_length=255)
    description = models.TextField(null=True, blank=True)
    permissions = models.ManyToManyField(Permission, through='RolePermission')
    created_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='created_roles', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='updated_roles', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'roles'
        unique_together = ('org', 'role_name')

    def __str__(self):
        return self.role_name


class RolePermission(models.Model):
    id = models.AutoField(primary_key=True)
    role = models.ForeignKey(Role, on_delete=models.CASCADE, db_column='role_id')
    permission = models.ForeignKey(Permission, on_delete=models.CASCADE, db_column='permission_id')
    created_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='created_role_permissions', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='updated_role_permissions', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'role_permissions'
        unique_together = ('role', 'permission')
