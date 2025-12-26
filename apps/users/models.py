from django.db import models
from apps.organizations.models import Organization
from apps.permissions.models import Role


class User(models.Model):
    id = models.AutoField(primary_key=True)
    org = models.ForeignKey(Organization, on_delete=models.CASCADE, db_column='org_id')
    username = models.CharField(max_length=255, unique=True)
    password = models.CharField(max_length=255)
    name = models.CharField(max_length=255)
    valid = models.BooleanField(default=True)
    attest_url = models.URLField(null=True, blank=True)
    role = models.ForeignKey(Role, on_delete=models.SET_NULL, null=True, db_column='role_id')
    created_by = models.ForeignKey('self', on_delete=models.RESTRICT, related_name='created_users', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey('self', on_delete=models.RESTRICT, related_name='updated_users', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'users'

    def __str__(self):
        return self.username
    
    @property
    def is_authenticated(self):
        return True