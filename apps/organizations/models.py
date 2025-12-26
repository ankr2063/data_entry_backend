from django.db import models


class Organization(models.Model):
    id = models.AutoField(primary_key=True)
    org_name = models.CharField(max_length=255)
    created_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='created_organizations', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey('users.User', on_delete=models.RESTRICT, related_name='updated_organizations', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'organizations'

    def __str__(self):
        return self.org_name