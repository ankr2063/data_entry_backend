from django.db import models


class Organization(models.Model):
    id = models.AutoField(primary_key=True)
    org_name = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'organizations'

    def __str__(self):
        return self.org_name