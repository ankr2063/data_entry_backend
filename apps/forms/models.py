from django.db import models
from apps.users.models import User


class Form(models.Model):
    id = models.AutoField(primary_key=True)
    form_name = models.CharField(max_length=255)
    source = models.CharField(max_length=255, null=True, blank=True)
    url = models.URLField(null=True, blank=True)
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.CharField(max_length=255, null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'forms'


class UserFormAccess(models.Model):
    id = models.AutoField(primary_key=True)
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    form = models.ForeignKey(Form, on_delete=models.CASCADE)
    role = models.CharField(max_length=100)
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.CharField(max_length=255, null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'user_form_access'


class FormEntryVersion(models.Model):
    id = models.AutoField(primary_key=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, db_column='form_id')
    form_entry_json = models.JSONField()
    form_version = models.CharField(max_length=50)
    approved = models.BooleanField(default=False)
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.CharField(max_length=255, null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'form_entry_versions'


class FormDisplayVersion(models.Model):
    id = models.AutoField(primary_key=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, db_column='form_id')
    form_display_json = models.JSONField()
    form_version = models.CharField(max_length=50)
    approved = models.BooleanField(default=False)
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.CharField(max_length=255, null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'form_display_versions'


class FormData(models.Model):
    id = models.AutoField(primary_key=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    created_by = models.CharField(max_length=255)
    updated_by = models.CharField(max_length=255, null=True, blank=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, db_column='form_id')
    form_entry_version = models.ForeignKey(FormEntryVersion, on_delete=models.CASCADE, db_column='form_entry_vid')
    form_values_json = models.JSONField()

    class Meta:
        db_table = 'form_datas'


class FormDataHistory(models.Model):
    id = models.AutoField(primary_key=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, db_column='form_id')
    form_entry_version = models.ForeignKey(FormEntryVersion, on_delete=models.CASCADE, db_column='form_entry_vid')
    form_values_json = models.JSONField()
    version = models.IntegerField()
    updated_by = models.CharField(max_length=255)
    updated_at = models.DateTimeField(auto_now=True)
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'form_data_history'