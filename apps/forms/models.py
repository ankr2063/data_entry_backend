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


class WorksheetMetadata(models.Model):
    id = models.AutoField(primary_key=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, related_name='worksheets')
    worksheet_name = models.CharField(max_length=255)
    row_count = models.IntegerField()
    column_count = models.IntegerField()
    sharepoint_url = models.URLField()
    raw_data = models.JSONField()
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.CharField(max_length=255, null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'worksheet_metadata'


class CellMetadata(models.Model):
    id = models.AutoField(primary_key=True)
    worksheet = models.ForeignKey(WorksheetMetadata, on_delete=models.CASCADE, related_name='cells')
    cell_address = models.CharField(max_length=20)
    row_index = models.IntegerField()
    column_index = models.IntegerField()
    value = models.TextField(null=True, blank=True)
    formula = models.TextField(null=True, blank=True)
    data_type = models.CharField(max_length=50)
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.CharField(max_length=255, null=True, blank=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'cell_metadata'
        unique_together = ['worksheet', 'cell_address']


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
    form_template_json = models.JSONField()
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
    form_config_json = models.JSONField()
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
    form_display_version = models.ForeignKey(FormDisplayVersion, on_delete=models.CASCADE, db_column='form_display_vid')
    form_values_json = models.JSONField()

    class Meta:
        db_table = 'form_datas'


class FormDataHistory(models.Model):
    id = models.AutoField(primary_key=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, db_column='form_id')
    form_entry_version = models.ForeignKey(FormEntryVersion, on_delete=models.CASCADE, db_column='form_entry_vid')
    form_display_version = models.ForeignKey(FormDisplayVersion, on_delete=models.CASCADE, db_column='form_display_vid')
    form_values_json = models.JSONField()
    updated_by = models.CharField(max_length=255)
    updated_at = models.DateTimeField(auto_now=True)
    created_by = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'form_data_history'