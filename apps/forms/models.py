from django.db import models
from apps.users.models import User
from apps.permissions.models import Role


class Form(models.Model):
    id = models.AutoField(primary_key=True)
    form_name = models.CharField(max_length=255)
    source = models.CharField(max_length=255, null=True, blank=True)
    url = models.URLField(null=True, blank=True)
    custom_scripts = models.JSONField(default=list, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='created_forms', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='updated_forms', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'forms'


class UserFormAccess(models.Model):
    id = models.AutoField(primary_key=True)
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    form = models.ForeignKey(Form, on_delete=models.CASCADE)
    role = models.ForeignKey(Role, on_delete=models.CASCADE, db_column='role_id')
    created_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='created_form_accesses', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='updated_form_accesses', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'user_form_access'


class FormEntryVersion(models.Model):
    id = models.AutoField(primary_key=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, db_column='form_id')
    form_entry_json = models.JSONField()
    form_version = models.CharField(max_length=50)
    approved = models.BooleanField(default=False)
    created_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='created_entry_versions', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='updated_entry_versions', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'form_entry_versions'


class FormDisplayVersion(models.Model):
    id = models.AutoField(primary_key=True)
    form = models.ForeignKey(Form, on_delete=models.CASCADE, db_column='form_id')
    form_display_json = models.JSONField()
    form_version = models.CharField(max_length=50)
    approved = models.BooleanField(default=False)
    created_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='created_display_versions', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='updated_display_versions', null=True, blank=True, db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        db_table = 'form_display_versions'


class FormData(models.Model):
    id = models.AutoField(primary_key=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    created_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='created_form_datas', db_column='created_by')
    updated_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='updated_form_datas', null=True, blank=True, db_column='updated_by')
    user = models.ForeignKey(User, on_delete=models.CASCADE, db_column='user_id')
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
    user = models.ForeignKey(User, on_delete=models.CASCADE, db_column='user_id')
    updated_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='updated_form_data_histories', db_column='updated_by')
    updated_at = models.DateTimeField(auto_now=True)
    created_by = models.ForeignKey(User, on_delete=models.RESTRICT, related_name='created_form_data_histories', db_column='created_by')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'form_data_history'