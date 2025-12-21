from django.contrib import admin
from .models import Form, WorksheetMetadata, CellMetadata, UserFormAccess


@admin.register(Form)
class FormAdmin(admin.ModelAdmin):
    list_display = ['id', 'form_name', 'source', 'created_by', 'created_at']
    search_fields = ['form_name']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(WorksheetMetadata)
class WorksheetMetadataAdmin(admin.ModelAdmin):
    list_display = ['id', 'form', 'worksheet_name', 'row_count', 'column_count', 'created_at']
    list_filter = ['form']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(CellMetadata)
class CellMetadataAdmin(admin.ModelAdmin):
    list_display = ['id', 'worksheet', 'cell_address', 'data_type', 'value']
    list_filter = ['data_type', 'worksheet']
    search_fields = ['cell_address', 'value']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(UserFormAccess)
class UserFormAccessAdmin(admin.ModelAdmin):
    list_display = ['id', 'user', 'form', 'role', 'created_at']
    list_filter = ['role']