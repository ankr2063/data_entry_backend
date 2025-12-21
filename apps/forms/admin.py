from django.contrib import admin
from .models import Form, UserFormAccess, FormDisplayVersion, FormEntryVersion, FormData, FormDataHistory


@admin.register(Form)
class FormAdmin(admin.ModelAdmin):
    list_display = ['id', 'form_name', 'source', 'created_by', 'created_at']
    search_fields = ['form_name']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(FormDisplayVersion)
class FormDisplayVersionAdmin(admin.ModelAdmin):
    list_display = ['id', 'form', 'form_version', 'approved', 'created_at']
    list_filter = ['approved', 'form']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(FormEntryVersion)
class FormEntryVersionAdmin(admin.ModelAdmin):
    list_display = ['id', 'form', 'form_version', 'approved', 'created_at']
    list_filter = ['approved', 'form']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(UserFormAccess)
class UserFormAccessAdmin(admin.ModelAdmin):
    list_display = ['id', 'user', 'form', 'role', 'created_at']
    list_filter = ['role']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(FormData)
class FormDataAdmin(admin.ModelAdmin):
    list_display = ['id', 'form', 'created_by', 'created_at']
    list_filter = ['form']
    readonly_fields = ['created_at', 'updated_at']


@admin.register(FormDataHistory)
class FormDataHistoryAdmin(admin.ModelAdmin):
    list_display = ['id', 'form', 'updated_by', 'updated_at']
    list_filter = ['form']
    readonly_fields = ['created_at', 'updated_at']