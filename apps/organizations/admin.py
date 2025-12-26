from django.contrib import admin
from .models import Organization


@admin.register(Organization)
class OrganizationAdmin(admin.ModelAdmin):
    list_display = ['id', 'org_name', 'created_at']
    search_fields = ['org_name']
    readonly_fields = ['created_at', 'updated_at']