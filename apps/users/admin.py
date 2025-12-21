from django.contrib import admin
from .models import User


@admin.register(User)
class UserAdmin(admin.ModelAdmin):
    list_display = ['id', 'username', 'name', 'org', 'role', 'valid', 'created_at']
    list_filter = ['valid', 'role', 'org']
    search_fields = ['username', 'name']
    readonly_fields = ['created_at', 'updated_at']