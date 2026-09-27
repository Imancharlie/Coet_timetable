from django.contrib import admin

from .models import CollisionReport, PortalSettings


@admin.register(PortalSettings)
class PortalSettingsAdmin(admin.ModelAdmin):
    list_display = ("current_semester",)

    def has_add_permission(self, request):
        return not PortalSettings.objects.exists()

    def has_delete_permission(self, request, obj=None):
        return False


@admin.register(CollisionReport)
class CollisionReportAdmin(admin.ModelAdmin):
    list_display = (
        "created_at", "reference", "timetable_type", "semester", "programme",
        "group", "status",
    )
    list_filter = ("status", "timetable_type", "semester", "created_at")
    search_fields = ("reference", "course_or_exam", "description", "contact_email")
    readonly_fields = ("reference", "created_at", "updated_at")
    list_editable = ("status",)
