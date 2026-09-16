from django.contrib import admin

from .models import (
    ActivityType,
    Day,
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    Venue,
    WorkshopAllocation,
)


@admin.register(Semester)
class SemesterAdmin(admin.ModelAdmin):
    list_display = ("id", "academic_year", "semester")
    list_display_links = list_display
    list_editable = ()
    ordering = ("academic_year", "semester")


@admin.register(Programme)
class ProgrammeAdmin(admin.ModelAdmin):
    list_display = ("id", "code", "name")
    list_display_links = list_display
    search_fields = ("code", "name")


class StudentGroupInline(admin.TabularInline):
    model = StudentGroup
    extra = 1
    fields = ("code",)


@admin.register(StudentGroup)
class StudentGroupAdmin(admin.ModelAdmin):
    list_display = ("id", "programme", "code")
    list_display_links = list_display
    list_filter = ("programme",)
    search_fields = ("code", "programme__code", "programme__name")
    raw_id_fields = ("programme",)


@admin.register(ProgrammeCourse)
class ProgrammeCourseAdmin(admin.ModelAdmin):
    list_display = ("id", "programme", "course_code")
    list_display_links = list_display
    list_filter = ("programme",)
    search_fields = ("course_code", "programme__code")
    raw_id_fields = ("programme",)


@admin.register(Venue)
class VenueAdmin(admin.ModelAdmin):
    list_display = ("id", "name", "capacity")
    list_display_links = list_display
    search_fields = ("name",)


class SessionGroupInline(admin.TabularInline):
    model = SessionGroup
    extra = 1
    raw_id_fields = ("group",)


@admin.register(Session)
class SessionAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "semester",
        "course_code",
        "activity_type",
        "day",
        "start_time",
        "end_time",
        "venue",
    )
    list_display_links = list_display
    list_filter = ("semester", "activity_type", "day")
    search_fields = ("course_code", "venue__name")
    raw_id_fields = ("semester", "venue")
    inlines = [SessionGroupInline]


@admin.register(SessionGroup)
class SessionGroupAdmin(admin.ModelAdmin):
    list_display = ("id", "session", "group")
    list_display_links = list_display
    list_filter = ("session__activity_type",)
    raw_id_fields = ("session", "group")


@admin.register(WorkshopAllocation)
class WorkshopAllocationAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "semester",
        "course_code",
        "group_code",
        "day",
        "start_time",
        "end_time",
        "venue",
    )
    list_display_links = list_display
    list_filter = ("semester", "day")
    search_fields = ("course_code", "group_code")


@admin.register(TechnicalDrawingAllocation)
class TechnicalDrawingAllocationAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "semester",
        "course_code",
        "group_code",
        "day",
        "start_time",
        "end_time",
        "venue",
    )
    list_display_links = list_display
    list_filter = ("semester", "day")
    search_fields = ("course_code", "group_code")