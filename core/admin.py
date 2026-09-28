from django.contrib import admin

from .models import (
    ActivityLog,
    ActivityType,
    AllocationChange,
    AllocationRun,
    Course,
    CourseActivityRequirement,
    Day,
    LogAction,
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
    # is_current is ticked straight from the list, the same one-value rule the
    # semesters page uses; the conditional unique constraint stops two rows
    # ever being ticked at once.
    list_display = ("id", "academic_year", "semester", "is_current")
    list_display_links = ("id", "academic_year", "semester")
    list_editable = ("is_current",)
    list_filter = ("is_current", "semester")
    ordering = ("-is_current", "academic_year", "semester")

    def save_model(self, request, obj, form, change):
        # Both the change form and the changelist tick write the boolean
        # straight to the database, bypassing Semester.set_current(), so the
        # previous holder has to be cleared here or the conditional unique
        # constraint would raise instead of switching terms.
        if obj.is_current:
            Semester.objects.exclude(pk=obj.pk).update(is_current=False)
        super().save_model(request, obj, form, change)


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


class CourseActivityRequirementInline(admin.TabularInline):
    model = CourseActivityRequirement
    extra = 1
    fields = ("activity_type", "count")


@admin.register(Course)
class CourseAdmin(admin.ModelAdmin):
    list_display = ("id", "code", "name")
    list_display_links = list_display
    search_fields = ("code", "name")
    inlines = [CourseActivityRequirementInline]


@admin.register(ProgrammeCourse)
class ProgrammeCourseAdmin(admin.ModelAdmin):
    list_display = ("id", "programme", "course", "semester")
    list_display_links = list_display
    list_filter = ("programme", "semester")
    search_fields = ("course__code", "course__name", "course_code", "programme__code")
    raw_id_fields = ("programme", "course")


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


@admin.register(ActivityLog)
class ActivityLogAdmin(admin.ModelAdmin):
    list_display = ("id", "created_at", "action", "resource", "target", "message")
    list_display_links = ("id", "message")
    list_filter = ("action", "resource")
    search_fields = ("resource", "target", "message")
    readonly_fields = ("action", "resource", "target", "message", "created_at")


class AllocationChangeInline(admin.TabularInline):
    model = AllocationChange
    extra = 0
    fields = ("action", "group", "session", "course_code", "activity_type")
    readonly_fields = fields
    raw_id_fields = ("group", "session")

    def has_add_permission(self, request, obj=None):
        return False


@admin.register(AllocationRun)
class AllocationRunAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "semester",
        "scope",
        "status",
        "assigned",
        "moved",
        "retained",
        "removed",
        "unresolved",
        "created_at",
    )
    list_display_links = ("id",)
    list_filter = ("semester", "scope", "status")
    search_fields = ("semester__academic_year",)
    readonly_fields = ("created_at", "applied_at", "reverted_at", "summary")
    inlines = [AllocationChangeInline]

    def has_add_permission(self, request):
        return False