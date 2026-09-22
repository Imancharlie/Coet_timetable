from django.conf import settings
from django.db import models


class ActivityType(models.TextChoices):
    LECTURE = "LECTURE", "Lecture"
    TUTORIAL = "TUTORIAL", "Tutorial"
    SEMINAR = "SEMINAR", "Seminar"
    PRACTICAL = "PRACTICAL", "Practical"
    WORKSHOP = "WORKSHOP", "Workshop"


class Day(models.TextChoices):
    MONDAY = "MONDAY", "Monday"
    TUESDAY = "TUESDAY", "Tuesday"
    WEDNESDAY = "WEDNESDAY", "Wednesday"
    THURSDAY = "THURSDAY", "Thursday"
    FRIDAY = "FRIDAY", "Friday"
    SATURDAY = "SATURDAY", "Saturday"
    SUNDAY = "SUNDAY", "Sunday"


class TimePeriod(models.TextChoices):
    MORNING = "MORNING", "Morning"
    AFTERNOON = "AFTERNOON", "Afternoon"


class Semester(models.Model):
    academic_year = models.CharField(max_length=20)
    semester = models.PositiveSmallIntegerField()

    class Meta:
        ordering = ["-academic_year", "-semester"]
        unique_together = ["academic_year", "semester"]

    def __str__(self):
        return f"{self.academic_year} - Semester {self.semester}"


class Programme(models.Model):
    code = models.CharField(max_length=20, unique=True)
    name = models.CharField(max_length=200)

    class Meta:
        ordering = ["code"]

    def __str__(self):
        return f"{self.code} - {self.name}"


class StudentGroup(models.Model):
    programme = models.ForeignKey(
        Programme, on_delete=models.CASCADE, related_name="student_groups"
    )
    code = models.CharField(max_length=20)

    class Meta:
        ordering = ["programme__code", "code"]
        unique_together = ["programme", "code"]

    def __str__(self):
        return f"{self.programme.code} {self.code}"


class ProgrammeCourse(models.Model):
    programme = models.ForeignKey(
        Programme, on_delete=models.CASCADE, related_name="programme_courses"
    )
    course_code = models.CharField(max_length=20)
    course_name = models.CharField(max_length=200)
    semester = models.PositiveSmallIntegerField()

    class Meta:
        ordering = ["programme__code", "semester", "course_code"]
        unique_together = ["programme", "course_code"]

    def __str__(self):
        return f"{self.programme.code} - {self.course_code} {self.course_name}"


class Venue(models.Model):
    name = models.CharField(max_length=50, unique=True)
    capacity = models.IntegerField()

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Session(models.Model):
    semester = models.ForeignKey(
        Semester, on_delete=models.CASCADE, related_name="sessions"
    )
    course_code = models.CharField(max_length=20)
    activity_type = models.CharField(
        max_length=20, choices=ActivityType.choices
    )
    day = models.CharField(max_length=10, choices=Day.choices)
    start_time = models.TimeField()
    end_time = models.TimeField()
    venue = models.ForeignKey(
        Venue,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="sessions",
    )

    class Meta:
        ordering = ["semester", "day", "start_time"]

    def __str__(self):
        venue = self.venue.name if self.venue else "-"
        return (
            f"{self.course_code} {self.activity_type} "
            f"{self.get_day_display()} {self.start_time:%H:%M}-{self.end_time:%H:%M} "
            f"({venue})"
        )


class SessionGroup(models.Model):
    session = models.ForeignKey(
        Session, on_delete=models.CASCADE, related_name="session_groups"
    )
    group = models.ForeignKey(
        StudentGroup, on_delete=models.CASCADE, related_name="session_groups"
    )

    class Meta:
        ordering = ["session", "group"]
        unique_together = ["session", "group"]

    def __str__(self):
        return f"{self.session} -> {self.group}"


class WorkshopAllocation(models.Model):
    semester = models.ForeignKey(
        Semester, on_delete=models.CASCADE, related_name="workshop_allocations"
    )
    course_code = models.CharField(max_length=20)
    group_code = models.CharField(max_length=20)
    day = models.CharField(max_length=10, choices=Day.choices)
    start_time = models.TimeField(null=True, blank=True)
    end_time = models.TimeField(null=True, blank=True)
    venue = models.CharField(max_length=50, blank=True)
    workshop = models.CharField(max_length=50, blank=True)
    time_period = models.CharField(
        max_length=10, choices=TimePeriod.choices, blank=True
    )
    position = models.PositiveSmallIntegerField(null=True, blank=True)
    schedule_section = models.CharField(max_length=20, blank=True)
    week_start = models.PositiveSmallIntegerField(null=True, blank=True)
    week_end = models.PositiveSmallIntegerField(null=True, blank=True)
    year_of_study = models.PositiveSmallIntegerField(null=True, blank=True)

    class Meta:
        ordering = ["semester", "day", "start_time"]

    def __str__(self):
        start = f"{self.start_time:%H:%M}" if self.start_time else "-"
        end = f"{self.end_time:%H:%M}" if self.end_time else "-"
        return (
            f"{self.course_code} {self.get_day_display()} "
            f"{start}-{end} {self.venue}"
        )


class TechnicalDrawingAllocation(models.Model):
    semester = models.ForeignKey(
        Semester, on_delete=models.CASCADE, related_name="td_allocations"
    )
    course_code = models.CharField(max_length=20)
    group_code = models.CharField(max_length=20)
    day = models.CharField(max_length=10, choices=Day.choices)
    start_time = models.TimeField()
    end_time = models.TimeField()
    venue = models.CharField(max_length=50)

    class Meta:
        ordering = ["semester", "day", "start_time"]

    def __str__(self):
        return (
            f"{self.course_code} {self.get_day_display()} "
            f"{self.start_time:%H:%M}-{self.end_time:%H:%M} {self.venue}"
        )


class LogAction(models.TextChoices):
    CREATE = "CREATE", "Created"
    UPDATE = "UPDATE", "Updated"
    DELETE = "DELETE", "Deleted"
    IMPORT = "IMPORT", "Imported"
    ASSIGN = "ASSIGN", "Assigned"
    REMOVE = "REMOVE", "Removed"
    CLEAR = "CLEAR", "Cleared"


class ActivityLog(models.Model):
    """Entry in the activity log (what happened / changes the user made)."""

    action = models.CharField(max_length=20, choices=LogAction.choices)
    resource = models.CharField(max_length=50, blank=True)
    target = models.CharField(max_length=300, blank=True)
    message = models.CharField(max_length=500)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]
        verbose_name = "Activity log"
        verbose_name_plural = "Activity logs"

    def __str__(self):
        return self.message


class ImportStatus(models.TextChoices):
    SUCCESS = "SUCCESS", "Success"
    PARTIAL = "PARTIAL", "Completed with issues"
    FAILED = "FAILED", "Failed"


class ImportHistory(models.Model):
    """Persistent summary of every file import, reviewable after the fact.

    Counts and the structured ``details`` JSON are kept so users can inspect
    exactly what changed, what errored and what to fix in the source document
    without ever storing the imported file's contents.
    """

    import_type = models.CharField(max_length=50)
    import_title = models.CharField(max_length=100)
    filename = models.CharField(max_length=300, blank=True)
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="import_history",
    )
    status = models.CharField(
        max_length=20, choices=ImportStatus.choices, default=ImportStatus.SUCCESS
    )
    created_at = models.DateTimeField(auto_now_add=True)
    rows_processed = models.IntegerField(default=0)
    created = models.IntegerField(default=0)
    updated = models.IntegerField(default=0)
    skipped = models.IntegerField(default=0)
    error_count = models.IntegerField(default=0)
    details = models.TextField(blank=True)

    class Meta:
        ordering = ["-created_at"]
        verbose_name = "Import history"
        verbose_name_plural = "Import history"
        indexes = [models.Index(fields=["import_type", "-created_at"])]

    def __str__(self):
        return f"{self.import_title} ({self.created_at:%Y-%m-%d %H:%M})"