import uuid

from django.db import models

from core.models import Day, Programme, Semester, StudentGroup


class PortalSettings(models.Model):
    """Portal-wide settings configured by staff in Django admin."""

    current_semester = models.ForeignKey(
        Semester,
        on_delete=models.PROTECT,
        related_name="current_for_portals",
        help_text="Semester selected by default on the public student portal.",
    )

    class Meta:
        verbose_name = "Student portal settings"
        verbose_name_plural = "Student portal settings"

    def __str__(self):
        return f"Student portal — {self.current_semester}"

    def save(self, *args, **kwargs):
        self.pk = 1
        return super().save(*args, **kwargs)


class CollisionReport(models.Model):
    class TimetableType(models.TextChoices):
        TEACHING = "TEACHING", "Teaching timetable"
        EXAMINATION = "EXAMINATION", "Examination timetable"

    class Status(models.TextChoices):
        NEW = "NEW", "New"
        REVIEWING = "REVIEWING", "Under review"
        RESOLVED = "RESOLVED", "Resolved"

    reference = models.UUIDField(default=uuid.uuid4, unique=True, editable=False)
    timetable_type = models.CharField(max_length=16, choices=TimetableType.choices)
    semester = models.ForeignKey(Semester, on_delete=models.PROTECT)
    programme = models.ForeignKey(
        Programme, on_delete=models.SET_NULL, null=True, blank=True
    )
    group = models.ForeignKey(
        StudentGroup, on_delete=models.SET_NULL, null=True, blank=True
    )
    course_or_exam = models.CharField(max_length=100, blank=True)
    day = models.CharField(max_length=10, choices=Day.choices, blank=True)
    time_description = models.CharField(max_length=80, blank=True)
    description = models.TextField()
    contact_email = models.EmailField(blank=True)
    status = models.CharField(max_length=12, choices=Status.choices, default=Status.NEW)
    staff_notes = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-created_at"]

    def __str__(self):
        return f"{self.reference} — {self.get_timetable_type_display()}"
