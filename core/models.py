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

    class Meta:
        ordering = ["programme__code", "course_code"]
        unique_together = ["programme", "course_code"]

    def __str__(self):
        return f"{self.programme.code} - {self.course_code}"


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