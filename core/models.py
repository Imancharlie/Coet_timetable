from django.conf import settings
from django.core.exceptions import ValidationError
from django.db import models


class ActivityType(models.TextChoices):
    LECTURE = "LECTURE", "Lecture"
    TUTORIAL = "TUTORIAL", "Tutorial"
    SEMINAR = "SEMINAR", "Seminar"
    PRACTICAL = "PRACTICAL", "Practical"
    WORKSHOP = "WORKSHOP", "Workshop"


#: Activities the group allocator is allowed to place groups into, in the
#: priority order the allocator searches them. Workshops are deliberately
#: excluded — they have their own allocation path — but they still count when
#: a group's availability is computed (see ``core.group_allocation``).
ALLOCATED_ACTIVITY_TYPES = (
    ActivityType.SEMINAR,
    ActivityType.TUTORIAL,
    ActivityType.PRACTICAL,
)


def normalise_course_code(value) -> str:
    """Canonical form of a course code: trimmed, internal whitespace collapsed.

    Course codes arrive from Excel with stray casing and padding (" mt161 ",
    "MT161"). Every code comparison in the shared-course model and in the
    group allocator goes through this one function so a code typed three
    different ways still names the same course. Sessions keep storing their
    own text ``course_code`` and are matched to a ``Course`` with this.
    """
    return " ".join(str(value or "").strip().split()).upper()


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


class Course(models.Model):
    """One shared record per course code, holding the global requirements.

    A course code has exactly one name and one set of required activities for
    the whole faculty: two programmes studying ``MT161`` are studying the same
    course, so the requirements cannot be per-programme. ``ProgrammeCourse``
    rows point here and only record *which programmes* study the course and in
    which semester.
    """

    code = models.CharField(max_length=20, unique=True)
    name = models.CharField(max_length=200, blank=True)
    name_variants = models.TextField(blank=True)

    class Meta:
        ordering = ["code"]

    def __str__(self):
        return f"{self.code} {self.name}".strip()

    def save(self, *args, **kwargs):
        # The code is the natural key every comparison relies on, so it is
        # normalised on the way in rather than trusted from the caller.
        self.code = normalise_course_code(self.code)
        stored = None
        if self.pk:
            stored = (
                type(self).objects.filter(pk=self.pk)
                .values_list("code", "name")
                .first()
            )
        super().save(*args, **kwargs)
        if stored is not None and stored != (self.code, self.name):
            self._push_to_programme_links()

    def _push_to_programme_links(self):
        """Re-mirror this course's code/name onto its ProgrammeCourse rows.

        The mirror columns exist so the many existing
        ``filter(course_code=...)`` queries keep working. They are only ever
        written from the shared record, so renaming a course here can never
        leave the programme links quoting a code or name that no longer
        exists. Going through ``sync_from_course`` is deliberately avoided --
        a bulk ``update()`` must not re-enter the save hooks per row.
        """
        ProgrammeCourse.objects.filter(course_id=self.pk).update(
            course_code=self.code, course_name=self.name
        )

    @property
    def requirements(self):
        """``[(activity_type, count), ...]`` in priority order (cached per
        instance)."""
        cached = getattr(self, "_requirements_cache", None)
        if cached is None:
            rows = self.activity_requirements.all()
            by_activity = {r.activity_type: r.count for r in rows}
            cached = tuple(
                (activity, by_activity[activity])
                for activity in ALLOCATED_ACTIVITY_TYPES
                if by_activity.get(activity)
            )
            self._requirements_cache = cached
        return cached

    def requirement_map(self) -> dict:
        """``{activity_type: count}`` for every configured activity."""
        return {activity: count for activity, count in self.requirements}

    def required_activities(self) -> tuple:
        """Configured activity types in priority order (blank when none)."""
        return tuple(activity for activity, _ in self.requirements)

    def required_count(self, activity_type: str) -> int:
        """How many sessions of ``activity_type`` each group needs (0 = none)."""
        return self.requirement_map().get(activity_type, 0)

    def activities_label(self) -> str:
        """"Tutorial; Practical" — for tables and forms (blank when none).

        Delegates to the one formatter in ``core.group_allocation`` so what the
        import accepts, what the pages show and what the logs record all read
        the same. Lazy-imported because that module imports this one.
        """
        from core.group_allocation import format_requirements

        return format_requirements(self.requirement_map())

    def set_requirements(self, mapping: dict, *, commit: bool = True):
        """Replace this course's requirements with ``{activity: count}``.

        Unknown activity keys and non-positive counts are ignored rather than
        raising, so a caller can hand over a parsed workbook cell directly.
        Returns the activities that were actually written.
        """
        wanted = {
            activity: int(count)
            for activity, count in (mapping or {}).items()
            if activity in ALLOCATED_ACTIVITY_TYPES and int(count or 0) > 0
        }
        self.activity_requirements.all().delete()
        written = []
        for activity in ALLOCATED_ACTIVITY_TYPES:
            count = wanted.get(activity)
            if not count:
                continue
            self.activity_requirements.create(
                activity_type=activity, count=count
            )
            written.append(activity)
        self._requirements_cache = None
        if commit:
            self.save()
        return tuple(written)

    def has_requirements(self) -> bool:
        return bool(self.requirements)

    def name_conflicts(self) -> list:
        """Other course names seen for this code before it was merged.

        A course code must have one name across the faculty. When the source
        data spelled the same course several ways ("Communication Skills for
        Engineers" vs "Communication Skills for Engineering") the migration
        picked one deterministically and kept the rest here, so the coordinator
        can review the choice instead of it being silently resolved.
        """
        import json

        if not self.name_variants:
            return []
        try:
            variants = json.loads(self.name_variants)
        except (ValueError, TypeError):
            return []
        return [v for v in variants if isinstance(v, str)]


class CourseActivityRequirement(models.Model):
    """How many sessions of one activity every group must attend for a course.

    One row per configured activity, so a course requiring a tutorial *and* a
    practical has two rows and a course requiring two practicals has one row
    with ``count = 2``. A course with no rows requires no allocation at all —
    which the allocation review flags for the coordinator rather than guessing.
    """

    course = models.ForeignKey(
        Course, on_delete=models.CASCADE, related_name="activity_requirements"
    )
    activity_type = models.CharField(
        max_length=20, choices=ActivityType.choices
    )
    count = models.PositiveSmallIntegerField(default=1)

    class Meta:
        ordering = ["course__code", "activity_type"]
        unique_together = ["course", "activity_type"]

    def __str__(self):
        plural = "" if self.count == 1 else "s"
        return f"{self.course.code} {self.count}x {ActivityType(self.activity_type).label}{plural}"

    def clean(self):
        super().clean()
        if self.activity_type not in ALLOCATED_ACTIVITY_TYPES:
            raise ValidationError(
                {
                    "activity_type": "Required activities must be Seminar, "
                    "Tutorial or Practical."
                }
            )
        if self.count < 1:
            raise ValidationError({"count": "A requirement count must be 1 or more."})


class ProgrammeCourse(models.Model):
    """A programme studies a shared course, in a given semester.

    ``course_code`` and ``course_name`` are kept as plain columns so the many
    existing ``filter(course_code=...)`` queries, exports and templates keep
    working, but they are *mirrors*: :meth:`save` copies them off the shared
    ``Course`` every time, so the shared record stays the single source of
    truth and neither column can be edited on its own.
    """

    programme = models.ForeignKey(
        Programme, on_delete=models.CASCADE, related_name="programme_courses"
    )
    course = models.ForeignKey(
        Course,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name="programme_courses",
    )
    course_code = models.CharField(max_length=20)
    course_name = models.CharField(max_length=200, blank=True)
    semester = models.PositiveSmallIntegerField()

    class Meta:
        ordering = ["programme__code", "semester", "course_code"]
        unique_together = ["programme", "course_code"]

    def __str__(self):
        return f"{self.programme.code} - {self.course_code} {self.course_name}"

    def sync_from_course(self):
        """Copy code/name off the shared course into the mirror columns."""
        if self.course_id is None:
            return
        code = normalise_course_code(self.course.code)
        name = (self.course.name or "").strip()
        if self.course_code != code:
            self.course_code = code
        if self.course_name != name:
            self.course_name = name

    def save(self, *args, **kwargs):
        if self.course_id is None:
            # A link created without naming its shared course (older code, a
            # fixture, a script) would leave the course with no requirements
            # and be invisible to the allocator. Attach it to a Course of the
            # right code rather than leave an orphan that cannot be reviewed.
            self.attach_course()
        self.sync_from_course()
        super().save(*args, **kwargs)

    def attach_course(self):
        """Point this link at the shared course for its code, creating one."""
        code = normalise_course_code(self.course_code)
        if not code:
            return None
        course, _ = Course.objects.get_or_create(
            code=code, defaults={"name": (self.course_name or "").strip()}
        )
        self.course = course
        return course

    @property
    def required_activities(self) -> str:
        """The shared course's requirements as text, for lists and details."""
        course = self.course
        if course is None:
            return "Not configured"
        return course.activities_label() or "Not configured"

    @property
    def has_requirements(self) -> bool:
        return bool(self.course and self.course.has_requirements())


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

    def clean(self):
        """Workshop sessions run only at the standard workshop times.

        Non-workshop sessions are never affected. Enforcement is lazy-imported
        so loading ``core.workshop_times`` (which imports this module) can
        never form an import cycle.
        """
        super().clean()
        if self.activity_type != "WORKSHOP" or not self.day:
            return
        from core.workshop_times import validate_workshop_session

        errors = validate_workshop_session(self)
        if errors:
            raise ValidationError({
                field: msgs
                for error in errors
                for field, msgs in error.error_dict.items()
            })


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

    def clean(self):
        """Workshop allocations keep the standard workshop session times.

        Records imported from the raw workshop matrix carry a period but no
        clock times and stay valid; manually entered times must equal one whole
        standard session for the day. Enforcement is lazy-imported so loading
        ``core.workshop_times`` (which imports this module) never forms an
        import cycle.
        """
        super().clean()
        if not self.day:
            return
        from core.workshop_times import validate_workshop_record

        errors = validate_workshop_record(self)
        if errors:
            raise ValidationError({
                field: msgs
                for error in errors
                for field, msgs in error.error_dict.items()
            })


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


class AllocationStatus(models.TextChoices):
    PREVIEWED = "PREVIEWED", "Previewed"
    APPLIED = "APPLIED", "Applied"
    REVERTED = "REVERTED", "Reverted"


class AllocationScope(models.TextChoices):
    SEMINAR = "SEMINAR", "Seminar"
    TUTORIAL = "TUTORIAL", "Tutorial"
    PRACTICAL = "PRACTICAL", "Practical"
    ALL = "ALL", "All configured activities"


class AllocationRun(models.Model):
    """One saved pass of the group allocator, kept so it can be reviewed and
    reverted long after the request that produced it ended.

    The plan itself lives in :attr:`summary` (a JSON snapshot: every proposed
    assignment, every unresolved reason, the warnings and the counts), and the
    individual before/after link states live in
    :class:`AllocationChange` so a revert can check that nothing has been
    edited since before it rewrites anything.
    """

    semester = models.ForeignKey(
        Semester, on_delete=models.CASCADE, related_name="allocation_runs"
    )
    scope = models.CharField(
        max_length=20,
        choices=AllocationScope.choices,
        default=AllocationScope.ALL,
    )
    status = models.CharField(
        max_length=20,
        choices=AllocationStatus.choices,
        default=AllocationStatus.PREVIEWED,
    )
    created_at = models.DateTimeField(auto_now_add=True)
    applied_at = models.DateTimeField(null=True, blank=True)
    reverted_at = models.DateTimeField(null=True, blank=True)
    search_limit_hit = models.BooleanField(default=False)
    assigned = models.IntegerField(default=0)
    moved = models.IntegerField(default=0)
    retained = models.IntegerField(default=0)
    removed = models.IntegerField(default=0)
    unresolved = models.IntegerField(default=0)
    summary = models.TextField(blank=True)

    class Meta:
        ordering = ["-created_at"]
        verbose_name = "Allocation run"
        verbose_name_plural = "Allocation runs"
        indexes = [models.Index(fields=["semester", "-created_at"])]

    def __str__(self):
        return f"{self.semester} — {self.get_scope_display()} ({self.status})"

    def plan(self) -> dict:
        """The stored :class:`core.group_allocation.AllocationPlan` snapshot."""
        import json

        try:
            return json.loads(self.summary or "{}")
        except (ValueError, TypeError):
            return {}

    @property
    def is_applied(self) -> bool:
        return self.status == AllocationStatus.APPLIED

    @property
    def change_count(self) -> int:
        return self.changes.count()


class AllocationChange(models.Model):
    """One SessionGroup link an :class:`AllocationRun` added or removed.

    ``before``/``after`` each hold ``None`` (the link was absent/present) or a
    ``{session_pk, session_label, course_code, activity_type, day,
    start_time, end_time, venue}`` snapshot. Reverting replays them, but only
    after checking that the link is still in the state this run left it in —
    a later manual edit is reported as a conflict instead of being overwritten.
    """

    class Action(models.TextChoices):
        ADD = "ADD", "Added"
        REMOVE = "REMOVE", "Removed"

    run = models.ForeignKey(
        AllocationRun, on_delete=models.CASCADE, related_name="changes"
    )
    group = models.ForeignKey(
        StudentGroup, on_delete=models.CASCADE, related_name="allocation_changes"
    )
    session = models.ForeignKey(
        Session,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="allocation_changes",
    )
    action = models.CharField(max_length=10, choices=Action.choices)
    course_code = models.CharField(max_length=20, blank=True)
    activity_type = models.CharField(max_length=20, blank=True)
    before = models.TextField(blank=True)
    after = models.TextField(blank=True)

    class Meta:
        ordering = ["run", "course_code", "activity_type", "group__code"]
        indexes = [models.Index(fields=["run", "action"])]

    def __str__(self):
        return f"{self.run_id}: {self.action} {self.group_id} -> {self.session_id}"

    def before_state(self) -> dict | None:
        import json

        return _load_state(self.before)

    def after_state(self) -> dict | None:
        import json

        return _load_state(self.after)


def _load_state(raw: str):
    """Decode a stored before/after link snapshot (``''`` means "no link")."""
    import json

    if not raw:
        return None
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        return None