"""Report shared courses whose data needs a coordinator decision.

Two problems block or degrade group allocation, and neither should be fixed
silently:

* **No allocation required - requirement not configured.** A course a
  programme studies has no Seminar / Tutorial / Practical requirement, so the
  allocator will not place any group in it. It is listed so the requirement can
  be set (on the course record, or by re-importing a workbook with a
  "Required Activities" column).
* **Conflicting course names.** A course code was spelled several ways in the
  source data. One name is kept on the shared record and the alternatives are
  stored on it, so the choice can be reviewed and corrected.

Non-destructive: nothing is written, renamed or deleted. Exits 1 when anything
needs attention.
"""

import sys

from django.core.management.base import BaseCommand

from core.group_allocation import courses_missing_requirements
from core.models import Course, ProgrammeCourse, Semester


class Command(BaseCommand):
    help = (
        "List courses with no required activities configured, and course codes "
        "that were imported under several different names. Read-only."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            "--semester",
            type=int,
            default=None,
            help="Only consider courses studied in this semester number.",
        )

    def handle(self, *args, **options):
        number = options.get("semester")
        semester = None
        if number is not None:
            semester = Semester.objects.filter(semester=number).first()
            if semester is None:
                self.stderr.write(
                    self.style.ERROR(f"Semester {number} does not exist.")
                )
                sys.exit(1)

        unconfigured = courses_missing_requirements(semester)
        conflicts = [
            (course, course.name_conflicts())
            for course in Course.objects.all()
            if course.name_conflicts()
        ]

        if not unconfigured and not conflicts:
            self.stdout.write(
                self.style.SUCCESS(
                    "Every course has its required activities configured and "
                    "only one name per code."
                )
            )
            return

        for course in unconfigured:
            programmes = sorted(
                {
                    row.programme.code
                    for row in ProgrammeCourse.objects.filter(
                        course=course
                    ).select_related("programme")
                }
            )
            self.stdout.write(
                f"NO REQUIREMENT {course.code} ({course.name or 'no name'}): "
                f"studied by {', '.join(programmes) or 'no programme'}"
            )

        for course, variants in conflicts:
            self.stdout.write(
                f"NAME CONFLICT {course.code}: keeping "
                f"'{course.name or 'no name'}', also seen as "
                + "; ".join(f"'{v}'" for v in variants)
            )

        self.stderr.write(
            self.style.WARNING(
                f"{len(unconfigured)} course(s) with no required activities, "
                f"{len(conflicts)} course code(s) with conflicting names. Set "
                f"the requirements and correct the names on the course records "
                f"(or re-import the source workbook), then run the allocator."
            )
        )
        sys.exit(1)
