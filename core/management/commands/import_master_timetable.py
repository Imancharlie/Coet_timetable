import sys

from django.core.management.base import BaseCommand

from core.importers import import_master_timetable_from_excel


class Command(BaseCommand):
    help = (
        "Import the master timetable into Session records.\n"
        "Requires --semester: the PK of the academic Semester record the "
        "timetable belongs to (never auto-created).\n"
        "Expected columns: course_code, activity_type, day, start_time, end_time\n"
        "Optional columns: venue, group / groups / group_code\n"
        "A reconciliation report is always produced first: missing reference "
        "data is reported instead of guessed at, and 'ALL' group rows are "
        "expanded through the ProgrammeCourse mapping. LECTURE sessions are "
        "automatically linked to every programme group that studies the course.\n"
    )

    def add_arguments(self, parser):
        parser.add_argument("--file", required=True, help="Path to Excel file")
        parser.add_argument(
            "--semester",
            type=int,
            required=True,
            help="PK of the Semester record this timetable belongs to",
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Reconcile and report without writing anything to the database",
        )

    def handle(self, *args, **options):
        result = import_master_timetable_from_excel(
            options["file"],
            semester_id=options["semester"],
            dry_run=options["dry_run"],
        )
        if options["dry_run"]:
            self.stdout.write("DRY RUN — no records were written.\n")
        self.stdout.write(result.summary())
        if result.errors or result.conflicts:
            sys.exit(1)