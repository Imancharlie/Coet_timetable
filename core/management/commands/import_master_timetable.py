import sys

from django.core.management.base import BaseCommand

from core.importers import import_master_timetable_from_excel


class Command(BaseCommand):
    help = (
        "Import the master timetable into Session records.\n"
        "Expected columns: course_code, activity_type, day, start_time, end_time\n"
        "Optional columns: venue, group / groups / group_code"
    )

    def add_arguments(self, parser):
        parser.add_argument("--file", required=True, help="Path to Excel file")
        parser.add_argument(
            "--semester",
            type=int,
            default=1,
            help="PK of the Semester record (default: 1)",
        )

    def handle(self, *args, **options):
        result = import_master_timetable_from_excel(
            options["file"], semester_id=options["semester"]
        )
        self.stdout.write(result.summary())
        if result.errors:
            sys.exit(1)