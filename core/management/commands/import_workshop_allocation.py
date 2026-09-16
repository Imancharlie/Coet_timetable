import sys

from django.core.management.base import BaseCommand

from core.importers import import_workshop_allocation_from_excel


class Command(BaseCommand):
    help = (
        "Import workshop allocation from Excel.\n"
        "Expected columns: course_code, group_code, day, start_time, end_time, venue"
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
        result = import_workshop_allocation_from_excel(
            options["file"], semester_id=options["semester"]
        )
        self.stdout.write(result.summary())
        if result.errors:
            sys.exit(1)