import sys

from django.core.management.base import BaseCommand

from core.importers import import_workshop_allocation_from_excel


class Command(BaseCommand):
    help = (
        "Import workshop allocation from Excel.\n"
        "Accepts both the flat format\n"
        "  (columns: course_code, group_code, day, start_time, end_time, venue)\n"
        "and the raw university Workshop Schedule workbook (matrix layout).\n"
        "Flat files default to --semester 1 when none is given."
    )

    def add_arguments(self, parser):
        parser.add_argument("--file", required=True, help="Path to Excel file")
        parser.add_argument(
            "--semester",
            type=int,
            default=None,
            help=(
                "PK of the Semester record. Defaults to auto-detection from the "
                "workbook title for the matrix format, otherwise semester 1."
            ),
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Reconcile and count what would be imported without writing",
        )

    def handle(self, *args, **options):
        result = import_workshop_allocation_from_excel(
            options["file"],
            semester_id=options["semester"],
            dry_run=options["dry_run"],
        )
        self.stdout.write(result.summary())
        if options["dry_run"]:
            self.stdout.write("DRY RUN — no records were written.")
        if result.errors:
            sys.exit(1)