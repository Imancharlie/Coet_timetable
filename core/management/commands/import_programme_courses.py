import sys

from django.core.management.base import BaseCommand

from core.importers import import_programme_courses_from_excel


class Command(BaseCommand):
    help = "Import programme-course mappings from Excel (columns: programme_code, course_code)"

    def add_arguments(self, parser):
        parser.add_argument("--file", required=True, help="Path to Excel file")

    def handle(self, *args, **options):
        result = import_programme_courses_from_excel(options["file"])
        self.stdout.write(result.summary())
        if result.errors:
            sys.exit(1)