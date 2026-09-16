import sys

from django.core.management.base import BaseCommand

from core.importers import import_venues_from_excel


class Command(BaseCommand):
    help = "Import venues from Excel (columns: name, capacity)"

    def add_arguments(self, parser):
        parser.add_argument("--file", required=True, help="Path to Excel file")

    def handle(self, *args, **options):
        result = import_venues_from_excel(options["file"])
        self.stdout.write(result.summary())
        if result.errors:
            sys.exit(1)