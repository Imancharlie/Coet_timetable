import sys

from django.core.management.base import BaseCommand

from core.importers import import_programme_courses_from_excel


class Command(BaseCommand):
    help = (
        "Import programme-course mappings from Excel and set each course's "
        "shared required activities.\n"
        "Required columns: programme_code, course_code, course_name, semester.\n"
        "Optional: a 'Required Activities' column (values like 'Seminar; "
        "Tutorial; Practical' or '2 practicals'), and/or separate count columns "
        "such as 'Tutorial Count' / 'Number of Practicals'. Anything else in "
        "those columns is reported, never ignored."
    )

    def add_arguments(self, parser):
        parser.add_argument("--file", required=True, help="Path to Excel file")
        parser.add_argument(
            "--requirements-column",
            dest="requirements_column",
            default=None,
            help=(
                "Exact header of the column holding the required activities. "
                "Use this when the heading is spelled in a way the importer "
                "does not recognise on its own."
            ),
        )

    def handle(self, *args, **options):
        result = import_programme_courses_from_excel(
            options["file"],
            requirements_column=options.get("requirements_column"),
        )
        self.stdout.write(result.summary())
        if result.errors:
            sys.exit(1)
