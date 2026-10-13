"""Build the synthetic dataset and print its statistics.

Runs inside a **throwaway test database** (Django's in-memory SQLite), so
``db.sqlite3`` is never touched. That makes it safe to run against a real
installation, and it is the same fixture the audit's own tests use.

    python manage.py audit_fixture            # stats only
    python manage.py audit_fixture --plan      # ...and the generated plan
"""

from django.core.management.base import BaseCommand
from django.test.utils import setup_databases, setup_test_environment, teardown_databases


class Command(BaseCommand):
    help = "Create the synthetic audit dataset in a temporary test database."

    def add_arguments(self, parser):
        parser.add_argument(
            "--year",
            default="2024/25",
            help="Academic year for the synthetic semester (default: 2024/25).",
        )
        parser.add_argument(
            "--semester", type=int, default=1, help="Semester number (default: 1)."
        )
        parser.add_argument(
            "--plan",
            action="store_true",
            help="Also print the full allocation plan snapshot.",
        )

    def handle(self, *args, **options):
        from audit.fixtures import generate

        setup_test_environment()
        databases = setup_databases(verbosity=0, interactive=False)
        try:
            stats = generate.build(
                academic_year=options["year"], semester_number=options["semester"]
            )
            width = max(len(k) for k in stats)
            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING("Synthetic fixture"))
            self.stdout.write("-" * (width + 22))
            for key, value in stats.items():
                if key == "planted":
                    continue
                self.stdout.write(f"  {key:<{width}}  {value}")
            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING("Planted problems"))
            self.stdout.write("-" * (width + 22))
            for key, value in stats["planted"].items():
                self.stdout.write(f"  {key:<{width}}  {value}")

            from core.models import AllocationRun

            run = AllocationRun.objects.get(pk=stats["run_id"])
            snapshot = run.plan()
            self.stdout.write("")
            self.stdout.write(
                f"Run #{run.pk}: {snapshot['requirement_total']} requirements, "
                f"{snapshot['added']} added, {snapshot['retained']} retained, "
                f"{snapshot['unresolved']} unresolved, "
                f"{snapshot['scanned']:,} candidates examined, "
                f"{snapshot['duration_ms']} ms"
            )
            if options["plan"]:
                self.stdout.write("")
                for row in snapshot["assignments"][:10]:
                    self.stdout.write(
                        f"  {row['group']:<8} {row['course']:<7} {row['activity']:<10}"
                        f" {row['day']:<10} {row['time']:<13} {row['status']}"
                    )
            self.stdout.write("")
            self.stdout.write(
                self.style.SUCCESS("Temporary database discarded; db.sqlite3 untouched.")
            )
        finally:
            teardown_databases(databases, verbosity=0)
