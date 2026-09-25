import sys

from django.core.management.base import BaseCommand

from core.workshop_times import (
    legacy_workshop_allocations,
    legacy_workshop_sessions,
)


class Command(BaseCommand):
    help = (
        "Identify workshop records that do not use the standard workshop "
        "session times (Monday-Friday morning 09:00-13:00 / afternoon "
        "15:00-19:00; Thursday morning 10:00-14:00).\n"
        "Non-destructive: nothing is written or deleted. Exits 1 only when "
        "legacy records are found."
    )

    def handle(self, *args, **options):
        allocations = legacy_workshop_allocations()
        sessions = legacy_workshop_sessions()

        if not allocations and not sessions:
            self.stdout.write(
                self.style.SUCCESS(
                    "No legacy workshop records found. All workshops use "
                    "the standard session times."
                )
            )
            return

        for rec, issue in allocations:
            time_label = (
                f"{rec.start_time:%H:%M}-{rec.end_time:%H:%M}"
                if rec.start_time and rec.end_time
                else (
                    f"{rec.get_time_period_display()}"
                    if rec.time_period
                    else "no time"
                )
            )
            self.stdout.write(
                f"ALLOCATION {rec.course_code} ({rec.group_code}) "
                f"{rec.get_day_display()} {time_label}: {issue}"
            )

        for session, issue in sessions:
            self.stdout.write(
                f"SESSION {session.course_code} "
                f"{session.get_day_display()} "
                f"{session.start_time:%H:%M}-{session.end_time:%H:%M}: {issue}"
            )

        self.stderr.write(
            self.style.WARNING(
                f"{len(allocations)} allocation(s), "
                f"{len(sessions)} session(s) found."
            )
        )
        sys.exit(1)