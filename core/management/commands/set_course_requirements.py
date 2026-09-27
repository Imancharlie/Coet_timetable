"""Set each course's required activities from a small file, or from the timetable.

The programme-course import reads requirements from its own workbook, which is
fine when that workbook has a "Required Activities" column — but it is the
wrong tool when it does not, and wrong too when all that is wanted is to
attach requirements to courses that are already in the database. This command
does that and nothing else: it never creates a course, never touches programme
links, and never touches the timetable.

Two sources:

``--file``
    A two-column sheet, ``course_code`` and ``Required Activities``. The same
    parsing rules as the import apply, so the file you already learned for the
    import works here unchanged — and the count-column layout ("Tutorial
    Count") is understood too. A blank cell means "this course requires no
    allocation", which is a statement, and clearing a course that had
    requirements is reported as ``CLEARED`` rather than passing unnoticed.

``--from-sessions``
    Derives the requirement for every course from the seminar / tutorial /
    practical sessions that already exist in the master timetable. A course
    that has sessions for an activity almost certainly requires one of them,
    so this is the fastest way to get from "nothing configured" to a plan worth
    reviewing. It only ever sets an activity that has at least one session, so
    it cannot invent a requirement nothing can satisfy. Courses with no
    small-group session at all are reported as skipped — deciding those is a
    judgement call, not a derivation.

Both accept ``--dry-run``, which reports every change and writes nothing.
"""

import sys

import pandas as pd
from django.core.management.base import BaseCommand

from core.group_allocation import (
    find_requirement_count_columns,
    find_requirements_column,
    format_requirements,
)
from core.importers import set_requirements_for_row
from core.models import ALLOCATED_ACTIVITY_TYPES, Course, Session, normalise_course_code

CODE_ALIASES = (
    "course_code", "course code", "code", "subject_code", "unit_code",
    "subject",
)


def _match_code_column(df):
    """The course-code column, matched on letters and digits like the imports."""
    for alias in CODE_ALIASES:
        key = "".join(c.lower() for c in alias if c.isalnum())
        for column in df.columns:
            if "".join(c.lower() for c in str(column) if c.isalnum()) == key:
                return column
    return None


def _read(path):
    """Read a CSV or Excel sheet, trimmed to strings."""
    text = str(path).lower()
    if text.endswith((".csv", ".txt")):
        return pd.read_csv(path, dtype=str).fillna("")
    return pd.read_excel(path, dtype=str).fillna("")


class Command(BaseCommand):
    help = (
        "Set the seminar / tutorial / practical requirements for courses that "
        "already exist.\n"
        "  --file <path>          a sheet with 'course_code' and "
        "'Required Activities'\n"
        "  --from-sessions        derive each requirement from the sessions "
        "that already exist\n"
        "  --requirements-column  exact header, when the heading is unusual\n"
        "  --dry-run              report every change, write nothing\n"
        "Exits 1 when a row could not be applied."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            "--file",
            default=None,
            help="CSV or Excel file with course_code and required activities",
        )
        parser.add_argument(
            "--from-sessions",
            action="store_true",
            help=(
                "Derive each course's requirement from the seminar/tutorial/"
                "practical sessions already in the timetable"
            ),
        )
        parser.add_argument(
            "--requirements-column",
            dest="requirements_column",
            default=None,
            help="Exact header of the column holding the required activities",
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Report what would change without writing anything",
        )

    def handle(self, *args, **options):
        if bool(options["file"]) == bool(options["from_sessions"]):
            self.stderr.write(
                self.style.ERROR(
                    "Give exactly one of --file or --from-sessions."
                )
            )
            sys.exit(1)

        if options["from_sessions"]:
            changes, skipped, problems = self._from_sessions(options["dry_run"])
        else:
            changes, skipped, problems = self._from_file(
                options["file"],
                options["requirements_column"],
                options["dry_run"],
            )

        self._report(changes, skipped, problems, options["dry_run"])
        if problems:
            sys.exit(1)

    # -- sources ------------------------------------------------------
    def _from_sessions(self, dry_run):
        """Derive a requirement per course from the sessions that exist."""
        available = {}
        for code, activity in Session.objects.values_list(
            "course_code", "activity_type"
        ):
            if activity not in ALLOCATED_ACTIVITY_TYPES:
                continue
            available.setdefault(
                normalise_course_code(code), set()
            ).add(activity)

        courses = list(
            Course.objects.prefetch_related("activity_requirements").all()
        )
        changes, skipped, problems = [], [], []
        for course in courses:
            present = [
                activity
                for activity in ALLOCATED_ACTIVITY_TYPES
                if activity in available.get(normalise_course_code(course.code), ())
            ]
            if not present:
                skipped.append(
                    f"{course.code}: no seminar/tutorial/practical session "
                    f"exists, so nothing was derived — decide whether the "
                    f"course needs any, or whether its sessions are still to "
                    f"be timetabled"
                )
                continue
            changes.append((course, {a: 1 for a in present}))
        return self._apply(changes, dry_run), skipped, problems

    def _from_file(self, path, explicit_column, dry_run):
        try:
            df = _read(path)
        except Exception as exc:  # pragma: no cover - surfaced to the user
            return [], [], [f"Could not read {path}: {exc}"]

        code_col = _match_code_column(df)
        if code_col is None:
            return [], [], [
                f"No 'course_code' column found in {path} "
                f"(found: {', '.join(map(str, df.columns))})"
            ]
        activities_col, unrecognised = find_requirements_column(
            df.columns, explicit=explicit_column
        )
        if explicit_column and not activities_col:
            return [], [], [
                f"Required-activities column '{explicit_column}' not found "
                f"(available: {', '.join(map(str, df.columns))})"
            ]
        if not activities_col:
            return [], [], [
                "No required-activities column found "
                f"(available: {', '.join(map(str, df.columns))}). Rename the "
                "heading to 'Required Activities', or pass "
                "--requirements-column '<the exact header>'."
            ]
        count_columns = find_requirement_count_columns(
            df.columns, exclude=[activities_col]
        )
        self.stdout.write(
            f"Reading '{activities_col}'"
            + (
                " plus count columns "
                + ", ".join(f"'{c}'" for c in count_columns.values())
                if count_columns
                else ""
            )
        )

        changes, skipped, problems = [], [], []
        for _, row in df.iterrows():
            code = normalise_course_code(row[code_col])
            if not code:
                problems.append("A row has no course code — skipped")
                continue
            course = Course.objects.filter(code=code).first()
            if course is None:
                problems.append(
                    f"Course '{code}' is not in the database — create it "
                    f"first (Courses -> Add New, or import programme "
                    f"courses); its requirements were not set"
                )
                continue
            requirements, issues = set_requirements_for_row(
                row, activities_col, count_columns
            )
            if issues:
                problems.extend(f"{code}: {issue}" for issue in issues)
                continue
            changes.append((course, requirements))
        return self._apply(changes, dry_run), skipped, problems

    # -- writing ------------------------------------------------------
    def _apply(self, changes, dry_run):
        """Write (or just report) each change, skipping ones already correct."""
        applied, unchanged = [], 0
        for course, requirements in changes:
            if course.requirement_map() == requirements:
                unchanged += 1
                continue
            was = course.activities_label() or "not configured"
            now = format_requirements(requirements) or "not configured"
            if not dry_run:
                course.set_requirements(requirements)
            cleared = bool(was != "not configured") and not requirements
            applied.append((course, was, now, cleared))
        return {"applied": applied, "unchanged": unchanged}

    def _report(self, changes, skipped, problems, dry_run):
        applied = changes["applied"]
        label = "Would set" if dry_run else "Set"
        if applied:
            self.stdout.write(
                self.style.SUCCESS(
                    f"{label} {len(applied)} course requirement(s); "
                    f"{changes['unchanged']} already correct."
                )
            )
            for course, was, now, cleared in applied:
                marker = "CLEARED " if cleared else "       "
                self.stdout.write(f"  {marker}{course.code}: {was} -> {now}")
        else:
            self.stdout.write(
                f"No changes needed "
                f"({changes['unchanged']} course(s) already correct)."
            )
        if skipped:
            self.stdout.write(
                self.style.WARNING(
                    f"{len(skipped)} course(s) left alone:"
                )
            )
            for line in skipped:
                self.stdout.write(f"  - {line}")
        if problems:
            self.stderr.write(
                self.style.ERROR(f"{len(problems)} problem(s):")
            )
            for line in problems:
                self.stderr.write(f"  - {line}")
        if dry_run and (applied or skipped):
            self.stdout.write(
                self.style.WARNING("DRY RUN - no records were written.")
            )
