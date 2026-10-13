import json
import re
from dataclasses import dataclass, field
from datetime import date, time
from pathlib import Path

import pandas as pd

from core.models import (
    ActivityType,
    Course,
    Day,
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    TimePeriod,
    Venue,
    WorkshopAllocation,
    normalise_course_code,
)
from core.workshop_times import (
    course_programme_codes,
    programme_codes_for_group_codes,
    workshop_time_issue,
)
from core.venue_quality import (
    adopt_official_name,
    base_key,
    detect_name_conflicts,
    fold_formatting_duplicates,
    formatting_key,
    kept_capacity,
    official_formatting_name,
    venue_by_formatting,
)
from core.workshop_parser import (
    detect_format,
    parse_workbook,
)


def _master_workshop_programme_codes(rec):
    """Programme code(s) a master-timetable WORKSHOP row belongs to.

    A literal group list resolves via the StudentGroup rows (unknown groups
    fall back to the default 10:00 Thursday rule); the legacy ``ALL`` group
    resolves via the programmes that study the course.
    """
    if str(rec["raw_groups"]).strip().upper() == "ALL":
        return course_programme_codes(rec["course_code"])
    group_codes = [
        g.strip()
        for g in str(rec["raw_groups"]).split(",")
        if g.strip() and g.strip().upper() != "ALL"
    ]
    return programme_codes_for_group_codes(group_codes)

DAY_MAP = {
    "mon": Day.MONDAY,
    "tue": Day.TUESDAY,
    "wed": Day.WEDNESDAY,
    "thu": Day.THURSDAY,
    "fri": Day.FRIDAY,
    "sat": Day.SATURDAY,
    "sun": Day.SUNDAY,
}


def normalise_day(raw: str) -> str:
    raw = str(raw).strip().upper()
    if raw in dict(Day.choices):
        return raw
    return DAY_MAP.get(raw[:3].lower(), raw.upper())


def parse_time(value) -> time:
    if isinstance(value, pd.Timestamp):
        return value.time()
    text = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p"):
        try:
            return time.fromisoformat(text.split(".")[0] if "." in text else text[:8])
        except ValueError:
            continue
    return time.fromisoformat(text[:8])


def normalise_activity_type(raw: str) -> str:
    upper = str(raw).strip().upper()
    mapping = {
        "LECTURE": ActivityType.LECTURE,
        "TUTORIAL": ActivityType.TUTORIAL,
        "SEMINAR": ActivityType.SEMINAR,
        "PRACTICAL": ActivityType.PRACTICAL,
        "WORKSHOP": ActivityType.WORKSHOP,
    }
    return mapping.get(upper, upper)


def _norm_col(name) -> str:
    """Normalise a header for alias matching: lowercase alphanumerics only."""
    return "".join(ch.lower() for ch in str(name) if ch.isalnum())


def _match_col(df, *aliases):
    """Return the real column matching any alias (case/punctuation-insensitive)."""
    norm = {}
    for c in df.columns:
        norm.setdefault(_norm_col(c), c)
    for alias in aliases:
        key = _norm_col(alias)
        if key in norm:
            return norm[key]
    return None


def _split_csv(value: str) -> list[str]:
    """Split a comma/semicolon separated cell into stripped, non-empty parts."""
    return [
        p.strip()
        for p in str(value).replace(";", ",").split(",")
        if p.strip()
    ]


def _split_range(value):
    """Split a '09:00-12:00' time range into (start, end) strings."""
    text = (
        str(value)
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace(" to ", "-")
    )
    parts = [p.strip() for p in text.split("-") if p.strip()]
    start = parts[0] if parts else ""
    end = parts[1] if len(parts) > 1 else ""
    return start, end


def _fill_down(values) -> list[str]:
    """Forward-fill blank cells produced by merged/pivoted Excel layouts."""
    out = []
    last = ""
    for v in values:
        s = "" if pd.isna(v) else str(v).strip()
        if s:
            last = s
        out.append(last)
    return out


_FILLER_WORDS = {
    "bsc",
    "bsc",
    "bachelor",
    "of",
    "in",
    "and",
    "the",
    "science",
    "sciences",
    "technology",
}


def _derive_code(name: str) -> str:
    """Build a programme code from a full programme name (e.g. 'BSc. in Chemical
    and Processing Engineering' -> 'CPE'). Existing-candidate codes are skipped."""
    words = re.split(r"[^A-Za-z0-9]+", name)
    significant = [w for w in words if w and w.lower() not in _FILLER_WORDS]
    letters = "".join(w[0].upper() for w in significant[:4])
    if len(letters) < 2 and significant:
        letters = significant[0][:2].upper()
    if not letters:
        letters = "PRG"
    base = letters[:20]
    n = 1
    while Programme.objects.filter(code__iexact=base).exists():
        n += 1
        base = f"{letters[:17]}{n}"
    return base


def _default_reference_semester():
    """The semester used by imports that do not need one supplied explicitly."""
    year = date.today().year
    semester, _ = Semester.objects.get_or_create(
        academic_year=f"{year}/{year + 1}", semester=1
    )
    return semester


@dataclass
class ImportResult:
    created: int = 0
    updated: int = 0
    skipped: int = 0
    errors: list = field(default_factory=list)
    matched_courses: list = field(default_factory=list)
    missing_courses: list = field(default_factory=list)
    missing_groups: list = field(default_factory=list)
    missing_venues: list = field(default_factory=list)
    all_rows: list = field(default_factory=list)
    conflicts: list = field(default_factory=list)
    format: str = ""
    detected_semester: str = ""
    missing_references: list = field(default_factory=list)
    unknown_keys: list = field(default_factory=list)
    invalid_positions: list = field(default_factory=list)
    ambiguous: list = field(default_factory=list)
    duplicates: list = field(default_factory=list)
    unrecognized_cells: list = field(default_factory=list)
    programmes_created: list = field(default_factory=list)
    lecture_sessions_processed: int = 0
    lecture_groups_linked: int = 0
    lecture_groups_existing: int = 0
    lecture_skipped_courses: list = field(default_factory=list)
    venue_resolutions: list = field(default_factory=list)
    venue_conflicts: list = field(default_factory=list)
    venue_capacity_issues: list = field(default_factory=list)
    invalid_requirements: list = field(default_factory=list)
    courses_created: list = field(default_factory=list)
    courses_updated: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    course_requirements: list = field(default_factory=list)
    requirements_column: str = ""
    requirement_count_columns: dict = field(default_factory=dict)
    derived_requirements: list = field(default_factory=list)
    derived_unresolved: list = field(default_factory=list)
    derived_skipped: list = field(default_factory=list)

    @property
    def total(self):
        return self.created + self.updated + self.skipped

    @property
    def venue_exact_count(self):
        return sum(1 for resolution in self.venue_resolutions if resolution["action"] == "exact")

    @property
    def venue_matched_count(self):
        return sum(1 for resolution in self.venue_resolutions if resolution["action"] == "matched")

    @property
    def venue_created_count(self):
        return sum(1 for resolution in self.venue_resolutions if resolution["action"] == "created")

    def summary(self) -> str:
        lines = [
            f"Created: {self.created}",
            f"Updated: {self.updated}",
            f"Skipped: {self.skipped}",
        ]
        if self.format:
            lines.append(f"Detected format: {self.format}")
        if self.detected_semester:
            lines.append(f"Detected semester: {self.detected_semester}")
        if self.requirements_column:
            lines.append(
                f"Required activities read from column "
                f"'{self.requirements_column}'"
            )
        if self.requirement_count_columns:
            lines.append(
                "Per-activity count columns: "
                + ", ".join(
                    f"{column} -> {ActivityType(activity).label}"
                    for activity, column in self.requirement_count_columns.items()
                )
            )
        if self.warnings:
            for warning in self.warnings:
                lines.append(f"Warning: {warning}")
        if self.notes:
            for note in self.notes:
                lines.append(f"Noted: {note}")
        if self.programmes_created:
            lines.append(
                f"Programmes created automatically: "
                f"{', '.join(self.programmes_created[:20])}"
            )
        if self.courses_created:
            lines.append(
                f"Courses created: {', '.join(self.courses_created[:20])}"
            )
        if self.courses_updated:
            lines.append(
                f"Courses updated: {', '.join(self.courses_updated[:20])}"
            )
        if self.invalid_requirements:
            lines.append(
                f"Invalid required activities: {len(self.invalid_requirements)} "
                f"row(s) - the course was left with no requirement set"
            )
            for issue in self.invalid_requirements[:10]:
                lines.append(f"  - {issue}")
        if self.derived_requirements:
            lines.append(
                f"Required activities derived from the timetable: "
                f"{len(self.derived_requirements)} course(s)"
            )
            for entry in self.derived_requirements[:20]:
                lines.append(f"  - {entry['code']}: {entry['requirements']}")
            if len(self.derived_requirements) > 20:
                lines.append(
                    f"  ... and "
                    f"{len(self.derived_requirements) - 20} more"
                )
        if self.derived_skipped:
            lines.append(
                f"Requirements already configured, left unchanged: "
                f"{len(self.derived_skipped)}"
            )
        if self.derived_unresolved:
            lines.append(
                f"Nothing derived, no small-group session exists: "
                f"{len(self.derived_unresolved)}"
            )
            for issue in self.derived_unresolved[:10]:
                lines.append(f"  - {issue}")
            if len(self.derived_unresolved) > 10:
                lines.append(
                    f"  ... and {len(self.derived_unresolved) - 10} more"
                )
        if self.lecture_sessions_processed:
            lines.append(
                f"Lecture groups assigned: {self.lecture_sessions_processed} "
                f"session(s), {self.lecture_groups_linked} link(s) added, "
                f"{self.lecture_groups_existing} already linked"
            )
        if self.matched_courses:
            lines.append(
                f"Reconciled courses: {', '.join(self.matched_courses[:20])}"
            )
        if self.missing_courses:
            lines.append(f"Missing courses: {', '.join(self.missing_courses)}")
        if self.missing_groups:
            lines.append(f"Missing groups: {', '.join(self.missing_groups)}")
        if self.missing_venues:
            lines.append(
                f"Venues created with capacity 0: {', '.join(self.missing_venues)}"
            )
        if self.venue_resolutions:
            exact = sum(1 for r in self.venue_resolutions if r["action"] == "exact")
            matched = sum(1 for r in self.venue_resolutions if r["action"] == "matched")
            created = sum(1 for r in self.venue_resolutions if r["action"] == "created")
            lines.append(
                f"Venue resolutions: {exact} exact, {matched} smart-matched, {created} created"
            )
        if self.missing_references:
            lines.append(
                f"Missing references: {', '.join(self.missing_references[:20])}"
            )
        if self.all_rows:
            lines.append(f"'ALL' group expansions: {len(self.all_rows)}")
        if self.conflicts:
            lines.append(f"Conflicts: {len(self.conflicts)}")
        if self.venue_conflicts:
            lines.append(
                f"Venue name conflicts: {len(self.venue_conflicts)} "
                f"(spacing/casing duplicates — resolve them interactively in "
                f"the web import result panel)"
            )
        if self.venue_capacity_issues:
            lines.append(
                f"Venues created with capacity 0 (missing capacity in source): "
                f"{', '.join(self.venue_capacity_issues)} — set a capacity "
                f"before finalising"
            )
        if self.unknown_keys:
            lines.append(f"Unknown KEY positions: {len(self.unknown_keys)}")
        if self.invalid_positions:
            lines.append(f"Invalid positions: {len(self.invalid_positions)}")
        if self.ambiguous:
            lines.append(f"Ambiguous entries: {len(self.ambiguous)}")
        if self.duplicates:
            lines.append(f"Duplicate keys: {len(self.duplicates)}")
        if self.unrecognized_cells:
            lines.append(f"Unrecognized cells: {len(self.unrecognized_cells)}")
        if self.errors:
            lines.append(f"Errors: {len(self.errors)}")
            for err in self.errors[:20]:
                lines.append(f"  - {err}")
            if len(self.errors) > 20:
                lines.append(f"  ... and {len(self.errors) - 20} more")
        return "\n".join(lines)

    def snapshot(self) -> dict:
        """A JSON-serialisable copy of everything worth reviewing later.

        Used to persist import history: errors (with their row/field wording),
        missing/invalid entities, conflicts, automatic fixes (venue matches,
        programmes created, lecture-group assignment) and source details.
        Never includes the file itself.
        """
        return {
            "created": self.created,
            "updated": self.updated,
            "skipped": self.skipped,
            "error_count": len(self.errors),
            "errors": list(self.errors),
            "matched_courses": list(self.matched_courses),
            "missing_courses": list(self.missing_courses),
            "missing_groups": list(self.missing_groups),
            "missing_venues": list(self.missing_venues),
            "missing_references": list(self.missing_references),
            "conflicts": list(self.conflicts),
            "duplicates": list(self.duplicates),
            "ambiguous": list(self.ambiguous),
            "unknown_keys": list(self.unknown_keys),
            "invalid_positions": list(self.invalid_positions),
            "unrecognized_cells": list(self.unrecognized_cells),
            "all_rows": list(self.all_rows),
            "programmes_created": list(self.programmes_created),
            "venue_resolutions": list(self.venue_resolutions),
            "venue_conflicts": list(self.venue_conflicts),
            "venue_capacity_issues": list(self.venue_capacity_issues),
            "invalid_requirements": list(self.invalid_requirements),
            "courses_created": list(self.courses_created),
            "courses_updated": list(self.courses_updated),
            "course_requirements": list(self.course_requirements),
            "requirements_column": self.requirements_column,
            "requirement_count_columns": dict(self.requirement_count_columns),
            "derived_requirements": list(self.derived_requirements),
            "derived_unresolved": list(self.derived_unresolved),
            "derived_skipped": list(self.derived_skipped),
            "warnings": list(self.warnings),
            "notes": list(self.notes),
            "lecture_sessions_processed": self.lecture_sessions_processed,
            "lecture_groups_linked": self.lecture_groups_linked,
            "lecture_groups_existing": self.lecture_groups_existing,
            "lecture_skipped_courses": list(self.lecture_skipped_courses),
            "format": self.format,
            "detected_semester": self.detected_semester,
        }


# ---------------------------------------------------------------------------
# Reference data imports
# ---------------------------------------------------------------------------

def import_programmes_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    code_col = _match_col(df, "code", "programme_code", "program code", "program_code", "programme", "program")
    name_col = _match_col(df, "name", "programme_name", "program_name", "title")
    missing = [c for c, col in (("code", code_col), ("name", name_col)) if not col]
    if missing:
        result.errors.append(
            f"Missing columns: {', '.join(missing)} "
            f"(found: {', '.join(map(str, df.columns))})"
        )
        return result

    for _, row in df.iterrows():
        code = str(row[code_col]).strip()
        name = str(row[name_col]).strip()
        if not code:
            result.errors.append(f"Row: empty programme code, skipped")
            result.skipped += 1
            continue
        _, created = Programme.objects.update_or_create(
            code=code, defaults={"name": name}
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


def import_student_groups_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    prog_col = _match_col(
        df,
        "programme_code",
        "programme",
        "program",
        "program_code",
        "program code",
        "programme code",
        "programme_name",
        "program_name",
        "programme name",
        "program name",
    )
    grp_col = _match_col(df, "group_code", "group", "groups", "group_name")
    missing = [
        c
        for c, col in (
            ("programme_code", prog_col),
            ("group_code", grp_col),
        )
        if not col
    ]
    if missing:
        result.errors.append(f"Missing columns: {', '.join(missing)}")
        return result

    for _, row in df.iterrows():
        prog_code = str(row[prog_col]).strip()
        grp_code = str(row[grp_col]).strip()
        programme = Programme.objects.filter(code=prog_code).first()
        if programme is None:
            programme = Programme.objects.filter(name__iexact=prog_code).first()
        if programme is None:
            result.errors.append(
                f"Programme '{prog_code}' not found for group '{grp_code}'"
            )
            result.skipped += 1
            continue
        _, created = StudentGroup.objects.update_or_create(
            programme=programme, code=grp_code
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


def import_programme_courses_from_excel(
    path: str | Path, requirements_column: str = None
) -> ImportResult:
    """Import programme-course rows, populating the shared Course records.

    Each row finds or creates the one shared ``Course`` for its course code and
    sets that course's name and required activities, then creates or updates
    the programme's link to it.

    **Requirements are read from an optional column.** The header is resolved by
    :func:`core.group_allocation.find_requirements_column`, which understands
    "Required Activities", "Allocation Requirements", "Activities Required" and
    friends (case, spacing and punctuation are ignored), falls back to a fuzzy
    "mentions an activity and a count word" match, and can be pointed at a
    specific header with ``requirements_column`` / ``--requirements-column``.

    Per-activity count columns ("Tutorial Count", "Number of Practicals") are
    read too, so a workbook can keep the numbers in their own cells instead of
    one list.

    A column that looks like it holds requirements but is not understood is
    **reported, never ignored** — that is the failure mode where a requirements
    column quietly does nothing. A workbook with no requirements column at all
    still imports; those courses are flagged during allocation review.

    A course code that appears more than once in the same workbook with
    conflicting names or requirements is reported as a conflict rather than
    silently resolved, and re-importing the same file changes nothing.
    """
    from core.group_allocation import (
        find_requirement_count_columns,
        find_requirements_column,
        parse_requirements,
    )
    from core.models import Course, normalise_course_code

    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    prog_col = _match_col(
        df,
        "programme_code",
        "programme",
        "program",
        "program_code",
        "program code",
        "programme code",
        "programme_name",
        "program_name",
        "programme name",
        "program name",
    )
    code_col = _match_col(
        df, "course_code", "course code", "code", "subject_code", "unit_code"
    )
    name_col = _match_col(
        df,
        "course_name",
        "course",
        "course_title",
        "course title",
        "subject",
        "subject_name",
        "unit_name",
    )
    sem_col = _match_col(df, "semester", "sem", "term")

    missing = [
        c
        for c, col in (
            ("programme_code", prog_col),
            ("course_code", code_col),
            ("course_name", name_col),
            ("semester", sem_col),
        )
        if not col
    ]
    if missing:
        result.errors.append(
            f"Missing required columns: {', '.join(missing)} "
            f"(found: {', '.join(map(str, df.columns))})"
        )
        return result

    activities_col, unrecognised = find_requirements_column(
        df.columns, explicit=requirements_column
    )
    if requirements_column and not activities_col:
        result.errors.append(
            f"Required-activities column '{requirements_column}' not found "
            f"(available columns: {', '.join(map(str, df.columns))})"
        )
        return result
    count_columns = find_requirement_count_columns(
        df.columns, exclude=[activities_col] if activities_col else []
    )
    result.requirements_column = str(activities_col) if activities_col else ""
    result.requirement_count_columns = {
        activity: str(column) for activity, column in count_columns.items()
    }
    if unrecognised:
        # A column that is clearly about requirements but was not understood.
        # Say so, rather than importing nothing and leaving the courses to be
        # discovered later as "requirement not configured".
        result.errors.append(
            "A column looks like it holds the required activities but was not "
            f"recognised: {', '.join(unrecognised)}. Rename it to one of the "
            "accepted headings (e.g. 'Required Activities'), or re-run with "
            "--requirements-column '<the exact header>' to point at it."
        )
    if not activities_col and not count_columns:
        result.warnings.append(
            "No required-activities column was found, so every imported course "
            "will have no seminar/tutorial/practical requirement and will be "
            "flagged during allocation review. Add a 'Required Activities' "
            "column to set them."
        )

    # Pass 1: resolve every row, and record what each course code is being asked
    # to say. Two rows disagreeing about one code is a conflict to report, not a
    # value to pick from.
    rows = []
    seen: dict = {}
    for idx, row in df.iterrows():
        row_no = idx + 2
        prog_raw = str(row[prog_col]).strip()
        crs_code = normalise_course_code(row[code_col])
        crs_name = str(row[name_col]).strip()
        try:
            sem = int(float(str(row[sem_col]).strip()))
        except (ValueError, TypeError):
            result.errors.append(
                f"Row {row_no}: invalid semester value for course "
                f"'{crs_code}' in programme '{prog_raw}'"
            )
            result.skipped += 1
            continue
        if not crs_code:
            result.errors.append(
                f"Row {row_no}: missing course_code, skipped"
            )
            result.skipped += 1
            continue
        requirements, problems = set_requirements_for_row(
            row, activities_col, count_columns
        )
        if problems:
            result.invalid_requirements.append(
                f"Row {row_no}: course {crs_code} — "
                + "; ".join(problems)
                + ". The course was left with no required activities "
                "configured; set them on the course record."
            )
            requirements = {}
        result.course_requirements.append(
            {
                "code": crs_code,
                "name": crs_name,
                "row_no": row_no,
                "requirements": _requirements_text(requirements),
            }
        )
        entry = {
            "row_no": row_no,
            "programme_raw": prog_raw,
            "code": crs_code,
            "name": crs_name,
            "semester": sem,
            "requirements": requirements,
            "first_for_code": crs_code not in seen,
            "name_conflict": False,
            "requirements_conflict": False,
        }
        rows.append(entry)
        previous = seen.get(crs_code)
        if previous is None:
            seen[crs_code] = entry
            continue
        # A blank cell means "this row has no opinion", not "no requirement".
        # Requirements and the course name describe the course, not the
        # programme, so a workbook that fills them in on only the first of a
        # course's rows is the normal shape — only two *different stated*
        # values are a conflict.
        if crs_name and previous["name"] and previous["name"] != crs_name:
            previous["name_conflict"] = entry["name_conflict"] = True
            result.conflicts.append(
                f"Course {crs_code} appears with conflicting names "
                f"('{previous['name']}' on row {previous['row_no']} and "
                f"'{crs_name}' on row {row_no}) — the shared course name was "
                f"left unchanged. Fix the source file and re-import."
            )
        if requirements and previous["requirements"] and (
            previous["requirements"] != requirements
        ):
            previous["requirements_conflict"] = entry["requirements_conflict"] = True
            result.conflicts.append(
                f"Course {crs_code} appears with conflicting required "
                f"activities ('{_requirements_text(previous['requirements'])}' "
                f"on row {previous['row_no']} and "
                f"'{_requirements_text(requirements)}' on row {row_no}) — the "
                f"shared course's requirements were left unchanged. Fix the "
                f"source file and re-import."
            )
        previous["semester"] = max(previous["semester"], sem)

    # Pass 2: write. The shared course is written first so the programme link
    # can mirror its code and name.
    for entry in rows:
        programme = Programme.objects.filter(code=entry["programme_raw"]).first()
        if programme is None:
            programme = Programme.objects.filter(
                name__iexact=entry["programme_raw"]
            ).first()
        if programme is None:
            programme = Programme.objects.create(
                code=_derive_code(entry["programme_raw"]),
                name=entry["programme_raw"],
            )
            result.programmes_created.append(
                f"{programme.name} -> {programme.code}"
            )
        course = _sync_course(entry, result)
        if course is None:
            continue
        _, created = ProgrammeCourse.objects.update_or_create(
            programme=programme,
            course_code=course.code,
            defaults={"course": course, "semester": entry["semester"]},
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


def _requirements_text(mapping: dict) -> str:
    from core.group_allocation import format_requirements

    return format_requirements(mapping) or "(none)"


def set_requirements_for_row(row, activities_col, count_columns: dict):
    """Requirements for one workbook row: the list cell plus any count cells.

    Returns ``({activity: count}, problems)``. A dedicated count column
    ("Tutorial Count" = 2) adds to whatever the combined cell said, so the two
    layouts can be mixed in one file. An unusable count cell is a problem
    against that row rather than a silent zero.

    Lives here so ``manage.py set_course_requirements`` and the programme-course
    import read a requirements sheet by exactly the same rules.
    """
    from core.group_allocation import parse_requirements

    counts: dict = {}
    problems: list = []
    if activities_col:
        raw = str(row.get(activities_col, "") or "").strip()
        parsed, issues = parse_requirements(raw)
        counts.update(parsed)
        problems.extend(issues)
    for activity, column in (count_columns or {}).items():
        raw = str(row.get(column, "") or "").strip()
        if not raw or raw.lower() in {"-", "n/a", "na"}:
            continue
        try:
            value = int(float(raw))
        except (TypeError, ValueError):
            problems.append(
                f"'{raw}' in column '{column}' is not a whole number of "
                f"sessions"
            )
            continue
        if value < 0:
            problems.append(
                f"column '{column}' has {value}; a count cannot be negative"
            )
            continue
        if value:
            counts[activity] = counts.get(activity, 0) + value
    return counts, problems


def _sync_course(entry, result: ImportResult):
    """Create or update the shared Course for one imported row.

    Two rules keep re-importing idempotent:

    * Only the first row of a course code writes the name or the requirements.
      A workbook normally repeats a course once per programme; letting every
      row write would mean the last row silently wins.
    * When the rows for one code disagree, the stored name and requirements are
      left exactly as they are. The disagreement is already reported as a
      conflict; choosing a side here would be the silent resolution the
      requirements forbid, and it would make every re-import flip the name.
      The losing spellings are still recorded as name variants so the
      coordinator can see what the file actually said.
    """
    code = entry["code"]
    name = entry["name"]
    course = Course.objects.filter(code=code).first()
    if course is None:
        course = Course(code=code, name=name)
        course.save()
        result.courses_created.append(code)
        if entry["requirements"]:
            course.set_requirements(entry["requirements"])
        return course

    variants = course.name_conflicts()
    if entry["name_conflict"]:
        # Keep the stored name; just record what else this file called it.
        if name and name != course.name and name not in variants:
            variants.append(name)
    elif entry["first_for_code"] and name and course.name != name:
        variants = [
            v for v in variants + [course.name] if v and v != name
        ]
        course.name = name
    variants = sorted(set(variants))
    if variants != course.name_conflicts():
        course.name_variants = json.dumps(variants)
        course.save(update_fields=["name_variants"])
    else:
        course.save()

    result.courses_updated.append(code)
    if entry["first_for_code"] and not entry["requirements_conflict"]:
        # Only a stated value is written: a workbook whose requirements column
        # is entirely blank must not wipe requirements someone set by hand.
        if entry["requirements"]:
            course.set_requirements(entry["requirements"])
    return course


def import_venues_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    name_col = _match_col(
        df,
        "name",
        "venue",
        "venue_name",
        "room",
        "room_name",
        "location",
        "location_name",
    )
    cap_col = _match_col(df, "capacity", "seats", "seat_capacity", "size")
    missing = [
        c for c, col in (("name", name_col), ("capacity", cap_col)) if not col
    ]
    if missing:
        result.errors.append(
            f"Missing columns: {', '.join(missing)} "
            f"(found: {', '.join(map(str, df.columns))})"
        )
        return result

    blanks = []
    for _, row in df.iterrows():
        name = str(row[name_col]).strip()
        if not name:
            result.errors.append(f"Row: empty venue name, skipped")
            result.skipped += 1
            continue
        raw_cap = str(row[cap_col]).strip()
        if not raw_cap:
            # A name with no capacity is still imported. Spacing and casing
            # duplicates of it are the same room and are merged below; the
            # missing capacity is flagged only when the room we keep still
            # has none.
            cap = 0
        else:
            try:
                cap = int(float(raw_cap))
            except (ValueError, TypeError):
                result.errors.append(f"Invalid capacity for venue '{name}'")
                result.skipped += 1
                continue
        existing = venue_by_formatting(name)
        if existing is not None:
            spellings = {existing.name, name}
            merging = len(spellings) > 1
            official = (
                official_formatting_name(spellings) if merging else existing.name
            )
            venue = adopt_official_name(existing, official)
            chosen = kept_capacity(venue.capacity, cap) if merging or cap <= 0 else cap
            if venue.capacity != chosen:
                venue.capacity = chosen
                venue.save(update_fields=["capacity"])
            result.updated += 1
            for spelling in sorted(spellings):
                if spelling != venue.name:
                    _remember_note(
                        result, f"{spelling} is the same venue as {venue.name}."
                    )
        else:
            Venue.objects.create(name=name, capacity=cap)
            result.created += 1
        if not raw_cap:
            blanks.append(name)

    for note in fold_formatting_duplicates():
        _remember_note(result, note)

    # A blank capacity is still reported, but only when the room we kept has
    # no capacity. A second spelling that already stated one fills it in.
    for raw_name in blanks:
        venue = venue_by_formatting(raw_name)
        label = venue.name if venue is not None else raw_name
        if venue is not None and (venue.capacity or 0) > 0:
            continue
        if label not in result.venue_capacity_issues:
            result.venue_capacity_issues.append(label)

    # Anything left that is only a spacing or casing pair was merged above.
    # detect_name_conflicts stays for a genuine leftover the merge could not
    # fold; a pure formatting difference does not reach the chooser.
    result.venue_conflicts = detect_name_conflicts(
        Venue.objects.all().values_list("name", flat=True)
    )

    return result


def _resolve_venue(venue_raw: str, dry_run: bool = False) -> tuple[Venue, str, int]:
    """Resolve one venue name for a master-timetable session.

    Returns (venue, action, capacity) where action is:
    - "exact": the stored name matches
    - "matched": the same room, ignoring spaces and capitals, already exists
    - "created": no match, so a new venue is created with capacity 0
    - "skipped": the cell is empty or still a comma-separated list

    An existing venue keeps its name. A list such as ``P 008, PB 06, S112``
    is never stored as a venue; the caller picks one room before this runs.
    """
    venue_raw = (venue_raw or "").strip()
    if not venue_raw or "," in venue_raw:
        return None, "skipped", 0

    venue = Venue.objects.filter(name=venue_raw).first()
    if venue:
        return venue, "exact", venue.capacity or 0

    match = venue_by_formatting(venue_raw)
    if match is not None:
        return match, "matched", match.capacity or 0

    raw_key = base_key(venue_raw)
    if raw_key and "," not in raw_key:
        for existing in Venue.objects.all():
            if "," in existing.name:
                continue
            if base_key(existing.name) == raw_key:
                return existing, "matched", existing.capacity or 0

    if not dry_run:
        venue, _ = Venue.objects.get_or_create(
            name=venue_raw, defaults={"capacity": 0}
        )
        return venue, "created", 0
    return None, "created", 0


def import_semesters_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    year_col = _match_col(df, "academic_year", "academic year", "year", "session")
    sem_col = _match_col(df, "semester", "sem", "term")
    missing = [
        c
        for c, col in (
            ("academic_year", year_col),
            ("semester", sem_col),
        )
        if not col
    ]
    if missing:
        result.errors.append(f"Missing columns: {', '.join(missing)}")
        return result

    for _, row in df.iterrows():
        year = str(row[year_col]).strip()
        try:
            sem = int(float(str(row[sem_col]).strip()))
        except (ValueError, TypeError):
            result.errors.append(f"Invalid semester value for year '{year}'")
            result.skipped += 1
            continue
        _, created = Semester.objects.update_or_create(
            academic_year=year, semester=sem
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


# ---------------------------------------------------------------------------
# Timetable imports
# ---------------------------------------------------------------------------

MASTER_REQUIRED = {"course_code", "activity_type", "day", "start_time", "end_time"}

MASTER_COLUMN_ALIASES = {
    "course_code": [
        "course_code", "course code", "course", "code", "subject_code", "unit_code",
    ],
    "activity_type": [
        "activity_type", "activity type", "activity", "type",
    ],
    "day": ["day", "day_of_week", "dayofweek"],
    "start_time": ["start_time", "start", "time_from"],
    "end_time": ["end_time", "end", "time_to"],
    "venue": ["venue", "room", "room_name", "venue_name"],
    "group": ["group", "groups", "group_code", "group_name"],
}


def _missing_columns(df, required, aliases) -> list[str]:
    """Canonical names that could not be matched to a real column."""
    return [
        name
        for name in required
        if _match_col(df, *aliases.get(name, [name])) is None
    ]


def _remember_note(result, note):
    if note and note not in result.notes:
        result.notes.append(note)


def _choose_session_venue(venue_raw, course_code, activity_raw):
    """Pick the one room a master-timetable cell is actually using.

    A comma-separated cell is a list of rooms, never a venue name, so nothing
    called ``P 008, PB 06, S112`` is created. An ME101 tutorial whose list
    includes S112 (any spacing or capitals) is that room. Any other list uses
    a listed room only when exactly one of them already exists.
    """
    text = str(venue_raw or "").strip()
    if not text or text.lower() in {"nan", "none"}:
        return "", ""
    if "," not in text:
        existing = venue_by_formatting(text)
        if existing is None or existing.name == text:
            return text, ""
        return existing.name, f"{text} is the same venue as {existing.name}."

    parts = [part.strip() for part in text.split(",") if part.strip()]
    me101_tutorial = (
        normalise_course_code(course_code) == "ME101"
        and normalise_activity_type(activity_raw) == ActivityType.TUTORIAL
    )
    if me101_tutorial and any(formatting_key(part) == "s112" for part in parts):
        existing = venue_by_formatting("S112")
        chosen = existing.name if existing is not None else "S112"
        return chosen, f"{text} for ME101 tutorial was taken as {chosen}."

    matches = []
    seen = set()
    for part in parts:
        existing = venue_by_formatting(part)
        if existing is not None and existing.pk not in seen:
            seen.add(existing.pk)
            matches.append(existing)
    if len(matches) == 1:
        return matches[0].name, f"{text} was taken as {matches[0].name}."
    if len(matches) > 1:
        names = ", ".join(venue.name for venue in matches)
        return "", (
            f"{text} lists several venues ({names}) and was not stored as one venue."
        )
    return "", f"{text} is a list of rooms and was not created as a venue."


def _programmes_for_course_codes():
    """``{canonical course code: {programme code, ...}}``.

    Stored codes are matched through :func:`normalise_course_code`, so a
    programme link saved as ``DS114(COET)`` is found from a timetable cell
    written ``DS114`` or ``DS 114 (COET)``.
    """
    found = {}
    for code, programme in ProgrammeCourse.objects.values_list(
        "course_code", "programme__code"
    ):
        key = normalise_course_code(code)
        if not key:
            continue
        found.setdefault(key, set()).add(programme)
    return found


def _programme_ids_for_course_codes():
    """``{canonical course code: {programme pk, ...}}`` for lecture linking."""
    found = {}
    for code, programme_id in ProgrammeCourse.objects.values_list(
        "course_code", "programme_id"
    ):
        key = normalise_course_code(code)
        if not key:
            continue
        found.setdefault(key, set()).add(programme_id)
    return found


def _note_course_spelling(result, raw, canonical):
    """Record that a spaced, padded or ``(COET)`` code was recognised."""
    shown = " ".join(str(raw or "").split()).upper()
    if not shown or shown == canonical:
        return
    _remember_note(result, f"{shown} is the same course as {canonical}.")


def _spans_four_hours(start, end) -> bool:
    """True when a session covers four hour-slots.

    A master-timetable period ends at :55, so ``09:00-12:55`` is four hours
    (09, 10, 11 and 12) — the same block as ``09:00-13:00``.
    """
    if start is None or end is None or end <= start:
        return False
    return (
        sum(
            1
            for hour in range(0, 23)
            if start < time(hour + 1, 0) and end > time(hour, 0)
        )
        == 4
    )


_FOUR_HOUR_WORKSHOP_NOTE = (
    "Sessions of four hours were stored as Workshop, at venue Workshop."
)
_ARCHITECTURE_STUDIO_NOTE = (
    "Four-hour architecture studios kept their own course and venue."
)


def _architecture_course_prefixes():
    """Prefixes of an architecture programme's own courses.

    Bachelor of Architecture is stored as programme ``AR``, and its own
    courses are written ``AR111``, ``AR121`` and so on. A course the
    architecture students also take (``CL111``, ``DS114``, ``TW107``) does
    not start with that code, so it is not one of these studios.

    When the programme list has not been imported yet, ``AR`` is still the
    faculty's code for those courses.
    """
    prefixes = set()
    for code, name in Programme.objects.values_list("code", "name"):
        if "architect" not in f"{code} {name}".casefold():
            continue
        prefix = normalise_course_code(code)
        if len(prefix) >= 2:
            prefixes.add(prefix)
    if not prefixes:
        prefixes.add("AR")
    return prefixes


def _is_architecture_course(course_code) -> bool:
    """True when ``course_code`` is an architecture programme's own course."""
    code = normalise_course_code(course_code)
    if not code:
        return False
    return any(code.startswith(prefix) for prefix in _architecture_course_prefixes())


def _mark_four_hour_workshop(rec) -> bool:
    """Store a four-hour master-timetable block as a workshop.

    The course is saved as ``WORKSHOP`` and the venue as ``Workshop``,
    whatever the sheet wrote in those cells. The clock times and the groups
    stay as imported. ``sheet_course_code`` keeps the original code so an
    ``ALL`` group list still expands through the course the row named.

    An architecture studio is the exception. ``AR111`` lasting four hours stays
    that studio, in its own room, so the row is left as imported.
    """
    if rec.get("four_hour_workshop"):
        return True
    if not _spans_four_hours(rec.get("start_time"), rec.get("end_time")):
        return False
    if _is_architecture_course(rec.get("course_code")):
        rec["architecture_studio"] = True
        return False
    rec["sheet_course_code"] = rec["course_code"]
    rec["sheet_activity_type"] = rec["activity_type"]
    rec["course_code"] = "WORKSHOP"
    rec["activity_type"] = ActivityType.WORKSHOP
    rec["venue_raw"] = "Workshop"
    rec["venue_note"] = ""
    rec["four_hour_workshop"] = True
    return True


def _note_four_hour_workshop(result, rec):
    """Record which four-hour rule the row fell under."""
    if _mark_four_hour_workshop(rec):
        _remember_note(result, _FOUR_HOUR_WORKSHOP_NOTE)
        return True
    if rec.get("architecture_studio"):
        _remember_note(result, _ARCHITECTURE_STUDIO_NOTE)
    return False


def _locate_master_session(basis, earlier=None):
    """The session this row should write, or None when it is new.

    Matches the stored row, the same slot under another spelling of the
    course, or — when a four-hour block was rewritten as a workshop — the
    session previously imported from the sheet's own course and activity.
    """
    canonical = basis["course_code"]
    existing = Session.objects.filter(**basis).first()
    if existing is not None:
        return existing
    slot = {key: value for key, value in basis.items() if key != "course_code"}
    for session in Session.objects.filter(**slot):
        if normalise_course_code(session.course_code) == canonical:
            return session
    if not earlier:
        return None
    sheet_code = normalise_course_code(earlier.get("course_code"))
    sheet_activity = earlier.get("activity_type") or ""
    if not sheet_code or not sheet_activity:
        return None
    for session in Session.objects.filter(
        semester=basis["semester"],
        day=basis["day"],
        start_time=basis["start_time"],
        end_time=basis["end_time"],
        activity_type=sheet_activity,
    ):
        if normalise_course_code(session.course_code) == sheet_code:
            return session
    return None


def _get_or_create_master_session(basis, earlier=None):
    """Find the session for this slot, including one stored under another spelling.

    ``DS 114 (COET)``, ``DS114(COET)`` and ``DS114`` are one session. A row
    that arrives under another spelling updates the stored code to the plain
    code instead of creating a second session. A four-hour block that used to
    be stored under the sheet's own course and activity is updated in place,
    so re-importing it does not leave the old row beside the workshop.
    """
    canonical = basis["course_code"]
    session = _locate_master_session(basis, earlier)
    if session is None:
        return Session.objects.create(**basis), True
    changed = []
    if session.course_code != canonical:
        session.course_code = canonical
        changed.append("course_code")
    if session.activity_type != basis["activity_type"]:
        session.activity_type = basis["activity_type"]
        changed.append("activity_type")
    venue = basis["venue"]
    venue_id = venue.pk if venue is not None else None
    if session.venue_id != venue_id:
        session.venue = venue
        changed.append("venue")
    if changed:
        session.save(update_fields=changed)
    return session, False


def _read_master_rows(df) -> list[dict]:
    venue_col = _match_col(df, *MASTER_COLUMN_ALIASES["venue"])
    code_col = _match_col(df, *MASTER_COLUMN_ALIASES["course_code"])
    act_col = _match_col(df, *MASTER_COLUMN_ALIASES["activity_type"])
    day_col = _match_col(df, *MASTER_COLUMN_ALIASES["day"])
    start_col = _match_col(df, *MASTER_COLUMN_ALIASES["start_time"])
    end_col = _match_col(df, *MASTER_COLUMN_ALIASES["end_time"])
    group_col = _match_col(df, *MASTER_COLUMN_ALIASES["group"])

    rows = []
    for idx, row in df.iterrows():
        venue_val = row.get(venue_col, None) if venue_col else None
        if venue_val is None or pd.isna(venue_val):
            venue_raw = ""
        else:
            venue_raw = str(venue_val).strip()
        raw_code = str(row.get(code_col, "")).strip() if code_col else ""
        # Comma-separated course codes are split into one session per code.
        # A comma-separated venue cell is a list of rooms, resolved to one.
        codes = _split_csv(raw_code) if "," in raw_code else ([raw_code] if raw_code else [])
        activity_raw = str(row.get(act_col, "") if act_col else "").strip()
        for code in codes:
            chosen_venue, venue_note = _choose_session_venue(
                venue_raw, normalise_course_code(code), activity_raw
            )
            rows.append(
                {
                    "row_no": idx + 2,
                    "course_code_raw": code,
                    "course_code": normalise_course_code(code),
                    "activity_raw": activity_raw,
                    "activity_type": None,
                    "day_raw": str(
                        row.get(day_col, "") if day_col else ""
                    ).strip(),
                    "day": None,
                    "start_raw": str(
                        row.get(start_col, "") if start_col else ""
                    ).strip(),
                    "end_raw": str(
                        row.get(end_col, "") if end_col else ""
                    ).strip(),
                    "start_time": None,
                    "end_time": None,
                    "venue_raw": chosen_venue,
                    "venue_note": venue_note,
                    "raw_groups": str(
                        row.get(group_col, "") if group_col else ""
                    ).strip(),
                    "group_codes": [],
                    "expanded_groups": [],
                    "conflict": False,
                }
            )
    return rows


def _validate_master_rows(df, result) -> list[dict]:
    rows = _read_master_rows(df)
    valid = []
    for rec in rows:
        if not rec["course_code"]:
            result.errors.append(f"Row {rec['row_no']}: missing course_code, skipped")
            result.skipped += 1
            continue
        _note_course_spelling(result, rec.get("course_code_raw"), rec["course_code"])
        rec["activity_type"] = normalise_activity_type(rec["activity_raw"])
        try:
            rec["day"] = normalise_day(rec["day_raw"])
        except Exception:
            result.errors.append(f"Row {rec['row_no']}: invalid day '{rec['day_raw']}'")
            result.skipped += 1
            continue
        try:
            rec["start_time"] = parse_time(rec["start_raw"])
            rec["end_time"] = parse_time(rec["end_raw"])
        except (ValueError, TypeError) as exc:
            result.errors.append(f"Row {rec['row_no']}: time parse error: {exc}")
            result.skipped += 1
            continue
        if rec["activity_type"] == ActivityType.WORKSHOP:
            issue = workshop_time_issue(
                rec["day"],
                rec["start_time"],
                rec["end_time"],
                programme_code=_master_workshop_programme_codes(rec),
            )
            if issue:
                result.errors.append(f"Row {rec['row_no']}: {issue}")
                result.skipped += 1
                continue
        _note_four_hour_workshop(result, rec)
        if rec["raw_groups"].upper() != "ALL":
            rec["group_codes"] = [
                g.strip()
                for g in rec["raw_groups"].split(",")
                if g.strip() and g.strip().upper() != "ALL"
            ]
        if rec.get("venue_note"):
            _remember_note(result, rec["venue_note"])
        valid.append(rec)
    return valid


def _reconcile_master(df, semester_id) -> ImportResult:
    """Read-only pass: validate rows and report what would be missing."""
    result = ImportResult()
    missing = _missing_columns(df, MASTER_REQUIRED, MASTER_COLUMN_ALIASES)
    if missing:
        result.errors.append(
            f"Missing required columns: {', '.join(missing)} "
            f"(found: {', '.join(map(str, df.columns))})"
        )
        return result

    if semester_id is None:
        result.errors.append(
            "No semester specified — choose the academic semester this "
            "timetable belongs to"
        )
        return result
    semester = Semester.objects.filter(pk=semester_id).first()
    if semester is None:
        result.errors.append(
            f"Semester {semester_id} does not exist — import blocked"
        )
        return result

    rows = _validate_master_rows(df, result)

    known = _programmes_for_course_codes()
    codes = sorted({r["course_code"] for r in rows if r["course_code"]})
    # A four-hour block is stored as WORKSHOP on purpose. That name is not a
    # course the workbook forgot to list.
    placeholder = any(r.get("four_hour_workshop") for r in rows)
    result.matched_courses = sorted(code for code in codes if code in known)
    result.missing_courses = sorted(
        code
        for code in codes
        if code not in known and not (placeholder and code == "WORKSHOP")
    )

    venue_req = {
        r["venue_raw"] for r in rows if r["venue_raw"].lower() not in ("nan", "none")
    }
    existing_venues = set(
        Venue.objects.filter(name__in=venue_req).values_list("name", flat=True)
    )
    result.missing_venues = sorted(venue_req - existing_venues)

    group_req = {g for r in rows for g in r["group_codes"]}
    existing_groups = set(
        StudentGroup.objects.filter(code__in=group_req).values_list(
            "code", flat=True
        )
    )
    result.missing_groups = sorted(group_req - existing_groups)

    for rec in rows:
        if rec["raw_groups"].upper() == "ALL":
            link_code = rec.get("sheet_course_code") or rec["course_code"]
            progs = sorted(known.get(link_code, ()))
            groups = StudentGroup.objects.filter(
                programme__code__in=progs
            ).select_related("programme")
            if not progs:
                rec["conflict"] = True
                result.conflicts.append(
                    f"Row {rec['row_no']}: course '{link_code}' uses "
                    f"'ALL' but has no ProgrammeCourse mapping — groups cannot "
                    f"be expanded"
                )
                continue
            rec["expanded_groups"] = list(groups)
            result.all_rows.append(
                f"Row {rec['row_no']}: ALL for {link_code} expanded to "
                f"{len(groups)} student group(s) across {', '.join(progs)}"
            )
        elif rec["group_codes"]:
            missing = [g for g in rec["group_codes"] if g not in existing_groups]
            if missing:
                result.errors.append(
                    f"Row {rec['row_no']}: group(s) {', '.join(missing)} not "
                    f"found in DB — not linked"
                )
    return result


def reconcile_master_timetable(
    path: str | Path, semester_id: int | None = None
) -> ImportResult:
    """Dry-run: build the reconciliation report without writing anything."""
    df = pd.read_excel(path, dtype=str).fillna("")
    return _reconcile_master(df, semester_id)


def assign_lecture_groups(semester=None):
    """Attach Student Groups to LECTURE sessions via ProgrammeCourse -> programme.

    Every group of every programme that studies the course is attached
    (regardless of subgroup), so a lecture holds the full cohort. Idempotent:
    re-running only adds missing links. When ``semester`` is given, only that
    semester's lecture sessions are processed.
    Returns a 4-tuple: (processed, created_links, already_linked, skipped).
    """
    sessions = Session.objects.filter(activity_type="LECTURE")
    if semester is not None:
        sessions = sessions.filter(semester=semester)
    processed = 0
    created = 0
    existing = 0
    skipped = set()
    programmes = _programme_ids_for_course_codes()
    for session in sessions:
        prog_ids = programmes.get(normalise_course_code(session.course_code), ())
        if not prog_ids:
            skipped.add(session.course_code)
            continue
        for group in StudentGroup.objects.filter(programme_id__in=prog_ids):
            _, was_created = SessionGroup.objects.get_or_create(
                session=session, group=group
            )
            if was_created:
                created += 1
            else:
                existing += 1
        processed += 1
    return processed, created, existing, sorted(skipped)


def import_master_timetable_from_excel(
    path: str | Path,
    semester_id: int | None = None,
    dry_run: bool = False,
    derive_requirements: bool = False,
) -> ImportResult:
    """Import the master timetable idempotently (get_or_create on a natural key).

    The target semester must be supplied explicitly (``semester_id``) — it is
    never guessed at. LECTURE sessions also get every group of every programme
    that studies the course attached automatically (same as the "Assign
    Lecture Groups" action), so full lecture cohorts are linked right after
    import. Missing reference data is reported, never guessed at. Rows with
    unresolvable 'ALL' expansions are imported with no group links and listed
    under conflicts. With dry_run=True nothing is written.

    With ``derive_requirements=True`` (and not dry-run) the sessions just
    imported are used to fill in each course's required activities: a course
    with tutorial sessions gets ``Tutorial``, one with tutorials *and*
    practicals gets both. Only an activity that has a session is ever set, and
    a course that already has a requirement — because it came from the
    programme-course workbook, or someone set it by hand — is never touched.
    Courses nothing could be said about are listed in
    ``result.derived_unresolved`` so they are not mistaken for "needs nothing".
    """
    df = pd.read_excel(path, dtype=str).fillna("")

    result = _reconcile_master(df, semester_id)

    missing = _missing_columns(df, MASTER_REQUIRED, MASTER_COLUMN_ALIASES)
    if missing:
        # Missing columns already reported by the reconcile pass.
        return result

    if semester_id is None:
        return result

    semester = Semester.objects.filter(pk=semester_id).first()
    if semester is None:
        return result

    result.detected_semester = str(semester)

    for rec in _read_master_rows(df):
        course_code = rec["course_code"]
        if not course_code:
            continue
        try:
            day = normalise_day(rec["day_raw"])
            start_time = parse_time(rec["start_raw"])
            end_time = parse_time(rec["end_raw"])
        except (ValueError, TypeError):
            continue
        activity_type = normalise_activity_type(rec["activity_raw"])
        rec["activity_type"] = activity_type
        rec["start_time"] = start_time
        rec["end_time"] = end_time
        rec["day"] = day
        if activity_type == ActivityType.WORKSHOP:
            issue = workshop_time_issue(
                day,
                start_time,
                end_time,
                programme_code=_master_workshop_programme_codes(rec),
            )
            if issue:
                # Reported and counted as skipped in the validate pass.
                continue
        if _note_four_hour_workshop(result, rec):
            course_code = rec["course_code"]
            activity_type = rec["activity_type"]

        venue_raw = rec["venue_raw"]
        if not venue_raw:
            venue = None
        else:
            venue, action, capacity = _resolve_venue(venue_raw, dry_run=dry_run)
            result.venue_resolutions.append({
                "raw": venue_raw,
                "matched": venue.name if venue else "None",
                "action": action,
                "capacity": capacity,
            })

        basis = {
            "semester": semester,
            "course_code": course_code,
            "activity_type": activity_type,
            "day": day,
            "start_time": start_time,
            "end_time": end_time,
            "venue": venue,
        }
        earlier = None
        if rec.get("four_hour_workshop"):
            earlier = {
                "course_code": rec.get("sheet_course_code"),
                "activity_type": rec.get("sheet_activity_type"),
            }
        if dry_run:
            exists = _locate_master_session(basis, earlier) is not None
            result.created += int(not exists)
            result.updated += int(exists)
            continue

        session, created = _get_or_create_master_session(basis, earlier)
        if created:
            result.created += 1
        else:
            result.updated += 1

        raw_groups = rec["raw_groups"]
        link_code = rec.get("sheet_course_code") or course_code

        if raw_groups.upper() == "ALL":
            progs = sorted(
                _programmes_for_course_codes().get(link_code, ())
            )
            if not progs:
                # Unresolvable: reported as a conflict in the reconcile pass.
                continue
            for group_obj in StudentGroup.objects.filter(
                programme__code__in=progs
            ):
                SessionGroup.objects.get_or_create(session=session, group=group_obj)
        else:
            for grp_code in [
                g.strip()
                for g in raw_groups.split(",")
                if g.strip() and g.strip().upper() != "ALL"
            ]:
                group_obj = StudentGroup.objects.filter(code=grp_code).first()
                if group_obj is None:
                    continue
                SessionGroup.objects.get_or_create(session=session, group=group_obj)

    if not dry_run:
        proc, linked, existing, skipped = assign_lecture_groups(semester)
        result.lecture_sessions_processed = proc
        result.lecture_groups_linked = linked
        result.lecture_groups_existing = existing
        result.lecture_skipped_courses = skipped
        if derive_requirements:
            _derive_requirements_after_import(result, semester)

    return result


def _derive_requirements_after_import(result: ImportResult, semester) -> None:
    """Fill in requirements from the timetable that was just imported.

    A course that already has a requirement is skipped: the programme-course
    workbook's "Required Activities" column, or a value someone set by hand,
    is a decision and this must not quietly replace it.
    """
    from core.group_allocation import derive_requirements_from_sessions

    mapping, unresolved = derive_requirements_from_sessions(semester)
    for course, requirements in mapping.items():
        if course.has_requirements():
            result.derived_skipped.append(
                f"{course.code} already has "
                f"'{course.activities_label()}' configured - left unchanged"
            )
            continue
        course.set_requirements(requirements)
        result.courses_updated.append(course.code)
        result.derived_requirements.append(
            {"code": course.code, "requirements": course.activities_label()}
        )
    for course in unresolved:
        result.derived_unresolved.append(
            f"{course.code}: no seminar/tutorial/practical session exists, so "
            f"nothing was derived - decide whether the course needs one, or "
            f"whether its sessions are still to be timetabled"
        )


def _detect_workshop_format(source) -> str:
    """Peek at the file to distinguish the raw matrix workbook from the flat format."""
    try:
        source.seek(0)
    except AttributeError:
        pass
    try:
        fmt = detect_format(source)
    except Exception:
        fmt = "flat"
    try:
        source.seek(0)
    except AttributeError:
        pass
    return fmt


def _workshop_natural_key(semester, rec):
    return {
        "semester": semester,
        "course_code": rec.course_code,
        "group_code": rec.group_code,
        "day": rec.day,
        "time_period": rec.time_period,
        "position": rec.position,
        "schedule_section": rec.schedule_section,
        "week_start": rec.week_start,
        "week_end": rec.week_end,
        "year_of_study": rec.year_of_study,
        "workshop": rec.workshop,
        "venue": rec.venue,
        "start_time": None,
        "end_time": None,
    }


def _resolve_workshop_semester(parsed, semester_id):
    if semester_id is not None:
        return Semester.objects.filter(pk=semester_id).first()
    if parsed.academic_year and parsed.semester:
        return Semester.objects.filter(
            academic_year=parsed.academic_year, semester=parsed.semester
        ).first()
    return Semester.objects.filter(pk=1).first()


def reconcile_workshop_workbook(
    path: str | Path, semester_id: int | None = None
) -> ImportResult:
    """Dry-run: parse the raw matrix workbook and report what would import.

    Nothing is written. Missing reference data (semester, student groups) is
    reported, never guessed at.
    """
    result = ImportResult()
    parsed = parse_workbook(path)
    result.format = "raw university workshop matrix"
    if parsed.academic_year and parsed.semester:
        result.detected_semester = (
            f"{parsed.academic_year} - Semester {parsed.semester}"
        )
    result.errors = list(parsed.errors)
    result.errors.extend(f"{item}" for item in parsed.invalid_groups)
    result.unrecognized_cells = list(parsed.unrecognized_cells)
    result.unknown_keys = list(parsed.unknown_keys)
    result.duplicates = list(parsed.duplicate_keys)
    result.ambiguous = []

    semester = _resolve_workshop_semester(parsed, semester_id)
    if semester is None:
        wanted = semester_id if semester_id is not None else None
        label = result.detected_semester or f"semester {wanted or 1}"
        result.missing_references.append(
            f"Semester {label} not found in DB - import blocked"
        )
    else:
        codes = sorted({rec.group_code for rec in parsed.records})
        existing = set(
            StudentGroup.objects.filter(code__in=codes).values_list(
                "code", flat=True
            )
        )
        result.missing_groups = sorted(set(codes) - existing)

    if not result.errors:
        if semester is None:
            result.skipped = len(parsed.records)
        else:
            for rec in parsed.records:
                key = _workshop_natural_key(semester, rec)
                exists = WorkshopAllocation.objects.filter(**key).exists()
                result.created += int(not exists)
                result.updated += int(exists)
    return result


def import_university_workshop_workbook(
    path: str | Path,
    semester_id: int | None = None,
    dry_run: bool = False,
) -> ImportResult:
    """Import the raw matrix workshop workbook idempotently.

    When no semester is supplied, the one detected from the workbook title is
    used (created automatically if it does not exist yet). Records are written
    under a full natural key so re-imports are no-ops. With dry_run=True nothing
    is written and a missing semester is reported, not created.
    """
    if semester_id is None and not dry_run:
        parsed = parse_workbook(path)
        semester = _resolve_workshop_semester(parsed, None)
        if semester is None:
            if parsed.academic_year and parsed.semester:
                semester = Semester.objects.get_or_create(
                    academic_year=parsed.academic_year, semester=parsed.semester
                )[0]
            else:
                semester = _default_reference_semester()
        semester_id = semester.pk

    result = reconcile_workshop_workbook(path, semester_id=semester_id)
    result.format = "raw university workshop matrix"

    if result.missing_references:
        result.errors.append(f"Import blocked — {result.missing_references[0]}")
        return result

    semester = _resolve_workshop_semester(
        parse_workbook(path), semester_id
    )
    if semester is None:
        return result

    parsed = parse_workbook(path)
    result.created = 0
    result.updated = 0
    for rec in parsed.records:
        key = _workshop_natural_key(semester, rec)
        if dry_run:
            exists = WorkshopAllocation.objects.filter(**key).exists()
            result.created += int(not exists)
            result.updated += int(exists)
            continue
        _, created = WorkshopAllocation.objects.update_or_create(
            **key, defaults={}
        )
        if created:
            result.created += 1
        else:
            result.updated += 1
    return result


def import_workshop_allocation_from_excel(
    path: str | Path,
    semester_id: int | None = None,
    dry_run: bool = False,
) -> ImportResult:
    if _detect_workshop_format(path) == "matrix":
        return import_university_workshop_workbook(
            path, semester_id=semester_id, dry_run=dry_run
        )

    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    course_col = _match_col(df, "course_code", "course code", "code")
    group_col = _match_col(df, "group_code", "group", "groups", "group_name")
    day_col = _match_col(df, "day", "day_of_week", "dayofweek")
    start_col = _match_col(df, "start_time", "start", "time_from")
    end_col = _match_col(df, "end_time", "end", "time_to")
    time_col = _match_col(
        df, "time", "time_range", "time range", "time_slot", "time slot", "period"
    )
    venue_col = _match_col(
        df, "workshop", "venue", "room", "room_name", "venue_name"
    )

    if (
        group_col is None
        or day_col is None
        or venue_col is None
        or (time_col is None and (start_col is None or end_col is None))
    ):
        required = [
            "group_code",
            "day",
            "a time column ('time' range, or 'start_time' + 'end_time')",
            "workshop",
        ]
        if course_col is not None:
            required.insert(0, "course_code")
        result.errors.append(
            "Invalid workshop format. Expected FORMAT A (course_code, group_code, "
            "day, start_time, end_time, workshop) or FORMAT B (group_code, day, "
            "start_time, end_time, workshop where course_code is derived from the "
            "workshop). "
            f"Required columns: {', '.join(required)}. "
            f"Found columns: {', '.join(map(str, df.columns))}"
        )
        return result

    if course_col is None:
        result.format = (
            "workshop format (group_code, day, start_time, end_time, workshop) — "
            "course_code derived from workshop"
        )
    else:
        result.format = (
            "full format (course_code, group_code, day, start_time, end_time, workshop)"
        )

    if semester_id is None:
        semester_id = 1

    semester = Semester.objects.filter(pk=semester_id).first()
    if semester is None and not dry_run:
        semester = _default_reference_semester()
        result.detected_semester = str(semester)
    if semester is None:
        result.errors.append(f"Semester {semester_id} does not exist")
        return result

    for idx, row in df.iterrows():
        venue_raw = str(row.get(venue_col, "")).strip() if venue_col else ""
        # FORMAT B has no course_code column: derive it from the workshop venue
        # (e.g. venue=Electrical => course_code=Electrical). A blank cell in
        # FORMAT A falls back to the venue the same way.
        course_code = (
            str(row.get(course_col, "")).strip() if course_col else ""
        ) or venue_raw
        group_code = str(row.get(group_col, "")).strip() if group_col else ""

        if not course_code or not group_code:
            result.errors.append(f"Row {idx + 2}: missing group_code or venue/course_code")
            result.skipped += 1
            continue

        day_raw = str(row.get(day_col, "")).strip() if day_col else ""
        try:
            day = normalise_day(day_raw)
        except Exception:
            result.errors.append(f"Row {idx + 2}: invalid day '{day_raw}'")
            result.skipped += 1
            continue

        if time_col:
            start_raw, end_raw = _split_range(row.get(time_col, ""))
        else:
            start_raw = str(row.get(start_col, "")).strip() if start_col else ""
            end_raw = str(row.get(end_col, "")).strip() if end_col else ""
        try:
            start_time = parse_time(start_raw)
            end_time = parse_time(end_raw)
        except (ValueError, TypeError) as exc:
            result.errors.append(f"Row {idx + 2}: time parse error: {exc}")
            result.skipped += 1
            continue

        issue = workshop_time_issue(
            day,
            start_time,
            end_time,
            programme_code=programme_codes_for_group_codes(group_code),
        )
        if issue:
            result.errors.append(f"Row {idx + 2}: {issue}")
            result.skipped += 1
            continue

        basis = dict(
            semester=semester,
            course_code=course_code,
            group_code=group_code,
            day=day,
            start_time=start_time,
            end_time=end_time,
            venue=venue_raw,
        )
        if dry_run:
            exists = WorkshopAllocation.objects.filter(**basis).exists()
            result.created += int(not exists)
            result.updated += int(exists)
            continue

        _, created = WorkshopAllocation.objects.update_or_create(**basis, defaults={})
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


def import_td_allocation_from_excel(
    path: str | Path, semester_id: int | None = None
) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    course_col = _match_col(df, "course_code", "course code", "course")
    group_col = _match_col(df, "group_code", "group", "groups", "group_name")
    day_col = _match_col(df, "day", "day_of_week", "dayofweek")
    start_col = _match_col(df, "start_time", "start", "time_from")
    end_col = _match_col(df, "end_time", "end", "time_to")
    time_col = _match_col(
        df, "time", "time_range", "time range", "time_slot", "time slot", "period"
    )
    venue_col = _match_col(df, "venue", "room", "room_name", "venue_name")

    missing = [
        name
        for name, col in (
            ("course_code", course_col),
            ("group_code", group_col),
            ("day", day_col),
            ("venue", venue_col),
        )
        if not col
    ]
    if time_col is None and (start_col is None or end_col is None):
        missing.append("a time column ('time' range, or 'start_time'+'end_time')")
    if missing:
        result.errors.append(
            f"Missing required columns: {', '.join(missing)} "
            f"(found: {', '.join(map(str, df.columns))})"
        )
        return result

    if semester_id is None:
        semester_id = 1

    semester = Semester.objects.filter(pk=semester_id).first()
    if semester is None:
        semester = _default_reference_semester()
        result.detected_semester = str(semester)
    if semester is None:
        result.errors.append(f"Semester {semester_id} does not exist")
        return result

    # Pivoted layouts (Day/Time/Group/Venue) only fill the merged header cell of
    # each group block — forward-fill day/time/venue so every group gets them.
    days = _fill_down(df[day_col])
    times = _fill_down(df[time_col]) if time_col else [""] * len(df)
    venues = _fill_down(df[venue_col]) if venue_col else [""] * len(df)

    for idx in range(len(df)):
        row = df.iloc[idx]
        course_code = str(row.get(course_col, "")).strip()
        group_code = str(row.get(group_col, "")).strip()

        if not group_code:
            continue
        if not course_code:
            result.errors.append(
                f"Row {idx + 2}: missing course_code (group '{group_code}')"
            )
            result.skipped += 1
            continue

        day_raw = days[idx]
        if not day_raw:
            result.skipped += 1
            continue
        try:
            day = normalise_day(day_raw)
        except Exception:
            result.errors.append(f"Row {idx + 2}: invalid day '{day_raw}'")
            result.skipped += 1
            continue

        if time_col:
            start_raw, end_raw = _split_range(row.get(time_col, ""))
        else:
            start_raw = str(row.get(start_col, "")).strip()
            end_raw = str(row.get(end_col, "")).strip()
        try:
            start_time = parse_time(start_raw)
            end_time = parse_time(end_raw)
        except (ValueError, TypeError) as exc:
            result.errors.append(f"Row {idx + 2}: time parse error: {exc}")
            result.skipped += 1
            continue

        venue_raw = venues[idx]

        _, created = TechnicalDrawingAllocation.objects.update_or_create(
            semester=semester,
            course_code=course_code,
            group_code=group_code,
            day=day,
            start_time=start_time,
            end_time=end_time,
            venue=venue_raw,
            defaults={},
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result