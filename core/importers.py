from dataclasses import dataclass, field
from datetime import time
from pathlib import Path

import pandas as pd

from core.models import (
    ActivityType,
    Day,
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    Venue,
    WorkshopAllocation,
)
from core.workshop_parser import (
    detect_format,
    parse_workbook,
)

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

    @property
    def total(self):
        return self.created + self.updated + self.skipped

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
        if self.missing_references:
            lines.append(
                f"Missing references: {', '.join(self.missing_references[:20])}"
            )
        if self.all_rows:
            lines.append(f"'ALL' group expansions: {len(self.all_rows)}")
        if self.conflicts:
            lines.append(f"Conflicts: {len(self.conflicts)}")
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


# ---------------------------------------------------------------------------
# Reference data imports
# ---------------------------------------------------------------------------

def import_programmes_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    required = {"code", "name"}
    cols = set(c.strip().lower() for c in df.columns)
    if not required.issubset(cols):
        missing = required - cols
        result.errors.append(f"Missing columns: {missing}")
        return result

    col_map = {c.strip().lower(): c for c in df.columns}

    for _, row in df.iterrows():
        code = str(row[col_map["code"]]).strip()
        name = str(row[col_map["name"]]).strip()
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

    required = {"programme_code", "group_code"}
    cols = set(c.strip().lower() for c in df.columns)
    if not required.issubset(cols):
        missing = required - cols
        result.errors.append(f"Missing columns: {missing}")
        return result

    col_map = {c.strip().lower(): c for c in df.columns}

    for _, row in df.iterrows():
        prog_code = str(row[col_map["programme_code"]]).strip()
        grp_code = str(row[col_map["group_code"]]).strip()
        try:
            programme = Programme.objects.get(code=prog_code)
        except Programme.DoesNotExist:
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


def import_programme_courses_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    required = {"programme_code", "course_code"}
    cols = set(c.strip().lower() for c in df.columns)
    if not required.issubset(cols):
        missing = required - cols
        result.errors.append(f"Missing columns: {missing}")
        return result

    col_map = {c.strip().lower(): c for c in df.columns}

    for _, row in df.iterrows():
        prog_code = str(row[col_map["programme_code"]]).strip()
        crs_code = str(row[col_map["course_code"]]).strip()
        try:
            programme = Programme.objects.get(code=prog_code)
        except Programme.DoesNotExist:
            result.errors.append(
                f"Programme '{prog_code}' not found for course '{crs_code}'"
            )
            result.skipped += 1
            continue
        _, created = ProgrammeCourse.objects.update_or_create(
            programme=programme, course_code=crs_code
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


def import_venues_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    required = {"name", "capacity"}
    cols = set(c.strip().lower() for c in df.columns)
    if not required.issubset(cols):
        missing = required - cols
        result.errors.append(f"Missing columns: {missing}")
        return result

    col_map = {c.strip().lower(): c for c in df.columns}

    for _, row in df.iterrows():
        name = str(row[col_map["name"]]).strip()
        try:
            cap = int(float(str(row[col_map["capacity"]]).strip()))
        except (ValueError, TypeError):
            result.errors.append(f"Invalid capacity for venue '{name}'")
            result.skipped += 1
            continue
        _, created = Venue.objects.update_or_create(
            name=name, defaults={"capacity": cap}
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


def import_semesters_from_excel(path: str | Path) -> ImportResult:
    df = pd.read_excel(path, dtype=str).fillna("")
    result = ImportResult()

    required = {"academic_year", "semester"}
    cols = set(c.strip().lower() for c in df.columns)
    if not required.issubset(cols):
        missing = required - cols
        result.errors.append(f"Missing columns: {missing}")
        return result

    col_map = {c.strip().lower(): c for c in df.columns}

    for _, row in df.iterrows():
        year = str(row[col_map["academic_year"]]).strip()
        try:
            sem = int(float(str(row[col_map["semester"]]).strip()))
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


def _read_master_rows(df) -> list[dict]:
    cols_lower = {c.strip().lower(): c for c in df.columns}

    def col(name):
        return cols_lower.get(name, None)

    rows = []
    group_col = col("group") or col("groups") or col("group_code")
    for idx, row in df.iterrows():
        venue_val = row.get(col("venue"), None)
        if pd.isna(venue_val):
            venue_raw = ""
        else:
            venue_raw = str(venue_val).strip()
        rows.append(
            {
                "row_no": idx + 2,
                "course_code": str(row.get(col("course_code"), "")).strip(),
                "activity_raw": str(row.get(col("activity_type"), "")).strip(),
                "activity_type": None,
                "day_raw": str(row.get(col("day"), "")).strip(),
                "day": None,
                "start_raw": str(row.get(col("start_time"), "")).strip(),
                "end_raw": str(row.get(col("end_time"), "")).strip(),
                "start_time": None,
                "end_time": None,
                "venue_raw": venue_raw,
                "raw_groups": str(row.get(group_col, "")).strip()
                if group_col
                else "",
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
        if rec["raw_groups"].upper() != "ALL":
            rec["group_codes"] = [
                g.strip()
                for g in rec["raw_groups"].split(",")
                if g.strip() and g.strip().upper() != "ALL"
            ]
        valid.append(rec)
    return valid


def _reconcile_master(df, semester_id) -> ImportResult:
    """Read-only pass: validate rows and report what would be missing."""
    result = ImportResult()
    cols_lower = {c.strip().lower(): c for c in df.columns}
    required = MASTER_REQUIRED - set(cols_lower)
    if required:
        result.errors.append(
            f"Missing required columns: {', '.join(sorted(required))}"
        )
        return result

    if semester_id is None:
        semester_id = 1
    semester = Semester.objects.filter(pk=semester_id).first()
    if semester is None:
        result.errors.append(
            f"Semester {semester_id} does not exist — import blocked"
        )
        return result

    rows = _validate_master_rows(df, result)

    codes = sorted({r["course_code"] for r in rows})
    known = set(
        ProgrammeCourse.objects.filter(course_code__in=codes).values_list(
            "course_code", flat=True
        )
    )
    result.matched_courses = sorted(known)
    result.missing_courses = sorted(set(codes) - known)

    prog_map = {}
    for code in codes:
        prog_map[code] = sorted(
            ProgrammeCourse.objects.filter(course_code=code).values_list(
                "programme__code", flat=True
            )
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
            progs = prog_map.get(rec["course_code"], [])
            groups = StudentGroup.objects.filter(
                programme__code__in=progs
            ).select_related("programme")
            if not progs:
                rec["conflict"] = True
                result.conflicts.append(
                    f"Row {rec['row_no']}: course '{rec['course_code']}' uses "
                    f"'ALL' but has no ProgrammeCourse mapping — groups cannot "
                    f"be expanded"
                )
                continue
            rec["expanded_groups"] = list(groups)
            result.all_rows.append(
                f"Row {rec['row_no']}: ALL for {rec['course_code']} expanded to "
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


def import_master_timetable_from_excel(
    path: str | Path, semester_id: int | None = None, dry_run: bool = False
) -> ImportResult:
    """Import the master timetable idempotently (get_or_create on a natural key).

    Reconciles first; missing reference data is reported, never guessed at.
    Rows with unresolvable 'ALL' expansions are imported with no group links and
    listed under conflicts. With dry_run=True nothing is written.
    """
    df = pd.read_excel(path, dtype=str).fillna("")
    result = _reconcile_master(df, semester_id)

    cols_lower = {c.strip().lower(): c for c in df.columns}
    if MASTER_REQUIRED - set(cols_lower):
        # Missing columns already reported by the reconcile pass.
        return result

    side = semester_id if semester_id is not None else 1
    semester = Semester.objects.filter(pk=side).first()
    if semester is None:
        return result

    for idx, row in df.iterrows():
        course_code = str(row.get(cols_lower.get("course_code"), "")).strip()
        if not course_code:
            continue
        try:
            day = normalise_day(str(row.get(cols_lower.get("day"), "")).strip())
            start_time = parse_time(str(row.get(cols_lower.get("start_time"), "")).strip())
            end_time = parse_time(str(row.get(cols_lower.get("end_time"), "")).strip())
        except (ValueError, TypeError):
            continue
        activity_type = normalise_activity_type(
            str(row.get(cols_lower.get("activity_type"), "")).strip()
        )

        venue_val = row.get(cols_lower.get("venue"), None)
        venue_raw = "" if pd.isna(venue_val) else str(venue_val).strip()
        if not venue_raw:
            venue = None
        else:
            venue = Venue.objects.filter(name=venue_raw).first()
            if not dry_run and venue is None:
                venue, _ = Venue.objects.get_or_create(
                    name=venue_raw, defaults={"capacity": 0}
                )

        basis = {
            "semester": semester,
            "course_code": course_code,
            "activity_type": activity_type,
            "day": day,
            "start_time": start_time,
            "end_time": end_time,
            "venue": venue,
        }
        if dry_run:
            exists = Session.objects.filter(**basis).exists()
            result.created += int(not exists)
            result.updated += int(exists)
            continue

        session, created = Session.objects.get_or_create(**basis)
        if created:
            result.created += 1
        else:
            result.updated += 1

        group_col = (
            cols_lower.get("group") or cols_lower.get("groups") or cols_lower.get("group_code")
        )
        raw_groups = str(row.get(group_col, "")).strip() if group_col else ""

        if raw_groups.upper() == "ALL":
            progs = list(
                ProgrammeCourse.objects.filter(course_code=course_code).values_list(
                    "programme__code", flat=True
                )
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

    return result


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

    Reconciles first; a missing Semester blocks the import. Records are written
    under a full natural key so re-imports are no-ops. With dry_run=True nothing
    is written.
    """
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
    result.format = "flat (course_code, group_code, day, start_time, end_time, venue)"

    cols_lower = {c.strip().lower(): c for c in df.columns}

    def col(name, default=None):
        return cols_lower.get(name.lower(), default)

    if semester_id is None:
        semester_id = 1

    semester = Semester.objects.filter(pk=semester_id).first()
    if semester is None:
        result.errors.append(f"Semester {semester_id} does not exist")
        return result

    for idx, row in df.iterrows():
        course_code = str(row.get(col("course_code", "course_code"), "")).strip()
        group_code = str(row.get(col("group_code", "group_code"), "")).strip()

        if not course_code or not group_code:
            result.errors.append(f"Row {idx + 2}: missing course_code or group_code")
            result.skipped += 1
            continue

        day_raw = str(row.get(col("day", "day"), "")).strip()
        try:
            day = normalise_day(day_raw)
        except Exception:
            result.errors.append(f"Row {idx + 2}: invalid day '{day_raw}'")
            result.skipped += 1
            continue

        start_raw = str(row.get(col("start_time", "start_time"), "")).strip()
        end_raw = str(row.get(col("end_time", "end_time"), "")).strip()
        try:
            start_time = parse_time(start_raw)
            end_time = parse_time(end_raw)
        except (ValueError, TypeError) as exc:
            result.errors.append(f"Row {idx + 2}: time parse error: {exc}")
            result.skipped += 1
            continue

        venue_raw = str(row.get(col("venue", "venue"), "")).strip()

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

    cols_lower = {c.strip().lower(): c for c in df.columns}

    def col(name, default=None):
        return cols_lower.get(name.lower(), default)

    if semester_id is None:
        semester_id = 1

    semester = Semester.objects.filter(pk=semester_id).first()
    if semester is None:
        result.errors.append(f"Semester {semester_id} does not exist")
        return result

    for idx, row in df.iterrows():
        course_code = str(row.get(col("course_code", "course_code"), "")).strip()
        group_code = str(row.get(col("group_code", "group_code"), "")).strip()

        if not course_code or not group_code:
            result.errors.append(f"Row {idx + 2}: missing course_code or group_code")
            result.skipped += 1
            continue

        day_raw = str(row.get(col("day", "day"), "")).strip()
        try:
            day = normalise_day(day_raw)
        except Exception:
            result.errors.append(f"Row {idx + 2}: invalid day '{day_raw}'")
            result.skipped += 1
            continue

        start_raw = str(row.get(col("start_time", "start_time"), "")).strip()
        end_raw = str(row.get(col("end_time", "end_time"), "")).strip()
        try:
            start_time = parse_time(start_raw)
            end_time = parse_time(end_raw)
        except (ValueError, TypeError) as exc:
            result.errors.append(f"Row {idx + 2}: time parse error: {exc}")
            result.skipped += 1
            continue

        venue_raw = str(row.get(col("venue", "venue"), "")).strip()

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