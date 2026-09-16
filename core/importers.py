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

    @property
    def total(self):
        return self.created + self.updated + self.skipped

    def summary(self) -> str:
        lines = [
            f"Created: {self.created}",
            f"Updated: {self.updated}",
            f"Skipped: {self.skipped}",
        ]
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

def _resolve_venue(raw: str) -> Venue | None:
    name = str(raw).strip()
    if not name or name.lower() in ("nan", "none", ""):
        return None
    venue, _ = Venue.objects.get_or_create(name=name, defaults={"capacity": 0})
    return venue


def _resolve_group_codes(raw: str) -> list[str]:
    text = str(raw).strip()
    if not text or text.lower() in ("nan", "none", ""):
        return []
    return [g.strip() for g in text.split(",") if g.strip() and g.strip().upper() != "ALL"]


def import_master_timetable_from_excel(
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
        if not course_code:
            result.errors.append(f"Row {idx + 2}: missing course_code, skipped")
            result.skipped += 1
            continue

        act_raw = str(row.get(col("activity_type", "activity_type"), "")).strip()
        activity_type = normalise_activity_type(act_raw)

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

        venue_raw = row.get(col("venue", "venue"), None)
        venue = _resolve_venue(venue_raw) if pd.notna(venue_raw) else None

        session, created = Session.objects.get_or_create(
            semester=semester,
            course_code=course_code,
            activity_type=activity_type,
            day=day,
            start_time=start_time,
            end_time=end_time,
            venue=venue,
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

        group_col = col("group", "groups") or col("group_code", "group_code")
        if group_col:
            raw_groups = str(row.get(group_col, "")).strip()
            for grp_code in _resolve_group_codes(raw_groups):
                group_obj = StudentGroup.objects.filter(
                    code=grp_code
                ).first()
                if group_obj is None:
                    result.errors.append(
                        f"Row {idx + 2}: group '{grp_code}' not found in DB"
                    )
                    continue
                SessionGroup.objects.get_or_create(
                    session=session, group=group_obj
                )

    return result


def import_workshop_allocation_from_excel(
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

        WorkshopAllocation.objects.create(
            semester=semester,
            course_code=course_code,
            group_code=group_code,
            day=day,
            start_time=start_time,
            end_time=end_time,
            venue=venue_raw,
        )
        result.created += 1

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

        TechnicalDrawingAllocation.objects.create(
            semester=semester,
            course_code=course_code,
            group_code=group_code,
            day=day,
            start_time=start_time,
            end_time=end_time,
            venue=venue_raw,
        )
        result.created += 1

    return result