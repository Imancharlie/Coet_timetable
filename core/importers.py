import re
from dataclasses import dataclass, field
from datetime import date, time
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
from core.venue_quality import base_key
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
        if self.programmes_created:
            lines.append(
                f"Programmes created automatically: "
                f"{', '.join(self.programmes_created[:20])}"
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


def import_programme_courses_from_excel(path: str | Path) -> ImportResult:
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

    for _, row in df.iterrows():
        prog_raw = str(row[prog_col]).strip()
        crs_code = str(row[code_col]).strip()
        crs_name = str(row[name_col]).strip()
        try:
            sem = int(float(str(row[sem_col]).strip()))
        except (ValueError, TypeError):
            result.errors.append(
                f"Invalid semester value for course '{crs_code}' "
                f"in programme '{prog_raw}'"
            )
            result.skipped += 1
            continue
        programme = Programme.objects.filter(code=prog_raw).first()
        if programme is None:
            programme = Programme.objects.filter(name__iexact=prog_raw).first()
        if programme is None:
            programme = Programme.objects.create(
                code=_derive_code(prog_raw), name=prog_raw
            )
            result.programmes_created.append(
                f"{programme.name} -> {programme.code}"
            )
        _, created = ProgrammeCourse.objects.update_or_create(
            programme=programme,
            course_code=crs_code,
            defaults={"course_name": crs_name, "semester": sem},
        )
        if created:
            result.created += 1
        else:
            result.updated += 1

    return result


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

    existing_keys = {
        base_key(n): n for n in Venue.objects.values_list("name", flat=True)
    }
    seen = {}
    for _, row in df.iterrows():
        name = str(row[name_col]).strip()
        if not name:
            result.errors.append(f"Row: empty venue name, skipped")
            result.skipped += 1
            continue
        try:
            cap = int(float(str(row[cap_col]).strip()))
        except (ValueError, TypeError):
            result.errors.append(f"Invalid capacity for venue '{name}'")
            result.skipped += 1
            continue
        key = base_key(name)
        clash = None
        if key in existing_keys and existing_keys[key] != name:
            clash = existing_keys[key]
        elif key in seen and seen[key] != name:
            clash = seen[key]
        if clash is not None:
            result.conflicts.append(
                f"Venue '{name}' matches existing venue '{clash}' "
                f"(normalised as '{key}') - same venue written with different "
                f"casing/spacing; it will be highlighted for recycling."
            )
        seen[key] = name
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
        # Comma-separated course codes are split into one session per code; the
        # venue cell is kept verbatim (a multi-room row stays a single session).
        codes = _split_csv(raw_code) if "," in raw_code else ([raw_code] if raw_code else [])
        for code in codes:
            rows.append(
                {
                    "row_no": idx + 2,
                    "course_code": code,
                    "activity_raw": str(
                        row.get(act_col, "") if act_col else ""
                    ).strip(),
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
                    "venue_raw": venue_raw,
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
    missing = _missing_columns(df, MASTER_REQUIRED, MASTER_COLUMN_ALIASES)
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

    When no semester is supplied and none exists yet, the current academic
    year's semester 1 is created automatically. Missing reference data is
    reported, never guessed at. Rows with unresolvable 'ALL' expansions are
    imported with no group links and listed under conflicts. With dry_run=True
    nothing is written.
    """
    df = pd.read_excel(path, dtype=str).fillna("")

    auto_sem = None
    if semester_id is None and not dry_run:
        semester = Semester.objects.filter(pk=1).first()
        if semester is None:
            semester = _default_reference_semester()
            auto_sem = semester
        semester_id = semester.pk

    result = _reconcile_master(df, semester_id)
    if auto_sem:
        result.detected_semester = str(auto_sem)

    missing = _missing_columns(df, MASTER_REQUIRED, MASTER_COLUMN_ALIASES)
    if missing:
        # Missing columns already reported by the reconcile pass.
        return result

    side = semester_id if semester_id is not None else 1
    semester = Semester.objects.filter(pk=side).first()
    if semester is None:
        return result

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

        venue_raw = rec["venue_raw"]
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

        raw_groups = rec["raw_groups"]

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
    venue_col = _match_col(df, "venue", "room", "room_name", "venue_name")

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
            "venue",
        ]
        if course_col is not None:
            required.insert(0, "course_code")
        result.errors.append(
            "Invalid workshop format. Expected FORMAT A (course_code, group_code, "
            "day, start_time, end_time, venue) or FORMAT B (group_code, day, "
            "start_time, end_time, venue where course_code is derived from venue). "
            f"Required columns: {', '.join(required)}. "
            f"Found columns: {', '.join(map(str, df.columns))}"
        )
        return result

    if course_col is None:
        result.format = (
            "workshop format (group_code, day, start_time, end_time, venue) — "
            "course_code derived from venue"
        )
    else:
        result.format = (
            "full format (course_code, group_code, day, start_time, end_time, venue)"
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