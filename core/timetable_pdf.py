"""Programme timetable PDF generation (reportlab).

Produces a classic weekly grid, A4 portrait: the DAYS run across the first row
and the hourly TIME slots (07:00-08:00 .. 19:00-20:00) run down the first
column, matching the on-screen timetable in ``core/timetable.html``. Sessions
that span multiple hourly slots merge their cells on that day. Each cell is
shaded by activity type — Lecture faded grey, Workshop faded green, Technical
Drawing faded pink. External activities (workshops / technical drawing) that
only carry a Morning/Afternoon period are placed in that day's standard
workshop slot — Monday/Tuesday/Wednesday/Friday morning 09:00-13:00, Thursday
morning 10:00-14:00, afternoons 15:00-19:00.

The PDF grid always spans the full configured daily window 07:00-20:00 so a
session at any hour within the day appears, even when the programme only has a
handful of sessions.

Workshop rotation
-----------------

When a student group has two or more *different* workshops allocated to the
same daily slot (same day and time, e.g. Electrical and Carpentry on Thursday
morning), those workshops rotate after every seven weeks. The cell then shows
``Electrical / Carpentry`` and a ``Workshop Rotation Key`` table is appended
below the grid listing, per programme/group/day: the workshop used in Week 1-7,
Week 8-14, and so on. Workshops that explicitly carry ``week_start``/``week_end``
ranges keep those ranges; otherwise the seven-week blocks are assigned in a
deterministic order (the workshop that matches the programme's name first, then
alphabetical). Non-rotating workshops render only their normal name and never
produce a rotation-key entry.
"""
import datetime
from collections import defaultdict
from xml.sax.saxutils import escape

from django.db.models import Q
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from core.models import (
    Day,
    Session,
    StudentGroup,
    TechnicalDrawingAllocation,
    TimePeriod,
    WorkshopAllocation,
)
from core.timetable_grid import build_time_day_grid, time_day_grid_to_table
from core.workshop_times import (
    allocation_programme_codes,
    workshop_hours,
    workshop_times_for,
)


def _ordinal(n):
    if n % 100 in (11, 12, 13):
        return f"{n}th"
    tail = n % 10
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(tail, "th")
    return f"{n}{suffix}"


def _display_times(hours, start_time=None, end_time=None):
    """(start, end) clock strings for an Activity Card header.

    Prefers the record's exact clock times; period-only records (matrix
    workshops) fall back to their covered hourly slots — the first slot is the
    start, one hour past the last slot is the end (e.g. Thursday morning
    hours {10, 11, 12, 13} => "10:00"-"14:00").
    """
    if start_time is not None and end_time is not None:
        return start_time.strftime("%H:%M"), end_time.strftime("%H:%M")
    if hours:
        return f"{min(hours):02d}:00", f"{max(hours) + 1:02d}:00"
    return "", ""


def _session_entry(session, groups_text=""):
    venue = session.venue.name if session.venue and session.venue.name else ""
    lines = [f"{session.course_code} {session.get_activity_type_display()}"]
    if groups_text:
        lines.append(groups_text)
    lines.append(venue or "-")
    start, end = _display_times(
        _covered_hours(session.start_time, session.end_time),
        session.start_time,
        session.end_time,
    )
    return {
        "key": ("session", session.pk),
        "day": session.day,
        "hours": _covered_hours(session.start_time, session.end_time),
        "label": "\n".join(lines),
        "kind": session.activity_type.lower(),
        "course_code": session.course_code,
        "name": session.course_code,
        "type_label": session.get_activity_type_display(),
        "venue": venue,
        "start": start,
        "end": end,
        "groups": groups_text,
        "note": "",
    }


def _workshop_display_name(rec):
    """The meaningful workshop/category name to show in a cell.

    Prefers the explicit workshop field, then the course code (both hold the
    workshop category for matrix imports), then the venue as a last resort.
    """
    return rec.workshop or rec.course_code or rec.venue


def _workshop_entry(rec):
    """Workshop cell: show only the meaningful workshop/category name.

    The group code, activity type, course code and venue are omitted from the
    cell; when two simultaneous entries share a name the grid builder appends
    the genuinely distinguishing details (week range, venue) rather than
    repeating every field for every entry.
    """
    note = ""
    if rec.week_start and rec.week_end:
        note = f"Wk {rec.week_start}-{rec.week_end}"
    name = _workshop_display_name(rec)
    hours = _workshop_hours(rec)
    start, end = _display_times(hours, rec.start_time, rec.end_time)
    return {
        "key": ("workshop", rec.pk),
        "day": rec.day,
        "hours": hours,
        "label": name,
        "kind": "workshop",
        "course_code": rec.course_code,
        "name": name,
        "type_label": "Workshop",
        "venue": rec.venue,
        "start": start,
        "end": end,
        "groups": rec.group_code,
        "note": note,
    }


def _td_entry(rec):
    start, end = _display_times(
        _covered_hours(rec.start_time, rec.end_time),
        rec.start_time,
        rec.end_time,
    )
    return {
        "key": ("td", rec.pk),
        "day": rec.day,
        "hours": _covered_hours(rec.start_time, rec.end_time),
        "label": "\n".join([
            f"{rec.course_code} TECHNICAL DRAWING",
            f"{rec.group_code}",
            rec.venue or "-",
        ]),
        "kind": "td",
        "course_code": rec.course_code,
        "name": rec.course_code,
        "type_label": "Technical Drawing",
        "venue": rec.venue,
        "start": start,
        "end": end,
        "groups": rec.group_code,
        "note": "",
    }


def _workshop_sort_time(rec):
    """Clock-time a workshop occupies, for a deterministic day ordering."""
    if rec.start_time is not None:
        return rec.start_time
    times = workshop_times_for(
        rec.day, rec.time_period, allocation_programme_codes(rec)
    )
    if times:
        return times[0]
    if rec.time_period == TimePeriod.MORNING:
        return datetime.time(9)
    if rec.time_period == TimePeriod.AFTERNOON:
        return datetime.time(15)
    return datetime.time(0)


def _workshop_slot_key(rec):
    """Identify the daily slot a workshop record belongs to.

    Records that share (group, day, slot) are the candidates for rotation: two
    *different* workshops in the same slot mean the group alternates between
    them every seven weeks. A different time on the same day is a separate
    session, not a rotation.
    """
    if rec.start_time is not None and rec.end_time is not None:
        return ("time", rec.start_time, rec.end_time)
    if rec.time_period:
        return ("period", rec.time_period)
    return ("time", None, None)


def _affinity_key(name, programme):
    """Deterministic ordering of workshop names inside a rotation.

    The workshop that matches the programme's identity (e.g. ``Electrical`` for
    the BSc. in Electrical Engineering programme) leads, so the natural
    workshop is shown first; everything else falls back to alphabetical order.
    Nothing is hard-coded — the rule is simply "a workshop whose name appears
    in the programme name comes first".
    """
    if programme and name and name.lower() in programme.name.lower():
        return (0, name.lower())
    return (1, name.lower())


def _rotation_blocks(records, programme):
    """Order a rotation's workshop names and assign their week ranges.

    When every record carries explicit ``week_start``/``week_end`` ranges those
    are kept (the matrix workbook shape). Otherwise the rotation is split into
    contiguous seven-week blocks — Week 1-7, Week 8-14, Week 15-21, ... — in
    the deterministic order produced by ``_affinity_key``.
    """
    names = sorted(
        {_workshop_display_name(r) for r in records},
        key=lambda n: _affinity_key(n, programme),
    )
    explicit = all(
        r.week_start is not None and r.week_end is not None for r in records
    )
    if explicit and len(records) == len({_workshop_display_name(r) for r in records}):
        ordered = sorted(
            records, key=lambda r: (r.week_start, r.week_end, _workshop_display_name(r))
        )
        return [
            (r.week_start, r.week_end, _workshop_display_name(r)) for r in ordered
        ]
    return [(1 + 7 * i, 7 + 7 * i, name) for i, name in enumerate(names)]


def _workshop_hours(rec):
    if rec.start_time is not None and rec.end_time is not None:
        return _covered_hours(rec.start_time, rec.end_time)
    return _period_hours(
        rec.time_period, rec.day, allocation_programme_codes(rec)
    )


def _build_rotation(group_code, day, records, programme):
    """Merge a group's same-slot workshops into one rotating cell + key row."""
    blocks = _rotation_blocks(records, programme)
    names = [name for _, _, name in blocks]
    label = " / ".join(names)
    note = " / ".join(f"Wk {wk_start}-{wk_end}" for wk_start, wk_end, _ in blocks)
    hours = set()
    for rec in records:
        hours.update(_workshop_hours(rec))
    courses = {rec.course_code for rec in records}
    course = next(iter(courses)) if len(courses) == 1 else ""
    keys = tuple(sorted(rec.pk for rec in records))
    start, end = _display_times(hours)
    entry = {
        "key": ("workshop-rotation", keys),
        "day": day,
        "hours": hours,
        "label": label,
        "kind": "workshop",
        "course_code": course or label,
        "name": label,
        "type_label": "Workshop",
        "venue": "",
        "start": start,
        "end": end,
        "groups": group_code,
        "note": note,
    }
    row = {
        "programme": programme.name if programme else "",
        "programme_code": programme.code if programme else "",
        "group": group_code,
        "day": day,
        "course": course,
        "blocks": blocks,
    }
    return entry, row


def _workshop_entries(qs, programme=None):
    """Turn a WorkshopAllocation queryset into display entries + rotation keys.

    Returns ``(entries, rotation_keys)``. A group's workshop for a given daily
    slot is either:

    - a single workshop: its week-run splits (e.g. Wk 1-6 and Wk 8-13 of the
      same workshop) are all preserved as separate entries, or
    - a rotation: two or more *different* workshops in the same slot merge into
      one ``Electrical / Carpentry`` entry, with a corresponding rotation-key
      row describing which workshop runs in Week 1-7, Week 8-14, ...

    Workshops in different slots of the same day are distinct sessions and are
    all kept. Record order within a slot is deterministic (day/time).
    """
    entries = []
    rotation_keys = []
    by_slot = defaultdict(list)
    for rec in qs:
        by_slot[(rec.group_code, rec.day, _workshop_slot_key(rec))].append(rec)

    for (group_code, day, _slot), records in by_slot.items():
        by_identity = defaultdict(list)
        for rec in records:
            by_identity[_workshop_display_name(rec)].append(rec)
        if len(by_identity) == 1:
            for recs in by_identity.values():
                sorted_recs = sorted(
                    recs, key=lambda r: (_workshop_sort_time(r), r.pk)
                )
                entries.extend(_workshop_entry(r) for r in sorted_recs)
            continue
        entry, row = _build_rotation(group_code, day, records, programme)
        entries.append(entry)
        rotation_keys.append(row)
    return entries, rotation_keys


def _covered_hours(start, end):
    """Hour slots a session covers given its start/end times.

    Considers the full clock day (00:00 .. 22:55 as the last whole slot) so a
    session that starts before 07:00 or ends after 19:55 is never clipped; the
    on-screen grid trims and the PDF's full-range grid unions these hours with
    the 07:00-20:00 display window.
    """
    hours = set()
    for h in range(0, 23):
        if start is not None and end is not None:
            if start < datetime.time(h + 1) and end > datetime.time(h):
                hours.add(h)
    return hours


def _period_hours(time_period, day=None, programme_code=None):
    """Hourly slots a morning/afternoon workshop covers.

    A morning workshop is one four-hour session (09:00-13:00 on most days,
    10:00-14:00 on Thursday) and an afternoon workshop one four-hour session
    15:00-19:00, matching the university rules.
    """
    hours = workshop_hours(day, time_period, programme_code)
    if hours:
        return hours
    if time_period == TimePeriod.MORNING:
        return set(range(9, 13))
    if time_period == TimePeriod.AFTERNOON:
        return set(range(15, 19))
    return set()


def _cell_markup(text):
    """Escape a cell's text for a reportlab Paragraph, keeping real line breaks.

    reportlab's Paragraph treats literal newlines as spaces, so the ``\n``
    separators used by entry labels are converted to explicit ``<br/>`` tags.
    XML-special characters (``&``, ``<``, ``>``) are escaped first so venue or
    course names cannot be misread as markup.
    """
    return escape(text).replace("\n", "<br/>")


def _apply_year_filter(qs, year):
    if year:
        return qs.filter(Q(year_of_study__isnull=True) | Q(year_of_study=year))
    return qs


def _collect_entries_and_rotations(programme, semester, group=None, year=None):
    """Gather timetable entries and workshop rotation keys for a selection.

    When ``group`` is given only that single student group's timetable is
    collected (sessions a group actually attends plus its workshops/TDs), which
    backs the per-group export. ``year`` optionally filters workshop
    allocations to a year of study (records without a year apply to every
    year). Returns ``(entries, rotation_keys)``.
    """
    if group is not None:
        return _collect_group_entries_and_rotations(group, semester, year=year)

    groups = list(StudentGroup.objects.filter(programme=programme))
    group_codes = {g.code for g in groups}

    entries = []

    sessions = (
        Session.objects.filter(
            semester=semester, session_groups__group__programme=programme
        )
        .select_related("venue")
        .distinct()
    )
    for session in sessions:
        attending = {
            sg.group.code
            for sg in session.session_groups.select_related("group")
            if sg.group.programme_id == programme.pk
        }
        groups_text = ""
        if 0 < len(attending) < len(groups):
            groups_text = ", ".join(sorted(attending))
        entries.append(_session_entry(session, groups_text))

    workshops = _apply_year_filter(
        WorkshopAllocation.objects.filter(
            semester=semester, group_code__in=group_codes
        ),
        year,
    )
    workshop_entries, rotation_keys = _workshop_entries(workshops, programme=programme)
    entries.extend(workshop_entries)

    tds = TechnicalDrawingAllocation.objects.filter(
        semester=semester, group_code__in=group_codes
    )
    for rec in tds:
        entries.append(_td_entry(rec))

    return entries, rotation_keys


def _collect_group_entries_and_rotations(group, semester, year=None):
    """As ``_collect_entries_and_rotations`` but restricted to ONE group."""
    entries = []

    sessions = (
        Session.objects.filter(semester=semester, session_groups__group=group)
        .select_related("venue")
        .distinct()
    )
    for session in sessions:
        entries.append(_session_entry(session, group.code))

    workshops = _apply_year_filter(
        WorkshopAllocation.objects.filter(semester=semester, group_code=group.code),
        year,
    )
    workshop_entries, rotation_keys = _workshop_entries(
        workshops, programme=group.programme
    )
    entries.extend(workshop_entries)

    tds = TechnicalDrawingAllocation.objects.filter(
        semester=semester, group_code=group.code
    )
    for rec in tds:
        entries.append(_td_entry(rec))

    return entries, rotation_keys


def collect_entries(programme, semester, group=None, year=None):
    """Gather timetable entries for one programme's groups in a semester.

    When ``group`` is given only that single student group's timetable is
    collected (sessions a group actually attends plus its workshops/TDs),
    which backs the per-group export. ``year`` optionally filters workshop
    allocations to a year of study (records without a year apply to every
    year).
    """
    return _collect_entries_and_rotations(programme, semester, group=group, year=year)[
        0
    ]


def collect_group_entries(group, semester, year=None):
    """Gather timetable entries for ONE student group in a semester.

    Sessions are those the group actually attends (via ``SessionGroup``);
    workshops and technical drawing slots are matched on the group's code.
    """
    return _collect_group_entries_and_rotations(group, semester, year=year)[0]


def collect_master_entries(semester, year=None):
    """Gather timetable entries for the whole COET First Year.

    The master timetable is the source of truth: every ``Session`` scheduled in
    the semester is included — not only lectures with group links — plus all
    workshop and technical-drawing allocations, so the default whole-year view
    shows the same sessions as the master timetable and renders in the same
    layout as a single programme's timetable. Session cards carry their
    actually assigned student group codes.
    """
    entries = []

    sessions = (
        Session.objects.filter(semester=semester)
        .select_related("venue")
        .prefetch_related("session_groups__group")
    )
    for session in sessions:
        groups_text = ", ".join(
            sorted(sg.group.code for sg in session.session_groups.all())
        )
        entries.append(_session_entry(session, groups_text))

    workshops = _apply_year_filter(
        WorkshopAllocation.objects.filter(semester=semester), year
    )
    workshop_entries, _rotation_keys = _workshop_entries(workshops)
    entries.extend(workshop_entries)

    tds = TechnicalDrawingAllocation.objects.filter(semester=semester)
    for rec in tds:
        entries.append(_td_entry(rec))

    return entries


def collect_workshop_rotations(programme, semester, group=None, year=None):
    """Rotation-key rows for a programme's (or one group's) workshops.

    Each row describes a rotating slot: the programme, group, day, shared
    course (where the records agree) and the ordered week blocks
    ``(week_start, week_end, workshop_name)``. Non-rotating workshops produce
    no rows.
    """
    return _collect_entries_and_rotations(programme, semester, group=group, year=year)[
        1
    ]


def build_grid(entries, show_groups=False):
    """Return table data (list of rows), SPAN commands and fill colours.

    Classic layout: column 0 is TIME, columns 1..n are the weekdays, row 0 is
    the DAY header. Uses the shared ``core.timetable_grid`` builder so the PDF
    matches the on-screen timetable. SPAN commands merge multi-slot session
    blocks vertically; fills carries one BACKGROUND command per block,
    colour-coded by activity type (Lecture grey, Workshop green, TD pink).
    ``show_groups`` appends owning group codes to workshop cells (the
    all-groups export).
    """
    grid = build_time_day_grid(entries, full_range=True)
    return time_day_grid_to_table(grid, show_groups=show_groups)


def render_programme_timetable(programme, semester, year_of_study=1, out=None):
    """Render the programme timetable PDF to `out` (file-like or a path)."""
    entries, rotation_keys = _collect_entries_and_rotations(
        programme, semester, year=year_of_study
    )
    show_groups = StudentGroup.objects.filter(programme=programme).count() > 1
    return _render_grid(
        entries,
        title=f"{programme.name.upper()}",
        subtitle=f"{_ordinal(year_of_study)} YEAR · {semester.academic_year}",
        semester=semester,
        doc_title=f"{programme.name} Timetable",
        show_groups=show_groups,
        rotation_keys=rotation_keys,
        out=out,
    )


def render_group_timetable(group, semester, year_of_study=1, out=None):
    """Render ONE student group's timetable PDF to `out` (file-like/path)."""
    entries, rotation_keys = _collect_group_entries_and_rotations(
        group, semester, year=year_of_study
    )
    return _render_grid(
        entries,
        title=f"{group.programme.name.upper()}",
        subtitle=(
            f"{_ordinal(year_of_study)} YEAR · GROUP {group.code} · "
            f"{semester.academic_year}"
        ),
        semester=semester,
        doc_title=f"{group.programme.code} {group.code} Timetable",
        show_groups=False,
        rotation_keys=rotation_keys,
        out=out,
    )


def _rotation_key_table(rotation_keys, cell_style, head_style):
    """Build the Workshop Rotation Key table flowables.

    Columns: Programme, Group, Day, Course/Session, then one week column per
    rotation block (Wk 1-7, Wk 8-14, ...).
    """
    usable = A4[0] - 22 * mm
    max_blocks = max(len(row["blocks"]) for row in rotation_keys)
    first = rotation_keys[0]["blocks"]
    week_headers = [f"Wk {ws}-{we}" for ws, we, _ in first]
    while len(week_headers) < max_blocks:
        week_headers.append(week_headers[-1] if week_headers else "Wk ?")

    headers = ["Programme", "Group", "Day", "Course/Session"] + week_headers
    data = [[Paragraph(head, head_style) for head in headers]]
    for row in rotation_keys:
        values = [
            row["programme"] or row["programme_code"],
            row["group"],
            Day(row["day"]).label,
            row["course"] or "—",
        ]
        values += [name for _, _, name in row["blocks"]]
        while len(values) < len(headers):
            values.append("—")
        data.append([Paragraph(str(v), cell_style) for v in values])

    n_week = max_blocks
    week_w = usable * 0.32 / max(n_week, 1)
    fixed = usable - week_w * n_week
    widths = [
        fixed * 0.34,
        fixed * 0.12,
        fixed * 0.16,
        fixed * 0.22,
    ] + [week_w] * n_week
    while len(widths) < len(headers):
        widths.append(week_w)

    table = Table(data, colWidths=widths, repeatRows=1, hAlign="CENTER")
    style = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dce6f1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    table.setStyle(TableStyle(style))
    return table


def _render_grid(
    entries,
    title,
    subtitle,
    semester,
    doc_title,
    out,
    show_groups=False,
    rotation_keys=None,
):
    """Render the weekly grid PDF to `out` (file-like or a path).

    Programmes/groups with no sessions get a clean "nothing scheduled" notice
    instead of a blank grid section; the Workshop Rotation Key is appended only
    when rotating workshops actually exist.
    """
    rotation_keys = rotation_keys or []

    doc = SimpleDocTemplate(
        out,
        pagesize=A4,
        leftMargin=11 * mm,
        rightMargin=11 * mm,
        topMargin=13 * mm,
        bottomMargin=14 * mm,
        title=doc_title,
    )

    title_style = ParagraphStyle(
        "tt",
        fontName="Helvetica-Bold",
        fontSize=14,
        alignment=TA_CENTER,
        leading=18,
        spaceAfter=2,
    )
    sub_style = ParagraphStyle(
        "sub", fontName="Helvetica", fontSize=9, alignment=TA_CENTER, leading=12
    )
    head_style = ParagraphStyle(
        "head",
        fontName="Helvetica-Bold",
        fontSize=8,
        alignment=TA_CENTER,
        leading=10,
    )
    cell_style = ParagraphStyle(
        "cell", fontName="Helvetica-Bold", fontSize=6.5, leading=8, alignment=TA_CENTER
    )
    time_style = ParagraphStyle(
        "time",
        fontName="Helvetica-Bold",
        fontSize=6.5,
        alignment=TA_CENTER,
        leading=8,
    )
    empty_style = ParagraphStyle(
        "empty",
        fontName="Helvetica",
        fontSize=10,
        alignment=TA_CENTER,
        leading=14,
        spaceBefore=6,
    )
    foot_style = ParagraphStyle(
        "foot", fontName="Helvetica", fontSize=8, alignment=TA_CENTER
    )
    key_title_style = ParagraphStyle(
        "keytitle",
        fontName="Helvetica-Bold",
        fontSize=9,
        alignment=TA_CENTER,
        leading=11,
        spaceAfter=3,
    )

    elements = [
        Paragraph(title, title_style),
        Paragraph(subtitle, sub_style),
        Paragraph(f"SEMESTER {semester.semester}", sub_style),
        Spacer(1, 4 * mm),
    ]

    if entries:
        data, spans, fills = build_grid(entries, show_groups=show_groups)

        n_days = max(len(data[0]) - 1, 1) if data else 1
        usable = A4[0] - 22 * mm
        time_w = 44
        day_w = (usable - time_w) / n_days
        widths = [time_w] + [day_w] * n_days

        table_data = [[Paragraph(item, head_style) for item in data[0]]]
        for row in data[1:]:
            table_data.append(
                [
                    Paragraph(row[0], time_style) if row[0] else "",
                ]
                + [
                    Paragraph(_cell_markup(cell), cell_style) if cell else ""
                    for cell in row[1:]
                ]
            )

        n_rows = len(data) - 1
        row_h = None
        if n_rows:
            title_block = 18 + 12 + 12 + 4 * mm
            footer_block = 6 * mm + 8
            usable_h = A4[1] - doc.topMargin - doc.bottomMargin
            avail = max(usable_h - title_block - footer_block, n_rows * 18)
            row_h = min(avail / n_rows, 42)

        grid = Table(
            table_data,
            colWidths=widths,
            repeatRows=1,
            rowHeights=[None] + [row_h] * n_rows if row_h else None,
        )
        style = [
            ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
            ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#eceff4")),
            ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#f8fafc")),
            ("BACKGROUND", (1, 0), (-1, 0), colors.HexColor("#dce6f1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ]
        for start, end in spans:
            style.append(("SPAN", start, end))
        for start, end, hex_color in fills:
            style.append(("BACKGROUND", start, end, colors.HexColor(hex_color)))
        grid.setStyle(TableStyle(style))

        elements.append(grid)
    else:
        elements.append(
            Paragraph(
                "No timetable sessions scheduled for this selection.",
                empty_style,
            )
        )

    if rotation_keys:
        elements.append(Spacer(1, 5 * mm))
        elements.append(
            Paragraph("WORKSHOP ROTATION KEY", key_title_style)
        )
        elements.append(_rotation_key_table(rotation_keys, cell_style, head_style))

    elements.append(Spacer(1, 6 * mm))
    export_date = datetime.date.today().strftime("%d %B %Y")
    elements.append(Paragraph(f"Prepared for personal use · {export_date}", foot_style))

    doc.build(elements)
    return doc