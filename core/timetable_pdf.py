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

from django.db.models import Q
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import (
    Flowable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.flowables import HRFlowable

from core.models import (
    Day,
    Session,
    StudentGroup,
    TechnicalDrawingAllocation,
    TimePeriod,
    WorkshopAllocation,
)
from core.timetable_grid import (
    DAY_ORDER,
    GROUPS_STYLE,
    GRID_HOUR_END,
    GRID_HOUR_START,
    WEEKEND_ORDER,
    build_time_day_grid,
    fill_color,
    time_day_grid_to_table,
)
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


def build_grid(entries, show_groups=False, markup=False):
    """Return table data (list of rows), SPAN commands and fill colours.

    Classic layout: column 0 is TIME, columns 1..n are the weekdays, row 0 is
    the DAY header. Uses the shared ``core.timetable_grid`` builder so the PDF
    matches the on-screen timetable. SPAN commands merge multi-slot session
    blocks vertically; fills carries one BACKGROUND command per block,
    colour-coded by activity type (Lecture grey, Workshop green, TD pink).
    ``show_groups`` appends owning group codes to workshop cells (the
    all-groups export). ``markup`` returns cell text as Paragraph markup with
    the assigned-groups line emphasised, already escaped.
    """
    grid = build_time_day_grid(entries, full_range=True)
    return time_day_grid_to_table(
        grid, show_groups=show_groups, markup=markup
    )


def render_programme_timetable(programme, semester, year_of_study=1, out=None):
    """Render the programme timetable PDF to `out` (file-like or a path)."""
    entries, rotation_keys = _collect_entries_and_rotations(
        programme, semester, year=year_of_study
    )
    groups = list(StudentGroup.objects.filter(programme=programme))
    show_groups = len(groups) > 1
    # Fold the records that describe one session before drawing: a workshop and
    # the placeholder practical written for it, and the same course recorded
    # once per group, become one cell listing every group.
    entries = fold_for_display(entries, {g.code for g in groups})
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
    # Same fold as the programme export, so a group's own workshop and the
    # placeholder practical for it are one cell rather than two.
    entries = fold_for_display(entries, {group.code})
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
        data, spans, fills = build_grid(
            entries, show_groups=show_groups, markup=True
        )

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
                    # Cell text arrives as ready-made markup: the
                    # assigned-groups line is already blue, bold and italic,
                    # and every other part is already escaped.
                    Paragraph(cell, cell_style) if cell else ""
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


# ─────────────────────────────────────────────────────────────────────────────
# STRUCTURE OF THE CODE — ALL-PROGRAMMES (MASTER) TIMETABLE
# Both the master timetable and the on-screen-format export are built as ONE
# reportlab `Table`:
#   - cells are real bordered Table cells, so they cannot visually overlap
#   - every cell's text is a real `Paragraph`, so line breaks actually render
#     (a raw string handed to a Table cell is treated as one unstyled line)
#   - every session cell is padded to a floor of three lines, so a short entry
#     is never a squeezed sliver and a long one just grows the row
#   - pagination is reportlab's own Table splitting; no manual page-fitting
#   - the heading is painted on the canvas via onFirstPage/onLaterPages, so it
#     repeats on every page however many the table ends up needing
# ─────────────────────────────────────────────────────────────────────────────

_SEMESTER_WORDS = {
    1: "FIRST",
    2: "SECOND",
    3: "THIRD",
    4: "FOURTH",
    5: "FIFTH",
    6: "SIXTH",
}


def _semester_word(number):
    """``1`` -> ``"FIRST"``; an unknown number falls back to an ordinal-ish tag."""
    try:
        number = int(number)
    except (TypeError, ValueError):
        return str(number or "").upper()
    return _SEMESTER_WORDS.get(number, f"{number}TH")


def _master_cell_text(entries, show_groups=False):
    """Build one master-timetable cell's text from its folded entries.

    Each entry is one line-block in UDSM field order, joined by a blank line:
    the type, the time it runs, the venue, the course, and the groups. A merged
    TD or workshop block drops the venue and course — the individual workshop
    names would swamp the cell — but keeps its time, like every other block.

    This delegates to :func:`_entry_lines`, the one place the field order is
    decided. It used to carry its own copy of that logic, which silently drifted
    and kept omitting the workshop time. ``show_groups`` is accepted for call
    compatibility but a folded block always carries the group list it was given.
    """
    blocks = [
        "\n".join(line for line in _entry_lines(entry) if line)
        for entry in entries or ()
    ]
    return "\n\n".join(block for block in blocks if block)


def _split_group_codes(text):
    """Split a stored group string into its individual codes.

    ``collect_master_entries`` stores a session's groups as a comma-separated
    list of codes, so ``"A1, A2"`` becomes ``["A1", "A2"]``. A code that is not
    comma-separated at all (a TD or workshop record carries a single
    ``group_code`` such as ``"EE C1"``) comes back as a one-item list.
    """
    if not text:
        return []
    return [part.strip() for part in str(text).split(",") if part.strip()]


def _compact_group_codes(codes):
    """Shorten a list of group codes by dropping a repeated programme prefix.

    Technical-drawing and workshop records name the group in full, so one
    block attending ``"EE C1"``, ``"EE C2"`` and ``"CE A1"`` reads much better
    as ``"EE C1, C2, CE A1"``: the first code of a run keeps its prefix and the
    rest of the run drops it. Order of first appearance is preserved so the
    output is stable and matches the order the records were collected in.
    """
    runs = []
    for code in codes:
        prefix, sep, suffix = str(code).rpartition(" ")
        if sep:
            for run_prefix, run_members in runs:
                if run_prefix == prefix:
                    run_members.append(suffix)
                    break
            else:
                runs.append((prefix, [suffix]))
        else:
            runs.append((None, [code]))

    parts = []
    for prefix, members in runs:
        if prefix is None:
            parts.extend(members)
        else:
            parts.append(f"{prefix} {members[0]}")
            parts.extend(members[1:])
    return ", ".join(parts)


def _entry_codes(entry):
    """The individual group codes an entry's ``groups`` string carries."""
    return _split_group_codes(entry.get("groups") or "")


def _groups_label(codes, all_groups):
    """Render the groups line for a block.

    A block reaching every known group reads ``ALL`` rather than listing twenty
    codes; otherwise the codes are shown compacted. A block with no groups
    assigned gets no line at all — the empty string.
    """
    if not codes:
        return ""
    if all_groups and set(codes) >= set(all_groups):
        return "ALL"
    return _compact_group_codes(codes)


# Courses whose LECTURES are taken by the entire first-year cohort, and so read
# "ALL" as their assigned group. Their records may name only some of the groups
# that actually attend — a lecture is not split by group — so listing group codes
# would misdescribe who has to be there.
#
# The rule is LECTURES ONLY. These same courses also run per-group tutorials and
# seminars (CL111 alone has 24 seminars), and those must keep naming their own
# groups, so the activity type is checked as well as the course code.
_WHOLE_COHORT_LECTURES = {"CL111", "MT161", "MT171", "ME101", "SC121"}

# Every lecture of a DS-prefixed course is likewise whole-cohort: the
# development-studies service courses run once for all programmes. This covers
# DS114, DS115, DS115_COET and any other DS course without listing them.
_WHOLE_COHORT_LECTURE_PREFIXES = ("DS",)


def _is_whole_cohort_lecture(entry):
    """True when this LECTURE is open to every group whatever its links say.

    Only lectures are affected: ``kind`` is the activity type lower-cased, so a
    tutorial, seminar or practical of the very same course returns False and
    keeps showing its own group codes. Workshop and technical-drawing blocks are
    never whole-cohort either.
    """
    if entry.get("kind") != "lecture":
        return False
    course = (entry.get("course_code") or "").strip().upper()
    if course in _WHOLE_COHORT_LECTURES:
        return True
    return course.startswith(_WHOLE_COHORT_LECTURE_PREFIXES)


# A session row whose course code is one of these names no real course: it is a
# stand-in for the workshop or technical-drawing allocations that carry the
# real detail. Drawing both would show one session twice -- a grey "WORKSHOP
# Practical" block sitting on top of the green workshop block for the very same
# slot. These are dropped whenever allocations actually describe their slot.
PLACEHOLDER_COURSE_CODES = {
    "WORKSHOP",
    "WORK SHOPS",
    "TD",
    "T/D",
    "T-DRAWING",
    "TECHNICAL DRAWING",
    "TECHNICALDRAWING",
}


def _is_placeholder(entry):
    """True when this session row is a stand-in for an allocation.

    A stand-in names no real course, only the activity it stands for. A row
    whose course code IS a real code is not a stand-in -- an ``ME101`` practical
    is a stand-in only because an ``ME101`` technical drawing shares its slot,
    which :func:`fold_same_sessions` decides.
    """
    return (entry.get("course_code") or "").strip().upper() in PLACEHOLDER_COURSE_CODES


def _slot_of(entry):
    """The slot an entry occupies: its day and the hours it covers.

    Keyed on the covered HOURS rather than on the clock strings, so a
    09:00-12:55 placeholder and a 09:00-13:00 allocation are recognised as the
    same slot even though their end times differ by five minutes. Matching on
    the raw start/end is what let the two be drawn as separate blocks.
    """
    return (entry.get("day"), frozenset(entry.get("hours") or ()))


def _identity_of(entry):
    """What an entry is a session OF, within its slot.

    Two entries in one slot that resolve to the same identity are the same
    session and collapse into a single block. A stand-in resolves to
    ``("allocation", "")``: it is whatever allocation shares its slot, decided
    by :func:`fold_same_sessions`.

    A workshop is identified by its SLOT ALONE, not by which craft it is. Every
    group's workshop in one daily slot is one workshop session for the week, so
    Carpentry, Welding and Masonry at Monday 09:00-13:00 belong in a single
    block reading ``Workshop`` with every attending group listed under it.
    Splitting them per craft was tried and rejected: it turned one session into
    three and buried the group list. Technical drawing keeps its course, so two
    genuinely different TD courses in one slot still read separately.
    """
    kind = entry.get("kind")
    course = (entry.get("course_code") or "").strip()
    if kind == "td":
        return ("td", course.upper())
    if kind == "workshop":
        return ("workshop", "")
    if _is_placeholder(entry):
        return ("allocation", "")
    return ("session", course.upper())


def fold_same_sessions(entries):
    """Collapse the records that describe one and the same session.

    Three pairs of records turn out to be a single session, and drawing both is
    what makes the timetable look like it has duplicated itself:

    * a workshop or technical-drawing allocation and the placeholder practical
      written to stand for it -- the placeholder is dropped and the allocation
      keeps its real group list;
    * every group's record of the same workshop slot -- one block whose groups
      line lists them all;
    * two session rows naming the same course in the same slot.

    A session whose course an allocation already covers in the same slot also
    folds in, which is how an ``ME101`` practical and the ``ME101`` technical
    drawing become the single block they really are.

    Returns one entry per surviving (slot, identity), each carrying the union of
    the hours and the groups of everything it absorbed, in first-seen order.
    Nothing is invented and nothing unrelated is merged: two different session
    courses in one slot stay two blocks, and a placeholder with no allocation
    behind it is kept rather than silently deleted.
    """
    order = []
    buckets = {}
    for entry in entries:
        key = (_slot_of(entry), _identity_of(entry))
        bucket = buckets.get(key)
        if bucket is None:
            bucket = _Bucket(key)
            buckets[key] = bucket
            order.append(key)
        bucket.append(entry)

    # Within each slot, work out which stand-ins are backed by a real
    # allocation, and which sessions an allocation already covers.
    by_slot = defaultdict(list)
    for key in order:
        by_slot[key[0]].append(key)

    keep = set()
    for keys in by_slot.values():
        families = {key[1][0] for key in keys}
        # A workshop covers its whole slot, so it needs no course to match on.
        has_workshop = "workshop" in families
        # Courses a technical drawing already speaks for in this slot, e.g. the
        # ME101 an ME101 practical is really the technical drawing of.
        td_courses = {
            key[1][1] for key in keys if key[1][0] == "td" and key[1][1]
        }
        for key in keys:
            family, course = key[1]
            if family == "allocation" and (has_workshop or td_courses):
                # The allocations already draw this slot, groups and all.
                continue
            if family == "session" and course in td_courses:
                # e.g. an ME101 practical that is really the ME101 technical
                # drawing: the allocation block stands in for both.
                continue
            keep.add(key)

    folded = []
    # Emitted in first-seen order, so the fold never reshuffles the week.
    for key in order:
        if key not in keep:
            continue
        bucket = buckets[key]
        merged = dict(bucket[0])
        merged["hours"] = set().union(
            *(set(e.get("hours") or ()) for e in bucket)
        )
        merged["key"] = key
        codes = []
        for entry in bucket:
            for code in _entry_codes(entry):
                if code not in codes:
                    codes.append(code)
        merged["_codes"] = codes
        folded.append(merged)
    return folded


class _Bucket(list):
    """A list of entries that collapse into one block, plus its merge key."""

    def __init__(self, key):
        super().__init__()
        self.key = key


def _merge_master_entries(entries, all_groups):
    """Group the raw master entries into the blocks the master grid draws.

    Delegates the "are these the same session?" judgement to
    :func:`fold_same_sessions` -- a workshop allocation and the placeholder
    practical written for it are one session, as are two allocations of the same
    course for different groups -- and then labels each surviving block: the
    groups line, ``ALL`` for a whole-cohort lecture, and the cell text.

    Genuinely different sessions in one slot (two different courses, say) are
    never merged.
    """
    return fold_for_display(entries, all_groups)


def fold_for_display(entries, all_group_codes):
    """Fold duplicate records and label what survives: what an export draws.

    This is the one place the "same day, same time, same course is one session"
    rule lives, shared by every export so they cannot drift apart. The
    on-screen timetable view deliberately does NOT use it yet -- it still draws
    one block per group for workshops and technical drawing.
    """
    return [
        _labelled_entry(entry, all_group_codes)
        for entry in fold_same_sessions(entries)
    ]


def _labelled_entry(entry, all_groups):
    """Attach the group list and the cell text to one folded block."""
    codes = entry.pop("_codes", None)
    if codes is None:
        codes = []
        for source in (entry,):
            for code in _entry_codes(source):
                if code not in codes:
                    codes.append(code)
    kind = entry.get("kind")
    if _is_whole_cohort_lecture(entry):
        # A whole-cohort lecture reads "ALL" even when only some groups are
        # linked to it: the codes would understate who attends.
        entry["groups"] = "ALL"
    else:
        entry["groups"] = _groups_label(codes, all_groups)

    if kind in ("td", "workshop"):
        # Type, time and the groups: the individual course/venue detail would
        # swamp the cell, but the time is stated like it is on every other
        # block so a workshop reads the same way as a lecture does.
        entry["type_label"] = "Technical Drawing" if kind == "td" else "Workshop"
        entry["label"] = "\n".join(
            part
            for part in (entry["type_label"], _time_span(entry), entry["groups"])
            if part
        )
    else:
        entry["label"] = "\n".join(
            part
            for part in (
                f"{entry.get('course_code') or ''} "
                f"{entry.get('type_label') or ''}".strip(),
                entry["groups"],
                entry.get("venue") or "",
            )
            if part
        )
    return entry


def _wrap_line(text, width, font, size):
    """Break one logical line of cell text so it fits `width` points.

    A single word can be wider than the box it belongs to (a long venue name
    with no spaces, say). Word wrapping alone would then leave it running out
    past the box border, so such a word is split mid-word as a last resort.
    """
    text = (text or "").strip()
    if not text or pdfmetrics.stringWidth(text, font, size) <= width:
        return [text] if text else [""]
    lines, cur = [], ""
    for word in text.split():
        if pdfmetrics.stringWidth(word, font, size) > width:
            # Too long to ever fit: flush what we have, then break it up.
            if cur:
                lines.append(cur)
                cur = ""
            for piece in _break_word(word, width, font, size):
                if (
                    cur
                    and pdfmetrics.stringWidth(f"{cur} {piece}", font, size)
                    > width
                ):
                    lines.append(cur)
                    cur = piece
                else:
                    cur = f"{cur} {piece}".strip()
            continue
        candidate = f"{cur} {word}".strip()
        if not cur or pdfmetrics.stringWidth(candidate, font, size) <= width:
            cur = candidate
        else:
            lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def _break_word(word, width, font, size):
    """Chop one over-long word into pieces that each fit `width`."""
    pieces, current = [], ""
    for char in word:
        candidate = current + char
        if current and pdfmetrics.stringWidth(candidate, font, size) > width:
            pieces.append(current)
            current = char
        else:
            current = candidate
    if current:
        pieces.append(current)
    return pieces


def _wrap_line(text, width, font, size):
    """Break one logical line of cell text so it fits `width` points.

    A single word can be wider than the box it belongs to (a long venue name
    with no spaces, say). Word wrapping alone would then leave it running out
    past the box border, so such a word is split mid-word as a last resort.
    """
    text = (text or "").strip()
    if not text or pdfmetrics.stringWidth(text, font, size) <= width:
        return [text] if text else [""]
    lines, cur = [], ""
    for word in text.split():
        if pdfmetrics.stringWidth(word, font, size) > width:
            # Too long to ever fit: flush what we have, then break it up.
            if cur:
                lines.append(cur)
                cur = ""
            for piece in _break_word(word, width, font, size):
                if (
                    cur
                    and pdfmetrics.stringWidth(f"{cur} {piece}", font, size)
                    > width
                ):
                    lines.append(cur)
                    cur = piece
                else:
                    cur = f"{cur} {piece}".strip()
            continue
        candidate = f"{cur} {word}".strip()
        if not cur or pdfmetrics.stringWidth(candidate, font, size) <= width:
            cur = candidate
        else:
            lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def _break_word(word, width, font, size):
    """Chop one over-long word into pieces that each fit `width`."""
    pieces, current = [], ""
    for char in word:
        candidate = current + char
        if current and pdfmetrics.stringWidth(candidate, font, size) > width:
            pieces.append(current)
            current = char
        else:
            current = candidate
    if current:
        pieces.append(current)
    return pieces


def _entry_parts(e):
    """A block's cell text as ``(text, is_groups)`` parts.

    The assigned-groups line is flagged so the renderer can draw it in
    :data:`~core.timetable_grid.GROUPS_STYLE` -- blue, bold and italic -- while
    the type, time, venue and course lines stay plain. Marking the line here,
    where the field order is decided, is what keeps the styling from having to
    guess that the groups are always last.
    """
    parts = [(e.get("type_label") or e.get("name") or "Session", False)]
    if e.get("start") and e.get("end"):
        parts.append((_time_span(e), False))
    if e.get("kind") in ("td", "workshop"):
        # Type, time and groups only: the course and venue would swamp the
        # cell, and the hour columns already place it. The time is still stated
        # here, exactly as it is on a taught session.
        if e.get("groups") and e["groups"] != "ALL":
            parts.append((e["groups"], True))
        return _flag_groups(parts, e)
    if e.get("venue"):
        parts.append((e["venue"], False))
    if e.get("course_code"):
        parts.append((e["course_code"], False))
    if e.get("groups"):
        # Includes the "ALL" marker: a whole-cohort lecture's assigned group is
        # the whole cohort, and the reader is told so on the page rather than
        # left to infer it from a missing line.
        parts.append((e["groups"], True))
    return _flag_groups(parts, e)


def _time_span(entry):
    """``"08:00-11:00"`` for an entry that states its own times, else "".

    Every block shows the time it runs, not only the taught sessions: a
    workshop or a technical drawing states it too, so all blocks read alike
    instead of leaving the reader to infer the hours from the grid columns.
    A period-only workshop (one imported from the raw matrix workbook) has no
    clock times of its own and is left out rather than invented.
    """
    if entry.get("start") and entry.get("end"):
        return f"{entry['start']}\u2013{entry['end']}"
    return ""


def _flag_groups(parts, entry):
    """Mark the line that IS the entry's group list, wherever it sits.

    A whole-cohort lecture's groups line reads ``ALL``, which says who attends
    just as much as a list of codes does, so it is emphasised as well.
    """
    groups = str(entry["groups"]).strip() if entry.get("groups") else ""
    if not groups:
        return parts
    return [
        (text, is_groups or text.strip() == groups) for text, is_groups in parts
    ]


def _entry_lines(e):
    """A block's cell text as plain strings (see :func:`_entry_parts`)."""
    return [text for text, _ in _entry_parts(e)]


def _stack_order(e):
    """Sort key deciding the order a day's blocks are laid down in.

    SHORTEST SESSION FIRST, then the earliest start, then the course code and
    finally the block's identity. The narrow blocks therefore reach the top of
    the day band and the wider ones fill in underneath, so a band reads as a
    one-hour session, then the two-hour ones, then the three-hour ones, and so
    on.

    This is a packing *preference*, never a hard constraint. The packer
    (``_DayFlowable._pack``) still drops every block onto the lowest position
    clear of the ones already laid down, so two simultaneous sessions stack
    instead of overlapping and no data can produce an invalid layout. A longer
    block that starts earlier than a shorter one simply ends up beneath it --
    the short one wins the top because it is shorter, and the grid is never
    distorted to let it.

    Laying the narrow blocks down first also pays off twice over: they keep the
    top of the band, and the wide ones left over get to sink into the gaps
    between them instead of each claiming a fresh full-height row.
    """
    hours = set(e.get("hours") or ())
    return (
        len(hours),
        min(hours) if hours else 0,
        e.get("course_code") or "",
        str(e.get("key")),
    )


def _occupy(runs, top, bottom):
    """Record ``[top, bottom)`` in a column's sorted, non-overlapping runs."""
    runs.append((top, bottom))
    runs.sort()
    merged = []
    for low, high in runs:
        if merged and low <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], high))
        else:
            merged.append((low, high))
    runs[:] = merged


class _DayFlowable(Flowable):
    """Render one day of the week as a single, tightly packed band.

    A band is a horizontal strip: a narrow day-label column on the left and one
    grid column per printed hour after it. The blocks inside are packed by
    ``_pack`` (shortest session on top, never overlapping) and the band is
    closed by a single heavy rule along its BOTTOM edge. That rule doubles as
    the separator between two consecutive days, so a week reads as one
    continuous grid with a clear line between the days and no gutter; a band's
    top edge is closed by whatever sits above it -- the neighbouring day's rule,
    or the hour header when the band starts a page.

    The band carries a ``BLOCK_GAP`` margin below its content, which is what
    puts that same clear space between the last block of one day and the first
    block of the next. The rule is drawn into the middle of that margin, so the
    space either side of the line is the same size as the space between any two
    sessions of a day.

    A day taller than the page frame is broken by ``split`` at a block
    boundary, never inside a block, so no session is ever sliced in two. Every
    chunk of a day keeps its own bordered label column, so the column never
    looks cut off, and the day's name is written exactly once -- on whichever
    chunk holds the MIDDLE of the day, not merely the first one. A chunk that
    runs on from a previous page therefore never repeats the name, and a chunk
    holding more than half the day has its name sit low, where the day's real
    centre falls rather than at the middle of the fragment on show.
    """

    FONT, BOLD = "Helvetica", "Helvetica-Bold"
    # Cell metrics. 6.5pt matches the cell size the classic A4 export already
    # uses; the leading is a comfortable 1.26x so consecutive lines of a session
    # read as separate lines rather than a paragraph of jammed text.
    CELL_SIZE, LEADING, MIN_LINES = 6.5, 8.2, 3
    PAD_X, PAD_Y, DAY_FONT = 3, 3, 10
    # A small clear margin drawn around every block, so two sessions stacked in
    # the same hour column never touch skin to skin. The packer reserves this
    # much space between neighbours, which is why the drawn box is inset by
    # half of it on all four sides.
    BLOCK_GAP = 1.5
    RULE_WIDTH = 1.1

    def __init__(
        self,
        day_entries,
        day_label,
        slots,
        col_width,
        day_width,
        boxes=None,
        day_offset=0.0,
        day_extent=None,
        height=None,
        label_here=True,
    ):
        super().__init__()
        self.day_entries = day_entries
        self.day_label = day_label
        self.slots = slots
        self.col_width = col_width
        self.day_width = day_width
        # Where this chunk sits inside its day, and how tall the whole day is.
        # An unsplit day is its own extent and starts at zero.
        self.day_offset = day_offset
        self.label_here = label_here
        self.boxes = self._pack(day_entries) if boxes is None else boxes
        self.content_height = self._content_height(self.boxes)
        self.day_extent = (
            self.content_height if day_extent is None else day_extent
        )
        # An empty band is nothing at all -- not even the trailing margin. A
        # chunk cut short by ``split`` is given the exact height it settled on,
        # which is its content plus whatever margin fitted.
        if height is not None:
            self.height = max(self.content_height, height)
        else:
            self.height = (
                self.content_height + self.BLOCK_GAP if self.boxes else 0.0
            )
        # The rule sits in the middle of whatever margin there is, so it is
        # always equally clear of the content above and below it.
        self.rule_y = (self.height - self.content_height) / 2.0

    def _content_height(self, boxes):
        """Height of the blocks themselves, ignoring the band margin."""
        if not boxes:
            return 0.0
        return max(box["bottom"] for box in boxes)

    def _clean_cut_levels(self):
        """The horizontal lines a day may be broken at without harm.

        A line is clean when, in every hour column, the lowest block above it
        ends before the highest block below it begins. Blocks in *different*
        columns may well overlap vertically -- they are drawn side by side -- so
        a line may pass through a block provided nothing sharing its columns
        sits on the other side of the break.

        Requiring the stricter "no block is crossed at all" looks tempting but
        is wrong: a densely packed day is a staircase, where the tallest block
        in every prefix runs past the next line, so NO line would qualify and
        the whole day would jump to a fresh page with the space below it empty.
        """
        columns = defaultdict(list)
        for box in self.boxes:
            for column in range(box["col"], box["col"] + box["colspan"]):
                columns[column].append(box)
        for blocks in columns.values():
            blocks.sort(key=lambda box: box["top"])

        levels = sorted({box["top"] for box in self.boxes})
        split_at = {column: 0 for column in columns}
        for level in levels:
            for column, blocks in columns.items():
                index = split_at[column]
                while index < len(blocks) and blocks[index]["top"] < level:
                    index += 1
                split_at[column] = index
            if all(
                not (0 < index < len(columns[column]))
                or columns[column][index - 1]["bottom"]
                <= columns[column][index]["top"]
                for column, index in split_at.items()
            ):
                # Columns the line does not actually cross are irrelevant: a
                # block alone in its column may span the break harmlessly.
                yield level

    def _label_offset(self):
        """Where to draw the rotated day name, or ``None`` for this chunk.

        The name belongs to the whole day, not to the fragment on show, so it
        is centred on the day's vertical midpoint and lands on whichever chunk
        owns that midpoint (``label_here``, decided in ``split``). A chunk that
        does not own it draws no name at all, which is what stops a continued
        day repeating itself. The position is clamped into this chunk's own
        content, so a midpoint that falls in the clear space at a page seam
        still puts the name against the blocks rather than in mid-air.
        """
        if not self.label_here or not self.boxes:
            return None
        return max(
            0.0,
            min(self.day_extent / 2.0 - self.day_offset, self.content_height),
        )

    # -- packing ---------------------------------------------------------
    def _pack(self, day_entries):
        """Lay a day's blocks out shortest-first without leaving dead bands.

        Each hour column keeps the vertical spans already taken. A block is
        dropped onto the LOWEST position free across every column it covers --
        including down into a hole left by a narrower neighbour -- so the band
        is as short as the data allows while the order the blocks were laid
        down in is preserved. Because the shortest blocks go down first they
        keep the top of the band; the wide ones then sink into whatever space
        is left instead of each claiming a fresh full-height row.

        Reclaiming those holes is worth about a quarter of the export's height:
        the busy sample week packs from 2633 to 1957 points, which is the
        difference between five pages and four.
        """
        idx = {h: i for i, h in enumerate(self.slots)}
        taken = [[] for _ in self.slots]
        boxes = []

        for e in sorted(day_entries, key=_stack_order):
            hours = sorted(set(e.get("hours") or ()))
            if not hours:
                continue
            start, end = idx.get(hours[0]), idx.get(hours[-1])
            if start is None or end is None:
                # An hour this band does not print is skipped, not squeezed in.
                continue
            span = end - start + 1
            columns = list(range(start, start + span))
            inner_w = self.col_width * span - 2 * self.PAD_X
            # Each wrapped piece keeps the is_groups flag of the field it came
            # from, so an emphasised group list survives being wrapped.
            wrapped = [
                (piece, is_groups)
                for line, is_groups in _entry_parts(e)
                for piece in _wrap_line(line, inner_w, self.FONT, self.CELL_SIZE)
            ]
            while len(wrapped) < self.MIN_LINES:
                wrapped.append(("", False))
            height = len(wrapped) * self.LEADING + 2 * self.PAD_Y
            top = self._lowest_free(taken, columns, height)
            bottom = top + height
            for column in columns:
                _occupy(taken[column], top, bottom)
            boxes.append(
                {
                    "col": start,
                    "colspan": span,
                    "top": top,
                    "bottom": bottom,
                    "lines": wrapped,
                    "fill": fill_color([e]),
                }
            )
        return boxes

    def _lowest_free(self, taken, columns, height):
        """The lowest y at which a ``height`` block clears all ``columns``.

        A new block has to clear every occupied run in the columns it covers by
        ``BLOCK_GAP``, so the only positions worth trying are the band floor and
        the top of each run plus that margin; the first that fits wins.
        """
        candidates = {0.0}
        for column in columns:
            candidates.update(high + self.BLOCK_GAP for _low, high in taken[column])
        for top in sorted(candidates):
            bottom = top + height
            if all(
                not any(top < high and low < bottom for low, high in taken[column])
                for column in columns
            ):
                return top
        return max(candidates)

    # -- page fitting ----------------------------------------------------
    def wrap(self, avail_width, avail_height):
        width = self.day_width + self.col_width * len(self.slots)
        return (width, self.height)

    def split(self, avail_width, avail_height):
        """Break a too-tall day between blocks, never through one.

        The break has to fall on a horizontal line that no block straddles, so
        the only candidates are the blocks' own top edges. The highest such
        line whose blocks all still fit the room left on the page is used, which
        puts as much of the day as possible on this page and sends the rest
        over whole. Every session is therefore drawn complete on exactly one
        page -- nothing is ever sliced through the middle.
        """
        if self.height <= avail_height or not self.boxes:
            return [self]

        ordered = sorted(self.boxes, key=lambda box: (box["top"], box["col"]))
        clean = set(self._clean_cut_levels())
        cut = None
        for level in sorted({box["top"] for box in ordered}):
            above = [box for box in ordered if box["top"] < level]
            below = [box for box in ordered if box["top"] >= level]
            if not above:
                # The band's own floor: no block lies above this line.
                continue
            if not below:
                # The topmost line: nothing left to carry over.
                break
            content = self._content_height(above)
            if content + self.BLOCK_GAP > avail_height:
                # Every line above this one holds at least as many blocks, so
                # none of them fits either: this is all the day that this page
                # can take, and the day starts here rather than being cut.
                break
            if level not in clean:
                # Something shares a column across this line and would collide.
                # A higher line may well be clear, so keep looking rather than
                # giving up and pushing the whole day to the next page.
                continue
            # The chunk takes the band margin too, but only when there is room
            # for it before the continuation starts; at a page-break seam a
            # missing margin is invisible, an overlap would not be.
            margin = self.BLOCK_GAP if content + self.BLOCK_GAP <= level else 0.0
            cut = (above, below, level, content + margin)
        if cut is None:
            # Not even the topmost block fits in what is left of this page, so
            # nothing is placed here and reportlab carries the whole day over.
            # A day that cannot fit an entire empty page raises reportlab's own
            # clear "too large on page" error rather than looping.
            return []

        above, below, level, height = cut
        # Exactly one chunk of a day carries its name, and ownership is settled
        # once and then handed straight down: a chunk that does not already own
        # the name never passes it on, so it can neither be written twice nor
        # lost. The owning chunk is the one holding the day's vertical midpoint;
        # a midpoint that lands in the clear space at a seam goes to the earlier
        # chunk, so the name appears as early in the day as it possibly can.
        midpoint = self.day_extent / 2.0
        first_owns = self.label_here and midpoint <= self.day_offset + height
        return [
            self._continuation(
                above, 0.0, self.day_offset, self.day_extent, height, first_owns
            ),
            self._continuation(
                below, level, self.day_offset + level, self.day_extent, None,
                # The continuation only inherits the name if this chunk was
                # holding it and the first half did not take it. Inheriting
                # "not first_owns" outright would hand the name to a page of a
                # day whose name was already written earlier.
                self.label_here and not first_owns,
            ),
        ]

    def _continuation(self, boxes, offset, day_offset, day_extent, height, label_here):
        """A band drawn from already-packed boxes, rebased onto y = 0.

        ``day_offset`` is this chunk's distance from the top of the whole day,
        which is what lets the day's name be centred on the day rather than on
        whichever fragment happens to be drawn. ``height`` is the clamped band
        height ``split`` settled on, or ``None`` to take the usual one.
        """
        rebased = [
            dict(box, top=box["top"] - offset, bottom=box["bottom"] - offset)
            for box in boxes
        ]
        return type(self)(
            self.day_entries,
            self.day_label,
            self.slots,
            self.col_width,
            self.day_width,
            boxes=rebased,
            day_offset=day_offset,
            day_extent=day_extent,
            height=height,
            label_here=label_here,
        )

    def draw(self):
        canvas = self.canv
        width = self.day_width + self.col_width * len(self.slots)
        canvas.saveState()
        
        # One heavy rule closes the band along its bottom edge, drawn into the
        # middle of the band margin so it sits equally clear of the day above
        # and the day below. It doubles as the separator between two
        # consecutive days, so the week reads as one continuous grid rather than
        # a stack of open-ended columns.
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(self.RULE_WIDTH)
        canvas.line(0, self.rule_y, width, self.rule_y)

        # The day-label column is bounded on every chunk of the day, so a day
        # carried over from the previous page never looks cut off down its side.
        canvas.setLineWidth(0.4)
        canvas.rect(0, 0, self.day_width, self.height, fill=0, stroke=1)

        # The name itself, on the chunk holding the middle of the whole day.
        label_offset = self._label_offset()
        if label_offset is not None:
            canvas.saveState()
            canvas.translate(
                self.day_width / 2.0, self.height - label_offset
            )
            canvas.rotate(90)
            canvas.setFillColor(colors.black)
            canvas.setFont(self.BOLD, self.DAY_FONT)
            canvas.drawCentredString(0, -self.DAY_FONT / 3.0, self.day_label)
            canvas.restoreState()
        
        # Draw column lines
        canvas.setStrokeColor(colors.HexColor("#9aa0a6"))
        canvas.setLineWidth(0.3)
        for i in range(len(self.slots) + 1):
            x = self.day_width + self.col_width * i
            canvas.line(x, 0, x, self.height)
        
        # Draw boxes. Vertically the drawn box is exactly the packed box, so the
        # BLOCK_GAP the packer reserved is the clear space a reader sees; at the
        # sides the box is inset by half of it so it never touches a column line.
        inset = self.BLOCK_GAP / 2.0
        for box in self.boxes:
            left = self.day_width + self.col_width * box["col"] + inset
            box_width = self.col_width * box["colspan"] - self.BLOCK_GAP
            box_top = self.height - box["top"]
            box_bottom = self.height - box["bottom"]
            canvas.setFillColor(colors.HexColor(box["fill"]))
            canvas.setStrokeColor(colors.black)
            canvas.setLineWidth(0.5)
            canvas.rect(left, box_bottom, box_width, box_top - box_bottom, fill=1, stroke=1)
            y = box_top - self.PAD_Y - self.LEADING * 0.78
            for line, is_groups in box["lines"]:
                if is_groups:
                    # The assigned groups: blue, bold and italic, so who
                    # attends a session is readable at a glance.
                    canvas.setFillColor(
                        colors.HexColor(GROUPS_STYLE["color"])
                    )
                    canvas.setFont(GROUPS_STYLE["font"], self.CELL_SIZE)
                else:
                    canvas.setFillColor(colors.black)
                    canvas.setFont(self.FONT, self.CELL_SIZE)
                canvas.drawString(left + self.PAD_X, y, line)
                y -= self.LEADING
        
        canvas.restoreState()


def _build_day_flowables(merged_entries, col_width, day_width):
    """Build one flowable per day, in week order, each closed by its own rule.

    Consecutive bands carry no spacer between them, so days butt up against
    each other and a single heavy rule marks where one ends and the next
    begins. A band too tall for the page is broken at a block boundary by
    ``_DayFlowable.split`` rather than being cut.
    """
    slots = list(range(GRID_HOUR_START, GRID_HOUR_END + 1))
    present = {e["day"] for e in merged_entries}
    day_order = [d for d in DAY_ORDER if d in present] + [d for d in WEEKEND_ORDER if d in present]
    by_day = {}
    for e in merged_entries:
        if e["hours"]:
            by_day.setdefault(e["day"], []).append(e)
    
    flowables = []
    for day in day_order:
        day_entries = by_day.get(day, [])
        if day_entries:
            day_label = Day(day).label
            flowable = _DayFlowable(day_entries, day_label, slots, col_width, day_width)
            flowables.append(flowable)
    
    return flowables, slots


# The hour-column header is painted on the canvas above the frame, so it costs
# no body space. It is kept deliberately shallow and sits flush on the frame
# top, which both saves vertical room and closes the top of the first day band
# on every page.
HEADER_BAND = 6.0 * mm
HEADER_TITLE_OFFSET = 6.6 * mm
HEADER_SUBTITLE_OFFSET = 10.4 * mm
HEADER_MARGIN = 17.5 * mm
FOOTER_MARGIN = 13.0 * mm
FOOTER_TEXT = "Generated from"
PORTAL_NAME = "CoET Timetable Portal"


def _draw_master_header(canvas, doc, heading_lines, slots, col_width, day_width):
    """Paint the title block and the hour-column header above the frame."""
    canvas.saveState()
    pw, ph = landscape(A4)
    canvas.setFillColor(colors.black)
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawCentredString(pw / 2, ph - HEADER_TITLE_OFFSET, heading_lines[0])
    canvas.setFont("Helvetica", 8.5)
    canvas.drawCentredString(pw / 2, ph - HEADER_SUBTITLE_OFFSET, heading_lines[1])

    # The band's lower edge is the frame top, so the first day band on the page
    # starts directly under a closed border instead of leaving the hour columns
    # hanging into a gap.
    x0 = doc.leftMargin
    bottom = ph - doc.topMargin
    top = bottom + HEADER_BAND
    canvas.setFillColor(colors.HexColor("#dce6f1"))
    canvas.setStrokeColor(colors.black)
    canvas.setLineWidth(0.4)
    canvas.rect(x0, bottom, day_width + col_width * len(slots), top - bottom, fill=1, stroke=1)
    canvas.setFillColor(colors.black)
    canvas.setFont("Helvetica-Bold", 7.5)
    for i, hour in enumerate(slots):
        x = x0 + day_width + col_width * i
        canvas.drawCentredString(
            x + col_width / 2.0,
            bottom + (top - bottom) / 2 - 2.7,
            f"{hour:02d}:00\u2013{hour + 1:02d}:00",
        )
        if i:
            canvas.setLineWidth(0.3)
            canvas.line(x, bottom, x, top)
    canvas.restoreState()


def _draw_master_footer(canvas, doc, portal_url=None, date_text=""):
    """Paint the footer strip: what generated the file, and which page it is.

    Drawn on the canvas rather than added to the element list, so every page
    carries it -- not just the last. The portal name is a live link back to the
    site that produced the export whenever the caller knows its address.
    """
    canvas.saveState()
    pw, ph = landscape(A4)
    right = pw - doc.rightMargin
    baseline = FOOTER_MARGIN - 4.6 * mm
    canvas.setFont("Helvetica", 7.5)

    lead = canvas.stringWidth(FOOTER_TEXT + " ", "Helvetica", 7.5)
    name = canvas.stringWidth(PORTAL_NAME, "Helvetica-Bold", 7.5)
    tail = canvas.stringWidth(" on " + date_text, "Helvetica", 7.5)
    total = lead + name + tail
    # Centred in the page, so the provenance line reads as a footer rather than
    # as a stray note in the left margin.
    x = (pw - total) / 2.0
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor("#444444"))
    canvas.drawString(x, baseline, FOOTER_TEXT + " ")
    x += lead
    canvas.setFillColor(colors.HexColor("#0b4f9e"))
    canvas.setFont("Helvetica-Bold", 7.5)
    canvas.drawString(x, baseline, PORTAL_NAME)
    if portal_url:
        # A clickable link back to the portal the export came from.
        canvas.linkURL(
            portal_url,
            (x - 1, baseline - 1.5, x + name + 1, baseline + 7.5),
            relative=0,
            thickness=0,
        )
    canvas.setLineWidth(0.3)
    canvas.setStrokeColor(colors.HexColor("#0b4f9e"))
    canvas.line(x, baseline - 1.4, x + name, baseline - 1.4)
    x += name
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor("#444444"))
    canvas.drawString(x, baseline, " on " + date_text)

    canvas.setFont("Helvetica-Bold", 8)
    canvas.setFillColor(colors.black)
    canvas.drawRightString(right, baseline, f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


def render_udsm_master_timetable(entries, semester, year_of_study=1, out=None, portal_url=None):
    all_groups = set(StudentGroup.objects.values_list("code", flat=True))
    merged_entries = _merge_master_entries(entries, all_groups)

    # A4 landscape, not A3: the grid is thirteen hour columns of short text,
    # so the old page was far wider than the content needed and had to be
    # scrolled sideways at 100% zoom. The day-label column is only as wide as
    # its rotated name needs, which hands the rest back to the hour columns.
    doc = SimpleDocTemplate(
        out,
        pagesize=landscape(A4),
        leftMargin=8 * mm,
        rightMargin=8 * mm,
        topMargin=HEADER_MARGIN,
        bottomMargin=FOOTER_MARGIN,
        title="University Master Timetable",
    )

    usable_w = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
    day_width = 22
    n_slots = GRID_HOUR_END - GRID_HOUR_START + 1
    col_width = (usable_w - day_width) / n_slots

    day_flowables, slots = _build_day_flowables(merged_entries, col_width, day_width)
    year_note = f" \u00b7 {_ordinal(int(year_of_study)).upper()} YEAR" if year_of_study and int(year_of_study) > 1 else ""
    subtitle = f"TEACHING TIMETABLE FOR {_semester_word(semester.semester)} SEMESTER {semester.academic_year}{year_note}"
    heading = [
        "UNIVERSITY OF DAR ES SALAAM",
        f"TEACHING TIMETABLE FOR {_semester_word(semester.semester)} SEMESTER "
        f"{semester.academic_year}{year_note} \u00b7 SEMESTER {semester.semester}",
    ]
    date_text = datetime.date.today().strftime("%d %B %Y")

    def page_cb(canvas, doc_):
        _draw_master_header(canvas, doc_, heading, slots, col_width, day_width)
        _draw_master_footer(canvas, doc_, portal_url=portal_url, date_text=date_text)

    elements = day_flowables if day_flowables else [
        Paragraph(
            "No timetable sessions scheduled for this selection.",
            ParagraphStyle("e", alignment=TA_CENTER),
        )
    ]
    doc.build(elements, onFirstPage=page_cb, onLaterPages=page_cb)
    return doc
