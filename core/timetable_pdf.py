"""Programme timetable PDF generation (reportlab).

Produces a classic weekly grid, A4 portrait: the DAYS run across the first row
and the hourly TIME slots (07:00-07:55 .. 19:00-19:55) run down the first
column, matching the on-screen timetable in ``core/timetable.html``. Sessions
that span multiple hourly slots merge their cells on that day. Each cell is
shaded by activity type — Lecture faded grey, Workshop faded green, Technical
Drawing faded pink. External activities (workshops / technical drawing) that
only carry a Morning/Afternoon period are placed in the full morning
(08:00-12:55) or afternoon (13:00-17:55) range.
"""

import datetime

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
    Session,
    StudentGroup,
    TechnicalDrawingAllocation,
    TimePeriod,
    WorkshopAllocation,
)
from core.timetable_grid import build_time_day_grid, time_day_grid_to_table

HOUR_START = 7
HOUR_END = 19  # inclusive last hour (19:00-19:55)


def _slot_hours():
    return list(range(HOUR_START, HOUR_END + 1))


def _ordinal(n):
    if n % 100 in (11, 12, 13):
        return f"{n}th"
    tail = n % 10
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(tail, "th")
    return f"{n}{suffix}"


def _session_entry(session, groups_text=""):
    venue = session.venue.name if session.venue and session.venue.name else ""
    lines = [f"{session.course_code} {session.get_activity_type_display()}"]
    if groups_text:
        lines.append(groups_text)
    lines.append(venue or "-")
    return {
        "key": ("session", session.pk),
        "day": session.day,
        "hours": _covered_hours(session.start_time, session.end_time),
        "label": "\n".join(lines),
        "kind": session.activity_type.lower(),
        "course_code": session.course_code,
        "type_label": session.get_activity_type_display(),
        "venue": venue,
        "groups": groups_text,
    }


def _workshop_entry(rec):
    label = f"{rec.course_code} WORKSHOP"
    if rec.workshop:
        label += f" · {rec.workshop}"
    lines = [label]
    bits = [rec.group_code]
    if rec.venue:
        bits.append(rec.venue)
    lines.append(" ".join(bits))
    note = ""
    if rec.week_start and rec.week_end:
        note = f"Wk {rec.week_start}-{rec.week_end}"
        lines.append(note)
    return {
        "key": ("workshop", rec.pk),
        "day": rec.day,
        "hours": (
            _covered_hours(rec.start_time, rec.end_time)
            if rec.start_time is not None
            else _period_hours(rec.time_period)
        ),
        "label": "\n".join(lines),
        "kind": "workshop",
        "course_code": rec.course_code,
        "type_label": "Workshop",
        "venue": rec.venue,
        "groups": rec.group_code,
        "note": note,
    }


def _td_entry(rec):
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
        "type_label": "Technical Drawing",
        "venue": rec.venue,
        "groups": rec.group_code,
    }


def _covered_hours(start, end):
    """Hour slots (7..19) a session covers given its start/end times."""
    hours = set()
    for h in _slot_hours():
        if start is not None and end is not None:
            if start < datetime.time(h + 1) and end > datetime.time(h):
                hours.add(h)
    return hours


def _period_hours(time_period):
    if time_period == TimePeriod.MORNING:
        return set(range(8, 13))
    if time_period == TimePeriod.AFTERNOON:
        return set(range(13, 18))
    return set()


def collect_entries(programme, semester, group=None):
    """Gather timetable entries for one programme's groups in a semester.

    When ``group`` is given only that single student group's timetable is
    collected (sessions a group actually attends plus its workshops/TDs),
    which backs the per-group export.
    """
    if group is not None:
        return collect_group_entries(group, semester)

    groups = list(StudentGroup.objects.filter(programme=programme))
    group_codes = {g.code for g in groups}
    group_pk_by_code = {g.code: g.pk for g in groups}

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

    workshops = WorkshopAllocation.objects.filter(
        semester=semester, group_code__in=group_codes
    )
    for rec in workshops:
        entries.append(_workshop_entry(rec))

    tds = TechnicalDrawingAllocation.objects.filter(
        semester=semester, group_code__in=group_codes
    )
    for rec in tds:
        entries.append(_td_entry(rec))

    return entries


def collect_group_entries(group, semester):
    """Gather timetable entries for ONE student group in a semester.

    Sessions are those the group actually attends (via ``SessionGroup``);
    workshops and technical drawing slots are matched on the group's code.
    """
    entries = []

    sessions = (
        Session.objects.filter(semester=semester, session_groups__group=group)
        .select_related("venue")
        .distinct()
    )
    for session in sessions:
        entries.append(_session_entry(session))

    workshops = WorkshopAllocation.objects.filter(
        semester=semester, group_code=group.code
    )
    for rec in workshops:
        entries.append(_workshop_entry(rec))

    tds = TechnicalDrawingAllocation.objects.filter(
        semester=semester, group_code=group.code
    )
    for rec in tds:
        entries.append(_td_entry(rec))

    return entries


def build_grid(entries):
    """Return table data (list of rows), SPAN commands and fill colours.

    Classic layout: column 0 is TIME, columns 1..n are the weekdays, row 0 is
    the DAY header. Uses the shared ``core.timetable_grid`` builder so the PDF
    matches the on-screen timetable. SPAN commands merge multi-slot session
    blocks vertically; fills carries one BACKGROUND command per block,
    colour-coded by activity type (Lecture grey, Workshop green, TD pink).
    """
    grid = build_time_day_grid(entries)
    return time_day_grid_to_table(grid)


def render_programme_timetable(programme, semester, year_of_study=1, out=None):
    """Render the programme timetable PDF to `out` (file-like or a path)."""
    entries = collect_entries(programme, semester)
    return _render_grid(
        entries,
        title=f"{programme.name.upper()}",
        subtitle=f"{_ordinal(year_of_study)} YEAR · {semester.academic_year}",
        semester=semester,
        doc_title=f"{programme.name} Timetable",
        out=out,
    )


def render_group_timetable(group, semester, year_of_study=1, out=None):
    """Render ONE student group's timetable PDF to `out` (file-like/path)."""
    entries = collect_entries(group.programme, semester, group=group)
    return _render_grid(
        entries,
        title=f"{group.programme.name.upper()}",
        subtitle=(
            f"{_ordinal(year_of_study)} YEAR · GROUP {group.code} · "
            f"{semester.academic_year}"
        ),
        semester=semester,
        doc_title=f"{group.programme.code} {group.code} Timetable",
        out=out,
    )


def _render_grid(entries, title, subtitle, semester, doc_title, out):
    """Render the weekly grid PDF to `out` (file-like or a path)."""
    data, spans, fills = build_grid(entries)

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
    foot_style = ParagraphStyle(
        "foot", fontName="Helvetica", fontSize=8, alignment=TA_CENTER
    )

    elements = [
        Paragraph(title, title_style),
        Paragraph(subtitle, sub_style),
        Paragraph(f"SEMESTER {semester.semester}", sub_style),
        Spacer(1, 4 * mm),
    ]

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
            + [Paragraph(cell, cell_style) if cell else "" for cell in row[1:]]
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
    elements.append(Spacer(1, 6 * mm))
    export_date = datetime.date.today().strftime("%d %B %Y")
    elements.append(Paragraph(f"Prepared for personal use · {export_date}", foot_style))

    doc.build(elements)
    return doc