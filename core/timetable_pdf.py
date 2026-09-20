"""Programme timetable PDF generation (reportlab).

Produces a weekly grid — TIME column (07:00-07:55 .. 19:00-19:55) plus
MONDAY..FRIDAY day columns — matching the layout of the university's
"skeleton" timetable PDF. Sessions that span multiple hourly slots merge
their cells on that day. External activities (workshops / technical drawing)
that only carry a Morning/Afternoon period are placed in the full morning
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

DAY_ORDER = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"]

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


def _entry_label(session, groups_text=""):
    venue = session.venue.name if session.venue and session.venue.name else "-"
    lines = [f"{session.course_code} {session.get_activity_type_display()}"]
    if groups_text:
        lines.append(groups_text)
    lines.append(venue)
    return "\n".join(lines)


def _workshop_entry(rec):
    label = f"{rec.course_code} WORKSHOP"
    if rec.workshop:
        label += f" · {rec.workshop}"
    lines = [label]
    bits = [rec.group_code]
    if rec.venue:
        bits.append(rec.venue)
    lines.append(" ".join(bits))
    if rec.week_start and rec.week_end:
        lines.append(f"Wk {rec.week_start}-{rec.week_end}")
    return "\n".join(lines)


def _td_entry(rec):
    return "\n".join([
        f"{rec.course_code} TECHNICAL DRAWING",
        f"{rec.group_code}",
        rec.venue or "-",
    ])


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


def collect_entries(programme, semester):
    """Gather timetable entries for one programme's groups in a semester."""
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
        entries.append(
            {
                "key": ("session", session.pk),
                "day": session.day,
                "hours": _covered_hours(session.start_time, session.end_time),
                "label": _entry_label(session, groups_text),
            }
        )

    workshops = WorkshopAllocation.objects.filter(
        semester=semester, group_code__in=group_codes
    )
    for rec in workshops:
        entries.append(
            {
                "key": ("workshop", rec.pk),
                "day": rec.day,
                "hours": (
                    _covered_hours(rec.start_time, rec.end_time)
                    if rec.start_time is not None
                    else _period_hours(rec.time_period)
                ),
                "label": _workshop_entry(rec),
            }
        )

    tds = TechnicalDrawingAllocation.objects.filter(
        semester=semester, group_code__in=group_codes
    )
    for rec in tds:
        entries.append(
            {
                "key": ("td", rec.pk),
                "day": rec.day,
                "hours": _covered_hours(rec.start_time, rec.end_time),
                "label": _td_entry(rec),
            }
        )

    return entries


def build_grid(entries):
    """Return table data (list of rows) plus (from,to) SPAN commands.

    Column layout: [TIME, MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY].
    Row 0 is the weekday header; rows 1..n are hourly slots. Consecutive
    slots whose entry signature is identical are merged via SPAN.
    """
    hours = _slot_hours()
    data = [["TIME"] + DAY_ORDER]
    for h in hours:
        data.append([f"{h:02d}:00 - {h:02d}:55"] + [""] * len(DAY_ORDER))

    by_col = {day: i + 1 for i, day in enumerate(DAY_ORDER)}
    spans = []

    for day in DAY_ORDER:
        col = by_col[day]
        cells = []
        for h in hours:
            cells.append(
                sorted(
                    (
                        e
                        for e in entries
                        if e["day"] == day and h in e["hours"]
                    ),
                    key=lambda e: e["label"],
                )
            )

        row = 1
        while row <= len(hours):
            current = cells[row - 1]
            if not current:
                row += 1
                continue
            sig = [e["key"] for e in current]
            end = row
            while end < len(hours) and [e["key"] for e in cells[end]] == sig:
                end += 1
            body = "\n\n".join(e["label"] for e in current)
            data[row][col] = body
            if end - row + 1 > 1:
                spans.append(((col, row), (col, end)))  # (col,row) start
            row = end + 1

    return data, spans


def render_programme_timetable(programme, semester, year_of_study=1, out=None):
    """Render the programme timetable PDF to `out` (file-like or a path)."""
    entries = collect_entries(programme, semester)
    data, spans = build_grid(entries)

    doc = SimpleDocTemplate(
        out,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title=f"{programme.name} Timetable",
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
        "cell", fontName="Helvetica", fontSize=7, leading=8.5, alignment=TA_CENTER
    )
    time_style = ParagraphStyle(
        "time",
        fontName="Helvetica-Bold",
        fontSize=7,
        alignment=TA_CENTER,
        leading=8.5,
    )
    foot_style = ParagraphStyle(
        "foot", fontName="Helvetica", fontSize=8, alignment=TA_CENTER
    )

    title = f"{programme.name.upper()} {_ordinal(year_of_study)} YEAR {semester.academic_year} TIME TABLE"
    elements = [
        Paragraph(title, title_style),
        Paragraph(f"SEMESTER {semester.semester} · {semester.academic_year}", sub_style),
        Spacer(1, 4 * mm),
    ]

    widths = None
    usable = A4[0] - 28 * mm
    time_w = 46
    day_w = (usable - time_w) / len(DAY_ORDER)
    widths = [time_w] + [day_w] * len(DAY_ORDER)

    table_data = [[Paragraph(item, head_style) for item in data[0]]]
    for row in data[1:]:
        table_data.append(
            [
                Paragraph(row[0], time_style),
            ]
            + [Paragraph(cell, cell_style) if cell else "" for cell in row[1:]]
        )

    grid = Table(table_data, colWidths=widths, repeatRows=1)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dce6f1")),
        ("SPAN", (0, 0), (0, 0)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]
    for start, end in spans:
        style.append(("SPAN", start, end))
    grid.setStyle(TableStyle(style))

    elements.append(grid)
    elements.append(Spacer(1, 6 * mm))
    elements.append(Paragraph("Prepared for personal use", foot_style))

    doc.build(elements)
    return doc