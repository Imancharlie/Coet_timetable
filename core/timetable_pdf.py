"""Programme timetable PDF generation (reportlab).

Produces a classic weekly grid, A4 portrait: the DAYS run across the first row
and the hourly TIME slots (07:00-07:55 .. 19:00-19:55) run down the first
column, matching the on-screen timetable in ``core/timetable.html``. Sessions
that span multiple hourly slots merge their cells on that day. Each cell is
shaded by activity type — Lecture faded grey, Workshop faded green, Technical
Drawing faded pink. External activities (workshops / technical drawing) that
only carry a Morning/Afternoon period are placed in the morning (09:00-12:55)
or afternoon (15:00-18:55) range.
"""

import datetime
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
        "name": session.course_code,
        "type_label": session.get_activity_type_display(),
        "venue": venue,
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
    return {
        "key": ("workshop", rec.pk),
        "day": rec.day,
        "hours": (
            _covered_hours(rec.start_time, rec.end_time)
            if rec.start_time is not None
            else _period_hours(rec.time_period)
        ),
        "label": name,
        "kind": "workshop",
        "course_code": rec.course_code,
        "name": name,
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
        "name": rec.course_code,
        "type_label": "Technical Drawing",
        "venue": rec.venue,
        "groups": rec.group_code,
        "note": "",
    }


def _workshop_sort_time(rec):
    """Clock-time a workshop occupies, for a deterministic day ordering."""
    if rec.start_time is not None:
        return rec.start_time
    if rec.time_period == TimePeriod.MORNING:
        return datetime.time(9)
    if rec.time_period == TimePeriod.AFTERNOON:
        return datetime.time(15)
    return datetime.time(0)


def _workshop_entries(qs):
    """Turn a WorkshopAllocation queryset into clean display entries.

    Applies the timetable rules that a student group has exactly one workshop
    per day: when the source carries two *different* workshops for the same
    group on the same day (a data inconsistency), only the first is kept so the
    grid never shows two workshops for one group on one day. Week-run splits of
    the same workshop (e.g. Wk 1-6 and Wk 8-13) are all preserved — they are
    one workshop with distinct valid sessions. The chosen records are ordered
    deterministically by day/time so cells render stably.
    """
    by_group_day = {}
    for rec in qs:
        by_group_day.setdefault((rec.group_code, rec.day), []).append(rec)

    kept = []
    for records in by_group_day.values():
        by_identity = {}
        for rec in records:
            by_identity.setdefault(_workshop_display_name(rec), []).append(rec)
        chosen = by_identity
        if len(by_identity) > 1:
            best = min(
                by_identity,
                key=lambda name: (
                    _workshop_sort_time(by_identity[name][0]),
                    by_identity[name][0].pk,
                ),
            )
            chosen = {best: by_identity[best]}
        for records_of_identity in chosen.values():
            kept.extend(
                sorted(
                    records_of_identity,
                    key=lambda r: (_workshop_sort_time(r), r.pk),
                )
            )
    return [_workshop_entry(rec) for rec in kept]


def _covered_hours(start, end):
    """Hour slots (7..19) a session covers given its start/end times."""
    hours = set()
    for h in _slot_hours():
        if start is not None and end is not None:
            if start < datetime.time(h + 1) and end > datetime.time(h):
                hours.add(h)
    return hours


def _period_hours(time_period):
    """Hourly slots a morning/afternoon workshop covers.

    A morning workshop is one four-hour session 09:00-12:55 and an afternoon
    workshop one four-hour session 15:00-18:55, matching the university rules.
    """
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


def collect_entries(programme, semester, group=None, year=None):
    """Gather timetable entries for one programme's groups in a semester.

    When ``group`` is given only that single student group's timetable is
    collected (sessions a group actually attends plus its workshops/TDs),
    which backs the per-group export. ``year`` optionally filters workshop
    allocations to a year of study (records without a year apply to every
    year).
    """
    if group is not None:
        return collect_group_entries(group, semester, year=year)

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
    if year:
        workshops = workshops.filter(
            Q(year_of_study__isnull=True) | Q(year_of_study=year)
        )
    entries.extend(_workshop_entries(workshops))

    tds = TechnicalDrawingAllocation.objects.filter(
        semester=semester, group_code__in=group_codes
    )
    for rec in tds:
        entries.append(_td_entry(rec))

    return entries


def collect_group_entries(group, semester, year=None):
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
    if year:
        workshops = workshops.filter(
            Q(year_of_study__isnull=True) | Q(year_of_study=year)
        )
    entries.extend(_workshop_entries(workshops))

    tds = TechnicalDrawingAllocation.objects.filter(
        semester=semester, group_code=group.code
    )
    for rec in tds:
        entries.append(_td_entry(rec))

    return entries


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
    grid = build_time_day_grid(entries)
    return time_day_grid_to_table(grid, show_groups=show_groups)


def render_programme_timetable(programme, semester, year_of_study=1, out=None):
    """Render the programme timetable PDF to `out` (file-like or a path)."""
    entries = collect_entries(programme, semester, year=year_of_study)
    show_groups = StudentGroup.objects.filter(programme=programme).count() > 1
    return _render_grid(
        entries,
        title=f"{programme.name.upper()}",
        subtitle=f"{_ordinal(year_of_study)} YEAR · {semester.academic_year}",
        semester=semester,
        doc_title=f"{programme.name} Timetable",
        show_groups=show_groups,
        out=out,
    )


def render_group_timetable(group, semester, year_of_study=1, out=None):
    """Render ONE student group's timetable PDF to `out` (file-like/path)."""
    entries = collect_entries(
        group.programme, semester, group=group, year=year_of_study
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
        out=out,
    )


def _render_grid(entries, title, subtitle, semester, doc_title, out, show_groups=False):
    """Render the weekly grid PDF to `out` (file-like or a path)."""
    data, spans, fills = build_grid(entries, show_groups=show_groups)

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
    elements.append(Spacer(1, 6 * mm))
    export_date = datetime.date.today().strftime("%d %B %Y")
    elements.append(Paragraph(f"Prepared for personal use · {export_date}", foot_style))

    doc.build(elements)
    return doc