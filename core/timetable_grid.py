"""Weekly timetable grid builder: TIME as rows, DAYS as columns.

Builds on the resolved entries from ``core.timetable_pdf.collect_entries`` so
the on-screen timetable (``core/timetable.html``) and the PDF export render
the same layout. The row headers are the hourly slots actually covered by the
selected programme/semester (07:00-07:55 .. 19:00-19:55 window); the columns
are MONDAY..FRIDAY (plus SATURDAY/SUNDAY when the timetable has weekend
activity). Sessions that occupy consecutive slots are merged into a single
vertical cell. Every render cell keeps the session type so templates and the
PDF can shade Lecture (grey), Workshop (green) and Technical Drawing (pink).
"""

from core.models import Day

DAY_ORDER = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"]
WEEKEND_ORDER = ["SATURDAY", "SUNDAY"]

# Light "faded" fills used by the PDF. Priority when a merged block contains
# several session types: TD > Workshop > Lecture > Tutorial > Seminar > Practical.
FILL_COLORS = {
    "td": "#fce7f3",
    "workshop": "#dcfce7",
    "lecture": "#d1d5db",
    "tutorial": "#dbeafe",
    "seminar": "#ede9fe",
    "practical": "#fef3c7",
}
_FILL_PRIORITY = ["td", "workshop", "lecture", "tutorial", "seminar", "practical"]


def _slot(hour):
    return {
        "hour": hour,
        "start": f"{hour:02d}:00",
        "end": f"{hour:02d}:55",
        "label": f"{hour:02d}:00-{hour:02d}:55",
    }


def build_time_day_grid(entries):
    """Turn ``collect_entries`` output into a time-row/day-column grid.

    Returns a dict with ``slots``, ``days`` (column headers) and ``rows``.
    ``rows`` is one dict per time slot: ``{"slot": {...}, "cols": [...]}``
    aligned to ``days``. Each column cell is ``None`` when the row is already
    consumed by the ``rowspan`` cell above it, ``{"empty": True}`` for a free
    period, or ``{"row", "rowspan", "entries"}`` for a merged session block.
    """
    if not entries:
        return {"slots": [], "days": [], "rows": []}

    present_days = {e["day"] for e in entries}
    day_order = [d for d in DAY_ORDER if d in present_days]
    day_order += [d for d in WEEKEND_ORDER if d in present_days]
    days = [{"day": d, "label": Day(d).label} for d in day_order]

    covered = set()
    for e in entries:
        covered.update(e["hours"])
    if not covered:
        return {"slots": [], "days": [], "rows": []}

    min_hour, max_hour = min(covered), max(covered)
    slots = list(range(min_hour, max_hour + 1))
    slot_defs = [_slot(hour) for hour in slots]
    n_slots = len(slots)

    by_slot = {}
    for e in entries:
        for hour in e["hours"]:
            by_slot.setdefault((e["day"], hour), []).append(e)

    rows = [
        {"slot": slot_defs[i], "cols": [None] * len(days)} for i in range(n_slots)
    ]

    for di, day in enumerate(day_order):
        r = 0
        while r < n_slots:
            hour = slots[r]
            covering = sorted(
                by_slot.get((day, hour), []),
                key=lambda e: (e["course_code"], str(e["key"])),
            )
            if not covering:
                rows[r]["cols"][di] = {"empty": True}
                r += 1
                continue
            sig = frozenset(e["key"] for e in covering)
            end = r + 1
            while end < n_slots:
                nxt = by_slot.get((day, slots[end]), [])
                if frozenset(e["key"] for e in nxt) == sig:
                    end += 1
                else:
                    break
            rows[r]["cols"][di] = {
                "row": r,
                "rowspan": end - r,
                "entries": covering,
            }
            r = end

    return {"slots": slot_defs, "days": days, "rows": rows}


def fill_color(entries):
    """Best shading fill for a set of entries (first matching priority kind)."""
    kinds = {e["kind"] for e in entries}
    for kind in _FILL_PRIORITY:
        if kind in kinds:
            return FILL_COLORS[kind]
    return "#ffffff"


def time_day_grid_to_table(grid):
    """Flatten a grid for reportlab's Table.

    Column 0 is the TIME column; columns 1..n are weekdays. Returns
    (data, spans, fills); ``spans`` are vertical SPAN commands for merged
    session blocks and ``fills`` carries one BACKGROUND command per block,
    colour-coded by session type.
    """
    day_headers = grid["days"]
    rows = grid["rows"]
    if not day_headers:
        return [["TIME", "DAY"]], [], []

    data = [["TIME"] + [d["day"] for d in day_headers]]
    for row in rows:
        line = [row["slot"]["label"]]
        for cell in row["cols"]:
            if cell is None or cell.get("empty"):
                line.append("")
            else:
                line.append("\n\n".join(e["label"] for e in cell["entries"]))
        data.append(line)

    spans = []
    fills = []
    for di in range(len(day_headers)):
        for r, row in enumerate(rows):
            cell = row["cols"][di]
            if cell is None or cell.get("empty"):
                continue
            col = di + 1
            rr = r + 1
            end_row = rr + cell["rowspan"] - 1
            if cell["rowspan"] > 1:
                spans.append(((col, rr), (col, end_row)))
            fills.append(((col, rr), (col, end_row), fill_color(cell["entries"])))

    return data, spans, fills