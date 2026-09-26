"""Weekly timetable grid builders.

``build_time_day_grid`` produces the classic time-row/day-column layout shared
with the PDF export (``core/timetable_pdf.build_grid``). ``build_day_time_grid``
produces the on-screen transposed layout — DAYS as rows (always the full
Monday..Friday range, so labels are the full weekday names) and HOURLY time
slots as columns — described by the default timetable view
(``core/timetable.html``). Both build on the resolved entries from
``core.timetable_pdf.collect_entries`` so the on-screen timetable and the PDF
export render the same data. Sessions that occupy consecutive slots are merged
into a single cell (vertical ``rowspan`` in the classic builder, horizontal
``colspan`` in the transposed one) spanning every covered hour, partial hours
included; on-screen, overlapping sessions are placed on separate lanes within a
day band (``_partition_lanes``). Every render cell keeps the session type so
templates and the PDF can shade Lecture (grey), Workshop (green) and Technical
Drawing (pink).
"""

from core.models import Day

DAY_ORDER = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"]
WEEKEND_ORDER = ["SATURDAY", "SUNDAY"]

# The configured daily timetable window. The on-screen grid trims to the slot
# range that actually has activity; the PDF export requests the full range
# (07:00-08:00 .. 19:00-20:00) via ``build_time_day_grid(..., full_range=True)``.
GRID_HOUR_START = 7
GRID_HOUR_END = 19

# Light "faded" fills used by the PDF and the on-screen grid, so a workshop
# reads as a workshop at a glance. Workshop is a clearly pale green rather than
# a near-white one: at 6.5pt on a white page the previous green-100 was
# indistinguishable from an empty slot, which made workshop blocks look blank.
# Black text stays comfortably readable on every one of these.
# Priority when a block contains several session types:
# TD > Workshop > Lecture > Tutorial > Seminar > Practical.
FILL_COLORS = {
    "td": "#fce7f3",
    "workshop": "#bbf7d0",
    "lecture": "#d1d5db",
    "tutorial": "#dbeafe",
    "seminar": "#ede9fe",
    "practical": "#fef3c7",
}
_FILL_PRIORITY = ["td", "workshop", "lecture", "tutorial", "seminar", "practical"]

def cell_text(entries, show_groups=False):
    """Render one cell's entries as a compact multi-line label.

    Every entry's label appears once. Genuinely distinct simultaneous entries
    are never silently merged: a week range, a venue that differs from the
    workshop name, and (in the all-groups view, ``show_groups``) the owning
    group code are appended so the reader can tell entries apart. Exact
    duplicate text is shown only once.
    """
    def render(e):
        bits = [e["label"]]
        note = e.get("note")
        if note:
            bits.append(str(note))
        if e.get("kind") == "workshop":
            venue = e.get("venue")
            if venue and venue != e.get("name"):
                bits.append(str(venue))
            if show_groups and e.get("groups"):
                bits.append(str(e["groups"]))
        return " · ".join(bits)

    seen = set()
    lines = []
    for same in _by_label(entries).values():
        for e in same:
            text = render(e)
            if text in seen:
                continue
            seen.add(text)
            lines.append(text)
    return "\n\n".join(lines)


def _by_label(entries):
    by_label = {}
    for e in entries:
        by_label.setdefault(e["label"], []).append(e)
    return by_label


# ─────────────────────────────────────────────────────────────
# STRUCTURE OF THE CODE START — ON-SCREEN TIMETABLE MAPPING
# This is the section that maps every session into the on-screen
# timetable: DAYS become rows, hourly TIME slots become columns,
# and a session spanning several hours becomes ONE merged cell
# (`colspan`), with overlapping sessions placed on separate lanes.
#   - `_slot`            → builds the hourly column labels for the PDF/classic grid
#   - `_screen_slot`     → builds the on-screen header labels (07:00-07:55, 08:00-08:55...)
#   - `_partition_lanes` → splits a day's sessions into band sub-rows
#   - `build_day_time_grid` → the entry point (used by core/views.py)
# ─────────────────────────────────────────────────────────────
def _slot(hour):
    return {
        "hour": hour,
        "start": f"{hour:02d}:00",
        "end": f"{hour + 1:02d}:00",
        "label": f"{hour:02d}:00-{hour + 1:02d}:00",
    }


def _screen_slot(hour):
    """Header slot for the on-screen grid: labels read 07:00-07:55 etc.

    Only the rendered header text differs from ``_slot``; ``start``/``end``
    and the ``hour`` value are unchanged so sessions stay mapped to the
    same hourly columns.
    """
    return {
        "hour": hour,
        "start": f"{hour:02d}:00",
        "end": f"{hour + 1:02d}:00",
        "label": f"{hour:02d}:00-{hour:02d}:55",
    }


def build_time_day_grid(entries, full_range=False):
    """Turn ``collect_entries`` output into a time-row/day-column grid.

    Returns a dict with ``slots``, ``days`` (column headers) and ``rows``.
    ``rows`` is one dict per time slot: ``{"slot": {...}, "cols": [...]}``
    aligned to ``days``. Each column cell is ``None`` when the row is already
    consumed by the ``rowspan`` cell above it, ``{"empty": True}`` for a free
    period, or ``{"row", "rowspan", "entries"}`` for a merged session block.

    By default the slot rows are trimmed to the hours actually covered by the
    entries. Passing ``full_range=True`` (used by the PDF export) always
    produces the configured daily window 07:00-08:00 .. 19:00-20:00 so the
    timetable covers the whole day even when few sessions exist; any session
    falling outside that window still adds its own slots rather than being
    clipped.
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

    if full_range:
        slots = sorted(set(range(GRID_HOUR_START, GRID_HOUR_END + 1)) | set(covered))
    else:
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


def _partition_lanes(entries, slots):
    """Split a day's sessions into non-overlapping lanes (band sub-rows).

    Greedy interval partitioning by the hours each entry covers: an entry is
    placed in the first lane whose last session ends before it starts, so a
    new lane is opened whenever sessions overlap. Each returned lane is a
    list of cells aligned to ``slots``: ``{"empty": True, "colspan"}`` for a
    free span of columns, or ``{"colspan", "entries"}`` where the entries
    visually merge across their whole covered range.
    """
    if not entries:
        return [[{"empty": True, "colspan": len(slots)}]]

    lane_runs = []
    lane_cells = []
    for e in sorted(
        entries,
        key=lambda e: (
            min(e["hours"]),
            -(max(e["hours"]) - min(e["hours"])),
            e["course_code"],
            str(e["key"]),
        ),
    ):
        hours = sorted(e["hours"])
        start, end = hours[0], hours[-1]
        start_idx = slots.index(start)
        span = end - start + 1
        lane_i = next(
            (i for i, run in enumerate(lane_runs) if run[-1][1] < start), None
        )
        if lane_i is None:
            lane_i = len(lane_runs)
            lane_runs.append([])
            lane_cells.append([])
        lane_runs[lane_i].append((start, end))
        lane_cells[lane_i].append((start_idx, span, e))

    lanes = []
    for cells in lane_cells:
        cells.sort()
        lane = []
        pos = 0
        for start_idx, span, e in cells:
            if start_idx > pos:
                lane.append({"empty": True, "colspan": start_idx - pos})
            lane.append({"colspan": span, "entries": [e]})
            pos = start_idx + span
        if pos < len(slots):
            lane.append({"empty": True, "colspan": len(slots) - pos})
        lanes.append(lane)
    return lanes


def build_day_time_grid(entries):
    """Transposed weekly grid for the default on-screen timetable view.

    DAYS are rows (always Monday..Friday, plus any weekend days present in
    the data) and the hourly time slots are columns, in ascending order from
    07:00 up to the latest hour actually covered. Each day is drawn as a
    vertical band of one or more sub-rows (lanes). A session spanning several
    hourly slots occupies ONE cell that merges those columns (``colspan``
    derived from every covered hour, including partial hours); sessions that
    overlap on the same day are pushed onto separate lanes instead of being
    sliced into one box per hour column.

    Returns ``{"slots", "days", "rows"}``. ``rows`` is one dict per day:
    ``{"day", "label", "lanes"}`` where ``label`` is the full weekday name,
    and ``lanes`` is the list of band rows from ``_partition_lanes``.
    """
    if not entries:
        return {"slots": [], "days": [], "rows": []}

    present_days = {e["day"] for e in entries}
    day_order = list(DAY_ORDER)
    day_order += [
        d for d in WEEKEND_ORDER if d in present_days and d not in day_order
    ]

    covered = set()
    for e in entries:
        covered.update(e["hours"])
    if not covered:
        return {"slots": [], "days": [], "rows": []}

    first_hour = min(GRID_HOUR_START, min(covered))
    slots = list(range(first_hour, max(covered) + 1))
    slot_defs = [_screen_slot(hour) for hour in slots]

    by_day = {}
    for e in entries:
        if e["hours"]:
            by_day.setdefault(e["day"], []).append(e)

    rows = []
    for day in day_order:
        rows.append(
            {
                "day": day,
                "label": Day(day).label,
                "lanes": _partition_lanes(by_day.get(day, []), slots),
            }
        )

    days = [{"day": row["day"], "label": row["label"]} for row in rows]
    return {"slots": slot_defs, "days": days, "rows": rows}
# ─────────────────────────────────────────────────────────────
# STRUCTURE OF THE CODE END — ON-SCREEN TIMETABLE MAPPING
# ─────────────────────────────────────────────────────────────


def fill_color(entries):
    """Best shading fill for a set of entries (first matching priority kind)."""
    kinds = {e["kind"] for e in entries}
    for kind in _FILL_PRIORITY:
        if kind in kinds:
            return FILL_COLORS[kind]
    return "#ffffff"


def time_day_grid_to_table(grid, show_groups=False):
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
                line.append(cell_text(cell["entries"], show_groups=show_groups))
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