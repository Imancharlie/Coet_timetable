from django import template

from core.models import ALLOCATED_ACTIVITY_TYPES, ActivityType

register = template.Library()


@register.filter
def get_attr(obj, attr):
    """Get an attribute from an object by name. Returns '—' for None."""
    value = getattr(obj, attr, None)
    if value is None:
        return "—"
    return value


@register.filter
def time_short(value):
    """Format a time as HH:MM."""
    if hasattr(value, "strftime"):
        return value.strftime("%H:%M")
    return value


_ACTIVITY_LABELS = {
    "lecture": "Lecture",
    "tutorial": "Tutorial",
    "practical": "Practical",
    "seminar": "Seminar",
    "workshop": "Workshop",
}


@register.filter
def activity_label(value):
    """Map an activity type label to one of the five canonical labels.

    Any variant text from the source data (case, spacing, plurals such as
    "Lectures" or "Workshops") maps to exactly one of Lecture / Tutorial /
    Practical / Seminar / Workshop; unrecognised labels pass through.
    """
    if not value:
        return value
    key = value.strip().lower()
    singular = key.rstrip("s")
    mapping = _ACTIVITY_LABELS.get(key) or _ACTIVITY_LABELS.get(singular)
    return mapping if mapping else value


@register.filter
def get_item(mapping, key):
    """Dictionary lookup by key; returns None when missing."""
    try:
        return mapping.get(key)
    except AttributeError:
        return None


@register.filter
def group_by_activity(rows):
    """Group a plan's assignment dicts by activity, in priority order.

    Returns ``[(label, rows), ...]`` ordered Seminar, Tutorial, Practical, so
    the review table reads in the same order the allocator searched in.
    """
    buckets: dict = {}
    for row in rows or []:
        buckets.setdefault(row.get("activity", ""), []).append(row)
    ordered = []
    for activity in ALLOCATED_ACTIVITY_TYPES:
        label = ActivityType(activity).label
        if label in buckets:
            ordered.append((label, buckets.pop(label)))
    # Anything unexpected still gets shown rather than silently dropped.
    ordered.extend(sorted(buckets.items()))
    return ordered


_NAV_SECTIONS = {
    "programme": "programmes",
    "group": "groups",
    "venue": "venues",
    "semester": "semesters",
    "session": "sessions",
    "workshop": "workshops",
    "td": "td",
    "course": "courses",
    # More specific than "allocation", and it must come first: the loop below
    # returns on the first prefix that matches, so the generic entry would
    # otherwise swallow these and light up the allocator link instead.
    # "audit-reports" likewise precedes "audit-report-*": the list page is the
    # one the sidebar link points at, so that is what must light up, and the
    # longer names are the *pages of* a report rather than the section itself.
    "audit-reports": "audit-reports",
    "audit-report-detail": "audit-reports",
    "audit-report-pdf": "audit-reports",
    "audit-report-json": "audit-reports",
    "allocation-groups": "allocation-progress",
    "allocation-group": "allocation-progress",
    "allocation-smart-page": "allocation-smart",
    "allocation-smart-preview": "allocation-smart",
    "allocation-advanced-page": "allocation-advanced",
    "allocation-advanced-preview": "allocation-advanced",
    "allocation": "allocation",
    "import": "imports",
    "export": "exports",
    "timetable": "timetable",
    "activity": "activity",
    "danger": "danger",
}


@register.simple_tag(takes_context=True)
def nav_section(context):
    """Return the active sidebar section derived from the URL name."""
    request = context.get("request")
    if request is None:
        return ""
    name = getattr(getattr(request, "resolver_match", None), "url_name", "") or ""
    if name == "dashboard":
        return "dashboard"
    for prefix, section in _NAV_SECTIONS.items():
        if name == prefix or name.startswith(prefix + "-"):
            return section
    return ""


@register.filter
def nav_is_any(nav_section_value, sections):
    """True when the active section is one of a comma-separated list.

    Used to light up the header of a collapsible section when one of the links
    inside it is the current page. The header deliberately does *not* get the
    same ``bg-slate-800 text-white`` the child links use, so a single sidebar
    item is never highlighted twice.
    """
    wanted = {s.strip() for s in str(sections).split(",") if s.strip()}
    return bool(nav_section_value) and nav_section_value in wanted


@register.filter
def active_cls(nav_section_value, section):
    """Sidebar link classes for the given section if it is the active one."""
    if nav_section_value == section:
        return "bg-slate-800 text-white"
    return "hover:bg-slate-800/60 text-slate-300"


@register.filter
def log_badge_cls(action):
    """Tailwind badge colour for a log action."""
    return {
        "CREATE": "bg-emerald-500",
        "UPDATE": "bg-blue-500",
        "DELETE": "bg-rose-500",
        "CLEAR": "bg-rose-700",
        "IMPORT": "bg-indigo-500",
        "ASSIGN": "bg-purple-500",
        "REMOVE": "bg-amber-500",
    }.get(action, "bg-slate-500")
