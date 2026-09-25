from django import template

from core.models import ActivityLog

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


_NAV_SECTIONS = {
    "programme": "programmes",
    "group": "groups",
    "venue": "venues",
    "semester": "semesters",
    "session": "sessions",
    "workshop": "workshops",
    "td": "td",
    "course": "courses",
    "import": "imports",
    "export": "exports",
    "timetable": "timetable",
    "activity": "activity",
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
def active_cls(nav_section_value, section):
    """Sidebar link classes for the given section if it is the active one."""
    if nav_section_value == section:
        return "bg-slate-800 text-white"
    return "hover:bg-slate-800/60 text-slate-300"


@register.simple_tag
def activity_logs(limit=10):
    """Latest activity log entries for the sidebar."""
    try:
        return list(ActivityLog.objects.all()[: int(limit)])
    except (TypeError, ValueError):
        return list(ActivityLog.objects.all()[:10])


@register.filter
def log_badge_cls(action):
    """Tailwind badge colour for a log action."""
    return {
        "CREATE": "bg-emerald-500",
        "UPDATE": "bg-blue-500",
        "DELETE": "bg-rose-500",
        "IMPORT": "bg-indigo-500",
        "ASSIGN": "bg-purple-500",
        "REMOVE": "bg-amber-500",
    }.get(action, "bg-slate-500")
