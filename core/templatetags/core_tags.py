from django import template

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
