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
