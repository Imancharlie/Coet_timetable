"""Read-only API for the semester the app is currently working in.

Every payload is for that one semester: the portal override when one is set,
otherwise the semester marked current, otherwise the newest semester.
"""

from django.http import HttpResponse, JsonResponse

from core.group_allocation import group_statuses
from core.models import (
    Day,
    Programme,
    Session,
    StudentGroup,
    TechnicalDrawingAllocation,
    Venue,
    WorkshopAllocation,
)
from core.timetable_pdf import collect_group_entries
from student_portal.views import current_semester

from .authentication import require_api_key
from .schema import openapi_schema


def _semester_payload(semester):
    return {
        "id": semester.pk,
        "academic_year": semester.academic_year,
        "semester": semester.semester,
        "label": str(semester),
    }


def _no_semester():
    return JsonResponse({"detail": "No current semester is set."}, status=404)


def _day_label(value):
    try:
        return Day(value).label
    except (ValueError, TypeError):
        return value or ""


def _entry_payload(entry):
    hours = entry.get("hours") or []
    return {
        "kind": entry.get("kind") or "",
        "day": entry.get("day") or "",
        "day_label": _day_label(entry.get("day")),
        "course_code": entry.get("course_code") or "",
        "name": entry.get("name") or "",
        "type": entry.get("type_label") or "",
        "venue": entry.get("venue") or "",
        "start": entry.get("start") or "",
        "end": entry.get("end") or "",
        "groups": entry.get("groups") or "",
        "note": entry.get("note") or "",
        "hours": sorted(hours),
    }


def _day_rank(entry):
    try:
        return list(Day).index(Day(entry["day"]))
    except (ValueError, TypeError):
        return len(Day)


@require_api_key
def group_timetable(request):
    """Sessions, workshops and technical drawing for one group."""
    semester = current_semester()
    if semester is None:
        return _no_semester()
    programme_code = (request.GET.get("programme") or "").strip()
    group_code = (request.GET.get("group") or "").strip()
    if not programme_code or not group_code:
        return JsonResponse(
            {"detail": "Provide programme and group."},
            status=400,
        )
    try:
        year = int(request.GET.get("year") or 1)
    except (TypeError, ValueError):
        return JsonResponse({"detail": "year must be a number."}, status=400)
    year = max(1, min(4, year))
    group = (
        StudentGroup.objects.select_related("programme")
        .filter(
            programme__code__iexact=programme_code,
            code__iexact=group_code,
        )
        .first()
    )
    if group is None:
        return JsonResponse(
            {"detail": "No group matches that programme and group code."},
            status=404,
        )
    entries = collect_group_entries(group, semester, year=year)
    entries = sorted(
        entries, key=lambda entry: (_day_rank(entry), entry.get("start") or "")
    )
    return JsonResponse(
        {
            "semester": _semester_payload(semester),
            "programme": {
                "code": group.programme.code,
                "name": group.programme.name,
            },
            "group": {"id": group.pk, "code": group.code},
            "year": year,
            "count": len(entries),
            "entries": [_entry_payload(entry) for entry in entries],
        }
    )


@require_api_key
def venue_list(request):
    """Rooms named on the current semester's timetable."""
    semester = current_semester()
    if semester is None:
        return _no_semester()
    names = set(
        Session.objects.filter(semester=semester, venue__isnull=False).values_list(
            "venue__name", flat=True
        )
    )
    names.update(
        name
        for name in WorkshopAllocation.objects.filter(semester=semester).values_list(
            "venue", flat=True
        )
        if name
    )
    names.update(
        name
        for name in TechnicalDrawingAllocation.objects.filter(
            semester=semester
        ).values_list("venue", flat=True)
        if name
    )
    venues = list(Venue.objects.filter(name__in=names).order_by("name"))
    return JsonResponse(
        {
            "semester": _semester_payload(semester),
            "count": len(venues),
            "venues": [
                {"id": venue.pk, "name": venue.name, "capacity": venue.capacity}
                for venue in venues
            ],
        }
    )


@require_api_key
def programme_groups(request):
    """Groups whose programme studies courses in the current semester."""
    semester = current_semester()
    if semester is None:
        return _no_semester()
    code = (request.GET.get("programme") or "").strip()
    groups = (
        StudentGroup.objects.filter(
            programme__programme_courses__semester=semester.semester
        )
        .select_related("programme")
        .distinct()
        .order_by("programme__code", "code")
    )
    if code:
        if not Programme.objects.filter(code__iexact=code).exists():
            return JsonResponse(
                {"detail": "No programme matches that code."},
                status=404,
            )
        groups = groups.filter(programme__code__iexact=code)
    programmes = {}
    for group in groups:
        bucket = programmes.setdefault(
            group.programme_id,
            {
                "code": group.programme.code,
                "name": group.programme.name,
                "groups": [],
            },
        )
        bucket["groups"].append({"id": group.pk, "code": group.code})
    listed = list(programmes.values())
    return JsonResponse(
        {
            "semester": _semester_payload(semester),
            "count": sum(len(item["groups"]) for item in listed),
            "programmes": listed,
        }
    )


@require_api_key
def timetable_complete(request):
    """Whether every group that has an allocation to place has finished it."""
    semester = current_semester()
    if semester is None:
        return _no_semester()
    groups = list(
        StudentGroup.objects.filter(
            programme__programme_courses__semester=semester.semester
        )
        .select_related("programme")
        .distinct()
    )
    statuses = group_statuses(semester, "ALL", groups)
    rows = []
    for group in sorted(groups, key=lambda item: (item.programme.code, item.code)):
        status = statuses.get(group.pk)
        met = status.met if status else 0
        total = status.total if status else 0
        rows.append(
            {
                "id": group.pk,
                "programme": group.programme.code,
                "code": group.code,
                "met": met,
                "total": total,
                "percent": status.percent if status else 0,
                "complete": bool(status and status.is_complete),
                "state": status.state if status else "none",
            }
        )
    complete_count = sum(1 for row in rows if row["complete"])
    incomplete_count = sum(1 for row in rows if row["total"] and not row["complete"])
    return JsonResponse(
        {
            "semester": _semester_payload(semester),
            "fully_set": complete_count > 0 and incomplete_count == 0,
            "group_count": len(rows),
            "complete_count": complete_count,
            "incomplete_count": incomplete_count,
            "groups": rows,
        }
    )


def schema(request):
    return JsonResponse(openapi_schema(request))


def docs(request):
    html = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>CoET Timetable API</title>
  <link rel="stylesheet" href="https://unpkg.com/swagger-ui-dist@5/swagger-ui.css">
</head>
<body>
  <div id="swagger-ui"></div>
  <script src="https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js"></script>
  <script>
    SwaggerUIBundle({
      url: "/api/schema/",
      dom_id: "#swagger-ui",
      persistAuthorization: true
    });
  </script>
</body>
</html>"""
    return HttpResponse(html)
