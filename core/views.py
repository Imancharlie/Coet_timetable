import io
import json
import uuid
from pathlib import Path
from urllib.parse import urlencode

from django.db.models import Q, Count
from django.http import HttpResponse, HttpResponseBadRequest, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone

from .timetable_grid import build_day_time_grid
from .timetable_pdf import (
    collect_entries,
    collect_group_entries,
    collect_master_entries,
    render_group_timetable,
    render_programme_timetable,
)

from .deletion_impact import deletion_impact
from .forms import (
    FileUploadForm,
    ProgrammeCourseForm,
    ProgrammeForm,
    SemesterForm,
    SessionForm,
    SessionGroupFormSet,
    StudentGroupForm,
    TechnicalDrawingAllocationForm,
    VenueForm,
    VenueRecycleForm,
    WorkshopAllocationForm,
)
from .importers import (
    ImportResult,
    assign_lecture_groups,
    import_master_timetable_from_excel,
    import_programme_courses_from_excel,
    import_programmes_from_excel,
    import_student_groups_from_excel,
    import_td_allocation_from_excel,
    import_venues_from_excel,
    import_workshop_allocation_from_excel,
)
from .models import (
    ActivityLog,
    Day,
    ImportHistory,
    ImportStatus,
    LogAction,
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    Venue,
    WorkshopAllocation,
)
from .venue_quality import (
    analyse_venues,
    base_key,
    detect_name_conflicts,
    issues_for,
    merge_venues,
    resolve_venue_name_conflict,
    suggested_name,
)
from .workshop_times import legacy_workshop_allocations, legacy_workshop_sessions

HTMX_HEADER = "HX-Request"


def _htmx(request):
    return request.headers.get(HTMX_HEADER)


def _log(action, message, resource="", target=""):
    """Record an entry in the activity log."""
    ActivityLog.objects.create(
        action=action, message=message, resource=resource, target=target
    )


def _write_response(request, trigger, redirect_name, *args):
    """Success response that works with and without htmx.

    For htmx requests return the trigger header the UI reacts to; for a plain
    browser POST fall back to a normal redirect so saves never appear to hang.
    """
    if _htmx(request):
        r = HttpResponse("")
        r["HX-Trigger"] = trigger
        return r
    return redirect(redirect_name, *args)


def _delete_context(request, item, back_url, delete_url):
    """Shared context for the confirm-delete page, including the impact preview."""
    return {
        "item": item,
        "title": f"Delete {item}?",
        "back_url": back_url,
        "delete_url": delete_url,
        "impact": deletion_impact(item),
    }


CLEAR_ALL_PHRASE = "DELETE ALL"


def _clear_all_context(
    request,
    list_url,
    clear_url,
    page_title,
    primary_label,
    primary_count,
    related,
    detached,
    kept_note="",
):
    """Shared context for the clear-all confirmation modal."""
    return {
        "clear_url": clear_url,
        "back_url": list_url,
        "page_title": page_title,
        "primary_label": primary_label,
        "primary_count": primary_count,
        "related": related,
        "detached": detached,
        "detached_total": sum(group["count"] for group in detached),
        "kept_note": kept_note,
        "filters_active": bool(request.GET),
        "has_records": bool(
            primary_count
            or any(group["count"] for group in related)
            or any(group["count"] for group in detached)
        ),
    }


def _clear_all(
    request,
    *,
    model,
    list_url,
    clear_url,
    page_title,
    primary_label,
    log_resource,
    redirect_name,
    related_count=(),
    detached_count=(),
    kept_note="",
):
    """Confirm + execute clearing every record of one resource type.

    ``related_count`` / ``detached_count`` are sequences of
    ``(queryset, label)`` pairs. ``related_count`` rows cascade away with the
    primary records (e.g. SessionGroup links for Session); ``detached_count``
    rows are kept with their reference cleared (e.g. Sessions when clearing
    Venues via the SET_NULL FK). Deleting happens through the ORM so Django
    honours the model CASCADE/SET_NULL rules; unrelated reference data is
    never touched.
    """
    ctx = _clear_all_context(
        request,
        list_url,
        clear_url,
        page_title,
        primary_label,
        model.objects.count(),
        [{"label": label, "count": qs.count()} for qs, label in related_count],
        [{"label": label, "count": qs.count()} for qs, label in detached_count],
        kept_note,
    )
    if request.method != "POST":
        return render(request, "core/clear_all.html", ctx)

    if request.POST.get("phrase", "").strip().upper() != CLEAR_ALL_PHRASE:
        ctx["error"] = (
            f"Type {CLEAR_ALL_PHRASE} to confirm — nothing was deleted."
        )
        return render(request, "core/clear_all.html", ctx)

    primary = model.objects.count()
    related = [(qs.count(), label) for qs, label in related_count]
    detached = [(qs.count(), label) for qs, label in detached_count]
    try:
        model.objects.all().delete()
    except Exception as exc:  # pragma: no cover - defensive
        ctx["error"] = (
            f"Clearing failed: {exc}. The deletion was rolled back — "
            "no records have been removed."
        )
        return render(request, "core/clear_all.html", ctx)
    if primary or any(count for count, _ in related) or any(
        count for count, _ in detached
    ):
        parts = [f"{primary} {log_resource} record(s)"]
        parts += [f"{count} {label}" for count, label in related if count]
        det_parts = [f"{count} {label}" for count, label in detached if count]
        message = f"Cleared all {primary_label}: "
        if parts:
            message += ", ".join(parts) + " removed"
        if det_parts:
            message += "; " if parts else ""
            message += ", ".join(det_parts) + " kept with reference cleared"
        _log(LogAction.CLEAR, message, log_resource)
    return _write_response(request, "close-modal,refresh-table", redirect_name)


def _search(qs, q, fields):
    if not q:
        return qs
    q_obj = Q()
    for f in fields:
        q_obj |= Q(**{f + "__icontains": q})
    return qs.filter(q_obj)


def _filters(request, specs):
    """Build filter context entries (name/label/type/options + active value)."""
    return [
        {**spec, "value": request.GET.get(spec["name"], "")}
        for spec in specs
    ]


def _query(request, param_names):
    """Build a ?... query string from the active GET params only."""
    pairs = [
        (p, request.GET.get(p, "")) for p in param_names if request.GET.get(p, "")
    ]
    return "?" + urlencode(pairs) if pairs else ""


def _paginate(request, qs, per_page=50):
    page = int(request.GET.get("page", 1))
    total = qs.count()
    total_pages = max((total + per_page - 1) // per_page, 1)
    page = min(page, total_pages)
    start = (page - 1) * per_page
    return qs[start : start + per_page], page, total_pages, total


# ──────────────────────────────────────────────
# Dashboard
# ──────────────────────────────────────────────


def dashboard(request):
    ctx = {
        "programme_count": Programme.objects.count(),
        "group_count": StudentGroup.objects.count(),
        "venue_count": Venue.objects.count(),
        "session_count": Session.objects.count(),
        "sessiongroup_count": SessionGroup.objects.count(),
        "workshop_count": WorkshopAllocation.objects.count(),
        "td_count": TechnicalDrawingAllocation.objects.count(),
        "course_count": ProgrammeCourse.objects.count(),
        "semester_count": Semester.objects.count(),
        "recent_sessions": Session.objects.select_related(
            "semester", "venue"
        ).order_by("-pk")[:10],
    }
    log_id = request.GET.get("log", "")
    if log_id:
        try:
            ctx["selected_log"] = ActivityLog.objects.get(pk=int(log_id))
        except (ValueError, ActivityLog.DoesNotExist):
            pass
    return render(request, "dashboard.html", ctx)


def activity_list(request):
    qs = ActivityLog.objects.all()
    q = request.GET.get("q", "")
    act = request.GET.get("action", "")
    qs = _search(qs, q, ["resource", "target", "message"])
    if act:
        qs = qs.filter(action=act)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "q": q,
        "filters": _filters(
            request,
            [
                {
                    "name": "action",
                    "label": "Action",
                    "type": "select",
                    "options": LogAction.choices,
                },
            ],
        ),
        "query": _query(request, ["q", "action"]),
        "page_title": "Activity Log",
        "list_url": "/activity/",
        "page": page,
        "pages": pages,
        "total": total,
    }
    if _htmx(request):
        return render(request, "core/_activity_log.html", ctx)
    return render(request, "core/activity_log.html", ctx)


# ──────────────────────────────────────────────
# Programme
# ──────────────────────────────────────────────

PROG_COLS = [{"key": "code", "label": "Code"}, {"key": "name", "label": "Name"}]
PROG_FIELDS = [{"label": "Code", "key": "code"}, {"label": "Name", "key": "name"}]


def programme_list(request):
    qs = Programme.objects.all()
    q = request.GET.get("q", "")
    qs = _search(qs, q, ["code", "name"])
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": PROG_COLS,
        "detail_fields": PROG_FIELDS,
        "q": q,
        "filters": _filters(request, []),
        "query": _query(request, ["q"]),
        "page_title": "Programmes",
        "list_url": "/programmes/",
        "create_url": "/programmes/create/",
        "clear_all_url": "/programmes/clear-all/",
        "edit_name": "programme-edit",
        "delete_name": "programme-delete",
        "detail_name": "programme-detail",
        "page": page,
        "pages": pages,
        "total": total,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/list.html", ctx)


def programme_detail(request, pk):
    item = get_object_or_404(Programme, pk=pk)
    ctx = {
        "item": item,
        "detail_fields": PROG_FIELDS,
        "page_title": str(item),
        "edit_url": f"/programmes/{pk}/edit/",
        "delete_url": f"/programmes/{pk}/delete/",
        "back_url": "/programmes/",
        "export_url": reverse("programme-timetable-export", args=[pk]),
        "export_label": "Export Timetable (PDF)",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def programme_create(request):
    if request.method == "POST":
        form = ProgrammeForm(request.POST)
        if form.is_valid():
            obj = form.save()
            _log(LogAction.CREATE, f"Created Programme {obj}", "Programme", str(obj))
            return _write_response(
                request, "close-modal,refresh-table", "programme-list"
            )
    else:
        form = ProgrammeForm()
    return render(
        request,
        "core/form.html",
        {"form": form, "title": "Create Programme", "action": "/programmes/create/"},
    )


def programme_edit(request, pk):
    item = get_object_or_404(Programme, pk=pk)
    if request.method == "POST":
        form = ProgrammeForm(request.POST, instance=item)
        if form.is_valid():
            obj = form.save()
            _log(LogAction.UPDATE, f"Updated Programme {obj}", "Programme", str(obj))
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "programme-list",
            )
    else:
        form = ProgrammeForm(instance=item)
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": f"Edit {item}",
            "action": f"/programmes/{pk}/edit/",
        },
    )


def programme_delete(request, pk):
    item = get_object_or_404(Programme, pk=pk)
    if request.method == "POST":
        label = str(item)
        item.delete()
        _log(LogAction.DELETE, f"Deleted Programme {label}", "Programme", label)
        return _write_response(request, "close-modal,refresh-table", "programme-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/programmes/", f"/programmes/{pk}/delete/"),
    )


def programme_clear_all(request):
    """Clear every Programme (cascading its courses and student groups)."""
    return _clear_all(
        request,
        model=Programme,
        list_url="/programmes/",
        clear_url="/programmes/clear-all/",
        page_title="Programmes",
        primary_label="Programmes",
        log_resource="Programme",
        redirect_name="programme-list",
        related_count=(
            (StudentGroup.objects.all(), "Student groups"),
            (ProgrammeCourse.objects.all(), "Programme courses"),
            (SessionGroup.objects.all(), "Session-group links"),
        ),
        kept_note=(
            "Linked student groups and programme courses are deleted with their "
            "programmes. Sessions and semesters are kept."
        ),
    )


def _latest_semester_with_data():
    """Most recent semester that holds any timetable data, else the latest one."""
    return (
        Semester.objects.filter(
            Q(sessions__isnull=False)
            | Q(workshop_allocations__isnull=False)
            | Q(td_allocations__isnull=False)
        )
        .order_by("-academic_year", "-semester")
        .distinct()
        .first()
        or Semester.objects.order_by("-academic_year", "-semester").first()
    )


def export_timetable(request):
    """Export hub: search programmes/groups, then export per group or programme."""
    q = request.GET.get("q", "").strip()
    programmes = Programme.objects.all()
    if q:
        programmes = programmes.filter(
            Q(code__icontains=q)
            | Q(name__icontains=q)
            | Q(student_groups__code__icontains=q)
        ).distinct()
    semester = _latest_semester_with_data()
    rows = []
    for p in programmes:
        groups = p.student_groups.all()
        if q:
            groups = [
                g
                for g in groups
                if q.lower() in g.code.lower()
                or q.lower() in p.code.lower()
                or q.lower() in p.name.lower()
            ]
        rows.append({"programme": p, "groups": groups})
    ctx = {
        "page_title": "Export Timetable",
        "rows": rows,
        "q": q,
        "total_groups": StudentGroup.objects.count(),
        "semesters": Semester.objects.all(),
        "default_semester": semester,
        "year_options": [1, 2, 3, 4],
    }
    return render(request, "core/export_timetable.html", ctx)


def programme_timetable_pdf(request, pk):
    """Export a single programme's timetable as a PDF grid (merged cells)."""
    programme = get_object_or_404(Programme, pk=pk)
    sem_id = request.GET.get("semester", "")
    semester = None
    if sem_id:
        semester = get_object_or_404(Semester, pk=sem_id)
    else:
        group_codes = list(
            StudentGroup.objects.filter(programme=programme).values_list(
                "code", flat=True
            )
        )
        semester = (
            Semester.objects.filter(
                Q(sessions__session_groups__group__programme=programme)
                | Q(workshop_allocations__group_code__in=group_codes)
                | Q(td_allocations__group_code__in=group_codes)
            )
            .order_by("-academic_year", "-semester")
            .distinct()
            .first()
            or _latest_semester_with_data()
        )
    try:
        year = int(request.GET.get("year", 1))
    except (TypeError, ValueError):
        year = 1
    response = HttpResponse(content_type="application/pdf")
    disposition = (
        f'attachment; filename="timetable_{programme.code}_{year}.pdf"'
    )
    response["Content-Disposition"] = disposition
    render_programme_timetable(programme, semester, year, out=response)
    return response


def group_timetable_pdf(request, pk):
    """Export a single student group's timetable as a PDF grid (merged cells)."""
    group = get_object_or_404(
        StudentGroup.objects.select_related("programme"), pk=pk
    )
    sem_id = request.GET.get("semester", "")
    semester = None
    if sem_id:
        semester = get_object_or_404(Semester, pk=sem_id)
    else:
        semester = (
            Semester.objects.filter(
                Q(sessions__session_groups__group=group)
                | Q(workshop_allocations__group_code=group.code)
                | Q(td_allocations__group_code=group.code)
            )
            .order_by("-academic_year", "-semester")
            .distinct()
            .first()
            or _latest_semester_with_data()
        )
    try:
        year = int(request.GET.get("year", 1))
    except (TypeError, ValueError):
        year = 1
    response = HttpResponse(content_type="application/pdf")
    filename = f"timetable_{group.programme.code}_{group.code}_{year}.pdf"
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    render_group_timetable(group, semester, year, out=response)
    return response


def timetable_view(request):
    """Screen timetable grid: TIME slots across the first row, DAYS down the first column.

    Built from the session/workshop/TD records of the selected programme and
    semester via ``core.timetable_grid.build_day_time_grid`` (same data source
    as the PDF export), so any programme renders correctly. Days run down the
    first column with full weekday labels (Monday..Friday, plus any weekend
    days with data), and each day is drawn as a vertical band of lanes:
    sessions are rendered as one Activity Card per cell spanning their hourly
    columns, and overlapping sessions sit on separate lanes.
    """
    programmes = list(Programme.objects.all())
    semesters = list(Semester.objects.order_by("-academic_year", "-semester"))

    active_programme = None
    programme_id = request.GET.get("programme")
    if programme_id and programme_id != "all":
        active_programme = Programme.objects.filter(pk=programme_id).first()
    # Empty/missing selection (or "all") means the whole COET First Year.
    show_all = active_programme is None

    active_group = None
    group_id = request.GET.get("group")
    if group_id and active_programme:
        active_group = StudentGroup.objects.filter(pk=group_id, programme=active_programme).first()

    active_semester = None
    if request.GET.get("semester"):
        active_semester = Semester.objects.filter(
            pk=request.GET["semester"]
        ).first()
    if active_semester is None:
        active_semester = _latest_semester_with_data()

    year = 1
    raw_year = request.GET.get("year", "")
    if raw_year.strip().isdigit():
        year = max(1, min(4, int(raw_year)))

    entries = []
    grid = {"slots": [], "days": [], "rows": []}
    if active_semester:
        if active_group:
            entries = collect_group_entries(active_group, active_semester, year=year)
        elif active_programme:
            entries = collect_entries(active_programme, active_semester, year=year)
        elif show_all:
            entries = collect_master_entries(active_semester, year=year)
        grid = build_day_time_grid(entries)

    ordinal = {1: "1st", 2: "2nd", 3: "3rd", 4: "4th"}
    
    groups_for_programme = []
    if active_programme:
        groups_for_programme = list(StudentGroup.objects.filter(programme=active_programme))
    
    ctx = {
        "page_title": "Timetable",
        "programmes": programmes,
        "semesters": semesters,
        "active_programme": active_programme,
        "active_group": active_group,
        "groups_for_programme": groups_for_programme,
        "active_semester": active_semester,
        "year": year,
        "year_options": [1, 2, 3, 4],
        "year_ordinal": ordinal.get(year, f"{year}th"),
        "grid": grid,
        "entry_count": len(entries),
        "show_all": show_all,
        "show_groups": (active_programme is not None and active_group is None) or show_all,
    }
    return render(request, "core/timetable.html", ctx)


def timetable_groups_json(request):
    """JSON endpoint to fetch groups for a selected programme."""
    programme_id = request.GET.get("programme")
    groups = []
    if programme_id:
        groups = list(
            StudentGroup.objects.filter(programme_id=programme).values("id", "code")
        )
    return JsonResponse({"groups": groups})


# ──────────────────────────────────────────────
# Student Group
# ──────────────────────────────────────────────

GRP_COLS = [
    {"key": "programme", "label": "Programme"},
    {"key": "code", "label": "Group Code"},
]
GRP_FIELDS = [
    {"label": "Programme", "key": "programme"},
    {"label": "Group Code", "key": "code"},
]


def studentgroup_list(request):
    qs = StudentGroup.objects.select_related("programme").all()
    q = request.GET.get("q", "")
    prog = request.GET.get("programme", "")
    qs = _search(qs, q, ["code", "programme__code", "programme__name"])
    if prog:
        qs = qs.filter(programme__code=prog)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": GRP_COLS,
        "detail_fields": GRP_FIELDS,
        "q": q,
        "filters": _filters(
            request,
            [
                {
                    "name": "programme",
                    "label": "Programme",
                    "type": "select",
                    "options": [
                        (p.code, p.name) for p in Programme.objects.all()
                    ],
                },
            ],
        ),
        "query": _query(request, ["q", "programme"]),
        "page_title": "Student Groups",
        "list_url": "/groups/",
        "create_url": "/groups/create/",
        "clear_all_url": "/groups/clear-all/",
        "edit_name": "group-edit",
        "delete_name": "group-delete",
        "detail_name": "group-detail",
        "page": page,
        "pages": pages,
        "total": total,
        "programmes": Programme.objects.all(),
        "active_programme": prog,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/list.html", ctx)


def studentgroup_detail(request, pk):
    item = get_object_or_404(StudentGroup.objects.select_related("programme"), pk=pk)
    ctx = {
        "item": item,
        "detail_fields": GRP_FIELDS,
        "page_title": str(item),
        "edit_url": f"/groups/{pk}/edit/",
        "delete_url": f"/groups/{pk}/delete/",
        "back_url": "/groups/",
        "export_url": reverse("group-timetable-export", args=[pk]),
        "export_label": "Export Timetable (PDF)",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def studentgroup_create(request):
    if request.method == "POST":
        form = StudentGroupForm(request.POST)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.CREATE,
                f"Created Student Group {obj}",
                "Student Group",
                str(obj),
            )
            return _write_response(
                request, "close-modal,refresh-table", "group-list"
            )
    else:
        form = StudentGroupForm()
    return render(
        request,
        "core/form.html",
        {"form": form, "title": "Create Student Group", "action": "/groups/create/"},
    )


def studentgroup_edit(request, pk):
    item = get_object_or_404(StudentGroup, pk=pk)
    if request.method == "POST":
        form = StudentGroupForm(request.POST, instance=item)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.UPDATE,
                f"Updated Student Group {obj}",
                "Student Group",
                str(obj),
            )
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "group-list",
            )
    else:
        form = StudentGroupForm(instance=item)
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": f"Edit {item}",
            "action": f"/groups/{pk}/edit/",
        },
    )


def studentgroup_delete(request, pk):
    item = get_object_or_404(StudentGroup, pk=pk)
    if request.method == "POST":
        label = str(item)
        item.delete()
        _log(
            LogAction.DELETE, f"Deleted Student Group {label}", "Student Group", label
        )
        return _write_response(request, "close-modal,refresh-table", "group-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/groups/", f"/groups/{pk}/delete/"),
    )


def group_clear_all(request):
    """Clear every StudentGroup, removing its SessionGroup links."""
    return _clear_all(
        request,
        model=StudentGroup,
        list_url="/groups/",
        clear_url="/groups/clear-all/",
        page_title="Student Groups",
        primary_label="Student Groups",
        log_resource="Student Group",
        redirect_name="group-list",
        related_count=((SessionGroup.objects.all(), "Session-group links"),),
        kept_note="Sessions, programmes, courses, semesters and venues are kept.",
    )


# ──────────────────────────────────────────────
# Venue
# ──────────────────────────────────────────────

VENUE_COLS = [
    {"key": "name", "label": "Name"},
    {"key": "capacity", "label": "Capacity"},
]
VENUE_FIELDS = [
    {"label": "Name", "key": "name"},
    {"label": "Capacity", "key": "capacity"},
]


def venue_list(request):
    qs = Venue.objects.all()
    q = request.GET.get("q", "")
    cmin = request.GET.get("capacity_min", "")
    cmax = request.GET.get("capacity_max", "")
    qs = _search(qs, q, ["name"])
    if cmin:
        qs = qs.filter(capacity__gte=cmin)
    if cmax:
        qs = qs.filter(capacity__lte=cmax)
    items, page, pages, total = _paginate(request, qs)
    issues_map, problem_groups, has_issues = analyse_venues()
    ctx = {
        "items": items,
        "columns": VENUE_COLS,
        "detail_fields": VENUE_FIELDS,
        "q": q,
        "filters": _filters(
            request,
            [
                {"name": "capacity_min", "label": "Capacity Min", "type": "number"},
                {"name": "capacity_max", "label": "Capacity Max", "type": "number"},
            ],
        ),
        "query": _query(request, ["q", "capacity_min", "capacity_max"]),
        "page_title": "Venues",
        "list_url": "/venues/",
        "create_url": "/venues/create/",
        "clear_all_url": "/venues/clear-all/",
        "edit_name": "venue-edit",
        "delete_name": "venue-delete",
        "detail_name": "venue-detail",
        "page": page,
        "pages": pages,
        "total": total,
        "row_issues": issues_map,
        "recycle_url": "/venues/recycle/",
        "recycle_count": len(issues_map),
        "has_issues": has_issues,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/venue_list.html", ctx)


def venue_detail(request, pk):
    item = get_object_or_404(Venue, pk=pk)
    ctx = {
        "item": item,
        "detail_fields": VENUE_FIELDS,
        "page_title": str(item),
        "edit_url": f"/venues/{pk}/edit/",
        "delete_url": f"/venues/{pk}/delete/",
        "back_url": "/venues/",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def venue_create(request):
    if request.method == "POST":
        form = VenueForm(request.POST)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.CREATE,
                f"Created Venue {obj.name} (capacity {obj.capacity})",
                "Venue",
                obj.name,
            )
            return _write_response(
                request, "close-modal,refresh-table", "venue-list"
            )
    else:
        form = VenueForm()
    return render(
        request,
        "core/form.html",
        {"form": form, "title": "Create Venue", "action": "/venues/create/"},
    )


def venue_edit(request, pk):
    item = get_object_or_404(Venue, pk=pk)
    if request.method == "POST":
        form = VenueForm(request.POST, instance=item)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.UPDATE,
                f"Updated Venue {obj.name} (capacity {obj.capacity})",
                "Venue",
                obj.name,
            )
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "venue-list",
            )
    else:
        form = VenueForm(instance=item)
    return render(
        request,
        "core/form.html",
        {"form": form, "title": f"Edit {item}", "action": f"/venues/{pk}/edit/"},
    )


def venue_delete(request, pk):
    item = get_object_or_404(Venue, pk=pk)
    if request.method == "POST":
        label = item.name
        item.delete()
        _log(LogAction.DELETE, f"Deleted Venue {label}", "Venue", label)
        return _write_response(request, "close-modal,refresh-table", "venue-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/venues/", f"/venues/{pk}/delete/"),
    )


def venue_clear_all(request):
    """Clear every Venue, clearing its Session references (sessions are kept)."""
    return _clear_all(
        request,
        model=Venue,
        list_url="/venues/",
        clear_url="/venues/clear-all/",
        page_title="Venues",
        primary_label="Venues",
        log_resource="Venue",
        redirect_name="venue-list",
        detached_count=((Session.objects.filter(venue__isnull=False), "sessions"),),
        kept_note=(
            "Sessions are kept and will simply have their venue reference "
            "cleared. Semesters, programmes, courses and student groups are kept."
        ),
    )


def _merge_venues(source, target):
    """Fold one venue into another and keep every reference in sync."""
    label = source.name
    target = merge_venues(source, target)
    _log(
        LogAction.UPDATE,
        f"Merged Venue {label} into {target.name}",
        "Venue",
        target.name,
    )


def _partner_for(venue):
    """Preferred venue to keep when folding a duplicate away (biggest capacity,
    then longest name, then lowest pk for a deterministic choice)."""
    key = base_key(venue.name)
    allies = [
        v
        for v in Venue.objects.exclude(pk=venue.pk)
        if base_key(v.name) == key
    ]
    if not allies:
        return None
    return max(allies, key=lambda v: (v.capacity or 0, len(v.name), -v.pk))


def _venue_impact(venue):
    """Preview for removing a venue on the recycle workbench.

    Pulling a duplicate venue away re-points every reference to the kept
    partner (merge) before deleting the row; a lone venue with no partner is
    plain-deleted and its Session FKs fall back to blank. Returns counts and a
    short human summary for each row in the recycle list.
    """
    sessions = Session.objects.filter(venue=venue).count()
    workshops = WorkshopAllocation.objects.filter(venue=venue.name).count()
    tds = TechnicalDrawingAllocation.objects.filter(venue=venue.name).count()
    partner = _partner_for(venue)
    total = sessions + workshops + tds
    if partner:
        summary = (
            f"No records lost — {total} reference(s) re-pointed to "
            f"'{partner.name}' before this row is removed."
        )
    else:
        summary = (
            f"No partner to merge into; {sessions} session(s) keep existing "
            f"but their venue is cleared."
        )
    return {
        "sessions": sessions,
        "workshops": workshops,
        "td": tds,
        "total": total,
        "partner": partner.name if partner else "",
        "summary": summary,
    }


def venue_recycle(request):
    """Review and clean up venues that were imported with casing, spacing or
    duplicate inconsistencies.

    GET renders the workbench; a POST with action=edit fixes a single venue
    (merging into any clash by normalised name) and action=delete removes a
    duplicate after its references have been re-pointed.  Every change is
    written straight to the database.
    """
    flash_success = []
    flash_errors = []
    if request.method == "POST":
        action = request.POST.get("action")
        try:
            pk = int(request.POST.get("pk", ""))
        except (TypeError, ValueError):
            pk = None
        if pk is None or action not in ("edit", "delete"):
            flash_errors.append("Invalid recycle request.")
        elif action == "edit":
            venue = get_object_or_404(Venue, pk=pk)
            new_name = request.POST.get("name", "").strip()
            raw_capacity = request.POST.get("capacity", "")
            clash = (
                Venue.objects.exclude(pk=pk).filter(name__iexact=new_name).first()
            )
            if clash:
                old_label = venue.name
                _merge_venues(venue, clash)
                flash_success.append(
                    f"Merged '{old_label}' into '{clash.name}'. Sessions and "
                    f"allocations now point to '{clash.name}'."
                )
            else:
                payload = {"name": new_name}
                if raw_capacity:
                    payload["capacity"] = raw_capacity
                # Resolving a casing/spacing duplicate is an independent action:
                # it only validates the name and never touches capacity unless a
                # new value was actually supplied, so existing capacities are
                # preserved and missing capacity never blocks the fix.
                form = (
                    VenueForm(payload, instance=venue)
                    if raw_capacity
                    else VenueRecycleForm(payload, instance=venue)
                )
                if form.is_valid():
                    form.save()
                    _log(
                        LogAction.UPDATE,
                        f"Venue recycled to '{new_name}'",
                        "Venue",
                        new_name,
                    )
                    flash_success.append(
                        f"Saved '{new_name}'. Database updated."
                    )
                else:
                    flash_errors.extend(
                        f"{field}: {' '.join(errs)}"
                        for field, errs in form.errors.items()
                    )
        else:
            venue = get_object_or_404(Venue, pk=pk)
            label = venue.name
            partner = _partner_for(venue)
            if partner:
                _merge_venues(venue, partner)
                flash_success.append(
                    f"Removed duplicate '{label}'; kept '{partner.name}'. "
                    f"Database updated."
                )
            else:
                venue.delete()
                _log(LogAction.DELETE, f"Deleted Venue {label}", "Venue", label)
                flash_success.append(f"Deleted '{label}'. Database updated.")

    issues_map, dup_groups, has_issues = analyse_venues()
    dup_pks = {v.pk for g in dup_groups for v in g["venues"]}
    fmt_only = [
        v
        for v in Venue.objects.all()
        if v.pk in issues_map and v.pk not in dup_pks
    ]
    venue_impacts = {
        v.pk: _venue_impact(v)
        for v in Venue.objects.filter(pk__in=set(issues_map) | set(dup_pks))
    }
    ctx = {
        "issues_map": issues_map,
        "dup_groups": dup_groups,
        "fmt_only": fmt_only,
        "has_issues": has_issues,
        "recycle_count": len(issues_map),
        "flash_success": flash_success,
        "flash_errors": flash_errors,
        "page_title": "Recycle Venue Data",
        "list_url": "/venues/",
        "recycle_url": "/venues/recycle/",
        "create_url": "/venues/create/",
        "venue_impacts": venue_impacts,
    }
    return render(request, "core/venue_recycle.html", ctx)


def venue_fix_conflict(request):
    """Resolve one interactive venue-name conflict from an import result.

    POSTs carry the import collision token, the normalised venue key, the
    issue id and the user-chosen official name. Every venue collapsing to that
    key is merged into the accepted name, an ActivityLog entry records the
    choice (original value, selected official value, action, time) and the
    conflicts panel is re-rendered so the remaining count updates in place.
    """
    token = request.POST.get("token", "").strip()
    issue_id = request.POST.get("issue_id", "").strip()
    key = request.POST.get("key", "").strip()
    official = request.POST.get("official", "").strip()
    flash_success = []
    flash_errors = []
    store_key = f"venue_conflicts:{token}" if token else ""
    store = request.session.get(store_key, {"open": [], "resolved": []}) if store_key else {"open": [], "resolved": []}

    issue = None
    if request.method == "POST":
        if not key or not official:
            flash_errors.append("Missing venue choice — nothing changed.")
        else:
            issue = next(
                (i for i in store.get("open", []) if i.get("id") == issue_id),
                None,
            )
            offered = set(issue["names"]) if issue else set()
            if issue and official not in offered:
                flash_errors.append(
                    f"'{official}' was not one of the offered names for this issue."
                )
            else:
                _, merged = resolve_venue_name_conflict(key, official)
                merged_text = ", ".join(merged)
                flash_success.append(
                    f"Accepted '{official}' as the official name"
                    + (f" — merged '{merged_text}'." if merged else ".")
                )
                _log(
                    LogAction.UPDATE,
                    f"Venue import conflict fixed: original '{merged_text or official}'"
                    f" → official '{official}' accepted",
                    "Venue",
                    official,
                )
                store["open"] = [
                    i
                    for i in store.get("open", [])
                    if i.get("id") != issue_id
                ]
                store["resolved"].append(
                    {
                        "id": issue_id,
                        "key": key,
                        "names": list(offered) or [official],
                        "reasons": issue.get("reasons", ["Casing", "Spacing"])
                        if issue
                        else [],
                        "accepted": official,
                        "fixed_at": timezone.now().isoformat(),
                    }
                )

    if store_key:
        current_names = set(Venue.objects.values_list("name", flat=True))
        store["open"] = [
            i
            for i in store.get("open", [])
            if set(i["names"]).issubset(current_names)
        ]
        request.session[store_key] = store

    ctx = {
        "open_conflicts": store.get("open", []),
        "resolved_conflicts": store.get("resolved", []),
        "token": token,
        "flash_success": flash_success,
        "flash_errors": flash_errors,
    }
    return render(request, "core/_venue_conflicts.html", ctx)


# ──────────────────────────────────────────────
# Semester
# ──────────────────────────────────────────────

SEM_COLS = [
    {"key": "academic_year", "label": "Academic Year"},
    {"key": "semester", "label": "Semester"},
]
SEM_FIELDS = [
    {"label": "Academic Year", "key": "academic_year"},
    {"label": "Semester", "key": "semester"},
]


def semester_list(request):
    qs = Semester.objects.all()
    q = request.GET.get("q", "")
    sem = request.GET.get("semester", "")
    qs = _search(qs, q, ["academic_year"])
    if sem:
        qs = qs.filter(semester=sem)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": SEM_COLS,
        "detail_fields": SEM_FIELDS,
        "q": q,
        "filters": _filters(
            request,
            [
                {
                    "name": "semester",
                    "label": "Semester",
                    "type": "select",
                    "options": [(str(i), str(i)) for i in range(1, 5)],
                },
            ],
        ),
        "query": _query(request, ["q", "semester"]),
        "page_title": "Semesters",
        "list_url": "/semesters/",
        "create_url": "/semesters/create/",
        "edit_name": "semester-edit",
        "delete_name": "semester-delete",
        "detail_name": "semester-detail",
        "page": page,
        "pages": pages,
        "total": total,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/list.html", ctx)


def semester_detail(request, pk):
    item = get_object_or_404(Semester, pk=pk)
    ctx = {
        "item": item,
        "detail_fields": SEM_FIELDS,
        "page_title": str(item),
        "edit_url": f"/semesters/{pk}/edit/",
        "delete_url": f"/semesters/{pk}/delete/",
        "back_url": "/semesters/",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def semester_create(request):
    if request.method == "POST":
        form = SemesterForm(request.POST)
        if form.is_valid():
            obj = form.save()
            _log(LogAction.CREATE, f"Created Semester {obj}", "Semester", str(obj))
            return _write_response(
                request, "close-modal,refresh-table", "semester-list"
            )
    else:
        form = SemesterForm()
    return render(
        request,
        "core/form.html",
        {"form": form, "title": "Create Semester", "action": "/semesters/create/"},
    )


def semester_edit(request, pk):
    item = get_object_or_404(Semester, pk=pk)
    if request.method == "POST":
        form = SemesterForm(request.POST, instance=item)
        if form.is_valid():
            obj = form.save()
            _log(LogAction.UPDATE, f"Updated Semester {obj}", "Semester", str(obj))
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "semester-list",
            )
    else:
        form = SemesterForm(instance=item)
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": f"Edit {item}",
            "action": f"/semesters/{pk}/edit/",
        },
    )


def semester_delete(request, pk):
    item = get_object_or_404(Semester, pk=pk)
    if request.method == "POST":
        label = str(item)
        item.delete()
        _log(LogAction.DELETE, f"Deleted Semester {label}", "Semester", label)
        return _write_response(request, "close-modal,refresh-table", "semester-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/semesters/", f"/semesters/{pk}/delete/"),
    )


# ──────────────────────────────────────────────
# Session (Master Timetable)
# ──────────────────────────────────────────────

SESS_COLS = [
    {"key": "course_code", "label": "Course"},
    {"key": "activity_type", "label": "Type"},
    {"key": "day", "label": "Day"},
    {"key": "start_time", "label": "Start"},
    {"key": "end_time", "label": "End"},
    {"key": "venue", "label": "Venue"},
]
SESS_FIELDS = [
    {"label": "Semester", "key": "semester"},
    {"label": "Course Code", "key": "course_code"},
    {"label": "Activity Type", "key": "activity_type"},
    {"label": "Day", "key": "day"},
    {"label": "Start Time", "key": "start_time"},
    {"label": "End Time", "key": "end_time"},
    {"label": "Venue", "key": "venue"},
]


def session_list(request):
    qs = Session.objects.select_related("semester", "venue").all()
    q = request.GET.get("q", "")
    act = request.GET.get("activity_type", "")
    day = request.GET.get("day", "")
    sem = request.GET.get("semester", "")
    ven = request.GET.get("venue", "")
    start_from = request.GET.get("start_from", "")
    start_to = request.GET.get("start_to", "")
    qs = _search(qs, q, ["course_code", "venue__name"])
    if act:
        qs = qs.filter(activity_type=act)
    if day:
        qs = qs.filter(day=day)
    if sem:
        qs = qs.filter(semester_id=sem)
    if ven:
        qs = qs.filter(venue__name=ven)
    if start_from:
        qs = qs.filter(end_time__gte=start_from)
    if start_to:
        qs = qs.filter(start_time__lte=start_to)
    items, page, pages, total = _paginate(request, qs)
    legacy_sessions = legacy_workshop_sessions()
    ctx = {
        "items": items,
        "columns": SESS_COLS,
        "detail_fields": SESS_FIELDS,
        "q": q,
        "filters": _filters(
            request,
            [
                {
                    "name": "activity_type",
                    "label": "Type",
                    "type": "select",
                    "options": Session._meta.get_field("activity_type").choices,
                },
                {
                    "name": "day",
                    "label": "Day",
                    "type": "select",
                    "options": Session._meta.get_field("day").choices,
                },
                {
                    "name": "semester",
                    "label": "Semester",
                    "type": "select",
                    "options": [
                        (str(s.pk), str(s)) for s in Semester.objects.all()
                    ],
                },
                {
                    "name": "venue",
                    "label": "Venue",
                    "type": "select",
                    "options": [
                        (v.name, v.name) for v in Venue.objects.all()
                    ],
                },
                {"name": "start_from", "label": "Start From", "type": "time"},
                {"name": "start_to", "label": "Start To", "type": "time"},
            ],
        ),
        "query": _query(
            request,
            [
                "q",
                "activity_type",
                "day",
                "semester",
                "venue",
                "start_from",
                "start_to",
            ],
        ),
        "page_title": "Master Timetable",
        "list_url": "/sessions/",
        "create_url": "/sessions/create/",
        "clear_all_url": "/sessions/clear-all/",
        "edit_name": "session-edit",
        "delete_name": "session-delete",
        "detail_name": "session-detail",
        "page": page,
        "pages": pages,
        "total": total,
        "activity_types": Session._meta.get_field("activity_type").choices,
        "days": Session._meta.get_field("day").choices,
        "semesters": Semester.objects.all(),
        "active_activity": act,
        "active_day": day,
        "active_semester": sem,
        "legacy_sessions_count": len(legacy_sessions),
        "legacy_sessions_examples": [
            {
                "course": ses.course_code,
                "day": ses.get_day_display(),
                "time": f"{ses.start_time:%H:%M}-{ses.end_time:%H:%M}",
                "issue": issue,
            }
            for ses, issue in legacy_sessions[:3]
        ],
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/session_list.html", ctx)


def _session_programmes(session):
    """Programme summary for a session's assigned groups.

    Returns a list of dicts: {programme, groups, total_groups, is_all} where
    is_all is True when every StudentGroup of that programme attends.
    """
    sgs = (
        SessionGroup.objects.filter(session=session)
        .select_related("group", "group__programme")
        .order_by("group__programme__code", "group__code")
    )
    by_prog = {}
    for sg in sgs:
        prog = sg.group.programme
        entry = by_prog.setdefault(prog.pk, {"programme": prog, "groups": []})
        entry["groups"].append(sg.group)
    summary = []
    for entry in by_prog.values():
        total = StudentGroup.objects.filter(programme=entry["programme"]).count()
        entry["total_groups"] = total
        entry["is_all"] = total > 0 and len(entry["groups"]) == total
        summary.append(entry)
    summary.sort(key=lambda e: e["programme"].code)
    return summary


def _session_groups_response(request, session):
    """Shared response for group add/remove/cancel actions."""
    groups = SessionGroup.objects.filter(session=session).select_related(
        "group__programme"
    )
    if not _htmx(request):
        return redirect("session-detail", pk=session.pk)
    return render(
        request,
        "core/_session_groups.html",
        {
            "groups": groups,
            "programmes": _session_programmes(session),
            "session_pk": session.pk,
            "all_groups": StudentGroup.objects.select_related("programme").all(),
            "add_group_url": f"/sessions/{session.pk}/add-group/",
        },
    )


def session_assign_lecture_groups(request):
    """Attach Student Groups to LECTURE sessions via ProgrammeCourse -> programme.

    All groups of every programme that studies the course are attached,
    regardless of subgroup (C1/C2), so a lecture holds the full cohort.
    Idempotent: re-running only adds missing links.
    """
    if request.method != "POST":
        return redirect("session-list")
    sessions = Session.objects.filter(activity_type="LECTURE")
    total_sessions = sessions.count()
    processed, created_links, already_linked, skipped_courses = (
        assign_lecture_groups()
    )
    _log(
        LogAction.ASSIGN,
        f"Lecture group assignment complete: {created_links} link(s) created, "
        f"{processed} lecture session(s) processed",
        "Session",
    )
    ctx = {
        "total_sessions": total_sessions,
        "processed": processed,
        "created_links": created_links,
        "already_linked": already_linked,
        "skipped_courses": skipped_courses,
    }
    if _htmx(request):
        # Result-only response. The success/warning summary lives in the green
        # results panel below the button (processed, created, already-linked,
        # skipped mappings) which the user dismisses manually or lets
        # auto-dismiss. No toast is dispatched here — that would duplicate the
        # panel. Genuine request failures are surfaced by the global
        # htmx:responseError / htmx:sendError handlers instead.
        return render(request, "core/_session_assign_result.html", ctx)
    return redirect("session-list")


def session_detail(request, pk):
    item = get_object_or_404(
        Session.objects.select_related("semester", "venue"), pk=pk
    )
    groups = SessionGroup.objects.filter(session=item).select_related("group__programme")
    all_groups = StudentGroup.objects.select_related("programme").all()
    ctx = {
        "item": item,
        "detail_fields": SESS_FIELDS,
        "groups": groups,
        "programmes": _session_programmes(item),
        "all_groups": all_groups,
        "session_pk": pk,
        "page_title": str(item),
        "edit_url": f"/sessions/{pk}/edit/",
        "delete_url": f"/sessions/{pk}/delete/",
        "back_url": "/sessions/",
        "add_group_url": f"/sessions/{pk}/add-group/",
    }
    if _htmx(request):
        hx_target = request.GET.get("target", "")
        if hx_target == "groups":
            return render(request, "core/_session_groups.html", ctx)
        return render(request, "core/_session_detail_content.html", ctx)
    return render(request, "core/session_detail.html", ctx)


def session_create(request):
    if request.method == "POST":
        form = SessionForm(request.POST)
        formset = SessionGroupFormSet(request.POST)
        if form.is_valid() and formset.is_valid():
            session = form.save()
            formset.instance = session
            formset.save()
            _log(
                LogAction.CREATE,
                f"Created Session {session.course_code} {session.activity_type}",
                "Session",
                str(session),
            )
            return _write_response(
                request, "close-modal,refresh-table", "session-list"
            )
    else:
        form = SessionForm()
        formset = SessionGroupFormSet()
    return render(
        request,
        "core/session_form.html",
        {
            "form": form,
            "formset": formset,
            "title": "Create Session",
            "action": "/sessions/create/",
        },
    )


def session_edit(request, pk):
    item = get_object_or_404(Session, pk=pk)
    if request.method == "POST":
        form = SessionForm(request.POST, instance=item)
        formset = SessionGroupFormSet(request.POST, instance=item)
        if form.is_valid() and formset.is_valid():
            session = form.save()
            formset.instance = session
            formset.save()
            _log(
                LogAction.UPDATE,
                f"Updated Session {session.course_code} {session.activity_type}",
                "Session",
                str(session),
            )
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "session-list",
            )
    else:
        form = SessionForm(instance=item)
        formset = SessionGroupFormSet(instance=item)
    return render(
        request,
        "core/session_form.html",
        {
            "form": form,
            "formset": formset,
            "title": f"Edit {item}",
            "action": f"/sessions/{pk}/edit/",
        },
    )


def session_delete(request, pk):
    item = get_object_or_404(Session, pk=pk)
    if request.method == "POST":
        label = str(item)
        item.delete()
        _log(LogAction.DELETE, f"Deleted Session {label}", "Session", label)
        return _write_response(request, "close-modal,refresh-table", "session-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/sessions/", f"/sessions/{pk}/delete/"),
    )


def session_clear_all(request):
    """Clear every Session (cascading its SessionGroup links) from the database."""
    return _clear_all(
        request,
        model=Session,
        list_url="/sessions/",
        clear_url="/sessions/clear-all/",
        page_title="Master Timetable",
        primary_label="Sessions",
        log_resource="Session",
        redirect_name="session-list",
        related_count=((SessionGroup.objects.all(), "Session-group links"),),
        kept_note="Semesters, programmes, courses, student groups and venues are kept.",
    )


def session_add_group(request, pk):
    session = get_object_or_404(Session, pk=pk)
    group_id = request.POST.get("group_id")
    if group_id:
        _, created = SessionGroup.objects.get_or_create(
            session=session, group_id=group_id
        )
        if created:
            group = StudentGroup.objects.filter(pk=group_id).first()
            _log(
                LogAction.ASSIGN,
                f"Assigned group {group} to session {session.course_code}",
                "Session",
                str(session),
            )
    return _session_groups_response(request, session)


def session_remove_group(request, pk, group_pk):
    session = get_object_or_404(Session, pk=pk)
    group = StudentGroup.objects.filter(pk=group_pk).first()
    SessionGroup.objects.filter(session_id=pk, group_id=group_pk).delete()
    _log(
        LogAction.REMOVE,
        f"Removed group {group or group_pk} from session {session.course_code}",
        "Session",
        str(session),
    )
    return _session_groups_response(request, session)


def session_remove_programme_groups(request, pk, programme_pk):
    """Cancel one programme's groups from the lecture session."""
    session = get_object_or_404(Session, pk=pk)
    programme = Programme.objects.filter(pk=programme_pk).first()
    SessionGroup.objects.filter(
        session=session, group__programme_id=programme_pk
    ).delete()
    _log(
        LogAction.REMOVE,
        f"Removed all groups of {programme or 'programme'} from session "
        f"{session.course_code}",
        "Session",
        str(session),
    )
    return _session_groups_response(request, session)


def session_clear_groups(request, pk):
    """Cancel the whole lecture allocation (remove all group links)."""
    session = get_object_or_404(Session, pk=pk)
    SessionGroup.objects.filter(session=session).delete()
    _log(
        LogAction.REMOVE,
        f"Cleared all group links from session {session.course_code}",
        "Session",
        str(session),
    )
    return _session_groups_response(request, session)


# ──────────────────────────────────────────────
# Workshop Allocation
# ──────────────────────────────────────────────

WS_COLS = [
    {"key": "course_code", "label": "Course"},
    {"key": "group_code", "label": "Group"},
    {"key": "day", "label": "Day"},
    {"key": "time_period", "label": "Period"},
    {"key": "start_time", "label": "Start"},
    {"key": "end_time", "label": "End"},
    {"key": "venue", "label": "Workshop"},
]
WS_FIELDS = [
    {"label": "Semester", "key": "semester"},
    {"label": "Course Code", "key": "course_code"},
    {"label": "Group Code", "key": "group_code"},
    {"label": "Workshop", "key": "workshop"},
    {"label": "Day", "key": "day"},
    {"label": "Time Period", "key": "time_period"},
    {"label": "Position", "key": "position"},
    {"label": "Schedule Section", "key": "schedule_section"},
    {"label": "Week Start", "key": "week_start"},
    {"label": "Week End", "key": "week_end"},
    {"label": "Start Time", "key": "start_time"},
    {"label": "End Time", "key": "end_time"},
    {"label": "Workshop", "key": "venue"},
    {"label": "Year of Study", "key": "year_of_study"},
]


def workshop_list(request):
    qs = WorkshopAllocation.objects.select_related("semester").all()
    q = request.GET.get("q", "")
    sem = request.GET.get("semester", "")
    day = request.GET.get("day", "")
    start_from = request.GET.get("start_from", "")
    start_to = request.GET.get("start_to", "")
    qs = _search(qs, q, ["course_code", "group_code", "venue"])
    if sem:
        qs = qs.filter(semester_id=sem)
    if day:
        qs = qs.filter(day=day)
    if start_from:
        qs = qs.filter(end_time__gte=start_from)
    if start_to:
        qs = qs.filter(start_time__lte=start_to)
    items, page, pages, total = _paginate(request, qs)
    legacy = legacy_workshop_allocations()
    ctx = {
        "items": items,
        "columns": WS_COLS,
        "detail_fields": WS_FIELDS,
        "q": q,
        "filters": _filters(
            request,
            [
                {
                    "name": "semester",
                    "label": "Semester",
                    "type": "select",
                    "options": [
                        (str(s.pk), str(s)) for s in Semester.objects.all()
                    ],
                },
                {
                    "name": "day",
                    "label": "Day",
                    "type": "select",
                    "options": Day.choices,
                },
                {"name": "start_from", "label": "Start From", "type": "time"},
                {"name": "start_to", "label": "Start To", "type": "time"},
            ],
        ),
        "query": _query(
            request, ["q", "semester", "day", "start_from", "start_to"]
        ),
        "page_title": "Workshop Allocations",
        "list_url": "/workshops/",
        "create_url": "/workshops/create/",
        "clear_all_url": "/workshops/clear-all/",
        "edit_name": "workshop-edit",
        "delete_name": "workshop-delete",
        "detail_name": "workshop-detail",
        "page": page,
        "pages": pages,
        "total": total,
        "semesters": Semester.objects.all(),
        "active_semester": sem,
        "legacy_count": len(legacy),
        "legacy_examples": [
            {
                "course": rec.course_code,
                "group": rec.group_code,
                "day": rec.get_day_display(),
                "time": (
                    f"{rec.start_time:%H:%M}-{rec.end_time:%H:%M} "
                    f"({rec.get_time_period_display()})"
                    if rec.time_period
                    else f"{rec.start_time:%H:%M}-{rec.end_time:%H:%M}"
                ),
                "issue": issue,
            }
            for rec, issue in legacy[:3]
        ],
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/workshop_list.html", ctx)


def workshop_detail(request, pk):
    item = get_object_or_404(
        WorkshopAllocation.objects.select_related("semester"), pk=pk
    )
    ctx = {
        "item": item,
        "detail_fields": WS_FIELDS,
        "page_title": str(item),
        "edit_url": f"/workshops/{pk}/edit/",
        "delete_url": f"/workshops/{pk}/delete/",
        "back_url": "/workshops/",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def workshop_create(request):
    if request.method == "POST":
        form = WorkshopAllocationForm(request.POST)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.CREATE,
                f"Created Workshop Allocation {obj}",
                "Workshop Allocation",
                str(obj),
            )
            return _write_response(
                request, "close-modal,refresh-table", "workshop-list"
            )
    else:
        form = WorkshopAllocationForm()
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": "Create Workshop Allocation",
            "action": "/workshops/create/",
        },
    )


def workshop_edit(request, pk):
    item = get_object_or_404(WorkshopAllocation, pk=pk)
    if request.method == "POST":
        form = WorkshopAllocationForm(request.POST, instance=item)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.UPDATE,
                f"Updated Workshop Allocation {obj}",
                "Workshop Allocation",
                str(obj),
            )
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "workshop-list",
            )
    else:
        form = WorkshopAllocationForm(instance=item)
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": f"Edit {item}",
            "action": f"/workshops/{pk}/edit/",
        },
    )


def workshop_delete(request, pk):
    item = get_object_or_404(WorkshopAllocation, pk=pk)
    if request.method == "POST":
        label = str(item)
        item.delete()
        _log(
            LogAction.DELETE,
            f"Deleted Workshop Allocation {label}",
            "Workshop Allocation",
            label,
        )
        return _write_response(request, "close-modal,refresh-table", "workshop-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/workshops/", f"/workshops/{pk}/delete/"),
    )


def workshop_clear_all(request):
    """Clear every WorkshopAllocation record from the database."""
    return _clear_all(
        request,
        model=WorkshopAllocation,
        list_url="/workshops/",
        clear_url="/workshops/clear-all/",
        page_title="Workshop Allocations",
        primary_label="Workshop Allocations",
        log_resource="Workshop Allocation",
        redirect_name="workshop-list",
        kept_note=(
            "Sessions, semesters, programmes, courses, student groups and venues are kept."
        ),
    )


# ──────────────────────────────────────────────
# Technical Drawing Allocation
# ──────────────────────────────────────────────

TD_COLS = [
    {"key": "course_code", "label": "Course"},
    {"key": "group_code", "label": "Group"},
    {"key": "day", "label": "Day"},
    {"key": "start_time", "label": "Start"},
    {"key": "end_time", "label": "End"},
    {"key": "venue", "label": "Venue"},
]
TD_FIELDS = [
    {"label": "Semester", "key": "semester"},
    {"label": "Course Code", "key": "course_code"},
    {"label": "Group Code", "key": "group_code"},
    {"label": "Day", "key": "day"},
    {"label": "Start Time", "key": "start_time"},
    {"label": "End Time", "key": "end_time"},
    {"label": "Venue", "key": "venue"},
]


def td_list(request):
    qs = TechnicalDrawingAllocation.objects.select_related("semester").all()
    q = request.GET.get("q", "")
    sem = request.GET.get("semester", "")
    day = request.GET.get("day", "")
    start_from = request.GET.get("start_from", "")
    start_to = request.GET.get("start_to", "")
    qs = _search(qs, q, ["course_code", "group_code", "venue"])
    if sem:
        qs = qs.filter(semester_id=sem)
    if day:
        qs = qs.filter(day=day)
    if start_from:
        qs = qs.filter(end_time__gte=start_from)
    if start_to:
        qs = qs.filter(start_time__lte=start_to)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": TD_COLS,
        "detail_fields": TD_FIELDS,
        "q": q,
        "filters": _filters(
            request,
            [
                {
                    "name": "semester",
                    "label": "Semester",
                    "type": "select",
                    "options": [
                        (str(s.pk), str(s)) for s in Semester.objects.all()
                    ],
                },
                {
                    "name": "day",
                    "label": "Day",
                    "type": "select",
                    "options": Day.choices,
                },
                {"name": "start_from", "label": "Start From", "type": "time"},
                {"name": "start_to", "label": "Start To", "type": "time"},
            ],
        ),
        "query": _query(
            request, ["q", "semester", "day", "start_from", "start_to"]
        ),
        "page_title": "Technical Drawing Allocations",
        "list_url": "/td/",
        "create_url": "/td/create/",
        "clear_all_url": "/td/clear-all/",
        "edit_name": "td-edit",
        "delete_name": "td-delete",
        "detail_name": "td-detail",
        "page": page,
        "pages": pages,
        "total": total,
        "semesters": Semester.objects.all(),
        "active_semester": sem,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/list.html", ctx)


def td_detail(request, pk):
    item = get_object_or_404(
        TechnicalDrawingAllocation.objects.select_related("semester"), pk=pk
    )
    ctx = {
        "item": item,
        "detail_fields": TD_FIELDS,
        "page_title": str(item),
        "edit_url": f"/td/{pk}/edit/",
        "delete_url": f"/td/{pk}/delete/",
        "back_url": "/td/",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def td_create(request):
    if request.method == "POST":
        form = TechnicalDrawingAllocationForm(request.POST)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.CREATE,
                f"Created TD Allocation {obj}",
                "TD Allocation",
                str(obj),
            )
            return _write_response(
                request, "close-modal,refresh-table", "td-list"
            )
    else:
        form = TechnicalDrawingAllocationForm()
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": "Create TD Allocation",
            "action": "/td/create/",
        },
    )


def td_edit(request, pk):
    item = get_object_or_404(TechnicalDrawingAllocation, pk=pk)
    if request.method == "POST":
        form = TechnicalDrawingAllocationForm(request.POST, instance=item)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.UPDATE,
                f"Updated TD Allocation {obj}",
                "TD Allocation",
                str(obj),
            )
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "td-list",
            )
    else:
        form = TechnicalDrawingAllocationForm(instance=item)
    return render(
        request,
        "core/form.html",
        {"form": form, "title": f"Edit {item}", "action": f"/td/{pk}/edit/"},
    )


def td_delete(request, pk):
    item = get_object_or_404(TechnicalDrawingAllocation, pk=pk)
    if request.method == "POST":
        label = str(item)
        item.delete()
        _log(LogAction.DELETE, f"Deleted TD Allocation {label}", "TD Allocation", label)
        return _write_response(request, "close-modal,refresh-table", "td-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/td/", f"/td/{pk}/delete/"),
    )


def td_clear_all(request):
    """Clear every TechnicalDrawingAllocation record from the database."""
    return _clear_all(
        request,
        model=TechnicalDrawingAllocation,
        list_url="/td/",
        clear_url="/td/clear-all/",
        page_title="Technical Drawing Allocations",
        primary_label="Technical Drawing Allocations",
        log_resource="TD Allocation",
        redirect_name="td-list",
        kept_note=(
            "Sessions, semesters, programmes, courses, student groups and venues are kept."
        ),
    )


# ──────────────────────────────────────────────
# Programme Course
# ──────────────────────────────────────────────

PC_COLS = [
    {"key": "programme", "label": "Programme"},
    {"key": "course_code", "label": "Course Code"},
    {"key": "course_name", "label": "Course Name"},
    {"key": "semester", "label": "Semester"},
]
PC_FIELDS = [
    {"label": "Programme", "key": "programme"},
    {"label": "Course Code", "key": "course_code"},
    {"label": "Course Name", "key": "course_name"},
    {"label": "Semester", "key": "semester"},
]


def course_list(request):
    qs = ProgrammeCourse.objects.select_related("programme").all()
    q = request.GET.get("q", "")
    prog = request.GET.get("programme", "")
    sem = request.GET.get("semester", "")
    qs = _search(
        qs, q, ["course_code", "course_name", "programme__code", "programme__name"]
    )
    if prog:
        qs = qs.filter(programme__code=prog)
    if sem:
        qs = qs.filter(semester=sem)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": PC_COLS,
        "detail_fields": PC_FIELDS,
        "q": q,
        "filters": _filters(
            request,
            [
                {
                    "name": "programme",
                    "label": "Programme",
                    "type": "select",
                    "options": [
                        (p.code, p.name) for p in Programme.objects.all()
                    ],
                },
                {
                    "name": "semester",
                    "label": "Semester",
                    "type": "select",
                    "options": [(str(i), str(i)) for i in range(1, 5)],
                },
            ],
        ),
        "query": _query(request, ["q", "programme", "semester"]),
        "page_title": "Programme Courses",
        "list_url": "/courses/",
        "create_url": "/courses/create/",
        "clear_all_url": "/courses/clear-all/",
        "edit_name": "course-edit",
        "delete_name": "course-delete",
        "detail_name": "course-detail",
        "page": page,
        "pages": pages,
        "total": total,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/list.html", ctx)


def course_detail(request, pk):
    item = get_object_or_404(
        ProgrammeCourse.objects.select_related("programme"), pk=pk
    )
    ctx = {
        "item": item,
        "detail_fields": PC_FIELDS,
        "page_title": str(item),
        "edit_url": f"/courses/{pk}/edit/",
        "delete_url": f"/courses/{pk}/delete/",
        "back_url": "/courses/",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def course_create(request):
    if request.method == "POST":
        form = ProgrammeCourseForm(request.POST)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.CREATE,
                f"Created Programme Course {obj}",
                "Programme Course",
                str(obj),
            )
            return _write_response(
                request, "close-modal,refresh-table", "course-list"
            )
    else:
        form = ProgrammeCourseForm()
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": "Create Programme Course",
            "action": "/courses/create/",
        },
    )


def course_edit(request, pk):
    item = get_object_or_404(ProgrammeCourse, pk=pk)
    if request.method == "POST":
        form = ProgrammeCourseForm(request.POST, instance=item)
        if form.is_valid():
            obj = form.save()
            _log(
                LogAction.UPDATE,
                f"Updated Programme Course {obj}",
                "Programme Course",
                str(obj),
            )
            return _write_response(
                request,
                "close-modal,refresh-table,refresh-detail",
                "course-list",
            )
    else:
        form = ProgrammeCourseForm(instance=item)
    return render(
        request,
        "core/form.html",
        {
            "form": form,
            "title": f"Edit {item}",
            "action": f"/courses/{pk}/edit/",
        },
    )


def course_delete(request, pk):
    item = get_object_or_404(ProgrammeCourse, pk=pk)
    if request.method == "POST":
        label = str(item)
        item.delete()
        _log(
            LogAction.DELETE,
            f"Deleted Programme Course {label}",
            "Programme Course",
            label,
        )
        return _write_response(request, "close-modal,refresh-table", "course-list")
    return render(
        request,
        "core/delete.html",
        _delete_context(request, item, "/courses/", f"/courses/{pk}/delete/"),
    )


def course_clear_all(request):
    """Clear every ProgrammeCourse mapping; programmes and other data are kept."""
    return _clear_all(
        request,
        model=ProgrammeCourse,
        list_url="/courses/",
        clear_url="/courses/clear-all/",
        page_title="Programme Courses",
        primary_label="Programme Courses",
        log_resource="Programme Course",
        redirect_name="course-list",
        kept_note="Programme, student group, session, semester and venue records are kept.",
    )


# ──────────────────────────────────────────────
# Import Views
# ──────────────────────────────────────────────

def _import_status(result) -> ImportStatus:
    """Classify an ImportResult as SUCCESS / PARTIAL / FAILED.

    FAILED means nothing was written and errors blocked the import; PARTIAL
    means some rows were written while others errored; SUCCESS is clean.
    """
    if result.errors:
        if (result.created + result.updated) > 0:
            return ImportStatus.PARTIAL
        return ImportStatus.FAILED
    return ImportStatus.SUCCESS


def _record_import(request, result, import_type, import_title, filename):
    """Persist an ImportHistory record summarising a finished import.

    Returns the saved record so the UI can link straight to its details.
    Only the structured summary is stored (via ``ImportResult.snapshot``) —
    never the uploaded file's contents.
    """
    status = _import_status(result)
    history = ImportHistory.objects.create(
        import_type=import_type,
        import_title=import_title,
        filename=filename,
        user=request.user if request.user.is_authenticated else None,
        status=status,
        rows_processed=result.created + result.updated + result.skipped,
        created=result.created,
        updated=result.updated,
        skipped=result.skipped,
        error_count=len(result.errors),
        details=json.dumps(result.snapshot(), ensure_ascii=True),
    )
    return history


def _import_toast(result) -> tuple[str, str]:
    """Important-event notification for an import (message, toast type).

    Routine row/field validation messages are intentionally NOT notified here
    — they are always visible in the results panel. Only the overall outcome
    triggers a pop-up.
    """
    status = _import_status(result)
    parts = [f"{result.created} created"]
    if result.updated:
        parts.append(f"{result.updated} updated")
    if result.skipped:
        parts.append(f"{result.skipped} skipped")
    summary = ", ".join(parts)
    if status == ImportStatus.FAILED:
        return (
            f"Import failed: {len(result.errors)} error(s) blocked "
            "the import. Review the errors below.",
            "error",
        )
    if status == ImportStatus.PARTIAL:
        return (
            f"Import completed with issues: {summary}, "
            f"{len(result.errors)} row(s) with errors — review below.",
            "error",
        )
    return (
        f"Import complete: {summary}.",
        "success",
    )

IMPORT_TYPES = {
    "programmes": {
        "title": "Programmes",
        "columns": "code, name",
        "hint": "Aliases accepted: 'code', 'programme_code', 'programme', 'program'; 'name', 'programme_name', 'title'.",
        "fn": import_programmes_from_excel,
    },
    "student-groups": {
        "title": "Student Groups",
        "columns": "programme_code, group_code",
        "hint": "Aliases accepted for both columns. Programme can be a code or a full name.",
        "fn": import_student_groups_from_excel,
    },
    "programme-courses": {
        "title": "Programme Courses",
        "columns": "programme_code, course_code, course_name, semester",
        "hint": (
            "Aliases accepted ('programme'/'program', 'course_code'/'course', "
            "'course_name'/'course'). A programme name (e.g. 'BSc. in Chemical and "
            "Processing Engineering') is recognised and the missing programme is "
            "created automatically."
        ),
        "fn": import_programme_courses_from_excel,
    },
    "venues": {
        "title": "Venues",
        "columns": "name, capacity",
        "hint": "Aliases accepted: 'name'/'venue'/'room'; 'capacity'/'seats'.",
        "fn": import_venues_from_excel,
    },
    "master-timetable": {
        "title": "Master Timetable",
        "columns": "course_code, activity_type, day, start_time, end_time, venue, group",
        "semester": "required",
        "hint": (
            "Readable aliases accepted ('course', 'type', 'start', 'end', 'room', "
            "'groups', ...). Comma-separated course codes are split into separate "
            "sessions. Choose the academic Semester this timetable belongs to above — "
            "it is never auto-detected. LECTURE sessions are automatically linked to "
            "every programme group that studies the course."
        ),
        "fn": import_master_timetable_from_excel,
    },
    "workshop-allocation": {
        "title": "Workshop Allocation",
        "columns": (
            "Flat: course_code, group_code, day, start_time, end_time, workshop — "
            "or drop in the raw university Workshop Schedule workbook (matrix) directly"
        ),
        "semester": "optional",
        "hint": (
            "Both formats supported automatically. The raw university workshop "
            "workbook (GROUPS/POSITION/SCHEDULE/KEY layout) is detected and parsed "
            "as-is; week ranges, workshop, position, day and Morning/Afternoon "
            "period come from the workbook, with the semester read from the title "
            "(created automatically if missing). The flat format also accepts a "
            "single 'time' range column such as 08:00-10:00. Leave the semester "
            "unset to auto-detect it."
        ),
        "fn": import_workshop_allocation_from_excel,
    },
    "td-allocation": {
        "title": "TD Allocation",
        "columns": "course_code, group_code, day, start_time, end_time, venue",
        "semester": "optional",
        "hint": (
            "Requires a course_code column. Accepts the pivoted layout "
            "(Day/Time/Group/Venue with merged cells): day, time and venue are "
            "forward-filled, and a 'time' range like 09:00-12:00 is split into "
            "start/end automatically."
        ),
        "fn": import_td_allocation_from_excel,
    },
}


def import_hub(request):
    latest_by_type = {}
    for key in IMPORT_TYPES:
        latest_by_type[key] = (
            ImportHistory.objects.filter(import_type=key).first()
        )
    ctx = {
        "page_title": "Import Data",
        "import_types": IMPORT_TYPES,
        "latest_imports": latest_by_type,
        "recent_imports": ImportHistory.objects.all()[:8],
        "import_history_count": ImportHistory.objects.count(),
        "history_url": reverse("import-history"),
    }
    return render(request, "core/import_hub.html", ctx)


def import_upload(request, import_type):
    if import_type not in IMPORT_TYPES:
        return HttpResponseBadRequest("Unknown import type")
    info = IMPORT_TYPES[import_type]
    semester_choice = info.get("semester", "")
    semesters = Semester.objects.all()
    no_semester = semester_choice == "required" and not semesters
    active_semester = ""
    saved_import = None
    import_status = None
    if request.method == "POST":
        form = FileUploadForm(request.POST, request.FILES)
        result: ImportResult = ImportResult()
        semester_id = None
        if semester_choice:
            active_semester = request.POST.get("semester", "").strip()
            if not active_semester:
                if semester_choice == "required":
                    result.errors.append(
                        "Please choose the academic semester this timetable belongs to."
                    )
            else:
                try:
                    sem = Semester.objects.get(pk=active_semester)
                except (ValueError, Semester.DoesNotExist):
                    result.errors.append(
                        f"Semester '{active_semester}' does not exist — create it first."
                    )
                else:
                    semester_id = sem.pk
        if form.is_valid() and not result.errors:
            uploaded = form.cleaned_data["file"]
            filename = getattr(uploaded, "name", "")
            try:
                if hasattr(uploaded, "temporary_file_path"):
                    source = uploaded.temporary_file_path()
                else:
                    uploaded.seek(0)
                    source = uploaded
                if semester_id is not None:
                    result = info["fn"](source, semester_id=semester_id)
                else:
                    result = info["fn"](source)
            except Exception as exc:
                result.errors.append(str(exc))
            # A complete summary is always saved so users can review this
            # import (and earlier imports) from the history pages.
            saved_import = _record_import(
                request, result, import_type, info["title"], filename
            )
            msg = f"Imported {info['title']}: {result.created} created, {result.updated} updated"
            if result.skipped:
                msg += f", {result.skipped} skipped"
            if result.errors:
                msg += f", {len(result.errors)} error(s)"
            _log(
                LogAction.IMPORT,
                msg,
                info["title"],
                filename[:300],
            )
        elif not form.is_valid():
            result.errors.append("No file attached or invalid upload.")
        import_status = _import_status(result)
        venue_conflict_token = ""
        if import_type == "venues" and result.venue_conflicts:
            venue_conflict_token = uuid.uuid4().hex[:10]
            request.session[f"venue_conflicts:{venue_conflict_token}"] = {
                "open": result.venue_conflicts,
                "resolved": [],
            }
            # Keep at most 5 concurrent conflict panels per browser.
            prefix = "venue_conflicts:"
            tokens = [
                k for k in request.session.keys() if k.startswith(prefix)
            ]
            for old in tokens[:-5]:
                del request.session[old]
            request.session.modified = True
        ctx = {
            "result": result,
            "import_type": import_type,
            "import_title": info["title"],
            "columns": info["columns"],
            "hint": info.get("hint", ""),
            "form": form,
            "semester_choice": semester_choice,
            "semesters": semesters,
            "active_semester": active_semester,
            "no_semester": no_semester,
            "import_status": import_status,
            "saved_import": saved_import,
            "venue_conflict_token": venue_conflict_token,
            "resolved_venue_conflicts": [],
            "page_title": f"Import {info['title']}",
        }
        if _htmx(request):
            response = render(request, "core/import_result.html", ctx)
            if saved_import is not None:
                toast_message, toast_type = _import_toast(result)
                response["HX-Trigger"] = json.dumps(
                    {"import-toast": {"message": toast_message, "type": toast_type}}
                )
            return response
        return render(request, "core/import_upload.html", ctx)
    form = FileUploadForm()
    return render(
        request,
        "core/import_upload.html",
        {
            "form": form,
            "import_type": import_type,
            "import_title": info["title"],
            "columns": info["columns"],
            "hint": info.get("hint", ""),
            "semester_choice": semester_choice,
            "semesters": semesters,
            "active_semester": active_semester,
            "no_semester": no_semester,
            "page_title": f"Import {info['title']}",
        },
    )


def import_history(request, import_type=None):
    """Overall import history, optionally narrowed to one import type."""
    valid_type = import_type in IMPORT_TYPES
    if import_type is not None and not valid_type:
        return HttpResponseBadRequest("Unknown import type")
    qs = ImportHistory.objects.all()
    if valid_type:
        qs = qs.filter(import_type=import_type)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "page_title": "Import History",
        "items": items,
        "import_types": IMPORT_TYPES,
        "active_type": import_type if valid_type else "",
        "type_options": [(key, info["title"]) for key, info in IMPORT_TYPES.items()],
        "page": page,
        "pages": pages,
        "total": total,
        "back_url": reverse("import-hub"),
    }
    return render(request, "core/import_history.html", ctx)


def import_history_detail(request, pk):
    """Full details of one saved import, including every recorded row/field
    error, missing entity and automatic fix."""
    record = get_object_or_404(ImportHistory, pk=pk)
    try:
        details = json.loads(record.details or "{}")
    except (ValueError, TypeError):
        details = {}
    details.setdefault("error_count", record.error_count)
    details.setdefault(
        "venue_exact_count",
        sum(1 for r in details.get("venue_resolutions", []) if r.get("action") == "exact"),
    )
    details.setdefault(
        "venue_matched_count",
        sum(1 for r in details.get("venue_resolutions", []) if r.get("action") == "matched"),
    )
    details.setdefault(
        "venue_created_count",
        sum(1 for r in details.get("venue_resolutions", []) if r.get("action") == "created"),
    )
    ctx = {
        "page_title": "Import Details",
        "record": record,
        "d": details,
        "back_url": reverse("import-history"),
        "type_history_url": reverse(
            "import-history-type", args=[record.import_type]
        )
        if record.import_type in IMPORT_TYPES
        else reverse("import-history"),
        "hub_url": reverse("import-hub"),
    }
    return render(request, "core/import_history_detail.html", ctx)
