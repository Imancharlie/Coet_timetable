import io
from pathlib import Path
from urllib.parse import urlencode

from django.db.models import Q, Count
from django.http import HttpResponse, HttpResponseBadRequest
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from .timetable_pdf import render_programme_timetable

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
    WorkshopAllocationForm,
)
from .importers import (
    ImportResult,
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
    issues_for,
    suggested_name,
)

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
        {"item": item, "title": f"Delete {item}?", "back_url": "/programmes/", "delete_url": f"/programmes/{pk}/delete/"},
    )


def export_timetable(request):
    """Page listing every registered programme, each with an Export button."""
    programmes = Programme.objects.all()
    semester = Semester.objects.order_by("-academic_year", "-semester").first()
    ctx = {
        "page_title": "Export Timetable",
        "programmes": programmes,
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
        semester = (
            Semester.objects.filter(
                sessions__session_groups__group__programme=programme
            )
            .order_by("-academic_year", "-semester")
            .first()
            or Semester.objects.first()
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
        {"item": item, "title": f"Delete {item}?", "back_url": "/groups/", "delete_url": f"/groups/{pk}/delete/"},
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
        {"item": item, "title": f"Delete {item}?", "back_url": "/venues/", "delete_url": f"/venues/{pk}/delete/"},
    )


def _merge_venues(source, target):
    """Fold one venue into another and keep every reference in sync."""
    target.refresh_from_db()
    Session.objects.filter(venue=source).update(venue=target)
    cleaned = target.name
    WorkshopAllocation.objects.filter(venue=source.name).update(venue=cleaned)
    TechnicalDrawingAllocation.objects.filter(venue=source.name).update(venue=cleaned)
    if (target.capacity or 0) <= 0 and (source.capacity or 0) > 0:
        target.capacity = source.capacity
        target.save(update_fields=["capacity"])
    label = source.name
    _log(
        LogAction.UPDATE,
        f"Merged Venue {label} into {target.name}",
        "Venue",
        target.name,
    )
    source.delete()


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
                form = VenueForm(payload, instance=venue)
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
    }
    return render(request, "core/venue_recycle.html", ctx)


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
        {"item": item, "title": f"Delete {item}?", "back_url": "/semesters/", "delete_url": f"/semesters/{pk}/delete/"},
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
    processed = 0
    created_links = 0
    already_linked = 0
    skipped_courses = set()
    for session in sessions:
        prog_ids = list(
            ProgrammeCourse.objects.filter(course_code=session.course_code).values_list(
                "programme_id", flat=True
            )
        )
        if not prog_ids:
            skipped_courses.add(session.course_code)
            continue
        for group in StudentGroup.objects.filter(programme_id__in=prog_ids):
            _, created = SessionGroup.objects.get_or_create(
                session=session, group=group
            )
            if created:
                created_links += 1
            else:
                already_linked += 1
        processed += 1
    _log(
        LogAction.ASSIGN,
        f"Lecture group assignment complete: {created_links} link(s) created, "
        f"{processed} lecture session(s) processed",
        "Session",
    )
    ctx = {
        "total_sessions": sessions.count(),
        "processed": processed,
        "created_links": created_links,
        "already_linked": already_linked,
        "skipped_courses": sorted(skipped_courses),
    }
    if _htmx(request):
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
        {"item": item, "title": f"Delete {item}?", "back_url": "/sessions/", "delete_url": f"/sessions/{pk}/delete/"},
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
    {"key": "venue", "label": "Venue"},
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
    {"label": "Venue", "key": "venue"},
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
        "edit_name": "workshop-edit",
        "delete_name": "workshop-delete",
        "detail_name": "workshop-detail",
        "page": page,
        "pages": pages,
        "total": total,
        "semesters": Semester.objects.all(),
        "active_semester": sem,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/list.html", ctx)


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
        {"item": item, "title": f"Delete {item}?", "back_url": "/workshops/", "delete_url": f"/workshops/{pk}/delete/"},
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
        {"item": item, "title": f"Delete {item}?", "back_url": "/td/", "delete_url": f"/td/{pk}/delete/"},
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
        {"item": item, "title": f"Delete {item}?", "back_url": "/courses/", "delete_url": f"/courses/{pk}/delete/"},
    )


# ──────────────────────────────────────────────
# Import Views
# ──────────────────────────────────────────────

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
        "hint": (
            "Readable aliases accepted ('course', 'type', 'start', 'end', 'room', "
            "'groups', ...). Comma-separated course codes are split into separate "
            "sessions. If no semester exists yet, the current academic year's "
            "semester 1 is created automatically."
        ),
        "fn": import_master_timetable_from_excel,
    },
    "workshop-allocation": {
        "title": "Workshop Allocation",
        "columns": (
            "Flat: course_code, group_code, day, start_time, end_time, venue — "
            "or drop in the raw university Workshop Schedule workbook (matrix) directly"
        ),
        "hint": (
            "Both formats supported automatically. The raw university workshop "
            "workbook (GROUPS/POSITION/SCHEDULE/KEY layout) is detected and parsed "
            "as-is; week ranges, workshop, position, day and Morning/Afternoon "
            "period come from the workbook, with the semester read from the title "
            "(created automatically if missing). The flat format also accepts a "
            "single 'time' range column such as 08:00-10:00."
        ),
        "fn": import_workshop_allocation_from_excel,
    },
    "td-allocation": {
        "title": "TD Allocation",
        "columns": "course_code, group_code, day, start_time, end_time, venue",
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
    ctx = {
        "page_title": "Import Data",
        "import_types": IMPORT_TYPES,
    }
    return render(request, "core/import_hub.html", ctx)


def import_upload(request, import_type):
    if import_type not in IMPORT_TYPES:
        return HttpResponseBadRequest("Unknown import type")
    info = IMPORT_TYPES[import_type]
    if request.method == "POST":
        form = FileUploadForm(request.POST, request.FILES)
        result: ImportResult = ImportResult()
        if form.is_valid():
            uploaded = form.cleaned_data["file"]
            try:
                if hasattr(uploaded, "temporary_file_path"):
                    result = info["fn"](uploaded.temporary_file_path())
                else:
                    uploaded.seek(0)
                    result = info["fn"](uploaded)
            except Exception as exc:
                result.errors.append(str(exc))
            msg = f"Imported {info['title']}: {result.created} created, {result.updated} updated"
            if result.skipped:
                msg += f", {result.skipped} skipped"
            if result.errors:
                msg += f", {len(result.errors)} error(s)"
            _log(
                LogAction.IMPORT,
                msg,
                info["title"],
                getattr(uploaded, "name", "")[:300],
            )
        else:
            result.errors.append("No file attached or invalid upload.")
        ctx = {
            "result": result,
            "import_type": import_type,
            "import_title": info["title"],
            "columns": info["columns"],
            "hint": info.get("hint", ""),
            "form": form,
            "page_title": f"Import {info['title']}",
        }
        if _htmx(request):
            return render(request, "core/import_result.html", ctx)
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
            "page_title": f"Import {info['title']}",
        },
    )
