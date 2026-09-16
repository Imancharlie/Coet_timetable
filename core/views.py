import io
from pathlib import Path

from django.db.models import Q, Count
from django.http import HttpResponse, HttpResponseBadRequest
from django.shortcuts import get_object_or_404, render

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
    import_semesters_from_excel,
    import_student_groups_from_excel,
    import_td_allocation_from_excel,
    import_venues_from_excel,
    import_workshop_allocation_from_excel,
)
from .models import (
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

NAV = "active_nav"
HTMX_HEADER = "HX-Request"


def _htmx(request):
    return request.headers.get(HTMX_HEADER)


def _search(qs, q, fields):
    if not q:
        return qs
    q_obj = Q()
    for f in fields:
        q_obj |= Q(**{f + "__icontains": q})
    return qs.filter(q_obj)


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
    return render(
        request,
        "dashboard.html",
        {
            NAV: "dashboard",
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
        },
    )


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
        NAV: "programmes",
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
        NAV: "programmes",
        "page_title": str(item),
        "edit_url": f"/programmes/{pk}/edit/",
        "delete_url": f"/programmes/{pk}/delete/",
        "back_url": "/programmes/",
    }
    if _htmx(request):
        return render(request, "core/_detail_content.html", ctx)
    return render(request, "core/detail.html", ctx)


def programme_create(request):
    if request.method == "POST":
        form = ProgrammeForm(request.POST)
        if form.is_valid():
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
    return render(
        request,
        "core/delete.html",
        {"item": item, "title": f"Delete {item}?", "back_url": "/programmes/", "delete_url": f"/programmes/{pk}/delete/"},
    )


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
        NAV: "groups",
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
        NAV: "groups",
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
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
    qs = _search(qs, q, ["name"])
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": VENUE_COLS,
        "detail_fields": VENUE_FIELDS,
        "q": q,
        NAV: "venues",
        "page_title": "Venues",
        "list_url": "/venues/",
        "create_url": "/venues/create/",
        "edit_name": "venue-edit",
        "delete_name": "venue-delete",
        "detail_name": "venue-detail",
        "page": page,
        "pages": pages,
        "total": total,
    }
    if _htmx(request):
        return render(request, "core/_table_and_cards.html", ctx)
    return render(request, "core/list.html", ctx)


def venue_detail(request, pk):
    item = get_object_or_404(Venue, pk=pk)
    ctx = {
        "item": item,
        "detail_fields": VENUE_FIELDS,
        NAV: "venues",
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
    return render(
        request,
        "core/delete.html",
        {"item": item, "title": f"Delete {item}?", "back_url": "/venues/", "delete_url": f"/venues/{pk}/delete/"},
    )


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
    qs = _search(qs, q, ["academic_year"])
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": SEM_COLS,
        "detail_fields": SEM_FIELDS,
        "q": q,
        NAV: "semesters",
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
        NAV: "semesters",
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
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
    qs = _search(qs, q, ["course_code", "venue__name"])
    if act:
        qs = qs.filter(activity_type=act)
    if day:
        qs = qs.filter(day=day)
    if sem:
        qs = qs.filter(semester_id=sem)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": SESS_COLS,
        "detail_fields": SESS_FIELDS,
        "q": q,
        NAV: "sessions",
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
    return render(request, "core/list.html", ctx)


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
        "all_groups": all_groups,
        "session_pk": pk,
        NAV: "sessions",
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
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
    return render(
        request,
        "core/delete.html",
        {"item": item, "title": f"Delete {item}?", "back_url": "/sessions/", "delete_url": f"/sessions/{pk}/delete/"},
    )


def session_add_group(request, pk):
    session = get_object_or_404(Session, pk=pk)
    group_id = request.POST.get("group_id")
    if group_id:
        SessionGroup.objects.get_or_create(
            session=session, group_id=group_id
        )
    groups = SessionGroup.objects.filter(session=session).select_related(
        "group__programme"
    )
    return render(
        request,
        "core/_session_groups.html",
        {
            "groups": groups,
            "session_pk": pk,
            "all_groups": StudentGroup.objects.select_related("programme").all(),
            "add_group_url": f"/sessions/{pk}/add-group/",
        },
    )


def session_remove_group(request, pk, group_pk):
    SessionGroup.objects.filter(session_id=pk, group_id=group_pk).delete()
    session = get_object_or_404(Session, pk=pk)
    groups = SessionGroup.objects.filter(session=session).select_related(
        "group__programme"
    )
    return render(
        request,
        "core/_session_groups.html",
        {
            "groups": groups,
            "session_pk": pk,
            "all_groups": StudentGroup.objects.select_related("programme").all(),
            "add_group_url": f"/sessions/{pk}/add-group/",
        },
    )


# ──────────────────────────────────────────────
# Workshop Allocation
# ──────────────────────────────────────────────

WS_COLS = [
    {"key": "course_code", "label": "Course"},
    {"key": "group_code", "label": "Group"},
    {"key": "day", "label": "Day"},
    {"key": "start_time", "label": "Start"},
    {"key": "end_time", "label": "End"},
    {"key": "venue", "label": "Venue"},
]
WS_FIELDS = [
    {"label": "Semester", "key": "semester"},
    {"label": "Course Code", "key": "course_code"},
    {"label": "Group Code", "key": "group_code"},
    {"label": "Day", "key": "day"},
    {"label": "Start Time", "key": "start_time"},
    {"label": "End Time", "key": "end_time"},
    {"label": "Venue", "key": "venue"},
]


def workshop_list(request):
    qs = WorkshopAllocation.objects.select_related("semester").all()
    q = request.GET.get("q", "")
    sem = request.GET.get("semester", "")
    qs = _search(qs, q, ["course_code", "group_code", "venue"])
    if sem:
        qs = qs.filter(semester_id=sem)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": WS_COLS,
        "detail_fields": WS_FIELDS,
        "q": q,
        NAV: "workshops",
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
        NAV: "workshops",
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
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
    qs = _search(qs, q, ["course_code", "group_code", "venue"])
    if sem:
        qs = qs.filter(semester_id=sem)
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": TD_COLS,
        "detail_fields": TD_FIELDS,
        "q": q,
        NAV: "td",
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
        NAV: "td",
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
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
]
PC_FIELDS = [
    {"label": "Programme", "key": "programme"},
    {"label": "Course Code", "key": "course_code"},
]


def course_list(request):
    qs = ProgrammeCourse.objects.select_related("programme").all()
    q = request.GET.get("q", "")
    qs = _search(qs, q, ["course_code", "programme__code", "programme__name"])
    items, page, pages, total = _paginate(request, qs)
    ctx = {
        "items": items,
        "columns": PC_COLS,
        "detail_fields": PC_FIELDS,
        "q": q,
        NAV: "courses",
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
        NAV: "courses",
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table"
            return r
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
            form.save()
            r = HttpResponse("")
            r["HX-Trigger"] = "close-modal,refresh-table,refresh-detail"
            return r
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
        item.delete()
        r = HttpResponse("")
        r["HX-Trigger"] = "close-modal,refresh-table"
        return r
    return render(
        request,
        "core/delete.html",
        {"item": item, "title": f"Delete {item}?", "back_url": "/courses/", "delete_url": f"/courses/{pk}/delete/"},
    )


# ──────────────────────────────────────────────
# Import Views
# ──────────────────────────────────────────────

IMPORT_TYPES = {
    "semesters": {
        "title": "Semesters",
        "columns": "academic_year, semester",
        "fn": import_semesters_from_excel,
    },
    "programmes": {
        "title": "Programmes",
        "columns": "code, name",
        "fn": import_programmes_from_excel,
    },
    "student-groups": {
        "title": "Student Groups",
        "columns": "programme_code, group_code",
        "fn": import_student_groups_from_excel,
    },
    "programme-courses": {
        "title": "Programme Courses",
        "columns": "programme_code, course_code",
        "fn": import_programme_courses_from_excel,
    },
    "venues": {
        "title": "Venues",
        "columns": "name, capacity",
        "fn": import_venues_from_excel,
    },
    "master-timetable": {
        "title": "Master Timetable",
        "columns": "course_code, activity_type, day, start_time, end_time, venue, group",
        "fn": import_master_timetable_from_excel,
    },
    "workshop-allocation": {
        "title": "Workshop Allocation",
        "columns": "course_code, group_code, day, start_time, end_time, venue",
        "fn": import_workshop_allocation_from_excel,
    },
    "td-allocation": {
        "title": "TD Allocation",
        "columns": "course_code, group_code, day, start_time, end_time, venue",
        "fn": import_td_allocation_from_excel,
    },
}


def import_hub(request):
    ctx = {
        NAV: "imports",
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
        if form.is_valid():
            uploaded = form.cleaned_data["file"]
            try:
                result: ImportResult = info["fn"](uploaded.temporary_file_path())
            except Exception as exc:
                result = ImportResult()
                result.errors.append(str(exc))
            ctx = {
                "result": result,
                "import_type": import_type,
                "import_title": info["title"],
                "columns": info["columns"],
            }
            return render(request, "core/import_result.html", ctx)
    form = FileUploadForm()
    return render(
        request,
        "core/import_upload.html",
        {
            "form": form,
            "import_type": import_type,
            "import_title": info["title"],
            "columns": info["columns"],
            NAV: "imports",
            "page_title": f"Import {info['title']}",
        },
    )
