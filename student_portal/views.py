import uuid

from django.conf import settings
from django.core.cache import cache
from django.db.models import Q
from django.http import HttpResponse, HttpResponseBadRequest
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from core.models import ActivityLog, LogAction, Programme, Semester, StudentGroup
from core.timetable_pdf import (
    collect_master_entries,
    render_group_timetable,
    render_programme_timetable,
    render_udsm_master_timetable,
)
from .forms import CollisionReportForm
from .models import CollisionReport, PortalSettings


def current_semester():
    """The semester the portal shows when the reader picks nothing.

    Three sources, in order: the portal's own override, the app-wide current
    semester staff set on /semesters/, and finally the newest semester. The
    override exists because the portal can be pointed at a different term than
    the staff tools, and it is a required field, so a settings row always names
    a semester -- creating one is a deliberate act, not an accident.
    """
    configured = PortalSettings.objects.select_related("current_semester").first()
    if configured:
        return configured.current_semester
    return Semester.current() or Semester.objects.order_by(
        "-academic_year", "-semester"
    ).first()


def student_home(request):
    semesters = Semester.objects.order_by("-academic_year", "-semester")
    programmes = Programme.objects.prefetch_related("student_groups").order_by("code")
    return render(request, "student_portal/home.html", {
        "semesters": semesters,
        "programmes": programmes,
        "current_semester": current_semester(),
        "year_options": range(1, 5),
    })


def student_timetable_pdf(request):
    scope = request.GET.get("scope", "group")
    # The combined master download always uses the configured current term.
    # Individual group/programme downloads can select a previous semester.
    semester_id = request.GET.get("semester") if scope != "all" else None
    semester = get_object_or_404(Semester, pk=semester_id) if semester_id else current_semester()
    if semester is None:
        return HttpResponseBadRequest("No semester has been configured yet.")
    try:
        year = max(1, min(4, int(request.GET.get("year", 1))))
    except (TypeError, ValueError):
        year = 1

    response = HttpResponse(content_type="application/pdf")
    if scope == "group":
        group_id = request.GET.get("group")
        programme = get_object_or_404(Programme, pk=request.GET.get("programme"))
        if group_id == "all":
            if not programme.student_groups.exists():
                return HttpResponseBadRequest("No student groups are registered for this programme yet.")
            response["Content-Disposition"] = (
                f'attachment; filename="timetable_{programme.code}_{year}.pdf"'
            )
            render_programme_timetable(programme, semester, year, out=response)
        else:
            group = get_object_or_404(
                StudentGroup.objects.select_related("programme"),
                pk=group_id,
                programme=programme,
            )
            response["Content-Disposition"] = (
                f'attachment; filename="timetable_{group.programme.code}_{group.code}.pdf"'
            )
            render_group_timetable(group, semester, year, out=response)
    elif scope == "programme":
        programme = get_object_or_404(Programme, pk=request.GET.get("programme"))
        response["Content-Disposition"] = (
            f'attachment; filename="timetable_{programme.code}_{year}.pdf"'
        )
        render_programme_timetable(programme, semester, year, out=response)
    elif scope == "all":
        response["Content-Disposition"] = (
            f'attachment; filename="master_timetable_{semester.academic_year}_semester_{semester.semester}.pdf"'
        )
        render_udsm_master_timetable(
            collect_master_entries(semester, year=None), semester, 1, out=response
        )
    else:
        return HttpResponseBadRequest("Choose a valid timetable scope.")
    return response


def collision_report(request):
    default = current_semester()
    if request.method == "POST":
        # Lightweight abuse control for the anonymous submission endpoint.
        address = request.META.get("REMOTE_ADDR", "unknown")
        key = f"collision-report:{address}"
        cache.add(key, 0, 60 * 30)
        used = cache.incr(key)
        if used > 5:
            return render(request, "student_portal/report.html", {
                "form": CollisionReportForm(request.POST, default_semester=default),
                "rate_limited": True,
            }, status=429)
        form = CollisionReportForm(request.POST, default_semester=default)
        if form.is_valid():
            report = form.save()
            return redirect("collision-report-thanks", reference=report.reference)
    else:
        form = CollisionReportForm(default_semester=default)
    return render(request, "student_portal/report.html", {"form": form})


def collision_report_thanks(request, reference):
    report = get_object_or_404(CollisionReport, reference=reference)
    return render(request, "student_portal/report_thanks.html", {"report": report})


def collision_report_list(request):
    reports = CollisionReport.objects.select_related("semester", "programme", "group")
    status = request.GET.get("status", "OPEN")
    if status == "OPEN":
        reports = reports.exclude(status=CollisionReport.Status.RESOLVED)
    elif status in dict(CollisionReport.Status.choices):
        reports = reports.filter(status=status)
    else:
        status = "OPEN"
        reports = reports.exclude(status=CollisionReport.Status.RESOLVED)
    q = request.GET.get("q", "").strip()
    if q:
        search = (
            Q(course_or_exam__icontains=q) | Q(description__icontains=q)
            | Q(group__code__icontains=q) | Q(programme__code__icontains=q)
        )
        try:
            search |= Q(reference=uuid.UUID(q))
        except (ValueError, TypeError, AttributeError):
            pass
        reports = reports.filter(search)
    counts = {
        "open": CollisionReport.objects.exclude(status=CollisionReport.Status.RESOLVED).count(),
        "new": CollisionReport.objects.filter(status=CollisionReport.Status.NEW).count(),
        "reviewing": CollisionReport.objects.filter(status=CollisionReport.Status.REVIEWING).count(),
        "resolved": CollisionReport.objects.filter(status=CollisionReport.Status.RESOLVED).count(),
    }
    return render(request, "student_portal/collision_list.html", {
        "reports": reports, "status_filter": status, "q": q, "counts": counts,
    })


def collision_report_detail(request, reference):
    report = get_object_or_404(
        CollisionReport.objects.select_related("semester", "programme", "group"),
        reference=reference,
    )
    if request.method == "POST":
        status = request.POST.get("status", "")
        notes = request.POST.get("staff_notes", "").strip()
        valid_statuses = dict(CollisionReport.Status.choices)
        if status not in valid_statuses:
            return render(request, "student_portal/collision_detail.html", {
                "report": report, "status_choices": CollisionReport.Status.choices,
                "error": "Choose a valid report status.",
            }, status=400)
        report.status = status
        report.staff_notes = notes
        report.save(update_fields=("status", "staff_notes", "updated_at"))
        ActivityLog.objects.create(
            action=LogAction.UPDATE,
            message=f"Updated student clash report {report.reference} to {report.get_status_display()}.",
            resource="Clash report", target=str(report.reference),
        )
        return redirect("collision-report-detail", reference=report.reference)
    return render(request, "student_portal/collision_detail.html", {
        "report": report, "status_choices": CollisionReport.Status.choices,
    })
