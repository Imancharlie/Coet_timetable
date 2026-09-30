"""Pages for reading, listing and exporting audit reports.

Four entry points, and the reason there are four is worth stating: a report is
generated *from* an allocation run, read *as* a version, listed *across* runs,
and downloaded as a file. Collapsing any two of those makes one of the
coordinator's actual questions unanswerable.

Nothing here writes to the timetable. The single write path in this app is
:func:`run_audit`, which writes an ``AuditReport`` row and nothing else.
"""

from __future__ import annotations

import json

from django.contrib.auth.decorators import login_required
from django.http import Http404, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.http import require_POST

from core.models import AllocationRun, LogAction
from core.views import _log
from audit.generator import (
    applied_runs_for_semester,
    generate_report,
    latest_report_for,
    recent_runs,
)
from audit.models import AuditReport, ReportStatus
from audit.render import PAYLOAD_VERSION


def _semester_from_request(request):
    from core.models import Semester

    raw = request.GET.get("semester", "")
    if raw:
        found = Semester.objects.filter(pk=raw).first()
        if found:
            return found
    return Semester.current() or Semester.objects.order_by("-pk").first()


def _next(request):
    target = request.GET.get("next") or ""
    if target and url_has_allowed_host_and_scheme(
        target, allowed_hosts={request.get_host()}, require_https=False
    ):
        return target
    return ""


# ---------------------------------------------------------------------------
# generate
# ---------------------------------------------------------------------------


@login_required
@require_POST
def run_audit(request, run_pk):
    """Generate a new report version for *run* and open it.

    POST only: this writes a row and burns a Chrome process, so it must never be
    reachable by a link, a prefetch or a crawler.
    """
    run = get_object_or_404(
        AllocationRun.objects.select_related("semester"), pk=run_pk
    )
    report = generate_report(run, user=request.user)

    if report.status == ReportStatus.FAILED:
        _log(
            LogAction.UPDATE,
            f"Audit report for run {run.pk} failed: {report.error[:200]}",
            "Audit Report",
            str(run),
        )
    else:
        _log(
            LogAction.CREATE,
            f"Audit report v{report.version} generated for run {run.pk}",
            "Audit Report",
            str(run),
        )

    if request.headers.get("HX-Request"):
        response = HttpResponse("")
        response["HX-Redirect"] = (
            f"{reverse('audit-report-detail', args=[report.pk])}"
            f"?run={run.pk}"
        )
        return response

    return redirect(f"{reverse('audit-report-detail', args=[report.pk])}?run={run.pk}")


# ---------------------------------------------------------------------------
# read
# ---------------------------------------------------------------------------


@login_required
def report_detail(request, pk):
    """One report version, read from its stored payload.

    The page renders ``summary_json``, not a fresh analysis. A report is a
    version of what was said at a moment; re-analysing on every view would make
    an old report quietly change under the reader and make the comparison in
    section 8 meaningless.
    """
    report = get_object_or_404(
        AuditReport.objects.select_related(
            "allocation_run", "allocation_run__semester", "created_by"
        ),
        pk=pk,
    )
    payload = report.summary()

    if payload.get("payload_version") != PAYLOAD_VERSION:
        # The stored payload predates the current shape, so the templates that
        # render it can no longer be trusted to read it — and a page that raises
        # is worse than one that is a version behind. The data is still there
        # and the analysis is deterministic, so regenerate into a NEW version
        # and send the reader there. The old row is kept: a version is a record
        # of what was said, not a cache slot to be overwritten.
        fresh = generate_report(report.allocation_run, user=request.user)
        if fresh.status != ReportStatus.FAILED:
            _log(
                LogAction.CREATE,
                f"Audit report v{report.version} was stored in an older payload "
                f"format; regenerated as v{fresh.version}",
                "Audit Report",
                str(report.allocation_run),
            )
            return redirect(
                f"{reverse('audit-report-detail', args=[fresh.pk])}?regenerated=1"
            )
        _log(
            LogAction.UPDATE,
            f"Could not re-render audit report v{report.version}: {fresh.error[:200]}",
            "Audit Report",
            str(report.allocation_run),
        )
        payload = {}

    # `day_pressure` is a list of rows (day / student_hours / sessions), not a
    # dict keyed by day: the payload builder flattens the rollup.
    day_pressure = payload.get("day_pressure") or []
    # The bar widths need a denominator, and a template cannot take a max() over
    # a list. A zero peak would also divide by zero in widthratio, hence the 1.
    day_peak = payload.get("day_peak") or max(
        (d.get("student_hours") or 0 for d in day_pressure), default=0
    ) or 1
    return render(
        request,
        "audit/report_detail.html",
        {
            "report": report,
            "payload": payload,
            "kpis": payload.get("kpis") or {},
            "recommendations": payload.get("recommendations") or [],
            "issues": payload.get("issues") or [],
            "groups": payload.get("groups") or [],
            "venues": payload.get("venues") or [],
            "day_pressure": day_pressure,
            "day_peak": day_peak,
            "verdict": payload.get("verdict") or {},
            "counts": payload.get("counts") or {},
            "unscored": payload.get("unscored") or [],
            "observations": payload.get("observations") or [],
            "previous": report.previous_complete(),
            "newer": (
                report.allocation_run.audit_reports.filter(
                    version__gt=report.version
                )
                .order_by("version")
                .first()
            ),
        },
    )


# ---------------------------------------------------------------------------
# list
# ---------------------------------------------------------------------------


@login_required
def report_index(request):
    """Every report ever generated, newest first, filterable by semester.

    This is the "which allocation was that?" page. It lists *versions* rather
    than runs on purpose: re-auditing a run after changing a threshold is a
    legitimate thing to want to see, and two rows for one run is the honest
    representation of that.
    """
    reports = AuditReport.objects.select_related(
        "allocation_run", "allocation_run__semester", "created_by"
    )
    semester = _semester_from_request(request)
    if semester and request.GET.get("semester"):
        reports = reports.filter(allocation_run__semester=semester)
    if request.GET.get("status"):
        reports = reports.filter(status=request.GET["status"])

    context = {
        "reports": reports[:200],
        "semester": semester,
        "semesters": _semester_choices(),
        "selected_status": request.GET.get("status", ""),
        "counts": _status_counts(),
        "recent_runs": recent_runs(10),
    }
    return render(request, "audit/report_index.html", context)


def _semester_choices():
    from core.models import Semester

    return Semester.objects.all().order_by("-academic_year", "-semester")


def _status_counts():
    from django.db.models import Count

    rows = AuditReport.objects.values_list("status").annotate(n=Count("pk"))
    return {status: n for status, n in rows}


# ---------------------------------------------------------------------------
# export
# ---------------------------------------------------------------------------


@login_required
def report_pdf(request, pk):
    """Download the stored PDF.

    ``Content-Disposition: attachment`` **and** a ``download`` attribute on the
    link that points here. The attribute is the part that matters: every export
    in this project carries it, because a download navigation never fires
    ``load``/``pageshow`` to disarm the page loader, and a missing attribute
    leaves the overlay spinning forever after the click.
    """
    report = get_object_or_404(AuditReport, pk=pk)
    payload = report.read_pdf()
    if payload is None:
        if report.status != ReportStatus.COMPLETE:
            raise Http404(f"Report v{report.version} is {report.status}.")
        # Complete but the file is gone from disk: regenerate rather than 404,
        # because the row is the record of a report that did exist.
        fresh = generate_report(report.allocation_run, user=request.user)
        payload = fresh.read_pdf()
        if payload is None:
            raise Http404(f"Report v{report.version} has no PDF on disk.")
        report = fresh

    response = HttpResponse(payload, content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="{report.filename}"'
    response["Content-Length"] = str(len(payload))
    return response


@login_required
def report_json(request, pk):
    """The stored payload as JSON — for anyone who wants the numbers."""
    report = get_object_or_404(AuditReport, pk=pk)
    return JsonResponse(
        {
            "run": report.allocation_run_id,
            "version": report.version,
            "data_hash": report.data_hash,
            "created_at": report.created_at,
            "summary": report.summary(),
        }
    )
