"""Turn an applied allocation run into a stored, versioned audit report.

The whole generation is one function, :func:`generate_report`, and it is
deliberately *synchronous*. A report is a few hundred milliseconds of number
crunching plus one Chrome invocation; a task queue would add a dependency and a
second state machine to explain for no benefit. The status column still carries
QUEUED/RUNNING/COMPLETE/FAILED because the model is the API, and a failed render
must be visible rather than retried invisibly.

Ordering is fixed and each step records itself on the row, so a report that
fails halfway says which stage failed:

1. collect (read-only)  2. analyse  3. recommend  4. store payload
5. render HTML  6. print PDF  7. mark complete
"""

from __future__ import annotations

import time
from pathlib import Path

from django.utils import timezone

from audit.analytics import analyse
from audit.collect import collect
from audit.config import default_config
from audit.render import build_payload, html_report, html_to_pdf, pdf_dir
from audit.version import APP_VERSION


def generate_report(
    run,
    *,
    user=None,
    config=None,
    config_path: str = "",
    write_pdf: bool = True,
    comparison_to=None,
) -> "object":
    """Build one report version for *run* and return the ``AuditReport`` row.

    Never raises for an expected failure: the row is marked FAILED with the
    reason and returned, because a coordinator pressing a button must see "the
    PDF could not be produced and here is why", not a 500.
    """
    from audit.models import AuditReport, ReportStatus

    started = time.monotonic()
    report = AuditReport.objects.create(
        allocation_run=run,
        version=report_version_for(run),
        created_by=user if (user is not None and getattr(user, "is_authenticated", False)) else None,
        app_version=APP_VERSION,
        status=ReportStatus.RUNNING,
        progress="Collecting timetable data",
        config_path=config_path,
    )

    try:
        cfg = config or default_config()
        report.mark_running("Collecting timetable data")
        dataset = collect(run, cfg)

        report.mark_running("Analysing")
        analysis = analyse(dataset)

        report.mark_running("Building recommendations")
        from audit.recommend import build as build_recommendations

        recommendations = build_recommendations(analysis, cfg)
        payload = build_payload(analysis, recommendations, run=run)

        report.mark_running("Saving report")
        report.summary_json = _dump(payload)
        report.thresholds_json = _dump(_thresholds(cfg))
        report.data_hash = payload.get("data_hash", "")
        report.comparison_to = comparison_to
        report.save(
            update_fields=[
                "summary_json",
                "thresholds_json",
                "data_hash",
                "comparison_to",
            ]
        )

        html_text = html_report(payload, report.document_title)

        if not write_pdf:
            report.mark_complete(
                "Complete (HTML only)",
                generation_seconds=round(time.monotonic() - started, 3),
            )
            return report

        report.mark_running("Printing PDF")
        target = pdf_dir() / report.filename
        html_to_pdf(html_text, target)
        size = target.stat().st_size

        report.mark_complete(
            "Complete",
            pdf_path=str(target),
            file_size=size,
            page_count=_page_count(target),
            generation_seconds=round(time.monotonic() - started, 3),
        )
        return report

    except Exception as exc:  # noqa: BLE001 - recorded, not swallowed
        report.mark_failed(f"{type(exc).__name__}: {exc}")
        report.generation_seconds = round(time.monotonic() - started, 3)
        report.save(update_fields=["generation_seconds"])
        return report


def report_version_for(run) -> int:
    from audit.models import AuditReport

    latest = (
        AuditReport.objects.filter(allocation_run=run)
        .order_by("-version")
        .values_list("version", flat=True)
        .first()
    )
    return (latest or 0) + 1


def _page_count(pdf_path: Path):
    """Page count from the PDF's own ``/Type /Page`` objects, or None.

    pypdf is the honest way to do this, but it is an optional extra, and a
    missing page count must not fail a report that rendered perfectly well.
    """
    try:
        from pypdf import PdfReader

        return len(PdfReader(str(pdf_path)).pages)
    except Exception:  # noqa: BLE001
        return None


def _dump(value) -> str:
    import json

    return json.dumps(value, default=str)


def _thresholds(cfg) -> dict:
    """The threshold subtree only, so a report can say what it judged against."""
    from audit.config import as_plain_dict

    data = as_plain_dict(cfg)
    keep = ("threshold", "working_day", "group_students", "analysis")
    return {k: v for k, v in data.items() if any(k.startswith(p) for p in keep)}


def latest_report_for(run):
    """The newest report of any status, or None."""
    from audit.models import AuditReport

    return (
        AuditReport.objects.filter(allocation_run=run)
        .order_by("-version")
        .first()
    )


def has_report(run) -> bool:
    from audit.models import AuditReport

    return AuditReport.objects.filter(allocation_run=run).exists()


def recent_runs(limit: int = 20):
    """Applied runs newest-first, for the 'audit this run' picker."""
    from core.models import AllocationRun

    qs = AllocationRun.objects.select_related("semester").order_by("-created_at")
    return list(qs[:limit])


def applied_runs_for_semester(semester):
    """Applied runs for one semester, newest first."""
    from core.models import AllocationRun, AllocationStatus

    return AllocationRun.objects.filter(
        semester=semester, status=AllocationStatus.APPLIED
    ).order_by("-created_at")
