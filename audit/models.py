"""Tables owned by the audit app.

Two, both of them outside everything the allocator writes to:

* :class:`AllocationRunMetrics` — additive instrumentation. The allocator
  never writes it; a signal reads the run's own plan snapshot and copies the
  performance counters out. Historical runs with no snapshot read as "Not
  recorded" rather than as zero.
* :class:`AuditReport` — one row per generated report *version*. Re-auditing
  the same run with different thresholds adds a row; it never overwrites one.
"""

from __future__ import annotations

import json

from django.conf import settings
from django.db import models
from django.utils import timezone

from audit.version import APP_VERSION, PROJECT_NAME


class ReportStatus(models.TextChoices):
    QUEUED = "QUEUED", "Queued"
    RUNNING = "RUNNING", "Running"
    COMPLETE = "COMPLETE", "Complete"
    FAILED = "FAILED", "Failed"


class AllocationRunMetrics(models.Model):
    """Performance counters for one allocation run, copied out read-only.

    The allocator's plan already records execution time and how many candidate
    placements it examined; this table gives the audit a stable, queryable place
    to read them from, and a home for counters a future allocator revision may
    add. Nothing here is ever *used* to make an allocation decision.
    """

    run = models.OneToOneField(
        "core.AllocationRun",
        on_delete=models.CASCADE,
        related_name="metrics",
    )
    duration_s = models.FloatField(null=True, blank=True)
    candidates_examined = models.IntegerField(null=True, blank=True)
    backtracks = models.IntegerField(null=True, blank=True)
    requirements_processed = models.IntegerField(null=True, blank=True)
    search_limit_hit = models.BooleanField(null=True, blank=True)
    recorded_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-recorded_at"]
        verbose_name = "Allocation run metrics"

    def __str__(self):
        return f"Metrics for run {self.run_id}"

    @property
    def throughput(self):
        """Requirements processed per second, or ``None`` if unrecorded."""
        if not self.duration_s or self.requirements_processed is None:
            return None
        if self.duration_s <= 0:
            return None
        return round(self.requirements_processed / self.duration_s, 2)


class AuditReportQuerySet(models.QuerySet):
    def complete(self):
        return self.filter(status=ReportStatus.COMPLETE)

    def for_semester(self, semester):
        return self.filter(allocation_run__semester=semester)


class AuditReport(models.Model):
    """One generated audit report — a version, not a mutable document."""

    allocation_run = models.ForeignKey(
        "core.AllocationRun",
        on_delete=models.CASCADE,
        related_name="audit_reports",
    )
    version = models.PositiveIntegerField(default=1)
    created_at = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="audit_reports",
    )
    app_version = models.CharField(max_length=40, default=APP_VERSION)
    thresholds_json = models.TextField(blank=True)
    summary_json = models.TextField(blank=True)
    data_hash = models.CharField(max_length=64, blank=True, db_index=True)
    pdf_path = models.CharField(max_length=400, blank=True)
    status = models.CharField(
        max_length=10, choices=ReportStatus.choices, default=ReportStatus.QUEUED
    )
    progress = models.CharField(max_length=200, blank=True)
    error = models.TextField(blank=True)
    page_count = models.IntegerField(null=True, blank=True)
    file_size = models.IntegerField(null=True, blank=True)
    generation_seconds = models.FloatField(null=True, blank=True)
    comparison_to = models.ForeignKey(
        "self",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="newer_reports",
    )
    config_path = models.CharField(max_length=400, blank=True)

    objects = AuditReportQuerySet.as_manager()

    class Meta:
        ordering = ["-created_at"]
        verbose_name = "Audit report"
        verbose_name_plural = "Audit reports"
        constraints = [
            models.UniqueConstraint(
                fields=["allocation_run", "version"], name="audit_report_version_unique"
            )
        ]
        indexes = [
            models.Index(fields=["allocation_run", "-version"]),
            models.Index(fields=["status"]),
        ]

    def __str__(self):
        return f"Audit report v{self.version} for run {self.allocation_run_id}"

    # -- state -------------------------------------------------------------
    @property
    def is_complete(self) -> bool:
        return self.status == ReportStatus.COMPLETE

    @property
    def is_running(self) -> bool:
        return self.status in {ReportStatus.QUEUED, ReportStatus.RUNNING}

    def mark_running(self, stage: str) -> None:
        self.status = ReportStatus.RUNNING
        self.progress = stage
        self.error = ""
        self.save(update_fields=["status", "progress", "error"])

    def mark_complete(self, stage: str, **fields) -> None:
        self.status = ReportStatus.COMPLETE
        self.progress = stage
        for key, value in fields.items():
            setattr(self, key, value)
        self.save(
            update_fields=["status", "progress", "error", "pdf_path", "page_count",
                           "file_size", "generation_seconds", "summary_json",
                           "thresholds_json", "data_hash", "comparison_to"]
        )

    def mark_failed(self, error: str) -> None:
        self.status = ReportStatus.FAILED
        self.progress = "Failed"
        self.error = str(error)[:4000]
        self.save(update_fields=["status", "progress", "error"])

    # -- stored payloads ---------------------------------------------------
    def summary(self) -> dict:
        try:
            return json.loads(self.summary_json or "{}")
        except (ValueError, TypeError):
            return {}

    def thresholds(self) -> dict:
        try:
            return json.loads(self.thresholds_json or "{}")
        except (ValueError, TypeError):
            return {}

    def next_version(self) -> int:
        latest = (
            AuditReport.objects.filter(allocation_run=self.allocation_run)
            .order_by("-version")
            .values_list("version", flat=True)
            .first()
        )
        return (latest or 0) + 1

    @property
    def title(self) -> str:
        return f"Post-Allocation Audit · {self.allocation_run.semester}"

    @property
    def document_title(self) -> str:
        return (
            f"Post-Allocation Audit Report — {self.allocation_run.semester} "
            f"(run #{self.allocation_run_id} v{self.version})"
        )

    @property
    def filename(self) -> str:
        run = self.allocation_run
        semester_slug = (
            f"{run.semester.academic_year}-S{run.semester.semester}"
        ).replace("/", "-")
        return (
            f"audit-report-run{self.allocation_run_id}-v{self.version}-"
            f"{semester_slug}.pdf"
        )

    def read_pdf(self) -> bytes | None:
        from pathlib import Path

        if not self.pdf_path:
            return None
        path = Path(self.pdf_path)
        if not path.exists():
            return None
        return path.read_bytes()

    def previous_complete(self):
        """The report this one is compared against in section 8."""
        if self.comparison_to_id:
            return self.comparison_to
        return (
            AuditReport.objects.filter(
                allocation_run__semester=self.allocation_run.semester,
                status=ReportStatus.COMPLETE,
            )
            .exclude(pk=self.pk)
            .order_by("-created_at")
            .first()
        )

    @property
    def created_label(self) -> str:
        stamp = self.created_at or timezone.now()
        return stamp.strftime("%d %b %Y, %H:%M")
