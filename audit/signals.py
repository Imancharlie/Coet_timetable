"""Additive instrumentation for the allocator's run records.

A single ``post_save`` listener on :class:`core.models.AllocationRun` copies
performance counters out of the run's own plan snapshot into
:mod:`audit`'s ``AllocationRunMetrics`` row.

It is a *listener*, not a hook the allocator calls: the allocator code is not
edited, does not import anything from ``audit``, and cannot fail because of
anything here. If the snapshot is missing or unreadable the listener simply
records ``None``, which the report prints as "Not recorded" — a run that
predates this table is never reported as a run that took zero seconds.
"""

from django.db.models.signals import post_save
from django.dispatch import receiver


@receiver(post_save, sender="core.AllocationRun")
def record_run_metrics(sender, instance, **kwargs):
    from audit.models import AllocationRunMetrics

    snapshot = instance.plan() or {}

    def number(value):
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    duration_s = None
    if snapshot.get("duration_ms"):
        try:
            duration_s = round(float(snapshot["duration_ms"]) / 1000, 3)
        except (TypeError, ValueError):
            duration_s = None

    search_limit = snapshot.get("search_limit_hit")
    if search_limit is None:
        search_limit = bool(instance.search_limit_hit) or None

    AllocationRunMetrics.objects.update_or_create(
        run=instance,
        defaults={
            "duration_s": duration_s,
            "candidates_examined": number(snapshot.get("scanned")),
            "backtracks": number(snapshot.get("backtracks")),
            "requirements_processed": number(snapshot.get("requirement_total")),
            "search_limit_hit": search_limit,
        },
    )
