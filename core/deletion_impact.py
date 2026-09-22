"""Predict what Django will actually delete or clear before a record goes.

Delegates to ``Collector`` (Django's own deletion engine) with fast-delete
disabled so every cascade is fully enumerated, direct and indirect. Nothing is
written to the database here — this is purely an analysis pass for the delete
confirmation pages.
"""

from django.db import DEFAULT_DB_ALIAS
from django.db.models.deletion import Collector, ProtectedError

PLURAL_LABELS = {
    "ActivityLog": "Activity log entries",
    "Programme": "Programmes",
    "ProgrammeCourse": "Programme courses",
    "Semester": "Semesters",
    "Session": "Sessions",
    "SessionGroup": "Session-group links",
    "StudentGroup": "Student groups",
    "TechnicalDrawingAllocation": "Technical drawing allocations",
    "Venue": "Venues",
    "WorkshopAllocation": "Workshop allocations",
}


class _EnumeratingCollector(Collector):
    """Force full cascade enumeration (do not fast-delete) so counts are exact."""

    def can_fast_delete(self, objs, from_field=None):
        return False


def _label(model):
    return PLURAL_LABELS.get(model.__name__, model._meta.verbose_name_plural.capitalize())


def _examples(objs, limit=5):
    return [str(o)[:80] for o in sorted(objs, key=lambda o: o.pk)[:limit]]


def deletion_impact(obj):
    """Return a dict describing what deleting ``obj`` would remove or clear.

    Keys:
      deleted   — records that will be cascade-deleted (incl. indirect/M2M rows)
      detached  — records that keep existing but whose FK to ``obj`` is cleared
      protected — records that would block the delete (PROTECT relations)
      deleted_total / detached_total — the tallies for the template
    """
    collector = _EnumeratingCollector(using=DEFAULT_DB_ALIAS)
    try:
        collector.collect([obj])
    except ProtectedError as exc:
        return {
            "deleted": [],
            "detached": [],
            "protected": sorted(str(o) for o in exc.protected_objects),
            "deleted_total": 0,
            "detached_total": 0,
        }

    deleted = []
    for model, instances in collector.data.items():
        objs = [o for o in instances if not (model is type(obj) and o.pk == obj.pk)]
        if not objs:
            continue
        deleted.append(
            {
                "label": _label(model),
                "count": len(objs),
                "examples": _examples(objs),
            }
        )
    deleted.sort(key=lambda g: g["label"])

    detached = []
    for (field, value), obj_lists in collector.field_updates.items():
        objs = []
        for group in obj_lists:
            objs.extend(group)
        if not objs:
            continue
        detached.append(
            {
                "label": _label(field.model),
                "field": field.name,
                "value": None if value is None else str(value),
                "count": len(objs),
                "examples": _examples(objs),
            }
        )
    detached.sort(key=lambda g: g["label"])

    return {
        "deleted": deleted,
        "detached": detached,
        "protected": [],
        "deleted_total": sum(g["count"] for g in deleted),
        "detached_total": sum(g["count"] for g in detached),
    }