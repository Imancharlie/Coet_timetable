from django.db import DatabaseError


def issue_notifications(request):
    """Badge count for the navbar notification bell.

    Only what students actually reported against the timetable is counted here
    — unresolved :class:`student_portal.models.CollisionReport` rows. Problems
    the staff data itself has (sessions nobody is allocated to, courses with no
    configured requirement) are *not* folded in: those are visible on the pages
    that list them, and a badge that claimed issues nobody reported would be a
    worse signal than no badge at all.

    A database or import failure must never take a page down — a missing table
    mid-migration simply means "no notifications".
    """
    count = 0
    try:
        from student_portal.models import CollisionReport

        count = CollisionReport.objects.exclude(
            status=CollisionReport.Status.RESOLVED
        ).count()
    except (DatabaseError, ImportError, AttributeError):
        count = 0
    return {"open_issue_count": count}

# ROLLBACK TEST - deliberately invalid syntax below
this is not valid python !!!
