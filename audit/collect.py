"""Normalise the database into the audit's own input structures.

Everything downstream of this module works on frozen dataclasses, never on
Django models. That is what makes the analytics testable without a database
and what makes the data hash in Appendix D reproducible: the same applied
timetable always produces the same tuples.

The mapping is strictly read-only. A few deliberate choices:

* **Sessions come from the live ``SessionGroup`` links**, not from the run's
  plan snapshot. The audit reports what students are *actually* timetabled,
  so a link that was added or removed by hand after the run is reflected.
* **Requirements come from the run's plan snapshot** (which recorded the
  requirement list at the time), but each one is re-checked against the live
  links, so a requirement whose assignment has since been removed is reported
  unresolved rather than quietly counted as allocated.
* **Origin** (retained / new / reassigned) only exists for runs that recorded
  a plan. Anything else is ``"unknown"`` and is printed as "Not recorded".
"""

from __future__ import annotations

import hashlib
import json
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from typing import Optional

from core.models import Day, normalise_course_code
from django.db.models import Q

#: ``Day`` TextChoices order is Monday..Sunday, which is exactly the 0..6
#: weekday index the spec uses, so this is a lookup rather than a guess.
DAY_INDEX = {value: index for index, value in enumerate(Day.values)}
DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_SHORT = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

ORIGIN_RETAINED = "retained"
ORIGIN_NEW = "new"
ORIGIN_REASSIGNED = "reassigned"
ORIGIN_UNKNOWN = "unknown"
ORIGIN_NA = "n/a"

_STATUS_TO_ORIGIN = {
    "added": ORIGIN_NEW,
    "retained": ORIGIN_RETAINED,
    "moved": ORIGIN_REASSIGNED,
}


@dataclass(frozen=True)
class Session:
    id: str
    course: str
    activity: str            # 'seminar' | 'tutorial' | 'practical' | 'lecture' | 'workshop'
    group: str               # e.g. 'EE C1'
    day: int                 # 0=Mon .. 4=Fri (5,6 if weekend used)
    start: int               # minutes from 00:00, e.g. 8:00 -> 480
    end: int
    venue: Optional[str]
    capacity: Optional[int]
    group_size: Optional[int]
    origin: str = ORIGIN_UNKNOWN
    session_pk: Optional[int] = None
    course_name: str = ""
    activity_label: str = ""

    @property
    def minutes(self) -> int:
        return max(self.end - self.start, 0)

    @property
    def hours(self) -> float:
        return round(self.minutes / 60, 4)


@dataclass(frozen=True)
class Requirement:
    course: str
    activity: str
    group: str
    status: str              # 'allocated' | 'unresolved'
    origin: str = ORIGIN_UNKNOWN
    reason: Optional[str] = None
    ordinal: int = 1
    session_pk: Optional[int] = None
    course_name: str = ""

    @property
    def key(self) -> tuple:
        return (self.group, self.course, self.activity)


@dataclass
class RunMeta:
    run_id: str
    semester: str
    academic_year: str
    started_at: Optional[str] = None
    duration_s: Optional[float] = None
    candidates_examined: Optional[int] = None
    backtracks: Optional[int] = None
    search_limit_reached: Optional[bool] = None
    completed: Optional[bool] = None
    applied: Optional[bool] = None
    errors: list = field(default_factory=list)
    # Additive context that is always available from the run row itself.
    run_pk: Optional[int] = None
    semester_pk: Optional[int] = None
    scope: str = "ALL"
    status: str = ""
    algorithm: str = ""
    created_at: Optional[str] = None
    applied_at: Optional[str] = None
    reverted_at: Optional[str] = None
    assigned: int = 0
    moved: int = 0
    retained: int = 0
    removed: int = 0
    unresolved: int = 0
    requirement_total: int = 0
    warnings: list = field(default_factory=list)
    unconfigured_courses: list = field(default_factory=list)

    @property
    def label(self) -> str:
        return f"Run #{self.run_id}"


@dataclass
class AuditDataset:
    """Everything the analytics need, and nothing they can accidentally write."""

    meta: RunMeta
    sessions: list = field(default_factory=list)
    requirements: list = field(default_factory=list)
    groups: list = field(default_factory=list)          # every group in scope
    scored_groups: list = field(default_factory=list)  # groups that have sessions
    group_programme: dict = field(default_factory=dict)
    group_size: dict = field(default_factory=dict)
    courses_in_scope: list = field(default_factory=list)
    config: object = None
    sessions_by_group: dict = field(default_factory=dict)
    requirements_by_group: dict = field(default_factory=dict)
    unconfigured_courses: list = field(default_factory=list)
    workshop_rows: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    notes: list = field(default_factory=list)

    def data_hash(self) -> str:
        """Stable hash of the analysed inputs (Appendix D lineage)."""
        payload = {
            "run": {
                "run_id": self.meta.run_id,
                "semester": self.meta.semester,
                "academic_year": self.meta.academic_year,
                "scope": self.meta.scope,
                "status": self.meta.status,
            },
            "sessions": sorted(
                [
                    s.id,
                    s.group,
                    s.course,
                    s.activity,
                    s.day,
                    s.start,
                    s.end,
                    s.venue or "",
                    s.capacity if s.capacity is not None else "",
                    s.group_size if s.group_size is not None else "",
                ]
                for s in self.sessions
            ),
            "requirements": sorted(
                [r.group, r.course, r.activity, r.status, r.origin]
                for r in self.requirements
            ),
        }
        blob = json.dumps(payload, sort_keys=True, default=str).encode("utf-8")
        return hashlib.sha256(blob).hexdigest()

    def sessions_for(self, group: str) -> list:
        return self.sessions_by_group.get(group, [])

    def requirements_for(self, group: str) -> list:
        return self.requirements_by_group.get(group, [])


def group_code(programme_code: str, group: str) -> str:
    """``("EE", "C1") -> "EE C1"`` — the same string ``StudentGroup.__str__`` makes."""
    return f"{programme_code} {group}".strip()


def _minutes(value) -> int:
    if value is None:
        return 0
    return int(value.hour) * 60 + int(value.minute)


def programme_of(code: str, pattern) -> str:
    """Read the programme off a group code with the configured regex."""
    import re

    match = re.match(pattern or r"^([A-Z]{2,4})\s", str(code or ""))
    if match:
        return match.group(1)
    return str(code or "").split(" ")[0] or "?"


def collect(run, config, *, groups=None) -> AuditDataset:
    """Read an :class:`core.models.AllocationRun` and its semester into a dataset.

    ``groups`` narrows the audit to specific groups (used by the "audit one
    programme" action and by tests); ``None`` means every group that studies a
    course in play for the semester.
    """
    from core.group_allocation import STUDENTS_PER_GROUP

    semester = run.semester
    snapshot = run.plan() or {}
    meta = _run_meta(run, snapshot)

    pattern = config.get_path("programme_regex") if config else None

    from core.models import SessionGroup, StudentGroup

    # "In scope" = the groups this audit speaks about: every group that has a
    # session-group link in the semester, plus every group whose programme
    # studies a course in it. A group in neither is not part of this timetable.
    group_qs = StudentGroup.objects.filter(
        Q(session_groups__session__semester=semester)
        | Q(programme__programme_courses__semester=semester.semester)
    ).distinct()
    if groups is not None:
        group_qs = group_qs.filter(pk__in=[g.pk if hasattr(g, "pk") else g for g in groups])
    group_rows = list(group_qs.order_by("programme__code", "code"))

    group_programme = {}
    group_size = {}
    for row in group_rows:
        code = group_code(row.programme.code, row.code)
        group_programme[code] = row.programme.code
        group_size[code] = STUDENTS_PER_GROUP

    # Every session-group link is one row of the group's own timetable.
    links = (
        SessionGroup.objects.filter(session__semester=semester)
        .select_related("session", "session__venue", "group", "group__programme")
        .order_by("group__programme__code", "group__code")
    )
    if groups is not None:
        links = links.filter(group__in=group_rows)

    origin_by_link = _origin_lookup(snapshot)

    sessions = []
    for link in links:
        sess = link.session
        code = group_code(link.group.programme.code, link.group.code)
        day_index = DAY_INDEX.get(sess.day)
        if day_index is None:
            continue
        capacity = sess.venue.capacity if sess.venue_id else None
        sessions.append(
            Session(
                id=f"S-{sess.pk}-{code}",
                course=normalise_course_code(sess.course_code),
                course_name="",
                activity=str(sess.activity_type or "").lower(),
                activity_label=str(sess.get_activity_type_display() or ""),
                group=code,
                day=day_index,
                start=_minutes(sess.start_time),
                end=_minutes(sess.end_time),
                venue=sess.venue.name if sess.venue_id else None,
                capacity=capacity if capacity else None,
                group_size=group_size.get(code, STUDENTS_PER_GROUP),
                origin=origin_by_link.get(
                    (code, sess.pk), ORIGIN_UNKNOWN
                ),
                session_pk=sess.pk,
            )
        )

    # Course names come from the shared Course record, for readable tables.
    from core.models import Course

    names = {
        code: name
        for code, name in Course.objects.filter(
            code__in={s.course for s in sessions}
        ).values_list("code", "name")
    }
    sessions = [
        Session(**{**asdict(s), "course_name": names.get(s.course, "")})
        for s in sessions
    ]

    requirements = _collect_requirements(snapshot, sessions, group_programme, names)
    groups_in_scope = sorted({s.group for s in sessions} | set(group_programme))
    workshop_rows = _collect_workshops(semester, _bare_group_index(groups_in_scope))

    dataset = AuditDataset(
        meta=meta,
        sessions=sessions,
        requirements=requirements,
        groups=groups_in_scope,
        group_programme=group_programme,
        group_size=group_size,
        config=config,
        unconfigured_courses=list(snapshot.get("unconfigured_courses") or []),
        workshop_rows=workshop_rows,
        warnings=list(snapshot.get("warnings") or []),
    )

    sessions_by_group: dict = {}
    for sess in sessions:
        sessions_by_group.setdefault(sess.group, []).append(sess)
    for rows in sessions_by_group.values():
        rows.sort(key=lambda s: (s.day, s.start, s.end, s.course))
    dataset.sessions_by_group = sessions_by_group
    dataset.scored_groups = sorted(sessions_by_group)

    requirements_by_group: dict = {}
    for req in requirements:
        requirements_by_group.setdefault(req.group, []).append(req)
    dataset.requirements_by_group = requirements_by_group

    dataset.courses_in_scope = sorted({r.course for r in requirements} | {s.course for s in sessions if s.activity in {"seminar", "tutorial", "practical"}})
    dataset.unconfigured_courses = list(snapshot.get("unconfigured_courses") or [])
    return dataset


def _bare_group_index(codes) -> dict:
    """``"C1" -> ["CE C1", "AR C1", ...]`` for the group codes in scope.

    A workshop names a *bare* group code, and the same code exists under every
    programme, so it is expanded to all of them. That is the meaning the rest of
    the app already gives it -- ``core.group_allocation`` blocks every group
    carrying the code when it checks availability -- and picking one programme
    would report a clash for one cohort and quietly hide it for the other
    thirteen.
    """
    index: dict = defaultdict(list)
    for code in codes:
        parts = str(code).split(" ")
        if len(parts) == 2:
            index[parts[1]].append(code)
    return {bare: sorted(full) for bare, full in index.items()}


def _collect_workshops(semester, by_bare: dict | None = None) -> list:
    """Workshop allocations, kept as plain dicts.

    They are *not* merged into any group's timetable: a workshop is its own
    class with its own room, and the audit reports an unresolvable workshop
    window as unverifiable rather than guessing at it.

    ``group_code`` is kept exactly as the model stores it -- a bare "C1" -- and
    ``groups`` carries the full "CE C1" style codes it stands for, because an
    issue that names "C1" points at nothing the reader can look up.
    """
    from core.models import WorkshopAllocation

    rows = []
    for rec in WorkshopAllocation.objects.filter(semester=semester).order_by("pk"):
        bare = rec.group_code or ""
        matched = (by_bare or {}).get(bare) or []
        rows.append(
            {
                "id": f"W-{rec.pk}",
                "group_code": bare,
                "groups": matched,
                "group": matched[0] if matched else bare,
                "day": rec.day,
                "start": _minutes(rec.start_time) if rec.start_time else None,
                "end": _minutes(rec.end_time) if rec.end_time else None,
                "time_period": rec.time_period,
                "venue": rec.venue,
                "workshop": rec.workshop,
                "week_start": rec.week_start,
                "week_end": rec.week_end,
            }
        )
    return rows


def _run_meta(run, snapshot: dict) -> RunMeta:
    """Assemble :class:`RunMeta`, marking anything the run never recorded."""
    metrics = getattr(run, "metrics", None)
    duration_s = None
    candidates = None
    backtracks = None
    if metrics is not None:
        duration_s = metrics.duration_s
        candidates = metrics.candidates_examined
        backtracks = metrics.backtracks
    if duration_s is None and snapshot.get("duration_ms"):
        duration_s = round(int(snapshot["duration_ms"]) / 1000, 3)
    if candidates is None and snapshot.get("scanned"):
        candidates = int(snapshot["scanned"])

    status = str(getattr(run, "status", "") or "")
    meta = RunMeta(
        run_id=str(run.pk),
        run_pk=run.pk,
        semester=str(run.semester),
        semester_pk=run.semester_id,
        academic_year=str(run.semester.academic_year),
        started_at=run.created_at.isoformat() if run.created_at else None,
        duration_s=duration_s,
        candidates_examined=candidates,
        backtracks=backtracks,
        search_limit_reached=(
            bool(run.search_limit_hit)
            if run.search_limit_hit is not None
            else (snapshot.get("search_limit_hit") if "search_limit_hit" in snapshot else None)
        ),
        completed=(bool(snapshot["complete"]) if "complete" in snapshot else None),
        applied=(status == "APPLIED"),
        scope=str(getattr(run, "scope", "ALL") or "ALL"),
        status=status,
        algorithm=str(getattr(run, "algorithm", "") or ""),
        created_at=run.created_at.isoformat() if run.created_at else None,
        applied_at=run.applied_at.isoformat() if run.applied_at else None,
        reverted_at=run.reverted_at.isoformat() if run.reverted_at else None,
        assigned=int(run.assigned or 0),
        moved=int(run.moved or 0),
        retained=int(run.retained or 0),
        removed=int(run.removed or 0),
        unresolved=int(run.unresolved or 0),
        requirement_total=int(snapshot.get("requirement_total") or 0),
        warnings=list(snapshot.get("warnings") or []),
        unconfigured_courses=list(snapshot.get("unconfigured_courses") or []),
    )
    if not meta.applied:
        meta.errors.append(
            f"This run is recorded as {status or 'unknown'}, not APPLIED. "
            "The audit describes the timetable as it stands now, which may not "
            "be what this run produced."
        )
    if meta.search_limit_reached:
        meta.errors.append(
            "The allocator stopped at its search limit, so it cannot "
            "conclusively establish whether a better or complete alternative "
            "allocation exists."
        )
    return meta


def _origin_lookup(snapshot: dict) -> dict:
    """``{(group code, session pk): origin}`` from the run's plan snapshot."""
    lookup: dict = {}
    for row in snapshot.get("assignments") or []:
        origin = _STATUS_TO_ORIGIN.get(str(row.get("status", "")))
        if not origin:
            continue
        key = (row.get("group"), row.get("session_pk"))
        if None in key:
            continue
        lookup[key] = origin
    return lookup


def _collect_requirements(snapshot, sessions, group_programme, names) -> list:
    """Rebuild the run's requirement list, re-checked against the live links.

    Every requirement the run recorded appears exactly once, with
    ``status='allocated'`` only when a matching ``SessionGroup`` link exists
    right now. That is the difference between "the run said it placed this" and
    "a student is in this class", and the audit reports the second.
    """
    live = {
        (s.group, s.course, s.activity)
        for s in sessions
        if s.activity in {"seminar", "tutorial", "practical"}
    }
    placed = {}
    for sess in sessions:
        if sess.activity in {"seminar", "tutorial", "practical"}:
            placed.setdefault((sess.group, sess.course, sess.activity), sess)

    out = []
    for row in snapshot.get("assignments") or []:
        key = (
            row.get("group"),
            normalise_course_code(row.get("course", "")),
            str(row.get("activity", "")).lower(),
        )
        if key in live:
            found = placed.get(key)
            out.append(
                Requirement(
                    course=key[1],
                    course_name=names.get(key[1], "") or row.get("course_name", ""),
                    activity=key[2],
                    group=key[0],
                    status="allocated",
                    origin=_STATUS_TO_ORIGIN.get(row.get("status"), ORIGIN_UNKNOWN),
                    session_pk=found.session_pk if found else row.get("session_pk"),
                )
            )
        else:
            out.append(
                Requirement(
                    course=key[1],
                    course_name=names.get(key[1], "") or row.get("course_name", ""),
                    activity=key[2],
                    group=key[0],
                    status="unresolved",
                    origin=_STATUS_TO_ORIGIN.get(row.get("status"), ORIGIN_UNKNOWN),
                    reason=(
                        "The run placed this group, but no session-group link "
                        "exists now — it was removed or never applied."
                    ),
                )
            )
    for row in snapshot.get("unresolved_items") or []:
        out.append(
            Requirement(
                course=normalise_course_code(row.get("course", "")),
                course_name=names.get(
                    normalise_course_code(row.get("course", "")), ""
                )
                or row.get("course_name", ""),
                activity=str(row.get("activity", "")).lower(),
                group=row.get("group"),
                status="unresolved",
                origin=ORIGIN_UNKNOWN,
                reason="; ".join(row.get("reasons") or []) or None,
                ordinal=int(row.get("ordinal") or 1),
            )
        )
    return sorted(out, key=lambda r: (r.group, r.course, r.activity, r.ordinal))
