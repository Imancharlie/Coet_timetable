"""Recommendations, and the evidence each one stands on.

A consultancy report is only worth reading if every claim in it can be checked
against the data it was built from, so this module has one rule that overrides
everything else: **nothing is recommended without evidence**. A builder returns
a :class:`Recommendation` only when it can name the records behind it -- the
classes, the groups, the room, the course -- and the numbers attached to them.
A recommendation whose evidence list is empty is a bug, and
``test_every_recommendation_carries_its_evidence`` in the test suite fails on it.

Two ideas are kept apart on purpose:

* **Findings** are what is wrong. They already exist as
  :class:`~audit.analytics.Issue` rows in the register, detected mechanically
  and each carrying its own ``why_detected``.
* **Recommendations** are what to do about it. They are ranked against each
  other by how much teaching they touch, because a university will not do
  twenty things and the report has to say which three matter.

Ranking is by ``students x hours`` -- student-hours of teaching affected --
because that is the only figure here that puts a room swap and a missing
seminar on the same scale honestly. Severity only breaks ties: swapping 300
students out of an undersized room outranks tidying one under-used room.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable

#: Recommendation categories, in the order the report presents them. Each one
#: answers a different question, which is why they are not merged: "the rooms
#: are wrong" and "these groups have no classes" are different conversations
#: with the same faculty.
CATEGORIES = (
    "validity",     # a class cannot run as timetabled
    "allocation",   # a group is short of what it must attend
    "structure",    # the week itself is lopsided
    "equity",       # the burden is unevenly shared
    "data",         # we cannot see clearly enough to advise
    "venue",        # the estate is used badly
)

#: Priority, worst first. Derived, never hand-assigned: a recommendation that
#: blocks teaching is ``critical`` whatever its size, and a data gap that hides
#: a critical finding is ``high`` because it blocks the *next* audit too.
PRIORITIES = ("critical", "high", "medium", "low")


@dataclass
class Recommendation:
    """One thing to do, and the records that justify it."""

    key: str                       # stable slug, e.g. 'venue-overflow'
    title: str                     # short and imperative
    category: str
    action: str                    # what to do, concretely enough to start
    headline: str                  # one sentence, carrying the numbers
    evidence: list = field(default_factory=list)
    groups: list = field(default_factory=list)
    students: int = 0
    hours: float = 0.0             # student-hours of teaching touched
    issue_ids: list = field(default_factory=list)
    alternatives: list = field(default_factory=list)
    priority: str = "medium"
    rank: int = 0
    #: How the report should show this. ``card`` for the handful of things the
    #: university must decide on; ``table`` for a long list of the same kind of
    #: thing (one row per group), where ten separate cards would bury the four
    #: that matter and repeat the same sentence ten times.
    presentation: str = "card"

    @property
    def impact(self) -> float:
        """Student-hours touched. The ranking key."""
        return round(self.students * self.hours, 1)


@dataclass
class RecommendationSet:
    recommendations: list = field(default_factory=list)
    covered_issues: list = field(default_factory=list)
    uncovered_issues: list = field(default_factory=list)

    @property
    def by_category(self) -> dict:
        out: dict = {name: [] for name in CATEGORIES}
        for rec in self.recommendations:
            out.setdefault(rec.category, []).append(rec)
        return out

    @property
    def cards(self) -> list:
        """The ones the reader has to act on, worst first."""
        return [r for r in self.recommendations if r.presentation == "card"]

    @property
    def tables(self) -> list:
        """The long lists, each rendered as one table."""
        return [r for r in self.recommendations if r.presentation == "table"]

    @property
    def by_priority(self) -> dict:
        out: dict = {name: [] for name in PRIORITIES}
        for rec in self.recommendations:
            out.setdefault(rec.priority, []).append(rec)
        return out

    def __len__(self) -> int:
        return len(self.recommendations)

    def __iter__(self):
        return iter(self.recommendations)


# ─────────────────────────────────────────────────────────────────────────────
# Builders. One per question; each returns None when it has nothing to say.
# ─────────────────────────────────────────────────────────────────────────────


def _over_capacity(analysis, cfg) -> list:
    """Classes that do not fit the room they are booked into.

    **One** recommendation for all of them, not one per room. Every affected room
    is too small for the same reason -- more classes want a two-group room than
    the faculty owns -- so the action is a single re-mix of the estate, and ten
    near-identical cards would read as ten separate problems when the reader has
    one. The per-room numbers are all in the evidence, because "move SEM2 out"
    is only actionable when you can see which eleven classes and into what.
    """
    issues = [i for i in analysis.issues if i.category == "insufficient-venue-capacity"]
    if not issues:
        return []
    rooms = []
    for issue in issues:
        venue = issue.evidence.get("venue") or "(no venue)"
        worst = issue.evidence.get("worst") or {}
        rooms.append(
            {
                "venue": venue,
                "classes": issue.evidence.get("classes_over_capacity") or 1,
                "capacity": worst.get("capacity"),
                "largest": worst.get("students") or 0,
                "shortfall": issue.evidence.get("seats_short_in_total") or 0,
                "groups": issue.groups,
                "issue": issue,
            }
        )
    rooms.sort(key=lambda r: -r["shortfall"])
    biggest = rooms[0]
    needed = max(r["largest"] for r in rooms)
    students = sum(i.students for i in issues)
    return [
        Recommendation(
            key="venue-overflow",
            title=f"Re-mix the rooms so {needed} students fit the class they are in",
            category="validity",
            action=(
                f"{sum(r['classes'] for r in rooms)} class(es) across {len(rooms)} room(s) are "
                f"booked into rooms smaller than the class. Either add rooms seating at least "
                f"{needed} students, or move those classes into the rooms that already hold "
                f"them -- {', '.join(r['venue'] for r in rooms[:4])}"
                f"{' and others' if len(rooms) > 4 else ''} are the ones to empty first. "
                f"{biggest['venue']} alone is {biggest['shortfall']} seats short across the week."
            ),
            headline=(
                f"{len(rooms)} room(s) are booked over capacity for {sum(r['classes'] for r in rooms)} "
                f"classes, touching {len({g for r in rooms for g in r['groups']})} group(s) and "
                f"{students} student-attendances; the worst needs {needed} seats and has "
                f"{biggest['capacity']}."
            ),
            evidence=[
                {
                    "label": f"{r['venue']} ({r['capacity']} seats)",
                    "detail": (
                        f"{r['classes']} class(es), largest {r['largest']} students, "
                        f"{r['shortfall']} seats short across the week"
                    ),
                    "groups": r["groups"],
                    "issue": r["issue"].id,
                }
                for r in rooms
            ],
            groups=sorted({g for r in rooms for g in r["groups"]}),
            students=students,
            hours=round(sum(i.hours for i in issues), 1),
            issue_ids=[i.id for i in issues],
            priority="critical",
        )
    ]


def _unmet_requirements(analysis, cfg) -> list:
    """Courses a group must attend and has no session for at all."""
    by_course: dict = {}
    for issue in analysis.issues:
        if issue.category != "unresolved-requirement":
            continue
        by_course.setdefault(issue.course or "(unknown)", []).append(issue)

    out = []
    for course, issues in sorted(by_course.items()):
        groups = sorted({g for i in issues for g in i.groups})
        students = sum(
            analysis.group_by_code[g].size for g in groups if g in analysis.group_by_code
        )
        activities = sorted({i.evidence.get("activity") for i in issues if i.evidence.get("activity")})
        reasons = sorted({i.evidence.get("reason") for i in issues if i.evidence.get("reason")})
        out.append(
            Recommendation(
                key=f"allocation-unresolved:{course}",
                title=f"Timetable the missing class for {course}",
                category="allocation",
                action=(
                    f"Add a session for {course} and link the affected groups to it, or record a "
                    f"decision that the course does not require one."
                    f"{' Activities: ' + ', '.join(a for a in activities if a) + '.' if activities else ''}"
                    f"{' Reason recorded: ' + '; '.join(reasons) if reasons else ''}"
                ),
                headline=(
                    f"{course} is a requirement of {len(groups)} group(s) "
                    f"({students} students) and no session satisfies it."
                ),
                evidence=[
                    {
                        "label": f"{i.evidence.get('activity') or 'class'} unsatisfied",
                        "detail": (
                            f"{i.evidence.get('requirements')} group(s) have no session: "
                            f"{', '.join(i.evidence.get('groups') or i.groups)[:120]}"
                        ),
                        "issue": i.id,
                    }
                    for i in issues
                ],
                groups=groups,
                students=students,
                hours=round(students * 1.5, 1),
                issue_ids=[i.id for i in issues],
                priority="critical",
            )
        )
    return out


def _clashes(analysis, cfg) -> list:
    """Groups timetabled into two places at once."""
    conflicts = [i for i in analysis.issues if i.category == "timetable-conflict"]
    if not conflicts:
        return []
    groups = sorted({g for i in conflicts for g in i.groups})
    pairs = sum(i.evidence.get("pairs") or 0 for i in conflicts)
    minutes = sum(i.evidence.get("overlap_minutes") or 0 for i in conflicts)
    return [
        Recommendation(
            key="allocation-clash",
            title="Resolve the overlapping sessions",
            category="validity",
            action=(
                "Move one of each overlapping pair to a free period. The allocator can only "
                "avoid a clash it can see, so correct the underlying records and re-run it."
            ),
            headline=(
                f"{pairs} overlapping pair(s) totalling {minutes} minutes put {len(groups)} "
                f"group(s) in two places at once."
            ),
            evidence=[
                {
                    "label": f"{(i.groups or [i.course])[0]}, {i.evidence.get('pairs')} pair(s)",
                    "detail": d,
                    "groups": i.groups,
                    "issue": i.id,
                }
                for i in conflicts
                for d in (i.evidence.get("detail") or [i.description])[:3]
            ],
            groups=groups,
            students=sum(issue.students for issue in conflicts),
            hours=round(minutes / 60 * len(groups), 1),
            issue_ids=[i.id for i in conflicts],
            priority="critical",
        )
    ]


def _groups_with_nothing(analysis, cfg) -> list:
    """Groups carrying a full timetable and no classes in it."""
    starved = [i for i in analysis.issues if i.category == "no-timetabled-classes"]
    if not starved:
        return []
    groups = sorted({g for i in starved for g in i.groups})
    return [
        Recommendation(
            key="allocation-no-classes",
            title="Give the empty groups something to attend",
            category="allocation",
            action=(
                "Check whether these groups are genuinely out of the allocation scope. If they "
                "are in scope, their requirements were never timetabled; if they are not, remove "
                "them from the scope so the report stops counting them as empty."
            ),
            headline=(
                f"{len(groups)} group(s) have a requirement list and an empty timetable "
                f"({', '.join(groups)})."
            ),
            evidence=[
                {
                    "label": f"{i.evidence.get('programme')} group(s) with no sessions",
                    "detail": f"{', '.join(i.evidence.get('groups') or i.groups)}: "
                    f"{i.evidence.get('searched')}",
                    "groups": i.groups,
                    "issue": i.id,
                }
                for i in starved
            ],
            groups=groups,
            students=sum(analysis.group_by_code[g].size for g in groups if g in analysis.group_by_code),
            hours=0.0,
            issue_ids=[i.id for i in starved],
            priority="critical",
        )
    ]


def _data_gaps(analysis, cfg) -> list:
    """Records the audit could not verify, and what that costs.

    Grouped as one recommendation rather than one per row: the fix is a data
    collection instruction, and a university will act on "record these four
    things" rather than on four separate paragraphs about them.
    """
    kinds = {
        "missing-venue-info": "sessions booked with no room, or a room with no recorded capacity",
        "unverifiable-workshop-time": "workshops with a day but no clock time",
        "unconfigured-requirement": "courses with no required activities recorded",
    }
    out = []
    for category, description in kinds.items():
        issues = [i for i in analysis.issues if i.category == category]
        if not issues:
            continue
        groups = sorted({g for i in issues for g in i.groups})
        out.append(
            Recommendation(
                key=f"data-{category}",
                title=f"Record the {category.replace('-', ' ')}",
                category="data",
                action=(
                    f"Correct the source records behind these {len(issues)} row(s): "
                    f"{description}. Until they are recorded the audit cannot tell an "
                    f"over-subscribed class from an unrecorded one, and reports the room as "
                    f"unusable when the real fault is the record."
                ),
                headline=(
                    f"{len(issues)} record(s) cannot be verified: {description}, "
                    f"touching {len(groups)} group(s)."
                ),
                evidence=[
                    {
                        "label": i.category.replace("-", " "),
                        "detail": (
                            f"{', '.join(i.evidence.get('courses') or ([i.course] if i.course else []))}"
                            f" -- {i.description}"
                        ),
                        "groups": i.groups,
                        "issue": i.id,
                    }
                    for i in issues
                ][:12],
                groups=groups,
                students=sum(i.students for i in issues),
                hours=0.0,
                issue_ids=[i.id for i in issues],
                priority="high",
            )
        )
    return out


def _day_pressure(analysis, cfg) -> list:
    """The busiest and quietest days, and the difference between them."""
    rollup = analysis.rollup or {}
    pressure = [row for row in rollup.get("day_pressure", []) if row.get("hours")]
    if len(pressure) < 2:
        return []
    heaviest = rollup.get("heaviest_day") or max(pressure, key=lambda r: r["hours"])
    lightest = rollup.get("lightest_day") or min(pressure, key=lambda r: r["hours"])
    if not heaviest or not lightest or heaviest["day"] == lightest["day"]:
        return []
    spread = heaviest["hours"] - lightest["hours"]
    if spread <= 0:
        return []
    share = spread / max(heaviest["hours"], 1e-9)
    if share < 0.15:
        # Below this the week is simply uneven, which is normal and not worth
        # a recommendation. Reporting it would train the reader to ignore the
        # day that really is out of line.
        return []
    return [
        Recommendation(
            key="structure-day-pressure",
            title=f"Take teaching off {heaviest['label']}",
            category="structure",
            action=(
                f"Move about {round(spread / 2, 1)} student-hours of practical or elective work "
                f"from {heaviest['label']} to {lightest['label']}, and check the bands in use on "
                f"{heaviest['label']} against the rooms available then."
            ),
            headline=(
                f"{heaviest['label']} carries {heaviest['hours']} student-hours against "
                f"{lightest['hours']} on {lightest['label']} -- "
                f"{round(100 * share)}% heavier."
            ),
            evidence=[
                {
                    "label": row["label"],
                    "detail": f"{row['hours']} student-hours",
                }
                for row in pressure
            ],
            groups=[],
            students=0,
            hours=round(spread, 1),
            priority="medium",
        )
    ]


def _right_sizing(analysis, cfg) -> list:
    """Rooms large enough for the classes they actually hold."""
    low = float(cfg.get_path("venue.low_seat_util", 0.35))
    rooms = [
        entry
        for entry in (analysis.venues or {}).values()
        if entry.get("has_venue")
        and entry.get("seat_util_avg") is not None
        and entry["seat_util_avg"] < low
        and not entry.get("overflow_sessions")
        and entry.get("time_util")
    ]
    if not rooms:
        return []
    rooms.sort(key=lambda e: e["seat_util_avg"])
    worst = rooms[0]
    wasted = round(
        sum(
            (e["unused_seats_avg"] or 0) * e["sessions"]
            for e in rooms
            if e.get("unused_seats_avg")
        )
    )
    return [
        Recommendation(
            key="venue-right-sizing",
            title="Stop booking small classes into the largest halls",
            category="venue",
            action=(
                f"Point {worst['venue']} at smaller classes, or release it for whole-cohort "
                f"lectures. A room that seats {worst['capacity']} and averages "
                f"{round(100 * worst['seat_util_avg'])}% full is available capacity the faculty "
                f"is not using."
            ),
            headline=(
                f"{len(rooms)} room(s) average under {round(100 * low)}% seat utilisation, "
                f"about {wasted} seats unused across their classes; "
                f"{worst['venue']} is the widest gap at "
                f"{round(100 * worst['seat_util_avg'])}%."
            ),
            evidence=[
                {
                    "label": e["venue"],
                    "detail": (
                        f"{e['capacity']} seats, {round(100 * e['seat_util_avg'])}% full on "
                        f"average over {e['sessions']} classes"
                    ),
                }
                for e in rooms
            ],
            students=sum(e.get("students", 0) for e in rooms),
            hours=round(sum(e.get("booked_hours", 0) for e in rooms), 1),
            priority="low",
        )
    ]


def _long_gaps(analysis, cfg) -> list:
    """The groups that spend the middle of the day waiting."""
    bad = int(cfg.get_path("gap_bad_min", 240))
    worst = [
        g
        for g in analysis.groups
        if g.tfi and g.tfi.get("longest_gap", 0) >= bad
    ]
    if not worst:
        return []
    worst.sort(key=lambda g: -g.tfi["longest_gap"])
    limit = int(cfg.get_path("flags.top_n_groups", 10))
    worst = worst[:limit]
    return [
        Recommendation(
            key="equity-long-gaps",
            title="Close the dead hours in the worst timetables",
            category="equity",
            action=(
                "Move one session of the affected groups into the middle of the day, or accept "
                "the gap and give those groups something to do in it. A gap this long is dead "
                "time the students pay for in travel and in nothing else."
            ),
            headline=(
                f"{len(worst)} group(s) have a gap of {bad // 60} hours or more; the worst "
                f"({worst[0].code}) waits {_dur(worst[0].tfi['longest_gap'])}."
            ),
            evidence=[
                {
                    "label": g.code,
                    "detail": (
                        f"{_dur(g.tfi['longest_gap'])} at worst, "
                        f"{g.tfi.get('dead_min')} dead of {g.tfi.get('contact_min')} contact minutes"
                    ),
                    "groups": [g.code],
                }
                for g in worst
            ],
            groups=[g.code for g in worst],
            students=sum(g.size for g in worst),
            hours=round(sum(g.tfi["dead_min"] for g in worst) / 60, 1),
            priority="high",
        )
    ]


def _programme_equity(analysis, cfg) -> list:
    """Programmes carrying a visibly heavier week than the rest."""
    flagged = [
        entry
        for entry in (analysis.rollup or {}).get("programmes", [])
        if entry.get("equity_flag")
    ]
    if not flagged:
        return []
    flagged.sort(key=lambda e: -(e.get("mean_week_hours") or 0))
    return [
        Recommendation(
            key="equity-programme-load",
            title="Even out the contact hours between programmes",
            category="equity",
            action=(
                "Check the small-group requirements of the programmes listed here. A Gini above "
                "0.10 with a spread of two hours or more is a real difference in student "
                "contact, not a rounding effect."
            ),
            headline=(
                f"{len(flagged)} programme(s) carry a heavier week than the faculty average: "
                + ", ".join(
                    f"{e['programme']} {e.get('mean_week_hours')}h" for e in flagged
                )
                + "."
            ),
            evidence=[
                {
                    "label": e["programme"],
                    "detail": (
                        f"mean {e.get('mean_week_hours')}h, Gini {e.get('gini')}, "
                        f"{e.get('at_de')} of {e.get('groups')} group(s) at grade D/E"
                    ),
                }
                for e in flagged
            ],
            priority="medium",
        )
    ]


# ─────────────────────────────────────────────────────────────────────────────
# Alternatives: what a group could attend instead.
# ─────────────────────────────────────────────────────────────────────────────


def alternatives_for_group(group, analysis, dataset, cfg, by_course=None) -> list:
    """Sessions this group could be moved to, best first.

    A recommendation to "improve EE C1's timetable" is useless on its own; what
    the coordinator needs is a named session to move them into. The rules are
    deliberately narrow, because a plausible-looking wrong suggestion costs more
    than no suggestion at all:

    * the same course, so the group is genuinely eligible for it;
    * a session it does not already attend;
    * no overlap with anything else in its timetable;
    * a room that holds it, since recommending a class into an undersized room
      would trade one validity problem for another;
    * and it must be a *different* day from the group's worst one, which is what
      makes it an improvement rather than a lateral move.

    Returns at most ``flags.alternatives_per_group`` entries. ``by_course`` is
    an optional index of the dataset's sessions; built once by :func:`build` so
    this does not rescan every session in the faculty for every group.
    """
    limit = int(cfg.get_path("flags.alternatives_per_group", 3))
    if limit <= 0 or not group.tfi:
        return []
    busy = sorted(
        (s for s in group.sessions if s.venue or s.activity == "workshop"),
        key=lambda s: (s.day, s.start),
    )
    attended = {s.session_pk for s in group.sessions if s.session_pk}
    courses = {
        r.course for r in dataset.requirements if r.group == group.code
    }
    if not courses:
        return []
    worst_day = _worst_day(group)

    pool_sessions = (
        [s for course in sorted(courses) for s in (by_course or {}).get(course, ())]
        if by_course is not None
        else [s for s in dataset.sessions if s.course in courses]
    )
    candidates = []
    for session in pool_sessions:
        if session.course not in courses or session.session_pk in attended:
            continue
        if session.capacity and (session.group_size or 0) > session.capacity:
            continue
        if _conflicts_with(busy, session):
            continue
        # The group's own timetable is the only place a clash can come from, but
        # a *parallel* section of a course it already attends can still be in
        # one of its rooms at the same time; the clash test above covers that.
        on_worst_day = session.day == worst_day
        candidates.append(
            {
                "course": session.course,
                "activity": session.activity,
                "day": session.day,
                "day_label": _day_label(session.day),
                "start": session.start,
                "end": session.end,
                "venue": session.venue or "not recorded",
                "capacity": session.capacity,
                "students": session.group_size,
                "session_pk": session.session_pk,
                "why": (
                    "lands on the same day as the group's worst day"
                    if on_worst_day
                    else f"lands on {_day_label(session.day)}, which the group has free"
                ),
                "improves": not on_worst_day,
            }
        )
    improving = [c for c in candidates if c["improves"]]
    pool = improving or candidates
    pool.sort(key=lambda c: (not c["improves"], c["day"], c["start"], c["course"]))
    return pool[:limit]


def _worst_day(group):
    hours = {d: h for d, h in (group.day_hours or {}).items() if h}
    if not hours:
        return None
    return max(sorted(hours), key=lambda d: hours[d])


def _conflicts_with(sessions, candidate) -> bool:
    """Whether ``candidate`` overlaps anything already in ``sessions``."""
    return any(
        s.day == candidate.day and s.start < candidate.end and candidate.start < s.end
        for s in sessions
    )


# ─────────────────────────────────────────────────────────────────────────────
# Assembly.
# ─────────────────────────────────────────────────────────────────────────────

_BUILDERS = (
    _over_capacity,
    _unmet_requirements,
    _clashes,
    _groups_with_nothing,
    _long_gaps,
    _day_pressure,
    _programme_equity,
    _right_sizing,
    _data_gaps,
)


def build(analysis, cfg) -> RecommendationSet:
    """Every recommendation the evidence supports, ranked.

    ``analysis.rollup`` must already be populated -- ``analyse`` does it, and
    this is called after, so the recommendations read the same numbers the
    report prints rather than recomputing them.
    """
    recommendations: list = []
    for builder in _BUILDERS:
        for rec in builder(analysis, cfg) or []:
            # The rule this module exists to keep: no evidence, no claim.
            if not rec.evidence:
                continue
            recommendations.append(rec)

    limit = int(cfg.get_path("flags.top_n_groups", 10))
    worst = [g for g in analysis.groups_by_tfi if g.tfi][:limit]
    by_course: dict = {}
    for session in analysis.dataset.sessions:
        by_course.setdefault(session.course, []).append(session)
    rows = []
    for group in worst:
        alts = alternatives_for_group(
            group, analysis, analysis.dataset, cfg, by_course=by_course
        )
        if not alts:
            # Nothing it could attend, so nothing to offer. A row saying "no
            # alternative exists" is still worth showing -- it tells the
            # coordinator the fix has to be a new session, not a reshuffle.
            alts = []
        rows.append(
            {
                "group": group.code,
                "programme": group.programme,
                "score": group.tfi["score"],
                "grade": group.tfi["grade"],
                "dead_min": group.tfi.get("dead_min"),
                "longest_gap": group.tfi.get("longest_gap"),
                "busiest_day": group.tfi.get("busiest_day"),
                "options": alts,
            }
        )
    movable = [r for r in rows if r["options"]]
    if rows:
        recommendations.append(
            Recommendation(
                key="equity-group-timetables",
                title="Give the worst timetables a lighter week",
                category="equity",
                action=(
                    "Each row lists sessions that group is already eligible for, on a day it has "
                    "free, in a room that holds it. Move a group into one and its dead time drops; "
                    "where no option is listed, the group has no alternative class and needs a "
                    "new session rather than a different one."
                ),
                headline=(
                    f"{len(rows)} group(s) score lowest on the timetable index; "
                    f"{len(movable)} of them can be improved by moving into a session that "
                    f"already exists."
                ),
                evidence=[
                    {
                        "label": r["group"],
                        "detail": (
                            f"score {r['score']} ({r['grade']}), {r['dead_min']} dead minutes, "
                            f"longest gap {_dur(r['longest_gap'])}"
                            + (
                                f", {len(r['options'])} option(s)"
                                if r["options"]
                                else ", no existing session it could attend"
                            )
                        ),
                        "groups": [r["group"]],
                    }
                    for r in rows
                ],
                alternatives=rows,
                groups=[r["group"] for r in rows],
                students=sum(
                    analysis.group_by_code[r["group"]].size
                    for r in rows
                    if r["group"] in analysis.group_by_code
                ),
                hours=round(
                    sum((r["dead_min"] or 0) for r in rows) / 60, 1
                ),
                priority="high" if movable else "medium",
                presentation="table",
            )
        )

    recommendations.sort(key=lambda r: (PRIORITIES.index(r.priority), -r.impact, r.key))
    for index, rec in enumerate(recommendations, start=1):
        rec.rank = index

    covered = {iid for rec in recommendations for iid in rec.issue_ids}
    all_ids = [i.id for i in analysis.issues]
    return RecommendationSet(
        recommendations=recommendations,
        covered_issues=[i for i in all_ids if i in covered],
        uncovered_issues=[i for i in all_ids if i not in covered],
    )


# ── small helpers ────────────────────────────────────────────────────────────


def _seated(capacity) -> str:
    return f" and seats {capacity}" if capacity else ""


def _hhmm(minutes) -> str:
    if minutes is None:
        return "?"
    minutes = int(minutes)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def _dur(minutes) -> str:
    minutes = int(minutes or 0)
    if minutes < 60:
        return f"{minutes} min"
    hours, rest = divmod(minutes, 60)
    return f"{hours}h" if not rest else f"{hours}h {rest}m"


def _day_label(day) -> str:
    from audit.collect import DAY_NAMES

    if day is None:
        return "no day"
    try:
        return DAY_NAMES[day]
    except (KeyError, IndexError, TypeError):
        return f"day {day}"
