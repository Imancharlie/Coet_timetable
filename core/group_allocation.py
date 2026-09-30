"""Tutorial, seminar and practical group allocation.

This module owns every decision about *which* group attends *which* session. It
is deliberately free of Django views and templates: the coordinator page, the
manual-assignment form and the tests all call the same functions, so a manual
edit can never be accepted on rules the automatic run would have rejected.

How a run is put together
-------------------------
1. **Requirements** — for the chosen semester, every student group of a
   programme that studies a course gets one requirement per configured activity
   of that course (``Course`` -> ``CourseActivityRequirement``). Courses with
   no requirements produce nothing and are reported for the coordinator's
   attention instead.
2. **Validation** — three independent checks, each of which can veto an
   assignment: course eligibility, timetable availability and venue capacity
   (see :func:`validate_assignment`).
3. **Search** — a depth-first search over the requirements in the order
   seminar -> tutorial -> practical, most-constrained requirement first inside
   each stage. Earlier stages are *not* frozen: a later stage that cannot be
   satisfied makes the search backtrack and try a different seminar or tutorial
   choice. When a requirement cannot be placed at all it is reported with the
   specific reason and the search continues with the next one, so a run always
   produces the best plan it can rather than nothing.

Data problems are never worked around. A venue with no capacity, a session
with no venue, an activity with no session and a workshop whose time cannot be
resolved all leave the affected requirement unresolved with an explanation.
"""

import html
import json
import re
import time
from dataclasses import dataclass, field

from django.db import transaction

from .models import (
    ALLOCATED_ACTIVITY_TYPES,
    ActivityType,
    AllocationChange,
    AllocationRun,
    AllocationStatus,
    Course,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    WorkshopAllocation,
    normalise_course_code,
)
from .workshop_times import is_workshop_day, workshop_times_for

#: Students per group for capacity arithmetic. A group is treated as a whole
#: cohort of this size; it is the only student-count assumption in the module.
STUDENTS_PER_GROUP = 30

#: Default ceiling on search nodes. Reaching it means "not proven impossible",
#: which is reported as such -- never as "no allocation exists".
DEFAULT_NODE_LIMIT = 250_000

REQUIREMENT_ACTIVITY_SCOPE = {
    "SEMINAR": (ActivityType.SEMINAR,),
    "TUTORIAL": (ActivityType.TUTORIAL,),
    "PRACTICAL": (ActivityType.PRACTICAL,),
    "ALL": ALLOCATED_ACTIVITY_TYPES,
}


def scope_activities(scope: str) -> tuple:
    """Activity types covered by a scope string ("ALL" or a single activity)."""
    return REQUIREMENT_ACTIVITY_SCOPE.get(scope, ALLOCATED_ACTIVITY_TYPES)

# Violation kinds, used for grouping in the UI and for counting in the summary.
KIND_ELIGIBILITY = "eligibility"
KIND_AVAILABILITY = "availability"
KIND_CAPACITY = "capacity"
KIND_DATA = "data"


# ──────────────────────────────────────────────
# Requirements: reading them out of a workbook
# ──────────────────────────────────────────────

#: Every spelling that names an allocatable activity, singular or plural, plus
#: the abbreviations that appear in real timetables. The lookup is by *substring*
#: (see ``_match_activities``) so an unseparated "Seminar Tutorial" still reads
#: as two activities rather than one unrecognised word.
_ACTIVITY_WORDS = {
    "seminar": ActivityType.SEMINAR,
    "seminars": ActivityType.SEMINAR,
    "sem": ActivityType.SEMINAR,
    "tutorial": ActivityType.TUTORIAL,
    "tutorials": ActivityType.TUTORIAL,
    "tut": ActivityType.TUTORIAL,
    "tuts": ActivityType.TUTORIAL,
    "practical": ActivityType.PRACTICAL,
    "practicals": ActivityType.PRACTICAL,
    "pract": ActivityType.PRACTICAL,
    "prac": ActivityType.PRACTICAL,
    "pracs": ActivityType.PRACTICAL,
}

#: Longest first, so "practical" wins over "prac" and "seminar" over "sem".
_ACTIVITY_RE = re.compile(
    r"\b(?:%s)\b" % "|".join(sorted(_ACTIVITY_WORDS, key=len, reverse=True)),
    re.IGNORECASE,
)

#: Word numbers, because "Two practicals" is a perfectly normal way to write it.
_COUNT_WORDS = {
    "a": 1, "an": 1, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11,
    "twelve": 12,
}

#: Cells that mean "this course needs no allocation" rather than a value to
#: parse. Matched against the whole cell and against each individual token, so
#: "None required" and "Tutorial; not applicable" both behave.
_BLANK_CELLS = {
    "-", "--", "n/a", "na", "n.a.", "none", "nil", "no", "nones", "nothing",
    "not required", "no requirement", "no requirements", "none required",
    "not applicable", "n/a (not applicable)", "no allocation",
    "no allocation required", "not allocated", "blank", "tbc",
}

#: What separates two activities in one cell. Newlines first: a cell copied
#: out of a Word table arrives with a line break between activities. "-" is
#: deliberately absent because it also carries counts ("Practical x-2").
_REQUIREMENT_SPLIT_RE = r"[\n\r;,\|/•·+&]|\band\b"

#: Count markers, tried in order; the first one that matches wins. All of them
#: have to agree on the number so "Practical (2)" is 2 and not 1.
_COUNT_PATTERNS = (
    re.compile(r"^\s*(\d+)\s*(?:x\s+|times\s+|per\s+)?", re.IGNORECASE),
    re.compile(
        r"^\s*(%s)\b\s*" % "|".join(sorted(_COUNT_WORDS, key=len, reverse=True)),
        re.IGNORECASE,
    ),
    re.compile(r"\bx\s*(\d+)\b", re.IGNORECASE),
    # A bracketed count has to be tried before a bare trailing number, because
    # "Practical (2)" does not *end* in a digit.
    re.compile(r"[\(\[\{]\s*(\d+)\s*[\)\]\}]\s*$"),
    re.compile(r"(\d+)\s*$"),
)


def _is_blank_cell(text: str) -> bool:
    """True when a cell/token means "no requirement stated"."""
    key = " ".join(str(text or "").strip().lower().split()).strip(" .;:")
    if not key:
        return True
    return key in _BLANK_CELLS


def _split_count(text: str):
    """``(text_without_the_count, count_or_None)``.

    One count is taken, never two: "2 practicals", "2 x practicals", "two
    practicals", "practical x2", "practical (2)" and "practical 2" all have to
    agree, and picking the first marker in the list above guarantees that
    rather than multiplying two numbers out of one cell.
    """
    for pattern in _COUNT_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        raw = match.group(1)
        if raw.isdigit():
            count = int(raw)
            remainder = text[: match.start()] + " " + text[match.end():]
        else:
            count = _COUNT_WORDS[raw.lower()]
            remainder = text[match.end():]
        if count >= 1:
            return remainder, count
    return text, None


def _match_activities(text: str) -> list:
    """Every allocatable activity named anywhere in ``text``, in order."""
    return [
        _ACTIVITY_WORDS[match.group(0).lower()] for match in _ACTIVITY_RE.finditer(text)
    ]


def parse_requirements(value):
    """Parse a "Required Activities" cell into ``({activity: count}, problems)``.

    Understands the ways this is actually written in spreadsheets:

    ==========================  ==========================================
    ``Tutorial``                one tutorial
    ``Tutorial; Practical``     both, one of each
    ``tutorials and practicals``  case and plurals do not matter
    ``Seminar | Practical``     any of ``; , / | • + &`` or a line break
    ``Seminar Tutorial``        an unseparated pair still reads as two
    ``2 practicals``            a count, in almost any position
    ``2 x Practical``           ditto, including ``x2`` and ``(2)``
    ``Two practicals``          spelled-out numbers
    ``Tutorials per week: 2``   a trailing count
    ``-`` / ``None`` / ``N/A``  no requirement stated
    ==========================  ==========================================

    Returns the counts plus a list of human-readable problems. An unrecognised
    word is reported rather than dropped, and **a cell that is understood but
    cannot be applied never returns a silently wrong count** — that is the one
    failure mode that would quietly mis-allocate a whole course.
    """
    if value is None:
        return {}, []
    text = str(value).strip()
    if not text:
        return {}, []
    # A cell pasted out of a web page arrives with its separators escaped.
    text = html.unescape(text)
    text = (
        text.replace("–", "-")
        .replace("—", "-")
        .replace("‘", "'")
        .replace("’", "'")
        .replace("“", '"')
        .replace("”", '"')
    )
    if _is_blank_cell(text):
        return {}, []

    counts: dict = {}
    problems: list = []
    for raw_token in re.split(_REQUIREMENT_SPLIT_RE, text, flags=re.IGNORECASE):
        token = " ".join(raw_token.split())
        if not token or _is_blank_cell(token):
            continue
        # A token can name several activities ("Seminar Tutorial"), so pull the
        # count off first and then look for every activity word in what is left.
        remainder, count = _split_count(token)
        activities = _match_activities(remainder) or _match_activities(token)
        if not activities:
            problems.append(
                f"'{token}' is not Seminar, Tutorial or Practical"
            )
            continue
        count = count if count is not None else 1
        if count < 1:
            problems.append(f"'{token}' needs a count of 1 or more")
            continue
        for activity in activities:
            # A repeated word adds up: "Tutorial, Tutorial" means two tutorials.
            counts[activity] = counts.get(activity, 0) + count
    return counts, problems


def format_requirements(mapping: dict) -> str:
    """``{TUTORIAL: 1, PRACTICAL: 2}`` -> "Tutorial; 2x Practical"."""
    parts = []
    for activity in ALLOCATED_ACTIVITY_TYPES:
        count = (mapping or {}).get(activity)
        if not count:
            continue
        label = ActivityType(activity).label
        parts.append(f"{count}x {label}" if count > 1 else label)
    return "; ".join(parts)


# ──────────────────────────────────────────────
# Finding the requirements in a workbook
# ──────────────────────────────────────────────

#: Headers accepted verbatim, most specific first. Matching is on letters and
#: digits only, so "Required Activities", "required_activities" and
#: "REQUIREDACTIVITIES" are the same column.
REQUIRED_ACTIVITIES_ALIASES = (
    "required activities",
    "required activity",
    "required activity types",
    "activities required",
    "required sessions",
    "required classes",
    "activity requirements",
    "allocation requirements",
    "allocation requirement",
    "group allocation",
    "requirements",
    "activities",
)

#: A header that is *about requirements* rather than about one activity.
#: Deliberately narrow: these are the words that mean "this column states what
#: is required", and matching on them (plus a word for the thing required) is
#: what lets "Allocated Activities" and "Required Sessions" be found without
#: listing every heading anyone has ever used.
_REQUIREMENT_HEADER_WORDS = ("required", "requirement", "requirements", "req", "alloc")

#: A word for the thing being required — the noun in "Required <thing>".
#: "requirement" is itself such a noun, which is what lets "Alloc
#: Requirements" resolve without needing the word "activities" in the header.
_REQUIRED_THING_WORDS = (
    "activit", "session", "class", "period", "contact", "meeting", "type",
    "requirement",
)

#: A word that makes a header a *number* of one activity, for the separate
#: count-column layout ("Tutorial Count", "Number of Practicals").
_COUNT_HEADER_WORDS = (
    "count", "number", "noof", "qty", "quantity", "howmany", "perweek",
    "persemester", "perterm", "peryear", "amount", "total", "sessions",
    "classes", "periods", "hours",
)

#: Activities are matched in the *original* header text, on word boundaries, so
#: "Semester" is not read as containing "sem" and "Assessment" not as "assess".
_HEADER_ACTIVITY_RE = re.compile(
    r"\b(?:%s)\b" % "|".join(sorted(_ACTIVITY_WORDS, key=len, reverse=True)),
    re.IGNORECASE,
)


def _normalise_header(name) -> str:
    """Letters and digits only, lower-case — the same rule the importers use."""
    return "".join(ch.lower() for ch in str(name) if ch.isalnum())


def _header_activities(name) -> list:
    """Distinct allocatable activities named in a header, in header order."""
    seen = []
    for match in _HEADER_ACTIVITY_RE.finditer(str(name)):
        activity = _ACTIVITY_WORDS[match.group(0).lower()]
        if activity not in seen:
            seen.append(activity)
    return seen


def find_requirements_column(columns, explicit: str = None):
    """Locate the "Required Activities" column in a workbook header.

    Returns ``(column_name_or_None, unrecognised_column_names)``.

    ``unrecognised`` is the important half: a header that is clearly about what
    is required but that this importer did not use is returned so the caller can
    *say so*. Silently ignoring a requirements column is how every course in the
    file quietly ends up "requirement not configured" much later, with nothing
    in between to explain it.

    Resolution is layered, most confident first:

    1. an explicit ``--requirements-column`` header (exact match on letters and
       digits, so spelling and spacing do not matter);
    2. the known aliases in :data:`REQUIRED_ACTIVITIES_ALIASES`;
    3. a header that pairs a requirement word with a noun for the thing required
       — "Allocated Activities", "Required Sessions", "Alloc Requirements".
    """
    names = [str(c) for c in columns]
    normalised = {_normalise_header(c): c for c in names}

    if explicit:
        return normalised.get(_normalise_header(explicit)), []

    for alias in REQUIRED_ACTIVITIES_ALIASES:
        found = normalised.get(_normalise_header(alias))
        if found is not None:
            return found, []

    for name in names:
        key = _normalise_header(name)
        if any(word in key for word in _REQUIREMENT_HEADER_WORDS) and any(
            thing in key for thing in _REQUIRED_THING_WORDS
        ):
            return name, []

    # Nothing was used. Flag any header that still talks about requirements, so
    # the coordinator is told rather than left to guess why nothing imported.
    unrecognised = [
        name
        for name in names
        if any(word in _normalise_header(name) for word in _REQUIREMENT_HEADER_WORDS)
    ]
    return None, unrecognised


def find_requirement_count_columns(columns, exclude=()):
    """Dedicated per-activity count columns, as ``{activity: column_name}``.

    For workbooks that keep the numbers in their own cells rather than in one
    list: "Tutorial Count", "Number of Practicals", "Tutorials per week". A
    header qualifies when it names **exactly one** activity, which is what stops
    the combined "Required Activities" column being read as three counts at
    once. Every such column is read and *added* to whatever the combined cell
    said, so the two layouts can even be mixed in one file.
    """
    found: dict = {}
    excluded = {_normalise_header(c) for c in exclude if c}
    for name in columns:
        key = _normalise_header(name)
        if not key or key in excluded:
            continue
        activities = _header_activities(name)
        if len(activities) != 1:
            continue
        looks_like_a_count = any(word in key for word in _COUNT_HEADER_WORDS)
        if not looks_like_a_count:
            continue
        found.setdefault(activities[0], name)
    return found


# ──────────────────────────────────────────────
# Violations — the shared validation vocabulary
# ──────────────────────────────────────────────


@dataclass(frozen=True)
class Violation:
    """One reason an assignment is not allowed."""

    code: str
    kind: str
    message: str

    def __str__(self):
        return self.message


def _courses_in_play(semester: Semester):
    """``{normalised_code: Course}`` studied by any programme this semester."""
    courses = {}
    rows = (
        ProgrammeCourse.objects.filter(
            semester=semester.semester, course__isnull=False
        )
        .select_related("course")
        .values_list("course_code", "course")
    )
    for raw_code, course in rows:
        code = normalise_course_code(raw_code or course.code)
        courses[code] = course
    return courses


# ──────────────────────────────────────────────
# Timetable availability
# ──────────────────────────────────────────────


@dataclass(frozen=True)
class BusyBlock:
    """Something already occupying a group at a point in the week."""

    kind: str  # "session" | "workshop" | "td" | "unverifiable"
    label: str
    day: str
    start: object = None
    end: object = None
    detail: str = ""
    # The course this block speaks for, normalised. Only meaningful for a
    # technical drawing, and it is what makes "the same class" decidable --
    # see ``AvailabilityIndex.conflicts_excluding``.
    course: str = ""

    @property
    def covers_whole_day(self) -> bool:
        return self.start is None or self.end is None

    def describe(self) -> str:
        """A sentence fragment for the coordinator, free of internal markers.

        A session block's ``detail`` carries a ``#<pk>:`` prefix that lets
        :meth:`AvailabilityIndex.conflicts_excluding` recognise the session's
        own blocks. That marker is an implementation detail and must never
        reach the review panel.
        """
        if self.kind == "unverifiable":
            return f"{self.label} (no time recorded)"
        when = f" {self.start:%H:%M}-{self.end:%H:%M}" if self.start else ""
        return f"{self.label}{when}"


def overlaps(block: BusyBlock, day: str, start, end) -> bool:
    """Interval overlap test used for every clock-based activity.

    A candidate conflicts when it starts before the block ends *and* ends after
    the block starts. A block with no known window (an imported workshop that
    only says which day) is treated as covering the whole day: the day is known
    to be occupied even though the hours are not, so nothing can be verified
    against it and nothing is placed there.
    """
    if block.day != day:
        return False
    if block.covers_whole_day:
        return True
    if start is None or end is None:
        return True
    return start < block.end and end > block.start


class AvailabilityIndex:
    """Every commitment that limits where a group can be placed.

    Four sources are folded together: existing ``SessionGroup`` links (lectures
    included), workshop allocations, technical-drawing allocations, and the
    proposed assignments of the run currently being computed.
    """

    def __init__(self, semester: Semester):
        self.semester = semester
        # group pk -> day -> [BusyBlock]
        self._blocks: dict = {}
        self.unverifiable: list = []
        self._load()

    # -- construction -------------------------------------------------
    def _add(self, group_pk, block: BusyBlock):
        self._blocks.setdefault(group_pk, {}).setdefault(block.day, []).append(block)

    def _groups_by_code(self, codes):
        return list(
            StudentGroup.objects.filter(code__in=codes).values_list("pk", flat=True)
        )

    def _load(self):
        sessions = {
            s.pk: s
            for s in Session.objects.filter(semester=self.semester).select_related(
                "venue"
            )
        }
        links = SessionGroup.objects.filter(
            session__semester=self.semester
        ).select_related("session", "group")
        for link in links:
            session = sessions.get(link.session_id)
            if session is None:
                continue
            self._add(
                link.group_id,
                _session_block(session),
            )

        # Workshop and technical-drawing allocations name a bare group code, so
        # they apply to every group carrying that code in any programme -- the
        # same resolution the timetable exports use.
        for rec in WorkshopAllocation.objects.filter(semester=self.semester):
            start, end = _workshop_window(rec)
            if start is None:
                block = BusyBlock(
                    kind="unverifiable",
                    label=f"Workshop {rec.workshop or rec.course_code}",
                    day=rec.day,
                    detail=(
                        f"{rec.workshop or rec.course_code} on "
                        f"{rec.get_day_display()} has a day but no resolvable "
                        f"time, so availability cannot be verified"
                    ),
                )
                self.unverifiable.append(block)
            else:
                block = BusyBlock(
                    kind="workshop",
                    label=f"Workshop {rec.workshop or rec.course_code}",
                    day=rec.day,
                    start=start,
                    end=end,
                    detail=str(rec),
                )
            for group_pk in self._groups_by_code([rec.group_code]):
                self._add(group_pk, block)

        for rec in TechnicalDrawingAllocation.objects.filter(semester=self.semester):
            block = BusyBlock(
                kind="td",
                label=f"Technical drawing {rec.course_code}",
                day=rec.day,
                start=rec.start_time,
                end=rec.end_time,
                detail=str(rec),
                course=normalise_course_code(rec.course_code or ""),
            )
            for group_pk in self._groups_by_code([rec.group_code]):
                self._add(group_pk, block)

    # -- queries ------------------------------------------------------
    def blocks_for(self, group_pk) -> dict:
        return self._blocks.get(group_pk, {})

    def days_busy(self, group_pk) -> int:
        return len(self._blocks.get(group_pk, {}))

    def commitments_on(self, group_pk, day) -> int:
        return len(self._blocks.get(group_pk, {}).get(day, []))

    def conflicts_excluding(
        self, group_pk, day, start, end, session_pk, course_code=None
    ):
        """Blocks overlapping the candidate window, ignoring ``session_pk``.

        Two kinds of block are deliberately *not* reported as clashes, and both
        exist so the allocator agrees with what the timetable draws:

        - The candidate session's own blocks. A group already linked to the
          session being tested must not be reported as clashing with itself;
          those blocks are recognised by the ``#<pk>:`` prefix every session
          block's ``detail`` carries.
        - **A technical drawing of the candidate's own course.** ``ME101``
          technical drawing 10:00-13:00 and the ``ME101`` tutorial 10:00-12:55
          are one class, not two: ``fold_same_sessions`` already merges them
          into a single block on every export, so calling the group "busy" for
          it would refuse a session that is not actually a second commitment.
          The match is on course, matching the export — a technical drawing of
          a *different* course still clashes, and a workshop still clashes
          whatever its course, because a workshop is its own class.
        """
        wanted = normalise_course_code(course_code or "")
        hits = []
        for block in self._blocks.get(group_pk, {}).get(day, ()):
            if block.kind == "session" and block.detail.startswith("#"):
                if block.detail.split(":", 1)[0] == f"#{session_pk}":
                    continue
            if block.kind == "td" and block.course and block.course == wanted:
                continue
            if overlaps(block, day, start, end):
                hits.append(block)
        return hits

    # -- mutation (used by the search) --------------------------------
    def add_session(self, group_pk, session: Session):
        self._add(group_pk, _session_block(session))

    def drop_session(self, group_pk, session_pk, day):
        blocks = self._blocks.get(group_pk, {}).get(day)
        if not blocks:
            return
        self._blocks[group_pk][day] = [
            b
            for b in blocks
            if not (b.kind == "session" and b.detail.startswith(f"#{session_pk}:"))
        ]


def _session_block(session: Session) -> BusyBlock:
    """The :class:`BusyBlock` for a session, tagged with its primary key."""
    return BusyBlock(
        kind="session",
        label=f"{session.course_code} {session.get_activity_type_display()}",
        day=session.day,
        start=session.start_time,
        end=session.end_time,
        detail=f"#{session.pk}:{session}",
    )


def _workshop_window(rec):
    """``(start, end)`` for a workshop allocation, or ``(None, None)``.

    Explicit clock times win. Otherwise the Morning/Afternoon period is
    resolved through the project's own workshop-time rules, which is what the
    raw matrix workbook leaves behind. A record with neither cannot be placed
    on the timetable at all, so the caller is told it is unverifiable instead of
    being handed a made-up window.
    """
    if rec.start_time is not None and rec.end_time is not None:
        return rec.start_time, rec.end_time
    if rec.day and rec.time_period and is_workshop_day(rec.day):
        return workshop_times_for(rec.day, rec.time_period)
    return None, None


# ──────────────────────────────────────────────
# Capacity
# ──────────────────────────────────────────────


def required_capacity(group_count: int) -> int:
    """Students a session must seat for ``group_count`` groups."""
    return max(int(group_count), 0) * STUDENTS_PER_GROUP


def capacity_status(session: Session, group_count: int) -> tuple[str, str]:
    """``(status, message)`` for seating ``group_count`` groups in ``session``.

    ``status`` is one of ``ok`` / ``no-venue`` / ``unknown-capacity`` /
    ``over-capacity``. A missing venue and a capacity of zero are *not* treated
    as "plenty of room": they are unresolved issues the coordinator has to fix,
    because assuming capacity would silently overfill an unverified room.
    """
    needed = required_capacity(group_count)
    venue = session.venue
    if venue is None:
        return "no-venue", (
            f"No venue is set for this session, so {needed} students "
            f"({group_count} group(s) x {STUDENTS_PER_GROUP}) cannot be "
            f"verified. Set the venue on the session and try again."
        )
    capacity = venue.capacity or 0
    if capacity <= 0:
        return "unknown-capacity", (
            f"Venue {venue.name} has no capacity recorded, so the "
            f"{needed} students required ({group_count} group(s) x "
            f"{STUDENTS_PER_GROUP}) cannot be verified. Set the capacity of "
            f"{venue.name} and try again."
        )
    if needed > capacity:
        return "over-capacity", (
            f"{venue.name} seats {capacity} but {group_count} group(s) need "
            f"{needed} places ({group_count} x {STUDENTS_PER_GROUP})."
        )
    return "ok", f"{venue.name} seats {capacity} for {needed} students."


# ──────────────────────────────────────────────
# The shared validator
# ──────────────────────────────────────────────


def check_eligibility(group, course, session, semester: Semester):
    """Course-eligibility violations for putting ``group`` in ``session``."""
    problems = []
    if session.semester_id != semester.pk:
        problems.append(
            Violation(
                "wrong-semester",
                KIND_ELIGIBILITY,
                f"This session belongs to {session.semester}, not "
                f"{semester}.",
            )
        )
    code = normalise_course_code(session.course_code)
    if code != normalise_course_code(course.code):
        problems.append(
            Violation(
                "course-mismatch",
                KIND_ELIGIBILITY,
                f"This session is for {session.course_code}, not {course.code}.",
            )
        )
    if session.activity_type not in ALLOCATED_ACTIVITY_TYPES:
        problems.append(
            Violation(
                "not-allocatable",
                KIND_ELIGIBILITY,
                f"{session.get_activity_type_display()} sessions are not "
                f"allocated by this feature (workshops and lectures are set "
                f"up elsewhere).",
            )
        )
    if session.activity_type not in course.required_activities():
        required = course.activities_label() or "none configured"
        problems.append(
            Violation(
                "activity-not-required",
                KIND_ELIGIBILITY,
                f"{course.code} does not require a "
                f"{session.get_activity_type_display()} (required: {required}).",
            )
        )
    studies = ProgrammeCourse.objects.filter(
        programme_id=group.programme_id,
        course_id=course.pk,
        semester=semester.semester,
    ).exists()
    if not studies:
        problems.append(
            Violation(
                "group-not-studying",
                KIND_ELIGIBILITY,
                f"{group} does not study {course.code} in semester "
                f"{semester.semester}.",
            )
        )
    return problems


def check_session_data(session: Session):
    """Data problems that make a session impossible to validate reliably."""
    problems = []
    if not session.day:
        problems.append(
            Violation("no-day", KIND_DATA, "The session has no day set.")
        )
    if session.start_time is None or session.end_time is None:
        problems.append(
            Violation(
                "no-time",
                KIND_DATA,
                "The session has no start/end time, so clashes cannot be "
                "checked.",
            )
        )
    elif session.start_time >= session.end_time:
        problems.append(
            Violation(
                "backwards-time",
                KIND_DATA,
                f"The session runs backwards ({session.start_time:%H:%M}-"
                f"{session.end_time:%H:%M}).",
            )
        )
    if session.activity_type not in ActivityType.values:
        problems.append(
            Violation(
                "unknown-activity",
                KIND_DATA,
                f"'{session.activity_type}' is not a known activity type.",
            )
        )
    return problems


def check_availability(group, session: Session, index: AvailabilityIndex):
    """Timetable clashes between ``group`` and ``session``."""
    problems = []
    for block in index.conflicts_excluding(
        group.pk,
        session.day,
        session.start_time,
        session.end_time,
        session.pk,
        course_code=session.course_code,
    ):
        if block.kind == "unverifiable":
            problems.append(
                Violation(
                    "unverifiable-workshop",
                    KIND_AVAILABILITY,
                    f"{group.code} has {block.describe()} on "
                    f"{session.get_day_display()}, so the availability of this "
                    f"session cannot be verified. Set a time period or start/end "
                    f"time on the workshop record, then run the allocator again.",
                )
            )
            continue
        problems.append(
            Violation(
                "clash",
                KIND_AVAILABILITY,
                f"{group.code} is already busy on "
                f"{session.get_day_display()} — {block.describe()}.",
            )
        )
    return problems


def validate_assignment(
    group,
    course,
    session: Session,
    semester: Semester,
    index: AvailabilityIndex,
    *,
    group_count_after: int = None,
):
    """Every reason ``group`` may not be placed in ``session``.

    This is the single validation path: the allocator calls it for each
    candidate, and the coordinator's manual-assignment form calls it for the one
    pair being edited, so the two can never disagree.

    ``group_count_after`` is how many groups would attend the session including
    this one; leave it ``None`` to use the count the index already knows.
    """
    problems = []
    problems += check_session_data(session)
    if problems:
        return problems
    problems += check_eligibility(group, course, session, semester)
    problems += check_availability(group, session, index)
    if group_count_after is None:
        group_count_after = _current_group_count(session)
    status, message = capacity_status(session, group_count_after)
    if status != "ok":
        problems.append(Violation(f"capacity-{status}", KIND_CAPACITY, message))
    return problems


def _current_group_count(session: Session) -> int:
    return SessionGroup.objects.filter(session_id=session.pk).count()


# ──────────────────────────────────────────────
# Requirements for one run
# ──────────────────────────────────────────────


@dataclass
class Requirement:
    """One group needing one session of one activity for one course."""

    group: StudentGroup
    course: Course
    activity_type: str
    ordinal: int  # 1..count, so a course needing 2 practicals makes two of these
    candidates: list = field(default_factory=list)
    existing: list = field(default_factory=list)  # linked candidate sessions

    @property
    def key(self):
        return (self.group.pk, self.course.pk, self.activity_type, self.ordinal)

    @property
    def label(self) -> str:
        label = (
            f"{self.group.code} - {self.course.code} "
            f"{ActivityType(self.activity_type).label}"
        )
        if self.ordinal > 1:
            label += f" #{self.ordinal}"
        return label


@dataclass
class Unresolved:
    """A requirement the run could not satisfy, with the reason why.

    ``reasons`` is grouped by *kind of problem* across **every** candidate
    session, not just whichever one happened to be tried last. With eight
    candidate sessions, "5 had no capacity recorded, 3 clashed with an existing
    commitment" tells the coordinator what to fix; the reasons of one arbitrary
    session do not.
    """

    requirement: Requirement
    reasons: list
    sessions_considered: int = 0
    #: ``{violation code: {"count": n, "example": str, "sessions": [...]}}``
    grouped: dict = field(default_factory=dict)

    @property
    def label(self) -> str:
        return self.requirement.label

    def summary_lines(self) -> list:
        """One line per distinct problem, with how many sessions it blocked."""
        lines = []
        for code, info in sorted(
            self.grouped.items(), key=lambda kv: (-kv[1]["count"], kv[0])
        ):
            count = info["count"]
            noun = "session" if count == 1 else "sessions"
            prefix = f"{count} {noun}: " if count > 1 else ""
            lines.append(f"{prefix}{info['example']}")
        return lines or list(self.reasons)


@dataclass
class Assignment:
    """A requirement the run placed."""

    requirement: Requirement
    session: Session
    status: str  # "added" | "retained" | "moved"
    moved_from: Session = None
    capacity_status: str = "ok"
    capacity_message: str = ""


@dataclass
class Removal:
    """A pre-existing link the run drops so a move can happen."""

    group: StudentGroup
    session: Session
    course: Course
    activity_type: str
    replaced_by: Session = None

    @property
    def label(self) -> str:
        return f"{self.group.code} - {self.course.code} {ActivityType(self.activity_type).label}"


@dataclass
class AllocationPlan:
    """The full, reviewable outcome of one allocation run."""

    semester: Semester
    scope: str
    assignments: list = field(default_factory=list)
    removals: list = field(default_factory=list)
    unresolved: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    unconfigured_courses: list = field(default_factory=list)
    search_limit_hit: bool = False
    scanned: int = 0
    duration_ms: int = 0
    run: AllocationRun = None

    # -- counts used by the summary and the applied run record ----------
    @property
    def added(self) -> int:
        return sum(1 for a in self.assignments if a.status == "added")

    @property
    def moved(self) -> int:
        return sum(1 for a in self.assignments if a.status == "moved")

    @property
    def retained(self) -> int:
        return sum(1 for a in self.assignments if a.status == "retained")

    @property
    def unresolved_count(self) -> int:
        return len(self.unresolved)

    @property
    def requirement_total(self) -> int:
        return len(self.assignments) + len(self.unresolved)

    def is_complete(self) -> bool:
        return not self.unresolved and not self.search_limit_hit

    def session_rollup(self) -> list:
        """One row per session the run touched, for the review table."""
        rows: dict = {}
        for assignment in self.assignments:
            session = assignment.session
            entry = rows.setdefault(
                session.pk,
                {
                    "session": session,
                    "groups": [],
                    "programmes": {},
                    "statuses": set(),
                },
            )
            entry["groups"].append(assignment.requirement.group)
            programme = assignment.requirement.group.programme
            entry["programmes"].setdefault(programme.code, 0)
            entry["programmes"][programme.code] += 1
            entry["statuses"].add(assignment.status)
        for entry in rows.values():
            entry["group_count"] = len(entry["groups"])
            entry["programme_count"] = len(entry["programmes"])
            entry["needed"] = required_capacity(entry["group_count"])
            entry["capacity_status"], entry["capacity_message"] = capacity_status(
                entry["session"], entry["group_count"]
            )
            entry["status"] = (
                "added"
                if entry["statuses"] == {"added"}
                else "retained"
                if entry["statuses"] == {"retained"}
                else "mixed"
            )
        return sorted(
            rows.values(),
            key=lambda e: (
                e["session"].day,
                e["session"].start_time,
                e["session"].course_code,
                e["session"].pk,
            ),
        )

    def snapshot(self) -> dict:
        """JSON-serialisable plan, stored on the AllocationRun for review."""
        return {
            "semester": str(self.semester),
            "semester_id": self.semester.pk,
            "scope": self.scope,
            "added": self.added,
            "moved": self.moved,
            "retained": self.retained,
            "removals": len(self.removals),
            "unresolved": self.unresolved_count,
            "requirement_total": self.requirement_total,
            "search_limit_hit": self.search_limit_hit,
            "scanned": self.scanned,
            "duration_ms": self.duration_ms,
            "complete": self.is_complete(),
            "warnings": list(self.warnings),
            "unconfigured_courses": [
                {
                    "code": entry["code"],
                    "name": entry["name"],
                    "programmes": entry["programmes"],
                    "group_count": entry["group_count"],
                    "variants": entry["variants"],
                }
                for entry in self.unconfigured_courses
            ],
            "assignments": [
                {
                    "group": a.requirement.group.code,
                    "programme": a.requirement.group.programme.code,
                    "course": a.requirement.course.code,
                    "course_name": a.requirement.course.name,
                    "activity": ActivityType(a.requirement.activity_type).label,
                    "ordinal": a.requirement.ordinal,
                    "status": a.status,
                    "session_pk": a.session.pk,
                    "session": _session_label(a.session),
                    "day": a.session.get_day_display(),
                    "time": (
                        f"{a.session.start_time:%H:%M}-{a.session.end_time:%H:%M}"
                    ),
                    "venue": a.session.venue.name if a.session.venue_id else "",
                    "moved_from_pk": a.moved_from.pk if a.moved_from else None,
                    "moved_from": _session_label(a.moved_from) if a.moved_from else "",
                    "group_count": _current_group_count(a.session),
                    "capacity_status": a.capacity_status,
                    "capacity_message": a.capacity_message,
                }
                for a in self.assignments
            ],
            "unresolved_items": [
                {
                    "group": u.requirement.group.code,
                    "programme": u.requirement.group.programme.code,
                    "course": u.requirement.course.code,
                    "course_name": u.requirement.course.name,
                    "activity": ActivityType(u.requirement.activity_type).label,
                    "ordinal": u.requirement.ordinal,
                    "sessions_considered": u.sessions_considered,
                    "reasons": u.summary_lines(),
                }
                for u in self.unresolved
            ],
        }

    @classmethod
    def from_snapshot(cls, data: dict):
        """Rebuild a template-friendly plan from a stored snapshot.

        Only used to re-render a saved run on the review page; the objects are
        plain dicts, not models, so nothing is written or trusted.
        """
        plan = cls(semester=None, scope=data.get("scope", "ALL"))
        plan.warnings = list(data.get("warnings", []))
        plan.unconfigured_courses = list(data.get("unconfigured_courses", []))
        plan.search_limit_hit = bool(data.get("search_limit_hit"))
        plan.scanned = int(data.get("scanned", 0))
        plan.duration_ms = int(data.get("duration_ms", 0))
        plan._snapshot = data
        return plan

    @property
    def snapshot_data(self) -> dict:
        return getattr(self, "_snapshot", None) or self.snapshot()


def _session_label(session) -> str:
    if session is None:
        return ""
    venue = session.venue.name if session.venue_id else "no venue"
    return (
        f"{session.get_day_display()} {session.start_time:%H:%M}-"
        f"{session.end_time:%H:%M} @ {venue}"
    )


# ──────────────────────────────────────────────
# The search
# ──────────────────────────────────────────────


class _SearchLimit(Exception):
    """Raised to unwind the DFS when the node budget is spent."""


def derive_requirements_from_sessions(semester: Semester = None):
    """Guess each course's requirement from the small-group sessions it has.

    A course with tutorial sessions in the timetable almost certainly requires
    every group to attend one, and the same for seminars and practicals — so
    the master timetable is good evidence for a starting point:

    ``CL111`` with only SEMINAR sessions   ->  ``{SEMINAR: 1}``
    ``MT111`` with TUTORIAL and PRACTICAL  ->  ``{TUTORIAL: 1, PRACTICAL: 1}``

    Two rules keep the guess honest. Only an activity that **has** a session is
    ever set, so this cannot invent a requirement nothing in the timetable
    could satisfy; and a course with no small-group session at all is returned
    in ``unresolved`` rather than being given a blank requirement, because
    whether it needs one is a judgement call, not a derivation.

    Only ever a *starting point*. The coordinator's own value always wins —
    :func:`_sync_course` and ``set_course_requirements`` skip any course that
    already has requirements — so re-running an import can never overwrite a
    decision that was made by hand.

    Returns ``(mapping, unresolved)``: ``{Course: {activity: count}}`` and
    ``[Course, ...]`` for the courses nothing could be said about.
    """
    sessions = Session.objects.all()
    if semester is not None:
        sessions = sessions.filter(semester=semester)
    available: dict = {}
    for code, activity in sessions.values_list("course_code", "activity_type"):
        if activity not in ALLOCATED_ACTIVITY_TYPES:
            continue
        available.setdefault(normalise_course_code(code), set()).add(activity)

    mapping, unresolved = {}, []
    for course in Course.objects.prefetch_related("activity_requirements").all():
        present = [
            activity
            for activity in ALLOCATED_ACTIVITY_TYPES
            if activity in available.get(normalise_course_code(course.code), ())
        ]
        if not present:
            unresolved.append(course)
        else:
            mapping[course] = {activity: 1 for activity in present}
    return mapping, unresolved


def build_requirements(semester: Semester, scope: str = "ALL"):
    """Requirements for the semester plus the courses that have no requirement.

    Returns ``(requirements, unconfigured)`` where ``unconfigured`` describes
    every course studied this semester whose requirements are blank — the
    "No allocation required - requirement not configured" notice.
    """
    activities = scope_activities(scope)
    courses = _courses_in_play(semester)
    groups = list(
        StudentGroup.objects.filter(
            programme__programme_courses__semester=semester.semester,
            programme__programme_courses__course__isnull=False,
        )
        .select_related("programme")
        .distinct()
    )

    by_course: dict = {}
    for group in groups:
        for row in ProgrammeCourse.objects.filter(
            programme_id=group.programme_id,
            semester=semester.semester,
            course__isnull=False,
        ).select_related("course"):
            by_course.setdefault(row.course_id, {"course": row.course, "groups": []})
            by_course[row.course_id]["groups"].append(group)

    sessions_by_key: dict = {}
    for session in Session.objects.filter(
        semester=semester, activity_type__in=activities
    ).select_related("venue"):
        key = (
            normalise_course_code(session.course_code),
            session.activity_type,
        )
        sessions_by_key.setdefault(key, []).append(session)

    requirements = []
    unconfigured = []
    for course_id in sorted(by_course):
        entry = by_course[course_id]
        course = entry["course"]
        groups_for_course = entry["groups"]
        configured = {
            activity: count
            for activity, count in course.requirements
            if activity in activities
        }
        if not configured:
            unconfigured.append(
                {
                    "course": course,
                    "code": course.code,
                    "name": course.name,
                    "programmes": sorted(
                        {g.programme.code for g in groups_for_course}
                    ),
                    "group_count": len(groups_for_course),
                    "variants": course.name_conflicts(),
                }
            )
            continue
        code = normalise_course_code(course.code)
        for activity in activities:
            count = configured.get(activity)
            if not count:
                continue
            candidates = sorted(
                sessions_by_key.get((code, activity), []),
                key=lambda s: (s.day, s.start_time, s.pk),
            )
            for group in groups_for_course:
                linked = set(
                    SessionGroup.objects.filter(
                        group_id=group.pk, session__in=candidates
                    ).values_list("session_id", flat=True)
                )
                for ordinal in range(1, count + 1):
                    requirements.append(
                        Requirement(
                            group=group,
                            course=course,
                            activity_type=activity,
                            ordinal=ordinal,
                            candidates=candidates,
                            existing=[
                                s for s in candidates if s.pk in linked
                            ],
                        )
                    )

    # Priority: seminar -> tutorial -> practical, most constrained first inside
    # each stage, then a stable label so identical input gives identical order.
    stage = {activity: index for index, activity in enumerate(activities)}
    requirements.sort(
        key=lambda r: (
            stage.get(r.activity_type, 99),
            len(r.candidates),
            r.course.code,
            r.group.programme.code,
            r.group.code,
            r.ordinal,
        )
    )
    return requirements, unconfigured


@dataclass
class SessionOption:
    """One candidate session for a requirement, and whether the group can take it.

    ``ok`` is the single answer to "can this group go here", and ``reasons``
    says why not when it cannot. ``free`` is split out from ``ok`` because the
    coordinator's question is usually narrower than "is this valid": *is the
    group actually free at that time?* A session can be free and still too
    small, and saying so separately is more useful than one blunt refusal.
    """

    session: Session
    ok: bool
    free: bool
    reasons: list
    already_assigned: bool = False
    group_count: int = 0
    capacity_status: str = "ok"
    capacity_message: str = ""

    @property
    def day_label(self) -> str:
        return self.session.get_day_display()

    @property
    def time_label(self) -> str:
        return f"{self.session.start_time:%H:%M}-{self.session.end_time:%H:%M}"

    @property
    def venue_label(self) -> str:
        return self.session.venue.name if self.session.venue_id else "no venue"


@dataclass
class RequirementStatus:
    """One requirement for one group: what it has, and what it could have."""

    requirement: Requirement
    assigned: Session = None
    options: list = field(default_factory=list)

    @property
    def is_met(self) -> bool:
        return self.assigned is not None

    @property
    def label(self) -> str:
        return self.requirement.label

    @property
    def activity_label(self) -> str:
        return ActivityType(self.requirement.activity_type).label

    @property
    def free_options(self) -> list:
        return [o for o in self.options if o.ok]

    @property
    def blocked_options(self) -> list:
        return [o for o in self.options if not o.ok]


@dataclass
class GroupStatus:
    """How far one group has got against the requirements of its courses."""

    group: StudentGroup
    semester: Semester
    entries: list = field(default_factory=list)
    unconfigured_courses: list = field(default_factory=list)
    detail_url: str = ""

    @property
    def total(self) -> int:
        return len(self.entries)

    @property
    def met(self) -> int:
        return sum(1 for e in self.entries if e.is_met)

    @property
    def outstanding(self) -> int:
        return self.total - self.met

    @property
    def percent(self) -> int:
        return round(self.met * 100 / self.total) if self.total else 0

    @property
    def is_complete(self) -> bool:
        return self.total > 0 and self.outstanding == 0

    @property
    def has_nothing_to_do(self) -> bool:
        """No requirement at all — nothing was ever asked of this group."""
        return self.total == 0

    @property
    def state(self) -> str:
        if self.is_complete:
            return "complete"
        if not self.total:
            return "none"
        return "partial" if self.met else "unassigned"

    def activity_progress(self) -> list:
        """``[(activity label, met, total), ...]`` in priority order."""
        buckets: dict = {}
        for entry in self.entries:
            key = entry.activity_label
            met, total = buckets.get(key, (0, 0))
            buckets[key] = (met + int(entry.is_met), total + 1)
        return [
            (label, buckets[label][0], buckets[label][1])
            for label in (
                ActivityType(a).label for a in ALLOCATED_ACTIVITY_TYPES
            )
            if label in buckets
        ]


def group_statuses(semester: Semester, scope: str = "ALL", groups=None):
    """Per-group allocation status, with every candidate session judged.

    This is the read-only counterpart of :func:`plan_allocation`: it never
    proposes a whole plan, it answers "where does each group stand, and where
    could it still go?". Each requirement carries its options, each option
    carrying the same verdict the allocator would give — because both call
    :func:`validate_assignment`, the guidance on this page can never offer a
    session the engine would refuse.
    """
    requirements, unconfigured = build_requirements(semester, scope)
    index = AvailabilityIndex(semester)
    occupancy: dict = {}
    for session_pk, group_pk in SessionGroup.objects.filter(
        session__semester=semester
    ).values_list("session_id", "group_id"):
        occupancy.setdefault(session_pk, set()).add(group_pk)

    wanted = None if groups is None else {g.pk for g in groups}
    unconfigured_by_group: dict = {}
    for entry in unconfigured:
        key = normalise_course_code(entry["code"])
        for row in ProgrammeCourse.objects.filter(
            course__code=entry["code"]
        ).values_list("programme_id", flat=True):
            for group in StudentGroup.objects.filter(
                programme_id=row
            ).values_list("pk", flat=True):
                unconfigured_by_group.setdefault(group, []).append(entry)

    statuses: dict = {}
    for requirement in requirements:
        if wanted is not None and requirement.group.pk not in wanted:
            continue
        status = statuses.get(requirement.group.pk)
        if status is None:
            status = GroupStatus(
                group=requirement.group,
                semester=semester,
                unconfigured_courses=unconfigured_by_group.get(
                    requirement.group.pk, []
                ),
            )
            statuses[requirement.group.pk] = status
        status.entries.append(_judge(requirement, semester, index, occupancy))

    # A group whose courses are all unconfigured still deserves a row, so the
    # board shows it as "nothing asked of it" rather than omitting it.
    if wanted is not None:
        for group in groups:
            statuses.setdefault(
                group.pk,
                GroupStatus(
                    group=group,
                    semester=semester,
                    unconfigured_courses=unconfigured_by_group.get(
                        group.pk, []
                    ),
                ),
            )
    return statuses


def _judge(requirement: Requirement, semester, index, occupancy: dict):
    """One requirement's current assignment plus every candidate's verdict.

    Each candidate goes through the same :func:`validate_assignment` the engine
    uses, so an option offered here can never be one the allocator would refuse.
    The session the group is already in is not re-validated — that one is the
    status, not a proposal.
    """
    group = requirement.group
    assigned = None
    for session in requirement.candidates:
        if group.pk in occupancy.get(session.pk, set()):
            assigned = session
            break

    options = []
    for session in requirement.candidates:
        already = assigned is not None and session.pk == assigned.pk
        if already:
            problems = []
        else:
            count_after = len(occupancy.get(session.pk, set()))
            if group.pk not in occupancy.get(session.pk, set()):
                count_after += 1
            problems = validate_assignment(
                group,
                requirement.course,
                session,
                semester,
                index,
                group_count_after=count_after,
            )
        group_count = len(occupancy.get(session.pk, ()))
        options.append(
            SessionOption(
                session=session,
                ok=not problems,
                free=not any(p.kind == KIND_AVAILABILITY for p in problems),
                reasons=problems,
                already_assigned=already,
                group_count=group_count + (0 if already else 1),
            )
        )
    options.sort(
        key=lambda o: (
            0 if o.already_assigned else 1,
            0 if o.ok else 1,
            o.day_label,
            o.time_label,
            o.session.pk,
        )
    )
    return RequirementStatus(
        requirement=requirement, assigned=assigned, options=options
    )


def _rank_candidates(requirement: Requirement, session: Session, index, occupancy):
    """Sort key for one candidate: lower is better.

    The tiers are the soft preferences, in order, and none of them can override
    a hard rule because they are only ever consulted for candidates that have
    already passed validation:

    1. keep a valid existing assignment
    2. balance group counts across sessions
    3. avoid timetable gaps and an overly concentrated day
    4. combine groups from different programmes

    Tiers 2 and 4 read the *current* occupancy, so this must be evaluated when
    the requirement is about to be placed rather than once up front — see
    ``try_place``. Ranking against a snapshot taken before anything was placed
    makes "balance" a no-op and piles every group into the lowest-numbered
    session.
    """
    group = requirement.group
    already = group.pk in occupancy.get(session.pk, set())
    count_after = len(occupancy.get(session.pk, ())) + (0 if already else 1)
    distinct_programmes = len(
        {
            g.programme_id
            for g in StudentGroup.objects.filter(
                pk__in=occupancy.get(session.pk, set())
            )
        }
    )
    return (
        0 if already else 1,
        count_after,
        index.commitments_on(group.pk, session.day) + (0 if already else 1),
        -distinct_programmes,
        session.pk,
    )


def _validate_all(requirement, session, semester, index, occupancy):
    """Hard checks for a candidate, evaluated against the partial plan."""
    count_after = len(occupancy.get(session.pk, set()))
    if requirement.group.pk not in occupancy.get(session.pk, set()):
        count_after += 1
    return validate_assignment(
        requirement.group,
        requirement.course,
        session,
        semester,
        index,
        group_count_after=count_after,
    )


def plan_allocation(semester: Semester, scope: str = "ALL", *, node_limit=None):
    """Compute the allocation for a semester without writing anything.

    Validates, searches with backtracking and returns an
    :class:`AllocationPlan`. Nothing is persisted, so the coordinator can review
    the whole plan -- proposals, moves, warnings and unresolved items -- before
    anything is applied.
    """
    started = time.monotonic()
    node_limit = int(node_limit or DEFAULT_NODE_LIMIT)

    requirements, unconfigured = build_requirements(semester, scope)
    plan = AllocationPlan(
        semester=semester, scope=scope, unconfigured_courses=unconfigured
    )
    if unconfigured:
        plan.warnings.append(
            f"{len(unconfigured)} course(s) have no required activities "
            f"configured and were not allocated. Set their requirements to "
            f"include them in a future run."
        )

    for requirement in requirements:
        if not requirement.candidates:
            plan.unresolved.append(
                Unresolved(
                    requirement=requirement,
                    reasons=[
                        f"No {ActivityType(requirement.activity_type).label.lower()} "
                        f"session exists for {requirement.course.code} in "
                        f"{semester}. Create the session, then run the "
                        f"allocator again."
                    ],
                )
            )

    placeable = [r for r in requirements if r.candidates]
    if not placeable:
        plan.duration_ms = int((time.monotonic() - started) * 1000)
        return plan

    index = AvailabilityIndex(semester)
    if index.unverifiable:
        plan.warnings.append(
            f"{len(index.unverifiable)} workshop record(s) name a day but no "
            f"time, so those days cannot be verified. Groups with such a "
            f"workshop are left unplaced on that day until the record is "
            f"corrected."
        )

    # Working state: which groups the plan has placed in which session so far.
    # ``initial_occupancy`` is the untouched starting point, kept so the result
    # can tell a link that already existed from one the run created.
    occupancy: dict = {}
    for session_pk, group_pk in SessionGroup.objects.filter(
        session__semester=semester
    ).values_list("session_id", "group_id"):
        occupancy.setdefault(session_pk, set()).add(group_pk)
    initial_occupancy = {pk: set(groups) for pk, groups in occupancy.items()}

    # The candidate ORDER is re-derived each time a requirement is reached,
    # against the occupancy as it stands then (see ``try_place``). Ranking it
    # once up front would freeze "balance the groups across sessions" against
    # an empty timetable and stack every group into the first session.
    candidate_pool = [requirement.candidates for requirement in placeable]

    # Two practicals for one group must land in two *different* sessions. The
    # already-linked sessions of one (group, course, activity) are therefore
    # handed out one per ordinal, best first, so a course needing two of
    # something keeps what it had instead of moving both.
    blocked = _block_repeat_sessions(placeable)

    state = {
        "chosen": [None] * len(placeable),
        "problems": [None] * len(placeable),
        # Accumulated across every candidate tried, so the report can say how
        # many sessions each distinct problem blocked instead of showing the
        # reasons of whichever session happened to be tried last.
        "grouped": [{} for _ in placeable],
        "considered": [0] * len(placeable),
        "nodes": 0,
    }

    def _record_problems(position, session, problems):
        for problem in problems:
            entry = state["grouped"][position].setdefault(
                problem.code,
                {"count": 0, "example": problem.message, "sessions": []},
            )
            if session.pk not in entry["sessions"]:
                entry["sessions"].append(session.pk)
                entry["count"] += 1

    def try_place(position):
        if position == len(placeable):
            return True
        requirement = placeable[position]
        group = requirement.group
        not_allowed = blocked[id(requirement)]
        # Re-ranked on arrival, because "balance the groups" and "combine
        # programmes" are only meaningful against what is already placed.
        order = sorted(
            candidate_pool[position],
            key=lambda s: _rank_candidates(requirement, s, index, occupancy),
        )
        for session in order:
            if session.pk in not_allowed:
                continue
            state["nodes"] += 1
            if state["nodes"] > node_limit:
                raise _SearchLimit
            problems = _validate_all(requirement, session, semester, index, occupancy)
            state["considered"][position] += 1
            if problems:
                state["problems"][position] = problems
                _record_problems(position, session, problems)
                continue
            occupancy.setdefault(session.pk, set()).add(group.pk)
            index.add_session(group.pk, session)
            state["chosen"][position] = session
            if try_place(position + 1):
                return True
            state["chosen"][position] = None
            occupancy[session.pk].discard(group.pk)
            index.drop_session(group.pk, session.pk, session.day)
        # Nothing left for this requirement here. Leave it unresolved, keeping
        # every reason gathered on the way, and carry on with the rest so the
        # coordinator still gets the assignments that do work.
        return try_place(position + 1)

    limit_hit = False
    try:
        try_place(0)
    except _SearchLimit:
        limit_hit = True
    plan.search_limit_hit = limit_hit
    plan.scanned = state["nodes"]

    if limit_hit:
        plan.warnings.append(
            f"The search stopped after {node_limit:,} steps, so this plan is "
            f"the best found so far -- it is not proof that no valid "
            f"allocation exists. Narrow the semester or the activity scope "
            f"and run again."
        )

    placed_by_position = {}
    for position, requirement in enumerate(placeable):
        session = state["chosen"][position]
        if session is not None:
            placed_by_position[position] = session
            continue
        problems = state["problems"][position] or [
            Violation(
                "search-exhausted",
                KIND_DATA,
                "No combination of sessions could satisfy every requirement; "
                "an earlier seminar or tutorial choice had to be given up.",
            )
        ]
        grouped = state["grouped"][position]
        plan.unresolved.append(
            Unresolved(
                requirement=requirement,
                # One message per *distinct* problem across every candidate, not
                # the reasons of whichever session happened to be tried last.
                reasons=[
                    info["example"] for info in grouped.values()
                ] or [p.message for p in problems],
                sessions_considered=state["considered"][position],
                grouped=grouped,
            )
        )

    _resolve_moves(plan, placeable, placed_by_position, initial_occupancy)
    plan.warnings.extend(_extra_link_warnings(plan, placeable, placed_by_position))
    plan.duration_ms = int((time.monotonic() - started) * 1000)
    return plan


def _block_repeat_sessions(placeable) -> dict:
    """Sessions each requirement may never use, so one group is never doubled.

    A group needing two practicals must attend two *different* practicals. The
    sessions it is already linked to are handed out one per ordinal (best rank
    first) and every other ordinal is blocked from them, so the search keeps
    what the group already has instead of moving it wholesale.
    """
    by_key: dict = {}
    for requirement in placeable:
        by_key.setdefault(
            (requirement.group.pk, requirement.course.pk, requirement.activity_type),
            [],
        ).append(requirement)
    blocked = {id(r): set() for r in placeable}
    for requirements in by_key.values():
        requirements.sort(key=lambda r: r.ordinal)
        if len(requirements) < 2:
            continue
        existing = requirements[0].existing
        for position, requirement in enumerate(requirements):
            keep = existing[position] if position < len(existing) else None
            for session in existing:
                if keep is None or session.pk != keep.pk:
                    blocked[id(requirement)].add(session.pk)
    return blocked


def _resolve_moves(plan, placeable, placed, initial_occupancy):
    """Work out which placements are retained/added/moved, and what to unlink.

    "Moved" means the group already sat in a *different* eligible session for
    the same course and activity, and that link is dropped in favour of the
    new one. A pre-existing link that is simply not used any more (the group
    was double-allocated) is reported but never removed: pruning links the
    coordinator did not ask about is not this feature's job.
    """
    by_key: dict = {}
    for position, requirement in enumerate(placeable):
        if position not in placed:
            continue
        by_key.setdefault(
            (requirement.group.pk, requirement.course.pk, requirement.activity_type),
            [],
        ).append((requirement, placed[position]))

    for entries in by_key.values():
        entries.sort(key=lambda pair: pair[0].ordinal)
        existing_pks = {s.pk for s in entries[0][0].existing}
        chosen_pks = {session.pk for _, session in entries}
        unused = [s for s in entries[0][0].existing if s.pk not in chosen_pks]
        # New placements (sessions the group was not in) are the ones that can
        # take over from a link being dropped; pair them in ordinal order.
        fresh = [
            (requirement, session)
            for requirement, session in entries
            if session.pk not in existing_pks
        ]
        # Keyed on identity: a Requirement is a mutable dataclass and therefore
        # unhashable.
        pairings = {
            id(requirement): old
            for (requirement, _), old in zip(fresh, unused)
        }

        for requirement, session in entries:
            was_linked = requirement.group.pk in initial_occupancy.get(
                session.pk, set()
            )
            replacement = pairings.get(id(requirement))
            if was_linked:
                status = "retained"
                count_after = len(initial_occupancy.get(session.pk, set()))
            else:
                status = "moved" if replacement is not None else "added"
                count_after = len(initial_occupancy.get(session.pk, set())) + 1
            cap_status, cap_message = capacity_status(session, count_after)
            assignment = Assignment(
                requirement=requirement,
                session=session,
                status=status,
                moved_from=replacement,
                capacity_status=cap_status,
                capacity_message=cap_message,
            )
            plan.assignments.append(assignment)
            if assignment.moved_from is not None:
                plan.removals.append(
                    Removal(
                        group=requirement.group,
                        session=assignment.moved_from,
                        course=requirement.course,
                        activity_type=requirement.activity_type,
                        replaced_by=session,
                    )
                )
        for leftover in unused[len(pairings):]:
            requirement = entries[0][0]
            plan.warnings.append(
                f"{requirement.group.code} is linked to "
                f"{_session_label(leftover)} for {requirement.course.code} "
                f"but that session is not used by the plan, so the extra link "
                f"was left in place."
            )


def _extra_link_warnings(plan, placeable, placed) -> list:
    """Report existing links on a requirement that ended up with no placement."""
    warnings = []
    planned = {}
    for position, requirement in enumerate(placeable):
        if position in placed:
            continue
        for session in requirement.existing:
            warnings.append(
                f"{requirement.group.code} stays linked to "
                f"{_session_label(session)} for {requirement.course.code}; the "
                f"run could not confirm that allocation."
            )
    return warnings


# ──────────────────────────────────────────────
# Applying and reverting a run
# ──────────────────────────────────────────────


def _link_state(session: Session, group: StudentGroup):
    """Snapshot of one SessionGroup link, or ``None`` when there is none."""
    if session is None:
        return None
    return {
        "session_pk": session.pk,
        "session": _session_label(session),
        "course_code": session.course_code,
        "activity_type": session.activity_type,
        "day": session.day,
        "start_time": session.start_time.strftime("%H:%M"),
        "end_time": session.end_time.strftime("%H:%M"),
        "venue": session.venue.name if session.venue_id else "",
    }


def save_plan(plan: AllocationPlan, algorithm: str = "smart") -> AllocationRun:
    """Persist a previewed run, with the link-level before/after states.

    The changes are recorded here, at plan time, because "before" means the
    state the timetable was in when the coordinator looked at the plan. That is
    what a revert has to put back.
    """
    run = AllocationRun.objects.create(
        semester=plan.semester,
        scope=plan.scope,
        status=AllocationStatus.PREVIEWED,
        search_limit_hit=plan.search_limit_hit,
        assigned=plan.added,
        moved=plan.moved,
        retained=plan.retained,
        removed=len(plan.removals),
        unresolved=plan.unresolved_count,
        summary=json.dumps(plan.snapshot()),
        algorithm=algorithm,
    )
    for removal in plan.removals:
        run.changes.create(
            group=removal.group,
            session=removal.session,
            action=AllocationChange.Action.REMOVE,
            course_code=removal.session.course_code,
            activity_type=removal.session.activity_type,
            before=json.dumps(_link_state(removal.session, removal.group)),
            after="",
        )
    for assignment in plan.assignments:
        if assignment.status == "retained":
            # Nothing about this link changes, so it is not recorded as a
            # change — a revert must not touch a link the run did not create.
            continue
        run.changes.create(
            group=assignment.requirement.group,
            session=assignment.session,
            action=AllocationChange.Action.ADD,
            course_code=assignment.session.course_code,
            activity_type=assignment.session.activity_type,
            before="",
            after=json.dumps(_link_state(assignment.session, assignment.requirement.group)),
        )
    return run


@transaction.atomic
def apply_run(run: AllocationRun) -> dict:
    """Write a saved plan: create the links, drop the ones being moved off.

    One transaction, so a failure part-way through leaves the timetable exactly
    as it was. Valid assignments are applied even when some requirements are
    unresolved -- those stay listed for the coordinator to resolve by hand.
    """
    from django.utils import timezone

    if run.status == AllocationStatus.APPLIED:
        return {
            "ok": False,
            "added": 0,
            "moved": 0,
            "retained": 0,
            "removed": 0,
            "message": "This run has already been applied.",
        }

    removed = added = 0
    # Removals first: a group is being moved off a session, and doing that
    # before the new link keeps the venue headroom honest inside the
    # transaction.
    for change in run.changes.filter(action=AllocationChange.Action.REMOVE):
        deleted, _ = SessionGroup.objects.filter(
            session_id=change.session_id, group_id=change.group_id
        ).delete()
        removed += deleted
    for change in run.changes.filter(action=AllocationChange.Action.ADD):
        if change.session_id is None:
            continue
        _, created = SessionGroup.objects.get_or_create(
            session_id=change.session_id, group_id=change.group_id
        )
        added += int(created)

    run.status = AllocationStatus.APPLIED
    run.applied_at = timezone.now()
    run.assigned = added
    run.removed = removed
    run.save()
    return {
        "ok": True,
        "added": added,
        "moved": int(added and removed and min(added, removed) or 0),
        "retained": run.retained,
        "removed": removed,
        "unresolved": run.unresolved,
        "message": (
            f"Applied: {added} assignment(s) added, {removed} previous "
            f"assignment(s) removed, {run.retained} kept. "
            f"{run.unresolved} requirement(s) remain unresolved."
        ),
    }


@transaction.atomic
def revert_run(run: AllocationRun) -> dict:
    """Undo an applied run, refusing to touch anything edited since.

    Every change is checked against the state the run left behind first. A
    single drifted link aborts the whole revert with the conflicts listed, so a
    later manual edit is never silently overwritten.
    """
    from django.utils import timezone

    if run.status != AllocationStatus.APPLIED:
        return {
            "ok": False,
            "conflicts": [],
            "message": "Only an applied run can be reverted.",
        }

    conflicts = []
    for change in run.changes.select_related("session", "group"):
        linked = SessionGroup.objects.filter(
            session_id=change.session_id, group_id=change.group_id
        ).exists()
        if change.action == AllocationChange.Action.ADD and not linked:
            conflicts.append(
                f"{change.group} is no longer linked to the "
                f"{change.activity_type} session this run added."
            )
        if change.action == AllocationChange.Action.REMOVE and linked:
            conflicts.append(
                f"{change.group} has been linked again to the "
                f"{change.activity_type} session this run moved it off."
            )
    if conflicts:
        return {
            "ok": False,
            "conflicts": conflicts,
            "message": (
                f"{len(conflicts)} assignment(s) have been edited since this "
                f"run was applied. Nothing was reverted - review the conflicts "
                f"and redo the changes by hand if needed."
            ),
        }

    removed = restored = 0
    for change in run.changes.select_related("session", "group"):
        if change.action == AllocationChange.Action.ADD:
            deleted, _ = SessionGroup.objects.filter(
                session_id=change.session_id, group_id=change.group_id
            ).delete()
            removed += deleted
        else:
            _, created = SessionGroup.objects.get_or_create(
                session_id=change.session_id, group_id=change.group_id
            )
            restored += int(created)

    run.status = AllocationStatus.REVERTED
    run.reverted_at = timezone.now()
    run.save()
    return {
        "ok": True,
        "conflicts": [],
        "removed": removed,
        "restored": restored,
        "message": (
            f"Run reverted: {removed} assignment(s) removed and "
            f"{restored} previous assignment(s) restored."
        ),
    }


# ──────────────────────────────────────────────
# Manual assignment (same validator as the engine)
# ──────────────────────────────────────────────


def find_course_for_session(session: Session):
    """The shared Course a session belongs to, matched on normalised code."""
    code = normalise_course_code(session.course_code)
    if not code:
        return None
    return Course.objects.filter(code__iexact=code).first()


def validate_manual_assignment(group, session: Session, index=None):
    """Validate one coordinator-entered (group, session) pair.

    Returns ``(course, violations)``. ``course`` is ``None`` when the session's
    course code names no known course, which is itself reported rather than
    guessed around.
    """
    if session is None:
        return None, [Violation("no-session", KIND_DATA, "Choose a session.")]
    if group is None:
        return None, [Violation("no-group", KIND_DATA, "Choose a group.")]
    course = find_course_for_session(session)
    if course is None:
        return None, [
            Violation(
                "unknown-course",
                KIND_ELIGIBILITY,
                f"No course record matches {session.course_code}. Create the "
                f"course (or fix the session's course code) first.",
            )
        ]
    if index is None:
        index = AvailabilityIndex(session.semester)
    # The group being added is not in the session's current headcount, so the
    # capacity check has to include it — otherwise a manual assignment could
    # quietly push a room over its seat count.
    already_linked = SessionGroup.objects.filter(
        session_id=session.pk, group_id=group.pk
    ).exists()
    count_after = _current_group_count(session) + (0 if already_linked else 1)
    return course, validate_assignment(
        group,
        course,
        session,
        session.semester,
        index,
        group_count_after=count_after,
    )


def manual_assign(group, session: Session, index=None):
    """Link a group to a session after validating it, or refuse with reasons.

    Returns ``(linked, violations)``. This is the only way a coordinator can
    place a group by hand, and it uses exactly the checks the automatic run
    uses, so a hand-made assignment can never be one the engine would reject.
    """
    course, problems = validate_manual_assignment(group, session, index=index)
    if problems:
        return False, problems
    SessionGroup.objects.get_or_create(session=session, group=group)
    return True, []


def manual_unassign(group, session: Session):
    """Remove a link. Always allowed: un-assigning cannot overfill anything."""
    deleted, _ = SessionGroup.objects.filter(
        session=session, group=group
    ).delete()
    return deleted


def courses_missing_requirements(semester: Semester = None):
    """Courses studied this semester with blank requirements, for review."""
    queryset = Course.objects.filter(
        programme_courses__semester__isnull=False
    ).distinct()
    if semester is not None:
        queryset = queryset.filter(
            programme_courses__semester=semester.semester
        ).distinct()
    return [
        course
        for course in queryset.select_related().prefetch_related(
            "activity_requirements"
        )
        if not course.has_requirements()
    ]


# ──────────────────────────────────────────────
# Min-Conflicts Allocation Algorithm (Layer 1)
# ──────────────────────────────────────────────


def plan_allocation_min_conflicts(
    semester: Semester, scope: str = "ALL", *, max_iterations=50000, random_restarts=5
):
    """Min-conflicts repair-based allocation algorithm.

    This algorithm uses a repair-based approach:
    1. Greedily assign all requirements (may have conflicts)
    2. Iteratively repair conflicts by reassigning the most conflicted requirements
    3. Use random restarts to escape local optima

    Returns an AllocationPlan compatible with the existing infrastructure.
    """
    import random

    started = time.monotonic()
    requirements, unconfigured = build_requirements(semester, scope)
    plan = AllocationPlan(
        semester=semester, scope=scope, unconfigured_courses=unconfigured
    )
    if unconfigured:
        plan.warnings.append(
            f"{len(unconfigured)} course(s) have no required activities "
            f"configured and were not allocated. Set their requirements to "
            f"include them in a future run."
        )

    for requirement in requirements:
        if not requirement.candidates:
            plan.unresolved.append(
                Unresolved(
                    requirement=requirement,
                    reasons=[
                        f"No {ActivityType(requirement.activity_type).label.lower()} "
                        f"session exists for {requirement.course.code} in "
                        f"{semester}. Create the session, then run the "
                        f"allocator again."
                    ],
                )
            )

    placeable = [r for r in requirements if r.candidates]
    if not placeable:
        plan.duration_ms = int((time.monotonic() - started) * 1000)
        return plan

    index = AvailabilityIndex(semester)
    if index.unverifiable:
        plan.warnings.append(
            f"{len(index.unverifiable)} workshop record(s) name a day but no "
            f"time, so those days cannot be verified. Groups with such a "
            f"workshop are left unplaced on that day until the record is "
            f"corrected."
        )

    # Initial occupancy from existing assignments
    occupancy: dict = {}
    for session_pk, group_pk in SessionGroup.objects.filter(
        session__semester=semester
    ).values_list("session_id", "group_id"):
        occupancy.setdefault(session_pk, set()).add(group_pk)
    initial_occupancy = {pk: set(groups) for pk, groups in occupancy.items()}

    # Block repeat sessions (same as standard allocator)
    blocked = _block_repeat_sessions(placeable)

    best_plan = None
    best_unresolved = len(placeable)

    for restart in range(random_restarts):
        # Reset occupancy for this restart
        occupancy = {pk: set(groups) for pk, groups in initial_occupancy.items()}
        index = AvailabilityIndex(semester)  # Rebuild index

        # Phase 1: Greedy initial assignment
        assignments = {}  # requirement.key -> session
        for requirement in placeable:
            not_allowed = blocked[id(requirement)]
            # Try candidates in ranked order
            candidates = sorted(
                requirement.candidates,
                key=lambda s: _rank_candidates(requirement, s, index, occupancy),
            )
            assigned = False
            for session in candidates:
                if session.pk in not_allowed:
                    continue
                count_after = len(occupancy.get(session.pk, set())) + 1
                problems = validate_assignment(
                    requirement.group,
                    requirement.course,
                    session,
                    semester,
                    index,
                    group_count_after=count_after,
                )
                if not problems:
                    occupancy.setdefault(session.pk, set()).add(requirement.group.pk)
                    index.add_session(requirement.group.pk, session)
                    assignments[requirement.key] = session
                    assigned = True
                    break
            if not assigned:
                # Leave unassigned for now
                assignments[requirement.key] = None

        # Phase 2: Min-conflicts repair
        for iteration in range(max_iterations):
            # Check if all requirements are satisfied
            unresolved_requirements = [
                r for r, s in assignments.items() if s is None
            ]
            if not unresolved_requirements:
                break  # Complete solution found

            # Find conflicted requirements (those with conflicts or unassigned)
            conflicted = []
            for req_key, session in assignments.items():
                # Find the requirement object for this key
                requirement = next((r for r in placeable if r.key == req_key), None)
                if not requirement:
                    continue
                if session is None:
                    conflicted.append(requirement)
                    continue
                # Check if this assignment still valid
                count_after = len(occupancy.get(session.pk, set()))
                problems = validate_assignment(
                    requirement.group,
                    requirement.course,
                    session,
                    semester,
                    index,
                    group_count_after=count_after,
                )
                if problems:
                    conflicted.append(requirement)

            if not conflicted:
                break  # No conflicts

            # Pick a random conflicted requirement
            requirement = random.choice(conflicted)

            # Find the session that minimizes conflicts
            not_allowed = blocked[id(requirement)]
            candidates = sorted(
                requirement.candidates,
                key=lambda s: _rank_candidates(requirement, s, index, occupancy),
            )

            best_session = None
            best_conflicts = float("inf")

            for session in candidates:
                if session.pk in not_allowed:
                    continue

                # Count conflicts if we move here
                conflicts = 0
                count_after = len(occupancy.get(session.pk, set())) + 1
                problems = validate_assignment(
                    requirement.group,
                    requirement.course,
                    session,
                    semester,
                    index,
                    group_count_after=count_after,
                )
                conflicts += len(problems)

                if conflicts < best_conflicts:
                    best_conflicts = conflicts
                    best_session = session

            # Apply the best move
            if best_session:
                # Remove old assignment if exists
                old_session = assignments.get(requirement.key)
                if old_session:
                    occupancy[old_session.pk].discard(requirement.group.pk)
                    index.drop_session(requirement.group.pk, old_session.pk, old_session.day)

                # Add new assignment
                if best_conflicts == 0:  # Only add if it's valid
                    occupancy.setdefault(best_session.pk, set()).add(requirement.group.pk)
                    index.add_session(requirement.group.pk, best_session)
                    assignments[requirement.key] = best_session
                else:
                    assignments[requirement.key] = None
            else:
                assignments[requirement.key] = None

        # Check if this restart was better
        unresolved_count = sum(1 for s in assignments.values() if s is None)
        if unresolved_count < best_unresolved:
            best_unresolved = unresolved_count
            best_plan = (assignments.copy(), occupancy.copy())

            if unresolved_count == 0:
                break  # Perfect solution found

    # Use the best solution found
    if best_plan:
        assignments, occupancy = best_plan
    else:
        # Use the last attempt if no restarts succeeded
        assignments = {r.key: None for r in placeable}
        occupancy = initial_occupancy.copy()

    # Build the plan from the final assignments
    placed_by_position = {}
    for position, requirement in enumerate(placeable):
        session = assignments.get(requirement.key)
        if session is not None:
            placed_by_position[position] = session
        else:
            # Mark as unresolved
            plan.unresolved.append(
                Unresolved(
                    requirement=requirement,
                    reasons=[
                        "Min-conflicts algorithm could not find a valid "
                        "assignment for this requirement after multiple "
                        "restarts and iterations."
                    ],
                    sessions_considered=len(requirement.candidates),
                )
            )

    _resolve_moves(plan, placeable, placed_by_position, initial_occupancy)
    plan.warnings.extend(_extra_link_warnings(plan, placeable, placed_by_position))
    plan.duration_ms = int((time.monotonic() - started) * 1000)
    plan.scanned = len(placeable) * max_iterations * random_restarts  # Approximate

    return plan


# ──────────────────────────────────────────────
# OR-Tools Constraint Programming Allocation (Layer 2)
# ──────────────────────────────────────────────


def plan_allocation_ortools(
    semester: Semester, scope: str = "ALL", *, time_limit_seconds=30
):
    """Constraint programming allocation using Google OR-Tools.

    This uses OR-Tools CP-SAT solver which provides:
    - Advanced constraint propagation
    - Sophisticated search strategies
    - Guaranteed optimality within time limits
    - Better handling of complex constraints

    Returns an AllocationPlan compatible with the existing infrastructure.
    """
    try:
        from ortools.sat.python import cp_model
    except ImportError:
        # OR-Tools not available, fall back to standard allocator
        plan = AllocationPlan(semester=semester, scope=scope)
        plan.warnings.append(
            "OR-Tools library is not installed. Falling back to standard "
            "allocation. Install with: pip install ortools"
        )
        return plan_allocation(semester, scope)

    started = time.monotonic()
    requirements, unconfigured = build_requirements(semester, scope)
    plan = AllocationPlan(
        semester=semester, scope=scope, unconfigured_courses=unconfigured
    )
    if unconfigured:
        plan.warnings.append(
            f"{len(unconfigured)} course(s) have no required activities "
            f"configured and were not allocated. Set their requirements to "
            f"include them in a future run."
        )

    for requirement in requirements:
        if not requirement.candidates:
            plan.unresolved.append(
                Unresolved(
                    requirement=requirement,
                    reasons=[
                        f"No {ActivityType(requirement.activity_type).label.lower()} "
                        f"session exists for {requirement.course.code} in "
                        f"{semester}. Create the session, then run the "
                        f"allocator again."
                    ],
                )
            )

    placeable = [r for r in requirements if r.candidates]
    if not placeable:
        plan.duration_ms = int((time.monotonic() - started) * 1000)
        return plan

    index = AvailabilityIndex(semester)
    if index.unverifiable:
        plan.warnings.append(
            f"{len(index.unverifiable)} workshop record(s) name a day but no "
            f"time, so those days cannot be verified. Groups with such a "
            f"workshop are left unplaced on that day until the record is "
            f"corrected."
        )

    # Build CP model
    model = cp_model.CpModel()

    # Create variables: each requirement -> session index
    # Map requirements to variable indices
    req_to_var = {}
    var_to_req = {}
    for i, requirement in enumerate(placeable):
        var = model.NewIntVar(0, len(requirement.candidates) - 1, f"req_{i}")
        req_to_var[requirement.key] = var
        var_to_req[i] = requirement

    # Initial occupancy from existing assignments
    initial_occupancy: dict = {}
    for session_pk, group_pk in SessionGroup.objects.filter(
        session__semester=semester
    ).values_list("session_id", "group_id"):
        initial_occupancy.setdefault(session_pk, set()).add(group_pk)

    # Block repeat sessions
    blocked = _block_repeat_sessions(placeable)

    # Add constraints
    # 1. Each requirement must be assigned to a valid session
    for requirement in placeable:
        var = req_to_var[requirement.key]
        not_allowed = blocked[id(requirement)]
        valid_indices = [
            i
            for i, session in enumerate(requirement.candidates)
            if session.pk not in not_allowed
        ]
        if valid_indices:
            # Only allow valid session indices
            model.AddAllowedAssignments(var, valid_indices)

    # 2. No conflicts between requirements
    # Build a conflict graph: requirements that conflict if assigned to certain sessions
    for i, req1 in enumerate(placeable):
        for j, req2 in enumerate(placeable):
            if i >= j:
                continue  # Avoid duplicate checks

            var1 = req_to_var[req1.key]
            var2 = req_to_var[req2.key]

            # Check if these requirements can conflict
            for idx1, session1 in enumerate(req1.candidates):
                for idx2, session2 in enumerate(req2.candidates):
                    # Check if assigning req1 to session1 and req2 to session2 would conflict
                    if _would_conflict(req1, req2, session1, session2, index):
                        # Add constraint: not (var1 == idx1 AND var2 == idx2)
                        # Implemented as: var1 != idx1 OR var2 != idx2
                        model.AddBoolOr(
                            [var1 != idx1, var2 != idx2]
                        )

    # 3. Capacity constraints
    # For each session, limit the number of groups assigned
    session_to_reqs = {}
    for requirement in placeable:
        for idx, session in enumerate(requirement.candidates):
            session_to_reqs.setdefault(session.pk, []).append((requirement, idx))

    for session_pk, req_indices in session_to_reqs.items():
        session = req_indices[0][0].candidates[req_indices[0][1]]
        if session.venue and session.venue.capacity:
            max_groups = session.venue.capacity // STUDENTS_PER_GROUP
            # Create a boolean variable for each potential assignment
            assigned_vars = []
            for requirement, idx in req_indices:
                var = req_to_var[requirement.key]
                is_assigned = model.NewBoolVar(f"assigned_{requirement.group.pk}_{session_pk}")
                model.Add(var == idx).OnlyEnforceIf(is_assigned)
                model.Add(var != idx).OnlyEnforceIf(is_assigned.Not())
                assigned_vars.append(is_assigned)

            # Sum of assigned vars <= max_groups
            model.Add(sum(assigned_vars) <= max_groups)

    # Solve with time limit
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_seconds
    solver.parameters.num_search_workers = 8  # Use multiple threads

    result = solver.Solve(model)

    if result == cp_model.OPTIMAL or result == cp_model.FEASIBLE:
        # Extract solution
        placed_by_position = {}
        for i, requirement in enumerate(placeable):
            var = req_to_var[requirement.key]
            session_idx = solver.Value(var)
            session = requirement.candidates[session_idx]
            placed_by_position[i] = session

        # Build plan
        _resolve_moves(plan, placeable, placed_by_position, initial_occupancy)
        plan.warnings.extend(_extra_link_warnings(plan, placeable, placed_by_position))
        plan.duration_ms = int((time.monotonic() - started) * 1000)
        plan.scanned = solver.NumBranches()

        # Add unresolved requirements that weren't placed
        for position, requirement in enumerate(placeable):
            if position not in placed_by_position:
                plan.unresolved.append(
                    Unresolved(
                        requirement=requirement,
                        reasons=[
                            "OR-Tools solver could not find a valid assignment "
                            "within the time limit."
                        ],
                        sessions_considered=len(requirement.candidates),
                    )
                )
    else:
        # Solver failed, fall back to unresolved
        for requirement in placeable:
            plan.unresolved.append(
                Unresolved(
                    requirement=requirement,
                    reasons=[
                        "OR-Tools solver failed to find a feasible solution. "
                        "Try increasing the time limit or use a different algorithm."
                    ],
                    sessions_considered=len(requirement.candidates),
                )
            )
        plan.duration_ms = int((time.monotonic() - started) * 1000)

    return plan


def _would_conflict(req1, req2, session1, session2, index):
    """Check if two requirements would conflict if assigned to these sessions."""
    # Same group can't be in two places at once
    if req1.group.pk == req2.group.pk:
        if session1.day == session2.day:
            # Check time overlap
            if not (session1.end_time <= session2.start_time or session2.end_time <= session1.start_time):
                return True

    # Check availability conflicts using the index's public method
    for req, session in [(req1, session1), (req2, session2)]:
        # Check if the group has any conflicts at this session time
        if index.conflicts_excluding(req.group.pk, session.day, session.start_time, session.end_time, session.pk):
            return True

    return False
