"""Deterministic synthetic data for developing, testing and demoing the audit.

About 80 student groups across 8 programmes (10 per programme), 40 configured
courses plus one deliberately unconfigured, 25 venues, one semester of
timetabled sessions, an applied allocation run with a plan snapshot, and
workshop / technical-drawing allocations.

The shape is realistic on purpose, because an audit of an unrealistic
timetable finds nothing useful. Each course is offered to two programmes, so 20
of the 80 groups study it. A seminar seats one group (30 students in a 30-seat
room); a tutorial and a practical seat two (60 students), and one class in four
is put in a hall, so the venue section has genuine right-sizing observations
alongside rooms that are genuinely full.

Six problems are **planted** so the golden test can assert each is found with
the right severity:

======================= ================================================
Packed Wednesday        every EE group gets a Wednesday afternoon of work
4-hour gap              ``ME C3`` sits 11:00-12:00 then 16:00-18:00
Over-capacity room      ``CH C1`` + ``CH C2`` share a 30-seat seminar room
Unresolved requirement  ``MT455`` seminar — no such session is timetabled
Missing venue           an ``SC121`` practical with no room at all
Overlapping pair        ``EE C1`` is in two tutorials that overlap by 30 min
======================= ================================================

Seeded with a fixed :class:`random.Random`, so every run produces identical
input and therefore an identical data hash.

This module writes real rows. Run it through ``manage.py audit_fixture`` or
``manage.py audit_analyse_fixture``, which do so inside a throwaway test
database, never the working one.
"""

from __future__ import annotations

import json
import math
import random
from collections import defaultdict
from datetime import time

from core.models import (
    ActivityType,
    AllocationRun,
    AllocationScope,
    AllocationStatus,
    Course,
    CourseActivityRequirement,
    Day,
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    Venue,
    WorkshopAllocation,
)

SEED = 20240915
N_GROUPS = 80
GROUPS_PER_PROGRAMME = 10
N_VENUES = 25
SMALL_ACTIVITIES = ("SEMINAR", "TUTORIAL", "PRACTICAL")

#: 30 students per group, matching core.group_allocation.STUDENTS_PER_GROUP.
STUDENTS_PER_GROUP = 30

PROGRAMMES = [
    ("EE", "BSc. Electrical Engineering"),
    ("ME", "BSc. Mechanical Engineering"),
    ("CE", "BSc. Civil Engineering"),
    ("CH", "BSc. Chemical and Process Engineering"),
    ("CS", "BSc. Computer Science"),
    ("AR", "BSc. Architecture"),
    ("ST", "BSc. Statistics"),
    ("NM", "BSc. Nautical Science"),
]

#: Two groups per class. A 30-student group in a 60-70 seat room is the
#: healthy, normal case, and it is what makes the deliberately bad choices
#: (a 30-seat room for two groups, a 300-seat hall for two) stand out.
GROUPS_PER_CLASS = 2

# A seminar and a tutorial are per group; a practical is commonly taken by two
# groups in one lab, which is the only place the fixture needs a room bigger
# than a seminar room.
VENUE_SPECS = [
    # Halls. Every lecture goes in one of these, and a hall is also where the
    # deliberately under-used *small* classes go, so they can be spotted as
    # oversized bookings. All five seat a whole cohort of ten groups (300
    # students), which is what a one-lecture-a-week course with two cohorts
    # needs; a faculty with fewer would be forcing a cohort into a room that
    # cannot hold it, which is a problem the audit would then report against
    # the fixture rather than against a real timetable.
    ("LT1", 300), ("LT2", 350), ("LT3", 300), ("AUD1", 300), ("AUD2", 320),
    # the healthy case: two groups (60 students) in a room that seats them
    ("ROOM-X", 60), ("ROOM-Y", 70), ("ROOM-Z", 60), ("TUT1", 60),
    ("TUT2", 65), ("LAB1", 60), ("LAB2", 60),
    # too small for the class put in them: the deliberate overflow
    ("SEM1", 30), ("SEM2", 30), ("SEM3", 30), ("TUT3", 25), ("TUT4", 25),
    ("LAB3", 24), ("LAB4", 24), ("LAB5", 36), ("STD1", 35), ("STD2", 35),
    # workshops
    ("WS-CARP", 20), ("WS-WELD", 20), ("WS-MAS", 20), ("WS-ELEC", 20),
    # a room nobody has measured
    ("UNSIZED", 0),
]
TWO_GROUP_ROOMS = [
    "ROOM-X", "ROOM-Y", "ROOM-Z", "TUT1", "TUT2", "LAB1", "LAB2",
]
TOO_SMALL_ROOMS = [
    "SEM1", "SEM2", "SEM3", "TUT3", "TUT4", "LAB3", "LAB4", "LAB5",
    "STD1", "STD2",
]
HALLS = ["LT1", "LT2", "LT3", "AUD1", "AUD2"]
#: The rooms a whole-cohort lecture may use, and the cap on how many lectures
#: can run at one (day, slot). The two must agree: two lectures in one hall at
#: the same time is one class of twice the size, which reads as an over-capacity
#: class nobody planted.
LECTURE_HALLS = list(HALLS)
UNSIZED_ROOM = "UNSIZED"

#: Every room a class can be booked into, and so the number of parallel classes
#: the estate can hold in one time slot.
ALL_ROOMS = TWO_GROUP_ROOMS + TOO_SMALL_ROOMS + HALLS + [UNSIZED_ROOM]

COURSE_SPECS = [
    # (code, name, {activity: count})
    #
    # 24 configured courses, each offered to exactly two programmes, so every
    # programme studies six of them: a realistic year, and small enough that a
    # group has ~24 contact hours rather than a fictional 60. The index of a
    # course decides its two programmes (i, i+3), which is what makes the
    # cohorts below repeatable.
    ("CL111", "Communication Skills for Engineers", {"SEMINAR": 1}),          # 0
    ("MT161", "Mathematics I", {"TUTORIAL": 1, "PRACTICAL": 1}),             # 1
    ("MT171", "Mathematics II", {"TUTORIAL": 1}),                            # 2
    ("ME101", "Engineering Drawing", {"TUTORIAL": 1, "PRACTICAL": 1}),       # 3
    ("ME102", "Engineering Mechanics", {"TUTORIAL": 1, "PRACTICAL": 1}),      # 4
    ("EE201", "Circuit Analysis", {"TUTORIAL": 1, "PRACTICAL": 1}),          # 5
    ("EE202", "Electromagnetics", {"TUTORIAL": 1}),                          # 6
    ("EE203", "Digital Systems", {"TUTORIAL": 1, "PRACTICAL": 1}),           # 7
    ("CE201", "Structural Mechanics", {"TUTORIAL": 1, "PRACTICAL": 1}),      # 8
    ("CE202", "Hydraulics", {"TUTORIAL": 1}),                                # 9
    ("CH201", "Material Balances", {"TUTORIAL": 1, "PRACTICAL": 1}),         # 10
    ("CH202", "Transport Phenomena", {"TUTORIAL": 1}),                       # 11
    ("CH203", "Unit Operations", {"PRACTICAL": 1}),                          # 12
    ("CS201", "Programming", {"TUTORIAL": 1, "PRACTICAL": 1}),               # 13
    ("CS202", "Data Structures", {"TUTORIAL": 1, "PRACTICAL": 1}),            # 14
    ("CS203", "Databases", {"TUTORIAL": 1}),                                 # 15
    ("AR201", "History of Architecture", {"SEMINAR": 1}),                    # 16
    ("AR202", "Building Technology", {"TUTORIAL": 1, "PRACTICAL": 1}),        # 17
    ("ST201", "Probability", {"TUTORIAL": 1, "PRACTICAL": 1}),               # 18
    ("ST202", "Statistical Computing", {"TUTORIAL": 1}),                      # 19
    ("NM201", "Nautical Science I", {"TUTORIAL": 1, "PRACTICAL": 1}),        # 20
    ("NM202", "Watchkeeping", {"TUTORIAL": 1}),                               # 21
    ("SC121", "Technical Writing", {"SEMINAR": 1, "PRACTICAL": 1}),          # 22
    # MT455's seminar is required by CE and ST and timetabled nowhere: their
    # requirement can never be satisfied, which is the point.
    ("MT455", "Numerical Methods", {"SEMINAR": 1}),                          # 23
    # EE399 has NO requirement rows on purpose: the audit must report it as a
    # course whose activity requirements were never configured.
    ("EE399", "Electronics Project", {}),
]
UNCONFIGURED_COURSE = "EE399"

#: A course that is required by two programmes and timetabled to neither, so no
#: allocator can place it. It therefore consumes no lecture slot or class band:
#: a course with no session must not crowd out one that has.
UNRESOLVED_COURSE = "MT455"

DAYS_ORDER = (Day.MONDAY, Day.TUESDAY, Day.WEDNESDAY, Day.THURSDAY, Day.FRIDAY)

#: Where the practical classes go. They pile onto Wednesday because that is
#: what makes it the faculty's heaviest day, and spill onto Tuesday only where a
#: programme has more practicals than Wednesday has bands.
PRACTICAL_DAYS = (Day.WEDNESDAY, Day.TUESDAY)

#: Where seminars and tutorials go, spreading across the rest of the week.
SEMINAR_TUTORIAL_DAYS = (Day.MONDAY, Day.TUESDAY, Day.THURSDAY, Day.FRIDAY)

#: Four time bands a day, interleaved with the lecture slots so the day reads
#: like a real one instead of a morning of tutorials followed by an afternoon of
#: lectures. A class series occupies one band on one day and runs parallel
#: sections in different rooms, so the bands are the faculty's parallel capacity:
#: four of them is what lets a programme put all its practicals on Wednesday
#: without two of its courses colliding.
DAY_BANDS = {
    day: [
        (8 * 60, 9 * 60),
        (10 * 60 + 30, 11 * 60 + 30),
        (13 * 60, 14 * 60),
        (15 * 60 + 30, 16 * 60 + 30),
    ]
    for day in DAYS_ORDER
}
#: Wednesday is the practical day: six bands and no lectures. The two extra
#: bands are 09:00-10:30 and 14:00-15:30 -- hours that are lecture slots on
#: every other day. A faculty that runs its labs on one day does exactly this,
#: and it is what makes Wednesday the busiest day on its own account: with the
#: same four bands everywhere the load is spread to within a few hours across
#: the week, and "the busiest day" would be decided by a tie-break rather than
#: by the shape of the teaching. Lectures keep out of the bands automatically
#: (see _slot_hits_band), so Wednesday simply has no lecture slots and the
#: courses take theirs on the other four days.
DAY_BANDS[Day.WEDNESDAY] = [
    (8 * 60, 9 * 60),
    (9 * 60, 10 * 60 + 30),
    (10 * 60 + 30, 11 * 60 + 30),
    (13 * 60, 14 * 60),
    (14 * 60, 15 * 60 + 30),
    (15 * 60 + 30, 16 * 60 + 30),
]

#: Whole-cohort lectures, 90 minutes each, in the gaps between the bands. A
#: lecture is the one class *every* group of a course attends, so unlike a
#: tutorial there is no distributing them to avoid a clash — the courses have to
#: be spread over the lecture slots and the days, and each one needs a hall that
#: seats a whole cohort.
LECTURE_SLOTS = [
    (9 * 60, 10 * 60 + 30),
    (11 * 60 + 30, 13 * 60),
    (14 * 60, 15 * 60 + 30),
]
LECTURES_PER_COURSE = 1
LECTURE_PRESENT = 0.85

#: The one window left in the day, after the last band. Every planted problem is
#: placed here, so a plant can never be mistaken for real timetabling: it lands
#: where the faculty genuinely has nothing left.
PLANT_WINDOW = (16 * 60 + 30, 18 * 60)

#: Reserved for the planted problems, all inside the 08:00-18:00 working day —
#: a session ending after it is a *data* error the audit would rightly report,
#: and a plant must never be two things at once. Everything but "gap" also sits
#: in the free window above; "gap" is exempt because the group it belongs to has
#: its whole day rebuilt around the two sessions.
#:
#: The packed Wednesday is three contiguous half-hour labs, not one 90-minute
#: block: a group cannot be in three rooms at once, so sharing a single window
#: would plant a timetable clash on top of the day-load plant. Contiguous rather
#: than gapped, because the point is a wall of Wednesday afternoon.
PLANT_SLOTS = {
    "clash": [
        (Day.MONDAY, 16 * 60 + 45, 17 * 60 + 45),
        (Day.MONDAY, 17 * 60, 18 * 60),
    ],
    "overflow": (Day.TUESDAY, 16 * 60 + 45, 17 * 60 + 45),
    "orphan": (Day.FRIDAY, 16 * 60 + 30, 17 * 60 + 30),
    "gap": [
        (Day.THURSDAY, 8 * 60, 9 * 60),
        (Day.THURSDAY, 13 * 60, 14 * 60),
    ],
    "packed_wednesday": [
        (Day.WEDNESDAY, 16 * 60 + 30, 17 * 60),
        (Day.WEDNESDAY, 17 * 60, 17 * 60 + 30),
        (Day.WEDNESDAY, 17 * 60 + 30, 18 * 60),
    ],
}
PLANTS_MAY_OVERLAP = {"gap"}
PLANTS_MAY_SELF_OVERLAP = {"clash"}

#: The room each planted session sits in, one entry per window of the plant
#: above (``None`` for the session that deliberately has no room at all).
#: Declared here, next to the times, because the rooms have to be *reserved*
#: before the real classes are placed: the gap plant owns TUT1 on Thursday
#: 08:00, and a series allocated into that same room and period would be a
#: venue clash the fixture never meant to plant. Reading the room from this map
#: when the session is created is what keeps the reservation and the session
#: from drifting apart.
PLANT_ROOMS = {
    "clash": ["TUT2", "TUT3"],
    "overflow": ["SEM3"],
    "orphan": [None],
    "gap": ["TUT1", "ROOM-X"],
    "packed_wednesday": ["LT1", "LT1", "LT1"],
}


def _plant_slots(name: str) -> list:
    """The planted time windows for one plant, always as a list."""
    slot = PLANT_SLOTS[name]
    return list(slot) if isinstance(slot[0], tuple) else [slot]


def _plant_rooms(name: str) -> list:
    """The room of each planted window, always as a list parallel to the times."""
    return list(PLANT_ROOMS[name])


def _plant_reservations() -> dict:
    """Every planted room/period, keyed the way ``_pick_room`` expects.

    Reserved up front so a real class is never placed in a room a planted
    session already holds. A planted clash is between two *groups* in different
    rooms, so this never hides one: it only stops a third class joining a room
    that is already busy.
    """
    booked: dict = {}
    for name in PLANT_SLOTS:
        for (day, start, _end), room in zip(_plant_slots(name), _plant_rooms(name)):
            if room:
                booked[(room, day, start)] = True
    return booked


def _assert_plants_own_their_slots() -> None:
    """Every planted time must be free of real classes and lectures.

    Without this the plants drift the moment someone edits a band, and the
    fixture quietly grows a second, unintended problem — which is how a golden
    test stops meaning anything.
    """
    day_start, day_end = 8 * 60, 18 * 60
    for name in PLANT_SLOTS:
        windows = _plant_slots(name)
        rooms = _plant_rooms(name)
        if len(rooms) != len(windows):
            raise AssertionError(
                f"planted {name} has {len(windows)} time(s) but {len(rooms)} "
                "room(s): every planted window needs the room it is booked in"
            )
        for day, start, end in windows:
            if start < day_start or end > day_end:
                raise AssertionError(
                    f"planted {name} at {day} {start}-{end} is outside the "
                    f"{day_start}-{day_end} working day"
                )
            if name not in PLANTS_MAY_OVERLAP:
                occupied = list(DAY_BANDS[day]) + list(LECTURE_SLOTS)
                for other_start, other_end in occupied:
                    if start < other_end and other_start < end:
                        raise AssertionError(
                            f"planted {name} at {day} {start}-{end} overlaps a "
                            f"real class or lecture at {other_start}-{other_end}"
                        )
        # Two windows of the same plant must not overlap either — a group cannot
        # be in two rooms at once, so a self-overlapping plant would report a
        # clash the fixture never meant to plant. "clash" is the one plant whose
        # whole purpose is to make its own windows overlap.
        if name in PLANTS_MAY_SELF_OVERLAP:
            continue
        ordered = sorted(windows, key=lambda w: w[1])
        for first, second in zip(ordered, ordered[1:]):
            if first[2] > second[1]:
                raise AssertionError(
                    f"planted {name} windows {first[1]}-{first[2]} and "
                    f"{second[1]}-{second[2]} overlap"
                )


def _t(minutes: int) -> time:
    return time(hour=minutes // 60, minute=minutes % 60)


def _bands(day) -> int:
    """How many parallel class bands a day has (Wednesday has a fifth)."""
    return len(DAY_BANDS[day])


def _slot_hits_band(day, start, end) -> bool:
    """Whether a lecture slot would land on a period the faculty is timetabling.

    The two passes keep one calendar between them: a lecture in a period that
    is also a class band would put a whole cohort in a room alongside whatever
    else the faculty is running, and the audit would report a venue clash that
    the fixture invented.
    """
    return any(start < band_end and band_start < end for band_start, band_end in DAY_BANDS[day])


def _groups_per_class() -> int:
    """How many groups share one class.

    Two: a 60-student class in a 60-70 seat room is the ordinary, healthy case
    for a faculty this size, and it is what makes the deliberately bad room
    choices in the fixture read as problems.
    """
    return GROUPS_PER_CLASS


def _pick_room(per_session, cursor, taken=None, day=None, start=None, pool=None):
    """Room *name* for one class. Deliberately mixes good and poor choices.

    Two groups in a room that seats 60 is the healthy majority. Every twelfth
    class is put in a room that cannot hold it, every sixth in a hall several
    times its size, and every twenty-third in the room nobody has measured, so
    the venue section has real overflow, right-sizing and unknown-capacity
    evidence instead of a uniform table of full rooms.

    That deliberate choice is only used when the room is actually free; the
    rest of the estate is then swept in preference order, so a class is never
    booked into a room another class already has. ``taken`` is keyed by room
    *name* and must be probed with a name too -- a Venue instance and its name
    are different keys, and a mismatch makes every room look free.

    Running out of rooms is an error, not a fallback. Returning the last
    candidate anyway books two classes into one room, which the audit then
    reports as a venue clash: the fixture would be inventing the problem it is
    supposed to be studying. ``pool`` restricts the estate (lectures use halls).
    """
    options = list(pool) if pool is not None else ALL_ROOMS
    wanted = _room_name(per_session, cursor, pool)
    for name in [wanted] + [r for r in _room_order(cursor, options) if r != wanted]:
        if not taken or (name, day, start) not in taken:
            return name
    raise AssertionError(
        f"all {len(options)} rooms are already booked on day {day} at "
        f"{_t(start).strftime('%H:%M')}, so a class would have to be double-booked"
    )


def _room_order(cursor, options):
    """``options`` rotated by ``cursor``, so the preference sweeps round."""
    size = len(options)
    offset = cursor % size
    return options[offset:] + options[:offset]


def _rotated(days, rotation):
    """``days`` starting at ``rotation``, so two programmes start differently."""
    offset = rotation % len(days)
    return days[offset:] + days[:offset]


def _sections(cohort):
    """Parallel classes one cohort of ``cohort`` groups is taught in."""
    per_session = _groups_per_class()
    return max(1, math.ceil(cohort / per_session)) if cohort else 0


def _room_name(per_session, cursor, pool=None):
    if pool is not None:
        return pool[cursor % len(pool)]
    if cursor % 23 == 11:
        return UNSIZED_ROOM
    if cursor % 12 == 5:
        # Rare on purpose. A class in a room that cannot hold it is a genuine
        # finding, but a rotation that put one in five classes would bury every
        # other kind of problem under a single cause and leave the report with
        # nothing to say except "too many small rooms".
        return TOO_SMALL_ROOMS[(cursor // 12) % len(TOO_SMALL_ROOMS)]
    if cursor % 6 == 4:
        return HALLS[(cursor // 6) % len(HALLS)]
    return TWO_GROUP_ROOMS[(cursor // 2) % len(TWO_GROUP_ROOMS)]


def build(academic_year="2024/25", semester_number=1) -> dict:
    """Create the whole synthetic world. Returns a stats dict for the console."""
    rng = random.Random(SEED)
    _assert_plants_own_their_slots()

    semester, _ = Semester.objects.get_or_create(
        academic_year=academic_year, semester=semester_number
    )
    Semester.set_current(semester)

    # ── venues ──────────────────────────────────────────────────────────────
    venues = {}
    for name, cap in VENUE_SPECS:
        venue, _ = Venue.objects.get_or_create(name=name, defaults={"capacity": cap})
        venues[name] = venue

    # ── programmes and groups ───────────────────────────────────────────────
    groups: list = []
    for pcode, pname in PROGRAMMES:
        programme, _ = Programme.objects.get_or_create(
            code=pcode, defaults={"name": pname}
        )
        for n in range(1, GROUPS_PER_PROGRAMME + 1):
            group, _ = StudentGroup.objects.get_or_create(
                programme=programme, code=f"C{n}"
            )
            groups.append(group)
    groups = groups[:N_GROUPS]
    group_rows = {f"{g.programme.code} {g.code}": g for g in groups}
    group_list = [f"{g.programme.code} {g.code}" for g in groups]
    groups_by_programme: dict = {}
    for key in group_list:
        groups_by_programme.setdefault(key.split(" ")[0], []).append(key)

    # ── courses and their requirements ──────────────────────────────────────
    courses = {}
    for code, name, reqs in COURSE_SPECS:
        course, _ = Course.objects.get_or_create(code=code, defaults={"name": name})
        if name and course.name != name:
            course.name = name
            course.save()
        if not course.activity_requirements.exists():
            for activity, count in reqs.items():
                if count:
                    CourseActivityRequirement.objects.create(
                        course=course, activity_type=activity, count=count
                    )
        courses[code] = course

    # Each course is offered to exactly two programmes, so 20 of the 80 groups
    # study it. The pairing is a stable rotation, not random, so the same
    # cohorts recur and the golden test can name them. Two courses per programme
    # over 23 courses keeps every programme to six or fewer, which is the most
    # that fits in the lecture slots with disjoint days — so no course is
    # hand-added to a programme below and the capacity is checked rather than
    # assumed.
    configured_codes = [c for c, _, r in COURSE_SPECS if r]
    programme_codes = [p for p, _ in PROGRAMMES]
    course_programmes: dict = {}
    for index, code in enumerate(configured_codes):
        if code == UNRESOLVED_COURSE:
            continue    # no sessions at all: the unresolved requirement
        course_programmes[code] = [
            programme_codes[index % len(programme_codes)],
            programme_codes[(index + 3) % len(programme_codes)],
        ]
    # MT455 is required by CE and ST, so both cohorts are left unresolved. It is
    # never offered a session, so it also costs those programmes no lecture slot.
    course_programmes[UNRESOLVED_COURSE] = ["CE", "ST"]

    programme_courses: dict = {}
    for code, plist in course_programmes.items():
        for pcode in dict.fromkeys(plist):
            programme_courses.setdefault(pcode, set()).add(code)
    programme_objs = {code: Programme.objects.get(code=code) for code in programme_codes}
    for pcode, codes in programme_courses.items():
        for code in sorted(codes):
            ProgrammeCourse.objects.get_or_create(
                programme=programme_objs[pcode],
                course=courses[code],
                defaults={"semester": semester.semester},
            )

    # ── small-group session series ──────────────────────────────────────────
    # series[(code, programme, activity)] = [Session, ...] — the parallel
    # classes one programme's groups are distributed across. A series never puts
    # two of its own classes in the same room at the same time, so a group that
    # attends one class of a course never clashes with a classmate's.
    pending: list = []

    def make_session(code, activity, day, start, end, venue):
        return Session(
            semester=semester,
            course_code=code,
            activity_type=activity,
            day=day,
            start_time=_t(start),
            end_time=_t(end),
            venue=venue,
        )

    # A group's week is chosen by *which day it takes out of each series*, so
    # the series have to be laid out so that no combination of choices can put a
    # group in two rooms at once. Two series collide for a group only if they
    # share a time band AND the group takes the same weekday out of both, so
    # A class series runs on ONE day a week, in ONE band, as parallel sections
    # in different rooms — which is how a faculty actually works, and is what
    # makes the faculty's day load lopsided instead of uniform. Practical classes
    # are the ones that pile onto one afternoon, so every PRACTICAL series is
    # placed on Wednesday; seminars and tutorials rotate over the other four.
    #
    # Because every group of a cohort attends its course on that one day, two
    # series of the same programme must never share a (day, band) cell — that
    # would double-book every group in the programme at once. Distinct days are
    # therefore handed out first, and only where a programme genuinely has more
    # series than days does a second band get used, which the assertion below
    # bounds to the bands available.
    series_slots: dict = {}
    demand: dict = defaultdict(int)
    # Balanced across the whole faculty, not per programme. Filling each
    # programme's bands independently puts every programme's first class of the
    # day in the same period, so one morning ends up holding 40 classes in a
    # faculty of 23 rooms -- and the estate cannot hold that, however the
    # classes are named.
    day_total: dict = defaultdict(int)
    cell_series: dict = defaultdict(int)
    per_cell = min(
        len(ALL_ROOMS) // _sections(len(cohort))
        for cohort in groups_by_programme.values()
        if cohort
    )
    for pcode, codes in sorted(programme_courses.items()):
        offered = [
            (code, activity)
            for code in sorted(codes)
            for activity in SMALL_ACTIVITIES
            if pcode in course_programmes.get(code, ())
            and courses[code].required_count(activity)
            # The course nobody timetabled. It must stay out of the *series* as
            # well as the lectures, not just the group links: a seminar session
            # here would quietly satisfy the requirement and the audit would
            # never report the missing class it exists to expose.
            and code != UNRESOLVED_COURSE
        ]
        cohort = len(groups_by_programme.get(pcode, []))
        sections = _sections(cohort)
        # Each programme starts on its own day. With one shared preference
        # order every programme opens on the same morning and the rest of the
        # week has to absorb the difference.
        rotation = programme_codes.index(pcode) if pcode in programme_codes else 0
        for code, activity in offered:
            practical = activity == ActivityType.PRACTICAL
            if practical:
                # Wednesday is the practical day and is filled before Tuesday
                # gets a look in. Choosing the quietest day here would make
                # Wednesday *quieter*: the extra band it has would simply never
                # be used, and the busiest day of the week would be decided by
                # whichever lecture day happened to fall behind.
                days = list(PRACTICAL_DAYS)
            else:
                days = _rotated(SEMINAR_TUTORIAL_DAYS, rotation)
            with_room = [d for d in days if day_total[d] < _bands(d) * per_cell]
            if not with_room:
                raise AssertionError(
                    f"programme {pcode} has run out of room on every day it may "
                    f"use ({[d.label for d in days]}): {code} would have to "
                    "share a period with another class, putting a group in two "
                    "classes at once"
                )
            day = with_room[0] if practical else min(
                with_room, key=lambda d: (day_total[d], days.index(d))
            )
            band = min(range(_bands(day)), key=lambda b: cell_series[(day, b)])
            cell_series[(day, band)] += 1
            day_total[day] += 1
            series_slots[(code, pcode, activity)] = (day, band)
            demand[(day, DAY_BANDS[day][band][0])] += sections

    # The estate has to hold every class that starts at the same moment, and
    # this is the check that says so *before* anything is written. The busiest
    # cells are listed because "the week does not fit" is only useful if it says
    # which periods do not fit.
    busiest = sorted(demand.items(), key=lambda kv: -kv[1])[:3]
    for (day, start), classes in busiest:
        if classes > len(ALL_ROOMS):
            raise AssertionError(
                f"{classes} classes start on {day} at "
                f"{_t(start).strftime('%H:%M')} and the estate only has "
                f"{len(ALL_ROOMS)} rooms, so some would have to share. "
                "Busiest periods: "
                + ", ".join(
                    f"{d.label} {_t(s).strftime('%H:%M')}={n}" for (d, s), n in busiest
                )
            )

    series: dict = {}
    taken: dict = _plant_reservations()
    room_cursor = 0
    for (code, pcode, activity), (day, band) in sorted(
        series_slots.items(), key=lambda kv: (kv[0][1], kv[0][0], kv[0][2])
    ):
        cohort = groups_by_programme.get(pcode, [])
        if not cohort:
            continue
        per_session = _groups_per_class()
        count = max(1, math.ceil(len(cohort) / per_session))
        start, end = DAY_BANDS[day][band]
        # Parallel sections of one series: same day, same band, different rooms.
        # ``taken`` is keyed by room *name* and is what stops two sections being
        # booked into one room, which would merge them into a single
        # over-capacity class. It has to be the name on both sides of the check:
        # a Venue instance and its name are different keys, and a mismatch
        # silently makes every room look free.
        for _ in range(count):
            name = _pick_room(per_session, room_cursor, taken, day, start)
            room_cursor += 1
            taken[(name, day, start)] = True
            series.setdefault((code, pcode, activity), []).append(
                make_session(code, activity, day, start, end, venues[name])
            )
    pending = [sess for rows in series.values() for sess in rows]

    # ── whole-cohort lectures ───────────────────────────────────────────────
    # One per course. A lecture is attended by *every* group of its course, so
    # there is no distributing them to dodge a clash: the courses themselves
    # have to be spread over the lecture slots and the days. A course takes the
    # lecture slots round-robin, and each one goes on whichever day its slot is
    # currently lightest, so the lectures fill the week evenly instead of piling
    # onto Monday. Every lecture needs a hall that seats a whole cohort, and
    # ``lecture_taken`` stops two sharing a hall at the same time — that would
    # read as a single class of twice the size, an over-capacity class nobody
    # planted. The cap is asserted, because crossing it means a slot is
    # oversubscribed and the estate cannot hold the week at all.
    #
    # The session always exists; who *turns up* to it is settled with the
    # group links below, at the same LECTURE_PRESENT rate. Deciding attendance
    # twice (a session that might not exist, and a group that might not attend)
    # would square the absence rate and quietly halve the faculty's lectures.
    lecture_sessions: list = []
    pending_lectures: dict = {}
    lecture_taken: dict = {}
    cell_load: dict = defaultdict(int)
    lecture_cursor = 0
    for pcode, codes in sorted(programme_courses.items()):
        offered = [
            code
            for code in sorted(codes)
            if pcode in course_programmes.get(code, ())
            and code != UNRESOLVED_COURSE
        ]
        cohort = groups_by_programme.get(pcode, [])
        for index, code in enumerate(offered):
            slot = index % len(LECTURE_SLOTS)
            start, end = LECTURE_SLOTS[slot]
            # A lecture may only use a day whose bands leave that period free.
            free_days = [d for d in DAYS_ORDER if not _slot_hits_band(d, start, end)]
            day = min(free_days, key=lambda d: (cell_load[(d, slot)], d))
            if cell_load[(day, slot)] + 1 > len(LECTURE_HALLS):
                raise AssertionError(
                    f"{cell_load[(day, slot)] + 1} lectures land on {day} at "
                    f"{_t(LECTURE_SLOTS[slot][0]).strftime('%H:%M')} and there "
                    f"are only {len(LECTURE_HALLS)} halls, so two of them "
                    "would share a hall and read as one over-capacity class"
                )
            name = _pick_room(
                len(cohort),
                lecture_cursor,
                lecture_taken,
                day,
                start,
                LECTURE_HALLS,
            )
            lecture_cursor += 1
            lecture_taken[(name, day, start)] = True
            sess = make_session(
                code, ActivityType.LECTURE, day, start, end, venues[name]
            )
            cell_load[(day, slot)] += 1
            lecture_sessions.append(sess)
            # Keyed by programme, not by course. A course offered to two
            # programmes is coloured — and therefore scheduled — once per
            # programme, so keying by course alone would hand every group both
            # cohorts' lectures and double-book the lot of them.
            pending_lectures.setdefault((pcode, code), []).append(sess)
    pending.extend(lecture_sessions)

    # ── the planted problems ────────────────────────────────────────────────
    plants: dict = {}

    # (1) Packed Wednesday: every EE group gets three extra practicals back to
    #     back on a Wednesday afternoon, so Wednesday carries visibly the most
    #     hours and no evening free of teaching.
    plants["packed_wednesday"] = "EE C1-C10"
    packed = _plant_slots("packed_wednesday")
    # LT1, the 300-seat hall: all ten EE groups attend each of these, so 300
    # students in a 70-seat lab would be a *second* planted problem (a capacity
    # overflow) hiding inside the day-load plant. The hall holds them exactly.
    ee_extra = [
        make_session(
            code,
            ActivityType.PRACTICAL,
            day,
            start,
            end,
            venues[room],
        )
        for (day, start, end), code, room in zip(
            packed, ("EE201", "CE201", "CS201"), _plant_rooms("packed_wednesday")
        )
    ]
    pending.extend(ee_extra)

    # (2) A 4-hour gap for ME C3: 08:00-09:00 then nothing until 13:00-14:00.
    plants["long_gap_group"] = "ME C3"
    gap_tutorial = make_session(
        "MT171",
        ActivityType.TUTORIAL,
        *_plant_slots("gap")[0],
        venues[_plant_rooms("gap")[0]],
    )
    gap_practical = make_session(
        "ME102",
        ActivityType.PRACTICAL,
        *_plant_slots("gap")[1],
        venues[_plant_rooms("gap")[1]],
    )
    pending.extend([gap_tutorial, gap_practical])

    # (3) Over-capacity: CH C1 and CH C2 both placed in a 30-seat seminar room.
    #     Both study CL111, so this is a real overflow and not an eligibility
    #     problem in its own right.
    plants["overflow"] = "CH C1 + CH C2 in SEM3"
    overflow = make_session(
        "CL111",
        ActivityType.SEMINAR,
        *PLANT_SLOTS["overflow"],
        venues[_plant_rooms("overflow")[0]],
    )
    pending.append(overflow)

    # (4) Unresolved: MT455 requires a seminar and has no seminar session.
    plants["unresolved"] = "MT455 seminar for CE and ST"

    # (5) A session with no room at all, so its capacity cannot be verified.
    #     ST studies SC121, so these two placements are not also an
    #     eligibility problem in their own right.
    plants["missing_venue"] = "SC121 practical, no room"
    orphan = make_session(
        "SC121", ActivityType.PRACTICAL, *PLANT_SLOTS["orphan"], None
    )
    pending.append(orphan)

    # (6) Two tutorials that overlap by 30 minutes, both attended by EE C1.
    #     EE201 and CE201 are both in the EE course set, so neither placement is
    #     for a course the group does not study.
    plants["overlap_group"] = "EE C1"
    clash_a = make_session(
        "EE201",
        ActivityType.TUTORIAL,
        *PLANT_SLOTS["clash"][0],
        venues[_plant_rooms("clash")[0]],
    )
    clash_b = make_session(
        "CE201",
        ActivityType.TUTORIAL,
        *PLANT_SLOTS["clash"][1],
        venues[_plant_rooms("clash")[1]],
    )
    pending.extend([clash_a, clash_b])

    sessions = Session.objects.bulk_create(pending, batch_size=500)
    by_pk = {s.pk: s for s in sessions}
    series = {k: [by_pk[s.pk] for s in v] for k, v in series.items()}
    lecture_sessions = [by_pk[s.pk] for s in lecture_sessions]
    ee_extra = [by_pk[s.pk] for s in ee_extra]
    gap_tutorial = by_pk[gap_tutorial.pk]
    gap_practical = by_pk[gap_practical.pk]
    overflow = by_pk[overflow.pk]
    orphan = by_pk[orphan.pk]
    clash_a = by_pk[clash_a.pk]
    clash_b = by_pk[clash_b.pk]
    lectures_by_programme = {
        key: [by_pk[row.pk] for row in rows]
        for key, rows in pending_lectures.items()
    }

    # ── assignment of groups to sessions ────────────────────────────────────
    # The link plan is built in memory and written in one insert: a row at a
    # time needs tens of thousands of queries on a cohort this size, and the
    # fixture has to stay fast enough to run inside a test.
    links: list = []
    assignments: list = []
    seen: set = set()
    counter = 0

    def plan_link(key, sess, record=True):
        nonlocal counter
        if (key, sess.pk) in seen:
            return
        seen.add((key, sess.pk))
        links.append(SessionGroup(session=sess, group=group_rows[key]))
        counter += 1
        if not record:
            # An extra commitment on top of a requirement the group has already
            # met. It is in the timetable, so the audit sees it, but it is not
            # an extra requirement.
            return
        # Every seventh link is presented as a pre-existing (retained) one, so
        # the origin split in the report is a real three-way split.
        status = "retained" if counter % 7 == 0 else "added"
        assignments.append(
            {
                "group": key,
                "programme": key.split(" ")[0],
                "course": sess.course_code,
                "course_name": courses[sess.course_code].name,
                "activity": ActivityType(sess.activity_type).label,
                "ordinal": 1,
                "status": status,
                "session_pk": sess.pk,
                "session": str(sess),
                "day": sess.get_day_display(),
                "time": f"{sess.start_time:%H:%M}-{sess.end_time:%H:%M}",
                "venue": sess.venue.name if sess.venue_id else "",
                "moved_from_pk": None,
                "moved_from": "",
                "group_count": 0,
                "capacity_status": "unknown",
                "capacity_message": "",
            }
        )

    for key in group_list:
        pcode = key.split(" ")[0]
        for code in sorted(programme_courses.get(pcode, set())):
            for activity in SMALL_ACTIVITIES:
                rows = series.get((code, pcode, activity)) or []
                if not rows:
                    continue
                # A series is one day a week, so a group has no day to choose —
                # only which of the parallel sections to sit in, which is what
                # ``plan_link`` below does. Two series of the same programme
                # never share a (day, band) cell, so this cannot double-book.
                section = (int(key.split(" ")[1][1:]) - 1) // _groups_per_class()
                plan_link(key, rows[min(section, len(rows) - 1)])
            for sess in lectures_by_programme.get((pcode, code), []):
                # A group's two lectures are the only per-group freedom left in
                # a week that is otherwise the same for all ten of its cohort,
                # and it is what gives the TFI some spread to rank.
                if rng.random() < LECTURE_PRESENT:
                    plan_link(key, sess, record=False)

    # The planted extras, on top of the normal assignment.
    for key in groups_by_programme["EE"]:
        for sess in ee_extra:
            plan_link(key, sess, record=False)
    for key in ("CH C1", "CH C2"):
        plan_link(key, overflow, record=False)
    for key in ("ST C3", "ST C4"):
        plan_link(key, orphan, record=False)
    plan_link("EE C1", clash_a, record=False)
    plan_link("EE C1", clash_b, record=False)
    # Only ME C3 gets the pair. Giving it to the rest of the cohort as well
    # would not make the gap any more real, and would put two more sessions in
    # everybody else's week.
    plan_link("ME C3", gap_tutorial, record=False)
    plan_link("ME C3", gap_practical, record=False)

    # Three groups end the week with nothing at all: an edge case the audit has
    # to list rather than score. Matched on pk, not on the code — every
    # programme has a C1.
    empty_groups = groups_by_programme["NM"][-3:]
    empty_pks = {group_rows[key].pk for key in empty_groups}
    plants["empty_groups"] = empty_groups
    links = [ln for ln in links if ln.group_id not in empty_pks]
    assignments = [a for a in assignments if a["group"] not in empty_groups]

    # The planted 4-hour gap is only real if nothing else sits in ME C3's own
    # Thursday timetable.
    keep = {gap_tutorial.pk, gap_practical.pk}
    me_c3_pk = group_rows["ME C3"].pk
    links = [
        ln
        for ln in links
        if ln.group_id != me_c3_pk
        or ln.session.day != Day.THURSDAY
        or ln.session_id in keep
    ]
    assignments = [
        a
        for a in assignments
        if not (a["group"] == "ME C3" and a["session_pk"] not in keep)
    ]

    SessionGroup.objects.bulk_create(links, ignore_conflicts=True, batch_size=500)

    # ── workshop + technical drawing allocations ────────────────────────────
    workshops = []
    for n, key in enumerate(group_list[:24]):
        craft, room = [
            ("Carpentry", "WS-CARP"),
            ("Welding", "WS-WELD"),
            ("Masonry", "WS-MAS"),
            ("Electrical", "WS-ELEC"),
        ][n % 4]
        workshops.append(
            WorkshopAllocation(
                semester=semester,
                course_code="WORKSHOP",
                group_code=key.split(" ")[1],
                day=Day.WEDNESDAY,
                start_time=_t(8 * 60),
                end_time=_t(13 * 60),
                venue=venues[room].name,
                workshop=craft,
                week_start=1 + (n % 3) * 7,
                week_end=7 + (n % 3) * 7,
            )
        )
    # A period-only row: a day is named, times are not — "unverifiable".
    workshops.append(
        WorkshopAllocation(
            semester=semester,
            course_code="WORKSHOP",
            group_code="C1",
            day=Day.FRIDAY,
            start_time=None,
            end_time=None,
            venue=venues["WS-CARP"].name,
            workshop="Carpentry",
            time_period="MORNING",
            week_start=1,
            week_end=7,
        )
    )
    WorkshopAllocation.objects.bulk_create(workshops, batch_size=200)

    TechnicalDrawingAllocation.objects.bulk_create(
        [
            TechnicalDrawingAllocation(
                semester=semester,
                course_code="ME101",
                group_code=key.split(" ")[1],
                day=Day.TUESDAY,
                start_time=_t(10 * 60),
                end_time=_t(13 * 60),
                venue=venues["STD1"].name,
            )
            for key in group_list[24:48]
        ],
        batch_size=200,
    )

    # ── the applied allocation run ──────────────────────────────────────────
    # Every group studying MT455 is unresolved: the course requires a seminar
    # and the semester has none, so no allocator can place it.
    unresolved_items = []
    for pcode in ("CE", "ST"):
        for key in groups_by_programme[pcode]:
            unresolved_items.append(
                {
                    "group": key,
                    "programme": pcode,
                    "course": "MT455",
                    "course_name": courses["MT455"].name,
                    "activity": "Seminar",
                    "ordinal": 1,
                    "sessions_considered": 0,
                    "reasons": ["no SEMINAR session of MT455 exists in the timetable"],
                }
            )
    added = sum(1 for a in assignments if a["status"] == "added")
    retained = sum(1 for a in assignments if a["status"] == "retained")
    distinct = {
        (a["group"], a["course"], a["activity"]) for a in assignments
    }
    snapshot = {
        "semester": str(semester),
        "semester_id": semester.pk,
        "scope": "ALL",
        "added": added,
        "moved": 0,
        "retained": retained,
        "removals": 0,
        "unresolved": len(unresolved_items),
        "requirement_total": len(distinct) + len(unresolved_items),
        "search_limit_hit": False,
        "scanned": 18422,
        "duration_ms": 4210,
        "complete": False,
        "warnings": [],
        "unconfigured_courses": [
            {
                "code": UNCONFIGURED_COURSE,
                "name": courses[UNCONFIGURED_COURSE].name,
                "programmes": ["EE"],
                "group_count": GROUPS_PER_PROGRAMME,
                "variants": [],
            }
        ],
        "assignments": assignments,
        "unresolved_items": unresolved_items,
    }

    payload = json.dumps(snapshot)
    run, _ = AllocationRun.objects.update_or_create(
        semester=semester,
        scope=AllocationScope.ALL,
        defaults={
            "status": AllocationStatus.APPLIED,
            "assigned": added,
            "retained": retained,
            "moved": 0,
            "removed": 0,
            "unresolved": len(unresolved_items),
            "summary": payload,
            "algorithm": "smart",
            "search_limit_hit": False,
        },
    )

    return {
        "semester": str(semester),
        "programmes": Programme.objects.count(),
        "groups": StudentGroup.objects.count(),
        "courses": Course.objects.count(),
        "venues": Venue.objects.count(),
        "sessions": Session.objects.filter(semester=semester).count(),
        "session_group_links": SessionGroup.objects.filter(
            session__semester=semester
        ).count(),
        "workshop_allocations": WorkshopAllocation.objects.filter(
            semester=semester
        ).count(),
        "td_allocations": TechnicalDrawingAllocation.objects.filter(
            semester=semester
        ).count(),
        "run_id": run.pk,
        "requirements": snapshot["requirement_total"],
        "allocated": len(distinct),
        "unresolved": len(unresolved_items),
        "planted": plants,
    }


if __name__ == "__main__":  # pragma: no cover - convenience entry point
    import os

    import django

    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "coet.settings")
    django.setup()

    for key, value in build().items():
        print(f"{key:24} {value}")
