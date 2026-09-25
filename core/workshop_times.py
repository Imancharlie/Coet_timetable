"""Standard workshop session times — the single source of truth.

Workshops run only during these sessions (the university rules, matching the
source workshop allocation Excel files):

    Monday/Tuesday/Wednesday/Friday
        Morning    09:00-13:00
        Afternoon  15:00-19:00
    Thursday
        Morning    10:00-14:00
        Afternoon  15:00-19:00

These are the only session times a workshop (either a ``WorkshopAllocation``
record or a ``Session`` with ``activity_type == WORKSHOP``) may use. The rules
drive the create/edit forms (inline validation messages), the flat workshop
importer, the master-timetable importer, the on-screen timetable grid and the
PDF export, and they back the non-destructive legacy-record identification
shown on the lists and via ``manage.py identify_legacy_workshops``. No other
activity type is affected.
"""

import datetime

from django.core.exceptions import ValidationError

from core.models import Day, TimePeriod

# day (upper-case Day choice) -> TimePeriod -> (start, end) of the ONE session
# a workshop may occupy. WEEKDAY keys below repeat intentionally so the rules
# stay explicit.
_WORKSHOP_MORNING = (datetime.time(9, 0), datetime.time(13, 0))
_WORKSHOP_AFTERNOON = (datetime.time(15, 0), datetime.time(19, 0))
_THURSDAY_MORNING = (datetime.time(10, 0), datetime.time(14, 0))

WORKSHOP_STANDARD_TIMES = {
    Day.MONDAY: {
        TimePeriod.MORNING: _WORKSHOP_MORNING,
        TimePeriod.AFTERNOON: _WORKSHOP_AFTERNOON,
    },
    Day.TUESDAY: {
        TimePeriod.MORNING: _WORKSHOP_MORNING,
        TimePeriod.AFTERNOON: _WORKSHOP_AFTERNOON,
    },
    Day.WEDNESDAY: {
        TimePeriod.MORNING: _WORKSHOP_MORNING,
        TimePeriod.AFTERNOON: _WORKSHOP_AFTERNOON,
    },
    Day.THURSDAY: {
        TimePeriod.MORNING: _THURSDAY_MORNING,
        TimePeriod.AFTERNOON: _WORKSHOP_AFTERNOON,
    },
    Day.FRIDAY: {
        TimePeriod.MORNING: _WORKSHOP_MORNING,
        TimePeriod.AFTERNOON: _WORKSHOP_AFTERNOON,
    },
}

WORKSHOP_DAYS = tuple(WORKSHOP_STANDARD_TIMES)


def _programme_codes(programme_code):
    """Normalise ``programme_code`` (one code, an iterable, or None) to a
    frozenset of codes, or None when nothing was given."""
    if programme_code is None:
        return None
    if isinstance(programme_code, str):
        codes = (programme_code,) if programme_code else ()
    else:
        codes = tuple(programme_code or ())
    return frozenset(c for c in codes if c)


def programme_codes_for_group_codes(group_codes):
    """Programme codes owning the given StudentGroup code(s) (() when unknown).

    ``group_codes`` may be a single code string or an iterable of codes.
    """
    from core.models import StudentGroup

    if isinstance(group_codes, str):
        group_codes = [group_codes]
    codes = [g for g in (group_codes or ()) if g]
    if not codes:
        return ()
    return tuple(
        StudentGroup.objects.filter(code__in=codes)
        .values_list("programme__code", flat=True)
        .distinct()
    )


def allocation_programme_codes(record):
    """Programme codes a WorkshopAllocation resolves to (via its group)."""
    return programme_codes_for_group_codes(record.group_code)


def session_programme_codes(session):
    """Programme codes a Session resolves to (via its attached groups).

    Before the groups of a new session are saved (e.g. inside a create form's
    ``clean()``) this is () and the record is judged against the default
    10:00 Thursday rule.
    """
    try:
        group_codes = [link.group.code for link in session.session_groups.all()]
    except (ValueError, TypeError):
        group_codes = []
    return programme_codes_for_group_codes(group_codes)


def course_programme_codes(course_code):
    """Programme codes that study the course (ProgrammeCourse rows)."""
    from core.models import ProgrammeCourse

    if not course_code:
        return ()
    return tuple(
        ProgrammeCourse.objects.filter(course_code=course_code)
        .values_list("programme__code", flat=True)
        .distinct()
    )


def day_label(day):
    try:
        return Day(day).label
    except ValueError:
        return str(day)


def period_label(period):
    try:
        return TimePeriod(period).label
    except ValueError:
        return str(period)


def is_workshop_day(day) -> bool:
    """True when ``day`` is a day workshops are scheduled on (Mon-Fri)."""
    return bool(day) and day in WORKSHOP_DAYS


def workshop_periods_for_day(day):
    """the {period: (start, end)} sessions for ``day``, or None when the day
    has no workshop sessions."""
    if not day:
        return None
    return WORKSHOP_STANDARD_TIMES.get(day)


def workshop_times_for(day, period, programme_code=None):
    """(start, end) for ``day``/``period``, or None when that session does not
    exist (e.g. a weekend day or an unknown period).

    ``programme_code`` is accepted for API compatibility but no longer changes
    the times: Thursday morning is 10:00-14:00 for every programme.
    """
    periods = workshop_periods_for_day(day)
    if not periods:
        return None
    if period not in periods:
        return None
    return periods[period]


def workshop_period_for_times(day, start, end, programme_code=None):
    """The period whose exact standard session matches ``start``/``end`` on
    ``day``, or None when the times match no workshop session."""
    periods = workshop_periods_for_day(day)
    if not periods or start is None or end is None:
        return None
    for period in periods:
        std_start, std_end = workshop_times_for(day, period, programme_code)
        if std_start == start and std_end == end:
            return period
    return None


def matches_standard_workshop_time(day, start, end, programme_code=None) -> bool:
    return workshop_period_for_times(day, start, end, programme_code) is not None


def workshop_hours(day, period, programme_code=None):
    """Hourly grid slots (07:00-08:00 .. 19:00-20:00) a workshop occupies.

    Morning 09:00-13:00 => {9, 10, 11, 12}; Thursday morning 10:00-14:00 =>
    {10, 11, 12, 13}; afternoon 15:00-19:00 => {15, 16, 17, 18}.
    """
    times = workshop_times_for(day, period, programme_code)
    if not times:
        return set()
    start, end = times
    return {
        h
        for h in range(7, 20)
        if start < datetime.time(h + 1) and end > datetime.time(h)
    }


def workshop_time_message(day=None, period=None, programme_code=None) -> str:
    """A clear description of the required time for a day and period.

    Used verbatim as the inline validation message when a workshop's times do
    not match the rules for the selected day/session.
    """
    periods = workshop_periods_for_day(day)
    if day and periods and period in periods:
        start, end = workshop_times_for(day, period, programme_code)
        return (
            f"{day_label(day)} {period_label(period)} workshops must run "
            f"{start:%H:%M}-{end:%H:%M}."
        )
    if day and periods:
        parts = " or ".join(
            f"{period_label(p)} {s:%H:%M}-{e:%H:%M}"
            for p in periods
            for s, e in (workshop_times_for(day, p, programme_code),)
        )
        return f"Workshops on {day_label(day)} must run {parts}."
    if day:
        return f"Workshops are not scheduled on {day_label(day)}."
    morning = _WORKSHOP_MORNING
    rough_afternoon = _WORKSHOP_AFTERNOON
    thu_morning = _THURSDAY_MORNING
    return (
        f"Workshops run Monday-Friday, {morning[0]:%H:%M}-{morning[1]:%H:%M} "
        f"(Morning) or {rough_afternoon[0]:%H:%M}-{rough_afternoon[1]:%H:%M} "
        f"(Afternoon); Thursday mornings start {thu_morning[0]:%H:%M}-"
        f"{thu_morning[1]:%H:%M}."
    )


def workshop_time_issue(day, start=None, end=None, period="", programme_code=None) -> str:
    """Describe why a workshop record breaks the rules, or '' when it is fine.

    - times given: they must equal one whole standard session for the day
      (and the selected ``period`` must be that same session when provided)
    - no clock times: a period alone is fine (matrix imports), but an unknown
      day/period or no session at all is a problem
    """
    if not day:
        return ""
    if not is_workshop_day(day):
        return workshop_time_message(day)
    if start is not None or end is not None:
        if start is None or end is None:
            return "Incomplete session time — both start and end are required."
        matched = workshop_period_for_times(day, start, end, programme_code)
        if matched is None:
            return (
                workshop_time_message(day, period, programme_code)
                if period
                else workshop_time_message(day, programme_code=programme_code)
            )
        if period and matched != period:
            return workshop_time_message(day, period, programme_code)
        return ""
    if not period:
        return "No session time — set a time period or start/end times."
    periods = workshop_periods_for_day(day)
    if not periods or period not in periods:
        return workshop_time_message(day, period, programme_code)
    return ""


def legacy_workshop_allocations(queryset=None):
    """Non-destructive identification of WorkshopAllocation records that break
    the workshop-time rules. Returns [(record, issue), ...]."""
    from core.models import StudentGroup, WorkshopAllocation

    records = queryset if queryset is not None else WorkshopAllocation.objects.all()
    group_programme = dict(
        StudentGroup.objects.values_list("code", "programme__code")
    )
    issues = []
    for rec in records.iterator():
        codes = (group_programme.get(rec.group_code),) if rec.group_code else ()
        issue = workshop_time_issue(
            rec.day, rec.start_time, rec.end_time, rec.time_period,
            programme_code=codes,
        )
        if issue:
            issues.append((rec, issue))
    return issues


def legacy_workshop_sessions(queryset=None):
    """Non-destructive identification of master-timetable ``Session`` records
    with ``activity_type == WORKSHOP`` that break the workshop-time rules.
    Returns [(session, issue), ...]."""
    from core.models import ActivityType, Session

    records = (
        queryset
        if queryset is not None
        else Session.objects.filter(activity_type=ActivityType.WORKSHOP)
    )
    return [
        (rec, issue)
        for rec in records.iterator()
        if (
            issue := workshop_time_issue(
                rec.day, rec.start_time, rec.end_time,
                programme_code=session_programme_codes(rec),
            )
        )
    ]


def validate_workshop_record(record) -> list[ValidationError]:
    """Field-keyed ValidationErrors for a WorkshopAllocation (e.g. raised by
    the model's ``clean()``). Empty list means the record is rule-compliant."""
    if not record.day:
        return []
    if not is_workshop_day(record.day):
        return [ValidationError({"day": [workshop_time_message(record.day)]})]
    issue = workshop_time_issue(
        record.day, record.start_time, record.end_time, record.time_period or "",
        programme_code=allocation_programme_codes(record),
    )
    if not issue:
        return []
    if record.start_time is not None and record.end_time is not None:
        return [
            ValidationError({"start_time": [issue]}),
            ValidationError({"end_time": [issue]}),
        ]
    if record.time_period:
        return [ValidationError({"time_period": [issue]})]
    return [ValidationError({"day": [issue]})]


def validate_workshop_session(session) -> list[ValidationError]:
    """Field-keyed ValidationErrors for a WORKSHOP ``Session`` record (used by
    the model's ``clean()``). Empty list means rule-compliant."""
    if not session.day:
        return []
    if not is_workshop_day(session.day):
        return [ValidationError({"day": [workshop_time_message(session.day)]})]
    issue = workshop_time_issue(
        session.day, session.start_time, session.end_time,
        programme_code=session_programme_codes(session),
    )
    if not issue:
        return []
    if session.start_time is not None and session.end_time is not None:
        return [
            ValidationError({"start_time": [issue]}),
            ValidationError({"end_time": [issue]}),
        ]
    return [ValidationError({"day": [issue]})]