"""The measured parts: day shape, the TFI, and robust outlier statistics.

This is a direct port of the audit specification's reference implementation,
including its iteration order, its ``max(runs)`` tie-breaking and its
half-and-half weighting of the mean day against the worst day. The reason to
port rather than "improve" is that the number in the PDF has to be the number
the method describes; anything cleverer would make Appendix C a lie.

Every function here is pure: dataclasses in, dataclasses out, no database.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from statistics import mean, median, pstdev


# ─────────────────────────────────────────────────────────────────────────────
# helpers
# ─────────────────────────────────────────────────────────────────────────────


def clamp(x, lo=0.0, hi=1.0):
    return max(lo, min(hi, x))


def lin(x, good, bad):
    """1.0 at or below ``good``, 0.0 at or above ``bad``, linear between.

    The one shape used by every TFI component, so a component can be read off
    the thresholds in Appendix C by eye.
    """
    if bad == good:
        return 1.0 if x <= good else 0.0
    return 1 - clamp((x - good) / (bad - good))


def robust_z(vals):
    """Median/MAD z-scores — "unusual" measured against the cohort, not a guess.

    Falls back to a plain z-score when more than half the values are identical
    (MAD of zero), which happens for real cohorts: most groups have the same
    two-day week.
    """
    vals = [float(v) for v in vals]
    if len(vals) < 2:
        return [0.0] * len(vals)
    m = median(vals)
    mad = median(abs(v - m) for v in vals)
    if mad == 0:
        sd = pstdev(vals) or 1.0
        return [(v - m) / sd for v in vals]
    return [0.6745 * (v - m) / mad for v in vals]


def gini(values):
    """Gini coefficient of workload spread (0 = perfectly equal, →1 = one group)."""
    vals = sorted(float(v) for v in values if v is not None)
    n = len(vals)
    if n == 0:
        return 0.0
    total = sum(vals)
    if total <= 0:
        return 0.0
    weighted = sum((i + 1) * v for i, v in enumerate(vals))
    return (2 * weighted) / (n * total) - (n + 1) / n


def percentile_rank(value, population):
    """Share of ``population`` at or below ``value``, as a percentage.

    The median of an even-length list therefore ranks 50, not 50.0001.
    """
    pop = [float(v) for v in population]
    if not pop:
        return 0.0
    below = sum(1 for v in pop if v < value)
    equal = sum(1 for v in pop if v == value)
    return 100.0 * (below + 0.5 * equal) / len(pop)


def iqr(values):
    """Interquartile range, by the median-of-halves definition."""
    vals = sorted(float(v) for v in values if v is not None)
    if len(vals) < 4:
        return 0.0
    mid = len(vals) // 2
    lower = vals[:mid]
    upper = vals[-mid:]
    return median(upper) - median(lower)


# ─────────────────────────────────────────────────────────────────────────────
# 4.1 per-group, per-day metrics
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class DayMetrics:
    sessions: int = 0
    hours: float = 0.0
    span_min: int = 0
    gaps: list = field(default_factory=list)
    longest_gap: int = 0
    longest_run_min: int = 0
    longest_run_n: int = 0
    overlaps: list = field(default_factory=list)
    first_start: int = 0
    last_end: int = 0
    session_ids: list = field(default_factory=list)

    @property
    def efficiency(self) -> float:
        """Contact time as a share of time spent on campus that day."""
        if self.span_min <= 0:
            return 0.0
        return clamp(self.hours * 60 / self.span_min, 0.0, 1.0)


def analyze_day(ss, cfg):
    """Shape of one day for one group.

    ``ss`` is that group's sessions on that day, in any order. A gap of at most
    ``back_to_back_tolerance_min`` continues the current run; a negative gap is
    an overlap and is recorded as a validity problem rather than being
    silently merged.
    """
    if not ss:
        return DayMetrics()
    ss = sorted(ss, key=lambda s: (s.start, s.end))
    m = DayMetrics(
        sessions=len(ss),
        hours=sum(s.end - s.start for s in ss) / 60,
        span_min=max(s.end for s in ss) - ss[0].start,
        first_start=ss[0].start,
        last_end=max(s.end for s in ss),
        session_ids=[s.id for s in ss],
    )
    run_start, run_n, prev_end, prev_id = ss[0].start, 1, ss[0].end, ss[0].id
    runs = []
    for s in ss[1:]:
        g = s.start - prev_end
        if g < 0:
            m.overlaps.append((prev_id, s.id))
        if g <= cfg.back_to_back_tolerance_min:
            run_n += 1
        else:
            runs.append((prev_end - run_start, run_n))
            if g >= cfg.gap_min:
                m.gaps.append(g)
            run_start, run_n = s.start, 1
        if s.end >= prev_end:
            prev_end, prev_id = s.end, s.id
    runs.append((prev_end - run_start, run_n))
    m.longest_run_min, m.longest_run_n = max(runs)
    m.longest_gap = max(m.gaps, default=0)
    return m


# ─────────────────────────────────────────────────────────────────────────────
# 4.2 Timetable Friendliness Index
# ─────────────────────────────────────────────────────────────────────────────


def grade(score, cfg):
    """A/B/C/D above their floor, E below all of them."""
    floors = cfg.get("grades") or {}
    for letter, floor in sorted(floors.items(), key=lambda kv: -float(kv[1])):
        if score >= float(floor):
            return letter
    return "E"


def friendliness(days, cfg):
    """TFI for one group from its per-day metrics, or ``None`` if it has none.

    ``days`` maps weekday index -> :class:`DayMetrics`. A group with no
    sessions has no score at all (``None``), which is different from a score of
    zero: nothing was measured, and the report says so.
    """
    active = [d for d in days.values() if d.sessions]
    if not active:
        return None
    hrs = [d.hours for d in active]
    day_soft = float(cfg.day_soft_hours)
    day_hard = float(cfg.day_hard_hours)
    run_soft = float(cfg.run_soft_hours)
    run_hard = float(cfg.run_hard_hours)

    load = 0.5 * mean(lin(h, day_soft, day_hard) for h in hrs) + 0.5 * lin(
        max(hrs), day_soft, day_hard
    )
    dead = sum(sum(d.gaps) for d in active)
    contact = sum(hrs) * 60
    dead_ratio = (dead / contact) if contact > 0 else 0.0
    gaps = 0.5 * lin(dead_ratio, 0.0, float(cfg.dead_ratio_bad)) + 0.5 * lin(
        max(d.longest_gap for d in active), float(cfg.gap_min), float(cfg.gap_bad_min)
    )
    cont = lin(
        max(d.longest_run_min for d in active) / 60, run_soft, run_hard
    )
    cv = (pstdev(hrs) / mean(hrs)) if len(hrs) > 1 and mean(hrs) > 0 else 0.0
    balance = lin(cv, 0.15, 0.8)
    parts = {"load": load, "gaps": gaps, "continuity": cont, "balance": balance}
    weights = cfg.get("weights") or {}
    score = 100 * sum(float(weights.get(k, 0.0)) * v for k, v in parts.items())
    score = max(0.0, min(100.0, score))

    return {
        "score": round(score, 1),
        "parts": {k: round(v, 4) for k, v in parts.items()},
        "grade": grade(score, cfg),
        "dead_min": dead,
        "contact_min": contact,
        "dead_ratio": dead_ratio,
        "busiest_day_h": max(hrs),
        "longest_gap": max(d.longest_gap for d in active),
        "longest_run_min": max(d.longest_run_min for d in active),
        "days_used": len(active),
        "week_hours": sum(hrs),
        "cv": cv,
    }


def group_days(sessions, working_days) -> dict:
    """Bucket one group's sessions into ``{weekday: [sessions]}``."""
    days: dict = {}
    for index in working_days:
        days[index] = []
    for sess in sessions:
        days.setdefault(sess.day, []).append(sess)
    return days


def score_group(sessions, cfg) -> dict:
    """Day metrics + TFI for a single group, in one call.

    Returns ``{day: DayMetrics}`` under ``days`` and the friendliness dict
    under ``tfi`` (``None`` when the group has no sessions at all).
    """
    working_days = list(cfg.get("working_days") or [0, 1, 2, 3, 4])
    days = group_days(sessions, working_days)
    metrics = {index: analyze_day(rows, cfg) for index, rows in days.items()}
    tfi = friendliness(metrics, cfg)
    if tfi is not None:
        active = [(index, m) for index, m in sorted(metrics.items()) if m.sessions]
        busiest = max(active, key=lambda pair: pair[1].hours)
        gap_day = max(active, key=lambda pair: pair[1].longest_gap)
        tfi["busiest_day"] = busiest[0]
        tfi["busiest_day_h"] = round(busiest[1].hours, 2)
        tfi["longest_gap_day"] = gap_day[0] if gap_day[1].longest_gap else None
        tfi["overlaps"] = [
            pair
            for index in sorted(metrics)
            for pair in metrics[index].overlaps
        ]
        tfi["span_min"] = sum(d.span_min for d in metrics.values() if d.sessions)
        tfi["free_days"] = sum(
            1 for index in working_days if not metrics.get(index)
            or not metrics[index].sessions
        )
        tfi["efficiency"] = (
            clamp(tfi["contact_min"] / tfi["span_min"], 0.0, 1.0)
            if tfi["span_min"] > 0
            else 0.0
        )
    return {"days": metrics, "tfi": tfi}
