"""The audit engine: allocation, venues, issues, roll-ups and the verdict.

Everything here is a pure function of the :class:`audit.collect.AuditDataset`
plus the config. No queries, no writes, no templates — which is why the
fixture can be analysed in a test and why re-running with different thresholds
can never change the timetable.

Two verdicts are kept apart on purpose and are never merged:

* **Validity** — did anything break a hard constraint? Conflicts, missing
  venues, invalid times, over-capacity rooms, unresolved requirements.
* **Friendliness** — is the timetable pleasant to live with? Dead time, brutal
  days, unbalanced weeks. A group can be perfectly valid and still grade D.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from statistics import mean, median

from audit.collect import (
    DAY_NAMES,
    ORIGIN_NEW,
    ORIGIN_REASSIGNED,
    ORIGIN_RETAINED,
    ORIGIN_UNKNOWN,
    AuditDataset,
    programme_of,
)
from audit.metrics import (
    clamp,
    gini,
    grade,
    iqr,
    percentile_rank,
    robust_z,
    score_group,
)

SEVERITY_ORDER = {"Critical": 0, "Review": 1, "Watch": 2, "Info": 3}
SEVERITY_CHIP = {
    "Critical": "critical",
    "Review": "review",
    "Watch": "watch",
    "Info": "ok",
}
SMALL_GROUP_ACTIVITIES = ("seminar", "tutorial", "practical")

ISSUE_CATEGORIES = {
    "unresolved-requirement": "Unresolved requirement",
    "missing-required-session": "Missing required session",
    "no-timetabled-classes": "Group with no timetabled classes",
    "insufficient-venue-capacity": "Insufficient venue capacity",
    "missing-venue-info": "Missing or invalid venue information",
    "timetable-conflict": "Timetable conflict (overlap)",
    "invalid-session-time": "Invalid session time",
    "unverifiable-workshop-time": "Unverifiable workshop time",
    "unconfigured-requirements": "Course with unconfigured requirements",
    "friendliness-concern": "Friendliness concern",
}

DETECTED = "System-detected"
JUDGMENT = "Needs human judgment"


# ─────────────────────────────────────────────────────────────────────────────
# results
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class Issue:
    id: str
    category: str
    severity: str
    affects_validity: bool
    detection: str
    course: str | None
    groups: list
    session_ids: list
    description: str
    evidence: dict
    action: str
    students: int = 0
    hours: float = 0.0
    why_detected: str = ""

    @property
    def category_label(self) -> str:
        return ISSUE_CATEGORIES.get(self.category, self.category)

    @property
    def chip(self) -> str:
        return SEVERITY_CHIP.get(self.severity, "ok")

    def as_dict(self) -> dict:
        return {
            "id": self.id,
            "category": self.category,
            "category_label": self.category_label,
            "severity": self.severity,
            "affects_validity": self.affects_validity,
            "detection": self.detection,
            "course": self.course,
            "groups": list(self.groups),
            "session_ids": list(self.session_ids),
            "description": self.description,
            "evidence": dict(self.evidence),
            "action": self.action,
            "students": self.students,
            "hours": self.hours,
        }


@dataclass
class GroupResult:
    code: str
    programme: str
    size: int
    sessions: list
    days: dict
    tfi: dict | None
    day_hours: dict = field(default_factory=dict)
    flags: list = field(default_factory=list)
    issues: list = field(default_factory=list)
    percentile_campus: float = 0.0
    percentile_programme: float = 0.0
    early_late_share: float = 0.0
    requirements_total: int = 0
    requirements_unresolved: int = 0
    day_slots: dict = field(default_factory=dict)
    gap_bands: dict = field(default_factory=dict)
    clash_ids: set = field(default_factory=set)

    @property
    def score(self) -> float:
        return float(self.tfi["score"]) if self.tfi else 0.0

    @property
    def grade_letter(self) -> str:
        return self.tfi["grade"] if self.tfi else "—"

    @property
    def week_hours(self) -> float:
        return round(self.tfi["week_hours"], 2) if self.tfi else 0.0

    @property
    def is_flagged(self) -> bool:
        return bool(self.flags) or self.grade_letter in {"D", "E"}

    @property
    def worst_component(self) -> tuple:
        """The component dragging this group's score down most (lowest value)."""
        if not self.tfi:
            return ("n/a", 1.0)
        parts = self.tfi["parts"]
        key = min(parts, key=lambda k: parts[k])
        return (key, parts[key])

    def as_row(self) -> dict:
        return {
            "group": self.code,
            "programme": self.programme,
            "score": self.tfi["score"] if self.tfi else None,
            "grade": self.grade_letter,
            "load": self.tfi["parts"]["load"] if self.tfi else None,
            "gaps": self.tfi["parts"]["gaps"] if self.tfi else None,
            "continuity": self.tfi["parts"]["continuity"] if self.tfi else None,
            "balance": self.tfi["parts"]["balance"] if self.tfi else None,
            "week_hours": self.week_hours,
            "busiest_day": DAY_NAMES[self.tfi["busiest_day"]] if self.tfi and self.tfi.get("busiest_day") is not None else None,
            "busiest_day_h": self.tfi["busiest_day_h"] if self.tfi else None,
            "longest_gap": self.tfi["longest_gap"] if self.tfi else 0,
            "longest_run": self.tfi["longest_run_min"] if self.tfi else 0,
            "dead_ratio": round(self.tfi["dead_ratio"], 4) if self.tfi else 0.0,
            "days_used": self.tfi["days_used"] if self.tfi else 0,
            "percentile": round(self.percentile_campus, 1),
        }


@dataclass
class Analysis:
    dataset: AuditDataset
    config: object
    groups: list = field(default_factory=list)
    group_by_code: dict = field(default_factory=dict)
    unscored_groups: list = field(default_factory=list)
    allocation: dict = field(default_factory=dict)
    venues: dict = field(default_factory=dict)
    observations: list = field(default_factory=list)
    concurrency: dict = field(default_factory=dict)
    issues: list = field(default_factory=list)
    pareto: list = field(default_factory=list)
    rollup: dict = field(default_factory=dict)
    performance: dict = field(default_factory=dict)
    verdict_label: str = ""
    verdict_colour: str = "green"
    kpis: dict = field(default_factory=dict)
    groups_by_tfi: list = field(default_factory=list)
    flagged: list = field(default_factory=list)

    # -- convenience accessors used by the report ---------------------------
    @property
    def meta(self):
        return self.dataset.meta

    @property
    def valid_issues(self) -> list:
        return [i for i in self.issues if i.affects_validity]

    @property
    def quality_issues(self) -> list:
        return [i for i in self.issues if not i.affects_validity]

    @property
    def needs_judgment(self) -> list:
        return [i for i in self.issues if i.detection == JUDGMENT]

    @property
    def critical_issues(self) -> list:
        return [i for i in self.issues if i.severity == "Critical"]

    def issues_for(self, group_code: str) -> list:
        return [i for i in self.issues if group_code in (i.groups or [])]

    def severity_counts(self) -> dict:
        counts = {k: 0 for k in SEVERITY_ORDER}
        for issue in self.issues:
            counts[issue.severity] = counts.get(issue.severity, 0) + 1
        return counts

    def validity_summary(self) -> dict:
        valid = self.valid_issues
        return {
            "valid": not valid,
            "count": len(valid),
            "students": sum(i.students for i in valid),
            "groups": len({g for i in valid for g in (i.groups or [])}),
            "critical": len([i for i in valid if i.severity == "Critical"]),
        }


# ─────────────────────────────────────────────────────────────────────────────
# top level
# ─────────────────────────────────────────────────────────────────────────────


def analyse(dataset: AuditDataset) -> Analysis:
    """The whole audit, in one call. Read-only with respect to the database."""
    cfg = dataset.config
    analysis = Analysis(dataset=dataset, config=cfg)

    analysis.groups = _analyse_groups(dataset, cfg)
    analysis.group_by_code = {g.code: g for g in analysis.groups}
    analysis.unscored_groups = sorted(set(dataset.groups) - set(analysis.group_by_code))
    analysis.groups_by_tfi = sorted(
        analysis.groups, key=lambda g: (g.score if g.tfi else 999)
    )
    analysis.flagged = [
        g
        for g in analysis.groups_by_tfi
        if g.is_flagged or g.issues
    ]

    analysis.allocation = allocation_metrics(dataset, analysis.group_by_code)
    analysis.venues = venue_metrics(dataset.sessions, cfg)
    # Computed once here: the issue register and the report both read the same
    # list, and a second pass over every physical class is not free.
    analysis.observations = overflow_observations(dataset.sessions, cfg)
    analysis.concurrency = concurrency_heatmap(dataset.sessions, cfg)
    analysis.issues = build_issues(dataset, analysis, cfg)
    _renumber(analysis.issues)
    analysis.pareto = pareto(analysis.issues)
    analysis.rollup = _rollups(dataset, analysis, cfg)
    analysis.performance = performance_metrics(dataset)
    analysis.kpis = kpis(dataset, analysis)
    analysis.verdict_label, analysis.verdict_colour = verdict(
        analysis.issues, analysis.allocation["completion_pct"]
    )
    for issue in analysis.issues:
        for code in issue.groups or []:
            if code in analysis.group_by_code:
                analysis.group_by_code[code].issues.append(issue)
    return analysis


# ─────────────────────────────────────────────────────────────────────────────
# 4.4 group-level analytics
# ─────────────────────────────────────────────────────────────────────────────


def _analyse_groups(dataset: AuditDataset, cfg) -> list:
    pattern = cfg.get_path("programme_regex")
    early = cfg.get_path("early_late.early_before", 510)
    late = cfg.get_path("early_late.late_after", 1020)
    out = []
    # Every group in scope, not just the ones that turned up in a session: a
    # group with no classes at all still has to be scored (to nothing) so the
    # report's denominators add up and the register can say so. Skipping it here
    # is what made "80 groups in scope" quietly become 77.
    for code in dataset.groups:
        sessions = dataset.sessions_by_group.get(code, [])
        result = score_group(sessions, cfg)
        tfi = result["tfi"]
        reqs = dataset.requirements_for(code)
        day_hours = {
            index: round(m.hours, 2) for index, m in result["days"].items() if m.sessions
        }
        clash_ids = set()
        for pairs in (tfi or {}).get("overlaps", []):
            clash_ids.update(pairs)
        out.append(
            GroupResult(
                code=code,
                programme=dataset.group_programme.get(code) or programme_of(code, pattern),
                size=dataset.group_size.get(code) or 0,
                sessions=sessions,
                days=result["days"],
                tfi=tfi,
                day_hours=day_hours,
                requirements_total=len(reqs),
                requirements_unresolved=sum(
                    1 for r in reqs if r.status == "unresolved"
                ),
                early_late_share=round(_early_late_share(sessions, early, late), 4),
                clash_ids=clash_ids,
                day_slots={d: [s for s in sessions if s.day == d] for d in day_hours},
                gap_bands=_gap_bands(sessions, cfg),
            )
        )
    _flag_groups(out, cfg)
    _rank_groups(out)
    return out


def _early_late_share(sessions, early: int, late: int) -> float:
    if not sessions:
        return 0.0
    hits = sum(1 for s in sessions if s.start < early or s.end > late)
    return hits / len(sessions)


def _gap_bands(sessions, cfg) -> dict:
    """Absolute ``{day: [{start, len, label}]}`` bands, for the Appendix A cards."""
    threshold = int(cfg.gap_min)
    bands: dict = defaultdict(list)
    for day in {s.day for s in sessions}:
        rows = sorted([s for s in sessions if s.day == day], key=lambda s: (s.start, s.end))
        prev_end = None
        for sess in rows:
            if prev_end is not None:
                gap = sess.start - prev_end
                if gap >= threshold:
                    bands[day].append(
                        {
                            "start": prev_end,
                            "len": gap,
                            "label": _fmt_dur(gap),
                        }
                    )
            prev_end = max(prev_end or 0, sess.end)
    return dict(bands)


def _fmt_dur(minutes) -> str:
    minutes = int(round(minutes))
    hours, mins = divmod(abs(minutes), 60)
    if hours and mins:
        return f"{hours}h {mins:02d}m"
    if hours:
        return f"{hours}h"
    return f"{mins}m"


def _flag_groups(groups: list, cfg) -> None:
    """Threshold breaches *and* relative outliers, with the reason recorded.

    A group is flagged when it breaks a hard threshold, when it is a robust
    outlier against its peers, or both. Recording which of the two fired is the
    point: "unusual" and "unacceptable" are different claims.
    """
    z_limit = float(cfg.outlier_z)
    if not groups:
        return
    scores = [g.tfi["score"] for g in groups if g.tfi]
    gaps = [g.tfi["longest_gap"] for g in groups if g.tfi]
    busy = [g.tfi["busiest_day_h"] for g in groups if g.tfi]
    dead = [g.tfi["dead_ratio"] for g in groups if g.tfi]
    z_scores = dict(zip([g.code for g in groups if g.tfi], robust_z(scores)))
    z_gaps = dict(zip([g.code for g in groups if g.tfi], robust_z(gaps)))
    z_busy = dict(zip([g.code for g in groups if g.tfi], robust_z(busy)))
    z_dead = dict(zip([g.code for g in groups if g.tfi], robust_z(dead)))

    gap_bad = int(cfg.gap_bad_min)
    day_hard = float(cfg.day_hard_hours)
    for group in groups:
        if not group.tfi:
            continue
        tfi = group.tfi
        if tfi["grade"] in {"D", "E"}:
            group.flags.append(
                {
                    "metric": "TFI",
                    "value": tfi["score"],
                    "grade": tfi["grade"],
                    "why": "threshold",
                    "z": z_scores.get(group.code, 0.0),
                    "note": f"TFI {tfi['score']} (grade {tfi['grade']})",
                }
            )
        elif z_scores.get(group.code, 0.0) >= z_limit:
            group.flags.append(
                {
                    "metric": "TFI",
                    "value": tfi["score"],
                    "grade": tfi["grade"],
                    "why": "outlier",
                    "z": z_scores[group.code],
                    "note": f"TFI {tfi['score']}, {z_scores[group.code]:.1f}σ below the cohort median",
                }
            )
        if tfi["longest_gap"] >= gap_bad:
            group.flags.append(
                {
                    "metric": "Longest gap",
                    "value": tfi["longest_gap"],
                    "why": "threshold",
                    "z": z_gaps.get(group.code, 0.0),
                    "note": f"{_fmt_dur(tfi['longest_gap'])} gap"
                    + (
                        f" on {DAY_NAMES[tfi['longest_gap_day']]}"
                        if tfi.get("longest_gap_day") is not None
                        else ""
                    ),
                }
            )
        elif z_gaps.get(group.code, 0.0) >= z_limit:
            group.flags.append(
                {
                    "metric": "Longest gap",
                    "value": tfi["longest_gap"],
                    "why": "outlier",
                    "z": z_gaps[group.code],
                    "note": f"{_fmt_dur(tfi['longest_gap'])} gap, "
                    f"{z_gaps[group.code]:.1f}σ above the cohort median",
                }
            )
        if tfi["busiest_day_h"] >= day_hard:
            group.flags.append(
                {
                    "metric": "Longest day",
                    "value": tfi["busiest_day_h"],
                    "why": "threshold",
                    "z": z_busy.get(group.code, 0.0),
                    "note": f"{tfi['busiest_day_h']:.1f} h on "
                    f"{DAY_NAMES[tfi['busiest_day']] if tfi.get('busiest_day') is not None else 'its busiest day'}",
                }
            )
        elif z_busy.get(group.code, 0.0) >= z_limit:
            group.flags.append(
                {
                    "metric": "Longest day",
                    "value": tfi["busiest_day_h"],
                    "why": "outlier",
                    "z": z_busy[group.code],
                    "note": f"{tfi['busiest_day_h']:.1f} h on its busiest day, "
                    f"{z_busy[group.code]:.1f}σ above the cohort median",
                }
            )
        if tfi["dead_ratio"] >= float(cfg.dead_ratio_bad) and tfi["contact_min"] > 0:
            group.flags.append(
                {
                    "metric": "Dead time",
                    "value": round(tfi["dead_ratio"], 3),
                    "why": "threshold",
                    "z": z_dead.get(group.code, 0.0),
                    "note": f"{tfi['dead_ratio'] * 100:.0f}% of contact time is gap",
                }
            )
        elif z_dead.get(group.code, 0.0) >= z_limit:
            group.flags.append(
                {
                    "metric": "Dead time",
                    "value": round(tfi["dead_ratio"], 3),
                    "why": "outlier",
                    "z": z_dead[group.code],
                    "note": f"{tfi['dead_ratio'] * 100:.0f}% dead time, "
                    f"{z_dead[group.code]:.1f}σ above the cohort median",
                }
            )


def _rank_groups(groups: list) -> None:
    scores = [g.tfi["score"] for g in groups if g.tfi]
    by_programme: dict = defaultdict(list)
    for group in groups:
        if group.tfi:
            by_programme[group.programme].append(group.tfi["score"])
    for group in groups:
        if not group.tfi:
            continue
        group.percentile_campus = percentile_rank(group.tfi["score"], scores)
        group.percentile_programme = percentile_rank(
            group.tfi["score"], by_programme.get(group.programme, [])
        )


# ─────────────────────────────────────────────────────────────────────────────
# 4.5 allocation metrics
# ─────────────────────────────────────────────────────────────────────────────


def allocation_metrics(dataset: AuditDataset, group_by_code: dict) -> dict:
    """Requirement counts by activity, course and group, plus the origin split."""
    reqs = dataset.requirements
    overall = _empty_bucket()
    by_activity: dict = defaultdict(_empty_bucket)
    by_course: dict = defaultdict(_empty_bucket)
    by_group: dict = defaultdict(_empty_bucket)
    origins = {ORIGIN_RETAINED: 0, ORIGIN_NEW: 0, ORIGIN_REASSIGNED: 0, ORIGIN_UNKNOWN: 0}

    convenient = 0
    convenient_total = 0
    for req in reqs:
        for bucket in (overall, by_activity[req.activity], by_course[req.course], by_group[req.group]):
            _tally(bucket, req)
        origins[req.origin] = origins.get(req.origin, 0) + 1
        if req.status == "allocated":
            convenient_total += 1
            group = group_by_code.get(req.group)
            if group and group.tfi and group.grade_letter in {"A", "B", "C"}:
                convenient += 1

    for bucket in by_activity.values():
        bucket["completion_pct"] = _pct(bucket["allocated"], bucket["total"])
    for bucket in by_course.values():
        bucket["completion_pct"] = _pct(bucket["allocated"], bucket["total"])
    for bucket in by_group.values():
        bucket["completion_pct"] = _pct(bucket["allocated"], bucket["total"])
    overall["completion_pct"] = _pct(overall["allocated"], overall["total"])

    return {
        "overall": overall,
        "by_activity": {
            k: by_activity[k] for k in sorted(by_activity)
        },
        "by_course": dict(
            sorted(
                by_course.items(),
                key=lambda kv: (-kv[1]["unresolved"], -kv[1]["total"], kv[0]),
            )
        ),
        "by_group": dict(
            sorted(
                by_group.items(),
                key=lambda kv: (-kv[1]["unresolved"], -kv[1]["total"], kv[0]),
            )
        ),
        "completion_pct": overall["completion_pct"],
        "origins": origins,
        "convenience_adjusted_completion": _pct(convenient, convenient_total),
        "convenience_allocated": convenient,
        "convenience_total": convenient_total,
        "unresolved_by_reasons": _reason_counts(reqs),
    }


def _empty_bucket() -> dict:
    return {
        "total": 0,
        "allocated": 0,
        "unresolved": 0,
        "completion_pct": 0.0,
        "origins": {},
        "reasons": [],
    }


def _tally(bucket, req) -> None:
    bucket["total"] += 1
    if req.status == "allocated":
        bucket["allocated"] += 1
        bucket["origins"][req.origin] = bucket["origins"].get(req.origin, 0) + 1
    else:
        bucket["unresolved"] += 1
        bucket["reasons"].append(req.reason)


def _pct(part, whole) -> float:
    if not whole:
        return 0.0
    return round(100.0 * part / whole, 1)


def _reason_counts(reqs) -> list:
    counts: dict = defaultdict(int)
    for req in reqs:
        if req.status == "unresolved" and req.reason:
            counts[req.reason] += 1
    return [
        {"reason": reason, "count": count}
        for reason, count in sorted(counts.items(), key=lambda kv: -kv[1])
    ]


# ─────────────────────────────────────────────────────────────────────────────
# 4.6 venue metrics
# ─────────────────────────────────────────────────────────────────────────────


def physical_sessions(sessions) -> dict:
    """Collapse one class attended by many groups into a single physical class.

    Keyed on ``(venue, day, start, end, course, activity)``: several groups
    sharing a room is one room booking, and counting it five times would
    overstate both student numbers and utilisation.
    """
    phys: dict = {}
    for s in sessions:
        key = (s.venue, s.day, s.start, s.end, s.course, s.activity)
        entry = phys.setdefault(
            key,
            {
                "groups": set(),
                "students": 0,
                "cap": s.capacity,
                "session_ids": [],
                "venue": s.venue,
                "day": s.day,
                "start": s.start,
                "end": s.end,
                "course": s.course,
                "activity": s.activity,
            },
        )
        entry["session_ids"].append(s.id)
        if s.group not in entry["groups"]:
            entry["groups"].add(s.group)
            entry["students"] += s.group_size or 0
        if s.capacity and not entry["cap"]:
            entry["cap"] = s.capacity
    return phys


def venue_metrics(sessions, cfg) -> dict:
    """Seat utilisation and time utilisation per venue (specification 4.6)."""
    phys = physical_sessions(sessions)
    by_venue: dict = defaultdict(list)
    for (venue, d, st, en, *_), p in phys.items():
        by_venue[venue].append((d, st, en, p))
    available = float(cfg.get_path("venue.available_hours_per_week", 50))
    out = {}
    for venue, items in by_venue.items():
        cap = next((p["cap"] for *_, p in items if p["cap"]), None)
        seat = [p["students"] / p["cap"] for *_, p in items if p["cap"]]
        booked = sum(en - st for _, st, en, _ in items) / 60
        entry = {
            "venue": venue or "(no venue)",
            "has_venue": bool(venue),
            "capacity": cap,
            "sessions": len(items),
            "groups_per_session": mean(len(p["groups"]) for *_, p in items),
            "avg_students": mean(p["students"] for *_, p in items),
            "avg_groups": mean(len(p["groups"]) for *_, p in items),
            "seat_util_avg": mean(seat) if seat else None,
            "seat_util_max": max(seat) if seat else None,
            "unused_seats_avg": (
                cap - mean(p["students"] for *_, p in items) if cap else None
            ),
            "overflow_sessions": sum(1 for x in seat if x > 1.0),
            "time_util": booked / available if available else None,
            "booked_hours": booked,
            "students": sum(p["students"] for *_, p in items),
            "groups": sorted({g for *_, p in items for g in p["groups"]}),
        }
        out[venue or "(no venue)"] = entry
    return out


def overflow_observations(sessions, cfg) -> list:
    """Every physical class whose group count outgrew its room, with numbers."""
    low = float(cfg.get_path("venue.low_seat_util", 0.35))
    high = float(cfg.get_path("venue.high_seat_util", 0.95))
    out = []
    for key, p in physical_sessions(sessions).items():
        cap = p["cap"]
        if not cap:
            continue
        ratio = p["students"] / cap
        if ratio <= 1.0 and ratio >= low:
            continue
        out.append(
            {
                "venue": p["venue"] or "(no venue)",
                "course": p["course"],
                "activity": p["activity"],
                "day": p["day"],
                "start": p["start"],
                "end": p["end"],
                "students": p["students"],
                "capacity": cap,
                "ratio": ratio,
                "groups": sorted(p["groups"]),
                "kind": "overflow" if ratio > 1.0 else "under-used",
                "session_ids": p["session_ids"],
            }
        )
    out.sort(key=lambda r: -r["ratio"])
    return out


def concurrency_heatmap(sessions, cfg) -> dict:
    """Students in class and rooms in use, per (day × 30-minute slot)."""
    step = 30
    start = int(cfg.get_path("working_day.start", 480))
    end = int(cfg.get_path("working_day.end", 1080))
    days = list(cfg.get("working_days") or [0, 1, 2, 3, 4])
    slots = list(range(start, end, step))
    phys = physical_sessions(sessions)
    students: dict = {d: {t: 0 for t in slots} for d in days}
    rooms: dict = {d: {t: 0 for t in slots} for d in days}
    for p in phys.values():
        if p["day"] not in students:
            students[p["day"]] = {t: 0 for t in slots}
            rooms[p["day"]] = {t: 0 for t in slots}
        for t in slots:
            if p["start"] < t + step and p["end"] > t:
                students[p["day"]][t] += p["students"]
                if p["venue"]:
                    rooms[p["day"]][t] += 1
    peak = max(
        ((students[d][t], d, t) for d in students for t in slots),
        default=(0, None, None),
    )
    quietest = min(
        ((students[d][t], d, t) for d in students for t in slots),
        default=(0, None, None),
    )
    return {
        "days": days,
        "slots": slots,
        "students": students,
        "rooms": rooms,
        "peak": {"students": peak[0], "day": peak[1], "slot": peak[2]},
        "quietest": {"students": quietest[0], "day": quietest[1], "slot": quietest[2]},
        "max": peak[0] or 1,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 4.6b system performance
# ─────────────────────────────────────────────────────────────────────────────


def performance_metrics(dataset: AuditDataset) -> dict:
    meta = dataset.meta
    requirements = meta.requirement_total or sum(
        1 for r in dataset.requirements
    )
    duration = meta.duration_s
    throughput = None
    if duration and duration > 0 and requirements:
        throughput = round(requirements / duration, 1)
    return {
        "run_id": meta.run_id,
        "status": meta.status,
        "applied": meta.applied,
        "algorithm": meta.algorithm or "Not recorded",
        "scope": meta.scope,
        "started_at": meta.started_at,
        "duration_s": duration,
        "candidates_examined": meta.candidates_examined,
        "backtracks": meta.backtracks,
        "search_limit_reached": meta.search_limit_reached,
        "completed": meta.completed,
        "requirements": requirements,
        "throughput": throughput,
        "assigned": meta.assigned,
        "moved": meta.moved,
        "retained": meta.retained,
        "removed": meta.removed,
        "unresolved": meta.unresolved,
        "errors": list(meta.errors),
        "warnings": list(meta.warnings),
        "run_pk": meta.run_pk,
    }


def run_trend(semester, exclude_pk=None) -> list:
    """Every stored run of the semester with its performance counters.

    Read-only query used by the section 6 trend chart; a single point renders as
    KPI tiles rather than a one-dot line.
    """
    from core.models import AllocationRun

    rows = []
    for run in AllocationRun.objects.filter(semester=semester).order_by("created_at"):
        if exclude_pk and run.pk == exclude_pk:
            continue
        plan = run.plan() or {}
        rows.append(
            {
                "run_id": run.pk,
                "created_at": run.created_at,
                "requirements": plan.get("requirement_total") or 0,
                "duration_s": (
                    round(plan["duration_ms"] / 1000, 2) if plan.get("duration_ms") else None
                ),
                "unresolved": run.unresolved,
                "status": run.status,
            }
        )
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# roll-ups
# ─────────────────────────────────────────────────────────────────────────────


def _rollups(dataset: AuditDataset, analysis: Analysis, cfg) -> dict:
    days = list(cfg.get("working_days") or [0, 1, 2, 3, 4])
    programme_rollup: dict = defaultdict(
        lambda: {
            "groups": 0,
            "scores": [],
            "grades": defaultdict(int),
            "day_hours": {d: [] for d in days},
            "week_hours": [],
        }
    )
    for group in analysis.groups:
        if not group.tfi:
            continue
        entry = programme_rollup[group.programme]
        entry["groups"] += 1
        entry["scores"].append(group.tfi["score"])
        entry["grades"][group.grade_letter] += 1
        entry["week_hours"].append(group.week_hours)
        for day in days:
            value = group.day_hours.get(day)
            if value is not None:
                entry["day_hours"][day].append(value)

    programmes = []
    for code, entry in sorted(programme_rollup.items()):
        hours = entry["week_hours"]
        scores = entry["scores"]
        low = sum(1 for g in analysis.groups if g.programme == code and g.grade_letter in {"D", "E"})
        mean_day = {
            d: round(mean(v), 2) if v else 0.0 for d, v in entry["day_hours"].items()
        }
        worst_day = max(mean_day, key=lambda d: mean_day[d]) if any(mean_day.values()) else None
        spread = (max(hours) - min(hours)) if hours else 0.0
        programmes.append(
            {
                "programme": code,
                "groups": entry["groups"],
                "mean_tfi": round(mean(scores), 1) if scores else None,
                "grades": {k: entry["grades"].get(k, 0) for k in "ABCDE"},
                "at_de": low,
                "at_de_pct": round(100 * low / entry["groups"], 1) if entry["groups"] else 0.0,
                "mean_week_hours": round(mean(hours), 2) if hours else 0.0,
                "hours_spread": round(spread, 2),
                "hours_iqr": round(iqr(hours), 2),
                "gini": round(gini(hours), 4),
                "mean_day_hours": mean_day,
                "worst_day": DAY_NAMES[worst_day] if worst_day is not None else None,
                "equity_flag": gini(hours) > 0.10 and spread >= 2.0,
            }
        )

    # Campus-wide day pressure: student-hours carried by each weekday.
    day_pressure = {d: 0.0 for d in days}
    for group in analysis.groups:
        if not group.tfi:
            continue
        for day, hours in group.day_hours.items():
            day_pressure[day] = day_pressure.get(day, 0.0) + hours
    pressure_rows = [
        {"day": d, "label": DAY_NAMES[d], "hours": round(day_pressure.get(d, 0.0), 1)}
        for d in days
    ]
    live = [r for r in pressure_rows if r["hours"] > 0]
    heaviest = max(live, key=lambda r: r["hours"]) if live else None
    lightest = min(live, key=lambda r: r["hours"]) if live else None

    campus_dead = [g.tfi["dead_ratio"] for g in analysis.groups if g.tfi]
    campus_scores = [g.tfi["score"] for g in analysis.groups if g.tfi]
    return {
        "programmes": programmes,
        "day_pressure": pressure_rows,
        "heaviest_day": heaviest,
        "lightest_day": lightest,
        "campus_mean_dead_ratio": round(mean(campus_dead), 4) if campus_dead else 0.0,
        "campus_mean_tfi": round(mean(campus_scores), 1) if campus_scores else None,
        "campus_median_tfi": round(median(campus_scores), 1) if campus_scores else None,
        "campus_dead_min": round(mean(campus_dead) * 100, 1) if campus_dead else 0.0,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 4.7 issue register
# ─────────────────────────────────────────────────────────────────────────────


def build_issues(dataset: AuditDataset, analysis: Analysis, cfg) -> list:
    """Every problem the audit can evidence, in one register."""
    issues: list = []
    issues += _issue_unresolved(dataset, analysis, cfg)
    issues += _issue_missing_required_session(dataset, analysis, cfg)
    issues += _issue_no_classes(analysis)
    issues += _issue_capacity(dataset, analysis, cfg)
    issues += _issue_venue_info(dataset, analysis, cfg)
    issues += _issue_conflicts(dataset, analysis, cfg)
    issues += _issue_invalid_times(dataset, analysis, cfg)
    issues += _issue_workshops(dataset, analysis, cfg)
    issues += _issue_unconfigured(dataset, analysis, cfg)
    issues += _issue_friendliness(analysis, cfg)
    issues.sort(
        key=lambda i: (
            SEVERITY_ORDER.get(i.severity, 9),
            -i.students,
            -len(i.groups or []),
            i.category,
        )
    )
    return issues


def _issue_unresolved(dataset, analysis, cfg):
    buckets: dict = defaultdict(list)
    for req in dataset.requirements:
        if req.status == "unresolved":
            buckets[(req.course, req.activity, req.reason or "No reason recorded")].append(
                req.group
            )
    out = []
    for (course, activity, reason), groups in sorted(buckets.items()):
        groups = sorted(set(groups))
        out.append(
            Issue(
                id="",
                category="unresolved-requirement",
                severity="Critical",
                affects_validity=True,
                detection=DETECTED,
                course=course,
                groups=groups,
                session_ids=[],
                description=(
                    f"{len(groups)} group{'s' if len(groups) != 1 else ''} "
                    f"{'have' if len(groups) != 1 else 'has'} no {activity} "
                    f"session for {course}."
                ),
                evidence={
                    "requirements": len(groups),
                    "groups": groups,
                    "activity": activity,
                    "reason": reason,
                },
                action=(
                    f"Timetable an extra {course} {activity} session with room for "
                    f"{len(groups)} group{'s' if len(groups) != 1 else ''}, or relax the "
                    f"course's requirement."
                ),
                students=len(groups) * (dataset.group_size.get(groups[0]) or 0),
            )
        )
    return out


def _issue_missing_required_session(dataset, analysis, cfg):
    """A requirement exists but the timetable has no session of that type at all."""
    from core.models import Course

    wanted = defaultdict(set)   # (course, activity) -> groups that need it
    for req in dataset.requirements:
        if req.status == "unresolved":
            wanted[(req.course, req.activity)].add(req.group)
    if not wanted:
        return []
    available: dict = defaultdict(set)
    for sess in dataset.sessions:
        if sess.activity in SMALL_GROUP_ACTIVITIES:
            available[(sess.course, sess.activity)].add(sess.id)
    from core.models import Course

    courses = {code for code, _ in wanted}
    configured = {
        c.code: (c.requirement_map(), c.name)
        for c in Course.objects.filter(code__in=courses)
    }
    out = []
    for (course, activity), groups in sorted(wanted.items()):
        entry = configured.get(course)
        if entry and not entry[0].get(activity):
            continue    # the requirement itself is not configured: separate issue
        if available.get((course, activity)):
            continue    # sessions exist, so this is a clash/capacity problem
        out.append(
            Issue(
                id="",
                category="missing-required-session",
                severity="Critical",
                affects_validity=True,
                detection=DETECTED,
                course=course,
                groups=sorted(groups),
                session_ids=[],
                description=(
                    f"The timetable has no {activity} session for {course} at all, so "
                    f"{len(groups)} requirement(s) cannot be satisfied by any session."
                ),
                evidence={
                    "activity": activity,
                    "requirements": len(groups),
                    "course_name": (entry[1] if entry else ""),
                    "searched": f"every {activity} session in the semester",
                },
                action=f"Add at least one {course} {activity} session to the master timetable.",
                students=len(groups) * (dataset.group_size.get(sorted(groups)[0]) or 0),
            )
        )
    return out


def _issue_no_classes(analysis):
    """Groups that are in scope but carry no classes at all.

    A group with an empty timetable has no TFI, so it drops out of every
    ranking and every average without a word: the report would say "80 groups
    in scope" while having analysed 77, and the three it silently dropped are
    exactly the ones nobody is timetabling. It is a timetable problem before it
    is a friendliness one, so it belongs in the register.
    """
    empty = [g for g in analysis.groups if not g.sessions]
    if not empty:
        return []
    by_programme: dict = defaultdict(list)
    for group in empty:
        by_programme[group.programme].append(group.code)
    out = []
    for programme, codes in sorted(by_programme.items()):
        codes = sorted(codes)
        out.append(
            Issue(
                id="",
                category="no-timetabled-classes",
                severity="Critical",
                affects_validity=True,
                detection=DETECTED,
                course=None,
                groups=codes,
                session_ids=[],
                description=(
                    f"{len(codes)} {programme} group{'s' if len(codes) != 1 else ''} "
                    f"({'are' if len(codes) != 1 else 'is'} not in any session, so "
                    "they have no timetable at all."
                ),
                evidence={
                    "programme": programme,
                    "groups": codes,
                    "searched": "every session in the semester, by group link",
                },
                action=(
                    f"Timetable the {programme} groups or withdraw them from the "
                    "semester before it is published."
                ),
                students=sum(g.size for g in empty if g.programme == programme),
            )
        )
    return out


def _issue_capacity(dataset, analysis, cfg):
    """Over-capacity classes, one issue per room rather than per class.

    A room that is booked too tightly all week is one decision to take, and one
    line in a report the coordinator can act on. Listing forty identical rows
    would bury the other findings and hide how much of the estate is affected.
    """
    by_venue: dict = defaultdict(list)
    for row in analysis.observations:
        if row["kind"] == "overflow":
            by_venue[row["venue"]].append(row)
    out = []
    for venue, rows in sorted(
        by_venue.items(), key=lambda kv: -max(r["ratio"] for r in kv[1])
    ):
        worst = max(rows, key=lambda r: r["ratio"])
        over_seats = sum(r["students"] - r["capacity"] for r in rows)
        groups = sorted({g for r in rows for g in r["groups"]})
        courses = sorted({r["course"] for r in rows})
        hours = sum((r["end"] - r["start"]) / 60 for r in rows)
        out.append(
            Issue(
                id="",
                category="insufficient-venue-capacity",
                severity="Critical",
                affects_validity=True,
                detection=DETECTED,
                course=", ".join(courses[:3]) + ("..." if len(courses) > 3 else ""),
                groups=groups,
                session_ids=sorted({sid for r in rows for sid in r["session_ids"]}),
                description=(
                    f"{len(rows)} class(es) in {venue} are over capacity. "
                    f"The worst is {worst['students']} students in "
                    f"{worst['capacity']} seats on {DAY_NAMES[worst['day']]} "
                    f"{_hhmm(worst['start'])}-{_hhmm(worst['end'])} "
                    f"({worst['ratio'] * 100:.0f}% of seats)."
                ),
                evidence={
                    "venue": venue,
                    "classes_over_capacity": len(rows),
                    "worst": {
                        "course": worst["course"],
                        "activity": worst["activity"],
                        "day": DAY_NAMES[worst["day"]],
                        "time": f"{_hhmm(worst['start'])}-{_hhmm(worst['end'])}",
                        "students": worst["students"],
                        "capacity": worst["capacity"],
                        "seat_util": round(worst["ratio"], 2),
                        "shortfall": worst["students"] - worst["capacity"],
                        "groups": worst["groups"],
                    },
                    "seats_short_in_total": over_seats,
                    "courses": courses,
                    "worst_examples": [
                        {
                            "course": r["course"],
                            "activity": r["activity"],
                            "day": DAY_NAMES[r["day"]],
                            "time": f"{_hhmm(r['start'])}-{_hhmm(r['end'])}",
                            "students": r["students"],
                            "capacity": r["capacity"],
                        }
                        for r in sorted(rows, key=lambda r: -r["ratio"])[:3]
                    ],
                },
                action=(
                    f"Re-room the {len(rows)} class(es) in {venue} or add capacity; "
                    f"at peak {venue} is {over_seats} seats short across the week."
                ),
                students=sum(r["students"] for r in rows),
                hours=round(hours, 2),
            )
        )
    return out


def _issue_venue_info(dataset, analysis, cfg):
    """Sessions with no room, or a room whose capacity is 0.

    Both are reported as validity problems because the capacity check the
    allocator performs cannot be trusted for them — an unknown room is not a
    room with plenty of space.
    """
    groups: dict = defaultdict(list)
    no_venue, unknown_cap = set(), set()
    missing_ids: set = set()
    missing_courses: set = set()
    for p in physical_sessions(dataset.sessions).values():
        if not p["venue"]:
            no_venue.update(p["groups"])
            missing_ids.update(p["session_ids"])
            missing_courses.add(p["course"])
        elif not p["cap"]:
            unknown_cap.add(p["venue"])
    issues = []
    if no_venue:
        issues.append(
            Issue(
                id="",
                category="missing-venue-info",
                severity="Critical",
                affects_validity=True,
                detection=DETECTED,
                course=", ".join(sorted(missing_courses)[:3]) or None,
                groups=sorted(no_venue),
                session_ids=sorted(missing_ids),
                description=(
                    f"{len(no_venue)} group(s) attend session(s) with no room assigned, "
                    f"so capacity and room clashes cannot be checked."
                ),
                evidence={
                    "groups": len(no_venue),
                    "courses": sorted(missing_courses),
                    "reason": "venue is null",
                },
                action="Assign a room to every session; an unverified room is not a room with space.",
                students=len(no_venue) * 30,
            )
        )
    if unknown_cap:
        rooms = sorted(unknown_cap)
        issues.append(
            Issue(
                id="",
                category="missing-venue-info",
                severity="Review",
                affects_validity=True,
                detection=DETECTED,
                course=None,
                groups=[],
                session_ids=[],
                description=(
                    f"{len(rooms)} room(s) have a capacity of 0 or no capacity recorded, "
                    f"so their utilisation cannot be computed."
                ),
                evidence={"venues": rooms, "reason": "capacity 0 / not recorded"},
                action="Set each room's real capacity so the capacity check can pass on evidence.",
            )
        )
    return issues


def _issue_conflicts(dataset, analysis, cfg):
    out = []
    for group in analysis.groups:
        if not group.tfi:
            continue
        overlaps = group.tfi.get("overlaps") or []
        if not overlaps:
            continue
        ids = sorted({i for pair in overlaps for i in pair})
        detail = []
        for a, b in overlaps:
            first = next((s for s in group.sessions if s.id == a), None)
            second = next((s for s in group.sessions if s.id == b), None)
            if first and second:
                detail.append(
                    f"{first.course} {DAY_NAMES[first.day]} "
                    f"{_hhmm(first.start)}-{_hhmm(first.end)} overlaps "
                    f"{second.course} {_hhmm(second.start)}-{_hhmm(second.end)}"
                )
        out.append(
            Issue(
                id="",
                category="timetable-conflict",
                severity="Critical",
                affects_validity=True,
                detection=DETECTED,
                course=None,
                groups=[group.code],
                session_ids=ids,
                description=(
                    f"{group.code} is timetabled into {len(overlaps)} overlapping pair(s): "
                    + "; ".join(detail[:2])
                    + ("." if len(detail) <= 2 else ", and more.")
                ),
                evidence={
                    "pairs": len(overlaps),
                    "detail": detail,
                    "overlap_minutes": _overlap_minutes(group.sessions),
                },
                action=(
                    "Move one of the pair to another session of the same course, or "
                    "unassign the group from the duplicate."
                ),
                students=group.size,
                hours=sum((s.minutes for s in group.sessions)) / 60,
            )
        )
    return out


def _overlap_minutes(sessions) -> int:
    total = 0
    for day in {s.day for s in sessions}:
        rows = sorted([s for s in sessions if s.day == day], key=lambda s: s.start)
        for i, first in enumerate(rows):
            for second in rows[i + 1 :]:
                overlap = min(first.end, second.end) - second.start
                if overlap <= 0:
                    break
                total += overlap
    return total


def _issue_invalid_times(dataset, analysis, cfg):
    start_bound = int(cfg.get_path("working_day.start", 480))
    end_bound = int(cfg.get_path("working_day.end", 1080))
    bad: dict = defaultdict(list)
    for sess in dataset.sessions:
        reasons = []
        if sess.end <= sess.start:
            reasons.append("end time is not after the start time")
        else:
            if sess.start < start_bound:
                reasons.append(f"starts before {_hhmm(start_bound)}")
            if sess.end > end_bound:
                reasons.append(f"ends after {_hhmm(end_bound)}")
        if reasons:
            bad[(sess.id, sess.group)].extend(reasons)
    if not bad:
        return []
    groups = sorted({g for (_, g) in bad})
    reasons_flat = sorted({r for rs in bad.values() for r in rs})
    return [
        Issue(
            id="",
            category="invalid-session-time",
            severity="Critical",
            affects_validity=True,
            detection=DETECTED,
            course=None,
            groups=groups,
            session_ids=sorted({sid for (sid, _) in bad}),
            description=(
                f"{len(bad)} timetabled session(s) fall outside the working window "
                f"{_hhmm(start_bound)}-{_hhmm(end_bound)} or have an impossible duration."
            ),
            evidence={"count": len(bad), "reasons": reasons_flat,
                      "window": f"{_hhmm(start_bound)}-{_hhmm(end_bound)}"},
            action="Correct the times in the master timetable, then re-import.",
            students=len(groups) * 30,
        )
    ]


def _issue_workshops(dataset, analysis, cfg):
    """Workshop rows that name a day but no resolvable time.

    Reported rather than assumed: the audit cannot confirm or deny a clash for a
    session whose window it does not know, and "unverifiable" is the honest
    verdict.
    """
    rows = getattr(dataset, "workshop_rows", None) or []
    unverifiable = [
        r
        for r in rows
        if r.get("day") and not (r.get("start") and r.get("end"))
    ]
    if not unverifiable:
        return []
    # A workshop's bare "C1" stands for every group's "C1" -- the collector has
    # already expanded it -- so the issue names groups the reader can look up
    # rather than a code that matches nothing in the register's own group list.
    groups = sorted(
        {g for r in unverifiable for g in (r.get("groups") or [r.get("group_code")]) if g}
    )
    students = sum(
        (dataset.group_size.get(g) or 0) for r in unverifiable for g in (r.get("groups") or [])
    )
    return [
        Issue(
            id="",
            category="unverifiable-workshop-time",
            severity="Review",
            affects_validity=True,
            detection=DETECTED,
            course="WORKSHOP",
            groups=[f"{g}" for g in groups],
            session_ids=[],
            description=(
                f"{len(unverifiable)} workshop allocation(s) name a day but no clock "
                f"time, so a clash with a timetabled session cannot be confirmed."
            ),
            evidence={
                "count": len(unverifiable),
                "days": sorted({r.get("day") for r in unverifiable if r.get("day")}),
                "periods": sorted({r.get("time_period") for r in unverifiable if r.get("time_period")}),
            },
            action="Enter the start and end times for these workshop allocations.",
            students=students,
        )
    ]


def _issue_unconfigured(dataset, analysis, cfg):
    """Courses in play with no configured activity requirements.

    Deliberately **not** a validity problem: the allocator did exactly what it
    was told, and only a person can say whether the course should have needed a
    seminar. It is the clearest example in the report of "needs human judgment".
    """
    entries = dataset.unconfigured_courses or []
    if not entries:
        return []
    groups = set()
    for entry in entries:
        for group in dataset.sessions_by_group:
            if group.split(" ")[0] in (entry.get("programmes") or []):
                groups.add(group)
    return [
        Issue(
            id="",
            category="unconfigured-requirements",
            severity="Review",
            affects_validity=False,
            detection=JUDGMENT,
            course=", ".join(e["code"] for e in entries[:4]),
            groups=sorted(groups)[:40],
            session_ids=[],
            description=(
                f"{len(entries)} course(s) in this timetable have no required-activity "
                f"configuration, so nothing was allocated for them."
            ),
            evidence={
                "courses": [e["code"] for e in entries],
                "names": {e["code"]: e.get("name", "") for e in entries},
                "groups": len(groups),
            },
            action=(
                "Confirm whether these courses need small-group sessions; if they do, "
                "set their required activities and re-run the allocator."
            ),
        )
    ]


def _issue_friendliness(analysis, cfg):
    """Quality, explicitly not validity: one issue per flagged group."""
    out = []
    gap_bad = int(cfg.gap_bad_min)
    for group in analysis.groups:
        if not group.tfi or not group.is_flagged:
            continue
        worst = group.worst_component
        tfi = group.tfi
        why = "threshold" if any(f["why"] == "threshold" for f in group.flags) else "outlier"
        out.append(
            Issue(
                id="",
                category="friendliness-concern",
                severity="Review" if tfi["grade"] in {"D", "E"} else "Watch",
                affects_validity=False,
                detection=DETECTED,
                course=None,
                groups=[group.code],
                session_ids=[],
                description=(
                    f"{group.code} scores {tfi['score']} (grade {tfi['grade']}); "
                    f"its weakest component is {worst[0]} ({worst[1]:.2f})"
                    + (
                        f" and its longest gap is {_fmt_dur(tfi['longest_gap'])}"
                        if tfi["longest_gap"] >= gap_bad
                        else ""
                    )
                    + "."
                ),
                evidence={
                    "tfi": tfi["score"],
                    "grade": tfi["grade"],
                    "parts": tfi["parts"],
                    "longest_gap": tfi["longest_gap"],
                    "busiest_day": tfi.get("busiest_day"),
                    "busiest_day_h": tfi["busiest_day_h"],
                    "dead_ratio": round(tfi["dead_ratio"], 3),
                    "why": why,
                    "flags": [f["note"] for f in group.flags],
                    "cohort_z": {
                        f["metric"]: round(f.get("z") or 0.0, 2) for f in group.flags
                    },
                },
                action=(
                    f"Check whether the {DAY_NAMES[tfi['busiest_day']]} "
                    f"{tfi['busiest_day_h']:.1f} h block can be moved."
                    if tfi.get("busiest_day") is not None
                    else "Review this group's week."
                ),
                students=group.size,
                hours=group.week_hours,
            )
        )
    return out


def _renumber(issues: list) -> None:
    """Assign ISS-001… in the order they will be displayed."""
    for index, issue in enumerate(issues, start=1):
        issue.id = f"ISS-{index:03d}"


def pareto(issues: list) -> list:
    """Issue counts per category, biggest first, with a cumulative share."""
    counts: dict = defaultdict(lambda: {"count": 0, "students": 0, "categories": set()})
    for issue in issues:
        entry = counts[issue.category]
        entry["count"] += 1
        entry["students"] += issue.students
        entry["categories"].add(issue.severity)
    rows = sorted(counts.values(), key=lambda e: -e["count"])
    total = sum(r["count"] for r in rows) or 1
    running = 0
    for row in rows:
        running += row["count"]
        row["cumulative_pct"] = round(100.0 * running / total, 1)
        row["labels"] = sorted(row["categories"], key=lambda s: SEVERITY_ORDER[s])
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# 4.9 verdict + KPIs
# ─────────────────────────────────────────────────────────────────────────────


def verdict(issues, completion_pct) -> tuple:
    """The one-line answer at the top of the report.

    Validity decides it: a single critical validity breach is "NOT READY" no
    matter how comfortable the timetable is. Convenience never blocks
    publication, it is reported separately.
    """
    if any(i.affects_validity and i.severity == "Critical" for i in issues):
        return "NOT READY", "red"
    if completion_pct < 100 or any(i.severity == "Review" for i in issues):
        return "READY WITH REVIEW", "amber"
    return "READY TO PUBLISH", "green"


def kpis(dataset: AuditDataset, analysis: Analysis) -> dict:
    venues = [v for v in analysis.venues.values() if v["seat_util_avg"] is not None]
    seat = mean([v["seat_util_avg"] for v in venues]) if venues else None
    time_util = (
        mean([v["time_util"] for v in venues if v["time_util"] is not None])
        if any(v["time_util"] is not None for v in venues)
        else None
    )
    return {
        "completion_pct": analysis.allocation["completion_pct"],
        "unresolved": analysis.allocation["overall"]["unresolved"],
        "requirements": analysis.allocation["overall"]["total"],
        "campus_tfi": analysis.rollup.get("campus_mean_tfi"),
        "campus_median_tfi": analysis.rollup.get("campus_median_tfi"),
        "groups_scored": len(analysis.groups),
        "groups_at_de": sum(1 for g in analysis.groups if g.grade_letter in {"D", "E"}),
        "venue_seat_util": round(100 * seat, 1) if seat is not None else None,
        "venue_time_util": round(100 * time_util, 1) if time_util is not None else None,
        "validity_violations": analysis.validity_summary()["count"],
        "critical_issues": len(analysis.critical_issues),
        "dead_time_pct": analysis.rollup.get("campus_dead_min"),
        "unconfigured_courses": len(dataset.unconfigured_courses or []),
    }


def grade_distribution(analysis: Analysis) -> list:
    counts = {k: 0 for k in "ABCDE"}
    for group in analysis.groups:
        letter = group.grade_letter
        if letter in counts:
            counts[letter] += 1
    return [{"grade": k, "count": v} for k, v in counts.items()]


def _hhmm(minutes) -> str:
    minutes = int(minutes or 0)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"
