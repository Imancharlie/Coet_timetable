"""Analyse the synthetic fixture and print what the engine found.

    python manage.py audit_analyse_fixture

Runs inside a throwaway test database, exactly like ``audit_fixture``, and
prints the register the report is built from. It is the fastest way to see the
engine working without waiting for a PDF.
"""

import json
import os

from django.core.management.base import BaseCommand
from django.test.utils import setup_databases, setup_test_environment, teardown_databases


class Command(BaseCommand):
    help = "Run the audit engine over the synthetic fixture and print the findings."

    def add_arguments(self, parser):
        parser.add_argument("--json", action="store_true", help="Print the summary as JSON.")
        parser.add_argument("--top", type=int, default=12, help="Rows per table (default 12).")
        parser.add_argument(
            "--explain",
            action="store_true",
            help="Print the evidence behind every non-friendliness issue.",
        )

    def handle(self, *args, **options):
        from audit.analytics import analyse
        from audit.collect import collect
        from audit.config import default_config
        from audit.fixtures import generate
        from audit.recommend import build as build_recommendations

        setup_test_environment()
        databases = setup_databases(verbosity=0, interactive=False)
        try:
            generate.build()
            from core.models import AllocationRun

            run = AllocationRun.objects.order_by("-created_at").first()
            cfg = default_config()
            dataset = collect(run, cfg)
            analysis = analyse(dataset)
            recs = build_recommendations(analysis, cfg)
            top = options["top"]

            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING("Collected"))
            self.stdout.write("-" * 60)
            self.stdout.write(
                f"  {len(dataset.groups)} groups in scope, {len(dataset.scored_groups)} "
                f"with sessions, {len(dataset.sessions)} timetable rows, "
                f"{len(dataset.requirements)} requirements"
            )
            self.stdout.write(f"  data hash: {dataset.data_hash()[:16]}...")

            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING("Verdict"))
            self.stdout.write("-" * 60)
            valid = analysis.validity_summary()
            self.stdout.write(f"  {analysis.verdict_label}")
            self.stdout.write(
                f"  Validity: {valid['count']} violation(s), "
                f"{valid['groups']} group(s), {valid['students']} students"
            )
            self.stdout.write(
                f"  Friendliness: mean TFI {analysis.kpis['campus_tfi']} - "
                f"{analysis.kpis['groups_at_de']} group(s) at grade D/E"
            )
            self.stdout.write(
                f"  Completion {analysis.allocation['completion_pct']}% - "
                f"convenience-adjusted {analysis.allocation['convenience_adjusted_completion']}%"
            )

            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING(f"Groups ranked by TFI (top {top})"))
            self.stdout.write("-" * 60)
            self.stdout.write(
                f"  {'group':<9}{'prog':<5}{'TFI':>6}{'gr':>4}{'wk h':>7}{'busy':>7}"
                f"{'gap':>7}{'run':>6}{'dead%':>7}  flags"
            )
            for group in analysis.groups_by_tfi[:top]:
                tfi = group.tfi
                self.stdout.write(
                    f"  {group.code:<9}{group.programme:<5}{tfi['score']:>6.1f}"
                    f"{group.grade_letter:>4}{group.week_hours:>7.1f}"
                    f"{tfi['busiest_day_h']:>7.1f}{_dur(tfi['longest_gap']):>7}"
                    f"{_dur(tfi['longest_run_min']):>6}"
                    f"{tfi['dead_ratio'] * 100:>6.0f}%  "
                    + (", ".join(f["metric"] for f in group.flags) or "-")
                )

            self.stdout.write("")
            self.stdout.write(
                self.style.MIGRATE_HEADING(f"Issue register ({len(analysis.issues)})")
            )
            self.stdout.write("-" * 60)
            for issue in analysis.issues[: top * 2]:
                self.stdout.write(
                    f"  {issue.id}  {issue.severity:<8} {issue.category_label:<38} "
                    f"{issue.affects_validity and 'validity' or 'quality':<8} "
                    f"{issue.groups[:3] if issue.groups else '-'}"
                )
            if options["explain"]:
                for issue in analysis.issues:
                    if issue.category != "friendliness-concern":
                        self.stdout.write(
                            f"  {issue.id} {issue.category} "
                            f"{issue.groups[:4]} {issue.description}"
                        )

            self.stdout.write("")
            self.stdout.write(
                self.style.MIGRATE_HEADING(f"Recommendations ({len(recs)})")
            )
            self.stdout.write("-" * 60)
            for rec in recs.cards:
                self.stdout.write(
                    f"  {rec.rank:>2}. [{rec.priority:<8}] {rec.title}"
                )
                self.stdout.write(f"      {rec.headline}")
                self.stdout.write(f"      do: {_wrap(rec.action)}")
                if options["explain"]:
                    for item in rec.evidence[:4]:
                        self.stdout.write(
                            f"        - {item.get('label', '')}: {item.get('detail', '')}"
                        )
            for rec in recs.tables:
                self.stdout.write(
                    f"  {rec.rank:>2}. [{rec.priority:<8}] {rec.title} "
                    f"({len(rec.alternatives)} row(s))"
                )
                self.stdout.write(f"      {rec.headline}")
            if recs.uncovered_issues:
                # Not a failure: a friendliness concern has no single action, it
                # is a ranking. But the reader should be told what the
                # recommendations do *not* speak to.
                self.stdout.write("")
                self.stdout.write(
                    f"  {len(recs.uncovered_issues)} finding(s) have no recommendation: "
                    f"{sorted({i.category for i in recs.uncovered_issues})}"
                )

            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING("Day pressure (student-hours)"))
            self.stdout.write("-" * 60)
            for row in analysis.rollup["day_pressure"]:
                bar = "#" * int(row["hours"] / 40)
                self.stdout.write(f"  {row['label']:<10}{row['hours']:>9.1f}  {bar}")

            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING("Venues (lowest seat utilisation)"))
            self.stdout.write("-" * 60)
            usable = [
                v
                for v in analysis.venues.values()
                if v["seat_util_avg"] is not None
            ]
            for venue in sorted(usable, key=lambda v: v["seat_util_avg"])[:top]:
                self.stdout.write(
                    f"  {venue['venue']:<12} cap {venue['capacity']:>4}  "
                    f"seat {venue['seat_util_avg'] * 100:>5.0f}%  "
                    f"max {venue['seat_util_max'] * 100:>5.0f}%  "
                    f"time {venue['time_util'] * 100:>5.0f}%  "
                    f"{venue['sessions']:>3} session(s)"
                )
            over = [v for v in usable if v["overflow_sessions"]]
            self.stdout.write("")
            self.stdout.write(
                f"  {len(over)} venue(s) with at least one over-capacity session: "
                + (", ".join(f"{v['venue']}x{v['overflow_sessions']}" for v in over) or "none")
            )

            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING("Planted problems"))
            self.stdout.write("-" * 60)
            for label, found in _planted(analysis):
                mark = self.style.SUCCESS("FOUND  ") if found else self.style.ERROR("MISSED ")
                self.stdout.write(f"  {mark}{label}")

            if options["json"]:
                self.stdout.write("")
                self.stdout.write(
                    json.dumps(
                        {
                            "verdict": analysis.verdict_label,
                            "kpis": analysis.kpis,
                            "issues": [i.as_dict() for i in analysis.issues],
                        },
                        indent=2,
                        default=str,
                    )
                )
        finally:
            teardown_databases(databases, verbosity=0)


def _dur(minutes) -> str:
    minutes = int(round(minutes or 0))
    hours, mins = divmod(abs(minutes), 60)
    if hours and mins:
        return f"{hours}h{mins:02d}"
    if hours:
        return f"{hours}h"
    return f"{mins}m"


def _wrap(text: str, width: int = 86, indent: str = "        ") -> str:
    """Fold an action onto indented lines, so a long one stays readable."""
    import textwrap

    lines = textwrap.wrap(text, width=width, break_on_hyphens=False) or [""]
    return f"\n{indent}".join(lines)


def _planted(analysis) -> list:
    """Check the six planted problems were found, so a regression is obvious.

    Returns ``[(label, found), ...]``. A missing plant means the engine stopped
    seeing a whole class of problem, which is worse than a wrong number: the
    numbers would still look plausible on the page.
    """
    def has(category, must_contain=()):
        for issue in analysis.issues:
            if issue.category != category:
                continue
            groups = set(issue.groups or [])
            if all(g in groups for g in must_contain):
                return True
        return False

    def group_result(code):
        return analysis.group_by_code.get(code)

    over = any(
        row.get("venue") == "SEM3"
        and {"CH C1", "CH C2"} <= set(row.get("groups") or [])
        and row.get("kind") == "overflow"
        for row in analysis.observations
    )
    me = group_result("ME C3")
    gap_ok = bool(me and me.tfi and me.tfi.get("longest_gap", 0) >= 240)
    day_pressure = analysis.rollup["day_pressure"]
    busiest = analysis.rollup.get("heaviest_day")
    unsized = [
        v for v in analysis.venues.values() if v["has_venue"] and not v["capacity"]
    ]
    return [
        (
            "over-capacity room (CH C1 + CH C2 in SEM3)",
            over or has("insufficient-venue-capacity", ("CH C1", "CH C2")),
        ),
        (
            "unresolved requirement (MT455 seminar)",
            has("unresolved-requirement", ("CE C1", "ST C1")),
        ),
        (
            "session with no venue (SC121)",
            has("missing-venue-info", ("ST C3", "ST C4")),
        ),
        (
            "group with no classes at all (NM C8-C10)",
            has("no-timetabled-classes", ("NM C8", "NM C9", "NM C10")),
        ),
        ("room with unrecorded capacity (UNSIZED)", bool(unsized)),
        ("overlapping pair (EE C1)", has("timetable-conflict", ("EE C1",))),
        ("4-hour gap (ME C3)", gap_ok),
        (
            "Wednesday is the busiest day",
            # ``day`` is the weekday *index* (0 = Monday), so match the label.
            # Comparing against "WEDNESDAY" can never be true, which is a check
            # that silently fails rather than one that reports a real problem.
            bool(busiest) and busiest.get("label") == "Wednesday",
        ),
    ]
