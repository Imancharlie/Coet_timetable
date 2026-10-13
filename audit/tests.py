"""Tests for the read-only audit engine.

Two kinds of test:

* :class:`MetricTests` are pure functions over hand-built timetables. They pin
  the *definitions* -- what a gap is, what dead time is, where the grade
  boundaries fall -- because a metric nobody can pin down is a number that
  cannot be defended in front of a committee.
* :class:`FixtureTests` builds the whole synthetic faculty and runs the real
  collector and analytics over it. That is where the golden numbers, the
  planted problems and the read-only guarantee live.

The world is rebuilt per test rather than cached on the class. It costs about a
second, and the alternative is worse than slow: a class-level fixture built
outside a test's transaction is exactly the kind of shared state that makes a
suite order-dependent, and the whole point of these tests is that they fail for
one reason.
"""

from django.test import SimpleTestCase, TestCase

from audit.collect import collect
from audit.config import default_config
from audit.metrics import grade, score_group

#: The severities the register is allowed to use.
SEVERITIES = {"Critical", "Review", "Watch", "Info"}


def sess(group, day, start, end, course="EE101", activity="tutorial",
         venue="ROOM-X", capacity=60, size=30):
    """One session row, for the metric tests."""
    from audit.collect import Session

    return Session(
        id=f"{group}-{course}-{activity}-{day}-{start}",
        course=course,
        activity=activity,
        group=group,
        day=day,
        start=start,
        end=end,
        venue=venue,
        capacity=capacity,
        group_size=size,
    )


class MetricTests(SimpleTestCase):
    """The TFI definitions, on timetables small enough to reason about."""

    def setUp(self):
        self.cfg = default_config()

    def test_no_sessions_scores_nothing_rather_than_a_zero(self):
        self.assertIsNone(
            score_group([], self.cfg)["tfi"],
            "an empty timetable has no score to give, and 0 would look like a bad one",
        )

    def test_a_back_to_back_week_scores_better_than_a_sparse_one(self):
        """The comparison that matters, stated without magic numbers.

        Both weeks carry the same ten contact hours. One is five clean
        two-hour blocks; the other is ten one-hour blocks with a three-hour hole
        between each pair. The holes have to cost something, because the holes
        are what a student actually feels.
        """
        packed = score_group(
            [sess("EE C1", day, 8 * 60, 10 * 60) for day in range(5)], self.cfg
        )["tfi"]
        sparse = score_group(
            [row for day in range(5) for row in (
                sess("EE C1", day, 8 * 60, 9 * 60),
                sess("EE C1", day, 12 * 60, 13 * 60),
            )],
            self.cfg,
        )["tfi"]
        self.assertEqual(packed["week_hours"], sparse["week_hours"], "same teaching")
        self.assertEqual(packed["dead_min"], 0)
        self.assertEqual(sparse["dead_min"], 15 * 60, "five days, three hours each")
        self.assertGreater(packed["score"], sparse["score"])

    def test_dead_time_is_the_hole_between_classes(self):
        rows = [
            sess("EE C1", 0, 8 * 60, 9 * 60),
            sess("EE C1", 0, 14 * 60, 15 * 60),
        ]
        tfi = score_group(rows, self.cfg)["tfi"]
        self.assertEqual(tfi["longest_gap"], 5 * 60, "09:00 to 14:00 is a five-hour hole")
        self.assertEqual(tfi["dead_min"], 5 * 60)
        # dead_ratio is dead minutes *per contact minute*, not a share of the
        # day, so it can exceed 1 -- five hours of teaching after a five-hour
        # hole is a ratio of 2.5, which is the point.
        self.assertEqual(tfi["dead_ratio"], 2.5)

    def test_a_long_stretch_is_penalised_for_continuity(self):
        """One ten-hour day is not a good week, however full it is."""
        marathon = score_group(
            [sess("EE C1", 0, 8 * 60 + i * 60, 9 * 60 + i * 60) for i in range(10)],
            self.cfg,
        )["tfi"]
        self.assertEqual(marathon["longest_run_min"], 10 * 60)
        self.assertLess(
            marathon["parts"]["continuity"], 1.0,
            "a ten-hour run must cost something, or the metric rewards endurance",
        )

    def test_overlaps_are_reported_rather_than_merged(self):
        rows = [
            sess("EE C1", 0, 8 * 60, 10 * 60, course="EE101"),
            sess("EE C1", 0, 9 * 60, 11 * 60, course="EE102"),
        ]
        tfi = score_group(rows, self.cfg)["tfi"]
        self.assertTrue(tfi["overlaps"], "two classes at once must surface")
        self.assertEqual(len(tfi["overlaps"]), 1)

    def test_a_free_day_is_counted_not_scored(self):
        rows = [sess("EE C1", 0, 8 * 60, 9 * 60), sess("EE C1", 3, 8 * 60, 9 * 60)]
        result = score_group(rows, self.cfg)
        self.assertEqual(result["days"][0].hours, 1.0)
        self.assertEqual(result["days"][3].hours, 1.0)
        self.assertEqual(result["days"][1].hours, 0.0, "Tuesday has no teaching")
        self.assertEqual(result["tfi"]["free_days"], 3)
        self.assertEqual(result["tfi"]["days_used"], 2)

    def test_the_grade_floors_are_respected(self):
        floors = self.cfg.get("grades")
        for letter, floor in sorted(floors.items(), key=lambda kv: -float(kv[1])):
            self.assertEqual(grade(float(floor), self.cfg), letter)
            if letter != "E":
                self.assertEqual(grade(float(floor) - 0.05, self.cfg) != letter, True)
        self.assertEqual(grade(0.0, self.cfg), "E")

    def test_the_score_stays_inside_its_own_scale(self):
        for rows in (
            [],
            [sess("EE C1", 0, 6 * 60, 7 * 60)],               # before the day starts
            [sess("EE C1", 0, 6 * 60, 7 * 60), sess("EE C1", 0, 6 * 60, 7 * 60)],
        ):
            tfi = score_group(rows, self.cfg)["tfi"]
            if tfi is None:
                continue
            self.assertGreaterEqual(tfi["score"], 0.0)
            self.assertLessEqual(tfi["score"], 100.0)
            self.assertIn(tfi["grade"], {"A", "B", "C", "D", "E"})


class FixtureTests(TestCase):
    """One synthetic faculty, analysed for real.

    ``EXPECTED`` is the report's headline. Changing a number there changes what
    the consultancy tells the university, so it has to be a deliberate edit --
    not something a re-run quietly rewrites.
    """

    #: Re-baselined 2026-09-30, when the generator stopped manufacturing problems
    #: it never meant to create. The room search gave up after a fixed number of
    #: probes and handed back an already-booked room, so a cell needing most of
    #: the estate put two classes in one room; the series allocator also filled
    #: each *programme's* bands independently, which stacked every programme's
    #: first class of the day into one period. Both are fixed, and the estate
    #: now genuinely holds the week: 20 classes in a 23-room period at worst.
    #: The visible effect is a healthier week -- fewer accidental over-capacity
    #: sessions (23 validity violations down to 16) and a better TFI (51.6 to
    #: 61.4) -- because what is left is what the plants actually planted.
    EXPECTED = {
        "groups_in_scope": 80,
        "groups_with_sessions": 77,
        "timetable_rows": 1096,
        "requirements": 706,
        "hash_prefix": "c2256cf75f684d65",
        "validity_violations": 16,
        "campus_tfi": 61.4,
        "groups_at_de": 25,
        "completion_pct": 97.2,
    }

    #: The categories the fixture is *allowed* to trip. Anything else means the
    #: generator has started manufacturing a finding nobody chose.
    EXPECTED_VALIDITY_CATEGORIES = {
        "insufficient-venue-capacity",   # the rooms that cannot hold their class
        "unresolved-requirement",        # MT455, which nobody timetabled
        "missing-venue-info",            # SC121 with no room, UNSIZED
        "timetable-conflict",            # the EE C1 overlapping pair
        "no-timetabled-classes",         # NM C8-C10, left with nothing
        "unverifiable-workshop-time",    # a workshop with a day but no clock
    }

    @classmethod
    def setUpTestData(cls):
        from audit.analytics import analyse
        from audit.fixtures import generate
        from core.models import AllocationRun

        generate.build()
        run = AllocationRun.objects.order_by("-created_at").first()
        cls.allocation_run = run
        cls.dataset = collect(run, default_config())
        cls.analysis = analyse(cls.dataset)

    def issues_with(self, category, groups=()):
        for issue in self.analysis.issues:
            if issue.category == category and set(groups) <= set(issue.groups or []):
                yield issue

    # ── the world is what the fixture says it is ────────────────────────────

    def test_collection_covers_every_group_not_just_the_busy_ones(self):
        self.assertEqual(len(self.dataset.groups), self.EXPECTED["groups_in_scope"])
        self.assertEqual(
            len(self.dataset.scored_groups), self.EXPECTED["groups_with_sessions"]
        )

    def test_the_data_hash_is_stable(self):
        """The lineage hash, so a report can name the inputs it was built from."""
        self.assertTrue(
            self.dataset.data_hash().startswith(self.EXPECTED["hash_prefix"]),
            f"the data hash moved: {self.dataset.data_hash()}",
        )

    def test_collecting_the_same_run_twice_gives_the_same_data(self):
        again = collect(self.allocation_run, default_config())
        self.assertEqual(again.data_hash(), self.dataset.data_hash())
        self.assertEqual(len(again.sessions), len(self.dataset.sessions))
        self.assertEqual(len(again.requirements), len(self.dataset.requirements))

    # ── every planted problem is found ──────────────────────────────────────

    def test_every_planted_problem_shows_up(self):
        def found(category, groups=()):
            return bool(list(self.issues_with(category, groups)))

        me = self.analysis.group_by_code["ME C3"].tfi or {}
        seen = {
            "over-capacity room": found("insufficient-venue-capacity", ("CH C1", "CH C2")),
            "unresolved requirement": found("unresolved-requirement", ("CE C1", "ST C1")),
            "session with no venue": found("missing-venue-info", ("ST C3", "ST C4")),
            "group with no classes": found(
                "no-timetabled-classes", ("NM C8", "NM C9", "NM C10")
            ),
            "overlapping pair": found("timetable-conflict", ("EE C1",)),
            "four-hour gap": me.get("longest_gap", 0) >= 240,
            "wednesday is the busiest day":
                (self.analysis.rollup["heaviest_day"] or {}).get("label") == "Wednesday",
            "unrecorded capacity": any(
                v["has_venue"] and not v["capacity"] for v in self.analysis.venues.values()
            ),
        }
        missing = sorted(k for k, ok in seen.items() if not ok)
        self.assertEqual(missing, [], f"the engine stopped seeing: {missing}")

    def test_no_unintended_validity_problems(self):
        """The fixture plants problems; it does not also *cause* them.

        A category outside the expected set means the generator has started
        manufacturing a finding nobody chose. This is the failure the module
        exists to catch: every number would still look plausible on the page,
        so nothing else would notice.
        """
        unexpected = {
            i.category for i in self.analysis.issues if i.affects_validity
        } - self.EXPECTED_VALIDITY_CATEGORIES
        self.assertEqual(
            unexpected, set(), f"the fixture invented validity problems: {sorted(unexpected)}"
        )

    def test_no_room_is_double_booked(self):
        """Two classes, one room, one time is a venue clash nobody planted.

        Keyed on ``session_pk``, not on the row id: a whole-cohort lecture is
        one class in one hall that ``collect`` expands into a row per attending
        group, so counting rows would report every lecture as a dozen classes
        sharing a room.
        """
        from collections import defaultdict

        taken = defaultdict(set)
        for session in self.dataset.sessions:
            if session.venue:
                taken[(session.venue, session.day, session.start)].add(session.session_pk)
        clashes = {
            key: sorted(pks)
            for key, pks in taken.items()
            if len(pks) > 1
        }
        self.assertEqual(clashes, {}, f"rooms double-booked: {clashes}")

    def test_every_lecture_is_in_a_hall_that_seats_it(self):
        """A whole cohort in a seminar room is a right-sizing finding, not a fit."""
        too_small = []
        for session in self.dataset.sessions:
            if session.activity != "lecture" or not session.capacity:
                continue
            students = (session.group_size or 0)
            if students > session.capacity:
                too_small.append((session.course, session.venue, students, session.capacity))
        self.assertEqual(too_small, [], f"lectures too big for their hall: {too_small[:4]}")

    def test_the_only_group_clash_is_the_planted_one(self):
        clashing = {
            code
            for code, group in self.analysis.group_by_code.items()
            if (group.tfi or {}).get("overlaps")
        }
        self.assertEqual(clashing, {"EE C1"})

    def test_every_session_sits_inside_the_working_day(self):
        day = self.analysis_cfg_day()
        outside = [
            s.id for s in self.dataset.sessions if s.start < day[0] or s.end > day[1]
        ]
        self.assertEqual(outside, [], f"sessions outside {day}: {outside[:4]}")

    def analysis_cfg_day(self):
        cfg = default_config().get("working_day") or {}
        return cfg.get("start", 8 * 60), cfg.get("end", 18 * 60)

    # ── the register is usable ──────────────────────────────────────────────

    def test_every_issue_carries_evidence_and_an_action(self):
        for issue in self.analysis.issues:
            with self.subTest(category=issue.category):
                self.assertTrue(issue.evidence, f"{issue.category} has no evidence")
                self.assertTrue(issue.action.strip(), f"{issue.category} has no action")
                self.assertTrue(issue.detection, f"{issue.category} does not say how")
                self.assertTrue(issue.category_label)

    def test_issue_ids_are_sequential_and_unique(self):
        ids = [i.id for i in self.analysis.issues]
        self.assertEqual(len(ids), len(set(ids)), "two issues share an id")
        self.assertEqual(ids, [f"ISS-{n:03d}" for n in range(1, len(ids) + 1)])

    def test_every_group_named_in_an_issue_really_exists(self):
        """A workshop names a bare "C1"; the report must still name a real group."""
        in_scope = set(self.dataset.groups)
        for issue in self.analysis.issues:
            for code in issue.groups or []:
                self.assertIn(
                    code, in_scope,
                    f"{issue.id} ({issue.category}) names {code!r}, "
                    "which is not a group in scope",
                )

    def test_severities_are_known(self):
        for issue in self.analysis.issues:
            self.assertIn(issue.severity, SEVERITIES, issue.category)

    def test_the_workload_numbers_add_up(self):
        """The headline counts must be the register's own, not a second tally."""
        validity = [
            i for i in self.analysis.issues if i.affects_validity
        ]
        self.assertEqual(
            self.analysis.kpis["validity_violations"], len(validity),
            "the KPI count and the register disagree",
        )
        self.assertEqual(
            self.analysis.kpis["groups_scored"], len(self.analysis.groups)
        )
        self.assertEqual(
            self.analysis.kpis["requirements"], len(self.dataset.requirements)
        )

    # ── the numbers the report leads with ───────────────────────────────────

    def test_the_headline_numbers_have_not_moved(self):
        kpis = self.analysis.kpis
        self.assertEqual(kpis["validity_violations"], self.EXPECTED["validity_violations"])
        self.assertEqual(kpis["campus_tfi"], self.EXPECTED["campus_tfi"])
        self.assertEqual(kpis["groups_at_de"], self.EXPECTED["groups_at_de"])
        self.assertEqual(kpis["completion_pct"], self.EXPECTED["completion_pct"])
        self.assertEqual(self.analysis.verdict_label, "NOT READY")

    def test_the_timetable_metric_is_not_saturated(self):
        """A friendliness score only ranks groups if it spreads them out.

        If every group scored the same, the league table would be decoration --
        and the report would be recommending things on the strength of noise.
        """
        scored = [g for g in self.analysis.groups if g.tfi]
        scores = [g.tfi["score"] for g in scored]
        self.assertGreater(max(scores) - min(scores), 25.0, "no spread to rank on")
        unflagged = [g for g in scored if not g.flags]
        self.assertTrue(unflagged, "every group is flagged, so nothing stands out")
        worst_ten = sorted(scored, key=lambda g: g.tfi["score"])[:10]
        self.assertGreater(
            len({g.programme for g in worst_ten}), 1,
            "the worst groups are all one programme, so the ranking is an artefact",
        )
        self.assertLessEqual(
            sum(1 for g in scored if g.grade_letter in {"D", "E"}), len(scored) * 0.6
        )

    # ── the guarantee the whole app rests on ────────────────────────────────

    def test_analysis_does_not_touch_the_database(self):
        """The audit is read-only, and this is the test that says so.

        It checksums every table the audit reads and the allocation tables it
        must never write, then runs a full collect-and-analyse pass. A report is
        usually run straight after an allocation, and a report that quietly
        edited what it measured would be worse than no report.
        """
        from audit.analytics import analyse
        from core.models import (
            AllocationChange,
            AllocationRun,
            Course,
            Session,
            SessionGroup,
            StudentGroup,
            TechnicalDrawingAllocation,
            Venue,
            WorkshopAllocation,
        )

        models = [
            Session, SessionGroup, WorkshopAllocation, TechnicalDrawingAllocation,
            AllocationRun, AllocationChange, Course, Venue, StudentGroup,
        ]

        def fingerprint():
            parts = []
            for model in models:
                rows = list(
                    model.objects.order_by("pk").values(
                        *[f.name for f in model._meta.concrete_fields]
                    )
                )
                parts.append(f"{model.__name__}:{len(rows)}:{repr(rows)!r}")
            return "|".join(parts)

        before = fingerprint()
        again = analyse(collect(self.allocation_run, default_config()))
        self.assertEqual(fingerprint(), before, "the analytics wrote to the database")
        self.assertEqual(
            len(again.issues), len(self.analysis.issues),
            "and produced a different answer from the same input",
        )

    # ── recommendations ────────────────────────────────────────────────────

    def setUpRecommendations(self):
        """Built once per test that needs them; cheap, but not free."""
        from audit import recommend

        if getattr(type(self), "_recs", None) is None:
            type(self)._recs = recommend.build(self.analysis, self.analysis.config)
        return type(self)._recs

    def test_every_recommendation_carries_its_evidence(self):
        """The rule the whole module exists to keep.

        A recommendation is a claim about a real faculty. If it cannot name the
        records behind it, it is an opinion, and a consultancy report full of
        opinions is worth nothing.
        """
        recs = self.setUpRecommendations()
        self.assertTrue(len(recs) >= 6, f"only {len(recs)} recommendations")
        for rec in recs:
            with self.subTest(rec=rec.key):
                self.assertTrue(rec.evidence, f"{rec.key} recommends without evidence")
                self.assertTrue(rec.headline and rec.action, f"{rec.key} is not actionable")
                self.assertTrue(
                    any(str(item.get("detail", "")).strip() for item in rec.evidence),
                    f"{rec.key} has evidence rows with no detail",
                )

    def test_recommendations_are_ranked_and_unique(self):
        recs = self.setUpRecommendations()
        keys = [r.key for r in recs]
        self.assertEqual(len(keys), len(set(keys)), "duplicate recommendation keys")
        self.assertEqual([r.rank for r in recs], list(range(1, len(recs) + 1)))
        priorities = [r.priority for r in recs]
        self.assertEqual(priorities, sorted(priorities, key=recommend_order), "out of rank order")

    def test_the_planted_problems_all_produce_a_recommendation(self):
        """Each planted problem must be actionable, not merely detected.

        The search covers the evidence as well as the prose: naming *which*
        eleven classes and *which* group is the part that makes a
        recommendation usable, and it is in the evidence rows.
        """
        recs = self.setUpRecommendations()
        text = " ".join(
            f"{r.key} {r.headline} {r.action} "
            + " ".join(
                f"{e.get('label', '')} {e.get('detail', '')} {' '.join(e.get('groups') or [])}"
                for e in r.evidence
            )
            for r in recs
        ).lower()
        for expected in ("sem3", "mt455", "sc121", "nm c8", "ee c1", "wednesday"):
            with self.subTest(expected=expected):
                self.assertIn(expected, text, f"no recommendation mentions {expected}")

    def test_alternatives_are_sessions_the_group_could_actually_attend(self):
        """A suggestion the coordinator cannot act on is worse than none.

        Every alternative has to be a class the group studies, at a time it is
        free, in a room big enough -- checked here against the raw dataset
        rather than against the search's own bookkeeping.
        """
        recs = self.setUpRecommendations()
        tables = [r for r in recs if r.presentation == "table"]
        self.assertTrue(tables, "no per-group table was produced")
        sessions = {s.session_pk: s for s in self.dataset.sessions if s.session_pk}
        limit = int(self.analysis.config.get_path("flags.alternatives_per_group", 3))
        seen = 0
        for rec in tables:
            for row in rec.alternatives:
                group = self.analysis.group_by_code[row["group"]]
                eligible = {r.course for r in self.dataset.requirements if r.group == group.code}
                attended = {s.session_pk for s in group.sessions}
                self.assertLessEqual(len(row["options"]), limit)
                for alt in row["options"]:
                    seen += 1
                    session = sessions.get(alt["session_pk"])
                    with self.subTest(group=group.code, course=alt["course"]):
                        self.assertIsNotNone(session)
                        self.assertIn(session.course, eligible, "not a course the group studies")
                        self.assertNotIn(session.session_pk, attended, "already attends it")
                        self.assertTrue(alt["why"], "no reason given")
                        if session.capacity:
                            self.assertLessEqual(
                                session.group_size or 0, session.capacity,
                                "the alternative is itself over capacity",
                            )
                        for other in group.sessions:
                            self.assertFalse(
                                other.day == session.day
                                and other.start < session.end
                                and session.start < other.end,
                                "the alternative clashes with the group's own timetable",
                            )
        self.assertTrue(seen, "not one alternative was offered")

    def test_a_group_with_no_timetable_gets_no_alternatives(self):
        """An empty timetable has nothing to clash with, so it is not a gap."""
        from audit import recommend

        for group in self.analysis.unscored_groups:
            self.assertEqual(
                recommend.alternatives_for_group(
                    self.analysis.group_by_code[group], self.analysis, self.dataset,
                    self.analysis.config,
                ),
                [],
                f"{group} has no timetable, so it cannot be offered a session to move to",
            )


def recommend_order(priority):
    from audit.recommend import PRIORITIES

    return PRIORITIES.index(priority)


# ── the report the coordinator actually reaches for ─────────────────────────
#
# Everything above proves the analysis is *right*. These prove it can be
# reached, read and exported — the parts a consultant notices first, and the
# only ones that fail silently if a URL is never reversed.
#
# These classes do NOT inherit FixtureTests. They would each re-run its 26
# analysis tests (three times over, 80s of duplicated work) to get hold of one
# built run, so the fixture is shared through this mixin instead: the test
# database is rolled back per test but the rows built here are the common
# starting point every one of them wants.


class BuiltRunMixin:
    """A generated world and one already-audited allocation run."""

    @classmethod
    def setUpTestData(cls):
        from audit.fixtures import generate
        from audit.generator import generate_report
        from core.models import AllocationRun

        generate.build()
        run = AllocationRun.objects.order_by("-created_at").first()
        cls.allocation_run = run
        cls.report = generate_report(run, write_pdf=False)


class ReportGenerationTests(BuiltRunMixin, TestCase):
    """A report row is produced, versioned, and carries its own payload."""

    def test_generating_a_report_stores_a_version_with_a_payload(self):
        from audit.generator import generate_report
        from audit.models import AuditReport

        report = generate_report(self.allocation_run, write_pdf=False)
        self.assertEqual(report.version, 2, "the second report of a run is v2")
        self.assertEqual(report.status, "COMPLETE")
        self.assertEqual(report.progress, "Complete (HTML only)")

        payload = report.summary()
        for key in (
            "verdict", "kpis", "recommendations", "issues", "groups",
            "day_pressure", "venues", "config_yaml",
        ):
            self.assertIn(key, payload, f"the payload has no {key!r}")
        self.assertTrue(payload["recommendations"], "no recommendations stored")
        self.assertTrue(
            payload["data_hash"],
            "the report did not record which data it was built from",
        )

    def test_the_report_says_it_built_on_the_same_data(self):
        """Two reports of unchanged data must agree on their lineage."""
        from audit.collect import collect
        from audit.config import default_config

        payload = self.report.summary()
        self.assertEqual(
            payload["data_hash"],
            collect(self.allocation_run, default_config()).data_hash(),
            "the stored report describes different data from the one collectable",
        )

    def test_re_auditing_adds_a_version_and_never_overwrites(self):
        from audit.generator import generate_report
        from audit.models import AuditReport

        first = generate_report(self.allocation_run, write_pdf=False)
        second = generate_report(self.allocation_run, write_pdf=False)

        self.assertEqual((first.version, second.version), (2, 3))
        self.assertEqual(
            AuditReport.objects.filter(allocation_run=self.allocation_run).count(), 3
        )
        first.refresh_from_db()
        self.assertEqual(first.status, "COMPLETE", "v2 was overwritten by v3")

    def test_a_failed_generation_is_recorded_rather_than_raised(self):
        """A coordinator pressing a button must see a reason, not a 500."""
        from unittest import mock

        from audit.generator import generate_report

        with mock.patch(
            "audit.generator.collect",
            side_effect=RuntimeError("the database went away"),
        ):
            report = generate_report(self.allocation_run, write_pdf=False)

        self.assertEqual(report.status, "FAILED")
        self.assertIn("the database went away", report.error)
        self.assertEqual(
            report.progress, "Failed", "a failed report still claims to be running"
        )


class ReportViewTests(BuiltRunMixin, TestCase):
    """The four access points: generate, read, list, export."""

    @classmethod
    def setUpTestData(cls):
        super().setUpTestData()
        from audit.generator import generate_report
        from audit.models import AuditReport

        cls.report = generate_report(cls.allocation_run, write_pdf=False)
        cls.runs_before = AuditReport.objects.count()

    def setUp(self):
        from django.contrib.auth import get_user_model
        from django.test import Client

        User = get_user_model()
        self.user = User.objects.create_user(
            username="auditor", password="pw", is_staff=True
        )
        self.client = Client()
        self.client.force_login(self.user)

    def test_the_list_page_names_every_report(self):
        response = self.client.get("/audit-reports/")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Audit Reports")
        self.assertContains(response, f"Version {self.report.version}")
        self.assertContains(response, "PDF", msg_prefix="no export offered")

    def test_the_detail_page_renders_the_stored_payload(self):
        response = self.client.get(f"/audit-reports/{self.report.pk}/")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, self.report.summary()["verdict"]["label"])
        # Every recommendation the analysis produced is on the page, with the
        # numbers that justify it.
        payload = self.report.summary()
        for rec in payload["recommendations"]:
            self.assertContains(response, rec["title"])

    def test_the_detail_page_reads_the_stored_report_not_a_fresh_analysis(self):
        """A version must not change under the reader.

        If the page re-analysed, deleting the underlying sessions would silently
        change what an already-issued report says — which is the one thing a
        signed-off report must never do.
        """
        from core.models import Session

        Session.objects.all().delete()
        response = self.client.get(f"/audit-reports/{self.report.pk}/")
        self.assertEqual(response.status_code, 200)
        self.assertContains(
            response, self.report.summary()["recommendations"][0]["title"]
        )

    def test_generating_from_the_page_writes_exactly_one_report(self):
        response = self.client.post(f"/allocation-runs/{self.allocation_run.pk}/audit/")
        self.assertEqual(response.status_code, 302)
        from audit.models import AuditReport

        self.assertEqual(
            AuditReport.objects.count(),
            self.runs_before + 1,
            "generating a report wrote something other than one report row",
        )
        fresh = AuditReport.objects.order_by("-pk").first()
        self.assertIn(f"/audit-reports/{fresh.pk}/", response["Location"])

    def test_generating_is_post_only(self):
        """It writes a row and prints a PDF; a GET must not be able to do that."""
        response = self.client.get(f"/allocation-runs/{self.allocation_run.pk}/audit/")
        self.assertEqual(response.status_code, 405)
        from audit.models import AuditReport

        self.assertEqual(AuditReport.objects.count(), self.runs_before)

    def test_a_report_stored_in_an_old_payload_format_is_re_rendered(self):
        """An unreadable page is worse than one that is a version behind.

        When the payload shape changes, every report already written is
        unrenderable by the new templates. The view regenerates into a new
        version and redirects rather than raising, and keeps the old row.
        """
        from audit.models import AuditReport
        from audit.render import PAYLOAD_VERSION

        import json

        stale = AuditReport.objects.get(pk=self.report.pk)
        stale.summary_json = json.dumps(
            {
                "verdict": {"label": "READY", "colour": "green"},
                "recommendations": [
                    {"rank": 1, "title": "Old shape", "evidence": [{"note": "x"}]}
                ],
            }
        )
        stale.save(update_fields=["summary_json"])

        response = self.client.get(f"/audit-reports/{stale.pk}/")
        self.assertEqual(response.status_code, 302, "an old payload rendered or errored")
        fresh_pk = int(response["Location"].split("/audit-reports/")[1].split("/")[0])
        self.assertNotEqual(fresh_pk, stale.pk, "it redirected to the same version")
        self.assertEqual(
            AuditReport.objects.get(pk=fresh_pk).summary()["payload_version"],
            PAYLOAD_VERSION,
        )
        stale.refresh_from_db()
        self.assertIn("Old shape", stale.summary_json, "the old version was overwritten")
        # And the redirect target actually renders.
        self.assertEqual(self.client.get(f"/audit-reports/{fresh_pk}/").status_code, 200)

    def test_the_pages_require_a_login(self):
        self.client.logout()
        for url in (
            "/audit-reports/",
            f"/audit-reports/{self.report.pk}/",
            f"/audit-reports/{self.report.pk}/json/",
        ):
            with self.subTest(url=url):
                self.assertEqual(self.client.get(url).status_code, 302)


class SidebarAuditLinkTests(BuiltRunMixin, TestCase):
    """The report has to be findable from the app it is about."""

    def test_the_sidebar_offers_audit_reports_under_allocation(self):
        from django.template.loader import render_to_string

        html = render_to_string("partials/sidebar.html", {"nav": ""})
        self.assertIn('href="/audit-reports/"', html)
        self.assertIn("Audit Reports", html)
        # Below Group Progress, inside the Allocation section: the coordinator
        # looks for it after the allocator, not before.
        self.assertLess(
            html.index("Group Progress"), html.index('href="/audit-reports/"')
        )

    def test_the_allocation_page_offers_the_audit_of_an_applied_run(self):
        from django.template.loader import render_to_string

        from audit.generator import generate_report

        report = generate_report(self.allocation_run, write_pdf=False)
        self.allocation_run.status = "APPLIED"
        self.allocation_run.save(update_fields=["status"])

        from core.views import _allocation_base_context

        request = self._request()
        ctx = _allocation_base_context(request, semester=self.allocation_run.semester)
        html = render_to_string("core/allocation.html", ctx)
        self.assertIn(f"/audit-reports/{report.pk}/", html)
        self.assertIn(f"/allocation-runs/{self.allocation_run.pk}/audit/", html)

    def test_audit_urls_highlight_the_sidebar_link(self):
        from core.templatetags.core_tags import _NAV_SECTIONS

        for name in (
            "audit-reports",
            "audit-report-detail",
            "audit-report-pdf",
        ):
            with self.subTest(name=name):
                self.assertEqual(_NAV_SECTIONS.get(name), "audit-reports")

    def _request(self):
        from django.test import RequestFactory

        request = RequestFactory().get("/allocation/")
        request.user = type(
            "U", (), {"is_staff": True, "is_authenticated": True, "pk": 1}
        )()
        return request
