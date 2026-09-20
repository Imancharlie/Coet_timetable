import os
import tempfile
from datetime import time
from pathlib import Path

import pandas as pd
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import Client, TestCase
from openpyxl import Workbook

from core.importers import (
    import_master_timetable_from_excel,
    import_programme_courses_from_excel,
    import_td_allocation_from_excel,
    import_venues_from_excel,
    import_workshop_allocation_from_excel,
    reconcile_master_timetable,
    reconcile_workshop_workbook,
)
from core.models import (
    ActivityLog,
    LogAction,
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
from core.venue_quality import analyse_venues, base_key, issues_for, suggested_name
from core.workshop_parser import parse_workbook

XLSX_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

MASTER_COLS = [
    "course_code",
    "activity_type",
    "day",
    "start_time",
    "end_time",
    "venue",
    "group",
]

_TMP_DIR = tempfile.mkdtemp()
_file_counter = {"n": 0}


def make_xlsx(rows, columns):
    """Write rows to a temp xlsx and return its path."""
    path = os.path.join(_TMP_DIR, "data_%d.xlsx" % _file_counter["n"])
    _file_counter["n"] += 1
    pd.DataFrame(rows, columns=columns).to_excel(path, index=False)
    return path


def _col(letter):
    from openpyxl.utils import column_index_from_string

    return column_index_from_string(letter)


_MATRIX_CATEGORIES = [
    ("A", "B", "Workshop"),
    ("C", "E", "Bench"),
    ("F", "J", "Building"),
    ("K", "O", "Carpentry"),
    ("P", "T", "CPE"),
    ("U", "Z", "Electrical"),
    ("AA", "AF", "Electronics"),
    ("AG", "AK", "M/Tools"),
    ("AL", "AM", "Plumbing"),
    ("AN", "AS", "Welding"),
]

_MATRIX_POSITIONS = {
    "C": 1, "D": 4, "E": 6,
    "F": 1, "G": 2, "H": 3, "I": 4, "J": 6,
    "K": 1, "L": 2, "M": 3, "N": 4, "O": 6,
    "P": 1, "Q": 2, "R": 3, "S": 4, "T": 6,
    "U": 1, "V": 2, "W": 3, "X": 4, "Y": 5, "Z": 6,
    "AA": 1, "AB": 2, "AC": 3, "AD": 4, "AE": 5, "AF": 6,
    "AG": 1, "AH": 3, "AI": 4, "AJ": 5, "AK": 6,
    "AL": 1, "AM": 5,
    "AN": 1, "AO": 2, "AP": 3, "AQ": 4, "AR": 5, "AS": 6,
}

_MATRIX_KEY = [
    (1, "Monday Morning and Wednesday Morning"),
    (2, "Tuesday Morning and Friday Morning"),
    (3, "Wednesday Afternoon and Thursday Morning"),
    (4, "Monday Afternoon"),
    (5, "Tuesday Morning and Thursday Afternoon"),
    (6, "Friday Afternoon"),
]


def make_workshop_matrix(path, schedule1=None, schedule2=None):
    """Build a representative copy of the raw university workshop workbook."""
    s1 = schedule1 or {"C": "E1", "G": "A2", "W": "C1", "AA": "C2", "AG": "D1"}
    s2 = schedule2 or {"C": "E2", "F": "C2", "M": "C1", "P": "B3"}
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"] = "FIRST YEAR - 2025/2026"
    ws["A2"] = "WORKSHOP TRAINING"
    ws["A3"] = "SEMESTER 1 : WORKSHOP SCHEDULES (Week 1-15): 24th November, 2025: Version 0"
    ws["A6"] = "Workshops and Attending Groups"

    for start, end, name in _MATRIX_CATEGORIES:
        ws.merge_cells(f"{start}7:{end}7")
        ws[f"{start}7"] = name

    ws.merge_cells("B8:B10")
    ws["B8"] = "WEEK No"

    ws["A9"] = "GROUPS"
    for start, end, name in _MATRIX_CATEGORIES[1:]:
        ws.merge_cells(f"{start}9:{end}9")
        ws[f"{start}9"] = name

    ws["A10"] = "POSITION"
    for letter, pos in _MATRIX_POSITIONS.items():
        ws[f"{letter}10"] = pos

    ws.merge_cells("A11:A18")
    ws["A11"] = "SCHEDULE 1"
    for i, week in enumerate(range(1, 8)):
        ws.cell(row=11 + i, column=2, value=week)
    for letter, code in s1.items():
        for i in range(7):
            ws.cell(row=11 + i, column=_col(letter), value=code)

    ws.merge_cells("A20:A27")
    ws["A20"] = "SCHEDULE 2"
    for i, week in enumerate(range(8, 15)):
        ws.cell(row=20 + i, column=2, value=week)
    for letter, code in s2.items():
        for i in range(7):
            ws.cell(row=20 + i, column=_col(letter), value=code)

    ws["Q18"] = 4
    ws["R18"] = 6

    ws.merge_cells("A30:A34")
    ws["A30"] = "KEY"
    for i, (num, text) in enumerate(_MATRIX_KEY):
        row = 30 + i
        if i < 3:
            ws.cell(row=row, column=2, value=num)
            ws.cell(row=row, column=3, value=text)
        else:
            ws.cell(row=row, column=20, value=num)
            ws.cell(row=row, column=21, value=text)

    wb.save(path)
    return path


class ImporterTestCase(TestCase):
    """Shared helpers for Excel-backed importer tests."""

    def _seed(self):
        self.sem1 = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.sem2 = Semester.objects.create(academic_year="2026/2027", semester=2)
        self.prog_a = Programme.objects.create(code="CE", name="Civil Engineering")
        self.prog_b = Programme.objects.create(code="ME", name="Mechanical Engineering")
        self.g1 = StudentGroup.objects.create(programme=self.prog_a, code="A1")
        self.g2 = StudentGroup.objects.create(programme=self.prog_a, code="A2")
        self.g3 = StudentGroup.objects.create(programme=self.prog_b, code="B1")
        ProgrammeCourse.objects.create(
            programme=self.prog_a, course_code="MT161", course_name="Mathematics 1", semester=1
        )
        ProgrammeCourse.objects.create(
            programme=self.prog_b, course_code="MT161", course_name="Mathematics 1", semester=1
        )
        ProgrammeCourse.objects.create(
            programme=self.prog_a, course_code="TG201", course_name="Technical Drawing 1", semester=1
        )
        Venue.objects.create(name="LH1", capacity=80)
        Venue.objects.create(name="NB102", capacity=40)


class ImportReconciliationTests(ImporterTestCase):
    def test_reconcile_is_read_only(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        result = reconcile_master_timetable(path, semester_id=self.sem1.pk)
        self.assertEqual(result.created, 0)
        self.assertEqual(Session.objects.count(), 0)
        self.assertEqual(Venue.objects.count(), 2)
        self.assertIn("MT161", result.matched_courses)
        self.assertEqual(result.missing_courses, [])

    def test_reconcile_reports_missing_reference_data(self):
        self._seed()
        path = make_xlsx(
            [
                ["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "NO_SUCH_VENUE", "A1"],
                ["GHOST100", "LECTURE", "MONDAY", "09:00", "11:00", "LH1", "A9"],
            ],
            MASTER_COLS,
        )
        result = reconcile_master_timetable(path, semester_id=self.sem1.pk)
        self.assertIn("GHOST100", result.missing_courses)
        self.assertIn("A9", result.missing_groups)
        self.assertIn("NO_SUCH_VENUE", result.missing_venues)
        self.assertTrue(result.errors)

    def test_reconcile_blocks_when_semester_missing(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        result = reconcile_master_timetable(path, semester_id=9999)
        self.assertTrue(result.errors)
        self.assertIn("blocked", result.errors[0])

def test_reconcile_requires_all_columns(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "LH1"]],
            ["course_code", "activity_type", "day", "start_time", "venue"],
        )
        result = reconcile_master_timetable(path, semester_id=self.sem1.pk)
        self.assertTrue(result.errors)
        self.assertIn("Missing required columns", result.errors[0])


class MasterTimetableImportTests(ImporterTestCase):
    def test_import_is_idempotent(self):
        self._seed()
        path = make_xlsx(
            [
                ["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"],
                ["TG201", "TUTORIAL", "TUESDAY", "11:00", "12:00", "LH1", "B1"],
                ["MT161", "PRACTICAL", "WEDNESDAY", "14:00", "16:00", "NB102", "A2"],
            ],
            MASTER_COLS,
        )
        first = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(first.created, 3)
        self.assertEqual(Session.objects.count(), 3)

        second = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(second.created, 0)
        self.assertEqual(second.updated, 3)
        self.assertEqual(Session.objects.count(), 3)
        self.assertEqual(SessionGroup.objects.count(), 3)

    def test_exact_course_codes_preserved(self):
        self._seed()
        path = make_xlsx(
            [["mt161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertTrue(Session.objects.filter(course_code="mt161").exists())
        self.assertFalse(Session.objects.filter(course_code="MT161").exists())

    def test_all_group_expands_through_programme_courses(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "ALL"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        session = Session.objects.get(course_code="MT161", semester=self.sem1)
        self.assertEqual(session.session_groups.count(), 3)  # CE A1,A2 + ME B1
        self.assertEqual(len(result.all_rows), 1)
        self.assertEqual(result.conflicts, [])

    def test_all_group_unresolvable_is_reported_not_dropped(self):
        self._seed()
        path = make_xlsx(
            [["GHOST100", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "ALL"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertTrue(result.conflicts)
        self.assertIn("GHOST100", result.missing_courses)
        session = Session.objects.get(course_code="GHOST100", semester=self.sem1)
        self.assertEqual(session.session_groups.count(), 0)

    def test_missing_groups_report_and_venue_autocreate(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "NO_SUCH_VENUE", "A9"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertIn("NO_SUCH_VENUE", result.missing_venues)
        self.assertIn("A9", result.missing_groups)
        self.assertTrue(result.errors)
        session = Session.objects.get(course_code="MT161", semester=self.sem1)
        self.assertEqual(session.session_groups.count(), 0)
        self.assertEqual(Venue.objects.get(name="NO_SUCH_VENUE").capacity, 0)

    def test_dry_run_writes_nothing(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(
            path, semester_id=self.sem1.pk, dry_run=True
        )
        self.assertEqual(result.created + result.updated, 1)
        self.assertEqual(Session.objects.count(), 0)
        self.assertEqual(Venue.objects.count(), 2)

    def test_import_blocked_when_semester_missing(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=9999)
        self.assertTrue(result.errors)
        self.assertEqual(Session.objects.count(), 0)


class WorkshopTdImportTests(ImporterTestCase):
    def test_workshop_import_is_idempotent(self):
        self._seed()
        path = make_xlsx(
            [["TG201", "C1", "MONDAY", "08:00", "10:00", "TW101"]],
            ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
        )
        first = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(first.created, 1)
        second = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(second.created, 0)
        self.assertEqual(second.updated, 1)
        self.assertEqual(WorkshopAllocation.objects.count(), 1)

    def test_workshop_format_b_derives_course_code_from_venue(self):
        self._seed()
        path = make_xlsx(
            [["E1", "MONDAY", "09:00", "13:00", "Electrical"]],
            ["group_code", "day", "start_time", "end_time", "venue"],
        )
        first = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(first.errors, [])
        self.assertEqual(first.created, 1)
        wa = WorkshopAllocation.objects.get()
        self.assertEqual(wa.course_code, "Electrical")
        self.assertEqual(wa.venue, "Electrical")
        self.assertEqual(wa.group_code, "E1")

        second = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(second.created, 0)
        self.assertEqual(second.updated, 1)
        self.assertEqual(WorkshopAllocation.objects.count(), 1)

    def test_workshop_format_a_blank_course_code_falls_back_to_venue(self):
        self._seed()
        path = make_xlsx(
            [["", "E1", "MONDAY", "09:00", "13:00", "Electrical"]],
            ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
        )
        first = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(first.created, 1)
        wa = WorkshopAllocation.objects.get()
        self.assertEqual(wa.course_code, "Electrical")
        self.assertEqual(wa.venue, "Electrical")

    def test_workshop_invalid_format_reports_clear_error(self):
        self._seed()
        path = make_xlsx(
            [["E1", "MONDAY", "Electrical"]],
            ["group_code", "day", "venue"],
        )
        result = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertTrue(result.errors)
        self.assertIn("Invalid workshop format", result.errors[0])
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_td_import_is_idempotent(self):
        self._seed()
        path = make_xlsx(
            [["TG201", "A1", "MONDAY", "08:00", "10:00", "TW101"]],
            ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
        )
        first = import_td_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(first.created, 1)
        second = import_td_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(second.updated, 1)
        self.assertEqual(TechnicalDrawingAllocation.objects.count(), 1)


class AdaptiveImportTests(ImporterTestCase):
    """Importers accept readable aliases and reshape data into the required model."""

    def test_programme_courses_auto_create_programmes_from_names(self):
        self._seed()
        path = make_xlsx(
            [
                ["BSc. in Chemical and Processing Engineering", "CPE100", "Process Units", 1],
                ["BSc. in Textile Design and Technology", "TDT101", "Weaving", 2],
            ],
            ["Program", "Course Code", "Course", "Semester"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertEqual(result.created, 2)
        self.assertEqual(
            sorted(Programme.objects.filter(code__iexact="CPE").values_list("code", flat=True)),
            ["CPE"],
        )
        self.assertTrue(
            ProgrammeCourse.objects.filter(
                course_code="TDT101", course_name="Weaving", semester=2
            ).exists()
        )
        self.assertEqual(len(result.programmes_created), 2)

    def test_programme_courses_take_program_code_column(self):
        self._seed()
        result = import_programme_courses_from_excel(
            make_xlsx(
                [["ME", "ST101", "Strength of Materials", 1]],
                ["Program Code", "Course Code", "Course", "Semester"],
            )
        )
        self.assertEqual(result.created, 1)
        self.assertEqual(result.programmes_created, [])
        pc = ProgrammeCourse.objects.get(course_code="ST101")
        self.assertEqual(pc.programme, self.prog_b)
        self.assertEqual(pc.semester, 1)

    def test_programme_courses_resolve_existing_code_or_name(self):
        self._seed()
        result = import_programme_courses_from_excel(
            make_xlsx(
                [["ME", "ST101", "Strength of Materials", 1]],
                ["Program", "Course Code", "Course", "Semester"],
            )
        )
        self.assertEqual(result.created, 1)
        self.assertEqual(result.programmes_created, [])

        result = import_programme_courses_from_excel(
            make_xlsx(
                [["Civil Engineering", "CV101", "Intro to Surveying", 1]],
                ["Program", "Course Code", "Course", "Semester"],
            )
        )
        self.assertEqual(result.created, 1)
        self.assertEqual(result.programmes_created, [])
        self.assertTrue(
            ProgrammeCourse.objects.filter(
                programme=self.prog_a, course_code="CV101"
            ).exists()
        )

    def test_venue_alias_headers_accepted(self):
        self._seed()
        result = import_venues_from_excel(
            make_xlsx(
                [["Hall A", 120], ["Lab 1", 30]],
                ["Room", "Seats"],
            )
        )
        self.assertEqual(result.created, 2)
        self.assertTrue(Venue.objects.filter(name="Hall A", capacity=120).exists())
        self.assertTrue(Venue.objects.filter(name="Lab 1", capacity=30).exists())

    def test_master_import_reads_alias_headers(self):
        self._seed()
        result = import_master_timetable_from_excel(
            make_xlsx(
                [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
                ["Course", "Type", "Day", "Start", "End", "Room", "Groups"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertEqual(result.created, 1)
        self.assertEqual(Session.objects.get(course_code="MT161").course_code, "MT161")

    def test_master_import_splits_comma_separated_course_codes(self):
        self._seed()
        result = import_master_timetable_from_excel(
            make_xlsx(
                [["MT161, TG201", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
                ["course_code", "activity_type", "day", "start_time", "end_time", "venue", "group"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertEqual(result.created, 2)
        self.assertTrue(Session.objects.filter(course_code="MT161").exists())
        self.assertTrue(Session.objects.filter(course_code="TG201").exists())

    def test_master_import_auto_creates_default_semester(self):
        Semester.objects.all().delete()
        result = import_master_timetable_from_excel(
            make_xlsx(
                [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
                MASTER_COLS,
            )
        )
        self.assertTrue(Semester.objects.exists())
        self.assertTrue(result.detected_semester)
        self.assertEqual(Session.objects.count(), 1)

    def test_workshop_flat_accepts_time_range_column(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "C1", "MONDAY", "08:00-10:00", "TW101"]],
                ["course_code", "group_code", "day", "time", "venue"],
            )
        )
        self.assertEqual(result.created, 1)
        wa = WorkshopAllocation.objects.get(course_code="TG201", group_code="C1")
        self.assertEqual(wa.start_time, time(8, 0))
        self.assertEqual(wa.end_time, time(10, 0))

    def test_td_pivoted_layout_requires_course_code(self):
        self._seed()
        result = import_td_allocation_from_excel(
            make_xlsx(
                [["A1", "09:00-12:00"]],
                ["Group", "Time"],
            )
        )
        self.assertIn("course_code", " ".join(result.errors))
        self.assertEqual(TechnicalDrawingAllocation.objects.count(), 0)

    def test_td_pivoted_layout_with_course_code(self):
        self._seed()
        result = import_td_allocation_from_excel(
            make_xlsx(
                [
                    ["TG201", "MONDAY", "A1", "09:00-12:00", "S112"],
                    ["TG201", "", "A2", "13:00-15:00", ""],
                ],
                ["course_code", "Day", "Group", "Time", "Venue"],
            )
        )
        self.assertEqual(result.created, 2)
        rec = TechnicalDrawingAllocation.objects.get(group_code="A1")
        self.assertEqual(rec.venue, "S112")
        self.assertEqual(rec.start_time, time(9, 0))
        self.assertEqual(rec.end_time, time(12, 0))
        # Merged-cell day/time/venue are filled down from the first group row.
        rec2 = TechnicalDrawingAllocation.objects.get(group_code="A2")
        self.assertEqual(rec2.day, "MONDAY")
        self.assertEqual(rec2.venue, "S112")


class CrudResponseTests(TestCase):
    """Save() responses must work with and without htmx, with CSRF enforced."""

    def setUp(self):
        self.client = Client(enforce_csrf_checks=True)
        self.programme = Programme.objects.create(code="CE", name="Civil Engineering")

    def _post(self, url, data, hx=False):
        self.client.get("/")
        token = self.client.cookies.get("csrftoken").value
        headers = {}
        if hx:
            headers["HTTP_HX_REQUEST"] = "true"
        return self.client.post(
            url, {**data, "csrfmiddlewaretoken": token}, **headers
        )

    def test_create_htmx_returns_trigger(self):
        resp = self._post(
            "/programmes/create/", {"code": "ME", "name": "Mech Eng"}, hx=True
        )
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertTrue(Programme.objects.filter(code="ME").exists())

    def test_create_plain_post_redirects(self):
        resp = self._post(
            "/programmes/create/", {"code": "ME", "name": "Mech Eng"}
        )
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, "/programmes/")
        self.assertTrue(Programme.objects.filter(code="ME").exists())

    def test_duplicate_programme_shows_error(self):
        self._post("/programmes/create/", {"code": "ME", "name": "Mech Eng"}, hx=True)
        resp = self._post(
            "/programmes/create/", {"code": "ME", "name": "Mech Eng"}, hx=True
        )
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("HX-Trigger", resp.headers)
        self.assertIn("Programme with this Code already exists", resp.content.decode())
        self.assertEqual(Programme.objects.filter(code="ME").count(), 1)

    def test_edit_and_delete_plain_post(self):
        created = Programme.objects.create(code="ME", name="Mech Eng")
        resp = self._post(
            "/programmes/%d/edit/" % created.pk,
            {"code": "ME", "name": "Mechanical Engineering"},
        )
        self.assertEqual(resp.status_code, 302)
        created.refresh_from_db()
        self.assertEqual(created.name, "Mechanical Engineering")

        resp = self._post("/programmes/%d/delete/" % created.pk, {})
        self.assertEqual(resp.status_code, 302)
        self.assertFalse(Programme.objects.filter(pk=created.pk).exists())

    def test_edit_htmx_preserves_refresh_detail_trigger(self):
        created = Programme.objects.create(code="ME", name="Mech Eng")
        resp = self._post(
            "/programmes/%d/edit/" % created.pk,
            {"code": "ME", "name": "Mechanical Engineering"},
            hx=True,
        )
        self.assertEqual(
            resp.headers["HX-Trigger"], "close-modal,refresh-table,refresh-detail"
        )

    def test_session_group_add_remove_fallback(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        venue = Venue.objects.create(name="LH1", capacity=80)
        group = StudentGroup.objects.create(programme=self.programme, code="A1")
        session = Session.objects.create(
            semester=sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=venue,
        )
        resp = self._post("/sessions/%d/add-group/" % session.pk, {"group_id": group.pk})
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, "/sessions/%d/" % session.pk)
        self.assertTrue(
            SessionGroup.objects.filter(session=session, group=group).exists()
        )
        resp = self._post("/sessions/%d/remove-group/%d/" % (session.pk, group.pk), {})
        self.assertEqual(resp.status_code, 302)
        self.assertFalse(
            SessionGroup.objects.filter(session=session, group=group).exists()
        )

    def test_session_crud_via_ui_saves(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        venue = Venue.objects.create(name="LH1", capacity=80)
        data = {
            "semester": sem.pk,
            "course_code": "MT161",
            "activity_type": "LECTURE",
            "day": "MONDAY",
            "start_time": "08:00",
            "end_time": "10:00",
            "venue": venue.pk,
            "session_groups-TOTAL_FORMS": "0",
            "session_groups-INITIAL_FORMS": "0",
            "session_groups-MIN_NUM_FORMS": "0",
            "session_groups-MAX_NUM_FORMS": "1000",
        }
        resp = self._post("/sessions/create/", data, hx=True)
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(Session.objects.count(), 1)

    def test_all_entities_create_htmx(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        cases = [
            (
                "/groups/create/",
                {"programme": self.programme.pk, "code": "A1"},
                StudentGroup,
            ),
            ("/venues/create/", {"name": "NB102", "capacity": 40}, Venue),
            (
                "/semesters/create/",
                {"academic_year": "2027/2028", "semester": 1},
                Semester,
            ),
            (
                "/courses/create/",
                {
                    "programme": self.programme.pk,
                    "course_code": "MT161",
                    "course_name": "Mathematics 1",
                    "semester": "1",
                },
                ProgrammeCourse,
            ),
            (
                "/workshops/create/",
                {
                    "semester": sem.pk,
                    "course_code": "TG201",
                    "group_code": "C1",
                    "day": "MONDAY",
                    "start_time": "08:00",
                    "end_time": "10:00",
                    "venue": "TW101",
                },
                WorkshopAllocation,
            ),
            (
                "/td/create/",
                {
                    "semester": sem.pk,
                    "course_code": "TG201",
                    "group_code": "A1",
                    "day": "MONDAY",
                    "start_time": "08:00",
                    "end_time": "10:00",
                    "venue": "TW101",
                },
                TechnicalDrawingAllocation,
            ),
        ]
        for create_url, data, model in cases:
            with self.subTest(url=create_url):
                resp = self._post(create_url, data, hx=True)
                self.assertEqual(resp.status_code, 200)
                self.assertEqual(
                    resp.headers["HX-Trigger"], "close-modal,refresh-table"
                )
                self.assertGreater(model.objects.count(), 0)


class SidebarTests(TestCase):
    def _active_count(self, resp):
        return resp.content.decode().count("bg-slate-800 text-white")

    def test_active_nav_highlighted(self):
        resp = self.client.get("/programmes/")
        html = resp.content.decode()
        self.assertIn("Programmes", html)
        self.assertEqual(self._active_count(resp), 1)

    def test_dashboard_link_active_on_root(self):
        resp = self.client.get("/")
        self.assertEqual(self._active_count(resp), 1)


class ImportUploadViewTests(TestCase):
    def test_htmx_upload_returns_partial(self):
        self.client.get("/import/programmes/")
        path = make_xlsx(
            [["CE", "Civil Engineering"]], ["code", "name"]
        )
        with open(path, "rb") as fh:
            resp = self.client.post(
                "/import/programmes/",
                {
                    "file": SimpleUploadedFile(
                        "prog.xlsx", fh.read(), content_type=XLSX_CONTENT_TYPE
                    )
                },
                HTTP_HX_REQUEST="true",
            )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Created", resp.content.decode())
        self.assertTrue(Programme.objects.filter(code="CE").exists())

    def test_plain_upload_returns_full_page_with_result(self):
        self.client.get("/import/programmes/")
        path = make_xlsx(
            [["ME", "Mechanical Engineering"]], ["code", "name"]
        )
        with open(path, "rb") as fh:
            resp = self.client.post(
                "/import/programmes/",
                {
                    "file": SimpleUploadedFile(
                        "prog.xlsx", fh.read(), content_type=XLSX_CONTENT_TYPE
                    )
                },
            )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Import successful", resp.content.decode())
        self.assertIn("#upload-result", resp.content.decode())
        self.assertTrue(Programme.objects.filter(code="ME").exists())

    def test_workshop_matrix_upload_auto_detects(self):
        sem = Semester.objects.create(academic_year="2025/2026", semester=1)
        path = os.path.join(_TMP_DIR, "matrix_upload_%d.xlsx" % _file_counter["n"])
        _file_counter["n"] += 1
        make_workshop_matrix(path)
        self.client.get("/import/workshop-allocation/")
        with open(path, "rb") as fh:
            resp = self.client.post(
                "/import/workshop-allocation/",
                {
                    "file": SimpleUploadedFile(
                        "ws.xlsx", fh.read(), content_type=XLSX_CONTENT_TYPE
                    )
                },
                HTTP_HX_REQUEST="true",
            )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Created", html)
        self.assertIn("2025/2026 - Semester 1", html)
        self.assertIn("raw university workshop matrix", html)
        self.assertEqual(WorkshopAllocation.objects.count(), 18)
        self.assertTrue(
            WorkshopAllocation.objects.filter(
                semester=sem,
                course_code="Electrical",
                group_code="C1",
                day="WEDNESDAY",
                time_period="AFTERNOON",
                week_start=1,
                week_end=7,
            ).exists()
        )

    def test_master_timetable_upload_shows_reconciliation(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        prog = Programme.objects.create(code="CE", name="Civil Engineering")
        StudentGroup.objects.create(programme=prog, code="A1")
        ProgrammeCourse.objects.create(
            programme=prog, course_code="MT161", course_name="Mathematics 1", semester=1
        )
        Venue.objects.create(name="LH1", capacity=80)
        self.client.get("/import/master-timetable/")
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "ALL"]],
            MASTER_COLS,
        )
        with open(path, "rb") as fh:
            resp = self.client.post(
                "/import/master-timetable/",
                {
                    "file": SimpleUploadedFile(
                        "mt.xlsx", fh.read(), content_type=XLSX_CONTENT_TYPE
                    )
                },
                HTTP_HX_REQUEST="true",
            )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Reconciled", html)
        self.assertIn("ALL", html)
        self.assertEqual(Session.objects.count(), 1)
        session = Session.objects.first()
        self.assertEqual(session.session_groups.count(), 1)


class WorkshopMatrixParserTests(TestCase):
    """Pure parser tests against both the built fixture and the real workbook."""

    def _matrix_path(self):
        path = os.path.join(_TMP_DIR, "matrix_%d.xlsx" % _file_counter["n"])
        _file_counter["n"] += 1
        return make_workshop_matrix(path)

    def test_builder_fixture_parses(self):
        parsed = parse_workbook(self._matrix_path())
        self.assertEqual(parsed.academic_year, "2025/2026")
        self.assertEqual(parsed.semester, 1)
        self.assertEqual(parsed.year_of_study, 1)
        self.assertEqual(sorted(parsed.key), [1, 2, 3, 4, 5, 6])
        self.assertEqual(
            parsed.key[3],
            [("WEDNESDAY", "AFTERNOON"), ("THURSDAY", "MORNING")],
        )
        self.assertEqual(
            [(s.name, s.weeks) for s in parsed.sections],
            [
                ("SCHEDULE 1", [1, 2, 3, 4, 5, 6, 7]),
                ("SCHEDULE 2", [8, 9, 10, 11, 12, 13, 14]),
            ],
        )
        c1 = [r for r in parsed.records if r.group_code == "C1"]
        self.assertEqual(len(c1), 4)
        elect = [r for r in c1 if r.workshop == "Electrical"]
        self.assertTrue(all(r.position == 3 for r in elect))
        self.assertTrue(
            all(r.week_start == 1 and r.week_end == 7 for r in elect)
        )
        self.assertEqual(
            sorted((r.day, r.time_period) for r in elect),
            [("THURSDAY", "MORNING"), ("WEDNESDAY", "AFTERNOON")],
        )
        carp = [r for r in c1 if r.workshop == "Carpentry"]
        self.assertTrue(
            all(r.position == 3 and r.week_start == 8 and r.week_end == 14 for r in carp)
        )
        c2 = [r for r in parsed.records if r.group_code == "C2"]
        self.assertEqual(
            sorted(r.workshop for r in c2),
            ["Building", "Building", "Electronics", "Electronics"],
        )
        self.assertIn("Q18='4' (no week number)", parsed.unrecognized_cells)
        self.assertEqual(parsed.errors, [])
        self.assertEqual(parsed.invalid_groups, [])
        self.assertEqual(len(parsed.records), 18)

    def test_real_workbook_parses(self):
        real = Path(__file__).resolve().parents[1] / "workshop_timetable.xlsx"
        if not real.exists():
            self.skipTest("workshop_timetable.xlsx not present")
        parsed = parse_workbook(real)
        self.assertEqual(parsed.academic_year, "2025/2026")
        self.assertEqual(parsed.semester, 1)
        self.assertEqual(sorted(parsed.key), [1, 2, 3, 4, 5, 6])
        self.assertEqual(
            [(s.name, s.weeks) for s in parsed.sections],
            [
                ("SCHEDULE 1", [1, 2, 3, 4, 5, 6, 7]),
                ("SCHEDULE 2", [8, 9, 10, 11, 12, 13, 14]),
            ],
        )
        self.assertEqual(len(parsed.records), 98)
        self.assertEqual(len(parsed.unrecognized_cells), 7)
        self.assertEqual(parsed.errors, [])
        self.assertEqual(parsed.invalid_groups, [])
        self.assertEqual(parsed.unknown_keys, [])
        c1 = [r for r in parsed.records if r.group_code == "C1"]
        self.assertEqual(len(c1), 4)
        elect = [r for r in c1 if r.workshop == "Electrical"]
        self.assertEqual(len(elect), 2)
        self.assertTrue(all(r.position == 3 for r in elect))
        self.assertTrue(
            all(r.week_start == 1 and r.week_end == 7 for r in elect)
        )
        self.assertEqual(
            sorted((r.day, r.time_period) for r in elect),
            [("THURSDAY", "MORNING"), ("WEDNESDAY", "AFTERNOON")],
        )
        carp = [r for r in c1 if r.workshop == "Carpentry"]
        self.assertEqual(len(carp), 2)
        self.assertTrue(
            all(r.position == 3 and r.week_start == 8 and r.week_end == 14 for r in carp)
        )
        self.assertTrue(all(r.source_cells[0].startswith("M") for r in carp))
        c2 = [r for r in parsed.records if r.group_code == "C2"]
        self.assertEqual(len(c2), 4)
        self.assertTrue(all(r.position == 1 for r in c2))
        self.assertEqual(
            sorted(r.workshop for r in c2),
            ["Building", "Building", "Electronics", "Electronics"],
        )


class WorkshopMatrixImportTests(ImporterTestCase):
    def setUp(self):
        self._seed()
        self.sem = Semester.objects.create(academic_year="2025/2026", semester=1)
        for code in ["A2", "B3", "C1", "C2", "D1", "E1", "E2"]:
            StudentGroup.objects.get_or_create(programme=self.prog_a, code=code)

    def _matrix_path(self):
        path = os.path.join(_TMP_DIR, "matrix_imp_%d.xlsx" % _file_counter["n"])
        _file_counter["n"] += 1
        return make_workshop_matrix(path)

    def test_import_matrix_is_idempotent_and_auto_detected(self):
        path = self._matrix_path()
        first = import_workshop_allocation_from_excel(path)
        self.assertTrue(first.format.startswith("raw"))
        self.assertEqual(first.detected_semester, "2025/2026 - Semester 1")
        self.assertEqual(first.created, 18)
        self.assertEqual(WorkshopAllocation.objects.count(), 18)

        self.assertTrue(
            WorkshopAllocation.objects.filter(
                semester=self.sem,
                course_code="Electrical",
                group_code="C1",
                day="WEDNESDAY",
                time_period="AFTERNOON",
                position=3,
                schedule_section="SCHEDULE 1",
                week_start=1,
                week_end=7,
            ).exists()
        )
        self.assertTrue(
            WorkshopAllocation.objects.filter(
                semester=self.sem,
                course_code="Carpentry",
                group_code="C1",
                position=3,
                schedule_section="SCHEDULE 2",
                week_start=8,
                week_end=14,
            ).exists()
        )
        self.assertTrue(
            WorkshopAllocation.objects.filter(
                semester=self.sem,
                course_code="Building",
                group_code="C2",
                position=1,
                schedule_section="SCHEDULE 2",
                week_start=8,
                week_end=14,
            ).exists()
        )

        second = import_workshop_allocation_from_excel(path)
        self.assertEqual(second.created, 0)
        self.assertEqual(second.updated, 18)
        self.assertEqual(WorkshopAllocation.objects.count(), 18)

    def test_import_matrix_auto_creates_missing_semester(self):
        Semester.objects.filter(academic_year="2025/2026", semester=1).delete()
        result = import_workshop_allocation_from_excel(self._matrix_path())
        self.assertEqual(result.detected_semester, "2025/2026 - Semester 1")
        self.assertFalse(result.missing_references)
        self.assertEqual(result.created, 18)
        self.assertTrue(
            Semester.objects.filter(academic_year="2025/2026", semester=1).exists()
        )

    def test_matrix_dry_run_writes_nothing(self):
        result = import_workshop_allocation_from_excel(
            self._matrix_path(), dry_run=True
        )
        self.assertEqual(result.created + result.updated, 18)
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_reconcile_reports_missing_groups(self):
        StudentGroup.objects.all().delete()
        result = reconcile_workshop_workbook(self._matrix_path())
        self.assertFalse(result.missing_references)
        self.assertEqual(
            sorted(result.missing_groups),
            ["A2", "B3", "C1", "C2", "D1", "E1", "E2"],
        )
        self.assertEqual(result.created, 18)
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_reconcile_reports_missing_semester(self):
        Semester.objects.filter(academic_year="2025/2026", semester=1).delete()
        result = reconcile_workshop_workbook(self._matrix_path())
        self.assertTrue(result.missing_references)
        self.assertEqual(result.skipped, 18)


class SessionAssignLectureGroupsTests(TestCase):
    """Assign button: LECTURE sessions get every group of the programme(s)
    that study the course, regardless of subgroup."""

    def setUp(self):
        self._seed()

    def _seed(self):
        self.sem = Semester.objects.create(academic_year="2025/2026", semester=1)
        self.prog_a = Programme.objects.create(code="CE", name="Civil Engineering")
        self.prog_b = Programme.objects.create(code="ME", name="Mechanical Engineering")
        ProgrammeCourse.objects.create(
            programme=self.prog_a, course_code="MT161", course_name="Mathematics 1", semester=1
        )
        ProgrammeCourse.objects.create(
            programme=self.prog_b, course_code="MT161", course_name="Mathematics 1", semester=1
        )
        ProgrammeCourse.objects.create(
            programme=self.prog_a, course_code="TG201", course_name="Technical Drawing 1", semester=1
        )
        self.groups = {
            "A1": StudentGroup.objects.create(programme=self.prog_a, code="A1"),
            "A2": StudentGroup.objects.create(programme=self.prog_a, code="A2"),
            "B1": StudentGroup.objects.create(programme=self.prog_b, code="B1"),
            "B2": StudentGroup.objects.create(programme=self.prog_b, code="B2"),
        }
        venue = Venue.objects.create(name="LH1", capacity=0)
        self.lecture_mt161 = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=venue,
        )
        self.lecture_tg201 = Session.objects.create(
            semester=self.sem,
            course_code="TG201",
            activity_type="LECTURE",
            day="TUESDAY",
            start_time="08:00",
            end_time="10:00",
            venue=venue,
        )
        self.lecture_nomapping = Session.objects.create(
            semester=self.sem,
            course_code="PHY200",
            activity_type="LECTURE",
            day="WEDNESDAY",
            start_time="08:00",
            end_time="10:00",
            venue=venue,
        )
        self.workshop_mt161 = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="WORKSHOP",
            day="THURSDAY",
            start_time="08:00",
            end_time="10:00",
            venue=venue,
        )

    def test_assigns_all_programme_groups_to_lectures_only(self):
        page = self.client.get("/sessions/")
        self.assertContains(page, "Assign Lecture Groups")
        self.client.get("/sessions/")
        resp = self.client.post(
            "/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true"
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Lecture group assignment complete", html)
        self.assertIn("Group links created: 6", html)
        self.assertIn("PHY200", html)
        self.assertEqual(
            set(
                SessionGroup.objects.filter(
                    session=self.lecture_mt161
                ).values_list("group__code", flat=True)
            ),
            {"A1", "A2", "B1", "B2"},
        )
        self.assertEqual(
            set(
                SessionGroup.objects.filter(
                    session=self.lecture_tg201
                ).values_list("group__code", flat=True)
            ),
            {"A1", "A2"},
        )
        self.assertEqual(
            SessionGroup.objects.filter(session=self.lecture_nomapping).count(), 0
        )
        self.assertEqual(
            SessionGroup.objects.filter(session=self.workshop_mt161).count(), 0
        )

    def test_reassign_is_idempotent(self):
        self.client.post("/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true")
        resp = self.client.post(
            "/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true"
        )
        html = resp.content.decode()
        self.assertIn("Group links created: 0", html)
        self.assertEqual(SessionGroup.objects.filter(session=self.lecture_mt161).count(), 4)

    def test_get_redirects_to_list(self):
        resp = self.client.get("/sessions/assign-lecture-groups/")
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, "/sessions/")


class SessionProgrammeSummaryTests(TestCase):
    """Session detail shows programme-level allocation, All marking and cancel."""

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2025/2026", semester=1)
        self.prog_a = Programme.objects.create(code="CE", name="Civil Engineering")
        self.prog_b = Programme.objects.create(code="ME", name="Mechanical Engineering")
        self.g1 = StudentGroup.objects.create(programme=self.prog_a, code="A1")
        self.g2 = StudentGroup.objects.create(programme=self.prog_a, code="A2")
        self.g3 = StudentGroup.objects.create(programme=self.prog_b, code="B1")
        self.g4 = StudentGroup.objects.create(programme=self.prog_b, code="B2")
        venue = Venue.objects.create(name="LH1", capacity=80)
        self.session = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=venue,
        )

    def _assign(self, *groups):
        for g in groups:
            SessionGroup.objects.create(session=self.session, group=g)

    def test_detail_shows_programmes_at_top_with_all_marker(self):
        self._assign(self.g1, self.g2, self.g3)  # CE all, ME partial
        resp = self.client.get("/sessions/%d/" % self.session.pk)
        html = resp.content.decode()
        self.assertIn("Civil Engineering", html)
        self.assertIn("Mechanical Engineering", html)
        self.assertIn("All (2 groups)", html)
        self.assertIn("1 of 2 groups", html)
        self.assertIn("CE A1", html)
        self.assertIn("CE A2", html)
        self.assertIn("ME B1", html)
        # programmes/groups panel renders above the detail fields
        self.assertLess(
            html.find("Programmes &amp; Groups Attending"),
            html.find("Course Code"),
        )

    def test_htmx_detail_partial_shows_summary(self):
        self._assign(self.g1, self.g2, self.g3)
        resp = self.client.get(
            "/sessions/%d/" % self.session.pk, HTTP_HX_REQUEST="true"
        )
        html = resp.content.decode()
        self.assertIn("All (2 groups)", html)
        self.assertIn("1 of 2 groups", html)

    def test_remove_programme_groups_cancels_only_that_programme(self):
        self._assign(self.g1, self.g2, self.g3, self.g4)
        resp = self.client.post(
            "/sessions/%d/remove-programme-groups/%d/"
            % (self.session.pk, self.prog_b.pk),
            {},
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("Mechanical Engineering", resp.content.decode())
        remaining = set(
            SessionGroup.objects.filter(session=self.session).values_list(
                "group__code", flat=True
            )
        )
        self.assertEqual(remaining, {"A1", "A2"})

    def test_clear_groups_cancels_whole_assignment(self):
        self._assign(self.g1, self.g2, self.g3)
        resp = self.client.post(
            "/sessions/%d/clear-groups/" % self.session.pk,
            {},
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("No groups assigned", resp.content.decode())
        self.assertEqual(SessionGroup.objects.filter(session=self.session).count(), 0)

    def test_clear_groups_plain_post_redirects(self):
        self._assign(self.g1)
        resp = self.client.post("/sessions/%d/clear-groups/" % self.session.pk, {})
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, "/sessions/%d/" % self.session.pk)
        self.assertEqual(SessionGroup.objects.filter(session=self.session).count(), 0)


class ListFilterTests(TestCase):
    """Column filters on list views, especially the master timetable."""

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.venue_lh1 = Venue.objects.create(name="LH1", capacity=80)
        self.venue_nb = Venue.objects.create(name="NB102", capacity=40)
        self.s_mon = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=self.venue_lh1,
        )
        self.s_tue = Session.objects.create(
            semester=self.sem,
            course_code="TG201",
            activity_type="PRACTICAL",
            day="TUESDAY",
            start_time="14:00",
            end_time="17:00",
            venue=self.venue_nb,
        )

    def test_session_list_renders_filter_controls(self):
        resp = self.client.get("/sessions/")
        html = resp.content.decode()
        for control in (
            'name="day"',
            'name="activity_type"',
            'name="semester"',
            'name="venue"',
            'name="start_from"',
            'name="start_to"',
        ):
            self.assertIn(control, html)

    def test_session_day_filter(self):
        resp = self.client.get("/sessions/", {"day": "MONDAY"})
        html = resp.content.decode()
        self.assertContains(resp, "MT161")
        self.assertNotContains(resp, "TG201")

    def test_session_time_range_filter(self):
        resp = self.client.get(
            "/sessions/", {"start_from": "09:00", "start_to": "12:00"}
        )
        self.assertContains(resp, "MT161")
        self.assertNotContains(resp, "TG201")

    def test_session_venue_filter(self):
        resp = self.client.get("/sessions/", {"venue": "LH1"})
        self.assertContains(resp, "MT161")
        self.assertNotContains(resp, "TG201")
        self.assertContains(resp, "1 record")

    def test_course_semester_filter(self):
        ProgrammeCourse.objects.create(
            programme=Programme.objects.create(
                code="CE", name="Civil Engineering"
            ),
            course_code="MT161",
            course_name="Mathematics 1",
            semester=1,
        )
        ProgrammeCourse.objects.create(
            programme=Programme.objects.get(code="CE"),
            course_code="TG201",
            course_name="Technical Drawing 1",
            semester=2,
        )
        resp = self.client.get("/courses/", {"semester": "1"})
        self.assertContains(resp, "MT161")
        self.assertNotContains(resp, "TG201")

    def test_venue_capacity_range_filter(self):
        resp = self.client.get("/venues/", {"capacity_min": "50"})
        self.assertContains(resp, "LH1")
        self.assertNotContains(resp, "NB102")


class ActivityLogTests(TestCase):
    """Sidebar vlogs: changes logged, clickable from dashboard, cancellable."""

    def setUp(self):
        self.client = Client(enforce_csrf_checks=True)
        self.programme = Programme.objects.create(code="CE", name="Civil Engineering")

    def _post(self, url, data, hx=False):
        self.client.get("/")
        token = self.client.cookies.get("csrftoken").value
        headers = {}
        if hx:
            headers["HTTP_HX_REQUEST"] = "true"
        return self.client.post(
            url, {**data, "csrfmiddlewaretoken": token}, **headers
        )

    def test_create_logs_event(self):
        resp = self._post(
            "/programmes/create/", {"code": "ME", "name": "Mech Eng"}, hx=True
        )
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        log = ActivityLog.objects.get(action=LogAction.CREATE)
        self.assertEqual(log.resource, "Programme")
        self.assertIn("ME", log.message)

    def test_edit_logs_event(self):
        self._post("/programmes/%d/edit/" % self.programme.pk,
                   {"code": "CE", "name": "Civil Updated"})
        log = ActivityLog.objects.get(action=LogAction.UPDATE)
        self.assertEqual(log.resource, "Programme")
        self.assertIn("Civil Updated", log.message)

    def test_delete_logs_event(self):
        self._post("/programmes/%d/delete/" % self.programme.pk, {})
        log = ActivityLog.objects.get(action=LogAction.DELETE)
        self.assertIn("Civil Engineering", log.message)
        self.assertFalse(Programme.objects.filter(pk=self.programme.pk).exists())

    def test_import_logs_event(self):
        self.client.get("/")
        token = self.client.cookies.get("csrftoken").value
        path = make_xlsx([["ME", "Mechanical Engineering"]], ["code", "name"])
        with open(path, "rb") as fh:
            resp = self.client.post(
                "/import/programmes/",
                {
                    "file": SimpleUploadedFile(
                        "prog.xlsx", fh.read(), content_type=XLSX_CONTENT_TYPE
                    ),
                    "csrfmiddlewaretoken": token,
                },
                HTTP_HX_REQUEST="true",
            )
        self.assertEqual(resp.status_code, 200)
        log = ActivityLog.objects.get(action=LogAction.IMPORT)
        self.assertEqual(log.resource, "Programmes")
        self.assertIn("1 created", log.message)

    def test_sidebar_lists_latest_events(self):
        ActivityLog.objects.create(
            action=LogAction.CREATE,
            message="Created Venue ZZZ99",
            resource="Venue",
            target="ZZZ99",
        )
        resp = self.client.get("/")
        html = resp.content.decode()
        self.assertIn("Activity Log", html)
        self.assertIn("Created Venue ZZZ99", html)
        self.assertIn("/?log=", html)

    def test_dashboard_shows_selected_log_and_cancel(self):
        log = ActivityLog.objects.create(
            action=LogAction.UPDATE,
            message="Updated Venue LH1",
            resource="Venue",
            target="LH1",
        )
        resp = self.client.get("/", {"log": log.pk})
        html = resp.content.decode()
        self.assertContains(resp, "Updated Venue LH1")
        self.assertEqual(html.count("Cancel"), 1)
        # Cancel is a plain link back to the bare dashboard (no ?log param).
        cancel_idx = html.find("Cancel")
        self.assertNotIn("?log=", html[max(0, cancel_idx - 200):cancel_idx])

    def test_invalid_log_param_is_ignored(self):
        resp = self.client.get("/", {"log": "999999"})
        self.assertEqual(resp.status_code, 200)
        self.assertNotContains(resp, "Cancel")

    def test_activity_page_lists_and_filters(self):
        ActivityLog.objects.create(
            action=LogAction.CREATE, message="alpha one", resource="Venue"
        )
        ActivityLog.objects.create(
            action=LogAction.DELETE, message="beta two", resource="Venue"
        )
        resp = self.client.get("/activity/")
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "alpha one")
        self.assertContains(resp, "beta two")
        # htmx partial has no sidebar, so the filter is isolated to the list.
        filtered = self.client.get(
            "/activity/", {"action": "CREATE"}, HTTP_HX_REQUEST="true"
        )
        self.assertContains(filtered, "alpha one")
        self.assertNotContains(filtered, "beta two")

    def test_session_group_ops_logged(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        venue = Venue.objects.create(name="LH1", capacity=80)
        group = StudentGroup.objects.create(programme=self.programme, code="A1")
        session = Session.objects.create(
            semester=sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=venue,
        )
        self._post("/sessions/%d/add-group/" % session.pk, {"group_id": group.pk})
        self.assertTrue(ActivityLog.objects.filter(action=LogAction.ASSIGN).exists())
        self._post(
            "/sessions/%d/remove-group/%d/" % (session.pk, group.pk), {}
        )
        self.assertTrue(ActivityLog.objects.filter(action=LogAction.REMOVE).exists())


class VenueQualityTests(TestCase):
    """Venue name normalisation, quality flags and the recycle workbench."""

    def setUp(self):
        self.client = Client()

    def test_base_key_normalises_case_and_spacing(self):
        self.assertEqual(base_key("A104"), "A104")
        self.assertEqual(base_key("a104"), "A104")
        self.assertEqual(base_key("A 104"), "A104")
        self.assertEqual(base_key("  a104  "), "A104")
        self.assertEqual(
            base_key("DO1 luhanga hall kijitonyama"),
            base_key("DO1 kijitonyama"),
        )
        # Compound codes are kept whole, never fused with a single room key.
        self.assertNotEqual(base_key("B4-206"), base_key("B4"))
        self.assertNotEqual(base_key("A104, A106"), base_key("A104"))
        self.assertEqual(base_key("A104, A106"), base_key(" a104 , a106 "))

    def test_issues_for_labels_problem_kinds(self):
        self.assertEqual(issues_for("a104"), ["Casing"])
        self.assertEqual(issues_for("A 104"), ["Spacing"])
        self.assertEqual(issues_for("THEATER 1"), ["Spacing"])
        self.assertEqual(issues_for("DO1 luhanga hall kijitonyama"), ["Casing"])
        self.assertEqual(suggested_name("a104"), "A104")
        self.assertEqual(suggested_name("A 104"), "A104")
        self.assertEqual(
            suggested_name("DO1 luhanga hall kijitonyama"), "DO1"
        )
        self.assertEqual(issues_for("A104"), [])

    def test_analyse_venues_flags_duplicates_and_formatting(self):
        v1 = Venue.objects.create(name="D01 KIJITONYAMA", capacity=200)
        v2 = Venue.objects.create(name="D01 Luhanga Hall Kijitonyama", capacity=220)
        Venue.objects.create(name="a104", capacity=60)
        Venue.objects.create(name="A104", capacity=60)
        Venue.objects.create(name="B4", capacity=30)
        Venue.objects.create(name="B4-206", capacity=40)
        clean = Venue.objects.create(name="LH1", capacity=80)

        issues_map, groups, has_issues = analyse_venues()
        self.assertTrue(has_issues)
        for name in (v1.name, v2.name, "a104"):
            pk = Venue.objects.get(name=name).pk
            self.assertIn(pk, issues_map)
            self.assertIn("Duplicate", issues_map[pk]["issues"])
        # Formatting-only venue is flagged but not in a duplicate group.
        self.assertIn(
            "Casing", issues_map[Venue.objects.get(name="a104").pk]["issues"]
        )
        self.assertNotIn(clean.pk, issues_map)
        self.assertNotIn(Venue.objects.get(name="B4").pk, issues_map)
        self.assertNotIn(Venue.objects.get(name="B4-206").pk, issues_map)
        keys = {g["key"] for g in groups}
        self.assertEqual(keys, {"D01", "A104"})

    def test_venue_list_shows_recycle_button_only_when_problems(self):
        Venue.objects.create(name="a104", capacity=60)
        resp = self.client.get("/venues/")
        html = resp.content.decode()
        self.assertIn("Recycle &amp; Clean", html)
        self.assertIn(">Quality<", html)
        self.assertContains(resp, "Casing")

        # A pristine dataset hides both the banner column and the recycle button.
        Venue.objects.get(name="a104").delete()
        resp = self.client.get("/venues/")
        self.assertNotIn("Recycle &amp; Clean", resp.content.decode())
        self.assertNotIn(">Quality<", resp.content.decode())

    def test_recycle_page_lists_problem_venues_and_offers_fixes(self):
        Venue.objects.create(name="A 104", capacity=60)
        Venue.objects.create(name="A104", capacity=60)
        resp = self.client.get("/venues/recycle/")
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("A 104", html)
        self.assertIn("name='action' value='edit'", html.replace('"', "'"))
        self.assertIn("name='action' value='delete'", html.replace('"', "'"))

    def test_recycle_edit_fixes_name_and_updates_database(self):
        v = Venue.objects.create(name="A 104", capacity=60)
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "edit", "pk": v.pk, "name": "A104", "capacity": "60"},
        )
        v.refresh_from_db()
        self.assertEqual(v.name, "A104")
        self.assertContains(resp, "Database updated")
        self.assertEqual(
            ActivityLog.objects.filter(
                action=LogAction.UPDATE, target="A104"
            ).count(),
            1,
        )

    def test_recycle_edit_merges_into_case_insensitive_clash(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        keep = Venue.objects.create(name="A104", capacity=0)
        dup = Venue.objects.create(name="a104", capacity=60)
        session = Session.objects.create(
            semester=sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=dup,
        )
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "edit", "pk": dup.pk, "name": "A104", "capacity": "60"},
        )
        self.assertContains(resp, "Merged")
        self.assertFalse(Venue.objects.filter(pk=dup.pk).exists())
        self.assertTrue(Venue.objects.filter(pk=keep.pk).exists())
        session.refresh_from_db()
        self.assertEqual(session.venue_id, keep.pk)

    def test_recycle_delete_folds_duplicate_and_repoints_references(self):
        sem = Semester.objects.create(academic_year="2025/2026", semester=1)
        dup = Venue.objects.create(name="D01 Luhanga Hall Kijitonyama", capacity=220)
        keep = Venue.objects.create(name="D01 KIJITONYAMA", capacity=200)
        session = Session.objects.create(
            semester=sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="TUESDAY",
            start_time="10:00",
            end_time="12:00",
            venue=dup,
        )
        WorkshopAllocation.objects.create(
            semester=sem,
            course_code="ME201",
            group_code="A1",
            day="TUESDAY",
            start_time="14:00",
            end_time="17:00",
            venue=dup.name,
        )
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "delete", "pk": dup.pk},
        )
        self.assertContains(resp, "Removed duplicate")
        self.assertFalse(Venue.objects.filter(pk=dup.pk).exists())
        session.refresh_from_db()
        self.assertEqual(session.venue_id, keep.pk)
        self.assertEqual(
            WorkshopAllocation.objects.get(course_code="ME201").venue,
            keep.name,
        )