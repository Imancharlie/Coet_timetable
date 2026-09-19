import os
import tempfile
from pathlib import Path

import pandas as pd
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import Client, TestCase
from openpyxl import Workbook

from core.importers import (
    import_master_timetable_from_excel,
    import_td_allocation_from_excel,
    import_workshop_allocation_from_excel,
    reconcile_master_timetable,
    reconcile_workshop_workbook,
)
from core.models import (
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
        ProgrammeCourse.objects.create(programme=self.prog_a, course_code="MT161")
        ProgrammeCourse.objects.create(programme=self.prog_b, course_code="MT161")
        ProgrammeCourse.objects.create(programme=self.prog_a, course_code="TG201")
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
                {"programme": self.programme.pk, "course_code": "MT161"},
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
        ProgrammeCourse.objects.create(programme=prog, course_code="MT161")
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

    def test_import_matrix_blocked_without_semester(self):
        Semester.objects.filter(academic_year="2025/2026", semester=1).delete()
        result = import_workshop_allocation_from_excel(self._matrix_path())
        self.assertTrue(result.missing_references)
        self.assertTrue(any("blocked" in err for err in result.errors))
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

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