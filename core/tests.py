import os
import json
import re
import tempfile
from datetime import time
from pathlib import Path
from unittest import mock

import pandas as pd
from django.core.exceptions import ValidationError
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
    ActivityType,
    ImportHistory,
    LogAction,
    Programme,
    ProgrammeCourse,
    Semester,
    Session,
    SessionGroup,
    StudentGroup,
    TechnicalDrawingAllocation,
    TimePeriod,
    Venue,
    WorkshopAllocation,
)
from core.timetable_grid import (
    build_day_time_grid,
    build_time_day_grid,
    time_day_grid_to_table,
)
from core.timetable_pdf import (
    _build_day_flowables,
    _compact_group_codes,
    _entry_lines,
    _master_cell_text,
    _merge_master_entries,
    _wrap_line,
    _DayFlowable,
    build_grid,
    collect_entries,
    collect_group_entries,
    collect_master_entries,
    collect_workshop_rotations,
    render_group_timetable,
    render_programme_timetable,
    render_udsm_master_timetable,
)
from core.venue_quality import (
    analyse_venues,
    base_key,
    conflict_reasons,
    detect_name_conflicts,
    issues_for,
    resolve_venue_name_conflict,
    suggested_name,
)
from core.workshop_parser import parse_workbook
from core.workshop_times import (
    allocation_programme_codes,
    course_programme_codes,
    legacy_workshop_allocations,
    legacy_workshop_sessions,
    session_programme_codes,
    validate_workshop_record,
    validate_workshop_session,
    workshop_hours,
    workshop_times_for,
    workshop_time_issue,
)

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
        self.assertEqual(SessionGroup.objects.count(), 5)

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
        self.assertNotIn(
            "A9", session.session_groups.values_list("group__code", flat=True)
        )
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
            [["TG201", "C1", "MONDAY", "09:00", "13:00", "TW101"]],
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

    def test_workshop_heading_accepted_as_venue_legacy(self):
        self._seed()
        path = make_xlsx(
            [["E1", "MONDAY", "09:00", "13:00", "Electrical"]],
            ["group_code", "day", "start_time", "end_time", "workshop"],
        )
        first = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(first.errors, [])
        self.assertEqual(first.created, 1)
        wa = WorkshopAllocation.objects.get()
        self.assertEqual(wa.course_code, "Electrical")
        self.assertEqual(wa.venue, "Electrical")

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

    def test_master_import_without_semester_is_blocked(self):
        Semester.objects.all().delete()
        result = import_master_timetable_from_excel(
            make_xlsx(
                [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
                MASTER_COLS,
            )
        )
        self.assertTrue(
            any("semester" in err.lower() for err in result.errors), result.errors
        )
        self.assertFalse(Semester.objects.exists())
        self.assertEqual(Session.objects.count(), 0)

    def test_master_import_auto_assigns_lecture_groups(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", ""]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        session = Session.objects.get(course_code="MT161", semester=self.sem1)
        self.assertEqual(
            set(session.session_groups.values_list("group__code", flat=True)),
            {"A1", "A2", "B1"},
        )
        self.assertEqual(result.lecture_sessions_processed, 1)
        self.assertEqual(result.lecture_groups_linked, 3)
        self.assertEqual(result.lecture_groups_existing, 0)

    def test_master_import_auto_assign_is_idempotent(self):
        self._seed()
        path = make_xlsx(
            [["MT161", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", ""]],
            MASTER_COLS,
        )
        import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        second = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(second.lecture_groups_linked, 0)
        self.assertEqual(second.lecture_groups_existing, 3)
        self.assertEqual(SessionGroup.objects.count(), 3)

    def test_master_import_auto_assign_skips_non_lectures(self):
        self._seed()
        path = make_xlsx(
            [["TG201", "TUTORIAL", "TUESDAY", "11:00", "12:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        session = Session.objects.get(course_code="TG201", semester=self.sem1)
        self.assertEqual(session.session_groups.count(), 1)  # only the explicit A1
        self.assertEqual(result.lecture_sessions_processed, 0)

    def test_master_import_auto_assign_reports_unmapped_courses(self):
        self._seed()
        path = make_xlsx(
            [["GHOST100", "LECTURE", "MONDAY", "08:00", "10:00", "LH1", ""]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertIn("GHOST100", result.lecture_skipped_courses)
        session = Session.objects.get(course_code="GHOST100", semester=self.sem1)
        self.assertEqual(session.session_groups.count(), 0)

    def test_workshop_flat_accepts_time_range_column(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "C1", "MONDAY", "09:00-13:00", "TW101"]],
                ["course_code", "group_code", "day", "time", "venue"],
            )
        )
        self.assertEqual(result.created, 1)
        wa = WorkshopAllocation.objects.get(course_code="TG201", group_code="C1")
        self.assertEqual(wa.start_time, time(9, 0))
        self.assertEqual(wa.end_time, time(13, 0))

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
                    "start_time": "09:00",
                    "end_time": "13:00",
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


class SidebarCollapseTests(TestCase):
    """Collapsible/expandable sidebar: toggle button, icon-only mode with
    tooltips, mobile drawer behaviour, a11y/keyboard use, active-state
    preservation, preference persistence and responsive content resizing."""

    def _html(self, path="/"):
        resp = self.client.get(path)
        self.assertEqual(resp.status_code, 200)
        return resp.content.decode()

    def test_collapse_button_at_top_with_controls(self):
        html = self._html()
        ctrl = html.find('aria-controls="app-sidebar"')
        self.assertGreater(ctrl, -1)
        self.assertLess(ctrl, html.find("<nav"))
        self.assertIn('id="app-sidebar"', html)
        btn = html[max(0, ctrl - 200):ctrl + 150]
        self.assertIn('type="button"', btn)
        self.assertIn('@click="toggle()"', btn)

    def test_animated_width_switch(self):
        html = self._html()
        self.assertIn("transition-[width,transform] duration-300", html)
        self.assertIn("collapsed ? 'lg:w-20' : 'lg:w-64'", html)
        self.assertIn(
            "collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'",
            html,
        )
        self.assertIn("transition-all duration-300", html)

    def test_icon_only_tooltips(self):
        html = self._html()
        for text in (
            "Dashboard",
            "Export Timetable",
            "Venues",
            "Workshop Allocation",
        ):
            self.assertIn("showTooltip($el, '%s')" % text, html)
        self.assertIn('role="tooltip"', html)
        self.assertIn("pointer-events-none", html)
        self.assertIn('x-show="collapsed && tooltip"', html)
        self.assertIn('x-text="tooltip"', html)

    def test_tooltips_hide_when_expanded(self):
        html = self._html()
        self.assertIn("window.innerWidth >= 1024", html)
        self.assertIn("showTooltip: function", html)
        self.assertIn("hideTooltip: function", html)

    def test_navigation_keeps_icons_and_hrefs(self):
        html = self._html()
        for href, aria in (
            ("/", "Dashboard"),
            ("/export/", "Export Timetable"),
            ("/venues/", "Venues"),
            ("/workshops/", "Workshop Allocation"),
        ):
            idx = html.find('aria-label="%s"' % aria)
            self.assertGreater(idx, -1)
            link = html[html.rfind("<a", 0, idx):html.find("</a>", idx)]
            self.assertIn('href="%s"' % href, link)
            self.assertIn("showTooltip($el, '%s')" % aria, link)
        self.assertGreater(html.count("<svg"), 15)

    def test_active_state_preserved_in_both_modes(self):
        html = self._html("/programmes/")
        start = html.find('href="/programmes/"')
        prog_link = html[start:html.find("</a>", start)]
        self.assertIn("bg-slate-800 text-white", prog_link)
        self.assertIn("lg:justify-center lg:gap-0", prog_link)
        self.assertIn("<svg", prog_link)
        self.assertIn('lg:max-w-0', html)
        self.assertEqual(html.count("bg-slate-800 text-white"), 1)
        dash_html = self._html("/")
        self.assertEqual(dash_html.count("bg-slate-800 text-white"), 1)

    def test_mobile_drawer_behaviour(self):
        html = self._html()
        self.assertIn("fixed lg:sticky", html)
        self.assertIn("translate-x-0", html)
        self.assertIn("-translate-x-full", html)
        self.assertIn("lg:hidden", html)
        self.assertIn('x-show="sidebarOpen"', html)
        self.assertIn('@click="sidebarOpen = false"', html)
        self.assertIn("hidden lg:inline-flex", html)

    def test_keyboard_and_accessibility(self):
        html = self._html()
        self.assertIn(":aria-expanded=", html)
        self.assertIn("Collapse sidebar", html)
        self.assertIn("Expand sidebar", html)
        self.assertIn("aria-controls=", html)
        self.assertIn("focus-visible:ring", html)
        self.assertIn(':aria-label="collapsed ? \'Expand sidebar\' : \'Collapse sidebar\'"', html)
        self.assertIn("@focus=\"showTooltip($el, 'Dashboard')\"", html)
        self.assertIn("@blur=\"hideTooltip()\"", html)
        self.assertIn("focus:outline-none", html)

    def test_responsive_content_resizing(self):
        html = self._html()
        self.assertIn('class="flex-1 flex flex-col min-w-0"', html)
        self.assertIn("shrink-0", html)
        self.assertIn("ease-in-out", html)
        self.assertIn("duration-300", html)

    def test_preference_persists_across_refreshes(self):
        html = self._html()
        self.assertIn("coet.sidebar.collapsed", html)
        self.assertIn("sessionStorage", html)
        self.assertIn("localStorage", html)
        self.assertIn("Alpine.data('sidebar',", html)
        self.assertIn('x-data="sidebar"', html)


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
                    ),
                    "semester": sem.pk,
                },
                HTTP_HX_REQUEST="true",
            )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Reconciled", html)
        self.assertIn("ALL", html)
        self.assertIn("2026/2027 - Semester 1", html)
        self.assertEqual(Session.objects.count(), 1)
        session = Session.objects.first()
        self.assertEqual(session.session_groups.count(), 1)

    def test_master_timetable_upload_requires_semester(self):
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
        self.assertIn("semester", html.lower())
        self.assertEqual(Session.objects.count(), 0)

    def test_master_timetable_upload_bad_semester_blocked(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        prog = Programme.objects.create(code="CE", name="Civil Engineering")
        StudentGroup.objects.create(programme=prog, code="A1")
        ProgrammeCourse.objects.create(
            programme=prog, course_code="MT161", course_name="Mathematics 1", semester=1
        )
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
                    ),
                    "semester": 9999,
                },
                HTTP_HX_REQUEST="true",
            )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("does not exist", resp.content.decode())
        self.assertEqual(Session.objects.count(), 0)

    def test_master_timetable_upload_shows_semester_select(self):
        Semester.objects.create(academic_year="2026/2027", semester=1)
        resp = self.client.get("/import/master-timetable/")
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Academic Semester", html)
        self.assertIn('name="semester"', html)
        self.assertIn('value="1"', html)
        self.assertIn("* required", html)


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

    def test_assignment_warnings_do_not_toast(self):
        # PHY200 has no programme mapping in the fixture. The warning lives in
        # the green results panel; no red toast is dispatched for it.
        resp = self.client.post(
            "/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true"
        )
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("HX-Trigger", resp.headers)
        html = resp.content.decode()
        self.assertIn("Lecture group assignment complete", html)
        self.assertIn("No programme mapping found for: PHY200", html)
        self.assertEqual(html.count('aria-label="Dismiss notification"'), 1)

    def test_clean_assignment_no_duplicate_toast(self):
        # Success details are already in the green panel; never a second toast.
        resp = self._clean_response()
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("HX-Trigger", resp.headers)
        html = resp.content.decode()
        self.assertIn("Lecture sessions processed: 3 of 3", html)
        self.assertIn("Group links created: 8", html)
        self.assertIn("Already linked: 0", html)

    def test_page_keeps_global_error_notification(self):
        # Genuine failures (server/db errors or lost connections) still surface
        # a red toast through the shared htmx error handlers on every page.
        page = self.client.get("/sessions/").content.decode()
        self.assertIn("htmx:responseError", page)
        self.assertIn("Request failed", page)
        self.assertIn("htmx:sendError", page)

    def test_no_lecture_sessions_no_toast(self):
        Session.objects.filter(activity_type="LECTURE").delete()
        resp = self.client.post(
            "/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true"
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("No lecture sessions found", resp.content.decode())
        self.assertNotIn("HX-Trigger", resp.headers)

    def _clean_response(self):
        ProgrammeCourse.objects.get_or_create(
            programme=self.prog_a,
            course_code="PHY200",
            defaults={"course_name": "Physics 2", "semester": 1},
        )
        return self.client.post(
            "/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true"
        )

    def test_success_message_supports_manual_dismiss(self):
        html = self._clean_response().content.decode()
        self.assertEqual(html.count('aria-label="Dismiss notification"'), 1)
        self.assertIn('type="button"', html)
        self.assertIn('@click="dismiss()"', html)
        self.assertIn("Dismiss notification", html)

    def test_clean_success_message_auto_dismisses(self):
        html = self._clean_response().content.decode()
        # Auto-dismiss marker only on the clean-success panel, keeping the
        # full summary visible while it is active.
        self.assertIn('data-auto-dismiss="7000"', html)
        self.assertIn("Lecture sessions processed: 3 of 3", html)
        self.assertIn("Group links created: 8", html)
        self.assertIn("Already linked: 0", html)

    def test_warning_message_persists_until_dismissed(self):
        # PHY200 has no programme mapping in the fixture -> warning.
        html = self.client.post(
            "/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true"
        ).content.decode()
        self.assertNotIn("data-auto-dismiss", html)
        self.assertEqual(html.count('aria-label="Dismiss notification"'), 1)
        self.assertIn("No programme mapping found for: PHY200", html)

    def test_no_sessions_message_persists_until_dismissed(self):
        Session.objects.filter(activity_type="LECTURE").delete()
        html = self.client.post(
            "/sessions/assign-lecture-groups/", {}, HTTP_HX_REQUEST="true"
        ).content.decode()
        self.assertNotIn("data-auto-dismiss", html)
        self.assertIn("No lecture sessions found", html)
        self.assertEqual(html.count('aria-label="Dismiss notification"'), 1)

    def test_repeated_assignments_never_accumulate_messages(self):
        # Each run swaps the result panel, so at most one message exists.
        for _ in range(3):
            html = self._clean_response().content.decode()
            self.assertEqual(html.count('aria-label="Dismiss notification"'), 1)
            self.assertEqual(html.count('x-data="assignResult"'), 1)


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

    def test_recycle_formatting_fix_without_capacity_succeeds_and_keeps_capacity(self):
        # The formatting-only Fix form posts just a name (no capacity input).
        # Resolving the casing/spacing duplicate must not be blocked by the
        # "capacity: This field is required" validation and must preserve the
        # venue's existing capacity.
        v = Venue.objects.create(name="a104", capacity=60)
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "edit", "pk": v.pk, "name": "A104"},
        )
        v.refresh_from_db()
        self.assertEqual(v.name, "A104")
        self.assertEqual(v.capacity, 60)
        self.assertContains(resp, "Database updated")
        self.assertNotIn("This field is required", resp.content.decode())
        self.assertEqual(
            ActivityLog.objects.filter(
                action=LogAction.UPDATE, target="A104"
            ).count(),
            1,
        )

    def test_recycle_formatting_fix_with_zero_capacity_succeeds(self):
        # A capacity-0 venue renders "" in the capacity input; the duplicate
        # fix for its name must still apply without touching capacity.
        v = Venue.objects.create(name="A 104", capacity=0)
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "edit", "pk": v.pk, "name": "A104"},
        )
        v.refresh_from_db()
        self.assertEqual(v.name, "A104")
        self.assertEqual(v.capacity, 0)
        self.assertNotIn("This field is required", resp.content.decode())

    def test_recycle_duplicate_fix_with_blank_capacity_merges_and_keeps_capacity(self):
        keep = Venue.objects.create(name="A104", capacity=60)
        dup = Venue.objects.create(name="a104", capacity=0)
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "edit", "pk": dup.pk, "name": "A104", "capacity": ""},
        )
        self.assertContains(resp, "Merged")
        self.assertFalse(Venue.objects.filter(pk=dup.pk).exists())
        keep.refresh_from_db()
        self.assertEqual(keep.name, "A104")
        self.assertEqual(keep.capacity, 60)

    def test_recycle_edit_still_validates_capacity_when_supplied(self):
        v = Venue.objects.create(name="A 104", capacity=60)
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "edit", "pk": v.pk, "name": "A104", "capacity": "abc"},
        )
        v.refresh_from_db()
        self.assertEqual(v.name, "A 104")
        html = resp.content.decode()
        self.assertIn("capacity", html.lower())
        self.assertIn("Enter a whole number", html)

    def test_recycle_edit_blank_name_still_rejected(self):
        v = Venue.objects.create(name="A 104", capacity=60)
        resp = self.client.post(
            "/venues/recycle/",
            {"action": "edit", "pk": v.pk, "name": ""},
        )
        v.refresh_from_db()
        self.assertEqual(v.name, "A 104")
        html = resp.content.decode()
        self.assertIn("name", html.lower())
        self.assertIn("This field is required", html)


class VenueImportConflictTests(TestCase):
    """Detect spacing/casing venue-name duplicates at import time and let the
    user interactively pick the official name, which is then merged."""

    def setUp(self):
        self.client = Client()

    # ---- Detection primitives ----------------------------------------------

    def test_conflict_reasons_labels_kinds(self):
        self.assertEqual(conflict_reasons("PB 06", "PB06"), ["Spacing"])
        self.assertEqual(conflict_reasons("PB 06", "pb 06"), ["Casing"])
        self.assertEqual(
            sorted(conflict_reasons("pb 06", "PB06")), ["Casing", "Spacing"]
        )
        self.assertEqual(conflict_reasons("A104", "B104"), [])
        self.assertEqual(conflict_reasons("A104", "A104"), [])
        # Locator-style duplicates (same room code, different notes) are not
        # formatting conflicts — they stay on the recycle workbench.
        self.assertEqual(
            conflict_reasons("D01 Luhanga Hall Kijitonyama", "D01 Kijitonyama"),
            [],
        )

    def test_detect_name_conflicts_lists_every_pair(self):
        conflicts = detect_name_conflicts(["PB 06", "PB06", "pb 06"])
        self.assertEqual(len(conflicts), 3)
        for issue in conflicts:
            self.assertEqual(issue["key"], "PB06")
            self.assertEqual(len(issue["names"]), 2)
            self.assertTrue(issue["reasons"])
        pairs = {tuple(sorted(i["names"])) for i in conflicts}
        self.assertEqual(
            pairs,
            {("PB 06", "PB06"), ("PB 06", "pb 06"), ("PB06", "pb 06")},
        )
        self.assertEqual(detect_name_conflicts(["A104", "A106", "B4"]), [])

    def test_detect_is_idempotent_and_case_insensitive(self):
        first = detect_name_conflicts(["A 104", "a104"])
        second = detect_name_conflicts(["A 104", "A104"])
        self.assertEqual(len(first), 1)
        self.assertEqual(len(second), 1)
        self.assertEqual(set(second[0]["names"]), {"A 104", "A104"})

    # ---- Import-time detection ---------------------------------------------

    def _import(self, rows):
        return import_venues_from_excel(make_xlsx(rows, ["name", "capacity"]))

    def test_venue_import_detects_spacing_conflict(self):
        result = self._import([["PB 06", 100], ["PB06", 100]])
        self.assertEqual(result.created, 2)
        self.assertEqual(len(result.venue_conflicts), 1)
        issue = result.venue_conflicts[0]
        self.assertEqual(issue["reasons"], ["Spacing"])
        self.assertEqual(set(issue["names"]), {"PB 06", "PB06"})
        # Venue formatting conflicts are structured, not free-text.
        self.assertEqual(result.conflicts, [])

    def test_venue_import_detects_casing_conflict(self):
        result = self._import([["PB 06", 100], ["pb 06", 100]])
        self.assertEqual(len(result.venue_conflicts), 1)
        self.assertEqual(result.venue_conflicts[0]["reasons"], ["Casing"])

    def test_venue_import_detects_combined_conflict(self):
        result = self._import([["pb 06", 100], ["PB06", 150]])
        issue = result.venue_conflicts[0]
        self.assertEqual(set(issue["reasons"]), {"Spacing", "Casing"})

    def test_venue_import_flags_clash_with_existing_venue(self):
        Venue.objects.create(name="A104", capacity=60)
        result = self._import([["A 104", 60]])
        self.assertEqual(result.created, 1)
        self.assertEqual(len(result.venue_conflicts), 1)
        self.assertEqual(set(result.venue_conflicts[0]["names"]), {"A104", "A 104"})

    def test_venue_import_clean_data_has_no_conflicts(self):
        result = self._import([["A104", 60], ["B106", 50]])
        self.assertEqual(result.venue_conflicts, [])
        self.assertEqual(len(result.venue_conflicts), 0)

    def test_venue_import_with_missing_capacity_creates_venue_and_flags_it(self):
        # A blank capacity no longer rejects the row: the venue is created with
        # capacity 0 so its name can be duplicate-fixed, and the missing
        # capacity is flagged separately (not as a blocking import error).
        result = self._import([["PB 06", ""]])
        self.assertEqual(result.created, 1)
        self.assertEqual(result.skipped, 0)
        self.assertEqual(result.errors, [])
        self.assertEqual(result.venue_capacity_issues, ["PB 06"])
        self.assertEqual(Venue.objects.get(name="PB 06").capacity, 0)

    def test_venue_import_invalid_capacity_still_errors_and_skips(self):
        # A non-numeric capacity stays a hard validation error: the row is
        # skipped, nothing is created and it is NOT added to the missing-
        # capacity (blank) list.
        result = self._import([["PB 06", "abc"]])
        self.assertEqual(result.created, 0)
        self.assertEqual(result.skipped, 1)
        self.assertEqual(len(result.errors), 1)
        self.assertIn("Invalid capacity for venue 'PB 06'", result.errors[0])
        self.assertEqual(result.venue_capacity_issues, [])
        self.assertFalse(Venue.objects.filter(name="PB 06").exists())

    def test_venue_import_missing_capacity_still_detects_conflict(self):
        result = self._import([["PB 06", ""], ["PB06", 150]])
        self.assertEqual(result.created, 2)
        self.assertEqual(result.venue_capacity_issues, ["PB 06"])
        self.assertEqual(len(result.venue_conflicts), 1)
        self.assertEqual(result.venue_conflicts[0]["reasons"], ["Spacing"])
        self.assertEqual(set(result.venue_conflicts[0]["names"]), {"PB 06", "PB06"})

    # ---- Interactive web fix flow ------------------------------------------

    def _upload_venues(self, rows):
        self.client.get("/import/venues/")
        path = make_xlsx(rows, ["name", "capacity"])
        with open(path, "rb") as fh:
            return self.client.post(
                "/import/venues/",
                {
                    "file": SimpleUploadedFile(
                        "venues.xlsx", fh.read(), content_type=XLSX_CONTENT_TYPE
                    )
                },
                HTTP_HX_REQUEST="true",
            )

    def _token(self):
        for key in self.client.session.keys():
            if key.startswith("venue_conflicts:"):
                return key.split(":", 1)[1]
        self.fail("no venue conflict token in session")

    def _open_issue(self, key):
        token = self._token()
        store = self.client.session["venue_conflicts:%s" % token]
        return token, next(issue for issue in store["open"] if issue["key"] == key)

    def test_upload_lists_conflicts_with_fix_chooser(self):
        resp = self._upload_venues([["PB 06", 100], ["PB06", 150]])
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("1 remaining", html)
        self.assertIn("Spacing difference", html)
        self.assertIn("PB 06", html)
        self.assertIn("PB06", html)
        self.assertIn("Accept &amp; Merge", html)
        self.assertIn("Which name should be the official venue name", html)

    def test_fix_accepts_official_name_and_merges_references(self):
        self._upload_venues([["PB 06", 100], ["PB06", 150]])
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        session = Session.objects.create(
            semester=sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=Venue.objects.get(name="PB 06"),
        )
        WorkshopAllocation.objects.create(
            semester=sem,
            course_code="ME201",
            group_code="A1",
            day="TUESDAY",
            start_time="14:00",
            end_time="17:00",
            venue="PB 06",
        )

        token, issue = self._open_issue("PB06")
        resp = self.client.post(
            "/venues/fix-conflict/",
            {
                "token": token,
                "key": issue["key"],
                "issue_id": issue["id"],
                "official": "PB06",
            },
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("All resolved", html)
        self.assertIn("Accepted: PB06", html)
        self.assertIn("Fixed", html)
        self.assertNotIn("remaining", html)

        # No duplicates left; every reference now uses the official name.
        self.assertFalse(Venue.objects.filter(name="PB 06").exists())
        self.assertTrue(Venue.objects.filter(name="PB06").exists())
        self.assertEqual(Venue.objects.count(), 1)
        session.refresh_from_db()
        self.assertEqual(session.venue.name, "PB06")
        self.assertEqual(
            WorkshopAllocation.objects.get(course_code="ME201").venue, "PB06"
        )

        # The audit trail keeps the original value, the chosen official value,
        # the action and the timestamp.
        logs = ActivityLog.objects.filter(action=LogAction.UPDATE, target="PB06")
        self.assertEqual(logs.count(), 1)
        log = logs.get()
        self.assertIn("PB 06", log.message)
        self.assertIn("official", log.message)
        self.assertIsNotNone(log.created_at)

    def test_remaining_count_updates_after_each_fix(self):
        resp = self._upload_venues(
            [["PB 06", 100], ["PB06", 150], ["LH 1", 80], ["LH1", 80]]
        )
        self.assertIn("2 remaining", resp.content.decode())

        token, lh_issue = self._open_issue("LH1")
        resp = self.client.post(
            "/venues/fix-conflict/",
            {
                "token": token,
                "key": lh_issue["key"],
                "issue_id": lh_issue["id"],
                "official": "LH1",
            },
            HTTP_HX_REQUEST="true",
        )
        html = resp.content.decode()
        self.assertIn("1 remaining", html)
        self.assertIn("Accepted: LH1", html)
        self.assertFalse(Venue.objects.filter(name="LH 1").exists())
        self.assertTrue(Venue.objects.filter(name="PB 06").exists())
        self.assertTrue(Venue.objects.filter(name="PB06").exists())

        token, pb_issue = self._open_issue("PB06")
        resp = self.client.post(
            "/venues/fix-conflict/",
            {
                "token": token,
                "key": pb_issue["key"],
                "issue_id": pb_issue["id"],
                "official": "PB06",
            },
            HTTP_HX_REQUEST="true",
        )
        html = resp.content.decode()
        self.assertIn("All resolved", html)
        self.assertNotIn("remaining", html)
        self.assertEqual(Venue.objects.count(), 2)
        self.assertEqual(Venue.objects.values_list("name", flat=True).count(), 2)

    def test_fix_spacing_conflict_with_missing_capacity_preserves_capacity(self):
        # One row has no capacity: the conflict still appears and fixing it is
        # not blocked, while the accepted venue keeps its real capacity.
        resp = self._upload_venues([["PB 06", ""], ["PB06", 150]])
        html = resp.content.decode()
        self.assertIn("1 remaining", html)
        self.assertIn("Spacing difference", html)
        self.assertIn("missing capacity", html)
        self.assertNotIn("capacity: This field is required", html)
        self.assertNotIn("Invalid capacity", html)
        self.assertEqual(Venue.objects.get(name="PB 06").capacity, 0)

        token, issue = self._open_issue("PB06")
        resp = self.client.post(
            "/venues/fix-conflict/",
            {
                "token": token,
                "key": issue["key"],
                "issue_id": issue["id"],
                "official": "PB06",
            },
            HTTP_HX_REQUEST="true",
        )
        html = resp.content.decode()
        self.assertIn("All resolved", html)
        self.assertIn("Accepted: PB06", html)
        self.assertNotIn("remaining", html)
        self.assertFalse(Venue.objects.filter(name="PB 06").exists())
        self.assertEqual(Venue.objects.get(name="PB06").capacity, 150)

    def test_fix_casing_conflict_with_empty_capacity(self):
        # Both colliding names lack capacity; accepting an official name must
        # still resolve the casing duplicate without any capacity error.
        resp = self._upload_venues([["PB 06", ""], ["pb 06", ""]])
        self.assertIn("Casing difference", resp.content.decode())
        self.assertNotIn("Invalid capacity", resp.content.decode())

        token, issue = self._open_issue("PB06")
        resp = self.client.post(
            "/venues/fix-conflict/",
            {
                "token": token,
                "key": issue["key"],
                "issue_id": issue["id"],
                "official": "pb 06",
            },
            HTTP_HX_REQUEST="true",
        )
        html = resp.content.decode()
        self.assertIn("All resolved", html)
        self.assertIn("Accepted: pb 06", html)
        self.assertEqual(Venue.objects.count(), 1)
        venue = Venue.objects.get()
        self.assertEqual(venue.name, "pb 06")
        self.assertEqual(venue.capacity, 0)

    def test_fix_rejects_official_not_offered(self):
        self._upload_venues([["PB 06", 100], ["PB06", 150]])
        token, issue = self._open_issue("PB06")
        resp = self.client.post(
            "/venues/fix-conflict/",
            {
                "token": token,
                "key": issue["key"],
                "issue_id": issue["id"],
                "official": "SOME OTHER VENUE",
            },
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("not one of the offered names", html)
        self.assertIn("1 remaining", html)
        self.assertTrue(Venue.objects.filter(name="PB 06").exists())
        self.assertTrue(Venue.objects.filter(name="PB06").exists())
        self.assertEqual(
            ActivityLog.objects.filter(target="SOME OTHER VENUE").count(), 0
        )

    def test_fix_without_choice_is_rejected(self):
        self._upload_venues([["PB 06", 100], ["PB06", 150]])
        resp = self.client.post(
            "/venues/fix-conflict/",
            {"token": "", "key": "", "issue_id": "", "official": ""},
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Missing venue choice", resp.content.decode())
        self.assertEqual(Venue.objects.count(), 2)

    def test_resolve_merges_every_colliding_name(self):
        Venue.objects.create(name="pb 06", capacity=100)
        Venue.objects.create(name="PB06", capacity=150)
        Venue.objects.create(name="PB 06", capacity=120)
        target, merged = resolve_venue_name_conflict("PB06", "PB06")
        self.assertEqual(target.name, "PB06")
        self.assertEqual(set(merged), {"pb 06", "PB 06"})
        self.assertEqual(Venue.objects.count(), 1)
        self.assertEqual(Venue.objects.get().name, "PB06")


class TimetableGridTests(TestCase):
    """Classic grid: DAYS as columns, TIME slots as rows, session spanning."""

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        self.g1 = StudentGroup.objects.create(programme=self.prog, code="C1")
        self.g2 = StudentGroup.objects.create(programme=self.prog, code="C2")
        self.venue = Venue.objects.create(name="YOMBO5", capacity=80)

        def make(course, atype, day, start, end):
            return Session.objects.create(
                semester=self.sem,
                course_code=course,
                activity_type=atype,
                day=day,
                start_time=start,
                end_time=end,
                venue=self.venue,
            )

        self.s1 = make("MT171", "LECTURE", "MONDAY", "08:00", "09:00")
        self.s2 = make("EE153", "LECTURE", "WEDNESDAY", "07:00", "09:55")
        self.s3 = make("EE131", "LECTURE", "MONDAY", "08:00", "10:00")
        SessionGroup.objects.create(session=self.s1, group=self.g1)
        SessionGroup.objects.create(session=self.s2, group=self.g1)
        SessionGroup.objects.create(session=self.s3, group=self.g2)

    def test_grid_slots_and_weekday_columns(self):
        grid = build_time_day_grid(collect_entries(self.prog, self.sem))
        self.assertEqual(grid["slots"][0]["label"], "07:00-08:00")
        labels = [d["day"] for d in grid["days"]]
        self.assertEqual(labels, ["MONDAY", "WEDNESDAY"])
        self.assertNotIn("SATURDAY", labels)
        self.assertEqual(len(grid["rows"]), 3)

    def test_multi_slot_session_spans_rows(self):
        grid = build_time_day_grid(collect_entries(self.prog, self.sem))
        di = [d["day"] for d in grid["days"]].index("WEDNESDAY")
        cell = grid["rows"][0]["cols"][di]
        self.assertEqual(cell["rowspan"], 3)
        self.assertEqual(cell["entries"][0]["course_code"], "EE153")
        self.assertIsNone(grid["rows"][1]["cols"][di])
        self.assertIsNone(grid["rows"][2]["cols"][di])

    def test_overlapping_sessions_share_a_block(self):
        grid = build_time_day_grid(collect_entries(self.prog, self.sem))
        di = [d["day"] for d in grid["days"]].index("MONDAY")
        cell = grid["rows"][1]["cols"][di]
        codes = sorted(e["course_code"] for e in cell["entries"])
        self.assertEqual(codes, ["EE131", "MT171"])
        self.assertEqual(grid["rows"][0]["cols"][di]["empty"], True)

    def test_pdf_grid_table_is_classic(self):
        entries = collect_entries(self.prog, self.sem)
        data, spans, fills = build_grid(entries)
        self.assertEqual(data[0][0], "TIME")
        self.assertEqual(data[0][1], "MONDAY")
        self.assertEqual(data[0][2], "WEDNESDAY")
        self.assertEqual(data[1][0], "07:00-08:00")
        self.assertIn(((2, 1), (2, 3)), spans)
        self.assertIn(((2, 1), (2, 3), "#d1d5db"), fills)

    def test_lecture_workshop_td_fill_colors(self):
        workshop = WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="EE151",
            workshop="W1",
            group_code="C1",
            day="WEDNESDAY",
            start_time="10:00",
            end_time="12:00",
            venue="WORKSHOP 1",
        )
        td = TechnicalDrawingAllocation.objects.create(
            semester=self.sem,
            course_code="EE153",
            group_code="C1",
            day="FRIDAY",
            start_time="13:00",
            end_time="17:00",
            venue="TD LAB",
        )
        data, spans, fills = build_grid(collect_entries(self.prog, self.sem))
        fill_map = {(start, end): color for start, end, color in fills}
        workshop_cell = (
            (2, 4),
            (2, 5),
        )
        self.assertEqual(fill_map[workshop_cell], "#dcfce7")
        self.assertIn("#fce7f3", fill_map.values())

    def test_collect_entries_for_single_group(self):
        g1_entries = collect_entries(self.prog, self.sem, group=self.g1)
        self.assertEqual(
            {e["course_code"] for e in g1_entries}, {"MT171", "EE153"}
        )
        g2_entries = collect_entries(self.prog, self.sem, group=self.g2)
        self.assertEqual({e["course_code"] for e in g2_entries}, {"EE131"})
        self.assertNotIn("EE131", {e["course_code"] for e in g1_entries})


class WorkshopCellDisplayTests(TestCase):
    """Workshop cells show the meaningful workshop/category names only.

    Covers multiple student groups, multiple programmes and multiple workshop
    categories, for both the flat (course_code derived from venue) and raw
    matrix (category in course_code + workshop) data shapes.
    """

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog_ee = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        self.prog_me = Programme.objects.create(
            code="ME", name="BSc. in Mechanical Engineering"
        )
        self.ee_c1 = StudentGroup.objects.create(programme=self.prog_ee, code="C1")
        self.ee_c2 = StudentGroup.objects.create(programme=self.prog_ee, code="C2")
        self.me_b1 = StudentGroup.objects.create(programme=self.prog_me, code="B1")

    def _flat_workshop(self, name, group="C2", day="MONDAY"):
        """Flat FORMAT B style: course_code is derived from the venue cell."""
        return WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code=name,
            group_code=group,
            day=day,
            start_time="09:00",
            end_time="13:00",
            venue=name,
        )

    def _flattened_cell(self, data, day, slot):
        header = data[0]
        self.assertIn(day, header)
        di = header.index(day)
        for row in data[1:]:
            if row[0] == slot:
                return row[di]
        self.fail(f"slot {slot} not found in grid")

    def test_group_workshop_cell_shows_names_only(self):
        self._flat_workshop("Building")
        self._flat_workshop("Electronics", day="TUESDAY")
        data, _, _ = build_grid(
            collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        )
        mon = self._flattened_cell(data, "MONDAY", "09:00-10:00")
        tue = self._flattened_cell(data, "TUESDAY", "09:00-10:00")
        self.assertEqual(mon, "Building")
        self.assertEqual(tue, "Electronics")
        for cell in (mon, tue):
            self.assertNotIn("WORKSHOP", cell)
            self.assertNotIn("C2", cell)
            self.assertNotIn("Building Building", cell)
            self.assertNotIn("Electronics Electronics", cell)

    def test_single_workshop_cell_shows_only_name(self):
        self._flat_workshop("Welding")
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        self.assertEqual([e["label"] for e in entries], ["Welding"])
        data, _, _ = build_grid(entries)
        self.assertEqual(
            self._flattened_cell(data, "MONDAY", "09:00-10:00"), "Welding"
        )

    def test_distinct_workshops_on_different_days_all_kept(self):
        self._flat_workshop("Carpentry", day="MONDAY")
        self._flat_workshop("Plumbing", day="TUESDAY")
        self._flat_workshop("Welding", day="WEDNESDAY")
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        self.assertEqual(
            {e["label"] for e in entries}, {"Carpentry", "Plumbing", "Welding"}
        )
        data, _, _ = build_grid(entries)
        self.assertEqual(
            self._flattened_cell(data, "MONDAY", "09:00-10:00"), "Carpentry"
        )
        self.assertEqual(
            self._flattened_cell(data, "TUESDAY", "09:00-10:00"), "Plumbing"
        )
        self.assertEqual(
            self._flattened_cell(data, "WEDNESDAY", "09:00-10:00"), "Welding"
        )

    def test_same_workshop_week_runs_are_distinguished(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="TUESDAY",
            time_period="MORNING",
            venue="",
            week_start=1,
            week_end=6,
        )
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="TUESDAY",
            time_period="MORNING",
            venue="",
            week_start=8,
            week_end=13,
        )
        data, _, _ = build_grid(
            collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        )
        cell = self._flattened_cell(data, "TUESDAY", "09:00-10:00")
        self.assertIn("Building \u00b7 Wk 1-6", cell)
        self.assertIn("Building \u00b7 Wk 8-13", cell)

    def test_morning_workshop_covers_09_to_1255(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="MONDAY",
            time_period="MORNING",
            venue="",
        )
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        self.assertEqual(entries[0]["hours"], {9, 10, 11, 12})
        grid = build_time_day_grid(entries)
        self.assertEqual(grid["slots"][0]["label"], "09:00-10:00")
        cell = grid["rows"][0]["cols"][0]
        self.assertEqual(cell["rowspan"], 4)

    def test_afternoon_workshop_covers_15_to_1855(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Electrical",
            workshop="Electrical",
            group_code="C2",
            day="MONDAY",
            time_period="AFTERNOON",
            venue="",
        )
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        self.assertEqual(entries[0]["hours"], {15, 16, 17, 18})
        grid = build_time_day_grid(entries)
        self.assertEqual(grid["slots"][0]["label"], "15:00-16:00")
        self.assertEqual(grid["rows"][0]["cols"][0]["rowspan"], 4)

    def test_same_day_different_time_workshops_are_separate_sessions(self):
        self._flat_workshop("Building")  # group C2, MONDAY morning
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Carpentry",
            group_code="C2",
            day="MONDAY",
            start_time="15:00",
            end_time="19:00",
            venue="Carpentry",
        )
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        labels = sorted(e["label"] for e in entries)
        self.assertEqual(labels, ["Building", "Carpentry"])
        # A shared name across week runs is still one workshop per day.
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="THURSDAY",
            time_period="MORNING",
            venue="",
            week_start=1,
            week_end=6,
        )
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="THURSDAY",
            time_period="MORNING",
            venue="",
            week_start=8,
            week_end=13,
        )
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        thu = [e for e in entries if e["day"] == "THURSDAY"]
        self.assertEqual(sorted(e["label"] for e in thu), ["Building", "Building"])
        self.assertEqual(sorted(e["note"] for e in thu), ["Wk 1-6", "Wk 8-13"])
        labels = [e["label"] for e in entries]  # Monday two + two Thursday runs
        self.assertEqual(labels.count("Building"), 3)

    def test_all_groups_cells_attribute_group_codes(self):
        self._flat_workshop("Building", group="C1", day="TUESDAY")
        self._flat_workshop("Building", group="C2", day="TUESDAY")
        data, _, _ = build_grid(
            collect_entries(self.prog_ee, self.sem), show_groups=True
        )
        cell = self._flattened_cell(data, "TUESDAY", "09:00-10:00")
        self.assertEqual(cell, "Building \u00b7 C1\n\nBuilding \u00b7 C2")
        # In a single-group export neither group code is repeated.
        data, _, _ = build_grid(
            collect_entries(self.prog_ee, self.sem, group=self.ee_c1)
        )
        cell = self._flattened_cell(data, "TUESDAY", "09:00-10:00")
        self.assertEqual(cell, "Building")

    def test_all_groups_never_merge_unrelated_group_sessions(self):
        self._flat_workshop("Building", group="C1")
        self._flat_workshop("Electronics", group="C2")
        cell_text = time_day_grid_to_table(
            build_time_day_grid(collect_entries(self.prog_ee, self.sem)),
            show_groups=True,
        )[0]
        mon = [
            row[1]
            for row in cell_text[1:]
            if row[1]
        ]
        joined = "\n".join(mon)
        self.assertIn("Building \u00b7 C1", joined)
        self.assertIn("Electronics \u00b7 C2", joined)

    def test_year_of_study_filters_workshops(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="MONDAY",
            time_period="MORNING",
            venue="",
            year_of_study=1,
        )
        self._flat_workshop("Plumbing", group="C2", day="TUESDAY")  # no year
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2, year=1)
        self.assertEqual({e["label"] for e in entries}, {"Building", "Plumbing"})
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2, year=2)
        self.assertEqual({e["label"] for e in entries}, {"Plumbing"})

    def test_on_screen_all_groups_shows_workshop_group(self):
        self._flat_workshop("Building", group="C1")
        self._flat_workshop("Electronics", group="C2")
        resp = self.client.get(
            "/timetable/",
            {
                "programme": self.prog_ee.pk,
                "semester": self.sem.pk,
                "year": "1",
            },
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Assigned groups: C1", html)
        self.assertIn("Assigned groups: C2", html)

    def test_workshop_names_simplified_across_programmes(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Electronics",
            workshop="Electronics",
            group_code="C1",
            day="MONDAY",
            start_time="08:00",
            end_time="12:00",
            venue="",
        )
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            group_code="B1",
            day="MONDAY",
            start_time="08:00",
            end_time="12:00",
            venue="Building",
        )
        self.assertEqual(
            {e["label"] for e in collect_entries(self.prog_ee, self.sem)},
            {"Electronics"},
        )
        self.assertEqual(
            {e["label"] for e in collect_entries(self.prog_me, self.sem)},
            {"Building"},
        )

    def test_group_collection_includes_only_its_own_workshops(self):
        self._flat_workshop("Building")  # group C2
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Electronics",
            group_code="C1",
            day="MONDAY",
            start_time="08:00",
            end_time="12:00",
            venue="Electronics",
        )
        c1_labels = [e["label"] for e in collect_entries(
            self.prog_ee, self.sem, group=self.ee_c1
        )]
        c2_labels = [e["label"] for e in collect_entries(
            self.prog_ee, self.sem, group=self.ee_c2
        )]
        self.assertEqual(c1_labels, ["Electronics"])
        self.assertEqual(c2_labels, ["Building"])
        self.assertNotIn("Building", c1_labels)
        self.assertNotIn("Electronics", c2_labels)

    def test_on_screen_workshop_card_renders_name_once(self):
        self._flat_workshop("Building")  # venue equals the workshop name
        resp = self.client.get(
            "/timetable/",
            {
                "programme": self.prog_ee.pk,
                "group": self.ee_c2.pk,
                "semester": self.sem.pk,
                "year": "1",
            },
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertEqual(html.count("Course: Building"), 1)
        self.assertIn("Venue: Building", html)
        self.assertIn("Assigned groups: C2", html)
        self.assertNotIn("Building Building", html)

    def test_group_pdf_renders_with_simplified_workshop_cell(self):
        self._flat_workshop("Building")
        self._flat_workshop("Electronics")
        from io import BytesIO

        buf = BytesIO()
        render_group_timetable(self.ee_c2, self.sem, 1, out=buf)
        pdf = buf.getvalue()
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertGreater(len(pdf), 1000)


class WorkshopRotationTests(TestCase):
    """Rotating workshops merge into one cell plus a Workshop Rotation Key.

    A slot (group + day + time period) holding two or more different workshop
    identities is a weekly rotation: the shared cell lists all names and the
    key maps each contiguous week block to a workshop. Different-time sessions
    on the same day stay separate. Programme affinity (workshop name in the
    programme name) orders the names without hard-coding.
    """

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog_ee = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        self.prog_cpe = Programme.objects.create(
            code="CPE", name="BSc. in Chemical and Processing Engineering"
        )
        self.ee_c1 = StudentGroup.objects.create(programme=self.prog_ee, code="C1")
        self.ee_c2 = StudentGroup.objects.create(programme=self.prog_ee, code="C2")
        self.cpe_b1 = StudentGroup.objects.create(programme=self.prog_cpe, code="B1")

    def _workshop(self, name, group, day="THURSDAY", period="MORNING", course=None):
        return WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code=course or name,
            workshop=name,
            group_code=group,
            day=day,
            time_period=period,
            venue="",
        )

    def test_rotating_workshops_merge_into_one_entry(self):
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c1)
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["label"], "Electrical / Carpentry")
        self.assertEqual(entries[0]["note"], "Wk 1-7 / Wk 8-14")

    def test_rotation_key_row_describes_the_slot(self):
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        rows = collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c1)
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["programme"], self.prog_ee.name)
        self.assertEqual(row["group"], "C1")
        self.assertEqual(row["day"], "THURSDAY")
        self.assertEqual(
            row["blocks"], [(1, 7, "Electrical"), (8, 14, "Carpentry")]
        )

    def test_non_rotating_workshop_has_no_rotation_row(self):
        self._workshop("Building", "C2")
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)
        self.assertEqual(entries[0]["label"], "Building")
        self.assertEqual(
            collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c2),
            [],
        )

    def test_three_way_rotation_merges_all_names(self):
        self._workshop("Welding", "C1")
        self._workshop("Electrical", "C1")
        self._workshop("M/Tools", "C1")
        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c1)
        self.assertEqual(
            entries[0]["label"], "Electrical / M/Tools / Welding"
        )
        row = collect_workshop_rotations(self.prog_ee, self.sem)[0]
        self.assertEqual(
            row["blocks"],
            [
                (1, 7, "Electrical"),
                (8, 14, "M/Tools"),
                (15, 21, "Welding"),
            ],
        )

    def test_rotations_are_isolated_per_programme(self):
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        self._workshop("Electronics", "B1")
        self._workshop("CPE", "B1")
        self.assertEqual(len(collect_workshop_rotations(self.prog_ee, self.sem)), 1)
        cpe_rows = collect_workshop_rotations(self.prog_cpe, self.sem)
        self.assertEqual(len(cpe_rows), 1)
        self.assertEqual(cpe_rows[0]["group"], "B1")
        self.assertEqual(
            cpe_rows[0]["blocks"], [(1, 7, "CPE"), (8, 14, "Electronics")]
        )

    def test_rotation_mapping_is_group_specific(self):
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        self._workshop("Building", "C2")
        self._workshop("Electronics", "C2")
        ee_rows = collect_workshop_rotations(self.prog_ee, self.sem)
        c1_row = [r for r in ee_rows if r["group"] == "C1"]
        c2_row = [r for r in ee_rows if r["group"] == "C2"]
        self.assertEqual(len(c1_row), 1)
        self.assertEqual(len(c2_row), 1)
        self.assertEqual(c1_row[0]["blocks"][0][2], "Electrical")
        self.assertEqual(c2_row[0]["blocks"][0][2], "Building")

    def test_shared_course_code_appears_in_the_key(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="WK-101",
            workshop="Carpentry",
            group_code="C1",
            day="THURSDAY",
            time_period="MORNING",
            venue="",
        )
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="WK-101",
            workshop="Electrical",
            group_code="C1",
            day="THURSDAY",
            time_period="MORNING",
            venue="",
        )
        row = collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c1)[0]
        self.assertEqual(row["course"], "WK-101")

    def test_explicit_week_blocks_are_honoured(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="MONDAY",
            time_period="MORNING",
            venue="",
            week_start=1,
            week_end=7,
        )
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Welding",
            workshop="Welding",
            group_code="C2",
            day="MONDAY",
            time_period="MORNING",
            venue="",
            week_start=17,
            week_end=18,
        )
        row = collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c2)[0]
        self.assertEqual(
            row["blocks"], [(1, 7, "Building"), (17, 18, "Welding")]
        )

    def test_empty_programme_has_no_rotation_rows(self):
        self.assertEqual(collect_workshop_rotations(self.prog_cpe, self.sem), [])

    def test_full_day_pdf_grid_starts_at_seven(self):
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        data, _, _ = build_grid(collect_entries(self.prog_ee, self.sem))
        self.assertEqual(data[0][0], "TIME")
        self.assertEqual(data[1][0], "07:00-08:00")
        self.assertEqual(data[-1][0], "19:00-20:00")
        self.assertEqual(len(data) - 1, 13)

    def test_late_evening_session_still_renders_on_full_grid(self):
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="Building",
            workshop="Building",
            group_code="C2",
            day="MONDAY",
            start_time="18:00",
            end_time="21:00",
            venue="",
        )
        data, _, _ = build_grid(collect_entries(self.prog_ee, self.sem))
        header = data[0]
        di = header.index("MONDAY")
        times = [row[0] for row in data[1:]]
        self.assertIn("07:00-08:00", times)
        self.assertIn("19:00-20:00", times)
        self.assertIn("20:00-21:00", times)
        mon = [row[di] for row in data[1:] if row[di]]
        self.assertTrue(any("Building" in cell for cell in mon))

    def test_pdf_contains_rotation_key_table(self):
        from io import BytesIO

        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        buf = BytesIO()
        render_programme_timetable(self.prog_ee, self.sem, 1, out=buf)
        pdf = buf.getvalue()
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertGreater(len(pdf), 1000)


class GroupTimetablePageTests(TestCase):
    """The on-screen timetable filters strictly to the selected student group."""

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        self.c1 = StudentGroup.objects.create(programme=self.prog, code="C1")
        self.c2 = StudentGroup.objects.create(programme=self.prog, code="C2")
        venue = Venue.objects.create(name="YOMBO5", capacity=80)
        c1_lecture = Session.objects.create(
            semester=self.sem,
            course_code="MT171",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="09:55",
            venue=venue,
        )
        SessionGroup.objects.create(session=c1_lecture, group=self.c1)
        c2_lecture = Session.objects.create(
            semester=self.sem,
            course_code="EE131",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="10:00",
            end_time="12:00",
            venue=venue,
        )
        SessionGroup.objects.create(session=c2_lecture, group=self.c2)
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="EE151",
            group_code="C1",
            day="TUESDAY",
            start_time="14:00",
            end_time="16:00",
            venue="WORKSHOP 1",
        )
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="EE152",
            group_code="C2",
            day="TUESDAY",
            start_time="14:00",
            end_time="16:00",
            venue="WORKSHOP 2",
        )
        TechnicalDrawingAllocation.objects.create(
            semester=self.sem,
            course_code="EE153",
            group_code="C1",
            day="WEDNESDAY",
            start_time="13:00",
            end_time="17:00",
            venue="TD LAB",
        )

    def _get(self, group):
        return self.client.get(
            "/timetable/",
            {
                "programme": self.prog.pk,
                "group": group.pk,
                "semester": self.sem.pk,
                "year": "1",
            },
            HTTP_HOST="localhost",
        )

    def test_selected_group_renders_without_error(self):
        resp = self._get(self.c1)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("GROUP C1", html)
        self.assertIn("MT171", html)
        self.assertIn("08:00", html)

    def test_selected_group_does_not_show_other_group_entries(self):
        c1_page = self._get(self.c1).content.decode()
        self.assertNotIn("EE131", c1_page)
        self.assertNotIn("EE152", c1_page)
        c2_page = self._get(self.c2).content.decode()
        self.assertNotIn("MT171", c2_page)
        self.assertNotIn("EE151", c2_page)
        self.assertNotIn("EE153", c2_page)

    def test_selected_group_shows_its_own_workshops_and_td(self):
        c1_page = self._get(self.c1).content.decode()
        self.assertIn("EE151", c1_page)
        self.assertIn("EE153", c1_page)
        c2_page = self._get(self.c2).content.decode()
        self.assertIn("EE152", c2_page)

    def test_group_grid_merges_multi_slot_entries(self):
        c1_page = self._get(self.c1).content.decode()
        self.assertIn('colspan="2"', c1_page)
        self.assertIn("09:00", c1_page)


class MasterTimetableExportTests(TestCase):
    """All-Programs export: merged blocks, ALL lectures, heading on every page.

    Covers the master renderer only. The individual programme/group exports are
    covered by ``TimetablePageTests`` and ``WorkshopRotationTests`` and must
    keep behaving exactly as before.

    The grid is masonry: each hour is its own column, and a session's top is the
    lowest edge any overlapping session has already reached in that column range.
    This class deliberately tests the properties that architecture is supposed to
    guarantee: the university heading is repeated on every page, an unused hour
    stays blank instead of being boxed, blocks never overlap, and splitting a
    page never drops a session.
    """

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/27", semester=1)
        self.ee = Programme.objects.create(code="EE", name="Electrical Engineering")
        self.ce = Programme.objects.create(code="CE", name="Civil Engineering")
        self.ee_c1 = StudentGroup.objects.create(programme=self.ee, code="C1")
        self.ee_c2 = StudentGroup.objects.create(programme=self.ee, code="C2")
        self.ce_a1 = StudentGroup.objects.create(programme=self.ce, code="A1")
        self.ce_a2 = StudentGroup.objects.create(programme=self.ce, code="A2")
        self.all_groups = {"C1", "C2", "A1", "A2"}
        self.venue = Venue.objects.create(name="R217", capacity=120)

    # -- helpers ---------------------------------------------------------
    def _session(self, course, day, start, end, groups=(), venue=True,
                 activity_type="LECTURE"):
        session = Session.objects.create(
            semester=self.sem,
            course_code=course,
            activity_type=activity_type,
            day=day,
            start_time=start,
            end_time=end,
            venue=self.venue if venue else None,
        )
        for group in groups:
            SessionGroup.objects.create(session=session, group=group)
        return session

    def _td(self, course, group_code, day, start, end, venue="S112"):
        return TechnicalDrawingAllocation.objects.create(
            semester=self.sem, course_code=course, group_code=group_code,
            day=day, start_time=start, end_time=end, venue=venue,
        )

    def _workshop(self, course, group_code, day, start, end, venue="Workshop Shed"):
        return WorkshopAllocation.objects.create(
            semester=self.sem, course_code=course, group_code=group_code,
            day=day, start_time=start, end_time=end, venue=venue, workshop=course,
        )

    def _merged(self, entries=None):
        if entries is None:
            entries = collect_master_entries(self.sem)
        return _merge_master_entries(entries, self.all_groups)

    def _blocks(self, kind):
        return [e for e in self._merged() if e["kind"] == kind]

    def _page_streams(self, pdf):
        """Decode each PDF page's content stream without needing pypdf."""
        import base64
        import zlib

        streams = []
        for match in re.finditer(rb"stream\r?\n", pdf):
            start = match.end()
            end = pdf.find(b"endstream", start)
            if end == -1:
                continue
            raw = pdf[start:end].strip()
            for decode in (
                lambda b: zlib.decompress(b),
                lambda b: zlib.decompress(base64.a85decode(b, adobe=True)),
                lambda b: b,
            ):
                try:
                    streams.append(decode(raw))
                    break
                except Exception:
                    continue
        return streams

    def _render(self):
        """Render the master PDF and count the pages reportlab produced."""
        from io import BytesIO

        import core.timetable_pdf as tp

        calls = []
        original = tp._draw_master_header

        def counting(canvas, doc, heading_lines, *args):
            calls.append(tuple(heading_lines))
            return original(canvas, doc, heading_lines, *args)

        buf = BytesIO()
        with mock.patch.object(tp, "_draw_master_header", counting):
            render_udsm_master_timetable(
                collect_master_entries(self.sem), self.sem, 1, out=buf
            )
        pdf = buf.getvalue()
        pages = len(re.findall(rb"/Type\s*/Page[^s]", pdf))
        return pdf, pages, calls

    def _busy_week(self):
        """Enough overlapping sessions to force the week onto several pages.

        Masonry grows a day downwards only where its sessions overlap in time.
        Twelve simultaneous sessions a day gives twelve stacked blocks, and five
        days of that cannot fit one A3 landscape page.
        """
        for day in ("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"):
            for i in range(12):
                self._session(f"MT{100 + i}", day, "08:00", "09:55")

    # -- 1. lectures: ALL instead of a long group list --------------------
    def test_lecture_for_every_group_shows_all(self):
        self._session(
            "QS125", "MONDAY", "09:00", "10:55",
            groups=[self.ee_c1, self.ee_c2, self.ce_a1, self.ce_a2],
        )
        lectures = self._blocks("lecture")
        self.assertEqual(len(lectures), 1)
        self.assertEqual(lectures[0]["groups"], "ALL")

    def test_lecture_block_prints_the_all_marker(self):
        """A lecture for every group prints "ALL" as its groups line.

        The block must not spell out the twenty-odd group codes, and it must not
        drop the group line either: "ALL" is the only thing telling the reader
        the session is open to every programme.
        """
        self._session(
            "QS125", "MONDAY", "09:00", "10:55",
            groups=[self.ee_c1, self.ee_c2, self.ce_a1, self.ce_a2],
        )
        lines = _master_cell_text(self._blocks("lecture"), show_groups=True).split("\n")
        self.assertEqual(lines[-1], "ALL")
        for code in ("C1", "C2", "A1", "A2"):
            self.assertNotIn(code, lines)

    def test_partial_lecture_block_lists_only_its_groups(self):
        self._session(
            "QS125", "MONDAY", "09:00", "10:55", groups=[self.ee_c1, self.ee_c2],
        )
        lines = _master_cell_text(self._blocks("lecture"), show_groups=True).split("\n")
        self.assertEqual(lines[-1], "C1, C2")
        self.assertNotIn("ALL", "\n".join(lines))

    def test_lecture_for_some_groups_keeps_them(self):
        self._session("TX101", "FRIDAY", "10:00", "10:55", groups=[self.ee_c1])
        lecture = self._blocks("lecture")[0]
        self.assertEqual(lecture["groups"], "C1")
        self.assertEqual(
            _master_cell_text([lecture], show_groups=True).split("\n")[-1], "C1"
        )

    def test_unassigned_lecture_lists_no_groups(self):
        self._session("AR131", "FRIDAY", "07:00", "07:55")
        lecture = self._blocks("lecture")[0]
        self.assertEqual(lecture["groups"], "")
        self.assertNotIn("ALL", _master_cell_text([lecture], show_groups=True))

    def test_distinct_lectures_at_one_slot_are_all_kept(self):
        self._session("AR111", "MONDAY", "08:00", "10:55")
        self._session("TR111", "MONDAY", "08:00", "10:55")
        self._session("SC121", "MONDAY", "08:00", "10:55")
        codes = sorted(e["course_code"] for e in self._blocks("lecture"))
        self.assertEqual(codes, ["AR111", "SC121", "TR111"])

    def test_identical_duplicate_lectures_collapse(self):
        for _ in range(3):
            self._session("AR111", "MONDAY", "08:00", "10:55")
        self.assertEqual(len(self._blocks("lecture")), 1)

    # -- 2. TD sessions merge -------------------------------------------
    def test_same_slot_td_sessions_merge_into_one_block(self):
        for group_code in ("EE C1", "EE C2", "CE A1"):
            self._td("ME101", group_code, "MONDAY", "09:00", "12:00")
        blocks = self._blocks("td")
        self.assertEqual(len(blocks), 1, "three TD sessions must share one block")
        self.assertEqual(blocks[0]["type_label"], "Technical Drawing")
        self.assertEqual(blocks[0]["groups"], "EE C1, C2, CE A1")
        self.assertEqual(blocks[0]["venue"], "S112")

    def test_td_at_different_times_stay_separate(self):
        for start, end in (("09:00", "12:00"), ("15:00", "18:00")):
            self._td("ME101", "A1", "MONDAY", start, end)
        self.assertEqual(len(self._blocks("td")), 2)

    def test_td_block_shows_type_and_groups_only(self):
        self._td("ME101", "A1", "MONDAY", "09:00", "12:00")
        text = _master_cell_text(self._blocks("td"), show_groups=True)
        self.assertEqual(text.split("\n")[0], "Technical Drawing")
        self.assertNotIn("ME101", text)

    def test_td_block_uses_the_full_technical_drawing_name(self):
        """The block reads "Technical Drawing", never the "TD" abbreviation."""
        self._td("ME101", "EE C1", "MONDAY", "09:00", "12:00")
        block = self._blocks("td")[0]
        self.assertEqual(block["type_label"], "Technical Drawing")
        text = _master_cell_text([block], show_groups=True)
        self.assertEqual(text.split("\n")[0], "Technical Drawing")
        self.assertNotIn("TD", text)

    def test_group_codes_drop_a_repeated_programme_prefix(self):
        self.assertEqual(
            _compact_group_codes(["EE C1", "EE C2", "CE A1"]), "EE C1, C2, CE A1"
        )
        self.assertEqual(_compact_group_codes(["C1", "C2", "A1"]), "C1, C2, A1")
        self.assertEqual(_compact_group_codes(["A1"]), "A1")

    # -- 2b. whole-cohort lectures always read ALL ------------------------
    def test_whole_cohort_lectures_read_all_not_their_linked_groups(self):
        """CL111, MT161, MT171, ME101 and every DS lecture are open to all.

        Their records may name only some of the groups that actually attend, so
        the block must read "ALL" rather than understating who has to be there.
        """
        for course in ("CL111", "MT161", "MT171", "ME101", "DS114", "DS176"):
            with self.subTest(course=course):
                self._session(
                    course, "MONDAY", "08:00", "09:55",
                    groups=[self.ee_c1, self.ee_c2],
                )
                block = [
                    e for e in self._merged() if e["course_code"] == course
                ][0]
                self.assertEqual(block["groups"], "ALL", course)
                text = _master_cell_text([block], show_groups=True)
                self.assertEqual(text.split("\n")[-1], "ALL", course)
                for code in ("C1", "C2"):
                    self.assertNotIn(code, text, course)

    def test_whole_cohort_rule_applies_to_lectures_only(self):
        """A per-group tutorial of the same course still shows its groups."""
        self._session(
            "CL111", "TUESDAY", "10:00", "10:55",
            groups=[self.ee_c1], activity_type="TUTORIAL",
        )
        block = [
            e for e in self._merged() if e["course_code"] == "CL111"
        ][0]
        self.assertEqual(block["groups"], "C1")

    def test_other_lectures_still_list_their_groups(self):
        self._session(
            "QS125", "MONDAY", "09:00", "10:55", groups=[self.ee_c1, self.ee_c2],
        )
        block = self._blocks("lecture")[0]
        self.assertEqual(block["groups"], "C1, C2")

    def test_ds_prefix_matches_any_development_studies_course(self):
        for course in ("DS101", "DS150", "DS2", "ds114"):
            with self.subTest(course=course):
                self._session(course, "WEDNESDAY", "08:00", "08:55",
                              groups=[self.ee_c1])
                block = [
                    e for e in self._merged() if e["course_code"] == course
                ][0]
                self.assertEqual(block["groups"], "ALL", course)

    def test_course_that_merely_contains_ds_is_not_whole_cohort(self):
        """Only a DS *prefix* counts; ADS101 or MTDS2 are ordinary lectures."""
        for course in ("ADS101", "MTDS2"):
            with self.subTest(course=course):
                self._session(course, "THURSDAY", "08:00", "08:55",
                              groups=[self.ee_c1])
                block = [
                    e for e in self._merged() if e["course_code"] == course
                ][0]
                self.assertEqual(block["groups"], "C1", course)

    # -- 3. workshops merge and simplify --------------------------------
    def test_same_slot_workshops_merge_and_hide_names(self):
        for group_code, workshop in (
            ("C1", "Carpentry"), ("C2", "Welding"), ("A1", "Masonry"),
        ):
            self._workshop(workshop, group_code, "MONDAY", "09:00", "13:00")
        blocks = self._blocks("workshop")
        self.assertEqual(len(blocks), 1, "three workshops must share one block")
        self.assertEqual(blocks[0]["type_label"], "Workshop")
        self.assertEqual(blocks[0]["groups"], "C1, C2, A1")
        text = _master_cell_text(blocks, show_groups=True)
        self.assertEqual(text.split("\n"), ["Workshop", "C1, C2, A1"])
        for detail in ("Carpentry", "Welding", "Masonry", "Workshop Shed"):
            self.assertNotIn(detail, text)

    def test_workshops_at_different_times_stay_separate(self):
        for start, end in (("09:00", "13:00"), ("15:00", "19:00")):
            self._workshop("Carpentry", "C1", "MONDAY", start, end)
        self.assertEqual(len(self._blocks("workshop")), 2)

    # -- 4. cell lines, wrapping, and the three-line floor ----------------
    def test_entry_lines_keep_every_field_on_its_own_line(self):
        # EE150 is not a whole-cohort course, so this block carries no groups
        # line; the "ALL" rule is covered by the lecture tests above.
        self._session("EE150", "MONDAY", "08:00", "09:55")
        entry = self._blocks("lecture")[0]
        self.assertEqual(
            _entry_lines(entry),
            ["Lecture", "08:00–09:55", "R217", "EE150"],
        )

    def test_entry_lines_really_split_on_newlines(self):
        """Each field is drawn separately, so they cannot run together."""
        self._session("EE150", "MONDAY", "08:00", "09:55")
        entry = self._blocks("lecture")[0]
        for field in ("Lecture", "08:00–09:55", "R217", "EE150"):
            self.assertIn(field, _entry_lines(entry))

    def test_boxes_are_padded_to_a_three_line_floor(self):
        """A short block still occupies three lines, so blocks stay uniform."""
        self._session("CL111", "MONDAY", "08:00", "08:55", venue=False)
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46
        )
        boxes = flowable.boxes
        height = flowable.height
        self.assertEqual(len(boxes), 1)
        self.assertGreaterEqual(len(boxes[0]["lines"]), 3)
        self.assertEqual(
            boxes[0]["bottom"] - boxes[0]["top"],
            len(boxes[0]["lines"]) * _WeekMasonryFlowable.LEADING
            + 2 * _WeekMasonryFlowable.PAD_Y,
        )
        self.assertAlmostEqual(height, boxes[0]["bottom"] - boxes[0]["top"])

    def test_longer_content_is_never_truncated_to_the_floor(self):
        self._session("CL111", "MONDAY", "08:00", "09:55", venue=False)
        merged = self._merged()
        boxes, _height = _masonry_day_boxes(
            [e for e in merged if e["day"] == "MONDAY"],
            list(range(7, 20)),
            col_width=60.0,
        )
        # Content is padded to minimum 3 lines for uniform appearance
        self.assertGreaterEqual(len(boxes[0]["lines"]), 3)

    def test_long_text_wraps_instead_of_overflowing_its_box(self):
        line = "R&D <lab> " + "VeryLongVenueName" * 6
        width = 40.0
        pieces = _wrap_line(line, width, "Helvetica", 7)
        from reportlab.pdfbase import pdfmetrics

        self.assertGreater(len(pieces), 1)
        for piece in pieces:
            self.assertLessEqual(
                pdfmetrics.stringWidth(piece, "Helvetica", 7), width
            )

    def test_short_text_is_left_as_one_line(self):
        self.assertEqual(_wrap_line("R217", 200.0, "Helvetica", 7), ["R217"])
        self.assertEqual(_wrap_line("", 200.0, "Helvetica", 7), [""])
        self.assertEqual(_wrap_line(None, 200.0, "Helvetica", 7), [""])

    # -- 4b. the masonry layout: independent cursors per time column ------
    def test_overlapping_blocks_stack_instead_of_sharing_a_row(self):
        """The whole point of masonry: no boxed-out empty rows.

        Three simultaneous sessions must produce three boxes stacked on top of
        one another, not one row stretched to the tallest of them.
        """
        for i in range(3):
            self._session(f"MT{600 + i}", "MONDAY", "08:00", "08:55")
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46
        )
        boxes = flowable.boxes
        height = flowable.height
        self.assertEqual(len(boxes), 3)
        # Every box occupies the same single hour column.
        self.assertEqual({box["colspan"] for box in boxes}, {1})
        # They are stacked: each starts where the previous one ended.
        for previous, current in zip(boxes, boxes[1:]):
            self.assertAlmostEqual(current["top"], previous["bottom"])
        # And the day is only as tall as the stack, not three times a row.
        self.assertAlmostEqual(height, boxes[-1]["bottom"])

    def test_sequential_blocks_do_not_stack(self):
        """Different hours are different columns, so they sit side by side."""
        self._session("MT601", "MONDAY", "08:00", "08:55")
        self._session("MT602", "MONDAY", "10:00", "10:55")
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46
        )
        boxes = flowable.boxes
        height = flowable.height
        self.assertEqual(len(boxes), 2)
        self.assertNotEqual(boxes[0]["col"], boxes[1]["col"])
        # Both start at the top: no empty row is drawn between them.
        self.assertAlmostEqual(boxes[0]["top"], 0.0)
        self.assertAlmostEqual(boxes[1]["top"], 0.0)
        self.assertAlmostEqual(height, boxes[0]["bottom"])

    def test_a_wide_block_starts_below_the_tallest_column_it_covers(self):
        """A block's top is the max cursor of the columns it spans."""
        # Two one-hour blocks in hour 8 stack; a two-hour block covering hours
        # 8 and 9 must start below both.
        for i in range(2):
            self._session(f"MT{610 + i}", "MONDAY", "08:00", "08:55")
        self._session("MT620", "MONDAY", "08:00", "09:55")
        merged = self._merged()
        boxes, _height = _masonry_day_boxes(
            [e for e in merged if e["day"] == "MONDAY"],
            list(range(7, 20)),
            col_width=60.0,
        )
        wide = [b for b in boxes if b["colspan"] == 2][0]
        narrow = [b for b in boxes if b["colspan"] == 1]
        self.assertAlmostEqual(wide["top"], max(b["bottom"] for b in narrow))

    def test_boxes_never_overlap_within_a_day(self):
        """No two boxes in a day may share space, whatever the data."""
        for hour in (8, 9, 10, 11):
            for i in range(4):
                self._session(
                    f"MT{700 + hour}{i}", "MONDAY",
                    f"{hour:02d}:00", f"{hour:02d}:55",
                )
        self._session("MT800", "MONDAY", "08:00", "11:55")
        merged = self._merged()
        boxes, _height = _masonry_day_boxes(
            [e for e in merged if e["day"] == "MONDAY"],
            list(range(7, 20)),
            col_width=60.0,
        )
        for i, first in enumerate(boxes):
            for second in boxes[i + 1:]:
                same_columns = (
                    first["col"] < second["col"] + second["colspan"]
                    and second["col"] < first["col"] + first["colspan"]
                )
                if not same_columns:
                    continue
                self.assertFalse(
                    first["top"] < second["bottom"] and second["top"] < first["bottom"],
                    f"boxes overlap: {first} and {second}",
                )

    def test_empty_time_is_left_blank_rather_than_drawn(self):
        """Nothing is emitted for an hour with no session at all."""
        self._session("MT601", "MONDAY", "14:00", "14:55")
        merged = self._merged()
        boxes, _height = _masonry_day_boxes(
            [e for e in merged if e["day"] == "MONDAY"],
            list(range(7, 20)),
            col_width=60.0,
        )
        self.assertEqual(len(boxes), 1)
        # One box for one hour: the twelve other printed hours draw nothing.
        self.assertEqual(boxes[0]["colspan"], 1)

    def test_a_session_outside_the_printed_hours_is_skipped_not_misplaced(self):
        self._session("MT601", "MONDAY", "05:00", "05:55")
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46
        )
        boxes = flowable.boxes
        height = flowable.height
        self.assertEqual(boxes, [])
        self.assertEqual(height, 0.0)

    # -- 4c. bands, days, and page breaks --------------------------------
    def test_bands_appear_in_week_order_and_carry_day_labels(self):
        for day in ("FRIDAY", "MONDAY", "WEDNESDAY"):
            self._session(f"MT{800 + len(day)}", day, "08:00", "08:55")
        merged = self._merged()
        day_flowables, slots = _build_day_flowables(merged, col_width=60.0, day_width=46)
        self.assertEqual(
            [flowable.day_label for flowable in day_flowables],
            ["Monday", "Wednesday", "Friday"],
        )
        self.assertEqual(slots, list(range(7, 20)))

    def test_a_day_with_no_sessions_gets_no_band(self):
        self._session("MT601", "MONDAY", "08:00", "08:55")
        merged = self._merged()
        bands, _slots = _build_week_masonry(merged, col_width=60.0, day_width=46)
        self.assertEqual([flowable.day_label for flowable in day_flowables], ["Monday"])

    def test_split_breaks_between_days(self):
        """Days continue into one another; only a full page forces a break."""
        for day in ("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"):
            self._session(f"MT{810 + len(day)}", day, "08:00", "08:55")
        merged = self._merged()
        day_flowables, slots = _build_day_flowables(merged, col_width=60.0, day_width=46)
        # With day-by-day flowables, days don't split - they move to next page
        # Test that each day can be placed independently
        for flowable in day_flowables:
            whole = flowable.split(1000, 5000)
            self.assertEqual(len(whole), 1)
        self.assertEqual(len(whole), 1)
        self.assertEqual(len(whole[0].bands), 5)
        # Room for two days only: the rest continues on the next page.
        two_days = sum(b["height"] for b in bands[:2])
        pieces = flowable.split(1000, two_days + 1)
        self.assertEqual(len(pieces), 2)
        self.assertEqual(
            [b["label"] for b in pieces[0].bands],
            ["Monday", "Tuesday"],
        )
        self.assertEqual(
            [b["label"] for b in pieces[1].bands],
            ["Wednesday", "Thursday", "Friday"],
        )

    def test_split_keeps_every_block_exactly_once(self):
        """Splitting must never drop or duplicate a block."""
        for i in range(6):
            for j in range(3):
                self._session(
                    f"MT{900 + i}{j}", "MONDAY", "08:00", "08:55"
                )
            self._session(f"MT{950 + i}", "TUESDAY", "08:00", "08:55")
        merged = self._merged()
        day_flowables, slots = _build_day_flowables(merged, col_width=60.0, day_width=46)
        # With day-by-day, each day stays intact
        for flowable in day_flowables:
            before = len(flowable.boxes)
            pieces = flowable.split(1000, flowable.height + 1)
            self.assertEqual(len(pieces), 1)
            after = len(pieces[0].boxes)
            self.assertEqual(before, after)

    def test_a_day_taller_than_a_page_moves_to_next_page(self):
        """A day that doesn't fit moves to the next page instead of being cut."""
        for i in range(60):
            self._session(f"MT{1000 + i}", "MONDAY", "08:00", "08:55")
        merged = self._merged()
        day_flowables, slots = _build_day_flowables(merged, col_width=60.0, day_width=46)
        self.assertEqual(len(day_flowables), 1)
        monday = day_flowables[0]
        self.assertGreater(monday.height, 700.0)
        # A day that doesn't fit should return empty split (move to next page)
        pieces = monday.split(1000, 700.0)
        self.assertEqual(len(pieces), 0)
        # Every piece still says Monday, so the day names itself on each page.
        for piece in pieces:
            for band in piece.bands:
                self.assertEqual(band["label"], "Monday")
        # And the split respects the page height it was given.
        for piece in pieces:
            for band in piece.bands:
                self.assertLessEqual(band["height"], 700.0 + 0.01)
        # No block is lost or duplicated by the cut.
        self.assertEqual(
            sum(len(b["boxes"]) for p in pieces for b in p.bands),
            len(bands[0]["boxes"]),
        )

    def test_cut_band_rebases_boxes_onto_their_new_chunk(self):
        band = {
            "label": "Monday",
            "height": 300.0,
            "boxes": [
                {"col": 0, "colspan": 1, "top": 0.0, "bottom": 30.0,
                 "lines": ["a"], "fill": "#fff"},
                {"col": 0, "colspan": 1, "top": 30.0, "bottom": 60.0,
                 "lines": ["b"], "fill": "#fff"},
                {"col": 0, "colspan": 1, "top": 60.0, "bottom": 90.0,
                 "lines": ["c"], "fill": "#fff"},
            ],
        }
        chunks = _cut_band(band, 70.0)
        self.assertEqual(len(chunks), 2)
        self.assertEqual([b["lines"][0] for b in chunks[0]["boxes"]], ["a", "b"])
        self.assertEqual([b["lines"][0] for b in chunks[1]["boxes"]], ["c"])
        # Each chunk starts at zero and its height is its own stack.
        self.assertEqual(chunks[0]["boxes"][0]["top"], 0.0)
        self.assertAlmostEqual(chunks[0]["height"], 60.0)
        self.assertEqual(chunks[1]["boxes"][0]["top"], 0.0)
        self.assertAlmostEqual(chunks[1]["height"], 30.0)

    def test_the_week_flowable_reports_its_full_size(self):
        for day in ("MONDAY", "TUESDAY"):
            self._session(f"MT{1100 + len(day)}", day, "08:00", "08:55")
        merged = self._merged()
        day_flowables, slots = _build_day_flowables(merged, col_width=60.0, day_width=46)
        # Test that day flowables report their correct size
        for flowable in day_flowables:
            width, height = flowable.wrap(10000, 10000)
            self.assertAlmostEqual(width, 46 + 60 * len(slots))
            self.assertGreater(height, 0)
        self.assertAlmostEqual(height, sum(b["height"] for b in bands))

    def test_export_draws_the_day_labels_rotated(self):
        """The day column is drawn through a rotated canvas, not as text."""
        self._session("AR111", "MONDAY", "08:00", "12:55")
        self._session("TR111", "WEDNESDAY", "08:00", "12:55")
        pdf, _pages, _calls = self._render()
        content = b"\n".join(self._page_streams(pdf))
        rotations = re.findall(
            rb"0(?:\.0+)? 1(?:\.0+)? -1(?:\.0+)? 0(?:\.0+)? "
            rb"(-?\d+\.?\d*) (-?\d+\.?\d*) cm",
            content,
        )
        self.assertTrue(rotations, "no 90-degree canvas rotation in the PDF")
        # One rotation per day that has a label.
        self.assertGreaterEqual(len(rotations), 2)

    def test_each_day_name_is_written_exactly_once(self):
        for day in ("MONDAY", "TUESDAY", "WEDNESDAY"):
            self._session(f"MT{1200 + len(day)}", day, "08:00", "08:55")
        pdf, _pages, _calls = self._render()
        content = b"\n".join(self._page_streams(pdf))
        for name in (b"(Monday)", b"(Tuesday)", b"(Wednesday)"):
            self.assertEqual(content.count(name), 1, name)

    # -- 5. the heading is drawn on EVERY page ---------------------------
    def test_heading_is_repeated_on_every_page(self):
        """Regression: the heading must not vanish on page 2+.

        The title block is painted via onFirstPage/onLaterPages rather than
        added once to the element list, so it reappears on however many pages
        reportlab's own table splitting produces.
        """
        self._busy_week()
        pdf, pages, calls = self._render()
        self.assertGreater(pages, 1, "this fixture must span more than one page")
        self.assertEqual(len(calls), pages)
        for heading in calls:
            self.assertEqual(heading[0], "UNIVERSITY OF DAR ES SALAAM")
            self.assertIn("TEACHING TIMETABLE FOR FIRST SEMESTER 2026/27", heading[1])

    def test_single_page_still_gets_the_heading(self):
        self._session("AR111", "MONDAY", "08:00", "12:55")
        pdf, pages, calls = self._render()
        self.assertEqual(pages, 1)
        self.assertEqual(len(calls), 1)
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertIn(b"Master Timetable", pdf)

    def test_export_is_a_pdf_with_the_right_title(self):
        self._session("AR111", "MONDAY", "08:00", "12:55")
        pdf, _pages, _calls = self._render()
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertIn(b"Master Timetable", pdf)

    def test_year_of_study_appears_in_the_subtitle(self):
        from io import BytesIO

        import core.timetable_pdf as tp

        calls = []
        original = tp._draw_master_header
        buf = BytesIO()

        def counting(canvas, doc, heading_lines, *args):
            calls.append(tuple(heading_lines))
            return original(canvas, doc, heading_lines, *args)

        with mock.patch.object(tp, "_draw_master_header", counting):
            render_udsm_master_timetable(
                collect_master_entries(self.sem), self.sem, 2, out=buf
            )
        self.assertIn("2ND YEAR", calls[0][1])

    def test_empty_selection_renders_a_message_not_a_crash(self):
        from io import BytesIO

        buf = BytesIO()
        render_udsm_master_timetable([], self.sem, 1, out=buf)
        self.assertTrue(buf.getvalue().startswith(b"%PDF-"))

    # -- 6. merging never loses data -------------------------------------
    def test_merging_never_increases_the_block_count(self):
        for i in range(4):
            self._td("ME101", f"G{i}", "MONDAY", "09:00", "12:00")
        raw = collect_master_entries(self.sem)
        self.assertLess(len(self._merged(raw)), len(raw))

    def test_merging_never_drops_an_entry(self):
        self._session("AR111", "MONDAY", "08:00", "12:55")
        self._session("TR111", "MONDAY", "09:00", "10:55")
        self._session("AR131", "FRIDAY", "07:00", "07:55")
        raw = collect_master_entries(self.sem)
        merged = self._merged(raw)
        codes = sorted(e["course_code"] for e in merged)
        self.assertEqual(codes, ["AR111", "AR131", "TR111"])
        self.assertEqual(len(merged), len(raw))

    def test_merged_block_covers_every_hour_of_its_bucket(self):
        self._td("ME101", "A1", "MONDAY", "09:00", "12:00")
        self._td("ME101", "A2", "MONDAY", "09:00", "12:00")
        block = self._blocks("td")[0]
        self.assertEqual(block["hours"], {9, 10, 11})

    def test_session_without_an_activity_type_still_gets_a_heading(self):
        session = Session.objects.create(
            semester=self.sem, course_code="ZZ999", activity_type="",
            day="MONDAY", start_time="08:00", end_time="08:55", venue=self.venue,
        )
        self.assertTrue(session.pk)
        blocks = [
            e for e in self._merged() if e["course_code"] == "ZZ999"
        ]
        self.assertEqual(len(blocks), 1)
        self.assertIn("ZZ999", _master_cell_text(blocks, show_groups=True))

    # -- 7. the individual exports are untouched -------------------------
    def test_individual_exports_are_untouched(self):
        self._session("AR111", "MONDAY", "08:00", "12:55", groups=[self.ee_c1])
        entries = collect_entries(self.ee, self.sem)
        labels = [e["label"] for e in entries]
        self.assertEqual(len(labels), 1)
        self.assertIn("AR111 Lecture", labels[0])
        data, spans, fills = build_grid(entries)
        self.assertEqual(data[0][0], "TIME")
        self.assertEqual(data[0][1], "MONDAY")

    def test_master_export_view_returns_a_pdf(self):
        self._session("AR111", "MONDAY", "08:00", "12:55")
        response = self.client.get(
            "/export/all-programmes/timetable.pdf/",
            {"semester": self.sem.pk, "year": "1"},
            HTTP_HOST="localhost",
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "application/pdf")
        self.assertTrue(response.content.startswith(b"%PDF-"))
        self.assertIn("master_timetable_1.pdf", response["Content-Disposition"])


class TimetablePageTests(TestCase):
    """On-screen timetable page and the PDF export (classic grid)."""

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog_a = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        Programme.objects.create(code="AB", name="Alpha Programme")
        group = StudentGroup.objects.create(programme=self.prog_a, code="C1")
        venue = Venue.objects.create(name="YOMBO5", capacity=80)
        session = Session.objects.create(
            semester=self.sem,
            course_code="MT171",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="09:00",
            venue=venue,
        )
        SessionGroup.objects.create(session=session, group=group)
        # An ungrouped session (as in the real master timetable) — tutorials,
        # seminars and practicals have no group links, so only the default
        # whole-year view (master timetable as source of truth) shows them.
        Session.objects.create(
            semester=self.sem,
            course_code="ST171",
            activity_type="TUTORIAL",
            day="TUESDAY",
            start_time="10:00",
            end_time="12:00",
            venue=venue,
        )
        WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code="EE151",
            group_code="C1",
            day="WEDNESDAY",
            start_time="09:00",
            end_time="12:00",
            venue="WORKSHOP 1",
        )
        TechnicalDrawingAllocation.objects.create(
            semester=self.sem,
            course_code="EE153",
            group_code="C1",
            day="THURSDAY",
            start_time="13:00",
            end_time="17:00",
            venue="TD LAB",
        )

    def test_page_renders_day_time_grid(self):
        resp = self.client.get(
            "/timetable/",
            {"programme": self.prog_a.pk, "semester": self.sem.pk},
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("TIME", html)
        # Days run down the first column, labelled with their full weekday name.
        self.assertIn("Monday", html)
        self.assertIn("Friday", html)
        self.assertNotIn("MONDAY</td>", html)
        # Hours run across the top, ascending from 07:00, headers in the
        # master-timetable format 07:00-07:55 etc.
        self.assertIn(">07:00-07:55<", html)
        self.assertIn(">08:00-08:55<", html)
        self.assertIn(">09:00-09:55<", html)
        self.assertNotIn(">07:00-08:00<", html)
        self.assertIn("MT171", html)
        self.assertIn("YOMBO5", html)
        self.assertIn("tt-card", html)
        self.assertIn("tt-kind-lecture", html)
        self.assertIn("tt-time", html)
        # Cards live in an inner flex wrapper — the <td> itself stays a plain
        # table cell so colspans align under the right hourly column.
        self.assertIn('class="tt-block-inner"', html)
        self.assertTrue(html.count('class="tt-block"') == html.count('class="tt-block-inner"'))
        # Entry box reads like the Excel grid cells.
        self.assertIn("Lecture, 08:00-09:00, Mon", html)
        self.assertIn("Course: MT171", html)
        self.assertIn("Venue: YOMBO5", html)
        self.assertIn("Assigned groups:", html)

    def test_page_falls_back_when_programme_unknown(self):
        resp = self.client.get(
            "/timetable/", {"programme": "9999"}, HTTP_HOST="localhost"
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Timetable", resp.content.decode())

    def test_empty_timetable_shows_message(self):
        prog = Programme.objects.get(code="AB")
        resp = self.client.get(
            "/timetable/", {"programme": prog.pk}, HTTP_HOST="localhost"
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("No sessions scheduled", resp.content.decode())

    def test_default_view_shows_complete_first_year_timetable(self):
        resp = self.client.get("/timetable/", HTTP_HOST="localhost")
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("ALL PROGRAMMES", html)
        self.assertIn(">TIME<", html)
        self.assertNotIn("No programmes registered yet", html)
        self.assertIn("MT171", html)
        # The master timetable is the source of truth: ungrouped sessions
        # (tutorials/seminars/practicals) appear too, with a blank groups line.
        self.assertIn("Course: ST171", html)
        self.assertRegex(html, r"Assigned groups:\s*</div>")

    def test_all_programmes_option_shows_complete_first_year_timetable(self):
        resp = self.client.get(
            "/timetable/", {"programme": "all"}, HTTP_HOST="localhost"
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("ALL PROGRAMMES", html)
        self.assertIn(">TIME<", html)
        self.assertNotIn("No programmes registered yet", html)
        self.assertIn("MT171", html)
        self.assertIn("ST171", html)

    def test_programme_filter_hides_ungrouped_master_sessions(self):
        resp = self.client.get(
            "/timetable/",
            {"programme": self.prog_a.pk, "semester": self.sem.pk},
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("MT171", html)
        self.assertIn("ELECTRICAL ENGINEERING", html)
        # Ungrouped master-timetable sessions are not part of any single
        # programme's timetable view.
        self.assertNotIn("ST171", html)

    def test_master_entries_include_every_session_in_the_semester(self):
        entries = collect_master_entries(self.sem)
        session_codes = {
            e["name"] for e in entries if e["key"][0] == "session"
        }
        self.assertIn("MT171", session_codes)
        self.assertIn("ST171", session_codes)
        self.assertEqual(
            len(session_codes), Session.objects.filter(semester=self.sem).count()
        )
        self.assertTrue(any(e["kind"] == "workshop" for e in entries))
        self.assertTrue(any(e["kind"] == "td" for e in entries))

    def test_classic_pdf_export(self):
        resp = self.client.get(
            "/export/programmes/%d/timetable.pdf/" % self.prog_a.pk,
            {"semester": self.sem.pk, "year": "1"},
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp["Content-Type"], "application/pdf")
        self.assertTrue(resp.content.startswith(b"%PDF-"))

    def test_group_pdf_export(self):
        group = StudentGroup.objects.get(programme=self.prog_a, code="C1")
        resp = self.client.get(
            "/export/groups/%d/timetable.pdf/" % group.pk,
            {"semester": self.sem.pk, "year": "1"},
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp["Content-Type"], "application/pdf")
        self.assertTrue(resp.content.startswith(b"%PDF-"))
        self.assertIn("attachment; filename=", resp["Content-Disposition"])
        self.assertIn("C1", resp["Content-Disposition"])

    def test_group_pdf_export_falls_back_to_a_semester(self):
        group = StudentGroup.objects.get(programme=self.prog_a, code="C1")
        resp = self.client.get(
            "/export/groups/%d/timetable.pdf/" % group.pk,
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(resp.content.startswith(b"%PDF-"))
        self.assertIn(b"/Title (EE C1 Timetable)", resp.content)

    def test_group_pdf_default_skips_empty_newer_semester(self):
        empty = Semester.objects.create(academic_year="2026/2027", semester=2)
        group = StudentGroup.objects.get(programme=self.prog_a, code="C1")
        resp = self.client.get(
            "/export/groups/%d/timetable.pdf/" % group.pk,
            HTTP_HOST="localhost",
        )
        empty_resp = self.client.get(
            "/export/groups/%d/timetable.pdf/" % group.pk,
            {"semester": empty.pk},
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(resp.content.startswith(b"%PDF-"))
        self.assertIn(b"/Title (EE C1 Timetable)", resp.content)
        # The default must not land on the newer-but-empty semester.
        self.assertGreater(len(resp.content), len(empty_resp.content))

    def test_export_page_defaults_to_semester_with_data(self):
        empty = Semester.objects.create(academic_year="2026/2027", semester=2)
        resp = self.client.get("/export/", HTTP_HOST="localhost")
        html = resp.content.decode()
        self.assertIn(f'value="{self.sem.pk}" selected', html)
        self.assertNotIn(f'value="{empty.pk}" selected', html)

    def test_export_page_searches_groups(self):
        resp = self.client.get("/export/", {"q": "C1"}, HTTP_HOST="localhost")
        html = resp.content.decode()
        self.assertEqual(resp.status_code, 200)
        self.assertIn("EE", html)
        self.assertIn("C1", html)
        self.assertNotIn("Alpha Programme", html)
        self.assertIn("/export/groups/", html)

    def test_export_page_empty_search(self):
        resp = self.client.get("/export/", {"q": "zzz"}, HTTP_HOST="localhost")
        html = resp.content.decode()
        self.assertEqual(resp.status_code, 200)
        self.assertIn("No programmes or groups match", html)


class DayTimeGridTests(TestCase):
    """The default on-screen grid: DAYS as rows (full weekday labels), TIME as
    columns, each day drawn as vertical lanes with merged session cells."""

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(code="EE", name="BSc. in Electrical Engineering")
        self.g = StudentGroup.objects.create(programme=self.prog, code="C1")
        StudentGroup.objects.create(programme=self.prog, code="C2")
        self.venue = Venue.objects.create(name="NB102", capacity=80)

    @staticmethod
    def _entry(**overrides):
        base = {
            "key": ("session", 1),
            "day": "MONDAY",
            "hours": {8, 9},
            "label": "MT171",
            "kind": "lecture",
            "course_code": "MT171",
            "name": "MT171",
            "type_label": "Lecture",
            "venue": "NB102",
            "start": "08:00",
            "end": "09:55",
            "groups": "C1",
            "note": "",
        }
        base.update(overrides)
        return base

    def _spanning_cells(self, grid):
        return [
            c
            for r in grid["rows"]
            for lane in r["lanes"]
            for c in lane
            if not c.get("empty")
        ]

    def test_orientation_full_day_labels_lane_span(self):
        g = build_day_time_grid([self._entry()])
        self.assertEqual(g["slots"][0]["start"], "07:00")
        self.assertEqual(g["slots"][0]["label"], "07:00-07:55")
        # All five weekdays are present, even with data on one day only.
        self.assertEqual(
            [r["label"] for r in g["rows"]],
            ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
        )
        self.assertEqual(g["rows"][0]["day"], "MONDAY")
        lanes = g["rows"][0]["lanes"]
        self.assertEqual(len(lanes), 1)
        self.assertEqual(lanes[0][0]["empty"], True)
        self.assertEqual(len(lanes[0]), 2)
        self.assertEqual(lanes[0][-1]["colspan"], 2)
        self.assertEqual(lanes[0][-1]["entries"][0]["course_code"], "MT171")

    def test_column_range_ends_at_latest_hour(self):
        g = build_day_time_grid(
            [
                self._entry(
                    key=("session", 1),
                    day="FRIDAY",
                    hours={13},
                    label="IE101",
                    course_code="IE101",
                    name="IE101",
                    start="13:00",
                    end="14:00",
                    groups="",
                )
            ]
        )
        headers = [s["start"] for s in g["slots"]]
        self.assertEqual(headers, ["07:00", "08:00", "09:00", "10:00", "11:00", "12:00", "13:00"])
        self.assertNotIn("14:00", headers)

    def test_day_time_grid_merges_across_hours(self):
        session = Session.objects.create(
            semester=self.sem,
            course_code="MT171",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="09:55",
            venue=self.venue,
        )
        SessionGroup.objects.create(session=session, group=self.g)
        grid = build_day_time_grid(collect_entries(self.prog, self.sem))
        self.assertEqual([c.get("colspan") for c in self._spanning_cells(grid)], [2])

    def test_sessions_begin_under_their_start_hour_column(self):
        g = build_day_time_grid(
            [
                self._entry(key=("session", 1), name="A", course_code="A",
                            day="MONDAY", hours={8}, start="08:00", end="09:00"),
                self._entry(key=("session", 2), name="B", course_code="B",
                            day="MONDAY", hours={10}, start="10:00", end="11:00"),
                self._entry(key=("session", 3), name="C", course_code="C",
                            day="MONDAY", hours={14}, start="14:00", end="15:00"),
            ]
        )
        lane = g["rows"][0]["lanes"][0]
        # A sits right after the 07:00 empty column, then an empty column
        # (09:00) before B at 10:00, then three empty columns (11:00-13:00)
        # before C at 14:00. The column the block starts on is its start hour.
        empties = [c["colspan"] for c in lane if c.get("empty")]
        blocks = [c["entries"][0]["course_code"] for c in lane if not c.get("empty")]
        self.assertEqual(empties, [1, 1, 3])
        self.assertEqual(blocks, ["A", "B", "C"])

    def test_partial_hour_session_spans_columns(self):
        g = build_day_time_grid(
            [
                self._entry(
                    key=("session", 1),
                    day="MONDAY",
                    hours={8, 9, 10},
                    start="08:30",
                    end="10:15",
                )
            ]
        )
        cells = [c for c in self._spanning_cells(g) if not c.get("empty")]
        self.assertEqual(cells[0]["colspan"], 3)

    def test_spanned_session_is_one_cell_not_per_hour(self):
        g = build_day_time_grid(
            [
                self._entry(
                    key=("session", 1),
                    day="MONDAY",
                    hours={9, 10, 11, 12},
                    start="09:00",
                    end="13:00",
                )
            ]
        )
        cells = [c for c in self._spanning_cells(g) if not c.get("empty")]
        self.assertEqual(len(cells), 1)
        self.assertEqual(cells[0]["colspan"], 4)
        self.assertEqual(len(cells[0]["entries"]), 1)

    def test_overlapping_sessions_use_separate_lanes(self):
        g = build_day_time_grid(
            [
                self._entry(key=("session", 1), day="MONDAY", hours={9, 10, 11, 12}),
                self._entry(key=("session", 2), day="MONDAY", hours={10, 11, 12, 13}),
            ]
        )
        row = g["rows"][0]
        self.assertEqual(len(row["lanes"]), 2)
        for lane in row["lanes"]:
            non_empty = [c for c in lane if not c.get("empty")]
            self.assertEqual(len(non_empty), 1)
            self.assertEqual(non_empty[0]["colspan"], 4)

    def test_adjacent_sessions_share_a_lane(self):
        g = build_day_time_grid(
            [
                self._entry(key=("session", 1), day="MONDAY", hours={8}, start="08:00", end="09:00"),
                self._entry(key=("session", 2), day="MONDAY", hours={9}, start="09:00", end="10:00"),
            ]
        )
        row = g["rows"][0]
        self.assertEqual(len(row["lanes"]), 1)
        cells = [c for c in row["lanes"][0] if not c.get("empty")]
        self.assertEqual(len(cells), 2)

    def test_empty_day_is_still_rendered_with_an_empty_lane(self):
        g = build_day_time_grid(
            [self._entry(key=("session", 1), day="TUESDAY", hours={9})]
        )
        monday = g["rows"][0]
        self.assertEqual(monday["day"], "MONDAY")
        self.assertEqual(len(monday["lanes"]), 1)
        self.assertEqual(len(monday["lanes"][0]), 1)
        self.assertEqual(monday["lanes"][0][0]["empty"], True)
        self.assertEqual(monday["lanes"][0][0]["colspan"], len(g["slots"]))

    def test_activity_card_renders_excel_style_text(self):
        session = Session.objects.create(
            semester=self.sem,
            course_code="MT171",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="09:00",
            venue=self.venue,
        )
        SessionGroup.objects.create(session=session, group=self.g)
        resp = self.client.get(
            "/timetable/",
            {"programme": self.prog.pk, "semester": self.sem.pk},
            HTTP_HOST="localhost",
        )
        html = resp.content.decode()
        # Box reads exactly like the Excel grid cells: header line, then
        # Course:/Venue:/Assigned groups: rows.
        self.assertIn("Lecture, 08:00-09:00, Mon", html)
        self.assertIn("Course: MT171", html)
        self.assertIn("Venue: NB102", html)
        self.assertIn("Assigned groups: C1", html)
        self.assertNotIn("Assigned groups: N/A", html)

    def test_activity_card_groups_line_blank_without_groups(self):
        Session.objects.create(
            semester=self.sem,
            course_code="MT171",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="09:00",
            venue=self.venue,
        )
        # Ungrouped sessions still appear (master timetable is the source of
        # truth) and their "Assigned groups:" line is left blank.
        resp = self.client.get("/timetable/", HTTP_HOST="localhost")
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Course: MT171", html)
        self.assertRegex(html, r"Assigned groups:\s*</div>")

    def test_activity_label_filter_maps_variants(self):
        from django.template import Template, Context
        from django.template import engines

        engine = engines["django"]
        tmpl = engine.from_string(
            "{% load core_tags %}"
            "{{ 'lecture'|activity_label }}|{{ 'Lectures'|activity_label }}"
            "|{{ 'Tutorial'|activity_label }}|{{ 'PRACTICAL'|activity_label }}"
            "|{{ ' Seminar '|activity_label }}|{{ 'Workshops'|activity_label }}"
        )
        out = tmpl.render({"request": None})
        self.assertEqual(out, "Lecture|Lecture|Tutorial|Practical|Seminar|Workshop")


class DeletionImpactTests(TestCase):
    """Delete previews must show the full cascade before confirming, and the
    actual delete must match what was previewed (Django cascades + detaches)."""

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(code="CE", name="Civil Engineering")
        self.g1 = StudentGroup.objects.create(programme=self.prog, code="A1")
        self.g2 = StudentGroup.objects.create(programme=self.prog, code="A2")
        self.course = ProgrammeCourse.objects.create(
            programme=self.prog, course_code="MT161", course_name="Maths 1", semester=1
        )
        self.venue = Venue.objects.create(name="NB102", capacity=100)
        self.session = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=self.venue,
        )
        self.link1 = SessionGroup.objects.create(session=self.session, group=self.g1)
        self.link2 = SessionGroup.objects.create(session=self.session, group=self.g2)
        self.workshop = WorkshopAllocation.objects.create(
            semester=self.sem, course_code="TG201", group_code="A1", day="TUESDAY"
        )
        self.td = TechnicalDrawingAllocation.objects.create(
            semester=self.sem,
            course_code="TG201",
            group_code="A1",
            day="WEDNESDAY",
            start_time="09:00",
            end_time="12:00",
            venue="TW101",
        )

    def _delete_page(self, url):
        return self.client.get(url)

    def test_programme_preview_lists_cascaded_records(self):
        resp = self._delete_page("/programmes/%d/delete/" % self.prog.pk)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Civil Engineering", html)
        self.assertIn("Student groups", html)
        self.assertIn("Programme courses", html)
        self.assertIn("Session-group links", html)
        # Badges: 2 student groups, 1 programme course, 2 session-group links.
        self.assertIn('bg-red-100 text-red-700 text-xs font-semibold">2</span>', html)
        self.assertIn('bg-red-100 text-red-700 text-xs font-semibold">1</span>', html)
        # Example rows are rendered for at least the groups and course.
        self.assertIn("CE A1", html)
        self.assertIn("Maths 1", html)
        # Sessions are NOT affected, so the Sessions group is never shown.
        self.assertNotIn(">Sessions<", html)

    def test_programme_delete_cascades_but_keeps_sessions(self):
        self.client.get("/")
        self.prog.delete()
        self.assertFalse(Programme.objects.filter(pk=self.prog.pk).exists())
        self.assertFalse(StudentGroup.objects.filter(pk=self.g1.pk).exists())
        self.assertFalse(ProgrammeCourse.objects.filter(pk=self.course.pk).exists())
        self.assertFalse(SessionGroup.objects.filter(session=self.session).exists())
        self.assertTrue(Session.objects.filter(pk=self.session.pk).exists())

    def test_session_preview_lists_its_group_links(self):
        resp = self._delete_page("/sessions/%d/delete/" % self.session.pk)
        html = resp.content.decode()
        self.assertIn("Session-group links", html)
        self.assertIn("Will be permanently deleted", html)

    def test_session_delete_removes_links_not_groups(self):
        self.session.delete()
        self.assertFalse(
            SessionGroup.objects.filter(session_id=self.session.pk).exists()
        )
        self.assertTrue(StudentGroup.objects.filter(pk=self.g1.pk).exists())
        self.assertTrue(Programme.objects.filter(pk=self.prog.pk).exists())

    def test_venue_preview_distinguishes_detached_not_deleted(self):
        resp = self._delete_page("/venues/%d/delete/" % self.venue.pk)
        html = resp.content.decode()
        self.assertIn("Not deleted, but", html)
        self.assertIn("Sessions", html)
        self.assertIn("venue", html)
        # Venue has no cascade children, so nothing is listed as deleted.
        self.assertNotIn("Will be permanently deleted", html)

    def test_venue_delete_clears_fk_keeps_session(self):
        self.session.refresh_from_db()
        self.assertEqual(self.session.venue, self.venue)
        self.venue.delete()
        self.assertFalse(Venue.objects.filter(pk=self.venue.pk).exists())
        self.assertTrue(Session.objects.filter(pk=self.session.pk).exists())
        self.session.refresh_from_db()
        self.assertIsNone(self.session.venue)

    def test_semester_preview_lists_all_kinds(self):
        resp = self._delete_page("/semesters/%d/delete/" % self.sem.pk)
        html = resp.content.decode()
        self.assertIn("Sessions", html)
        self.assertIn("Workshop allocations", html)
        self.assertIn("Technical drawing allocations", html)
        self.assertIn("Session-group links", html)

    def test_semester_delete_cascades_everything(self):
        self.sem.delete()
        self.assertFalse(Session.objects.filter(pk=self.session.pk).exists())
        self.assertFalse(
            WorkshopAllocation.objects.filter(pk=self.workshop.pk).exists()
        )
        self.assertFalse(TechnicalDrawingAllocation.objects.filter(pk=self.td.pk).exists())
        self.assertFalse(
            SessionGroup.objects.filter(session_id=self.session.pk).exists()
        )
        # The programme/group/course are untouched by a semester delete.
        self.assertTrue(Programme.objects.filter(pk=self.prog.pk).exists())
        self.assertTrue(ProgrammeCourse.objects.filter(pk=self.course.pk).exists())

    def test_no_related_records_says_so(self):
        # A programme course has no cascade children.
        resp = self._delete_page("/courses/%d/delete/" % self.course.pk)
        html = resp.content.decode()
        self.assertIn("No related records will be deleted.", html)
        self.assertNotIn("Will be permanently deleted", html)

    def test_group_delete_preview_shows_its_links(self):
        resp = self._delete_page("/groups/%d/delete/" % self.g1.pk)
        html = resp.content.decode()
        self.assertIn("Session-group links", html)
        self.assertIn("Will be permanently deleted", html)

    def test_group_delete_removes_only_its_own_links(self):
        self.g1.delete()
        self.assertFalse(SessionGroup.objects.filter(pk=self.link1.pk).exists())
        self.assertTrue(SessionGroup.objects.filter(pk=self.link2.pk).exists())
        self.assertTrue(Session.objects.filter(pk=self.session.pk).exists())

    def test_delete_preview_get_renders_form_and_csrf(self):
        resp = self._delete_page("/programmes/%d/delete/" % self.prog.pk)
        html = resp.content.decode()
        self.assertIn("csrfmiddlewaretoken", html)
        self.assertIn("Yes, Delete", html)
        self.assertIn("Cancel", html)

    def test_recycle_remove_shows_venue_impact(self):
        v2 = Venue.objects.create(name="nb102", capacity=90)
        self.session.venue = v2
        self.session.save()
        resp = self.client.get("/venues/recycle/")
        html = resp.content.decode()
        self.assertIn("re-pointed", html)


class ClearAllTests(TestCase):
    """Page-level 'Clear All' on Sessions / Workshop Allocations / TD Allocations:
    GET preview counts, POST-only with a typed phrase and CSRF, exact deletion
    scope with SessionGroup cascade, preserved reference data, htmx + plain POST
    responses, one ActivityLog entry, and a graceful empty dataset."""

    def setUp(self):
        self.client = Client(enforce_csrf_checks=True)
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(code="CE", name="Civil Engineering")
        self.g1 = StudentGroup.objects.create(programme=self.prog, code="A1")
        self.g2 = StudentGroup.objects.create(programme=self.prog, code="A2")
        self.course = ProgrammeCourse.objects.create(
            programme=self.prog, course_code="MT161", course_name="Maths 1", semester=1
        )
        self.venue = Venue.objects.create(name="NB102", capacity=100)
        self.session1 = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=self.venue,
        )
        self.session2 = Session.objects.create(
            semester=self.sem,
            course_code="PH130",
            activity_type="TUTORIAL",
            day="TUESDAY",
            start_time="08:00",
            end_time="09:00",
            venue=self.venue,
        )
        self.link1 = SessionGroup.objects.create(session=self.session1, group=self.g1)
        self.link2 = SessionGroup.objects.create(session=self.session2, group=self.g2)
        self.workshop = WorkshopAllocation.objects.create(
            semester=self.sem, course_code="TG201", group_code="A1", day="WEDNESDAY"
        )
        self.td = TechnicalDrawingAllocation.objects.create(
            semester=self.sem,
            course_code="TG201",
            group_code="A1",
            day="THURSDAY",
            start_time="09:00",
            end_time="12:00",
            venue="TW101",
        )

    def _post(self, url, data, hx=False):
        self.client.get("/")
        token = self.client.cookies.get("csrftoken").value
        headers = {}
        if hx:
            headers["HTTP_HX_REQUEST"] = "true"
        return self.client.post(url, {**data, "csrfmiddlewaretoken": token}, **headers)

    def _confirm(self, url, query=""):
        resp = self.client.get(url + query, HTTP_HX_REQUEST="true")
        self.assertEqual(resp.status_code, 200)
        return resp.content.decode()

    def test_confirm_get_shows_counts_phrase_and_csrf(self):
        for url, label in [
            ("/sessions/clear-all/", "Sessions"),
            ("/workshops/clear-all/", "Workshop Allocations"),
            ("/td/clear-all/", "Technical Drawing Allocations"),
        ]:
            with self.subTest(url=url):
                html = self._confirm(url)
                self.assertIn("csrfmiddlewaretoken", html)
                self.assertIn("DELETE ALL", html)
                self.assertIn(label, html)
                self.assertIn("permanently removes", html)
        # Session preview counts the primary rows and its SessionGroup links.
        html = self._confirm("/sessions/clear-all/")
        self.assertIn("Session-group links", html)
        self.assertIn(">2</span>", html)

    def test_confirm_get_does_not_delete(self):
        self._confirm("/sessions/clear-all/")
        self.assertEqual(Session.objects.count(), 2)
        self.assertEqual(SessionGroup.objects.count(), 2)

    def test_post_requires_phrase(self):
        resp = self._post("/sessions/clear-all/", {"phrase": "WRONG"}, hx=True)
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("HX-Trigger", resp.headers)
        self.assertIn("DELETE ALL", resp.content.decode())
        self.assertEqual(Session.objects.count(), 2)
        self.assertFalse(ActivityLog.objects.filter(action=LogAction.CLEAR).exists())

    def test_post_requires_csrf(self):
        self.client.get("/")
        resp = self.client.post(
            "/sessions/clear-all/", {"phrase": "DELETE ALL"}
        )
        self.assertEqual(resp.status_code, 403)
        self.assertEqual(Session.objects.count(), 2)

    def test_htmx_post_clears_sessions_and_cascades_links(self):
        resp = self._post("/sessions/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(Session.objects.count(), 0)
        self.assertEqual(SessionGroup.objects.count(), 0)
        # Reference data is preserved.
        self.assertTrue(Semester.objects.filter(pk=self.sem.pk).exists())
        self.assertTrue(Programme.objects.filter(pk=self.prog.pk).exists())
        self.assertTrue(StudentGroup.objects.filter(pk=self.g1.pk).exists())
        self.assertTrue(StudentGroup.objects.filter(pk=self.g2.pk).exists())
        self.assertTrue(ProgrammeCourse.objects.filter(pk=self.course.pk).exists())
        self.assertTrue(Venue.objects.filter(pk=self.venue.pk).exists())
        # Other allocation types are untouched.
        self.assertTrue(WorkshopAllocation.objects.filter(pk=self.workshop.pk).exists())
        self.assertTrue(TechnicalDrawingAllocation.objects.filter(pk=self.td.pk).exists())

    def test_plain_post_redirects(self):
        resp = self._post("/sessions/clear-all/", {"phrase": "DELETE ALL"})
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, "/sessions/")
        self.assertEqual(Session.objects.count(), 0)
        self.assertEqual(SessionGroup.objects.count(), 0)

    def test_workshop_clear_all_only_removes_workshops(self):
        resp = self._post(
            "/workshops/clear-all/", {"phrase": "DELETE ALL"}, hx=True
        )
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(WorkshopAllocation.objects.count(), 0)
        self.assertEqual(Session.objects.count(), 2)
        self.assertEqual(TechnicalDrawingAllocation.objects.count(), 1)

    def test_td_clear_all_only_removes_td(self):
        resp = self._post("/td/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(TechnicalDrawingAllocation.objects.count(), 0)
        self.assertEqual(Session.objects.count(), 2)
        self.assertEqual(WorkshopAllocation.objects.count(), 1)

    def test_clear_all_logs_one_entry_with_counts(self):
        self._post("/sessions/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        logs = ActivityLog.objects.filter(action=LogAction.CLEAR)
        self.assertEqual(logs.count(), 1)
        log = logs.get()
        self.assertEqual(log.resource, "Session")
        self.assertIn("Cleared all Sessions", log.message)
        self.assertIn("2 Session record(s)", log.message)
        self.assertIn("2 Session-group links", log.message)
        # The per-row delete-style entries are not affected.
        self._post("/workshops/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(ActivityLog.objects.filter(action=LogAction.CLEAR).count(), 2)

    def test_filters_note_only_when_query_active(self):
        html = self._confirm("/sessions/clear-all/", "?q=MT161")
        self.assertIn("Filtered view", html)
        self.assertIn("not limited", html)
        plain = self._confirm("/sessions/clear-all/")
        self.assertNotIn("Filtered view", plain)

    def test_empty_dataset_is_graceful(self):
        Session.objects.all().delete()
        WorkshopAllocation.objects.all().delete()
        TechnicalDrawingAllocation.objects.all().delete()
        html = self._confirm("/sessions/clear-all/")
        self.assertIn(">0</span>", html)
        self.assertIn("Nothing will be deleted", html)
        self.assertNotIn('name="phrase"', html)
        before = ActivityLog.objects.count()
        resp = self._post("/sessions/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(ActivityLog.objects.count(), before)
        self.assertEqual(Session.objects.count(), 0)


class ClearAllReferenceDataTests(TestCase):
    """Universal 'Clear All' on Venues / Programmes / Programme Courses /
    Student Groups: same confirmation, POST-only with phrase + CSRF, exact
    deletion scope per section, preserved unrelated data and references, htmx
    + plain POST responses, one ActivityLog entry, and graceful empty data."""

    def setUp(self):
        self.client = Client(enforce_csrf_checks=True)
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(code="CE", name="Civil Engineering")
        self.g1 = StudentGroup.objects.create(programme=self.prog, code="A1")
        self.g2 = StudentGroup.objects.create(programme=self.prog, code="A2")
        self.course = ProgrammeCourse.objects.create(
            programme=self.prog, course_code="MT161", course_name="Maths 1", semester=1
        )
        self.venue = Venue.objects.create(name="NB102", capacity=100)
        self.session1 = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=self.venue,
        )
        self.session2 = Session.objects.create(
            semester=self.sem,
            course_code="PH130",
            activity_type="TUTORIAL",
            day="TUESDAY",
            start_time="08:00",
            end_time="09:00",
            venue=self.venue,
        )
        self.link1 = SessionGroup.objects.create(session=self.session1, group=self.g1)
        self.link2 = SessionGroup.objects.create(session=self.session2, group=self.g2)
        self.workshop = WorkshopAllocation.objects.create(
            semester=self.sem, course_code="TG201", group_code="A1", day="WEDNESDAY"
        )
        self.td = TechnicalDrawingAllocation.objects.create(
            semester=self.sem,
            course_code="TG201",
            group_code="A1",
            day="THURSDAY",
            start_time="09:00",
            end_time="12:00",
            venue="TW101",
        )

    def _post(self, url, data, hx=False):
        self.client.get("/")
        token = self.client.cookies.get("csrftoken").value
        headers = {}
        if hx:
            headers["HTTP_HX_REQUEST"] = "true"
        return self.client.post(url, {**data, "csrfmiddlewaretoken": token}, **headers)

    def _confirm(self, url, query=""):
        resp = self.client.get(url + query, HTTP_HX_REQUEST="true")
        self.assertEqual(resp.status_code, 200)
        return resp.content.decode()

    def test_list_pages_show_clear_all_button(self):
        for url, label in [
            ("/venues/", "Venues"),
            ("/programmes/", "Programmes"),
            ("/courses/", "Programme Courses"),
            ("/groups/", "Student Groups"),
        ]:
            resp = self.client.get(url)
            html = resp.content.decode()
            self.assertEqual(resp.status_code, 200)
            self.assertIn("Clear All", html)
            self.assertIn('id="record-total"', html)
            self.assertIn(label, html)

    def test_confirm_get_heading_phrase_and_csrf(self):
        for url, label in [
            ("/venues/clear-all/", "venues"),
            ("/programmes/clear-all/", "programmes"),
            ("/courses/clear-all/", "programme courses"),
            ("/groups/clear-all/", "student groups"),
        ]:
            with self.subTest(url=url):
                html = self._confirm(url)
                self.assertIn("csrfmiddlewaretoken", html)
                self.assertIn("DELETE ALL", html)
                self.assertIn(f"clear all {label} data?", html)
                self.assertIn("This action cannot be undone.", html)
                self.assertIn("permanently removes", html)

    def test_programme_confirm_lists_cascaded_related(self):
        html = self._confirm("/programmes/clear-all/")
        for group in ("Student groups", "Programme courses", "Session-group links"):
            self.assertIn(group, html)
        # Groups (2), courses (1), session-group links (2) all rendered with badges.
        self.assertIn(">2</span>", html)
        self.assertIn(">1</span>", html)

    def test_venue_confirm_lists_detached_sessions(self):
        html = self._confirm("/venues/clear-all/")
        self.assertIn("Not deleted", html)
        self.assertIn("have their reference to these records cleared", html)
        self.assertIn("sessions", html)
        self.assertIn(">2</span>", html)  # the two sessions that reference a venue

    def test_get_does_not_delete(self):
        for url in (
            "/venues/clear-all/",
            "/programmes/clear-all/",
            "/courses/clear-all/",
            "/groups/clear-all/",
        ):
            self._confirm(url)
        self.assertGreater(Programme.objects.count(), 0)
        self.assertGreater(Venue.objects.count(), 0)
        self.assertGreater(StudentGroup.objects.count(), 0)
        self.assertGreater(ProgrammeCourse.objects.count(), 0)
        self.assertGreater(SessionGroup.objects.count(), 0)

    def test_post_requires_phrase(self):
        resp = self._post("/venues/clear-all/", {"phrase": "WRONG"}, hx=True)
        self.assertEqual(resp.status_code, 200)
        self.assertNotIn("HX-Trigger", resp.headers)
        self.assertIn("DELETE ALL", resp.content.decode())
        self.assertGreater(Venue.objects.count(), 0)
        self.assertFalse(ActivityLog.objects.filter(action=LogAction.CLEAR).exists())

    def test_post_requires_csrf(self):
        self.client.get("/")
        resp = self.client.post("/programmes/clear-all/", {"phrase": "DELETE ALL"})
        self.assertEqual(resp.status_code, 403)
        self.assertGreater(Programme.objects.count(), 0)

    def test_venue_clear_all_clears_reference_not_sessions(self):
        resp = self._post("/venues/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(Venue.objects.count(), 0)
        self.assertEqual(Session.objects.count(), 2)
        self.session1.refresh_from_db()
        self.session2.refresh_from_db()
        self.assertIsNone(self.session1.venue)
        self.assertIsNone(self.session2.venue)
        for model in (Semester, Programme, StudentGroup, ProgrammeCourse):
            self.assertGreater(model.objects.count(), 0)
        self.assertEqual(WorkshopAllocation.objects.count(), 1)
        self.assertEqual(TechnicalDrawingAllocation.objects.count(), 1)
        # Session-group links survive a venue clear.
        self.assertEqual(SessionGroup.objects.count(), 2)

    def test_programme_clear_all_only_cascades_linked(self):
        resp = self._post("/programmes/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(Programme.objects.count(), 0)
        self.assertEqual(StudentGroup.objects.count(), 0)
        self.assertEqual(ProgrammeCourse.objects.count(), 0)
        self.assertEqual(SessionGroup.objects.count(), 0)
        self.assertEqual(Session.objects.count(), 2)
        self.assertEqual(Venue.objects.count(), 1)
        self.assertEqual(Semester.objects.count(), 1)
        self.assertEqual(WorkshopAllocation.objects.count(), 1)
        self.assertEqual(TechnicalDrawingAllocation.objects.count(), 1)

    def test_course_clear_all_only_removes_mappings(self):
        resp = self._post("/courses/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(ProgrammeCourse.objects.count(), 0)
        self.assertEqual(Programme.objects.count(), 1)
        self.assertEqual(StudentGroup.objects.count(), 2)
        self.assertEqual(Session.objects.count(), 2)
        self.assertEqual(SessionGroup.objects.count(), 2)
        self.assertEqual(Venue.objects.count(), 1)
        self.assertEqual(Semester.objects.count(), 1)

    def test_group_clear_all_removes_links_keeps_rest(self):
        resp = self._post("/groups/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(StudentGroup.objects.count(), 0)
        self.assertEqual(SessionGroup.objects.count(), 0)
        self.assertEqual(Session.objects.count(), 2)
        self.assertEqual(Programme.objects.count(), 1)
        self.assertEqual(ProgrammeCourse.objects.count(), 1)
        self.assertEqual(Semester.objects.count(), 1)

    def test_plain_post_redirects(self):
        resp = self._post("/groups/clear-all/", {"phrase": "DELETE ALL"})
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, "/groups/")
        self.assertEqual(StudentGroup.objects.count(), 0)
        resp = self._post("/venues/clear-all/", {"phrase": "DELETE ALL"})
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, "/venues/")

    def test_activity_logged_with_counts(self):
        self._post(
            "/programmes/clear-all/", {"phrase": "DELETE ALL"}, hx=True
        )
        logs = ActivityLog.objects.filter(action=LogAction.CLEAR)
        self.assertEqual(logs.count(), 1)
        log = logs.get()
        self.assertEqual(log.resource, "Programme")
        self.assertIn("Cleared all Programmes", log.message)
        self.assertIn("1 Programme record(s)", log.message)
        self.assertIn("2 Student groups", log.message)
        self.assertIn("1 Programme courses", log.message)
        self.assertIn("2 Session-group links", log.message)
        # One entry per action: a venue clear logs a detached-reference note.
        self._post("/venues/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(ActivityLog.objects.filter(action=LogAction.CLEAR).count(), 2)
        venue_log = ActivityLog.objects.filter(
            action=LogAction.CLEAR, resource="Venue"
        ).get()
        self.assertIn("1 Venue record(s) removed", venue_log.message)
        self.assertIn("2 sessions kept with reference cleared", venue_log.message)

    def test_empty_dataset_is_graceful(self):
        Venue.objects.all().delete()
        html = self._confirm("/venues/clear-all/")
        self.assertIn(">0</span>", html)
        self.assertIn("Nothing will be deleted", html)
        self.assertNotIn('name="phrase"', html)
        before = ActivityLog.objects.count()
        resp = self._post("/venues/clear-all/", {"phrase": "DELETE ALL"}, hx=True)
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(ActivityLog.objects.count(), before)
        self.assertEqual(Venue.objects.count(), 0)

def test_filters_note_when_query_active(self):
        html = self._confirm("/venues/clear-all/", "?q=NB")
        self.assertIn("Filtered view", html)
        self.assertNotIn("not limited", html)
        plain = self._confirm("/venues/clear-all/")
        self.assertNotIn("Filtered view", plain)


class ImportHistoryTests(TestCase):
    """Progress states, notification priority, persistence and retrieval of
    the per-import history summaries recorded after every import."""

    def _upload(self, import_type, filename, rows, columns, extra=None):
        self.client.get(f"/import/{import_type}/")
        path = make_xlsx(rows, columns)
        data = {
            "file": SimpleUploadedFile(
                filename, Path(path).read_bytes(), content_type=XLSX_CONTENT_TYPE
            )
        }
        if extra:
            data.update(extra)
        return self.client.post(f"/import/{import_type}/", data, HTTP_HX_REQUEST="true")

    def test_success_import_records_complete_summary(self):
        resp = self._upload(
            "venues", "venues_ok.xlsx", [["LH1", "80"], ["LB2", "40"]], ["name", "capacity"]
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("Imported", html)
        self.assertIn("Created", html)
        rec = ImportHistory.objects.get(import_type="venues")
        self.assertEqual(rec.status, "SUCCESS")
        self.assertEqual(rec.import_title, "Venues")
        self.assertEqual(rec.filename, "venues_ok.xlsx")
        self.assertIsNone(rec.user)
        self.assertEqual(rec.rows_processed, 2)
        self.assertEqual(rec.created, 2)
        self.assertEqual(rec.updated, 0)
        self.assertEqual(rec.skipped, 0)
        self.assertEqual(rec.error_count, 0)
        details = json.loads(rec.details)
        self.assertEqual(details["errors"], [])
        self.assertEqual(details["created"], 2)
        self.assertTrue(Venue.objects.filter(name="LH1").exists())

    def test_partial_import_counts_status_and_errors(self):
        resp = self._upload(
            "venues", "venues_partial.xlsx", [["LH1", "80"], ["BAD", "abc"]], ["name", "capacity"]
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Imported with issues", resp.content.decode())
        rec = ImportHistory.objects.get(import_type="venues")
        self.assertEqual(rec.status, "PARTIAL")
        self.assertEqual(rec.created, 1)
        self.assertEqual(rec.skipped, 1)
        self.assertEqual(rec.error_count, 1)
        self.assertEqual(rec.rows_processed, 2)
        details = json.loads(rec.details)
        self.assertIn("Invalid capacity for venue 'BAD'", details["errors"])
        self.assertEqual(Venue.objects.count(), 1)
        self.assertTrue(Venue.objects.filter(name="LH1").exists())

    def test_failed_import_writes_nothing_and_is_failed(self):
        resp = self._upload(
            "venues", "venues_bad.xlsx", [["X1", "abc"], ["Y1", "xyz"]], ["name", "capacity"]
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Import Failed", resp.content.decode())
        rec = ImportHistory.objects.get(import_type="venues")
        self.assertEqual(rec.status, "FAILED")
        self.assertEqual(rec.created, 0)
        self.assertEqual(rec.skipped, 2)
        self.assertEqual(rec.rows_processed, 2)
        self.assertEqual(rec.error_count, 2)
        self.assertEqual(Venue.objects.count(), 0)

    def test_htmx_success_progress_state_and_toast(self):
        resp = self._upload(
            "programmes", "progs_ok.xlsx", [["CE", "Civil Engineering"]], ["code", "name"]
        )
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        # "Importing..." progress state is present on the hosting page and
        # "Imported" is the terminal success state after completion.
        self.assertIn("Imported", html)
        trigger = resp.headers.get("HX-Trigger", "")
        self.assertIn("import-toast", trigger)
        self.assertIn("Import complete", trigger)
        self.assertNotIn("progs_ok.xlsx", trigger)

    def test_routine_validation_errors_do_not_toast_or_persist(self):
        # No file attached: a routine form-level validation message. It is
        # visible in the results panel below the form but must never pop up
        # as a toast, and nothing is persisted because the import never ran.
        self.client.get("/import/venues/")
        resp = self.client.post("/import/venues/", {}, HTTP_HX_REQUEST="true")
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("No file attached", html)
        self.assertIn("Import Failed", html)
        self.assertNotIn("HX-Trigger", resp.headers)
        self.assertEqual(ImportHistory.objects.count(), 0)

    def test_missing_semester_blocker_does_not_toast_or_persist(self):
        Semester.objects.create(academic_year="2026/2027", semester=1)
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
        self.assertIn("academic semester", resp.content.decode().lower())
        self.assertNotIn("HX-Trigger", resp.headers)
        self.assertEqual(ImportHistory.objects.count(), 0)
        self.assertEqual(Session.objects.count(), 0)

    def test_failed_real_import_toasts_failure(self):
        resp = self._upload(
            "venues", "venues_fail.xlsx", [["X1", "abc"]], ["name", "capacity"]
        )
        trigger = resp.headers.get("HX-Trigger", "")
        self.assertIn("import-toast", trigger)
        self.assertIn("Import failed", trigger)
        self.assertIn("blocked", trigger)

    def test_history_listing_and_type_filtering(self):
        self._upload("venues", "venues_a.xlsx", [["LH1", "80"]], ["name", "capacity"])
        self._upload("programmes", "progs_a.xlsx", [["CE", "Civil Engineering"]], ["code", "name"])
        self.assertEqual(ImportHistory.objects.count(), 2)

        all_page = self.client.get("/import/history/").content.decode()
        self.assertIn("venues_a.xlsx", all_page)
        self.assertIn("progs_a.xlsx", all_page)

        venues_page = self.client.get("/import/history/venues/").content.decode()
        self.assertIn("venues_a.xlsx", venues_page)
        self.assertNotIn("progs_a.xlsx", venues_page)

        programmes_page = self.client.get("/import/history/programmes/").content.decode()
        self.assertIn("progs_a.xlsx", programmes_page)
        self.assertNotIn("venues_a.xlsx", programmes_page)

    def test_history_detail_shows_errors_and_skips_file_contents(self):
        self._upload(
            "venues", "venues_detail.xlsx", [["LH1", "80"], ["BAD", "abc"]], ["name", "capacity"]
        )
        rec = ImportHistory.objects.get(import_type="venues")
        resp = self.client.get(f"/import/history/{rec.pk}/")
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        self.assertIn("This import completed with issues", html)
        self.assertIn("Invalid capacity for venue", html)
        self.assertIn("BAD", html)
        self.assertIn("completed with issues", html)
        self.assertIn("structured import summary", html)
        # The affected-row error states exactly what to correct; the raw file
        # value "abc" is never stored.
        self.assertNotIn("abc", rec.details)
        self.assertIn('"errors"', rec.details)

    def test_import_hub_shows_latest_and_history(self):
        self._upload("programmes", "hub_progs.xlsx", [["CE", "Civil Engineering"]], ["code", "name"])
        html = self.client.get("/import/").content.decode()
        self.assertIn("Previous Import Details", html)
        self.assertIn("hub_progs.xlsx", html)
        self.assertIn("SUCCESS", html)


class WorkshopStandardTimesTests(ImporterTestCase):
    """The standard workshop session times are enforced everywhere.

    Monday/Tuesday/Wednesday/Friday morning 09:00-13:00, afternoon
    15:00-19:00; Thursday morning 10:00-14:00, afternoon 15:00-19:00.
    Workshops on other days or at other times are rejected with a clear
    message; non-workshop sessions are never affected. The same times apply
    to every programme, regardless of the record's student group.
    """

    def _workshop(self, **overrides):
        fields = dict(
            semester=self.sem1,
            course_code="TG201",
            group_code="A1",
            day="MONDAY",
            start_time=time(9, 0),
end_time=time(13, 0),
            venue="TW101",
        )
        fields.update(overrides)
        return WorkshopAllocation.objects.create(**fields)

    # --------------------------------------------------------- rules module

    def test_standard_times_table(self):
        cases = [
            ("MONDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),
            ("MONDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),
            ("TUESDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),
            ("TUESDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),
            ("WEDNESDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),
            ("WEDNESDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),
            ("FRIDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),
            ("FRIDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),
            ("THURSDAY", TimePeriod.MORNING, time(10, 0), time(14, 0)),
            ("THURSDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),
        ]
        for day, period, start, end in cases:
            with self.subTest(day=day, period=period):
                self.assertEqual(workshop_times_for(day, period), (start, end))

    def test_weekend_has_no_workshop_sessions(self):
        self.assertIsNone(workshop_times_for("SATURDAY", TimePeriod.MORNING))
        self.assertIsNone(workshop_times_for("SUNDAY", TimePeriod.AFTERNOON))
        self.assertIn("not scheduled", workshop_time_issue("SATURDAY", time(9, 0), time(13, 0)))

    def test_thursday_morning_is_1000_to_1400_for_all_programmes(self):
        for code in ("CE", "TE", "EE", "ME", "QS", "cpE", "UNKNOWN"):
            with self.subTest(programme=code):
                self.assertEqual(
                    workshop_times_for("THURSDAY", TimePeriod.MORNING, code),
                    (time(10, 0), time(14, 0)),
                )
                self.assertEqual(
                    workshop_hours("THURSDAY", TimePeriod.MORNING, code),
                    {10, 11, 12, 13},
                )
        # A mixed bag of programmes still uses the single Thursday morning
        # session, and Thursday-only: Monday morning stays 09:00-13:00.
        self.assertEqual(
            workshop_times_for("THURSDAY", TimePeriod.MORNING, ("ME", "CE")),
            (time(10, 0), time(14, 0)),
        )
        self.assertEqual(
            workshop_times_for("MONDAY", TimePeriod.MORNING, "CE"),
            (time(9, 0), time(13, 0)),
        )
        self.assertEqual(
            workshop_times_for("MONDAY", TimePeriod.MORNING, "ME"),
            (time(9, 0), time(13, 0)),
        )

    def test_workshop_hours_are_day_aware(self):
        self.assertEqual(
            workshop_hours("MONDAY", TimePeriod.MORNING), {9, 10, 11, 12}
        )
        self.assertEqual(
            workshop_hours("THURSDAY", TimePeriod.MORNING), {10, 11, 12, 13}
        )
        self.assertEqual(
            workshop_hours("THURSDAY", TimePeriod.MORNING, "CE"), {10, 11, 12, 13}
        )
        self.assertEqual(
            workshop_hours("FRIDAY", TimePeriod.AFTERNOON), {15, 16, 17, 18}
        )

    def test_time_issue_accepts_standard_and_rejects_bad(self):
        for day in ("MONDAY", "TUESDAY", "WEDNESDAY", "FRIDAY"):
            self.assertEqual(workshop_time_issue(day, time(9, 0), time(13, 0)), "")
            self.assertEqual(workshop_time_issue(day, time(15, 0), time(19, 0)), "")
        self.assertEqual(workshop_time_issue("THURSDAY", time(10, 0), time(14, 0)), "")
        self.assertEqual(workshop_time_issue("THURSDAY", time(15, 0), time(19, 0)), "")
        self.assertIn("must run", workshop_time_issue("MONDAY", time(8, 0), time(10, 0)))
        # Thursday morning is 10:00-14:00 for every programme.
        self.assertIn(
            "must run",
            workshop_time_issue("THURSDAY", time(9, 0), time(13, 0), programme_code="CE"),
        )
        self.assertIn(
            "must run",
            workshop_time_issue("THURSDAY", time(10, 0), time(13, 0), programme_code="ME"),
        )
        self.assertIn(
            "must run",
            workshop_time_issue("THURSDAY", time(10, 0), time(13, 0), programme_code="CE"),
        )
        self.assertIn("must run", workshop_time_issue("THURSDAY", time(9, 0), time(13, 0)))
        self.assertIn("required", workshop_time_issue("MONDAY", time(9, 0), None))

    # ----------------------------------------------------------- model clean

    def test_model_clean_rejects_non_standard_workshop(self):
        self._seed()
        rec = WorkshopAllocation(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="MONDAY", start_time=time(8, 0), end_time=time(10, 0), venue="W",
        )
        with self.assertRaises(ValidationError):
            rec.full_clean()

    def test_model_clean_rejects_weekend_workshop(self):
        self._seed()
        rec = WorkshopAllocation(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="SATURDAY", start_time=time(9, 0), end_time=time(13, 0), venue="W",
        )
        with self.assertRaises(ValidationError) as ctx:
            rec.full_clean()
        self.assertIn("not scheduled", str(ctx.exception))

    def test_model_clean_accepts_matrix_period_only_record(self):
        self._seed()
        rec = WorkshopAllocation(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",
        )
        rec.full_clean()  # must not raise

    def test_model_clean_accepts_thursday_morning(self):
        self._seed()
        rec = WorkshopAllocation(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="THURSDAY", start_time=time(10, 0), end_time=time(14, 0), venue="W",
        )
        rec.full_clean()  # must not raise — Thursday morning is 10:00-14:00

    def test_model_clean_rejects_thursday_0900(self):
        self._seed()
        rec = WorkshopAllocation(
            semester=self.sem1, course_code="TG201", group_code="B1",
            day="THURSDAY", start_time=time(9, 0), end_time=time(13, 0), venue="W",
        )
        with self.assertRaises(ValidationError) as ctx:
            rec.full_clean()
        self.assertIn("10:00-14:00", str(ctx.exception))

    def test_model_clean_rejects_thursday_1000_1300(self):
        self._seed()
        rec = WorkshopAllocation(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="THURSDAY", start_time=time(10, 0), end_time=time(13, 0), venue="W",
        )
        with self.assertRaises(ValidationError) as ctx:
            rec.full_clean()
        self.assertIn("10:00-14:00", str(ctx.exception))

    def test_session_clean_rejects_non_standard_workshop(self):
        self._seed()
        ses = Session(
            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,
            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,
        )
        with self.assertRaises(ValidationError):
            ses.full_clean()

    def test_session_clean_ignores_non_workshop(self):
        self._seed()
        ses = Session(
            semester=self.sem1, course_code="MT161", activity_type=ActivityType.LECTURE,
            day="MONDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,
        )
        ses.full_clean()  # must not raise

    def test_orm_create_bypasses_validation(self):
        self._seed()
        self._workshop(day="MONDAY", start_time=time(8, 0), end_time=time(10, 0))
        self.assertEqual(WorkshopAllocation.objects.count(), 1)

    # ---------------------------------------------------- form inline errors

    def test_workshop_create_shows_inline_time_error(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        resp = self.client.post("/workshops/create/", {
            "semester": sem.pk,
            "course_code": "TG201",
            "group_code": "A1",
            "day": "MONDAY",
            "start_time": "08:00",
            "end_time": "10:00",
            "venue": "TW101",
        })
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "must run")
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_session_create_shows_inline_workshop_time_error(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        resp = self.client.post("/sessions/create/", {
            "semester": sem.pk,
            "course_code": "TG201",
            "activity_type": "WORKSHOP",
            "day": "THURSDAY",
            "start_time": "08:00",
            "end_time": "10:00",
            "venue": "",
            "session_groups-TOTAL_FORMS": "0",
            "session_groups-INITIAL_FORMS": "0",
            "session_groups-MIN_NUM_FORMS": "0",
            "session_groups-MAX_NUM_FORMS": "1000",
        })
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "must run")
        self.assertEqual(Session.objects.count(), 0)

    def test_lecture_session_create_unaffected(self):
        sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        resp = self.client.post("/sessions/create/", {
            "semester": sem.pk,
            "course_code": "MT161",
            "activity_type": "LECTURE",
            "day": "MONDAY",
            "start_time": "08:00",
            "end_time": "10:00",
            "venue": "",
            "session_groups-TOTAL_FORMS": "0",
            "session_groups-INITIAL_FORMS": "0",
            "session_groups-MIN_NUM_FORMS": "0",
            "session_groups-MAX_NUM_FORMS": "1000",
        })
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(Session.objects.get(activity_type="LECTURE").start_time, time(8, 0))
        self.assertEqual(Session.objects.get(activity_type="LECTURE").end_time, time(10, 0))

    # ----------------------------------------------------------- imports

    def test_flat_import_rejects_non_standard_times(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "C1", "MONDAY", "08:00", "10:00", "TW101"]],
                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertTrue(result.errors)
        self.assertIn("must run", result.errors[0])
        self.assertEqual(result.skipped, 1)
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_flat_import_rejects_saturday(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "C1", "SATURDAY", "09:00", "13:00", "TW101"]],
                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertIn("not scheduled", result.errors[0])
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_flat_import_accepts_thursday_morning(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "C1", "THURSDAY", "10:00", "14:00", "TW101"]],
                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertEqual(result.errors, [])
        self.assertEqual(result.created, 1)
        rec = WorkshopAllocation.objects.get()
        self.assertEqual(rec.start_time, time(10, 0))
        self.assertEqual(rec.end_time, time(14, 0))

    def test_flat_import_accepts_thursday_morning_for_any_group(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "A1", "THURSDAY", "10:00", "14:00", "TW101"]],
                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertEqual(result.errors, [])
        self.assertEqual(result.created, 1)
        rec = WorkshopAllocation.objects.get()
        self.assertEqual(rec.start_time, time(10, 0))
        self.assertEqual(rec.end_time, time(14, 0))

    def test_flat_import_rejects_thursday_0900(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "B1", "THURSDAY", "09:00", "13:00", "TW101"]],
                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertTrue(result.errors)
        self.assertIn("10:00-14:00", result.errors[0])
        self.assertEqual(result.skipped, 1)
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_flat_import_rejects_thursday_1000_1300(self):
        self._seed()
        result = import_workshop_allocation_from_excel(
            make_xlsx(
                [["TG201", "A1", "THURSDAY", "10:00", "13:00", "TW101"]],
                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],
            ),
            semester_id=self.sem1.pk,
        )
        self.assertTrue(result.errors)
        self.assertIn("10:00-14:00", result.errors[0])
        self.assertEqual(WorkshopAllocation.objects.count(), 0)

    def test_master_import_rejects_non_standard_workshop(self):
        self._seed()
        path = make_xlsx(
            [["TG201", "WORKSHOP", "MONDAY", "08:00", "10:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertTrue(result.errors)
        self.assertFalse(
            Session.objects.filter(activity_type=ActivityType.WORKSHOP).exists()
        )

    def test_master_import_accepts_default_thursday_morning(self):
        self._seed()
        path = make_xlsx(
            [["TG201", "WORKSHOP", "THURSDAY", "10:00", "14:00", "LH1", "B1"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(result.errors, [])
        self.assertEqual(
            Session.objects.filter(
                activity_type=ActivityType.WORKSHOP,
                start_time=time(10, 0),
                end_time=time(14, 0),
            ).count(),
            1,
        )

    def test_master_import_accepts_thursday_morning_for_any_group(self):
        self._seed()
        path = make_xlsx(
            [["TG201", "WORKSHOP", "THURSDAY", "10:00", "14:00", "LH1", "A1"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertEqual(result.errors, [])
        self.assertEqual(
            Session.objects.filter(
                activity_type=ActivityType.WORKSHOP,
                start_time=time(10, 0),
                end_time=time(14, 0),
            ).count(),
            1,
        )

    def test_master_import_rejects_thursday_0900(self):
        self._seed()
        path = make_xlsx(
            [["TG201", "WORKSHOP", "THURSDAY", "09:00", "13:00", "LH1", "B1"]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)
        self.assertTrue(result.errors)
        self.assertFalse(
            Session.objects.filter(activity_type=ActivityType.WORKSHOP).exists()
        )

    # ---------------------------------------------------------- grid / PDF

    def test_grid_thursday_morning_uses_10_to_14_slot(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.sem1, course_code="TG201", group_code="B1",
            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",
        )
        entries = collect_entries(self.prog_b, self.sem1, group=self.g3)
        self.assertEqual(entries[0]["hours"], {10, 11, 12, 13})
        grid = build_time_day_grid(entries)
        self.assertEqual(grid["slots"][0]["label"], "10:00-11:00")
        self.assertEqual(grid["rows"][0]["cols"][0]["rowspan"], 4)

    def test_grid_thursday_morning_uses_10_to_14_slot_for_any_programme(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",
        )
        entries = collect_entries(self.prog_a, self.sem1, group=self.g1)
        self.assertEqual(entries[0]["hours"], {10, 11, 12, 13})
        grid = build_time_day_grid(entries)
        self.assertEqual(grid["slots"][0]["label"], "10:00-11:00")
        self.assertEqual(grid["rows"][0]["cols"][0]["rowspan"], 4)

    def test_grid_monday_morning_uses_09_to_1255_slot(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="MONDAY", time_period=TimePeriod.MORNING, venue="",
        )
        entries = collect_entries(self.prog_a, self.sem1, group=self.g1)
        self.assertEqual(entries[0]["hours"], {9, 10, 11, 12})
        grid = build_time_day_grid(entries)
        self.assertEqual(grid["slots"][0]["label"], "09:00-10:00")

    def test_pdf_renders_thursday_morning_workshop(self):
        from io import BytesIO

        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",
        )
        out = BytesIO()
        render_programme_timetable(self.prog_a, self.sem1, out=out)
        self.assertGreater(len(out.getvalue()), 1000)

    # -------------------------------------------------------------- legacy

    def test_legacy_workshop_allocations_detected(self):
        self._seed()
        self._workshop()
        self._workshop(day="TUESDAY", start_time=time(8, 0), end_time=time(10, 0))
        found = legacy_workshop_allocations()
        self.assertEqual(len(found), 1)
        rec, issue = found[0]
        self.assertEqual(rec.day, "TUESDAY")
        self.assertIn("must run", issue)

    def test_legacy_workshop_allocations_flag_old_thursday_times(self):
        self._seed()
        # Neither CE (group A1) nor ME (group B1) may start Thursday at 09:00
        # anymore — the morning session is 10:00-14:00 for every programme.
        self._workshop(
            day="THURSDAY", group_code="A1",
            start_time=time(9, 0), end_time=time(13, 0),
        )
        self._workshop(
            day="THURSDAY", group_code="B1",
            start_time=time(9, 0), end_time=time(13, 0),
        )
        found = legacy_workshop_allocations()
        self.assertEqual(len(found), 2)
        self.assertEqual({rec.group_code for rec, _ in found}, {"A1", "B1"})

    def test_programme_resolution_helpers(self):
        self._seed()
        rec = WorkshopAllocation(
            semester=self.sem1, course_code="TG201", group_code="A1",
            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",
        )
        self.assertEqual(allocation_programme_codes(rec), ("CE",))
        self.assertEqual(allocation_programme_codes(
            WorkshopAllocation(
                semester=self.sem1, course_code="TG201", group_code="UNKNOWN",
                day="MONDAY", time_period=TimePeriod.MORNING, venue="",
            )
        ), ())
        self.assertEqual(
            course_programme_codes("TG201"), ("CE",)
        )
        ses = Session(
            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,
            day="THURSDAY", start_time=time(9, 0), end_time=time(13, 0), venue=None,
        )
        ses.save()
        self.assertEqual(session_programme_codes(ses), ())
        SessionGroup.objects.create(session=ses, group=self.g1)
        self.assertEqual(session_programme_codes(ses), ("CE",))
        ses2 = Session(
            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,
            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,
        )
        # An unsaved session's groups have not been assigned yet.
        self.assertEqual(session_programme_codes(ses2), ())

    def test_legacy_workshop_sessions_detected(self):
        self._seed()
        Session.objects.create(
            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,
            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,
        )
        found = legacy_workshop_sessions()
        self.assertEqual(len(found), 1)
        session, issue = found[0]
        self.assertEqual(session.day, "THURSDAY")
        self.assertIn("must run", issue)

    def test_workshop_list_shows_legacy_banner(self):
        self._seed()
        self._workshop(day="TUESDAY", start_time=time(8, 0), end_time=time(10, 0))
        resp = self.client.get("/workshops/")
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "standard workshop session times")

    def test_session_list_shows_legacy_banner(self):
        self._seed()
        Session.objects.create(
            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,
            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,
        )
        resp = self.client.get("/sessions/")
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "standard workshop session times")