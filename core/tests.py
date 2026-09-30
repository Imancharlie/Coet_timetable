import io
import os
import json
import re
import tempfile
from datetime import time
from pathlib import Path
from unittest import mock

import pandas as pd
from django.conf import settings
from django.contrib.auth.models import User
from django.core.exceptions import ValidationError
from django.core.files.uploadedfile import SimpleUploadedFile
from django.db import DatabaseError, IntegrityError, transaction
from django.db.models.deletion import ProtectedError
from django.test import Client, TestCase
from openpyxl import Workbook

from core.forms import ProgrammeCourseForm
from core.group_allocation import (
    apply_run,
    AvailabilityIndex,
    build_requirements,
    capacity_status,
    check_availability,
    derive_requirements_from_sessions,
    find_requirement_count_columns,
    find_requirements_column,
    format_requirements,
    group_statuses,
    manual_assign,
    manual_unassign,
    parse_requirements,
    plan_allocation,
    required_capacity,
    revert_run,
    save_plan,
    validate_assignment,
    validate_manual_assignment,
)
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
    AllocationChange,
    AllocationRun,
    AllocationStatus,
    Course,
    CourseActivityRequirement,
    Day,
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
    normalise_course_code,
)
from core.timetable_grid import (
    FILL_COLORS,
    GROUPS_STYLE,
    build_day_time_grid,
    build_time_day_grid,
    cell_markup,
    cell_text,
    time_day_grid_to_table,
)
from core.timetable_pdf import (
    _build_day_flowables,
    _collect_group_entries_and_rotations,
    _compact_group_codes,
    _entry_lines,
    _entry_parts,
    _master_cell_text,
    _merge_master_entries,
    _wrap_line,
    _DayFlowable,
    build_grid,
    collect_entries,
    collect_group_entries,
    collect_master_entries,
    collect_workshop_rotations,
    fold_for_display,
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
from core.views import _latest_semester_with_data, CLEAR_ALL_TARGETS
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

    def test_the_submit_button_leaves_its_importing_state_when_the_request_ends(self):
        """The busy state is cleared by the request finishing, not by a reload.

        The button is disabled and reads "Importing..." while ``uploading`` is
        true, and nothing else on the page ever clears it — the result panel
        lands in a swap, the URL never changes. Without this handler the
        coordinator is left with a stuck button above a finished result, and
        the only way out is to reload the page.
        """
        html = self.client.get("/import/programmes/").content.decode()
        self.assertIn('@submit="uploading = true"', html)
        self.assertIn(
            '@htmx:after-request.window="if ($event.detail.elt === $el)'
            ' uploading = false"',
            html,
        )

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


class SidebarStructureTests(TestCase):
    """The collapsible sections, the loader and the issue bell.

    The committed student-portal work serves the portal at ``/`` and
    ``StaffAccessMiddleware`` 302s anonymous clients away from staff URLs, so
    these log in as staff and read the ``/staff/`` paths directly. Testing the
    markup any other way would only ever measure the 302.
    """

    def setUp(self):
        self.staff = User.objects.create_user(
            "section-tester", password="pw", is_staff=True
        )
        self.client = Client()
        self.client.force_login(self.staff)
        self.semester = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.programme = Programme.objects.create(code="CE", name="Civil Engineering")

    @staticmethod
    def _sidebar(html):
        """Just the aside, so the mobile bar's duplicate links don't count."""
        start = html.index('id="app-sidebar"')
        end = html.index("</aside>", start)
        return html[start:end]

    @staticmethod
    def _topbar(html):
        """Just the <header> bar, so the sidebar's wordmark doesn't count."""
        start = html.index('<header class="topbar')
        return html[start:html.index("</header>", start)]

    def _html(self, path):
        resp = self.client.get(path)
        self.assertEqual(resp.status_code, 200, path)
        return resp.content.decode()

    def test_the_three_sections_exist_and_nothing_else_is_collapsible(self):
        sidebar = self._sidebar(self._html("/staff/"))
        for key, label in (
            ("reference", "Reference Data"),
            ("timetable", "Timetable"),
            ("allocation", "Allocation"),
        ):
            self.assertIn('aria-controls="nav-section-%s"' % key, sidebar)
            self.assertIn("toggleSection('%s')" % key, sidebar)
            self.assertIn("isSectionOpen('%s')" % key, sidebar)
            self.assertIn(label, sidebar)
        # Import, Export, the Activity Log and the Danger Zone stay flat -- a
        # section header on them would strand their links behind a chevron.
        self.assertEqual(sidebar.count("toggleSection("), 3)
        for flat in ("Import Data", "Export", "Activity Log", "Danger Zone"):
            self.assertIn(flat, sidebar)
            self.assertNotIn('aria-controls="nav-%s"' % flat.lower()[:6], sidebar)

    def test_the_group_allocator_and_progress_share_the_allocation_section(self):
        sidebar = self._sidebar(self._html("/staff/"))
        block = sidebar[sidebar.index('id="nav-section-allocation"'):]
        self.assertLess(block.index('href="/allocation/"'), block.index('href="/allocation/groups/"'))
        # The progress board is "Group Progress" — the full name states what the
        # page tracks, so it cannot be mistaken for the allocator's own progress.
        # It still lives inside the collapsible Allocation section.
        self.assertIn(">Group Progress<", block)
        self.assertIn(">Group Allocation<", block)

    def test_a_section_body_is_never_hid_without_a_persistent_key(self):
        """Every x-show in the sidebar is bound to a key the store persists.

        The store's own key list is what section state is written against; a
        body bound to a key missing from it would open on load and refuse to
        remember its state.
        """
        html = self._html("/staff/")
        sidebar = self._sidebar(html)
        keys = set(re.findall(r"isSectionOpen\('([a-z]+)'\)", sidebar))
        self.assertEqual(keys, {"reference", "timetable", "allocation"})
        store = html[html.index("Alpine.data('sidebar'"):]
        for key in keys:
            self.assertIn("'%s'" % key, store)
        # Persistence is what makes a collapse survive a reload.
        self.assertIn("coet.sidebar.sections", store)
        self.assertIn("localStorage", store)
        # A collapsed desktop rail must not be allowed to hide the icons, or
        # the 16px sidebar would lose every child page.
        self.assertIn("isSectionDivider()", sidebar)
        self.assertIn(':disabled="isSectionDivider()"', sidebar)

    def test_the_section_header_lights_up_for_every_link_it_contains(self):
        """A page highlights its own section header as well as its own link.

        The header carries ``nav|nav_is_any`` listing each child key, so a new
        child link without a matching key would leave the header dark.
        """
        for path, header in (
            ("/venues/", "nav-section-reference"),
            ("/timetable/", "nav-section-timetable"),
            ("/allocation/groups/", "nav-section-allocation"),
        ):
            sidebar = self._sidebar(self._html(path))
            # Slice from the <button> itself: class comes before aria-controls.
            at = sidebar.index('aria-controls="%s"' % header)
            start = sidebar.rindex("<button", 0, at)
            button = sidebar[start: sidebar.index("</button>", at)]
            self.assertIn("is-active", button, path)

    def test_td_allocation_is_reachable_but_not_listed(self):
        """The view works; the sidebar link was dropped as unused."""
        self.assertEqual(self.client.get("/td/").status_code, 200)
        self.assertNotIn("TD Allocations", self._sidebar(self._html("/staff/")))

    def test_a_tap_is_tinted_in_the_theme_instead_of_the_browsers_own_flash(self):
        """A phone tap must look deliberate, the way it does in the portal.

        Left to itself the browser paints a grey-blue box over whatever was
        tapped, which on a themed button reads as a rendering fault, and the
        student portal -- the first thing a user sees -- does not do it. The
        dark chrome cannot carry the app blue either, so it gets a light wash.
        """
        html = self._html("/staff/")
        self.assertIn("-webkit-tap-highlight-color: rgba(49, 93, 131, 0.22)", html)
        self.assertIn(
            ".topbar a, .topbar button, .mobile-bottom-nav a, .mobile-bottom-nav button",
            html,
        )
        self.assertIn("rgba(255, 255, 255, 0.16)", html)
        # :active has to reach the themed hovers, otherwise the press colour is
        # gone the moment the finger lifts and only the flash is left.
        for hover in ("hover:bg-blue-700", "hover:bg-emerald-700", "hover:bg-slate-800/60"):
            block = html[html.index('[class~="%s"]' % hover):]
            self.assertIn(":active", block[: block.index("}") + 1], hover)

    def test_the_phone_top_bar_is_the_height_of_the_student_portals(self):
        """Tapping through from the portal must not feel like the app shrank.

        The portal's bar is `.7rem` of padding around a `2.8rem` menu, so
        anything under 4.25rem is visibly smaller than the page the user just
        came from.

        The height is a `calc()` and not a `h-*` utility because it has to grow
        by the notch inset as well -- see the next test. So the assertion is on
        the rule, not on a class, and it also pins the no-dual-source point: a
        height class back on the element would be a second value for the same
        box and would silently drop the inset.
        """
        html = self._html("/staff/")
        topbar = html[html.index('<header class="topbar'):]
        self.assertNotIn("h-[4.25rem]", topbar[: topbar.index(">")])

        rule = html[html.index("body .topbar {"):]
        self.assertIn("height: calc(4.25rem + env(safe-area-inset-top))", rule[: rule.index("}") + 1])
        self.assertIn("padding-top: env(safe-area-inset-top)", rule[: rule.index("}") + 1])

    def test_the_dashboard_bar_names_the_app_rather_than_the_page(self):
        """On the front door the bar wears the wordmark; elsewhere it names the page.

        "Dashboard" in the bar of the dashboard said nothing a user could not
        already see, while the sidebar rail was already stating the app's name in
        the one form the brand has. Both now include ``partials/brand.html``, so
        the cream-on-white pairing cannot be retyped and drift -- that is the
        reason for the partial, and the reason the pair is asserted on both
        bars rather than just the one being changed.

        The size passed in is the bar's own type scale, a step below the rail's:
        the phone's bar is narrower, and the link is the flex item that has to
        truncate (``text-overflow`` does nothing on the inline span inside it),
        so the wordmark gives way before the issue bell is pushed off the end.
        """
        topbar = self._topbar(self._html("/staff/"))
        self.assertIn('aria-label="CoET Timetable, dashboard"', topbar)
        self.assertIn('<span class="tt-brand-accent">CoET</span> Timetable', topbar)
        # The link is the flex item, so `truncate` belongs to it and the size to
        # the span inside -- `text-overflow` does nothing on the inline span, and
        # a size on both would be two values for one box.
        brand = topbar[topbar.index('class="tt-brand-text'):]
        brand = brand[: brand.index("</a>")]
        self.assertIn("text-base sm:text-lg lg:text-xl", brand)
        link = topbar[: topbar.index('class="tt-brand-text')]
        self.assertIn("truncate min-w-0", link)
        self.assertNotIn("text-base sm:text-lg lg:text-xl", link)
        # It is a link, not a second heading: dashboard.html owns the page's h1,
        # and repeating "Dashboard" in the bar is exactly what is going away.
        self.assertNotIn(">Dashboard</h1>", topbar)

        sidebar = self._sidebar(self._html("/staff/"))
        for shared in ('<span class="tt-brand-accent">CoET</span> Timetable', "tt-brand-text"):
            self.assertIn(shared, sidebar)
            self.assertIn(shared, topbar)

        # Every other page still names itself -- a brand in the bar everywhere
        # would leave the user with no idea which page they are on.
        venues = self._topbar(self._html("/venues/"))
        self.assertIn(">Venues</h1>", venues)
        self.assertNotIn("tt-brand-text", venues)

    def test_the_browsers_own_chrome_is_painted_in_the_theme_not_white(self):
        """The pale URL bar is the one thing that says "this is a web page".

        The OS draws that bar, so the app cannot colour it with CSS -- it has to
        ask, via `theme-color`, and the media variants stop the OS inverting it
        on a dark-mode phone. The light/dark pair matters as much as the base
        tag: one untagged `theme-color` is a white bar on a dark-mode device.
        """
        html = self._html("/staff/")
        self.assertIn('<meta name="theme-color" content="#244b6b">', html)
        for scheme in ("light", "dark"):
            self.assertIn(
                'name="theme-color" media="(prefers-color-scheme: %s)" content="#244b6b"' % scheme,
                html,
            )

    def test_the_app_can_be_installed_and_owns_the_status_bar(self):
        """Standalone mode, so there is no browser furniture between the user and the app.

        `black-translucent` is the iOS tag that actually does anything: it makes
        the status bar transparent so the page's own navy sits behind it. It only
        has that effect alongside `viewport-fit=cover`, which is why both are
        asserted here -- the translucent status bar without the viewport fit just
        leaves a white strip, which is the bug being fixed.
        """
        html = self._html("/staff/")
        for tag in (
            'name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover"',
            'name="apple-mobile-web-app-capable" content="yes"',
            'name="mobile-web-app-capable" content="yes"',
            'name="apple-mobile-web-app-status-bar-style" content="black-translucent"',
        ):
            self.assertIn(tag, html)

        # The area above the document (the rubber-band region when a phone
        # scrolls past the top) is <html>, not <body>, so it has to carry the
        # navy itself or it flashes white where the app's colour should be.
        self.assertIn("html { background: var(--coet-blue-dark); }", html)

    def test_typing_in_a_field_never_flashes_a_tap_tint(self):
        """A tinted flash over a field being typed into is pure noise."""
        html = self._html("/staff/")
        rule = html[html.index("input, textarea, select {"):]
        self.assertIn("-webkit-tap-highlight-color: transparent", rule[: rule.index("}") + 1])

    def test_the_loader_exists_and_ignores_htmx_writes(self):
        html = self._html("/staff/")
        self.assertIn('id="page-loader"', html)
        # An htmx write swaps a fragment; a full-page loader over it would
        # flash on every save.
        self.assertIn("htmxDriven", html)
        self.assertIn("afterRequest", html)
        # A full navigation, a back/forward restore and a failed request all
        # have to be able to take the overlay back down.
        for event in ("beforeunload", "load", "pageshow"):
            self.assertIn("addEventListener('%s'" % event, html)

    def test_the_loader_always_comes_back_down(self):
        """A download never loads a document, so nothing else can clear it.

        A PDF export answers with ``Content-Disposition: attachment``: the browser
        takes the file, abandons the navigation and stays put. ``beforeunload``
        has already armed the overlay by then, and ``load``/``pageshow``/
        ``htmx:afterRequest`` can never fire -- so the spinner sat there for good
        after every export. The overlay is ``pointer-events:none``, which is why
        it read as a page that had frozen rather than one that had wedged.
        """
        html = self._html("/staff/")
        # The backstop, for any path that still manages to arm it.
        self.assertIn("MAX_HOLD_MS", html)
        self.assertIn("setTimeout(hideLoader, MAX_HOLD_MS)", html)
        # And the primary fix: a download link is recognised and left alone.
        self.assertIn("link.hasAttribute('download')", html)

    def test_no_page_arms_the_loader_for_a_pdf_export(self):
        """Every entry point to a PDF carries ``download``.

        The loader's click handler returns early for a link with the attribute,
        and a ``download`` click does not navigate at all, so ``beforeunload``
        never fires either. The export hub used to assign ``window.location``,
        which *is* a navigation and was the one path that still stuck.
        """
        for path in (
            "/export/",
            "/timetable/?programme=%s&semester=%s&year=1"
            % (self.programme.pk, self.semester.pk),
            # The detail pages carry their own export link (detail.html and
            # _detail_content.html), so they are a second and third entry point.
            "/programmes/%s/" % self.programme.pk,
        ):
            html = self._html(path)
            for anchor in re.findall(r"<a\b[^>]*>", html):
                if "timetable.pdf" not in anchor and "timetable-export" not in anchor:
                    continue
                self.assertIn(
                    "download", anchor, f"export link would stick the loader: {anchor}"
                )

        hub = self._html("/export/")
        # Not a navigation, and the click it makes carries the attribute too.
        self.assertNotIn("window.location.href = url", hub)
        self.assertIn("a.download = ''", hub)
        self.assertIn("a.click()", hub)

    def test_the_bell_counts_only_unresolved_issues(self):
        from student_portal.models import CollisionReport

        def report(**kw):
            kw.setdefault("timetable_type", CollisionReport.TimetableType.TEACHING)
            kw.setdefault("semester", self.semester)
            return CollisionReport.objects.create(**kw)

        # The bell lives in the top bar, not the sidebar.
        self.assertIn('href="/staff/collision-reports/"', self._html("/staff/"))
        self.assertNotIn("reported issue", self._html("/staff/"))  # nothing yet

        report(description="Double booked")
        self.assertIn("1 reported issue awaiting review", self._html("/staff/"))

        report(description="Already fixed", status=CollisionReport.Status.RESOLVED)
        self.assertIn("1 reported issue awaiting review", self._html("/staff/"))

        report(description="Second open one")
        self.assertIn("2 reported issues awaiting review", self._html("/staff/"))

    def test_the_bell_degrades_to_zero_instead_of_taking_the_page_down(self):
        """A missing table must cost the badge, not the page.

        The processor swallows the database error so a half-migrated deploy
        still renders every management page. Called directly rather than
        through a page render, because the dashboard's own clash counter is a
        separate query and is allowed to fail loudly on its own terms.
        """
        from core.context_processors import issue_notifications

        with mock.patch("student_portal.models.CollisionReport") as broken:
            broken.objects.exclude.return_value.count.side_effect = DatabaseError(
                "no such table"
            )
            self.assertEqual(
                issue_notifications(self._request()), {"open_issue_count": 0}
            )
        # And the real query still works, counting only unresolved rows.
        self.assertEqual(issue_notifications(self._request()), {"open_issue_count": 0})

    def _request(self):
        from django.test import RequestFactory

        request = RequestFactory().get("/staff/")
        request.user = self.staff
        return request

    def test_the_allocation_page_says_simulate_not_calculate(self):
        html = self._html("/allocation/")
        self.assertIn("Simulate Allocation", html)
        self.assertNotIn("Calculate Allocation", html)
        self.assertIn("Simulating", html)  # the running state, too


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

    def test_the_sidebar_links_the_activity_log_but_does_not_preview_it(self):
        """The latest-events list is gone from the sidebar; the page is not.

        Six log lines under the Activity Log link pushed the Danger Zone -- the
        only way to reach an irreversible action -- off the bottom of a laptop
        sidebar, and a preview is a poor place for it anyway: the full log
        filters and paginates, and every entry is a click away.
        """
        ActivityLog.objects.create(
            action=LogAction.CREATE,
            message="Created Venue ZZZ99",
            resource="Venue",
            target="ZZZ99",
        )
        resp = self.client.get("/")
        html = resp.content.decode()
        start = html.index('id="app-sidebar"')
        sidebar = html[start : html.index("</aside>", start)]
        self.assertIn("Activity Log", sidebar)
        self.assertIn('href="/activity/"', sidebar)
        self.assertNotIn("Created Venue ZZZ99", sidebar)
        # Nothing anywhere in the shell previews the log any more.
        self.assertNotIn("No activity yet.", html)
        # The full log still holds it, and the dashboard still takes ?log=.
        self.assertContains(self.client.get("/activity/"), "Created Venue ZZZ99")

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
        # A workshop is shaded a pale green, and only that cell uses it.
        self.assertEqual(fill_map[workshop_cell], FILL_COLORS["workshop"])
        self.assertEqual(
            [c for (s, e), c in fill_map.items() if (s, e) != workshop_cell
             and c == FILL_COLORS["workshop"]],
            [],
            "the workshop green leaked into a non-workshop cell",
        )
        self.assertIn(FILL_COLORS["td"], fill_map.values())

    def test_workshop_shade_is_clearly_pale_green(self):
        """The workshop fill must read as green, not as an empty slot.

        Regression: the fill was a near-white green, so a workshop block looked
        blank on the page. It has to be recognisably green while still pale
        enough for 6.5pt black text.
        """
        rgb = tuple(int(FILL_COLORS["workshop"][i:i + 2], 16) for i in (1, 3, 5))
        r, g, b = rgb
        self.assertGreater(g, r, "not green-dominant")
        self.assertGreater(g, b, "not green-dominant")
        # Pale: every channel is light, so black text stays readable.
        self.assertGreater(min(rgb), 150, f"too dark to read black text on: {rgb}")
        # And it is not so pale that it disappears against the white page.
        self.assertLess(max(rgb) - min(rgb), 200, f"too saturated: {rgb}")

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
        # The grid and the narrow-screen agenda are two renderings of the same
        # entries, so "once" is per view: a card must not repeat the workshop
        # name that its venue already spells out.
        self.assertIn("tt-agenda", html)
        grid, _, agenda = html.partition('class="tt-agenda')
        self.assertEqual(grid.count("Course: Building"), 1)
        self.assertEqual(agenda.count("Course: Building"), 1)
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


def _pdf_page_streams(pdf):
    """Decode each PDF page's content stream without needing pypdf.

    reportlab writes page content as ASCII85 + Flate, so the words a test wants
    to assert on are not visible in the raw bytes.
    """
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


def _pdf_text(pdf):
    """Every page's decoded content stream, joined, as latin-1 text."""
    return b"\n".join(_pdf_page_streams(pdf)).decode("latin-1")


def _drawn_text_y(text, needle):
    """The ``y`` a string was drawn at, from reportlab's ``Tm`` operator.

    reportlab emits ``1 0 0 1 <x> <y> Tm`` immediately before the text it
    places, so this is the real vertical position on the page -- which is the
    only honest way to assert that a footer is at the *bottom* rather than
    merely present somewhere on the sheet.
    """
    pattern = re.escape(needle).replace("\\ ", "\\s*\\)?\\s*")
    for match in re.finditer(r"1 0 0 1 (-?[\d.]+) (-?[\d.]+) Tm\s*\(" + pattern,
                             text):
        return float(match.group(2))
    return None


class WorkshopRotationKeyTableTests(TestCase):
    """The rotation key is three columns and one row per week block.

    A row has to read as a sentence -- *in weeks 1-7 this group attends this
    workshop* -- so the key is transposed out of the old "one row per group, one
    column per week" shape.
    """

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        self.c1 = StudentGroup.objects.create(programme=self.prog, code="C1")
        self.c2 = StudentGroup.objects.create(programme=self.prog, code="C2")
        self.c3 = StudentGroup.objects.create(programme=self.prog, code="C3")

    def _workshop(self, name, group, day="THURSDAY", period="MORNING",
                  start=None, end=None, week=None):
        return WorkshopAllocation.objects.create(
            semester=self.sem,
            course_code=name,
            workshop=name,
            group_code=group,
            day=day,
            time_period=period,
            start_time=start,
            end_time=end,
            week_start=week[0] if week else None,
            week_end=week[1] if week else None,
            venue="",
        )

    def _pdf(self, group=None):
        from io import BytesIO

        buf = BytesIO()
        if group is not None:
            render_group_timetable(group, self.sem, 1, out=buf)
        else:
            render_programme_timetable(self.prog, self.sem, 1, out=buf)
        return _pdf_text(buf.getvalue())

    def _key_rows(self, text):
        """The key's body rows as ``(weeks, group, workshop)`` triples.

        reportlab draws table cell text as ``(string) Tj`` lines, so the row
        contents can be read off in order once the header is skipped.
        """
        head, sep, body = text.partition("WORKSHOP ROTATION KEY")
        self.assertTrue(sep, "no rotation key in the PDF")
        cells = re.findall(r"\(((?:[^()\\]|\\.)*)\)\s*Tj", body)
        cells = [c for c in cells if c.strip()]
        self.assertEqual(cells[:3], ["Weeks", "Group", "Workshop"], cells[:3])
        rows = cells[3:]
        return [tuple(rows[i : i + 3]) for i in range(0, len(rows) - 2, 3)]

    def test_the_key_has_exactly_three_columns(self):
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        text = self._pdf(self.c1)
        header = re.findall(r"\(((?:[^()\\]|\\.)*)\)\s*Tj", text)
        self.assertIn("Weeks", header)
        self.assertIn("Group", header)
        self.assertIn("Workshop", header)
        # The columns that used to be there must be gone.
        for dropped in ("Programme", "Day", "Course/Session"):
            self.assertNotIn(dropped, header)

    def test_a_two_way_rotation_is_two_readable_rows(self):
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        self.assertEqual(
            self._key_rows(self._pdf(self.c1)),
            [
                ("Week 1-7", "C1", "Electrical"),
                ("Week 8-14", "C1", "Carpentry"),
            ],
        )

    def test_a_three_way_rotation_orders_by_week_not_by_text(self):
        """"Week 15-21" sorts before "Week 8-14" as a string.

        A rotation long enough to reach a two-digit week is the only place that
        shows, and it is exactly the case that would silently print out of
        order.
        """
        self._workshop("Welding", "C1")
        self._workshop("Electrical", "C1")
        self._workshop("M/Tools", "C1")
        rows = self._key_rows(self._pdf(self.c1))
        self.assertEqual(
            rows,
            [
                ("Week 1-7", "C1", "Electrical"),
                ("Week 8-14", "C1", "M/Tools"),
                ("Week 15-21", "C1", "Welding"),
            ],
        )
    def test_each_group_gets_its_own_rows(self):
        """A different schedule earns its own row -- that is the useful case."""
        self._workshop("Carpentry", "C1")
        self._workshop("Electrical", "C1")
        self._workshop("Welding", "C2")
        self._workshop("Building", "C2")
        rows = self._key_rows(self._pdf())
        self.assertEqual(
            rows,
            [
                ("Week 1-7", "C2", "Building"),
                ("Week 1-7", "C1", "Electrical"),
                ("Week 8-14", "C1", "Carpentry"),
                ("Week 8-14", "C2", "Welding"),
            ],
        )

    def test_groups_on_the_same_schedule_share_one_row(self):
        """Not one row per group. Printing C1's weeks and then C2's says the
        same thing twice over; the Group column lists who follows the schedule."""
        for group in ("C1", "C2", "C3"):
            self._workshop("Carpentry", group)
            self._workshop("Electrical", group)
        rows = self._key_rows(self._pdf())
        self.assertEqual(
            rows,
            [
                ("Week 1-7", "C1, C2, C3", "Electrical"),
                ("Week 8-14", "C1, C2, C3", "Carpentry"),
            ],
        )

    def test_a_repeated_programme_prefix_is_collapsed_in_the_group_column(self):
        for code in ("EE C1", "EE C2", "CE A1"):
            StudentGroup.objects.create(programme=self.prog, code=code)
            self._workshop("Carpentry", code)
            self._workshop("Electrical", code)
        rows = self._key_rows(self._pdf())
        self.assertEqual(
            rows,
            [
                ("Week 1-7", "EE C1, C2, CE A1", "Electrical"),
                ("Week 8-14", "EE C1, C2, CE A1", "Carpentry"),
            ],
        )

    def test_explicit_week_ranges_are_kept_verbatim(self):
        self._workshop("Carpentry", "C1", start=time(10, 0), end=time(14, 0),
                       week=(3, 8))
        self._workshop("Electrical", "C1", start=time(10, 0), end=time(14, 0),
                       week=(9, 14))
        self.assertEqual(
            self._key_rows(self._pdf(self.c1)),
            [
                ("Week 3-8", "C1", "Carpentry"),
                ("Week 9-14", "C1", "Electrical"),
            ],
        )

    def test_the_cell_no_longer_repeats_the_week_range_on_a_group_sheet(self):
        """The range came off a cell that cannot act on it.

        A folded workshop cell states the type, the time and the groups; the
        workshop *names* only ever appear in the key below, because the export
        drops the course detail to stop it swamping the box. A week range beside
        that is therefore unreadable noise -- and the key prints the mapping
        properly, one row per block.
        """
        self._workshop("Carpentry", "C1", start=time(10, 0), end=time(14, 0),
                       week=(1, 7))
        self._workshop("Electrical", "C1", start=time(10, 0), end=time(14, 0),
                       week=(8, 14))
        text = self._pdf(self.c1)
        cell = text.partition("WORKSHOP ROTATION KEY")[0]
        self.assertIn("Workshop", cell)
        self.assertNotIn("Wk 1-7", cell)
        self.assertNotIn("Wk 8-14", cell)
        # ...and the key is now the single place the mapping is stated.
        self.assertEqual(
            self._key_rows(text),
            [
                ("Week 1-7", "C1", "Carpentry"),
                ("Week 8-14", "C1", "Electrical"),
            ],
        )

    def test_a_programme_sheet_keeps_the_range_as_a_disambiguator(self):
        """There the range is sometimes the only thing telling two entries apart.

        Two groups doing the *same* workshop in the same slot but in different
        week blocks collapse to one cell. Without the range the two rows print
        identically and nobody can tell which group is in which week.
        """
        for group, start, end in (("C1", 1, 7), ("C2", 8, 14)):
            WorkshopAllocation.objects.create(
                semester=self.sem,
                course_code="Carpentry",
                workshop="Carpentry",
                group_code=group,
                day="THURSDAY",
                time_period="MORNING",
                week_start=start,
                week_end=end,
                venue="",
            )
        text = self._pdf()
        self.assertIn("Wk 1-7", text)
        self.assertIn("Wk 8-14", text)

    def test_the_two_week_ranges_are_not_folded_into_one_wrong_cell(self):
        """Folding them would keep only the first range and claim both groups.

        C1 is in weeks 1-7 and C2 in weeks 8-14. Merged into one cell reading
        "ALL groups, weeks 1-7", the sheet would be wrong about C2 -- and
        neither group rotates here, so this is a cell question, not a key one.
        """
        for group, start, end in (("C1", 1, 7), ("C2", 8, 14)):
            WorkshopAllocation.objects.create(
                semester=self.sem,
                course_code="Carpentry",
                workshop="Carpentry",
                group_code=group,
                day="THURSDAY",
                time_period="MORNING",
                week_start=start,
                week_end=end,
                venue="",
            )
        cell = self._pdf().partition("WORKSHOP ROTATION KEY")[0]
        # Both ranges must survive, and neither may sit beside an "ALL" that
        # would claim a group it does not cover.
        self.assertIn("Wk 1-7", cell)
        self.assertIn("Wk 8-14", cell)
        self.assertNotIn("ALL \\267 Wk 1-7", cell)
        self.assertNotIn("Wk 1-7 \\267 ALL", cell)


class GridExportFooterTests(TestCase):
    """The two grid sheets carry the master timetable's footer, on every page.

    It used to be a flowable appended to the element list, so it printed wherever
    the content happened to end -- near the top of page one for a short
    timetable -- and never appeared on any other page.
    """

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(
            code="EE", name="BSc. in Electrical Engineering"
        )
        self.c1 = StudentGroup.objects.create(programme=self.prog, code="C1")
        self.c2 = StudentGroup.objects.create(programme=self.prog, code="C2")
        self.venue = Venue.objects.create(name="R217", capacity=200)

    def _sessions(self, n=9):
        from datetime import time as _t

        for i in range(n):
            Session.objects.create(
                semester=self.sem,
                course_code="EE150",
                activity_type=ActivityType.TUTORIAL,
                day="MONDAY",
                start_time=_t(7 + (i % 12), 0),
                end_time=_t(8 + (i % 12), 0),
                venue=self.venue,
            )

    def _group_pdf(self, **kwargs):
        from io import BytesIO

        buf = BytesIO()
        render_group_timetable(self.c1, self.sem, 1, out=buf, **kwargs)
        return buf.getvalue()

    def _programme_pdf(self, **kwargs):
        from io import BytesIO

        buf = BytesIO()
        render_programme_timetable(self.prog, self.sem, 1, out=buf, **kwargs)
        return buf.getvalue()

    def test_the_group_sheet_names_the_portal_the_date_and_the_page(self):
        import datetime as _dt

        self._sessions()
        text = _pdf_text(self._group_pdf())
        self.assertIn("Generated from", text)
        self.assertIn("CoET Timetable Portal", text)
        self.assertIn(_dt.date.today().strftime("%d %B %Y"), text)
        self.assertIn("(Page 1)", text)

    def test_the_programme_sheet_carries_the_same_footer(self):
        import datetime as _dt

        self._sessions()
        text = _pdf_text(self._programme_pdf())
        for fragment in (
            "Generated from",
            "CoET Timetable Portal",
            _dt.date.today().strftime("%d %B %Y"),
            "(Page 1)",
        ):
            self.assertIn(fragment, text)

    def test_the_footer_sits_at_the_bottom_of_the_page(self):
        """Not merely present -- down in the margin, where a footer belongs.

        A4 portrait is 842pt tall, so anything above ~60pt is body text. This is
        the assertion that catches the old flowable placement, which put the
        line immediately under the table.
        """
        self._sessions()
        text = _pdf_text(self._group_pdf())
        y = _drawn_text_y(text, "Generated from")
        self.assertIsNotNone(y, "footer text not found in the content stream")
        self.assertLess(y, 60, f"footer drawn at y={y}, which is not the bottom")
        self.assertGreater(y, 0, "footer drawn off the bottom of the page")

    def test_the_footer_repeats_on_every_page(self):
        """A long rotation key forces a second page, and the footer must be on it.

        The old version was a flowable, so it could only ever be on one page --
        and on a short sheet, near the top of the first.
        """
        for n in range(30):
            code = f"X{n:02d}"
            StudentGroup.objects.create(programme=self.prog, code=code)
            # Two workshops per group, so each one really is a rotation and
            # really earns key rows, and a week range unique to that group so
            # the rows cannot collapse together into a page-fitting handful.
            for offset, name in ((1, f"Shop{n:02d}"), (8, f"Shop{(n + 1) % 30:02d}")):
                WorkshopAllocation.objects.create(
                    semester=self.sem,
                    course_code=name,
                    workshop=name,
                    group_code=code,
                    day="THURSDAY",
                    time_period="MORNING",
                    week_start=n * 7 + offset,
                    week_end=n * 7 + offset + 6,
                    venue="",
                )
        self._sessions()
        pdf = self._programme_pdf()
        pages = len(re.findall(rb"/Type\s*/Page[^s]", pdf))
        self.assertGreater(pages, 1, "this fixture must span more than one page")
        text = _pdf_text(pdf)
        self.assertEqual(text.count("Generated from"), pages)
        for number in range(1, pages + 1):
            self.assertIn(f"(Page {number})", text)

    def test_the_portal_name_links_back_to_the_site_on_both_sheets(self):
        for pdf in (
            self._group_pdf(portal_url="http://example.test/"),
            self._programme_pdf(portal_url="http://example.test/"),
        ):
            self.assertEqual(
                pdf.count(b"/URI (http://example.test/)"),
                len(re.findall(rb"/Type\s*/Page[^s]", pdf)),
            )

    def test_no_sheet_still_carries_the_old_floating_line(self):
        self._sessions()
        for pdf in (self._group_pdf(), self._programme_pdf()):
            self.assertNotIn("Prepared for personal use", _pdf_text(pdf))

    def test_the_export_views_hand_the_renderer_their_own_address(self):
        for url in (
            f"/export/groups/{self.c1.pk}/timetable.pdf/",
            f"/export/programmes/{self.prog.pk}/timetable.pdf/",
        ):
            resp = self.client.get(
                url,
                {"semester": self.sem.pk, "year": "1"},
                HTTP_HOST="timetable.test",
            )
            self.assertEqual(resp.status_code, 200, url)
            self.assertIn(b"/URI (http://timetable.test/)", resp.content)


class WorkshopRotationTests(TestCase):

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

    def test_every_listed_whole_cohort_course_reads_all(self):
        """The courses the timetable names: SC121, MT161, MT171, CL111, DS114, DS115.

        Each of these also runs per-group tutorials and seminars, so this checks
        the LECTURE rows only -- the per-group rows of the same course must keep
        naming their own groups.
        """
        for course in ("SC121", "MT161", "MT171", "CL111", "DS114", "DS115"):
            with self.subTest(course=course):
                Session.objects.filter(course_code=course).delete()
                self._session(
                    course, "MONDAY", "08:00", "09:55",
                    groups=[self.ee_c1], activity_type="LECTURE",
                )
                block = [
                    e for e in self._merged() if e["course_code"] == course
                ][0]
                self.assertEqual(block["groups"], "ALL", course)
                self.assertEqual(_entry_parts(block)[-1], ("ALL", True), course)

                # The same course's tutorial is per-group and stays that way.
                self._session(
                    course, "TUESDAY", "10:00", "10:55",
                    groups=[self.ee_c1], activity_type="TUTORIAL",
                )
                tutorial = [
                    e for e in self._merged()
                    if e["course_code"] == course and e["kind"] == "tutorial"
                ][0]
                self.assertEqual(tutorial["groups"], "C1", course)

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
        """One daily workshop slot is ONE session, whoever attends it.

        Every group's workshop in one slot belongs to a single workshop session
        for the week, so Carpentry, Welding and Masonry at Monday 09:00-13:00
        are one block reading "Workshop" with all three groups listed -- not
        three blocks, which buried the group list.
        """
        for group_code, workshop in (
            ("C1", "Carpentry"), ("C2", "Welding"), ("A1", "Masonry"),
        ):
            self._workshop(workshop, group_code, "MONDAY", "09:00", "13:00")
        blocks = self._blocks("workshop")
        self.assertEqual(len(blocks), 1, "one slot must be one workshop block")
        self.assertEqual(blocks[0]["type_label"], "Workshop")
        self.assertEqual(blocks[0]["groups"], "C1, C2, A1")
        self.assertEqual(blocks[0]["hours"], {9, 10, 11, 12})
        # Type, the time it runs, then the groups -- like every other block.
        text = _master_cell_text(blocks, show_groups=True)
        self.assertEqual(
            text.split("\n"), ["Workshop", "09:00\u201313:00", "C1, C2, A1"]
        )
        for detail in ("Carpentry", "Welding", "Masonry", "Workshop Shed"):
            self.assertNotIn(detail, text)

    def test_same_workshop_for_many_groups_becomes_one_block(self):
        """The plain case: one craft, many groups, one block."""
        for group_code in ("C1", "C2", "A1"):
            self._workshop("Carpentry", group_code, "MONDAY", "09:00", "13:00")
        blocks = self._blocks("workshop")
        self.assertEqual(len(blocks), 1, "one workshop must be one block")
        self.assertEqual(blocks[0]["type_label"], "Workshop")
        self.assertEqual(blocks[0]["groups"], "C1, C2, A1")
        text = _master_cell_text(blocks, show_groups=True)
        self.assertEqual(
            text.split("\n"), ["Workshop", "09:00\u201313:00", "C1, C2, A1"]
        )
        for detail in ("Carpentry", "Workshop Shed"):
            self.assertNotIn(detail, text)

    def test_workshop_placeholder_practical_is_replaced_by_the_workshop(self):
        """A four-hour practical named WORKSHOP is the workshop, not a second one.

        This is the duplicate the export used to draw: a grey "WORKSHOP
        Practical" block sitting on top of the green workshop block for the
        very same slot.
        """
        self._workshop("Carpentry", "C1", "MONDAY", "09:00", "13:00")
        self._workshop("Carpentry", "C2", "MONDAY", "09:00", "13:00")
        # 09:00-12:55 rather than 09:00-13:00: the same hours, a different
        # clock string, which is exactly how the duplicate slipped through.
        self._session(
            "WORKSHOP", "MONDAY", "09:00", "12:55",
            activity_type=ActivityType.PRACTICAL,
        )
        blocks = self._blocks("workshop")
        self.assertEqual(len(blocks), 1, "the placeholder must not add a block")
        self.assertEqual(blocks[0]["groups"], "C1, C2")
        practicals = self._blocks("practical")
        self.assertEqual(practicals, [], "the placeholder practical is gone")

    def test_placeholder_with_no_workshop_behind_it_is_kept(self):
        """A WORKSHOP practical with nothing behind it is a real session.

        Dropping it would silently lose a scheduled class, so it stays.
        """
        self._session(
            "WORKSHOP", "MONDAY", "09:00", "12:55",
            activity_type=ActivityType.PRACTICAL,
        )
        blocks = self._blocks("practical")
        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0]["course_code"], "WORKSHOP")

    def test_me101_practical_folds_into_the_me101_technical_drawing(self):
        """An ME101 practical IS the ME101 technical drawing: one block."""
        self._td("ME101", "C1", "WEDNESDAY", "09:00", "12:00")
        self._td("ME101", "C2", "WEDNESDAY", "09:00", "12:00")
        self._session(
            "ME101", "WEDNESDAY", "09:00", "11:55",
            activity_type=ActivityType.PRACTICAL,
        )
        blocks = self._blocks("td")
        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0]["type_label"], "Technical Drawing")
        self.assertEqual(blocks[0]["groups"], "C1, C2")
        self.assertEqual(self._blocks("practical"), [])

    def test_two_groups_doing_one_course_at_one_time_become_one_block(self):
        """Same day, same time, same course: one block listing both groups."""
        for group in (self.ee_c1, self.ee_c2):
            self._session("EE150", "TUESDAY", "10:00", "11:55", groups=[group])
        blocks = self._blocks("lecture") or self._blocks("tutorial")
        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0]["course_code"], "EE150")
        self.assertEqual(len(blocks[0]["groups"].split(",")), 2)

    def test_different_courses_in_one_slot_are_never_merged(self):
        self._session("EE150", "MONDAY", "08:00", "09:55")
        self._session("EE152", "MONDAY", "08:00", "09:55")
        codes = {e["course_code"] for e in self._merged()}
        self.assertEqual(codes, {"EE150", "EE152"})

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

    # -- 3c. the assigned-groups line is emphasised -----------------------
    def test_the_groups_line_is_the_only_emphasised_one(self):
        """Blue/bold/italic marks the groups, and nothing else.

        The flag is decided where the field order is decided, so it identifies
        the line by matching the entry's own group list rather than by assuming
        it comes last.
        """
        self._session(
            "EE150", "MONDAY", "08:00", "09:55",
            groups=[self.ee_c1, self.ee_c2],
        )
        parts = _entry_parts(self._blocks("lecture")[0])
        self.assertEqual(
            [text for text, _ in parts],
            ["Lecture", "08:00\u201309:55", "R217", "EE150", "C1, C2"],
        )
        self.assertEqual([flag for _, flag in parts], [False] * 4 + [True])

    def test_a_technical_drawing_states_its_time_and_emphasises_its_groups(self):
        """A TD block reads type, time, groups -- and the time is stated."""
        self._td("ME101", "C1", "MONDAY", "09:00", "12:00")
        self._td("ME101", "C2", "MONDAY", "09:00", "12:00")
        parts = _entry_parts(self._blocks("td")[0])
        self.assertEqual(
            [text for text, _ in parts],
            ["Technical Drawing", "09:00\u201312:00", "C1, C2"],
        )
        self.assertEqual([flag for _, flag in parts], [False, False, True])

    def test_a_workshop_states_its_time(self):
        """A workshop states the time it runs, like any other block."""
        self._workshop("Carpentry", "C1", "MONDAY", "08:00", "11:00")
        parts = _entry_parts(self._blocks("workshop")[0])
        self.assertEqual(
            [text for text, _ in parts],
            ["Workshop", "08:00\u201311:00", "C1"],
        )
        self.assertEqual([flag for _, flag in parts], [False, False, True])

    def test_a_workshop_without_clock_times_invents_none(self):
        """A period-only workshop (raw matrix import) states no clock time.

        Its hours are only known as grid slots, so no "09:00-13:00" is
        fabricated; the block simply omits the line.
        """
        entry = {
            "kind": "workshop", "course_code": "Carpentry", "name": "Carpentry",
            "type_label": "Workshop", "groups": "C1", "hours": {9, 10, 11, 12},
            "label": "Carpentry", "venue": "", "note": "",
            "start": "", "end": "",
        }
        parts = _entry_parts(entry)
        self.assertEqual([text for text, _ in parts], ["Workshop", "C1"])
        self.assertEqual([flag for _, flag in parts], [False, True])

    def test_a_whole_cohort_lecture_writes_all_as_its_group_line(self):
        """A whole-cohort lecture says so on the page, not by going silent.

        The ``ALL`` marker is the assigned group, so it is drawn and emphasised
        exactly like a list of codes.
        """
        self._session(
            "QS125", "MONDAY", "09:00", "10:55",
            groups=[self.ee_c1, self.ee_c2, self.ce_a1, self.ce_a2],
        )
        block = self._blocks("lecture")[0]
        self.assertEqual(block["groups"], "ALL")
        parts = _entry_parts(block)
        self.assertEqual(parts[-1], ("ALL", True))
        for code in ("C1", "C2", "A1", "A2"):
            self.assertNotIn(code, " ".join(text for text, _ in parts))

    def test_a_block_with_no_groups_emphasises_nothing(self):
        self._session("EE150", "MONDAY", "08:00", "09:55")
        parts = _entry_parts(self._blocks("lecture")[0])
        self.assertTrue(all(not flag for _text, flag in parts))

    def test_the_emphasis_survives_wrapping(self):
        """A group list too wide for its box stays emphasised once wrapped."""
        entry = {
            "kind": "workshop", "course_code": "Carpentry", "name": "Carpentry",
            "type_label": "Workshop",
            "groups": " ".join("EE%d" % i for i in range(1, 14)),
            "hours": {8}, "label": "Carpentry", "venue": "", "note": "",
            "start": "", "end": "",
        }
        flowable = _DayFlowable([entry], "Monday", list(range(7, 20)), 40.0, 22)
        lines = flowable.boxes[0]["lines"]
        self.assertGreater(len(lines), 2, "the group list should have wrapped")
        self.assertFalse(lines[0][1], "the heading must not be emphasised")
        self.assertTrue(all(flag for _text, flag in lines[1:]))

    def test_cell_markup_emphasises_groups_and_escapes_the_rest(self):
        """The classic export styles the groups and escapes every other part."""
        entry = {
            "kind": "lecture", "course_code": "R&D <lab>", "name": "R&D <lab>",
            "type_label": "Lecture", "start": "08:00", "end": "08:55",
            "venue": "A & B", "groups": "EE C1, C2", "hours": {8},
            "label": "R&D <lab> Lecture\n08:00\u201308:55\nA & B\nEE C1, C2",
            "note": "",
        }
        markup = cell_markup([entry], show_groups=True)
        self.assertIn(GROUPS_STYLE["color"], markup)
        self.assertIn("<b><i>EE C1, C2</i></b>", markup)
        # Dangerous characters are escaped, so they can never become tags.
        self.assertIn("R&amp;D &lt;lab&gt;", markup)
        self.assertIn("A &amp; B", markup)
        self.assertNotIn("<lab>", markup)
        # The plain-text form is unchanged and carries no markup at all.
        self.assertEqual(
            cell_text([entry], show_groups=True),
            "R&D <lab> Lecture\n08:00\u201308:55\nA & B\nEE C1, C2",
        )

    def test_cell_markup_emphasises_the_all_marker(self):
        entry = {
            "kind": "lecture", "course_code": "QS125", "name": "QS125",
            "type_label": "Lecture", "start": "09:00", "end": "10:55",
            "venue": "R217", "groups": "ALL", "hours": {9},
            "label": "QS125 Lecture\n09:00\u201310:55\nR217\nALL", "note": "",
        }
        markup = cell_markup([entry], show_groups=True)
        self.assertIn("<b><i>ALL</i></b>", markup)
        self.assertIn(GROUPS_STYLE["color"], markup)

    def test_cell_markup_leaves_a_block_without_groups_untouched(self):
        entry = {
            "kind": "lecture", "course_code": "EE150", "name": "EE150",
            "type_label": "Lecture", "start": "08:00", "end": "08:55",
            "venue": "R217", "groups": "", "hours": {8},
            "label": "EE150 Lecture\n08:00\u201308:55\nR217", "note": "",
        }
        markup = cell_markup([entry], show_groups=True)
        self.assertNotIn(GROUPS_STYLE["color"], markup)
        self.assertNotIn("<i>", markup)

    def test_both_exports_really_draw_the_emphasis(self):
        """The blue fill and the bold-oblique face reach the actual PDF bytes."""
        from io import BytesIO

        # A third EE group the session does NOT attend: without it the session
        # would cover every group in the programme, and a block that needs no
        # group list has nothing for the styling to emphasise.
        ee_c3 = StudentGroup.objects.create(programme=self.ee, code="C3")
        self._session(
            "EE150", "MONDAY", "08:00", "09:55",
            groups=[self.ee_c1, self.ee_c2],
        )
        # Folded blocks must carry a group list for the styling to have anything
        # to emphasise, so this also guards the fold's output shape.
        folded = fold_for_display(
            collect_entries(self.ee, self.sem),
            {self.ee_c1.code, self.ee_c2.code, ee_c3.code},
        )
        self.assertTrue(
            [e for e in folded if e.get("groups") == "C1, C2"],
            f"no folded block carried the group list: "
            f"{[e.get('groups') for e in folded]}",
        )

        blue = int(GROUPS_STYLE["color"][5:7], 16) / 255
        needle = f"{blue:.6f}".lstrip("0").encode()
        self.assertEqual(GROUPS_STYLE["font"], "Helvetica-BoldOblique")

        master = BytesIO()
        render_udsm_master_timetable(
            collect_master_entries(self.sem), self.sem, 1, out=master
        )
        self.assertIn(b"/Helvetica-BoldOblique", master.getvalue())

        classic = BytesIO()
        render_programme_timetable(self.ee, self.sem, 1, out=classic)
        classic_pdf = classic.getvalue()
        self.assertIn(b"/Helvetica-BoldOblique", classic_pdf)

        # The blue reaches the page: it is a fill colour in a content stream.
        for pdf in (master.getvalue(), classic_pdf):
            content = b"\n".join(self._page_streams(pdf))
            self.assertIn(needle, content)

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
            len(boxes[0]["lines"]) * _DayFlowable.LEADING
            + 2 * _DayFlowable.PAD_Y,
        )
        # The band adds the trailing margin that separates it from the next day.
        self.assertAlmostEqual(
            height,
            boxes[0]["bottom"] - boxes[0]["top"] + _DayFlowable.BLOCK_GAP,
        )

    def test_longer_content_is_never_truncated_to_the_floor(self):
        self._session("CL111", "MONDAY", "08:00", "09:55", venue=False)
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        boxes = flowable.boxes
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
        # They are stacked, each one a clear BLOCK_GAP below the previous, so
        # two sessions never meet skin to skin.
        for previous, current in zip(boxes, boxes[1:]):
            self.assertAlmostEqual(
                current["top"], previous["bottom"] + _DayFlowable.BLOCK_GAP
            )
        # The band is only as tall as the stack plus the trailing margin that
        # separates it from the next day -- not three times a row.
        self.assertAlmostEqual(
            height, boxes[-1]["bottom"] + _DayFlowable.BLOCK_GAP
        )

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
        self.assertAlmostEqual(
            height, boxes[0]["bottom"] + _DayFlowable.BLOCK_GAP
        )

    def test_a_wide_block_starts_below_the_tallest_column_it_covers(self):
        """A wide block clears the deepest column it covers, plus the margin."""
        # Two one-hour blocks in hour 8 stack; a two-hour block covering hours
        # 8 and 9 must start below both.
        for i in range(2):
            self._session(f"MT{610 + i}", "MONDAY", "08:00", "08:55")
        self._session("MT620", "MONDAY", "08:00", "09:55")
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        boxes = flowable.boxes
        wide = [b for b in boxes if b["colspan"] == 2][0]
        narrow = [b for b in boxes if b["colspan"] == 1]
        self.assertAlmostEqual(
            wide["top"],
            max(b["bottom"] for b in narrow) + _DayFlowable.BLOCK_GAP,
        )

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
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        boxes = flowable.boxes
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
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        boxes = flowable.boxes
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
    def test_the_shortest_session_is_stacked_on_top(self):
        """A band's blocks read shortest-first, whatever the start times are.

        A four-hour block starting at 08:00 used to own the top of the band and
        push the one-hour block at 10:00 underneath it. The shorter block now
        wins the top, which is what the export is read for.
        """
        self._session("MT501", "MONDAY", "08:00", "11:55")  # four hours
        self._session("MT502", "MONDAY", "10:00", "10:55")  # one hour
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        by_span = {box["colspan"]: box for box in flowable.boxes}
        self.assertEqual(set(by_span), {1, 4})
        self.assertEqual(by_span[1]["top"], 0.0)
        self.assertLess(by_span[1]["top"], by_span[4]["top"])

    def test_stacking_is_shortest_first_across_a_mixed_day(self):
        """Blocks of equal length stay in start-time order among themselves."""
        for course, start, end in (
            ("MT401", "08:00", "10:55"),   # three hours
            ("MT402", "09:00", "10:55"),   # two hours
            ("MT403", "10:00", "10:55"),   # one hour
            ("MT404", "14:00", "14:55"),   # one hour
        ):
            self._session(course, "MONDAY", start, end)
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        rows = sorted(flowable.boxes, key=lambda b: (b["top"], b["col"]))
        # One hour, then one hour, then two, then three.
        self.assertEqual(
            [box["colspan"] for box in rows], [1, 1, 2, 3]
        )
        # The two one-hour blocks sit side by side at the very top.
        self.assertEqual([box["top"] for box in rows[:2]], [0.0, 0.0])
        self.assertNotEqual(rows[0]["col"], rows[1]["col"])

    def test_shortest_first_never_forces_blocks_to_overlap(self):
        """The shorter-on-top rule is a preference, never a layout constraint."""
        for course, start, end in (
            ("MT301", "08:00", "12:55"),   # five hours
            ("MT302", "09:00", "09:55"),   # one hour inside it
            ("MT303", "10:00", "11:55"),   # two hours inside it
            ("MT304", "11:00", "11:55"),   # one hour inside that
        ):
            self._session(course, "MONDAY", start, end)
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        boxes = flowable.boxes
        for i, first in enumerate(boxes):
            for second in boxes[i + 1:]:
                share_columns = (
                    first["col"] < second["col"] + second["colspan"]
                    and second["col"] < first["col"] + first["colspan"]
                )
                if not share_columns:
                    continue
                self.assertFalse(
                    first["top"] < second["bottom"] and second["top"] < first["bottom"],
                    f"boxes overlap: {first} and {second}",
                )

    def test_every_stacked_pair_is_separated_by_the_block_gap(self):
        """No two blocks ever touch: a clear margin is always reserved.

        The packer, not the renderer, is what reserves the space, so this holds
        for any data -- including blocks that only partly share columns.
        """
        for hour in (8, 9, 10):
            for i in range(3):
                self._session(
                    f"MT{760 + hour}{i}", "MONDAY",
                    f"{hour:02d}:00", f"{hour:02d}:55",
                )
        self._session("MT790", "MONDAY", "08:00", "10:55")
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        gap = _DayFlowable.BLOCK_GAP
        self.assertGreater(gap, 0.0)
        checked = 0
        for i, first in enumerate(flowable.boxes):
            for second in flowable.boxes[i + 1:]:
                share_columns = (
                    first["col"] < second["col"] + second["colspan"]
                    and second["col"] < first["col"] + first["colspan"]
                )
                if not share_columns:
                    continue
                upper, lower = sorted((first, second), key=lambda b: b["top"])
                if lower["top"] < upper["bottom"]:
                    continue  # side by side, not stacked
                self.assertGreaterEqual(
                    lower["top"] - upper["bottom"], gap - 1e-6,
                    f"only {lower['top'] - upper['bottom']}pt clear",
                )
                checked += 1
        self.assertGreater(checked, 0, "fixture produced no stacked pairs")

    def test_cell_leading_leaves_the_lines_readable(self):
        """Consecutive lines of a session get real breathing room.

        A block's height is derived from the leading, so this is what stops the
        fields of a session reading as one jammed-together paragraph.
        """
        self.assertGreater(_DayFlowable.LEADING, _DayFlowable.CELL_SIZE)
        self.assertGreaterEqual(
            _DayFlowable.LEADING / _DayFlowable.CELL_SIZE, 1.2
        )
        self._session("CL111", "MONDAY", "08:00", "08:55", venue=False)
        merged = self._merged()
        flowable = _DayFlowable(
            [e for e in merged if e["day"] == "MONDAY"],
            "MONDAY",
            list(range(7, 20)),
            60.0,
            46,
        )
        box = flowable.boxes[0]
        self.assertEqual(
            box["bottom"] - box["top"],
            len(box["lines"]) * _DayFlowable.LEADING + 2 * _DayFlowable.PAD_Y,
        )
        # The block is tall enough for its lines plus the padding, and every
        # line sits LEADING apart rather than jammed against the next.
        self.assertGreaterEqual(len(box["lines"]), 3)
        self.assertEqual(
            box["bottom"] - box["top"],
            len(box["lines"]) * _DayFlowable.LEADING + 2 * _DayFlowable.PAD_Y,
        )

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
        day_flowables, _slots = _build_day_flowables(
            merged, col_width=60.0, day_width=46
        )
        self.assertEqual([flowable.day_label for flowable in day_flowables], ["Monday"])

    def test_split_breaks_between_blocks_and_never_through_one(self):
        """A too-tall day is broken in the gap between two blocks.

        The whole point of the export is that no session is ever sliced through
        the middle and shown half on one page, half on the next. Every block
        must come out whole, exactly once, across the pieces.
        """
        for i in range(60):
            self._session(f"MT{1000 + i}", "MONDAY", "08:00", "08:55")
        merged = self._merged()
        day_flowables, _slots = _build_day_flowables(
            merged, col_width=60.0, day_width=46
        )
        self.assertEqual(len(day_flowables), 1)
        monday = day_flowables[0]
        self.assertGreater(monday.height, 700.0)

        pieces = monday.split(1000, 700.0)
        self.assertGreater(len(pieces), 1)
        self.assertLessEqual(pieces[0].height, 700.0 + 0.01)
        # The continuation is the same day, and the day's name is written on
        # exactly one of the two pages.
        self.assertEqual(pieces[1].day_label, monday.day_label)
        self.assertEqual(
            [p._label_offset() is not None for p in pieces].count(True), 1
        )
        # Every page keeps a bounded day-label column, so a day carried over
        # from the previous page never looks cut off down its side.
        for piece in pieces:
            self.assertGreater(piece.day_width, 0)
        # The two pages abut in whole-day coordinates: no gap and no overlap.
        self.assertAlmostEqual(
            pieces[1].day_offset, pieces[0].day_offset + pieces[0].height
        )

        # Drive reportlab's own loop: keep splitting whatever is still too tall
        # and check that every block lands whole on exactly one page.
        pages, queue, guard = [], list(pieces), 0
        while queue:
            guard += 1
            self.assertLess(guard, 50, "splitting did not converge")
            piece = queue.pop(0)
            if piece.height <= 700.0:
                pages.append(piece)
                continue
            queue = list(piece.split(1000, 700.0)) + queue
        self.assertGreater(len(pages), 1)
        seen = []
        for page in pages:
            self.assertLessEqual(page.height, 700.0 + 0.01)
            for box in page.boxes:
                self.assertGreaterEqual(box["top"], -0.01)
                seen.append((tuple(box["lines"]), box["colspan"]))
        # Nothing lost, nothing duplicated: 60 distinct blocks, once each.
        self.assertEqual(len(seen), 60)
        self.assertEqual(len(set(seen)), 60)
        # The day's name is written on exactly one page, never repeated.
        self.assertEqual(
            sum(1 for page in pages if page._label_offset() is not None), 1
        )
        # And no page overlaps the next in whole-day coordinates.
        for earlier, later in zip(pages, pages[1:]):
            self.assertGreaterEqual(
                later.day_offset, earlier.day_offset + earlier.height - 0.01
            )

    def test_a_day_with_no_room_left_moves_whole_to_the_next_page(self):
        """With no space left, nothing is drawn and the day carries over."""
        for i in range(40):
            self._session(f"MT{1100 + i}", "MONDAY", "08:00", "08:55")
        merged = self._merged()
        day_flowables, _slots = _build_day_flowables(
            merged, col_width=60.0, day_width=46
        )
        monday = day_flowables[0]
        # Too little room for even the first block: reportlab is told "nothing
        # here", so it starts the day again on a fresh page, whole.
        self.assertEqual(monday.split(1000, 4.0), [])
        # And given a full page it is placed in one piece.
        self.assertEqual(monday.split(1000, 5000), [monday])

    def test_a_day_starts_on_a_page_that_still_has_room(self):
        """A day must use the space left on a page, not skip to the next one.

        Regression: a cut line was rejected whenever ANY block crossed it, even
        one in a different hour column that is drawn alongside and cannot
        collide. A densely packed day is a staircase where the tallest block in
        every prefix crosses the next line, so no line qualified at all and the
        whole day jumped to a fresh page, leaving the space below it empty.
        """
        # Three hours, three different block heights, three rows: the day packs
        # as a staircase, where the tallest block in every prefix crosses the
        # next candidate line. Built as plain entries so the shape is exact.
        def entry(course, hour, n_lines):
            return {
                "key": ("session", course), "day": "MONDAY", "hours": {hour},
                "label": course, "kind": "lecture", "course_code": course,
                "name": course, "type_label": "Lecture",
                "venue": "R217" if n_lines >= 4 else None,
                "start": "08:00", "end": "08:55",
                "groups": "EE C1" if n_lines >= 5 else "ALL", "note": "",
            }

        rows = [
            entry(course, hour, n_lines)
            for repeat in range(3)
            for course, hour, n_lines in (
                ("MT20%d" % (1 + repeat), 8, 3),
                ("MT30%d" % (1 + repeat), 10, 4),
                ("MT40%d" % (1 + repeat), 12, 5),
            )
        ]
        monday = _DayFlowable(rows, "Monday", list(range(7, 20)), 60.0, 46)
        gap = _DayFlowable.BLOCK_GAP
        spaces = (120.0, 100.0, 80.0, 60.0)
        # The day is a staircase, and every line is crossed by a block in
        # another column -- so a cut exists only because crossing is allowed.
        self.assertGreater(monday.height, max(spaces))
        self.assertGreater(
            sum(1 for space in spaces if len(monday.split(1000, space)) > 1),
            0,
            "no space in the fixture forces a split, so it proves nothing",
        )

        for space in spaces:
            # How much of the day this space can actually hold, derived
            # independently: the highest line that collides with nothing in a
            # shared column and does not overflow.
            best = None
            for level in sorted({b["top"] for b in monday.boxes}):
                above = [b for b in monday.boxes if b["top"] < level]
                below = [b for b in monday.boxes if b["top"] >= level]
                if not above or not below:
                    continue
                content = max(b["bottom"] for b in above)
                if content + gap > space:
                    continue
                collides = any(
                    a["col"] < c["col"] + c["colspan"]
                    and c["col"] < a["col"] + a["colspan"]
                    and a["bottom"] > c["top"]
                    for a in above
                    for c in below
                )
                if collides:
                    continue
                margin = gap if content + gap <= level else 0.0
                best = content + margin
            pieces = monday.split(1000, space)
            if best is None:
                # Nothing fits: the day moves whole to a fresh page.
                self.assertEqual(pieces, [monday])
                continue
            # The day starts here and the first page takes the most it can.
            self.assertGreaterEqual(len(pieces), 2)
            self.assertAlmostEqual(pieces[0].height, best)
            self.assertLessEqual(pieces[0].height, space)
            # Every block is on exactly one page, and none collides across a
            # seam in a column they share.
            self.assertEqual(sum(len(p.boxes) for p in pieces), len(monday.boxes))
            for first, second in zip(pieces, pieces[1:]):
                for a in first.boxes:
                    for c in second.boxes:
                        share = (
                            a["col"] < c["col"] + c["colspan"]
                            and c["col"] < a["col"] + a["colspan"]
                        )
                        if not share:
                            continue
                        self.assertLessEqual(
                            a["bottom"] + first.day_offset,
                            c["top"] + second.day_offset,
                            f"blocks collide across the seam: {a} / {c}",
                        )
                self.assertGreaterEqual(second.day_offset, first.day_offset)

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

    def test_the_day_flowable_reports_its_full_size(self):
        for day in ("MONDAY", "TUESDAY"):
            self._session(f"MT{1100 + len(day)}", day, "08:00", "08:55")
        merged = self._merged()
        day_flowables, slots = _build_day_flowables(merged, col_width=60.0, day_width=46)
        # Every day flowable reports its own width and height.
        for flowable in day_flowables:
            width, height = flowable.wrap(10000, 10000)
            self.assertAlmostEqual(width, 46 + 60 * len(slots))
            self.assertGreater(height, 0)
            # The band is exactly as tall as the blocks it holds, plus the
            # margin that separates it from the next day.
            self.assertAlmostEqual(
                height,
                max(box["bottom"] for box in flowable.boxes)
                + _DayFlowable.BLOCK_GAP,
            )

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

    def test_page_is_a4_landscape_so_it_needs_no_sideways_scrolling(self):
        """The grid must fit the page at 100% zoom.

        Regression: the export was A3 landscape, whose 1191pt width is far more
        than thirteen short hour columns need, so the reader had to scroll
        sideways to see the timetable at all.
        """
        self._session("AR111", "MONDAY", "08:00", "12:55")
        pdf, _pages, _calls = self._render()
        boxes = re.findall(rb"/MediaBox\s*\[\s*([\d.\-]+)\s+([\d.\-]+)\s+"
                           rb"([\d.\-]+)\s+([\d.\-]+)\s*\]", pdf)
        self.assertTrue(boxes, "no MediaBox in the PDF")
        widths = {round(float(b[2])) for b in boxes}
        heights = {round(float(b[3])) for b in boxes}
        self.assertEqual(widths, {842})
        self.assertEqual(heights, {595})

    def test_footer_names_the_portal_the_date_and_the_page_on_every_page(self):
        """Every page carries the provenance line and its own page number."""
        import datetime as _dt

        self._busy_week()
        pdf, pages, _calls = self._render()
        self.assertGreater(pages, 1, "this fixture must span more than one page")
        content = b"\n".join(self._page_streams(pdf)).decode("latin-1")
        # Once per page: the line, and a page number counting 1..n.
        self.assertEqual(content.count("Generated from"), pages)
        self.assertEqual(content.count("CoET Timetable Portal"), pages)
        for number in range(1, pages + 1):
            self.assertIn(f"(Page {number})", content)
        self.assertNotIn(f"(Page {pages + 1})", content)
        # Today's date, written out, on every page.
        self.assertEqual(
            content.count(_dt.date.today().strftime("%d %B %Y")), pages
        )

    def test_portal_name_is_a_link_back_to_the_site(self):
        """The portal name in the footer is a live link, once per page."""
        from io import BytesIO

        buf = BytesIO()
        render_udsm_master_timetable(
            collect_master_entries(self.sem), self.sem, 1, out=buf,
            portal_url="http://example.test/",
        )
        pdf = buf.getvalue()
        self.assertEqual(pdf.count(b"/URI (http://example.test/)"),
                         len(re.findall(rb"/Type\s*/Page[^s]", pdf)))

    def test_export_view_passes_its_own_address_as_the_portal_link(self):
        """The download view hands the renderer the site it was served from."""
        resp = self.client.get(
            "/export/all-programmes/timetable.pdf/",
            {"semester": self.sem.pk, "year": "1"},
            HTTP_HOST="timetable.test",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp["Content-Type"], "application/pdf")
        # The portal name in the footer links back to this very host.
        self.assertIn(b"/URI (http://timetable.test/)", resp.content)

    def test_every_day_is_closed_by_a_boundary_rule(self):
        """Each day band is closed by one rule, not left hanging open.

        Regression: the day frame stroked a heavy rule along the top but none
        along the bottom, so the hour columns trailed away into nothing under
        the last block. One heavy rule along the bottom edge closes the band
        and doubles as the separator from the day that follows it.
        """
        for day in ("MONDAY", "TUESDAY", "WEDNESDAY"):
            self._session(f"MT{1300 + len(day)}", day, "08:00", "08:55")
        pdf, _pages, _calls = self._render()
        content = b"\n".join(self._page_streams(pdf)).decode("latin-1")

        # reportlab writes "w" for a line width, then "x1 y1 m x2 y2 l S".
        width = None
        rules = []
        for match in re.finditer(
            r"([\d.]+) w|n ([\d.]+) ([\d.]+) m ([\d.]+) ([\d.]+) l S", content
        ):
            if match.group(1) is not None:
                width = float(match.group(1))
                continue
            x1, y1, x2, y2 = (float(match.group(i)) for i in (2, 3, 4, 5))
            if width and width > 1.0 and abs(y1 - y2) < 0.01 and abs(x1 - x2) > 1:
                rules.append((x1, y1, x2))
        # Exactly one heavy rule per day band, drawn at y = 0 inside the band's
        # own coordinate space. Three days therefore produce three rules -- not
        # six, because the same rule separates one day from the next instead of
        # every band stroking a line along both of its edges.
        self.assertEqual(len(rules), 3, rules)
        for x1, y1, x2 in rules:
            # The rule sits in the middle of the band's trailing margin, so it
            # stands equally clear of the day above and the day below.
            self.assertAlmostEqual(y1, _DayFlowable.BLOCK_GAP / 2.0)
            # The rule spans the whole grid, day-label column included, so no
            # hour column is ever left hanging open-ended. The page is A4
            # landscape, so the grid ends just short of 842pt.
            self.assertAlmostEqual(x1, 0.0)
            self.assertGreater(x2, 700.0)
            self.assertLess(x2, 841.89)

    def test_a_continued_day_does_not_repeat_its_name(self):
        """A day carried onto the next page does not name itself again."""
        for i in range(60):
            self._session(f"MT{1400 + i}", "MONDAY", "08:00", "08:55")
        self._session("MT1500", "TUESDAY", "08:00", "08:55")
        pdf, pages, _calls = self._render()
        content = b"\n".join(self._page_streams(pdf)).decode("latin-1")
        self.assertGreater(pages, 1, "this fixture must span more than one page")
        # The name is written exactly once for the whole export, however many
        # pages the day runs on to.
        self.assertEqual(content.count("(Monday)"), 1)
        self.assertEqual(content.count("(Tuesday)"), 1)

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


class ProgrammeExportLayoutTests(TestCase):
    """The whole-programme sheet is the master's packed band layout.

    The programme export used to be a classic grid: a cell per hour, with a
    session spanning five hours drawn as five cells and two sessions in one hour
    stacked into the one box they share. Both are unreadable for a real
    programme, where a whole cohort is in a lecture at 08:00 while six groups
    are in six different rooms at the same hour. It now uses the same
    ``_DayFlowable`` packing as the all-programmes export -- days down a narrow
    left column, hours across the top, and each block laid down on the lowest
    free position across the hours it covers, so nothing is ever sliced or
    overlapped. The single-group sheet is deliberately left as the classic
    portrait grid: one group's week has no simultaneous sessions to separate.
    """

    def setUp(self):
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(
            code="CE", name="BSc. in Civil Engineering"
        )
        self.groups = [
            StudentGroup.objects.create(programme=self.prog, code="C%d" % i)
            for i in range(1, 8)
        ]
        self.venue = Venue.objects.create(name="YOMBO5", capacity=200)

    def _session(self, code, day, start, end, groups=(), activity="LECTURE", venue=None):
        session = Session.objects.create(
            semester=self.sem,
            course_code=code,
            activity_type=activity,
            day=day,
            start_time=start,
            end_time=end,
            venue=venue or self.venue,
        )
        for group in groups:
            SessionGroup.objects.create(session=session, group=group)
        return session

    def _entries(self):
        return fold_for_display(
            collect_entries(self.prog, self.sem, year=1), [g.code for g in self.groups]
        )

    def test_the_programme_sheet_is_the_packed_landscape_layout(self):
        # Six groups in six rooms at the same hour, plus a cohort lecture on
        # top: the case the classic grid could not draw at all.
        for i, group in enumerate(self.groups[:6]):
            self._session(
                "AR%d" % (i + 1), "MONDAY", "08:00", "10:55", groups=[group],
                activity="TUTORIAL", venue=Venue.objects.create(
                    name="ROOM%d" % i, capacity=40
                ),
            )
        # A whole-cohort lecture has to be linked to every group to be collected
        # at all -- ``collect_entries`` selects on ``session_groups``, so an
        # unlinked session belongs to the master sheet, not this one.
        self._session("MT171", "MONDAY", "08:00", "09:55", groups=self.groups)

        day_flowables, _slots = _build_day_flowables(self._entries(), 60.0, 46)
        self.assertEqual(len(day_flowables), 1)
        monday = day_flowables[0]
        self.assertEqual(monday.day_label, "Monday")

        # Every block is whole: no session is cut into one box per hour, and the
        # 08:00-09:55 lecture is a single block rather than two.
        spans = [box["colspan"] for box in monday.boxes]
        self.assertIn(3, spans, "the 3-hour lecture was not drawn as one block")
        self.assertEqual(len(monday.boxes), 7)

        # And nothing overlaps. Two blocks may share an hour column only if one
        # ends before the other starts.
        for i, upper in enumerate(monday.boxes):
            for lower in monday.boxes[i + 1:]:
                same_columns = (
                    upper["col"] < lower["col"] + lower["colspan"]
                    and lower["col"] < upper["col"] + upper["colspan"]
                )
                if not same_columns:
                    continue
                self.assertFalse(
                    upper["top"] < lower["bottom"] and lower["top"] < upper["bottom"],
                    f"blocks overlap: {upper['lines']} / {lower['lines']}",
                )

    def test_the_programme_export_renders_a_packed_pdf_that_keeps_its_key(self):
        self._session("AR111", "MONDAY", "08:00", "10:55", groups=self.groups[:3])
        # A real rotation is one *group* doing two *different* workshops in the
        # same slot on different week blocks. Two groups each on one block is
        # not a rotation and earns no key row, so the fixture below is the only
        # shape that can prove the key survived the change of engine.
        for workshop, (start, end) in (("Carpentry", (1, 7)), ("Welding", (8, 14))):
            WorkshopAllocation.objects.create(
                semester=self.sem, group_code="C1", day="TUESDAY",
                start_time="08:00", end_time="11:00", venue="YOMBO5",
                workshop=workshop, week_start=start, week_end=end,
            )
        buf = io.BytesIO()
        render_programme_timetable(self.prog, self.sem, 1, out=buf)
        pdf = buf.getvalue()
        self.assertTrue(pdf.startswith(b"%PDF-"))

        text = _pdf_text(pdf)
        self.assertIn("WORKSHOP ROTATION KEY", text)
        # The key states the mapping the folded cell cannot: which craft runs in
        # which weeks.
        self.assertIn("Carpentry", text)
        self.assertIn("Welding", text)
        self.assertIn("1-7", text.replace("\\226", "-"))
        self.assertIn("8-14", text.replace("\\226", "-"))
        # Packed band, not the classic grid: an hour header naming the columns,
        # and day names drawn rotated in their own left column.
        self.assertIn("07:00", text)
        self.assertIn("Monday", text)
        # A4 landscape, like the master sheet it now shares an engine with.
        boxes = re.findall(
            rb"/MediaBox\s*\[\s*([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s*\]",
            pdf,
        )
        self.assertTrue(boxes, "no MediaBox in the PDF")
        self.assertEqual({round(float(b[2])) for b in boxes}, {842})
        self.assertEqual({round(float(b[3])) for b in boxes}, {595})

    def test_the_single_group_sheet_still_uses_the_classic_portrait_grid(self):
        """One group's week has no simultaneity, so the grid is right for it."""
        self._session("AR111", "MONDAY", "08:00", "10:55", groups=self.groups[:1])
        buf = io.BytesIO()
        render_group_timetable(self.groups[0], self.sem, 1, out=buf)
        pdf = buf.getvalue()

        boxes = re.findall(
            rb"/MediaBox\s*\[\s*([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s+([\d.\-]+)\s*\]",
            pdf,
        )
        self.assertTrue(boxes, "no MediaBox in the PDF")
        # Portrait, where the master and programme sheets are landscape.
        self.assertEqual({round(float(b[2])) for b in boxes}, {595})
        self.assertEqual({round(float(b[3])) for b in boxes}, {842})

    def test_both_packed_exports_share_one_engine(self):
        """Two renderers, one packer -- so they cannot drift apart."""
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate

        calls = []

        def fake_bands(*args, **kwargs):
            calls.append(kwargs)
            SimpleDocTemplate(io.BytesIO(), pagesize=A4).build([])

        with mock.patch("core.timetable_pdf._render_day_bands", fake_bands):
            render_programme_timetable(self.prog, self.sem, 1, out=io.BytesIO())
            render_udsm_master_timetable([], self.sem, 1, out=io.BytesIO())

        self.assertEqual(len(calls), 2)
        # The programme sheet keeps the workshop week range (two groups in one
        # slot on different week blocks must stay tellable apart); the master
        # sheet, which has no rotation key, drops it.
        self.assertIs(calls[0]["show_notes"], True)
        self.assertIs(calls[1]["show_notes"], False)

    def test_the_week_range_survives_into_the_programme_cell(self):
        """The range is the only thing telling two groups in one slot apart.

        ``fold_same_sessions`` buckets on ``(slot, identity, note)``, so a
        differing week range is a differing block -- the two groups must not be
        merged into one box, or the cell would keep whichever range was seen
        first and print "Wk 1-7" against everyone. The packed band draws them
        as two adjacent blocks, which is also the only way both can be read.
        """
        self._session("AR111", "MONDAY", "08:00", "10:55", groups=self.groups[:1])
        for code, (start, end) in (("C1", (1, 7)), ("C2", (8, 14))):
            WorkshopAllocation.objects.create(
                semester=self.sem, group_code=code, day="TUESDAY",
                start_time="08:00", end_time="11:00", venue="YOMBO5",
                workshop="Carpentry", week_start=start, week_end=end,
            )
        entries = self._entries()
        workshop = [e for e in entries if e["kind"] == "workshop"]
        self.assertEqual(len(workshop), 2, "the week range must keep them apart")

        with_notes, _ = _build_day_flowables(entries, 60.0, 46, show_notes=True)
        without_notes, _ = _build_day_flowables(entries, 60.0, 46, show_notes=False)
        noted = [ln for f in with_notes for b in f.boxes for ln in b["lines"]]
        plain = [ln for f in without_notes for b in f.boxes for ln in b["lines"]]
        # Both ranges, on the programme sheet...
        self.assertTrue(any("1-7" in ln[0] for ln in noted), noted)
        self.assertTrue(any("8-14" in ln[0] for ln in noted), noted)
        # ...and neither on the single-group one, whose key states them instead.
        self.assertFalse(any("1-7" in ln[0] for ln in plain), plain)
        self.assertFalse(any("8-14" in ln[0] for ln in plain), plain)


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
            activity_type="TUTORIAL",
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
        self.assertIn("Tutorial, 08:00-09:00, Mon", html)
        self.assertIn("Course: MT171", html)
        self.assertIn("Venue: NB102", html)
        self.assertIn("Assigned groups: C1", html)
        self.assertNotIn("Assigned groups: N/A", html)

    def _grouped(self, course_code, activity_type, day, start, end, group=None):
        """A session linked to a group, so it survives the programme filter.

        A programme view only shows sessions that have ``SessionGroup`` links,
        so every fixture below needs one or the page renders empty.
        """
        session = Session.objects.create(
            semester=self.sem,
            course_code=course_code,
            activity_type=activity_type,
            day=day,
            start_time=start,
            end_time=end,
            venue=self.venue,
        )
        SessionGroup.objects.create(session=session, group=group or self.g)
        return session

    def _timetable_html(self, **params):
        resp = self.client.get(
            "/timetable/",
            {"programme": self.prog.pk, "semester": self.sem.pk, **params},
            HTTP_HOST="localhost",
        )
        self.assertEqual(resp.status_code, 200)
        return resp.content.decode()

    def test_lecture_card_omits_the_assigned_groups_line(self):
        # A whole-cohort lecture is not "assigned" to anyone, so the line is
        # left out of the box entirely rather than printed as "ALL" or a list.
        # (The legend still names the concept, hence the check is on the card.)
        self._grouped("MT171", "LECTURE", "MONDAY", "08:00", "09:00")
        html = self._timetable_html()
        self.assertIn("Lecture, 08:00-09:00, Mon", html)
        self.assertIn("Course: MT171", html)
        self.assertIn("Venue: NB102", html)
        self.assertIn("tt-kind-lecture", html)
        self.assertNotIn('class="tt-line groups"', html)

    def test_assigned_groups_line_is_emphasised_in_blue(self):
        # The groups line is the one a reader looks for, so it is drawn in the
        # same blue the PDF exports use (GROUPS_STYLE in core/timetable_grid).
        self._grouped("MT171", "SEMINAR", "MONDAY", "08:00", "09:00")
        html = self._timetable_html()
        self.assertIn('class="tt-line groups"', html)
        self.assertIn("Assigned groups: C1", html)
        self.assertRegex(html, r"\.tt-card \.groups \{[^}]*color:\s*#0b4f9e")

    def test_activity_card_groups_line_blank_without_groups(self):
        Session.objects.create(
            semester=self.sem,
            course_code="MT171",
            activity_type="TUTORIAL",
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

    def test_grid_claims_the_full_page_width(self):
        # The container is uncapped and the grid's overflow floor is the sum of
        # its column minimums, so the table grows into whatever the sidebar
        # leaves rather than stopping at a fixed max width.
        self._grouped("MT171", "LECTURE", "MONDAY", "08:00", "09:00")
        html = self._timetable_html()
        self.assertIn("#tt-page { width: 100%; max-width: none; }", html)
        self.assertIn("tt-table--grid", html)
        self.assertIn("min-width: calc(var(--tt-day) + var(--tt-hour) * var(--tt-slot-count))", html)
        # The overflow floor is fed the real number of hourly columns.
        self.assertRegex(html, r'--tt-slot-count: \d+')
        # The old 72rem cap on this page's own wrapper is gone (the base
        # template still ships a max-w-6xl phone rule for every other page).
        self.assertNotIn("max-width: 72rem", html)
        self.assertNotIn('class="max-w-6xl', html)
        # Collapsing the sidebar is what hands the grid its extra width, so the
        # collapsed state has to be allowed to use it.
        self.assertIn("html.tt-collapsed .tt-scroll", html)

    def test_narrow_screens_get_an_agenda_of_the_same_sessions(self):
        self._grouped("MT171", "LECTURE", "MONDAY", "08:00", "09:00")
        self._grouped("CL111", "SEMINAR", "MONDAY", "10:00", "11:00")
        html = self._timetable_html()
        self.assertIn("tt-agenda", html)
        # Both views are rendered from the same grid, so a session that is in
        # the grid is in the agenda too.
        self.assertIn("tt-agenda-day", html)
        self.assertGreaterEqual(html.count("tt-card"), 4)
        # A day with nothing on it says so instead of leaving a hole.
        self.assertIn("Nothing scheduled", html)
        # The grid is what prints; the agenda would print every session twice.
        self.assertIn(".tt-agenda { display: none !important; }", html)

    def test_day_entries_are_chronological_for_the_agenda(self):
        late = self._entry(
            key=("session", 2), day="MONDAY", hours={13}, course_code="B",
            name="B", start="13:00", end="14:00",
        )
        early = self._entry(
            key=("session", 1), day="MONDAY", hours={8, 9}, course_code="A",
            name="A", start="08:00", end="09:55",
        )
        row = build_day_time_grid([late, early])["rows"][0]
        self.assertEqual([e["course_code"] for e in row["entries"]], ["A", "B"])
        # The lanes keep the same two sessions; the agenda list is an extra
        # ordering of them, not a different set.
        self.assertEqual(
            sorted(
                e["course_code"]
                for lane in row["lanes"]
                for cell in lane
                if not cell.get("empty")
                for e in cell["entries"]
            ),
            ["A", "B"],
        )

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

    def test_no_template_leaks_its_own_comment_markup(self):
        """`{# #}` is single-line; a multi-line one renders as visible text.

        Django's short comment syntax does not span newlines -- only
        `{% comment %}` does. A long explanatory note written as
        `{# line one\n   line two #}` is therefore emitted into the page
        verbatim, which is how a navbar once showed a sentence about the
        issue bell. This walks every template and fails on the pattern, so
        the mistake cannot come back silently.
        """
        from pathlib import Path

        # BASE_DIR, not django.__file__: walking the interpreter's
        # site-packages would scan Django's own templates and find nothing.
        root = Path(settings.BASE_DIR)
        offenders = []
        for path in sorted(root.rglob("*.html")):
            if any(
                part in {".venv", "venv", ".git", "node_modules", "__pycache__"}
                for part in path.parts
            ):
                continue
            text = path.read_text(encoding="utf-8", errors="replace")
            # An opener whose closing `#}` is not on the same line.
            for match in re.finditer(r"\{#", text):
                line_end = text.find("\n", match.start())
                line = text[match.start(): line_end if line_end != -1 else len(text)]
                if "#}" not in line:
                    rel = path.relative_to(root).as_posix()
                    offenders.append("%s: %s" % (rel, line.strip()[:70]))
                    break
        self.assertEqual(offenders, [], "multi-line {# #} comments leak to the page")

    def test_every_template_actually_compiles(self):
        """One unparsable template takes down every page that includes it.

        ``sidebar.html`` once closed a ``{% comment %}`` with ``#}``, which
        Django cannot parse: the template fails to *compile*, so every staff
        page raised TemplateSyntaxError rather than merely rendering oddly.
        Nothing caught it because the only tests that rendered a page never
        got past the login redirect, so the bug sat in the tree being shipped.

        Compiling rather than rendering is the point -- a template can be
        unparsable and still be on disk looking fine.
        """
        from django.template.loader import get_template

        root = Path(settings.BASE_DIR)
        unparsable = []
        for path in sorted(root.rglob("*.html")):
            if any(
                part in {".venv", "venv", ".git", "node_modules", "__pycache__"}
                for part in path.parts
            ):
                continue
            # Templates are addressed by name, and a stray .html under a
            # directory that is not a template dir has no name to load by.
            try:
                name = path.relative_to(root / "templates").as_posix()
            except ValueError:
                continue
            try:
                get_template(name)
            except Exception as exc:  # TemplateSyntaxError and friends
                unparsable.append("%s: %s" % (name, str(exc).splitlines()[0][:80]))
        self.assertEqual(unparsable, [], "templates that do not compile")


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

    def test_list_pages_show_no_clear_all_button(self):
        """The destructive action moved to the Danger Zone, off the list pages.

        It used to sit in `list.html`'s toolbar, next to a search box and a set
        of filters, which is the worst possible place for "delete everything" --
        and the filters never limited it, so the amber note had to warn about
        them. Nothing in the list UI may offer it now.
        """
        for url in (
            "/venues/",
            "/programmes/",
            "/courses/",
            "/groups/",
            "/sessions/",
            "/workshops/",
            "/td/",
            "/course-requirements/",
        ):
            resp = self.client.get(url)
            html = resp.content.decode()
            self.assertEqual(resp.status_code, 200, url)
            self.assertNotIn("Clear All", html, url)
            self.assertNotIn("clear-all", html, url)
            # The toolbar itself is untouched: records are still added and counted.
            self.assertIn("Add New", html, url)
            self.assertIn('id="record-total"', html, url)

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
        self.assertIn("not limited", html)
        plain = self._confirm("/venues/clear-all/")
        self.assertNotIn("Filtered view", plain)


class DangerZoneTests(TestCase):
    """The one page holding every 'clear all' action.

    The actions themselves are unchanged -- they are the same endpoints, the
    same confirmation modal and the same typed phrase, and the two ClearAll
    classes above still drive them. What is asserted here is that the board is
    the only place they are reachable from, that it lists every one of them
    with a live count, and that the sidebar gets staff to it.
    """

    ALL_TARGETS = (
        ("/programmes/clear-all/", "Programmes"),
        ("/groups/clear-all/", "Student Groups"),
        ("/courses/clear-all/", "Programme Courses"),
        ("/course-requirements/clear-all/", "Courses"),
        ("/venues/clear-all/", "Venues"),
        ("/sessions/clear-all/", "Sessions"),
        ("/workshops/clear-all/", "Workshop Allocations"),
        ("/td/clear-all/", "Technical Drawing Allocations"),
    )

    def setUp(self):
        self.client = Client()
        # The portal middleware 302s an anonymous client away from every staff
        # URL, so log in as staff or every request here measures the redirect.
        self.client.force_login(
            User.objects.create_user("danger-tester", password="pw", is_staff=True)
        )
        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)
        self.prog = Programme.objects.create(code="CE", name="Civil Engineering")
        self.g1 = StudentGroup.objects.create(programme=self.prog, code="A1")
        self.venue = Venue.objects.create(name="NB102", capacity=100)
        self.mapping = ProgrammeCourse.objects.create(
            programme=self.prog, course_code="MT161", course_name="Maths 1", semester=1
        )
        # Saving the link created the shared course record; both must be clearable.
        self.course = Course.objects.get(code="MT161")
        self.session = Session.objects.create(
            semester=self.sem,
            course_code="MT161",
            activity_type="LECTURE",
            day="MONDAY",
            start_time="08:00",
            end_time="10:00",
            venue=self.venue,
        )
        self.link = SessionGroup.objects.create(session=self.session, group=self.g1)
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

    def test_the_board_offers_every_clear_all_and_nothing_else(self):
        resp = self.client.get("/danger-zone/")
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode()
        for clear_url, label in self.ALL_TARGETS:
            with self.subTest(target=clear_url):
                self.assertIn("openModal('%s'" % clear_url, html)
                self.assertIn("Clear all %s" % label, html)
        # One card per target, so the count matches the registry exactly.
        # Counted on the onclick attribute: base.html also defines openModal().
        self.assertEqual(html.count('onclick="openModal('), len(self.ALL_TARGETS))
        # The board is a page about irreversible actions, so it says so.
        self.assertIn("There is no undo", html)
        self.assertIn("across the whole database", html)

    def test_each_card_shows_the_live_count_of_what_it_would_delete(self):
        html = self.client.get("/danger-zone/").content.decode()
        self.assertIn("Every programme, and with it its student groups", html)
        groups = self.client.get("/danger-zone/").context["groups"]
        # Counts come from the same registry the endpoint deletes through.
        for group in groups:
            for target in group["targets"]:
                with self.subTest(target=target["key"]):
                    model = CLEAR_ALL_TARGETS[target["key"]]["model"]
                    self.assertEqual(target["count"], model.objects.count())
                    # A cascading target says so before the modal opens.
                    self.assertEqual(
                        target["has_records"],
                        bool(
                            target["count"]
                            or target["related_total"]
                            or target["detached_total"]
                        ),
                    )

    def test_the_board_refreshes_its_counts_after_a_clear(self):
        """A board that still says "1 session" after the clear would be a lie."""
        html = self.client.get("/danger-zone/").content.decode()
        # A swappable fragment with a stable id is what `refresh-table` targets.
        self.assertIn('id="danger-zone-body"', html)
        self.assertIn("hx-get=\"/danger-zone/\"", html)
        self.assertIn("refresh from:window", html)
        self.assertIn("htmx.trigger('#danger-zone-body', 'refresh')", html)
        # The htmx GET serves the same fragment on its own, like a list view's table.
        fragment = self.client.get("/danger-zone/", HTTP_HX_REQUEST="true")
        self.assertEqual(fragment.status_code, 200)
        body = fragment.content.decode()
        self.assertIn('id="danger-zone-body"', body)
        self.assertNotIn("<html", body)

    def test_a_card_with_nothing_to_delete_offers_nothing(self):
        for model in (
            SessionGroup,
            Session,
            WorkshopAllocation,
            TechnicalDrawingAllocation,
            ProgrammeCourse,
            Course,
            StudentGroup,
            Programme,
        ):
            model.objects.all().delete()
        html = self.client.get("/danger-zone/").content.decode()
        self.assertIn("Nothing to clear", html)
        for clear_url, _label in self.ALL_TARGETS:
            if clear_url == "/venues/clear-all/":
                continue
            self.assertNotIn("openModal('%s'" % clear_url, html)
        # The one record left is the venue, and it is still clearable.
        self.assertIn("openModal('/venues/clear-all/'", html)
        self.assertEqual(html.count('onclick="openModal('), 1)

    def test_clearing_from_the_board_runs_the_unchanged_endpoint(self):
        """The board is a new door onto the old flow, not a second one."""
        token = self.client.get("/danger-zone/").cookies["csrftoken"].value
        preview = self.client.get(
            "/sessions/clear-all/", HTTP_HX_REQUEST="true"
        ).content.decode()
        self.assertIn("csrfmiddlewaretoken", preview)
        self.assertIn("DELETE ALL", preview)
        resp = self.client.post(
            "/sessions/clear-all/",
            {"phrase": "DELETE ALL", "csrfmiddlewaretoken": token},
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.headers["HX-Trigger"], "close-modal,refresh-table")
        self.assertEqual(Session.objects.count(), 0)
        self.assertEqual(SessionGroup.objects.count(), 0)
        log = ActivityLog.objects.get(action=LogAction.CLEAR)
        self.assertIn("Cleared all Sessions", log.message)
        # And the board now reports the sessions as gone.
        self.assertNotIn("openModal('/sessions/clear-all/'", self.client.get("/danger-zone/").content.decode())

    def test_the_sidebar_links_here_from_the_bottom(self):
        html = self.client.get("/danger-zone/").content.decode()
        start = html.index('id="app-sidebar"')
        sidebar = html[start : html.index("</aside>", start)]
        self.assertIn('href="/danger-zone/"', sidebar)
        self.assertIn(">Clear all data</span>", sidebar)
        # The section header names the area, the link names the action.
        self.assertIn(">Danger Zone</div>", sidebar)
        # It is the last link in the nav: getting here must not look routine.
        self.assertLess(sidebar.index('href="/activity/"'), sidebar.index('href="/danger-zone/"'))
        self.assertGreater(sidebar.index('href="/danger-zone/"'), sidebar.index('href="/allocation/groups/"'))
        # The Danger Zone link is the one sidebar item tinted red.
        self.assertIn("#nav-danger-zone:hover", html)
        self.assertIn("#nav-danger-zone.bg-slate-800", html)

    def test_the_link_lights_up_only_on_its_own_page(self):
        for path, lit in (("/danger-zone/", True), ("/activity/", False), ("/venues/", False)):
            with self.subTest(path=path):
                html = self.client.get(path).content.decode()
                start = html.index('id="app-sidebar"')
                sidebar = html[start : html.index("</aside>", start)]
                at = sidebar.index("nav-danger-zone")
                link = sidebar[sidebar.rindex("<a ", 0, at) : sidebar.index("</a>", at)]
                self.assertEqual("bg-slate-800 text-white" in link, lit, path)


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

    def test_import_hub_hides_the_td_card_but_keeps_it_importable(self):
        """The TD upload card is withdrawn from the hub, not deleted.

        Hiding it is a presentation choice; the importer and the upload
        endpoint keep working, so a coordinator who needs it can still post to
        /import/upload/td-allocation/ directly.
        """
        hub = self.client.get("/import/").content.decode()
        self.assertNotIn("TD Allocation", hub)
        self.assertNotIn("import/td-allocation/", hub)
        # Every other type still gets its card.
        for key in ("programmes", "venues", "master-timetable", "workshop-allocation"):
            self.assertIn("import/%s/" % key, hub)
        # The type is still registered, so the upload view does not 400 with
        # "Unknown import type" and its GET form still renders.
        from core.views import IMPORT_TYPES

        self.assertIn("td-allocation", IMPORT_TYPES)
        form = self.client.get("/import/td-allocation/")
        self.assertEqual(form.status_code, 200)
        self.assertIn("TD Allocation", form.content.decode())


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


# --------------------------------------------------------------------------
# Shared courses, required activities and group allocation
# --------------------------------------------------------------------------


# Shared courses and required activities
# --------------------------------------------------------------------------


class CourseCodeNormalisationTests(TestCase):
    def test_trims_and_uppercases(self):
        self.assertEqual(normalise_course_code("  mt161 "), "MT161")
        self.assertEqual(normalise_course_code("Mt 161"), "MT 161")
        self.assertEqual(normalise_course_code(None), "")

    def test_code_is_normalised_on_save(self):
        course = Course.objects.create(code="  ee131 ", name="Electronics")
        self.assertEqual(course.code, "EE131")
        self.assertEqual(Course.objects.get(pk=course.pk).code, "EE131")


class SharedCourseModelTests(TestCase):
    def test_blank_requirements_means_no_allocation_needed(self):
        course = Course.objects.create(code="MT161", name="Mathematics 1")
        self.assertEqual(course.requirements, ())
        self.assertEqual(course.required_activities(), ())
        self.assertFalse(course.has_requirements())
        self.assertEqual(course.activities_label(), "")

    def test_requirements_support_a_combination(self):
        course = Course.objects.create(code="TG201", name="Technical Drawing 1")
        course.set_requirements({"TUTORIAL": 1, "PRACTICAL": 1})
        self.assertEqual(
            course.required_activities(), ("TUTORIAL", "PRACTICAL")
        )
        self.assertEqual(course.activities_label(), "Tutorial; Practical")

    def test_requirements_are_returned_in_priority_order(self):
        course = Course.objects.create(code="CL111")
        course.set_requirements({"PRACTICAL": 1, "SEMINAR": 1, "TUTORIAL": 1})
        self.assertEqual(
            course.required_activities(),
            ("SEMINAR", "TUTORIAL", "PRACTICAL"),
        )

    def test_requirements_carry_a_count(self):
        course = Course.objects.create(code="ME101")
        course.set_requirements({"PRACTICAL": 2})
        self.assertEqual(course.required_count("PRACTICAL"), 2)
        self.assertEqual(course.activities_label(), "2x Practical")

    def test_lectures_and_workshops_are_never_requirements(self):
        course = Course.objects.create(code="XX100")
        course.set_requirements({"LECTURE": 1, "WORKSHOP": 1, "SEMINAR": 1})
        self.assertEqual(course.required_activities(), ("SEMINAR",))

    def test_set_requirements_replaces_rather_than_accumulates(self):
        course = Course.objects.create(code="XX100")
        course.set_requirements({"TUTORIAL": 1, "PRACTICAL": 1})
        course.set_requirements({"SEMINAR": 1})
        self.assertEqual(course.required_activities(), ("SEMINAR",))

    def test_requirement_row_rejects_a_lecture(self):
        course = Course.objects.create(code="XX100")
        row = CourseActivityRequirement(
            course=course, activity_type="LECTURE", count=1
        )
        with self.assertRaises(ValidationError):
            row.clean()

    def test_name_conflicts_are_kept_for_review(self):
        course = Course.objects.create(
            code="CL111",
            name="Communication Skills for Engineers",
            name_variants=json.dumps(["Communication Skills for Engineering"]),
        )
        self.assertEqual(
            course.name_conflicts(), ["Communication Skills for Engineering"]
        )


class ProgrammeCourseLinkTests(TestCase):
    def setUp(self):
        self.prog = Programme.objects.create(code="CE", name="Civil Engineering")
        self.course = Course.objects.create(code="MT161", name="Mathematics 1")

    def test_code_and_name_mirror_the_shared_course(self):
        link = ProgrammeCourse.objects.create(
            programme=self.prog, course=self.course, semester=1
        )
        self.assertEqual(link.course_code, "MT161")
        self.assertEqual(link.course_name, "Mathematics 1")

    def test_renaming_the_shared_course_updates_every_link(self):
        ProgrammeCourse.objects.create(
            programme=self.prog, course=self.course, semester=1
        )
        Programme.objects.create(code="ME", name="Mechanical")
        me = Programme.objects.get(code="ME")
        ProgrammeCourse.objects.create(
            programme=me, course=self.course, semester=1
        )
        self.course.name = "Mathematics I"
        self.course.save()
        link = ProgrammeCourse.objects.get(programme=self.prog)
        link.refresh_from_db()
        self.assertEqual(link.course_name, "Mathematics I")
        self.assertEqual(
            ProgrammeCourse.objects.filter(course_name="Mathematics I").count(), 2
        )

    def test_a_shared_course_cannot_be_deleted_out_from_under_a_programme(self):
        ProgrammeCourse.objects.create(
            programme=self.prog, course=self.course, semester=1
        )
        with self.assertRaises(ProtectedError):
            self.course.delete()

    def test_form_only_exposes_the_shared_course(self):
        form = ProgrammeCourseForm(
            data={
                "programme": self.prog.pk,
                "course": self.course.pk,
                "semester": 1,
            }
        )
        self.assertTrue(form.is_valid(), form.errors)
        self.assertNotIn("course_code", form.fields)
        self.assertNotIn("course_name", form.fields)

    def test_form_rejects_a_duplicate_programme_course_semester(self):
        ProgrammeCourse.objects.create(
            programme=self.prog, course=self.course, semester=1
        )
        form = ProgrammeCourseForm(
            data={
                "programme": self.prog.pk,
                "course": self.course.pk,
                "semester": 1,
            }
        )
        self.assertFalse(form.is_valid())
        self.assertIn("semester", form.errors)


class RequiredActivitiesParsingTests(TestCase):
    def test_blank_means_no_requirement(self):
        self.assertEqual(parse_requirements(""), ({}, []))
        self.assertEqual(parse_requirements(None), ({}, []))

    def test_every_spelling_of_no_requirement(self):
        for text in (
            "-", "--", "N/A", "n/a", "NA", "none", "None", "nil", "no",
            "not required", "None required", "no requirements",
            "not applicable", "no allocation", "No allocation required",
            "TBC", "  ",
        ):
            counts, problems = parse_requirements(text)
            self.assertEqual((counts, problems), ({}, []), text)

    def test_single_activity(self):
        self.assertEqual(parse_requirements("Tutorial"), ({"TUTORIAL": 1}, []))

    def test_combinations_with_every_separator(self):
        for text in (
            "Tutorial; Practical",
            "Tutorial, Practical",
            "Tutorial + Practical",
            "Tutorial and Practical",
            "Seminar & Tutorial",
            "Tutorial | Practical",
            "Seminar / Practical",
            "Seminar\nPractical",
            "Seminar\r\nPractical",
            "• Tutorial; • Practical",
        ):
            counts, problems = parse_requirements(text)
            self.assertEqual(problems, [], text)
            self.assertEqual(len(counts), 2, text)

    def test_an_unseparated_pair_is_still_two_activities(self):
        # A cell pasted out of a table often loses its separators.
        for text in ("Seminar Tutorial", "Practical Seminar", "TUTORIAL SEMINAR"):
            counts, problems = parse_requirements(text)
            self.assertEqual(problems, [], text)
            self.assertEqual(len(counts), 2, text)

    def test_run_together_words_are_reported_rather_than_guessed(self):
        # "TutorialTutorial" is not a real spreadsheet value, and guessing at it
        # is the one thing that must not happen.
        counts, problems = parse_requirements("TutorialTutorial")
        self.assertEqual(counts, {})
        self.assertEqual(len(problems), 1)

    def test_case_and_plural_variants(self):
        counts, problems = parse_requirements("tutorials; PRACTICALS")
        self.assertEqual(problems, [])
        self.assertEqual(counts, {"TUTORIAL": 1, "PRACTICAL": 1})

    def test_abbreviations(self):
        for text, expected in (
            ("Tut", {"TUTORIAL": 1}),
            ("Prac", {"PRACTICAL": 1}),
            ("Pract", {"PRACTICAL": 1}),
            ("Sem", {"SEMINAR": 1}),
        ):
            counts, problems = parse_requirements(text)
            self.assertEqual(problems, [], text)
            self.assertEqual(counts, expected, text)

    def test_every_way_of_writing_a_count(self):
        """A count that is read as 1 is a silently mis-allocated course."""
        for text in (
            "2 practicals",
            "2 x Practical",
            "2x Practical",
            "Practical x2",
            "Practical (2)",
            "Practical [2]",
            "Practical 2",
            "2 times practicals",
        ):
            counts, problems = parse_requirements(text)
            self.assertEqual(problems, [], text)
            self.assertEqual(counts, {"PRACTICAL": 2}, text)

    def test_a_count_can_appear_after_a_noun_phrase(self):
        counts, problems = parse_requirements("Tutorials per week: 2")
        self.assertEqual(problems, [])
        self.assertEqual(counts, {"TUTORIAL": 2})

    def test_spelled_out_numbers(self):
        counts, problems = parse_requirements("Two practicals")
        self.assertEqual(problems, [])
        self.assertEqual(counts, {"PRACTICAL": 2})

    def test_count_with_a_combination(self):
        counts, problems = parse_requirements("One seminar; 2 practicals")
        self.assertEqual(problems, [])
        self.assertEqual(counts, {"SEMINAR": 1, "PRACTICAL": 2})

    def test_counted_combination_round_trips_through_the_importer_format(self):
        counts, _ = parse_requirements("Seminar (1); Practical (2)")
        self.assertEqual(counts, {"SEMINAR": 1, "PRACTICAL": 2})
        self.assertEqual(
            format_requirements(counts), "Seminar; 2x Practical"
        )

    def test_repeated_word_adds_up(self):
        counts, _ = parse_requirements("Tutorial; Tutorial")
        self.assertEqual(counts, {"TUTORIAL": 2})

    def test_html_escaped_separators_from_a_web_paste(self):
        counts, problems = parse_requirements("Tutorial &amp; Practical")
        self.assertEqual(problems, [])
        self.assertEqual(counts, {"TUTORIAL": 1, "PRACTICAL": 1})

    def test_unicode_dashes_do_not_break_a_count(self):
        counts, problems = parse_requirements("Practical – 2")
        self.assertEqual(problems, [])
        self.assertEqual(counts, {"PRACTICAL": 2})

    def test_unknown_words_are_reported_not_dropped(self):
        counts, problems = parse_requirements("Tutorial; Fieldwork")
        self.assertEqual(counts, {"TUTORIAL": 1})
        self.assertEqual(len(problems), 1)
        self.assertIn("Fieldwork", problems[0])

    def test_a_lecture_or_workshop_is_never_a_requirement(self):
        for text in ("Lecture", "Workshop", "Workshop Training"):
            counts, problems = parse_requirements(text)
            self.assertEqual(counts, {}, text)
            self.assertEqual(len(problems), 1, text)

    def test_semester_is_not_read_as_a_seminar(self):
        counts, problems = parse_requirements("Semester 2")
        self.assertEqual(counts, {})
        self.assertEqual(len(problems), 1)

    def test_format_round_trip(self):
        for mapping in (
            {"SEMINAR": 1},
            {"TUTORIAL": 1, "PRACTICAL": 1},
            {"PRACTICAL": 2},
            {"SEMINAR": 1, "TUTORIAL": 1, "PRACTICAL": 1},
            {"TUTORIAL": 2, "PRACTICAL": 3},
        ):
            label = format_requirements(mapping)
            counts, problems = parse_requirements(label)
            self.assertEqual(problems, [], label)
            self.assertEqual(counts, mapping, label)

    def test_format_of_nothing_is_blank(self):
        self.assertEqual(format_requirements({}), "")


class RequirementsColumnDetectionTests(TestCase):
    BASE = ["Programme", "Course Code", "Course Name", "Semester"]

    def test_documented_headings_are_found(self):
        for header in (
            "Required Activities",
            "Required_Activities",
            "REQUIREDACTIVITIES",
            "Required  Activities",
            "Allocation Requirements",
            "Alloc Requirements",
            "Activities Required",
            "Required Activity Types",
            "Required Sessions",
            "Required Classes",
        ):
            column, unrecognised = find_requirements_column(self.BASE + [header])
            self.assertEqual(column, header)
            self.assertEqual(unrecognised, [], header)

    def test_an_unlisted_but_sensible_heading_is_still_found(self):
        for header in (
            "Allocated Activities",
            "Required contact hours",
            "Weird: Req. Activities Col",
        ):
            column, _ = find_requirements_column(self.BASE + [header])
            self.assertEqual(column, header)

    def test_other_columns_are_never_mistaken_for_it(self):
        for header in ("Semester", "Course Name", "Programme", "Course Code"):
            column, unrecognised = find_requirements_column(self.BASE + [header])
            self.assertIsNone(column, header)
            self.assertEqual(unrecognised, [], header)

    def test_semester_does_not_look_like_a_requirement_column(self):
        # "Semester" contains "sem"; word boundaries keep it out of the report.
        _, unrecognised = find_requirements_column(
            self.BASE + ["Required Activities"]
        )
        self.assertEqual(unrecognised, [])

    def test_a_workbook_with_no_requirements_column_is_not_an_error(self):
        column, unrecognised = find_requirements_column(self.BASE)
        self.assertIsNone(column)
        self.assertEqual(unrecognised, [])

    def test_an_unrecognised_requirements_heading_is_reported(self):
        # The whole point: never silently import nothing.
        column, unrecognised = find_requirements_column(
            self.BASE + ["Semesterly Required Stuff"]
        )
        self.assertIsNone(column)
        self.assertEqual(unrecognised, ["Semesterly Required Stuff"])

    def test_an_explicit_header_wins(self):
        column, _ = find_requirements_column(
            self.BASE + ["Allocated Activities"], explicit="Allocated Activities"
        )
        self.assertEqual(column, "Allocated Activities")

    def test_an_explicit_header_is_matched_loosely(self):
        column, _ = find_requirements_column(
            self.BASE + ["Allocated Activities"], explicit="allocated_activities"
        )
        self.assertEqual(column, "Allocated Activities")

    def test_an_explicit_header_that_is_absent_returns_nothing(self):
        column, unrecognised = find_requirements_column(
            self.BASE, explicit="Does Not Exist"
        )
        self.assertIsNone(column)
        self.assertEqual(unrecognised, [])

    def test_dedicated_count_columns_are_found(self):
        found = find_requirement_count_columns(
            self.BASE
            + ["Required Activities", "Tutorial Count", "Number of Practicals"],
            exclude=["Required Activities"],
        )
        self.assertEqual(
            found,
            {ActivityType.TUTORIAL: "Tutorial Count",
             ActivityType.PRACTICAL: "Number of Practicals"},
        )

    def test_the_combined_column_is_never_read_as_a_count_column(self):
        found = find_requirement_count_columns(
            self.BASE + ["Required Activities"], exclude=["Required Activities"]
        )
        self.assertEqual(found, {})

    def test_count_column_words(self):
        found = find_requirement_count_columns(
            self.BASE + ["Tutorials per week", "Practical sessions"]
        )
        self.assertEqual(
            found,
            {ActivityType.TUTORIAL: "Tutorials per week",
             ActivityType.PRACTICAL: "Practical sessions"},
        )


class ProgrammeCoursesImportRequirementsTests(ImporterTestCase):
    COLS = [
        "programme_code",
        "course_code",
        "course_name",
        "semester",
        "Required Activities",
    ]

    def test_import_populates_the_shared_course(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Seminar; Tutorial"]],
            self.COLS,
        )
        result = import_programme_courses_from_excel(path)
        self.assertFalse(result.errors, result.errors)
        course = Course.objects.get(code="MT161")
        self.assertEqual(course.name, "Mathematics 1")
        self.assertEqual(
            course.required_activities(), ("SEMINAR", "TUTORIAL")
        )
        self.assertIn("MT161", result.courses_updated)

    def test_course_is_created_when_missing(self):
        self._seed()
        path = make_xlsx(
            [["CE", "ZZ999", "Brand New", 1, "Practical"]], self.COLS
        )
        import_programme_courses_from_excel(path)
        course = Course.objects.get(code="ZZ999")
        self.assertEqual(course.required_activities(), ("PRACTICAL",))

    def test_counted_requirements_are_kept(self):
        self._seed()
        path = make_xlsx(
            [["CE", "ZZ999", "Brand New", 1, "2 practicals"]], self.COLS
        )
        import_programme_courses_from_excel(path)
        self.assertEqual(
            Course.objects.get(code="ZZ999").required_count("PRACTICAL"), 2
        )

    def test_column_aliases_are_accepted(self):
        self._seed()
        for header in (
            "required activities",
            "Required_Activities",
            "allocation requirements",
            "Activities Required",
        ):
            path = make_xlsx(
                [["CE", "MT161", "Mathematics 1", 1, "Seminar"]],
                ["programme_code", "course_code", "course_name", "semester", header],
            )
            import_programme_courses_from_excel(path)
            self.assertTrue(
                Course.objects.get(code="MT161").has_requirements(), header
            )
            Course.objects.get(code="MT161").set_requirements({})

    def test_older_workbook_without_the_column_still_imports(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1]],
            ["programme_code", "course_code", "course_name", "semester"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertFalse(result.errors, result.errors)
        self.assertEqual(result.created + result.updated, 1)
        self.assertFalse(Course.objects.get(code="MT161").has_requirements())

    def test_invalid_activity_value_is_reported_and_sets_nothing(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Tutorial; Fieldwork"]],
            self.COLS,
        )
        result = import_programme_courses_from_excel(path)
        self.assertEqual(len(result.invalid_requirements), 1)
        self.assertIn("Fieldwork", result.invalid_requirements[0])
        # The course still exists and the programme link is still written.
        self.assertFalse(Course.objects.get(code="MT161").has_requirements())
        self.assertTrue(ProgrammeCourse.objects.filter(course_code="MT161").exists())

    def test_conflicting_names_in_one_workbook_are_reported(self):
        self._seed()
        path = make_xlsx(
            [
                ["CE", "MT161", "Mathematics 1", 1, "Tutorial"],
                ["ME", "MT161", "Mathematics One", 1, "Tutorial"],
            ],
            self.COLS,
        )
        result = import_programme_courses_from_excel(path)
        self.assertEqual(len(result.conflicts), 1)
        self.assertIn("conflicting names", result.conflicts[0])
        # Both programmes still get their link.
        self.assertEqual(ProgrammeCourse.objects.filter(course_code="MT161").count(), 2)

    def test_conflicting_requirements_in_one_workbook_are_reported(self):
        self._seed()
        path = make_xlsx(
            [
                ["CE", "MT161", "Mathematics 1", 1, "Tutorial"],
                ["ME", "MT161", "Mathematics 1", 1, "Practical"],
            ],
            self.COLS,
        )
        result = import_programme_courses_from_excel(path)
        self.assertEqual(len(result.conflicts), 1)
        self.assertIn("conflicting required activities", result.conflicts[0])

    def test_a_conflicted_name_is_not_silently_replaced_on_reimport(self):
        self._seed()
        path = make_xlsx(
            [
                ["CE", "MT161", "Mathematics 1", 1, "Tutorial"],
                ["ME", "MT161", "Mathematics One", 1, "Tutorial"],
            ],
            self.COLS,
        )
        import_programme_courses_from_excel(path)
        course = Course.objects.get(code="MT161")
        self.assertEqual(course.name, "Mathematics 1")
        self.assertIn("Mathematics One", course.name_conflicts())
        # Re-importing the same file changes nothing at all.
        before = (course.name, tuple(course.name_conflicts()))
        import_programme_courses_from_excel(path)
        course.refresh_from_db()
        self.assertEqual(
            (course.name, tuple(course.name_conflicts())), before
        )

    def test_requirements_are_written_once_per_course_not_per_row(self):
        self._seed()
        path = make_xlsx(
            [
                ["CE", "MT161", "Mathematics 1", 1, "Tutorial; Practical"],
                ["ME", "MT161", "Mathematics 1", 1, ""],
            ],
            self.COLS,
        )
        import_programme_courses_from_excel(path)
        self.assertEqual(
            Course.objects.get(code="MT161").required_activities(),
            ("TUTORIAL", "PRACTICAL"),
        )

    def test_a_blank_requirements_column_never_wipes_requirements(self):
        self._seed()
        course = Course.objects.get(code="MT161")
        course.set_requirements({"TUTORIAL": 1})
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, ""]], self.COLS
        )
        import_programme_courses_from_excel(path)
        self.assertEqual(
            Course.objects.get(code="MT161").required_activities(), ("TUTORIAL",)
        )

    def test_import_is_idempotent(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Seminar"]], self.COLS
        )
        import_programme_courses_from_excel(path)
        second = import_programme_courses_from_excel(path)
        self.assertEqual(second.created, 0)
        self.assertEqual(
            ProgrammeCourse.objects.filter(course_code="MT161").count(), 2
        )

    def test_course_codes_are_normalised_on_import(self):
        self._seed()
        path = make_xlsx(
            [["CE", "  mt161 ", "Mathematics 1", 1, "Tutorial"]], self.COLS
        )
        import_programme_courses_from_excel(path)
        self.assertTrue(Course.objects.filter(code="MT161").exists())
        self.assertFalse(Course.objects.filter(code="  mt161 ").exists())

    def test_issue_lists_reach_the_import_snapshot(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Nonsense"]], self.COLS
        )
        snapshot = import_programme_courses_from_excel(path).snapshot()
        self.assertEqual(len(snapshot["invalid_requirements"]), 1)
        self.assertIn("courses_created", snapshot)
        self.assertIn("courses_updated", snapshot)
        self.assertIn("course_requirements", snapshot)
        self.assertEqual(snapshot["requirements_column"], "Required Activities")

    def test_the_column_used_is_reported(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Tutorial"]],
            ["programme_code", "course_code", "course_name", "semester",
             "Allocated Activities"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertEqual(result.requirements_column, "Allocated Activities")
        self.assertEqual(
            Course.objects.get(code="MT161").required_activities(), ("TUTORIAL",)
        )

    def test_the_parsed_requirements_are_reported_per_row(self):
        self._seed()
        path = make_xlsx(
            [
                ["CE", "MT161", "Mathematics 1", 1, "Seminar; Tutorial"],
                ["CE", "ZZ999", "Counted", 1, "2 practicals"],
                ["CE", "YY888", "None needed", 1, "-"],
            ],
            self.COLS,
        )
        result = import_programme_courses_from_excel(path)
        parsed = {row["code"]: row["requirements"] for row in result.course_requirements}
        self.assertEqual(parsed["MT161"], "Seminar; Tutorial")
        self.assertEqual(parsed["ZZ999"], "2x Practical")
        self.assertEqual(parsed["YY888"], "(none)")

    def test_an_explicit_column_can_be_named(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Tutorial"]],
            ["programme_code", "course_code", "course_name", "semester",
             "Column X"],
        )
        result = import_programme_courses_from_excel(
            path, requirements_column="Column X"
        )
        self.assertFalse(result.errors, result.errors)
        self.assertEqual(result.requirements_column, "Column X")
        self.assertEqual(
            Course.objects.get(code="MT161").required_activities(), ("TUTORIAL",)
        )

    def test_an_explicit_column_that_does_not_exist_is_a_clear_error(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Tutorial"]], self.COLS
        )
        result = import_programme_courses_from_excel(
            path, requirements_column="Nope"
        )
        self.assertEqual(len(result.errors), 1)
        self.assertIn("'Nope' not found", result.errors[0])
        self.assertIn("Required Activities", result.errors[0])

    def test_an_unrecognised_requirements_heading_is_reported_not_ignored(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Tutorial"]],
            ["programme_code", "course_code", "course_name", "semester",
             "Semesterly Required Stuff"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertTrue(
            any("was not recognised" in e for e in result.errors), result.errors
        )
        self.assertTrue(
            any("Semesterly Required Stuff" in e for e in result.errors)
        )
        self.assertTrue(
            any("--requirements-column" in e for e in result.errors)
        )

    def test_a_workbook_with_no_requirements_column_warns(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1]],
            ["programme_code", "course_code", "course_name", "semester"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertFalse(result.errors, result.errors)
        self.assertTrue(
            any("no seminar/tutorial/practical requirement" in w for w in result.warnings)
        )

    def test_dedicated_count_columns_are_imported(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Seminar", 2]],
            self.COLS + ["Tutorial Count"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertFalse(result.errors, result.errors)
        self.assertEqual(
            result.requirement_count_columns,
            {"TUTORIAL": "Tutorial Count"},
        )
        course = Course.objects.get(code="MT161")
        self.assertEqual(course.required_activities(), ("SEMINAR", "TUTORIAL"))
        self.assertEqual(course.required_count("TUTORIAL"), 2)

    def test_a_count_column_adds_to_the_combined_cell(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Tutorial", 1]],
            self.COLS + ["Number of Practicals"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertFalse(result.errors, result.errors)
        course = Course.objects.get(code="MT161")
        self.assertEqual(course.required_activities(), ("TUTORIAL", "PRACTICAL"))
        self.assertEqual(course.required_count("PRACTICAL"), 1)

    def test_a_count_column_that_is_not_a_number_is_reported(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "", "two"]],
            self.COLS + ["Tutorial Count"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertEqual(len(result.invalid_requirements), 1)
        self.assertIn("not a whole number", result.invalid_requirements[0])
        self.assertFalse(Course.objects.get(code="MT161").has_requirements())

    def test_a_blank_count_column_is_simply_no_opinion(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Seminar", ""]],
            self.COLS + ["Tutorial Count"],
        )
        result = import_programme_courses_from_excel(path)
        self.assertFalse(result.invalid_requirements, result.invalid_requirements)
        self.assertEqual(
            Course.objects.get(code="MT161").required_activities(), ("SEMINAR",)
        )

    def test_counts_survive_a_real_import(self):
        self._seed()
        path = make_xlsx(
            [["CE", "MT161", "Mathematics 1", 1, "Seminar (1); Practical (2)"]],
            self.COLS,
        )
        result = import_programme_courses_from_excel(path)
        self.assertFalse(result.errors, result.errors)
        self.assertEqual(result.invalid_requirements, [])
        course = Course.objects.get(code="MT161")
        self.assertEqual(course.required_activities(), ("SEMINAR", "PRACTICAL"))
        self.assertEqual(course.required_count("PRACTICAL"), 2)


# --------------------------------------------------------------------------
# Group allocation
# --------------------------------------------------------------------------


STUDENTS_PER_GROUP = 30


class AllocationTestCase(TestCase):
    """Two programmes, four groups, one course that requires a tutorial.

    ``GH1`` seats 200 (all four groups), ``LH1`` seats 90 (three groups),
    ``NB102`` seats 30 (one group) and ``GAP`` records no capacity at all, so
    every capacity outcome — all fit, some fit, none fit, unknown — is one
    argument away.
    """

    def _seed(self, *, semester_number=1, requirements=None):
        self.semester = Semester.objects.create(
            academic_year="2026/2027", semester=semester_number
        )
        self.other_semester = Semester.objects.create(
            academic_year="2026/2027", semester=2
        )
        self.ce = Programme.objects.create(code="CE", name="Civil Engineering")
        self.me = Programme.objects.create(code="ME", name="Mechanical Engineering")
        self.a1 = StudentGroup.objects.create(programme=self.ce, code="A1")
        self.a2 = StudentGroup.objects.create(programme=self.ce, code="A2")
        self.d1 = StudentGroup.objects.create(programme=self.me, code="D1")
        self.d2 = StudentGroup.objects.create(programme=self.me, code="D2")
        self.maths = Course.objects.create(code="MT161", name="Mathematics 1")
        for programme in (self.ce, self.me):
            ProgrammeCourse.objects.create(
                programme=programme, course=self.maths, semester=semester_number
            )
        self.maths.set_requirements(requirements or {"TUTORIAL": 1})
        self.hall = Venue.objects.create(name="GH1", capacity=200)
        self.big = Venue.objects.create(name="LH1", capacity=90)
        self.small = Venue.objects.create(name="NB102", capacity=30)
        self.no_capacity = Venue.objects.create(name="GAP", capacity=0)
        return self.semester

    @staticmethod
    def _at(value):
        """``9`` -> 09:00, ``(9, 30)`` -> 09:30. Keeps the tests readable."""
        if isinstance(value, tuple):
            return time(*value)
        return time(value, 0)

    def _session(self, course_code, activity, day, start, end, venue=None):
        return Session.objects.create(
            semester=self.semester,
            course_code=course_code,
            activity_type=activity,
            day=day,
            start_time=self._at(start),
            end_time=self._at(end),
            venue=venue,
        )

    def _tutorial(self, day, start, end, venue=None, course="MT161"):
        return self._session(
            course, ActivityType.TUTORIAL, day, start, end, venue
        )


class RequirementBuildingTests(AllocationTestCase):
    def test_every_group_gets_a_requirement_per_configured_activity(self):
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        requirements, _ = build_requirements(self.semester, "ALL")
        self.assertEqual(len(requirements), 8)  # 4 groups x 2 activities
        self.assertEqual(
            {r.activity_type for r in requirements},
            {ActivityType.TUTORIAL, ActivityType.PRACTICAL},
        )

    def test_a_counted_requirement_makes_one_requirement_per_count(self):
        self._seed(requirements={"PRACTICAL": 2})
        requirements, _ = build_requirements(self.semester, "ALL")
        self.assertEqual(len(requirements), 8)
        self.assertEqual(sorted(r.ordinal for r in requirements[:2]), [1, 2])

    def test_activities_run_in_priority_order(self):
        self._seed(requirements={"PRACTICAL": 1, "SEMINAR": 1, "TUTORIAL": 1})
        requirements, _ = build_requirements(self.semester, "ALL")
        order = []
        for requirement in requirements:
            if requirement.activity_type not in order:
                order.append(requirement.activity_type)
        self.assertEqual(
            order, [ActivityType.SEMINAR, ActivityType.TUTORIAL, ActivityType.PRACTICAL]
        )

    def test_the_most_constrained_requirement_is_searched_first(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9)
        self._tutorial("TUESDAY", 8, 9)
        self._tutorial("WEDNESDAY", 8, 9)
        # ME101 has exactly one session, so it must come first.
        drawing = Course.objects.create(code="ME101", name="Technical Drawing 1")
        for programme in (self.ce, self.me):
            ProgrammeCourse.objects.create(
                programme=programme, course=drawing, semester=1
            )
        drawing.set_requirements({"TUTORIAL": 1})
        self._tutorial("THURSDAY", 8, 9, course="ME101")
        requirements, _ = build_requirements(self.semester, "ALL")
        self.assertEqual(requirements[0].course.code, "ME101")

    def test_a_course_with_blank_requirements_is_reported_not_allocated(self):
        self._seed()
        other = Course.objects.create(code="ZZ999", name="Unconfigured")
        for programme in (self.ce, self.me):
            ProgrammeCourse.objects.create(
                programme=programme, course=other, semester=1
            )
        requirements, unconfigured = build_requirements(self.semester, "ALL")
        self.assertNotIn("ZZ999", {r.course.code for r in requirements})
        self.assertEqual([e["code"] for e in unconfigured], ["ZZ999"])
        self.assertEqual(unconfigured[0]["group_count"], 4)

    def test_the_scope_limits_which_activities_are_built(self):
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        requirements, _ = build_requirements(self.semester, "TUTORIAL")
        self.assertEqual(
            {r.activity_type for r in requirements}, {ActivityType.TUTORIAL}
        )

    def test_groups_of_other_semesters_are_left_out(self):
        self._seed()
        other = Course.objects.create(code="MT171", name="Calculus")
        for programme in (self.ce, self.me):
            ProgrammeCourse.objects.create(
                programme=programme, course=other, semester=2
            )
        other.set_requirements({"TUTORIAL": 1})
        requirements, _ = build_requirements(self.semester, "ALL")
        self.assertNotIn("MT171", {r.course.code for r in requirements})


class ValidatorTests(AllocationTestCase):
    def test_a_group_studying_the_course_may_attend_its_required_activity(self):
        self._seed()
        session = self._tutorial("MONDAY", 8, 9, self.big)
        index = AvailabilityIndex(self.semester)
        self.assertEqual(
            validate_assignment(self.a1, self.maths, session, self.semester, index), []
        )

    def test_a_group_that_does_not_study_the_course_is_refused(self):
        self._seed()
        outsider = Course.objects.create(code="EE131", name="Electronics")
        session = self._tutorial("MONDAY", 8, 9, self.big, course="EE131")
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, outsider, session, self.semester, index
        )
        self.assertIn("group-not-studying", [p.code for p in problems])

    def test_an_activity_the_course_does_not_require_is_refused(self):
        self._seed()
        practical = self._session(
            "MT161", ActivityType.PRACTICAL, "MONDAY", 8, 9, self.big
        )
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, self.maths, practical, self.semester, index
        )
        self.assertIn("activity-not-required", [p.code for p in problems])

    def test_a_workshop_session_is_never_allocatable(self):
        self._seed()
        workshop = self._session(
            "MT161", ActivityType.WORKSHOP, "MONDAY", 9, 13, self.big
        )
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, self.maths, workshop, self.semester, index
        )
        self.assertIn("not-allocatable", [p.code for p in problems])

    def test_a_session_in_another_semester_is_refused(self):
        self._seed()
        # A course the group genuinely studies, but in semester 2 only.
        other = Course.objects.create(code="MT171", name="Calculus")
        for programme in (self.ce, self.me):
            ProgrammeCourse.objects.create(
                programme=programme, course=other, semester=2
            )
        session = Session.objects.create(
            semester=self.other_semester,
            course_code="MT171",
            activity_type=ActivityType.TUTORIAL,
            day="MONDAY",
            start_time=time(8, 0),
            end_time=time(9, 0),
            venue=self.big,
        )
        index = AvailabilityIndex(self.other_semester)
        problems = validate_assignment(
            self.a1, other, session, self.semester, index
        )
        self.assertIn("wrong-semester", [p.code for p in problems])

    def test_an_overlapping_lecture_blocks_the_candidate(self):
        self._seed()
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 10, self.big
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        tutorial = self._tutorial("MONDAY", 9, 10, self.big)
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, self.maths, tutorial, self.semester, index
        )
        self.assertIn("clash", [p.code for p in problems])

    def test_adjacent_sessions_do_not_clash(self):
        self._seed()
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 9, self.big
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        tutorial = self._tutorial("MONDAY", 9, 10, self.big)
        index = AvailabilityIndex(self.semester)
        self.assertEqual(
            validate_assignment(
                self.a1, self.maths, tutorial, self.semester, index
            ),
            [],
        )

    def test_a_workshop_period_resolves_into_standard_hours(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            group_code="A1",
            day="MONDAY",
            time_period="MORNING",
            venue="TW101",
        )
        index = AvailabilityIndex(self.semester)
        clashes = index.conflicts_excluding(
            self.a1.pk, "MONDAY", time(10, 0), time(11, 0), None
        )
        self.assertEqual(len(clashes), 1)
        self.assertEqual(clashes[0].start, time(9, 0))
        self.assertEqual(clashes[0].end, time(13, 0))

    def test_a_workshop_outside_its_standard_hours_does_not_clash(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            group_code="A1",
            day="MONDAY",
            time_period="MORNING",
            venue="TW101",
        )
        index = AvailabilityIndex(self.semester)
        self.assertEqual(
            index.conflicts_excluding(
                self.a1.pk, "MONDAY", time(14, 0), time(15, 0), None
            ),
            [],
        )

    def test_a_workshop_with_no_resolvable_time_blocks_the_whole_day(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            group_code="A1",
            day="MONDAY",
            venue="TW101",
        )
        index = AvailabilityIndex(self.semester)
        self.assertEqual(len(index.unverifiable), 1)
        clashes = index.conflicts_excluding(
            self.a1.pk, "MONDAY", time(8, 0), time(9, 0), None
        )
        self.assertEqual([c.kind for c in clashes], ["unverifiable"])

    def test_a_clash_message_never_leaks_an_internal_session_marker(self):
        # The BusyBlock detail carries a "#<pk>:" prefix so the index can
        # recognise a session's own blocks. It must not reach the review panel.
        self._seed()
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 10, self.hall
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        tutorial = self._tutorial("MONDAY", 9, 10, self.hall)
        index = AvailabilityIndex(self.semester)
        problems = check_availability(self.a1, tutorial, index)
        self.assertEqual([p.code for p in problems], ["clash"])
        self.assertNotIn(f"#{lecture.pk}:", problems[0].message)
        self.assertNotIn("#", problems[0].message)
        self.assertIn("MT161", problems[0].message)

    def test_an_unverifiable_workshop_message_says_what_to_do(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            group_code="A1",
            day="MONDAY",
            venue="TW101",
        )
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        index = AvailabilityIndex(self.semester)
        problems = check_availability(self.a1, tutorial, index)
        self.assertEqual([p.code for p in problems], ["unverifiable-workshop"])
        self.assertIn("cannot be verified", problems[0].message)

    def test_a_technical_drawing_blocks_the_candidate(self):
        self._seed()
        TechnicalDrawingAllocation.objects.create(
            semester=self.semester,
            course_code="TG201",
            group_code="A1",
            day="MONDAY",
            start_time=time(8, 0),
            end_time=time(11, 0),
            venue="TW101",
        )
        index = AvailabilityIndex(self.semester)
        self.assertEqual(
            len(
                index.conflicts_excluding(
                    self.a1.pk, "MONDAY", time(10, 0), time(10, 30), None
                )
            ),
            1,
        )

    def test_a_technical_drawing_of_the_same_course_is_not_a_clash(self):
        """The reported case: ME101 technical drawing 10:00-13:00 vs the ME101
        tutorial 10:00-12:55, both in S112.

        ``fold_same_sessions`` already draws those two as ONE block on every
        export -- the allocation stands in for the session. Calling the group
        "busy" therefore refuses a session that is not a second commitment at
        all, and the group could never be placed.
        """
        self._seed()
        TechnicalDrawingAllocation.objects.create(
            semester=self.semester,
            course_code="ME101",
            group_code="A1",
            day="TUESDAY",
            start_time=time(10, 0),
            end_time=time(13, 0),
            venue="S112",
        )
        tutorial = self._session(
            "ME101", ActivityType.TUTORIAL, "TUESDAY", 10, (12, 55), self.hall
        )
        index = AvailabilityIndex(self.semester)
        self.assertEqual(check_availability(self.a1, tutorial, index), [])

    def test_the_same_course_is_matched_through_normalised_codes(self):
        self._seed()
        TechnicalDrawingAllocation.objects.create(
            semester=self.semester,
            course_code="  me101 ",  # as an import may well leave it
            group_code="A1",
            day="TUESDAY",
            start_time=time(10, 0),
            end_time=time(13, 0),
            venue="S112",
        )
        tutorial = self._session(
            "ME101", ActivityType.TUTORIAL, "TUESDAY", 10, (12, 55), self.hall
        )
        self.assertEqual(
            check_availability(self.a1, tutorial, AvailabilityIndex(self.semester)),
            [],
        )

    def test_a_technical_drawing_of_another_course_still_blocks(self):
        """The exemption is per course, not blanket: a TD201 drawing in the same
        slot really is a different class from the ME101 tutorial."""
        self._seed()
        TechnicalDrawingAllocation.objects.create(
            semester=self.semester,
            course_code="TG201",
            group_code="A1",
            day="TUESDAY",
            start_time=time(10, 0),
            end_time=time(13, 0),
            venue="TW101",
        )
        tutorial = self._session(
            "ME101", ActivityType.TUTORIAL, "TUESDAY", 10, (12, 55), self.hall
        )
        problems = check_availability(
            self.a1, tutorial, AvailabilityIndex(self.semester)
        )
        self.assertEqual([p.code for p in problems], ["clash"])
        self.assertIn("Technical drawing TG201", problems[0].message)

    def test_a_workshop_still_blocks_even_a_matching_course_name(self):
        """A workshop is its own class. The export folds a *placeholder* session
        into a workshop slot, but a real taught session is drawn alongside it --
        so a group in a workshop is genuinely busy for a tutorial."""
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            workshop="Carpentry",
            group_code="A1",
            day="TUESDAY",
            start_time=time(9, 0),
            end_time=time(13, 0),
            venue="TW101",
        )
        tutorial = self._session(
            "ME101", ActivityType.TUTORIAL, "TUESDAY", 10, (12, 55), self.hall
        )
        problems = check_availability(
            self.a1, tutorial, AvailabilityIndex(self.semester)
        )
        self.assertEqual([p.code for p in problems], ["clash"])
        self.assertIn("Workshop Carpentry", problems[0].message)

    def test_a_same_course_drawing_on_another_day_is_irrelevant(self):
        self._seed()
        TechnicalDrawingAllocation.objects.create(
            semester=self.semester,
            course_code="ME101",
            group_code="A1",
            day="WEDNESDAY",
            start_time=time(10, 0),
            end_time=time(13, 0),
            venue="S112",
        )
        tutorial = self._session(
            "ME101", ActivityType.TUTORIAL, "TUESDAY", 10, (12, 55), self.hall
        )
        self.assertEqual(
            check_availability(self.a1, tutorial, AvailabilityIndex(self.semester)),
            [],
        )

    def test_a_group_can_actually_be_placed_over_its_own_drawing(self):
        """End to end: the exemption has to survive the whole validator, or the
        fix is cosmetic."""
        self._seed()
        drawing = Course.objects.create(code="ME101", name="Technical Drawing 1")
        for programme in (self.ce, self.me):
            ProgrammeCourse.objects.create(
                programme=programme, course=drawing, semester=1
            )
        drawing.set_requirements({"TUTORIAL": 1})
        TechnicalDrawingAllocation.objects.create(
            semester=self.semester,
            course_code="ME101",
            group_code="A1",
            day="TUESDAY",
            start_time=time(10, 0),
            end_time=time(13, 0),
            venue="S112",
        )
        tutorial = self._session(
            "ME101", ActivityType.TUTORIAL, "TUESDAY", 10, (12, 55), self.hall
        )
        problems = validate_assignment(
            self.a1,
            drawing,
            tutorial,
            self.semester,
            AvailabilityIndex(self.semester),
        )
        self.assertEqual(problems, [])
        linked, problems = manual_assign(self.a1, tutorial)
        self.assertTrue(linked, [p.message for p in problems])

    def test_a_workshop_only_blocks_the_groups_carrying_its_code(self):
        self._seed()
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            group_code="A1",
            day="MONDAY",
            time_period="MORNING",
            venue="TW101",
        )
        index = AvailabilityIndex(self.semester)
        self.assertTrue(
            index.conflicts_excluding(
                self.a1.pk, "MONDAY", time(10, 0), time(10, 30), None
            )
        )
        self.assertFalse(
            index.conflicts_excluding(
                self.a2.pk, "MONDAY", time(10, 0), time(10, 30), None
            )
        )

    def test_a_group_already_in_the_session_does_not_clash_with_itself(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.big)
        SessionGroup.objects.create(session=tutorial, group=self.a1)
        index = AvailabilityIndex(self.semester)
        self.assertEqual(
            validate_assignment(
                self.a1, self.maths, tutorial, self.semester, index
            ),
            [],
        )

    def test_capacity_uses_thirty_students_per_group(self):
        session = Session(
            activity_type=ActivityType.TUTORIAL,
            day="MONDAY",
            start_time=time(8, 0),
            end_time=time(9, 0),
            venue=Venue(name="V", capacity=90),
        )
        self.assertEqual(required_capacity(3), 90)
        status, _ = capacity_status(session, 3)
        self.assertEqual(status, "ok")
        status, message = capacity_status(session, 4)
        self.assertEqual(status, "over-capacity")
        self.assertIn("120", message)

    def test_a_missing_venue_never_passes(self):
        self._seed()
        session = self._tutorial("MONDAY", 8, 9, venue=None)
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, self.maths, session, self.semester, index
        )
        self.assertIn("capacity-no-venue", [p.code for p in problems])

    def test_a_zero_capacity_venue_never_passes(self):
        self._seed()
        session = self._tutorial("MONDAY", 8, 9, self.no_capacity)
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, self.maths, session, self.semester, index
        )
        codes = [p.code for p in problems]
        self.assertIn("capacity-unknown-capacity", codes)
        self.assertIn("Set the capacity", " ".join(p.message for p in problems))

    def test_a_seminar_gets_no_special_minimum_capacity(self):
        self._seed(requirements={"SEMINAR": 1})
        # One group of 30 fits exactly in a 30-seat room.
        session = self._session(
            "MT161", ActivityType.SEMINAR, "MONDAY", 8, 9, self.small
        )
        index = AvailabilityIndex(self.semester)
        self.assertEqual(
            validate_assignment(
                self.a1, self.maths, session, self.semester, index
            ),
            [],
        )

    def test_a_backwards_session_time_is_reported_not_accepted(self):
        self._seed()
        session = self._tutorial("MONDAY", 10, 9, self.big)
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, self.maths, session, self.semester, index
        )
        self.assertIn("backwards-time", [p.code for p in problems])

    def test_an_unknown_activity_type_is_reported(self):
        self._seed()
        session = self._session("MT161", "UNSPECIFIED", "MONDAY", 8, 9, self.big)
        index = AvailabilityIndex(self.semester)
        problems = validate_assignment(
            self.a1, self.maths, session, self.semester, index
        )
        self.assertIn("unknown-activity", [p.code for p in problems])


class AllocationEngineTests(AllocationTestCase):
    def test_a_group_is_assigned_to_an_eligible_session(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.added, 4)
        self.assertEqual(plan.unresolved_count, 0)
        for assignment in plan.assignments:
            self.assertEqual(assignment.session.pk, tutorial.pk)

    def test_a_configured_activity_with_no_session_is_unresolved(self):
        self._seed()
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.added, 0)
        self.assertEqual(plan.unresolved_count, 4)
        for item in plan.unresolved:
            self.assertIn("No tutorial session exists", item.reasons[0])

    def test_nothing_is_written_by_planning(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        self.assertEqual(SessionGroup.objects.count(), 0)
        plan_allocation(self.semester, "ALL")
        self.assertEqual(SessionGroup.objects.count(), 0)

    def test_capacity_limits_how_many_groups_share_a_session(self):
        self._seed()
        # 30 seats: exactly one group fits.
        self._tutorial("MONDAY", 8, 9, self.small)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.added, 1)
        self.assertEqual(plan.unresolved_count, 3)
        self.assertIn("seats 30", " ".join(plan.unresolved[0].reasons))

    def test_a_session_with_no_venue_leaves_the_requirement_unresolved(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, venue=None)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.added, 0)
        self.assertEqual(plan.unresolved_count, 4)
        self.assertIn("No venue is set", " ".join(plan.unresolved[0].reasons))

    def test_a_zero_capacity_venue_leaves_the_requirement_unresolved(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.no_capacity)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.added, 0)
        self.assertIn(
            "Set the capacity of GAP", " ".join(plan.unresolved[0].reasons)
        )

    def test_groups_are_spread_across_sessions_to_fit_capacity(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.small)  # 1 group
        self._tutorial("TUESDAY", 8, 9, self.big)  # 3 groups
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.unresolved_count, 0)
        by_session = {}
        for assignment in plan.assignments:
            by_session.setdefault(assignment.session.day, []).append(
                assignment.requirement.group.code
            )
        self.assertEqual(len(by_session["MONDAY"]), 1)
        self.assertEqual(len(by_session["TUESDAY"]), 3)

    def test_a_timetable_clash_moves_the_group_to_another_session(self):
        self._seed()
        monday = self._tutorial("MONDAY", 8, 9, self.big)
        self._tutorial("TUESDAY", 8, 9, self.big)
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 10, self.big
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.unresolved_count, 0)
        placed = {
            a.requirement.group.code: a.session.day for a in plan.assignments
        }
        self.assertEqual(placed["A1"], "TUESDAY")
        self.assertEqual(placed["A2"], "MONDAY")

    def test_existing_valid_links_are_retained(self):
        self._seed()
        monday = self._tutorial("MONDAY", 8, 9, self.big)
        self._tutorial("TUESDAY", 8, 9, self.big)
        SessionGroup.objects.create(session=monday, group=self.a1)
        plan = plan_allocation(self.semester, "ALL")
        retained = [a for a in plan.assignments if a.status == "retained"]
        self.assertEqual(len(retained), 1)
        self.assertEqual(retained[0].requirement.group.code, "A1")
        self.assertEqual(retained[0].session.pk, monday.pk)

    def test_the_engine_keeps_earlier_choices_when_they_all_fit(self):
        self._seed(requirements={"SEMINAR": 1, "TUTORIAL": 1})
        self._session("MT161", ActivityType.SEMINAR, "MONDAY", 8, 9, self.hall)
        self._tutorial("MONDAY", 10, 11, self.hall)
        self._tutorial("TUESDAY", 10, 11, self.hall)
        seminar = Session.objects.get(
            course_code="MT161", activity_type=ActivityType.SEMINAR
        )
        tutorial = Session.objects.get(
            course_code="MT161", activity_type=ActivityType.TUTORIAL,
            day="TUESDAY",
        )
        SessionGroup.objects.create(session=seminar, group=self.a1)
        SessionGroup.objects.create(session=tutorial, group=self.a1)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.unresolved_count, 0)
        self.assertEqual(plan.moved, 0)
        self.assertEqual(plan.retained, 2)
        self.assertEqual(plan.added, 6)

    def test_the_engine_backtracks_an_earlier_stage_for_a_later_one(self):
        """A greedy first dive strands the practical; only backtracking saves it.

        Four groups need one seminar and one practical. The two Monday rooms
        (08:00-10:00 seminar, 10:00-12:00 practical) are the ones the
        most-constrained-first order reaches for first, but taking both for
        one group blocks that group's other requirement. Backtracking has to
        reconsider the seminar choice to place all eight.
        """
        self._seed(requirements={"SEMINAR": 1, "PRACTICAL": 1})
        monday = self._session(
            "MT161", ActivityType.SEMINAR, "MONDAY", 8, 10, self.hall
        )
        tuesday = self._session(
            "MT161", ActivityType.SEMINAR, "TUESDAY", 8, 10, self.hall
        )
        monday_p = self._session(
            "MT161", ActivityType.PRACTICAL, "MONDAY", 10, 12, self.hall
        )
        self._session("MT161", ActivityType.PRACTICAL, "TUESDAY", 10, 12, self.hall)
        # Force the first dive to take the Monday pair for A1, which is
        # impossible: the practical then clashes with the seminar.
        SessionGroup.objects.create(session=monday, group=self.a1)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.unresolved_count, 0)
        self.assertEqual(plan.added + plan.retained, 8)
        self.assertEqual(plan.retained, 1)  # A1 keeps its Monday seminar
        placed = {
            (a.requirement.group.code, a.requirement.activity_type): a.session.pk
            for a in plan.assignments
        }
        # A1 keeps its Monday seminar and takes the Tuesday practical.
        self.assertEqual(placed[("A1", "SEMINAR")], monday.pk)
        self.assertNotEqual(placed[("A1", "PRACTICAL")], monday_p.pk)
        # Every other group is placed too.
        self.assertEqual(len(placed), 8)
        self.assertIn(tuesday.pk, set(placed.values()))

    def test_a_later_stage_that_cannot_be_met_reports_the_specific_reason(self):
        self._seed(requirements={"SEMINAR": 1, "TUTORIAL": 1})
        self._session("MT161", ActivityType.SEMINAR, "MONDAY", 8, 9, self.hall)
        # One tutorial room for four groups, and it holds one.
        self._tutorial("MONDAY", 10, 11, self.small)
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.unresolved_count, 3)
        self.assertIn("seats 30", " ".join(plan.unresolved[0].reasons))

    def test_an_unverifiable_workshop_day_is_left_alone_and_reported(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.big)
        self._tutorial("TUESDAY", 8, 9, self.big)
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            group_code="A1",
            day="MONDAY",
            venue="TW101",
        )
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.unresolved_count, 0)
        self.assertTrue(any("cannot be verified" in w for w in plan.warnings))
        placed = {a.requirement.group.code: a.session.day for a in plan.assignments}
        self.assertEqual(placed["A1"], "TUESDAY")

    def test_reaching_the_search_limit_is_reported_not_called_impossible(self):
        self._seed()
        for day in ("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"):
            for hour in (8, 10, 12, 14, 16):
                self._tutorial(day, hour, hour + 1, self.hall)
        plan = plan_allocation(self.semester, "ALL", node_limit=2)
        self.assertTrue(plan.search_limit_hit)
        self.assertFalse(plan.is_complete())
        self.assertTrue(
            any("not proof that no valid allocation exists" in w for w in plan.warnings)
        )
        # Nothing was written even though the search was abandoned.
        self.assertEqual(SessionGroup.objects.count(), 0)

    def test_a_full_search_within_the_budget_is_reported_as_complete(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        plan = plan_allocation(self.semester, "ALL")
        self.assertFalse(plan.search_limit_hit)
        self.assertTrue(plan.is_complete())

    def test_the_plan_is_deterministic(self):
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        for day in ("MONDAY", "TUESDAY", "WEDNESDAY"):
            self._tutorial(day, 8, 9, self.hall)
            self._session("MT161", ActivityType.PRACTICAL, day, 10, 12, self.hall)
        first = [
            (a.requirement.group.code, a.requirement.activity_type, a.session.pk)
            for a in plan_allocation(self.semester, "ALL").assignments
        ]
        second = [
            (a.requirement.group.code, a.requirement.activity_type, a.session.pk)
            for a in plan_allocation(self.semester, "ALL").assignments
        ]
        self.assertEqual(first, second)

    def test_an_unverifiable_workshop_is_never_silently_accepted(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        WorkshopAllocation.objects.create(
            semester=self.semester,
            course_code="WT107",
            group_code="A1",
            day="MONDAY",
            venue="TW101",
        )
        plan = plan_allocation(self.semester, "ALL")
        self.assertEqual(plan.unresolved_count, 1)
        self.assertIn("cannot be verified", plan.unresolved[0].reasons[0])

    def test_two_groups_from_different_programmes_may_share_a_session(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        plan = plan_allocation(self.semester, "ALL")
        programmes = {a.requirement.group.programme.code for a in plan.assignments}
        self.assertEqual(programmes, {"CE", "ME"})
        self.assertEqual(plan.session_rollup()[0]["programme_count"], 2)

    def test_the_scope_restricts_the_run(self):
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        self._tutorial("MONDAY", 8, 9, self.hall)
        self._session("MT161", ActivityType.PRACTICAL, "MONDAY", 10, 12, self.hall)
        plan = plan_allocation(self.semester, "TUTORIAL")
        self.assertEqual(plan.added, 4)
        self.assertEqual(
            {a.requirement.activity_type for a in plan.assignments},
            {ActivityType.TUTORIAL},
        )


class ManualAssignmentTests(AllocationTestCase):
    def test_a_valid_manual_assignment_is_accepted(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.big)
        linked, problems = manual_assign(self.a1, tutorial)
        self.assertTrue(linked, problems)
        self.assertTrue(
            SessionGroup.objects.filter(session=tutorial, group=self.a1).exists()
        )

    def test_a_manual_assignment_uses_the_engine_rules(self):
        self._seed()
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 10, self.big
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        tutorial = self._tutorial("MONDAY", 9, 10, self.big)
        linked, problems = manual_assign(self.a1, tutorial)
        self.assertFalse(linked)
        self.assertIn("clash", [p.code for p in problems])

    def test_a_manual_assignment_respects_capacity(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.small)
        self.assertTrue(manual_assign(self.a1, tutorial)[0])
        linked, problems = manual_assign(self.a2, tutorial)
        self.assertFalse(linked)
        self.assertIn("capacity-over-capacity", [p.code for p in problems])

    def test_a_manual_assignment_refuses_a_venue_with_no_capacity(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.no_capacity)
        linked, problems = manual_assign(self.a1, tutorial)
        self.assertFalse(linked)
        self.assertIn("capacity-unknown-capacity", [p.code for p in problems])

    def test_a_session_whose_course_is_unknown_is_reported(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.big, course="ZZ999")
        course, problems = validate_manual_assignment(self.a1, tutorial)
        self.assertIsNone(course)
        self.assertIn("unknown-course", [p.code for p in problems])

    def test_the_course_is_matched_on_the_normalised_code(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.big, course="  mt161 ")
        course, problems = validate_manual_assignment(self.a1, tutorial)
        self.assertEqual(course, self.maths)
        self.assertEqual(problems, [])

    def test_unassign_removes_the_link(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.big)
        manual_assign(self.a1, tutorial)
        self.assertEqual(manual_unassign(self.a1, tutorial), 1)
        self.assertFalse(
            SessionGroup.objects.filter(session=tutorial, group=self.a1).exists()
        )


class AllocationRunTests(AllocationTestCase):
    def _plan_and_run(self, scope="ALL"):
        plan = plan_allocation(self.semester, scope)
        return plan, save_plan(plan, algorithm="smart")

    def _a_group_pinned_to_monday(self):
        """A group that must be moved off Monday to its Tuesday session.

        The Monday tutorial is the one the group is already linked to, and a
        lecture overlapping it means the plan has to move it — which is what
        produces the REMOVE half of a revert.
        """
        monday = self._tutorial("MONDAY", 8, 9, self.hall)
        self._tutorial("TUESDAY", 8, 9, self.hall)
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 10, self.hall
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        SessionGroup.objects.create(session=monday, group=self.a1)
        return monday

    def test_applying_a_run_creates_the_links(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        plan, run = self._plan_and_run()
        result = apply_run(run)
        self.assertTrue(result["ok"])
        self.assertEqual(result["added"], 4)
        self.assertEqual(SessionGroup.objects.count(), 4)
        self.assertEqual(run.status, AllocationStatus.APPLIED)
        self.assertIsNotNone(run.applied_at)

    def test_applying_twice_is_refused(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        _, run = self._plan_and_run()
        apply_run(run)
        second = apply_run(run)
        self.assertFalse(second["ok"])
        self.assertEqual(SessionGroup.objects.count(), 4)

    def test_valid_assignments_apply_while_items_stay_unresolved(self):
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        self._tutorial("MONDAY", 8, 9, self.hall)
        plan, run = self._plan_and_run()
        self.assertEqual(plan.added, 4)
        self.assertEqual(plan.unresolved_count, 4)
        apply_run(run)
        self.assertEqual(SessionGroup.objects.count(), 4)
        self.assertEqual(run.unresolved, 4)

    def test_reverting_restores_the_original_state(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        _, run = self._plan_and_run()
        apply_run(run)
        result = revert_run(run)
        self.assertTrue(result["ok"])
        self.assertEqual(result["removed"], 4)
        self.assertEqual(SessionGroup.objects.count(), 0)
        self.assertEqual(run.status, AllocationStatus.REVERTED)

    def test_reverting_a_move_puts_the_group_back(self):
        self._seed()
        self._a_group_pinned_to_monday()
        before = set(
            SessionGroup.objects.values_list("session_id", "group_id")
        )
        plan, run = self._plan_and_run()
        self.assertEqual(plan.moved, 1)
        apply_run(run)
        self.assertNotEqual(
            set(SessionGroup.objects.values_list("session_id", "group_id")), before
        )
        revert_run(run)
        self.assertEqual(
            set(SessionGroup.objects.values_list("session_id", "group_id")), before
        )

    def test_revert_is_refused_when_an_added_link_was_removed_by_hand(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        _, run = self._plan_and_run()
        apply_run(run)
        SessionGroup.objects.all().delete()
        result = revert_run(run)
        self.assertFalse(result["ok"])
        self.assertTrue(result["conflicts"])
        self.assertIn("edited since this run was applied", result["message"])
        self.assertEqual(AllocationRun.objects.get(pk=run.pk).status, "APPLIED")

    def test_revert_is_refused_when_a_moved_group_was_linked_back(self):
        self._seed()
        monday = self._a_group_pinned_to_monday()
        _, run = self._plan_and_run()
        apply_run(run)
        # A1 was moved off Monday; putting it back is a later manual edit.
        SessionGroup.objects.get_or_create(session=monday, group=self.a1)
        result = revert_run(run)
        self.assertFalse(result["ok"])
        self.assertTrue(
            any("linked again" in c for c in result["conflicts"]), result
        )

    def test_only_an_applied_run_can_be_reverted(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        _, run = self._plan_and_run()
        result = revert_run(run)
        self.assertFalse(result["ok"])
        self.assertIn("Only an applied run", result["message"])

    def test_the_run_keeps_the_link_level_history(self):
        self._seed()
        monday = self._a_group_pinned_to_monday()
        _, run = self._plan_and_run()
        actions = set(run.changes.values_list("action", flat=True))
        self.assertEqual(actions, {"ADD", "REMOVE"})
        # The run that moved A1 off Monday records both halves of that move.
        self.assertTrue(
            run.changes.filter(
                action=AllocationChange.Action.REMOVE, group=self.a1, session=monday
            ).exists()
        )
        self.assertTrue(
            run.changes.filter(
                action=AllocationChange.Action.ADD, group=self.a1
            ).exists()
        )

    def test_a_retained_link_is_never_recorded_as_a_change(self):
        self._seed()
        monday = self._tutorial("MONDAY", 8, 9, self.hall)
        SessionGroup.objects.create(session=monday, group=self.a1)
        _, run = self._plan_and_run()
        self.assertFalse(
            run.changes.filter(group=self.a1, session=monday).exists()
        )
        # Reverting therefore leaves that pre-existing link alone.
        apply_run(run)
        revert_run(run)
        self.assertTrue(
            SessionGroup.objects.filter(session=monday, group=self.a1).exists()
        )

    def test_the_snapshot_records_every_proposal(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        _, run = self._plan_and_run()
        data = run.plan()
        self.assertEqual(data["added"], 4)
        self.assertEqual(len(data["assignments"]), 4)
        entry = data["assignments"][0]
        for key in (
            "group", "programme", "course", "activity", "day", "time",
            "venue", "group_count", "capacity_status", "status",
        ):
            self.assertIn(key, entry)
        self.assertEqual(data["requirement_total"], 4)
        self.assertTrue(data["complete"])

    def test_the_snapshot_records_unresolved_reasons(self):
        self._seed()
        _, run = self._plan_and_run()
        data = run.plan()
        self.assertEqual(len(data["unresolved_items"]), 4)
        self.assertTrue(data["unresolved_items"][0]["reasons"])
        self.assertFalse(data["complete"])

    def test_unresolved_reasons_are_grouped_across_every_candidate(self):
        # Five zero-capacity sessions and one clashable one: the report must say
        # so, not show only whichever session was tried last.
        self._seed()
        for hour in (8, 9, 10, 11, 12):
            self._tutorial("MONDAY", hour, hour + 1, self.no_capacity)
        clashable = self._tutorial("TUESDAY", 8, 9, self.hall)
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "TUESDAY", 8, 10, self.hall
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        plan = plan_allocation(self.semester, "ALL")
        item = next(
            u for u in plan.unresolved if u.requirement.group.code == "A1"
        )
        self.assertEqual(item.sessions_considered, 6)
        lines = " ".join(item.summary_lines())
        # A problem that blocked several sessions says so; a one-off does not
        # need a prefix. The "N sessions checked" line gives the total.
        self.assertIn("5 sessions", lines)
        self.assertIn("no capacity recorded", lines)
        self.assertIn("already busy", lines)
        self.assertNotIn("1 session:", lines)
        # One message per distinct problem, however many sessions it blocked.
        self.assertEqual(len(item.reasons), 2)
        self.assertEqual(item.grouped["capacity-unknown-capacity"]["count"], 5)
        self.assertEqual(item.grouped["clash"]["count"], 1)

    def test_applying_is_atomic(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.big)
        _, run = self._plan_and_run()
        original = SessionGroup.objects.count()
        with mock.patch(
            "core.group_allocation.SessionGroup.objects.get_or_create",
            side_effect=RuntimeError("boom"),
        ):
            with self.assertRaises(RuntimeError):
                apply_run(run)
        self.assertEqual(SessionGroup.objects.count(), original)
        self.assertEqual(AllocationRun.objects.get(pk=run.pk).status, "PREVIEWED")


class AllocationPageTests(AllocationTestCase):
    def test_the_page_renders_with_a_semester_and_activity_scope(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        resp = self.client.get("/allocation/")
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "Group Allocation")
        self.assertContains(resp, "Simulate Allocation")
        self.assertContains(resp, "Seminar")
        self.assertContains(resp, "Tutorial")
        self.assertContains(resp, "Practical")

    def test_the_sidebar_links_to_the_page(self):
        self._seed()
        resp = self.client.get("/allocation/")
        self.assertContains(resp, 'href="/allocation/"')

    def test_the_preview_reports_the_plan_without_writing(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        resp = self.client.post(
            "/allocation/preview/",
            {"semester": self.semester.pk, "scope": "ALL"},
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "Allocation plan")
        self.assertContains(resp, "Apply Allocation")
        self.assertEqual(SessionGroup.objects.count(), 0)
        self.assertEqual(AllocationRun.objects.count(), 1)

    def test_the_preview_shows_unresolved_reasons(self):
        self._seed()
        resp = self.client.post(
            "/allocation/preview/",
            {"semester": self.semester.pk, "scope": "ALL"},
        )
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "Unresolved")
        self.assertContains(resp, "No tutorial session exists")

    def test_the_preview_flags_courses_with_no_requirement(self):
        self._seed()
        other = Course.objects.create(code="ZZ999", name="Unconfigured")
        ProgrammeCourse.objects.create(
            programme=self.ce, course=other, semester=1
        )
        resp = self.client.post(
            "/allocation/preview/",
            {"semester": self.semester.pk, "scope": "ALL"},
        )
        self.assertContains(resp, "requirement not configured")
        self.assertContains(resp, "ZZ999")

    def test_the_plan_is_readable_on_a_phone_not_a_nine_column_table(self):
        """The plan overflowed the viewport on a phone.

        Nine columns cannot be read at 375px, and the page then forced a
        horizontal scroll for the whole document. The rows are now presented
        twice -- a labelled card list below `lg`, the table from `lg` up -- and
        each presentation is hidden at the other one's breakpoint, so exactly
        one is ever on screen and no field is dropped to achieve it.

        The assertion is scoped to one view for that reason: "this text appears
        once" is no longer true of the served HTML, exactly as with the
        timetable's grid/agenda pair.
        """
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        resp = self.client.post(
            "/allocation/preview/",
            {"semester": self.semester.pk, "scope": "ALL"},
        )
        html = resp.content.decode()

        cards = html[html.index('class="tt-plan-mobile'):]
        cards = cards[: cards.index("hidden lg:block")]
        table = html[html.index("hidden lg:block overflow-x-auto"):]

        # A card carries every field the table row does, labelled, so nothing is
        # lost by not showing the table.
        for field in ("Group", "Course", "Day", "Time", "Venue",
                      "Groups in session", "Capacity"):
            self.assertIn(field, cards)
        self.assertIn("MT161", cards)
        # And the card list is the table's rows, not a different plan.
        self.assertIn("MT161", table)

        # Each presentation is hidden at the other's breakpoint, so the two
        # never both render.
        self.assertIn("lg:hidden", cards[: cards.index(">")])
        self.assertTrue(table.startswith("hidden lg:block overflow-x-auto"))

    def test_the_page_wrapper_cannot_be_pushed_wider_than_the_screen(self):
        """A grid/flex item defaults to `min-width: auto`.

        That is what let the plan's wide table widen the whole page instead of
        scrolling inside `overflow-x-auto`, so the container itself has to opt
        out. Asserted because the fix is a class that looks incidental.
        """
        self._seed()
        resp = self.client.get("/allocation/")
        html = resp.content.decode()
        self.assertIn('class="max-w-7xl mx-auto space-y-4 min-w-0 overflow-x-hidden"', html)

    def test_apply_then_revert_through_the_views(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        run = self.client.post(
            "/allocation/preview/",
            {"semester": self.semester.pk, "scope": "ALL"},
        ).context["run"]
        self.client.post(
            "/allocation/apply/", {"run": run.pk}, HTTP_HX_REQUEST="true"
        )
        self.assertEqual(SessionGroup.objects.count(), 4)
        self.client.post(
            "/allocation/revert/", {"run": run.pk}, HTTP_HX_REQUEST="true"
        )
        self.assertEqual(SessionGroup.objects.count(), 0)

    def test_a_manual_assignment_is_validated_and_feedback_returned(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": tutorial.pk},
            HTTP_HX_REQUEST="true",
        )
        self.assertContains(resp, "assigned to")
        self.assertTrue(
            SessionGroup.objects.filter(session=tutorial, group=self.a1).exists()
        )

    def test_a_refused_manual_assignment_explains_itself(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.small)
        self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": tutorial.pk},
            HTTP_HX_REQUEST="true",
        )
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a2.pk, "session": tutorial.pk},
            HTTP_HX_REQUEST="true",
        )
        self.assertContains(resp, "Assignment refused")
        self.assertContains(resp, "seats 30")
        self.assertFalse(
            SessionGroup.objects.filter(session=tutorial, group=self.a2).exists()
        )

    def test_a_group_can_be_unassigned_by_hand(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        manual_assign(self.a1, tutorial)
        resp = self.client.post(
            "/allocation/unassign/",
            {"group": self.a1.pk, "session": tutorial.pk},
            HTTP_HX_REQUEST="true",
        )
        self.assertContains(resp, "removed from")
        self.assertFalse(
            SessionGroup.objects.filter(session=tutorial, group=self.a1).exists()
        )

    def test_applying_logs_the_change(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        run = self.client.post(
            "/allocation/preview/",
            {"semester": self.semester.pk, "scope": "ALL"},
        ).context["run"]
        self.client.post("/allocation/apply/", {"run": run.pk})
        messages = list(ActivityLog.objects.values_list("message", flat=True))
        self.assertTrue(any("Applied group allocation run" in m for m in messages))
        self.assertTrue(
            ActivityLog.objects.filter(
                action=LogAction.ASSIGN, resource="Group Allocation"
            ).exists()
        )

    def test_reverting_logs_a_removal(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.big)
        run = self.client.post(
            "/allocation/preview/",
            {"semester": self.semester.pk, "scope": "ALL"},
        ).context["run"]
        self.client.post("/allocation/apply/", {"run": run.pk})
        self.client.post("/allocation/revert/", {"run": run.pk})
        self.assertTrue(
            ActivityLog.objects.filter(
                action=LogAction.REMOVE, resource="Group Allocation"
            ).exists()
        )

    def test_get_on_the_mutating_endpoints_redirects(self):
        self._seed()
        for url in (
            "/allocation/preview/",
            "/allocation/apply/",
            "/allocation/revert/",
            "/allocation/assign/",
            "/allocation/unassign/",
        ):
            resp = self.client.get(url)
            self.assertEqual(resp.status_code, 302, url)
            self.assertEqual(resp.url, "/allocation/")

    def test_the_page_survives_an_empty_database(self):
        resp = self.client.get("/allocation/")
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, "Create a semester before allocating groups")


class GroupStatusTests(AllocationTestCase):
    """``group_statuses`` answers "where does this group stand?" without
    proposing anything. It must never claim a group is done when a requirement
    is unplaced, and it must never offer a session the engine would refuse."""

    def test_a_group_with_nothing_placed_is_wholly_outstanding(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        status = group_statuses(self.semester, "ALL", [self.a1])[self.a1.pk]
        self.assertEqual(status.total, 1)
        self.assertEqual(status.met, 0)
        self.assertEqual(status.outstanding, 1)
        self.assertEqual(status.state, "unassigned")
        self.assertFalse(status.is_complete)

    def test_a_fully_placed_group_is_complete_and_not_have_nothing_to_do(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        manual_assign(self.a1, tutorial)
        status = group_statuses(self.semester, "ALL", [self.a1])[self.a1.pk]
        self.assertEqual((status.met, status.total), (1, 1))
        self.assertEqual(status.percent, 100)
        self.assertEqual(status.state, "complete")
        self.assertTrue(status.is_complete)
        # "complete" must mean "finished", never "nothing was ever asked".
        self.assertFalse(status.has_nothing_to_do)

    def test_a_group_whose_courses_require_nothing_is_reported_as_such(self):
        self._seed()
        self.maths.set_requirements({})
        status = group_statuses(self.semester, "ALL", [self.a1])[self.a1.pk]
        self.assertEqual(status.total, 0)
        self.assertTrue(status.has_nothing_to_do)
        self.assertFalse(status.is_complete)  # nothing asked is not "done"
        self.assertEqual(status.state, "none")
        self.assertEqual(status.percent, 0)

    def test_a_partly_placed_group_counts_only_what_it_has(self):
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        manual_assign(self.a1, tutorial)
        status = group_statuses(self.semester, "ALL", [self.a1])[self.a1.pk]
        self.assertEqual((status.met, status.total), (1, 2))
        self.assertEqual(status.state, "partial")
        self.assertEqual(status.percent, 50)
        self.assertEqual(
            [(label, met, total) for label, met, total in status.activity_progress()],
            [("Tutorial", 1, 1), ("Practical", 0, 1)],
        )

    def test_an_unplaced_requirement_is_matched_to_its_session(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        entry = group_statuses(self.semester, "ALL", [self.a1])[self.a1.pk].entries[0]
        self.assertIsNone(entry.assigned)
        placed = [o for o in entry.options if o.already_assigned]
        self.assertEqual(placed, [])
        self.assertIn(tutorial, [o.session for o in entry.options])

    def test_the_session_the_group_is_in_is_marked_assigned(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        manual_assign(self.a1, tutorial)
        entry = group_statuses(self.semester, "ALL", [self.a1])[self.a1.pk].entries[0]
        self.assertTrue(entry.is_met)
        self.assertEqual(entry.assigned, tutorial)
        assigned = [o for o in entry.options if o.already_assigned]
        self.assertEqual([o.session for o in assigned], [tutorial])
        # The current placement is the status, not a proposal, so it is never
        # re-judged -- it cannot come back with a "not possible".
        self.assertTrue(assigned[0].ok)
        self.assertEqual(assigned[0].reasons, [])

    def test_a_clashing_session_is_listed_but_marked_not_free(self):
        self._seed()
        monday = self._tutorial("MONDAY", 8, 9, self.hall)
        # Overlaps the one above, so the group is busy at the time.
        clashing = self._tutorial("MONDAY", 8, (9, 30), self.hall)
        manual_assign(self.a1, monday)
        entry = group_statuses(
            self.semester, "ALL", [self.a1]
        )[self.a1.pk].entries[0]
        blocked = {o.session.pk: o for o in entry.options if not o.ok}
        self.assertEqual(list(blocked), [clashing.pk])
        self.assertFalse(blocked[clashing.pk].free)
        self.assertTrue(blocked[clashing.pk].reasons)
        # The session the group is already in is still a valid option; nothing
        # else is, so this requirement has nowhere left to move to.
        self.assertEqual(
            [o.session for o in entry.free_options if not o.already_assigned], []
        )

    def test_a_too_small_session_is_offered_as_free_but_not_ok(self):
        """"Free" and "valid" are different questions and must be reported apart.

        NB102 seats 30, so it takes exactly one group. The second group is free
        at that time -- it simply cannot fit. Reporting one blunt "no" would
        hide that it is a room-size problem, not a timetable clash.
        """
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.small)
        manual_assign(self.a1, tutorial)
        only = group_statuses(
            self.semester, "ALL", [self.a2]
        )[self.a2.pk].entries[0].options[0]
        self.assertTrue(only.free)  # nothing clashes
        self.assertFalse(only.ok)  # but the room is full
        self.assertIn("seats 30", " ".join(r.message for r in only.reasons))

    def test_an_option_counts_the_group_that_would_be_added(self):
        self._seed()
        tutorial = self._tutorial("MONDAY", 8, 9, self.big)  # seats 90
        manual_assign(self.a1, tutorial)
        option = group_statuses(
            self.semester, "ALL", [self.a2]
        )[self.a2.pk].entries[0].options[0]
        self.assertEqual(option.group_count, 2)
        self.assertTrue(option.ok)

    def test_the_options_are_ordered_with_the_current_placement_first(self):
        self._seed()
        later = self._tutorial("FRIDAY", 8, 9, self.hall)
        earlier = self._tutorial("MONDAY", 8, 9, self.hall)
        manual_assign(self.a1, later)
        options = group_statuses(
            self.semester, "ALL", [self.a1]
        )[self.a1.pk].entries[0].options
        self.assertEqual(options[0].session, later)
        self.assertTrue(options[0].already_assigned)
        self.assertEqual([o.session for o in options[1:]], [earlier])

    def test_unconfigured_courses_are_attached_to_the_groups_that_study_them(self):
        self._seed(requirements={})
        other = Course.objects.create(code="ZZ999", name="Unconfigured")
        ProgrammeCourse.objects.create(
            programme=self.ce, course=other, semester=1
        )
        statuses = group_statuses(self.semester, "ALL", [self.a1, self.d1])
        self.assertEqual(
            [e["code"] for e in statuses[self.a1.pk].unconfigured_courses],
            ["ZZ999"],
        )
        # D1 studies the other programme, so the same course is not its problem.
        self.assertEqual(statuses[self.d1.pk].unconfigured_courses, [])

    def test_the_scope_limits_the_activities_being_asked_about(self):
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        self._tutorial("MONDAY", 8, 9, self.hall)
        self._session("MT161", ActivityType.PRACTICAL, "TUESDAY", 8, 9, self.hall)
        status = group_statuses(self.semester, "TUTORIAL", [self.a1])[self.a1.pk]
        self.assertEqual(status.total, 1)
        self.assertEqual(status.entries[0].activity_label, "Tutorial")

    def test_a_group_with_no_requirement_still_gets_a_row(self):
        """The board must show "nothing asked of it", not quietly omit it."""
        self._seed(requirements={})
        self.assertIn(self.a1.pk, group_statuses(self.semester, "ALL", [self.a1]))


class GroupProgressBoardTests(AllocationTestCase):
    def setUp(self):
        super().setUp()
        self._seed(requirements={"TUTORIAL": 1, "PRACTICAL": 1})
        self.tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        self.practical = self._session(
            "MT161", ActivityType.PRACTICAL, "TUESDAY", 8, 9, self.hall
        )

    def _board(self):
        return self.client.get(
            f"/allocation/groups/?semester={self.semester.pk}&scope=ALL"
        )

    def test_the_board_lists_every_group_with_a_progress_bar(self):
        resp = self._board()
        self.assertEqual(resp.status_code, 200)
        rows = resp.context["rows"]
        self.assertEqual({r.group.pk for r in rows}, {g.pk for g in (self.a1, self.a2, self.d1, self.d2)})
        for row in rows:
            self.assertEqual((row.met, row.total), (0, 2))
            self.assertEqual(row.outstanding, 2)

    def test_a_completed_group_is_counted_and_sorts_ahead(self):
        manual_assign(self.a1, self.tutorial)
        manual_assign(self.a1, self.practical)
        rows = self._board().context["rows"]
        self.assertEqual(rows[0].group, self.a1)
        self.assertTrue(rows[0].is_complete)
        self.assertEqual(self._board().context["complete"], 1)
        self.assertEqual(self._board().context["outstanding"], 6)

    def test_the_summary_buckets_add_up_to_the_number_of_groups(self):
        manual_assign(self.a1, self.tutorial)  # 1/2 -> partial
        ctx = self._board().context
        self.assertEqual(
            ctx["complete"] + ctx["partial"] + ctx["unassigned"] + ctx["nothing"],
            len(ctx["rows"]),
        )
        self.assertEqual(ctx["partial"], 1)
        self.assertEqual(ctx["unassigned"], 3)

    def test_each_row_links_to_that_group_and_keeps_the_filters(self):
        row = self._board().context["rows"][0]
        self.assertEqual(row.detail_url, f"/allocation/group/{row.group.pk}/")
        resp = self._board()
        self.assertContains(
            resp, f"/allocation/group/{row.group.pk}/?semester={self.semester.pk}"
        )

    def test_the_board_survives_an_empty_database(self):
        resp = self.client.get("/allocation/groups/")
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.context["rows"], [])

    def test_the_widest_cell_is_constrained_so_a_phone_is_not_scrolled_sideways(self):
        """The board is one row of facts per group, so it stays a table at every width.

        That makes the "Still needed" list the thing to watch: left unbounded it
        is the widest thing in the row, and the table ends up far wider than a
        phone, so the page has to be dragged sideways to read it. Constraining
        that one cell (rather than only wrapping the container in an
        `overflow-x-auto`, which does nothing while a child is that wide) is
        what keeps the scroll to the table's own box.
        """
        html = self._board().content.decode()
        self.assertIn("max-w-[16rem] sm:max-w-none", html)
        self.assertIn("text-amber-800 break-safe", html)
        # min-w-0 on the fragment and the scroll box, for the same reason
        # min-width:auto would otherwise let the table widen the page.
        self.assertIn('id="allocation-groups-board"\n     class="space-y-4 min-w-0"', html)
        self.assertIn('class="overflow-x-auto min-w-0"', html)

    def test_a_row_opens_that_group_in_the_slide_panel_from_anywhere_on_it(self):
        """A tap anywhere on a row is enough -- the last column is not the way in.

        The board is the working view, so reaching the placement form should not
        mean travelling to the end of the row, or to another page.
        """
        row = self._board().context["rows"][0]
        self.assertEqual(
            row.slide_url,
            f"/allocation/group/{row.group.pk}/slide/"
            f"?semester={self.semester.pk}&scope=ALL",
        )
        resp = self._board()
        self.assertContains(
            resp, f'hx-get="{row.slide_url.replace("&", "&amp;")}"'
        )
        self.assertContains(resp, 'hx-target="#slide-content"')
        self.assertContains(resp, '@click="slideOpen = true"')

    def test_the_row_keeps_its_own_link_without_opening_the_panel_over_it(self):
        row = self._board().context["rows"][0]
        self.assertContains(self._board(), "@click.stop")

    def test_the_slide_serves_the_group_panel_the_page_would_show(self):
        row = self._board().context["rows"][0]
        resp = self.client.get(row.slide_url, HTTP_HX_REQUEST="true")
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.context["group"], row.group)
        self.assertContains(resp, 'id="allocation-requirements"')
        self.assertContains(resp, self.a1.code)
        # The programme the narrow-screen table drops is named in the panel.
        self.assertContains(resp, self.a1.programme.name)
        # And there is still a way through to the full page.
        self.assertContains(resp, f"/allocation/group/{row.group.pk}/")

    def test_the_programme_column_is_scoped_away_on_a_narrow_screen_not_dropped(self):
        resp = self._board()
        # Hidden below md, back from md up -- a phone loses the width, not the
        # column, and the panel still names the programme.
        self.assertContains(resp, "hidden md:table-cell")
        self.assertContains(resp, self.a1.programme.code)
        self.assertContains(resp, "hidden sm:table-cell")

    def test_the_board_is_served_on_its_own_so_a_placement_can_redraw_it(self):
        resp = self.client.get(
            f"/allocation/groups/?semester={self.semester.pk}&scope=ALL",
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertTemplateUsed(resp, "core/_allocation_groups_board.html")
        # The fragment carries the filters it is showing, so a refresh cannot
        # quietly redraw the board against a different semester.
        self.assertContains(resp, 'id="allocation-groups-board"')
        self.assertContains(
            resp,
            f'hx-get="/allocation/groups/?semester={self.semester.pk}&amp;scope=ALL"',
        )

    def test_a_placement_made_from_the_panel_asks_the_board_to_redraw(self):
        """Otherwise the row the panel was opened from keeps claiming the work."""
        row = self._board().context["rows"][0]
        entry = self.client.get(
            row.slide_url, HTTP_HX_REQUEST="true"
        ).context["status"].entries[0]
        resp = self.client.post(
            "/allocation/assign/",
            {
                "group": self.a1.pk,
                "session": entry.options[0].session.pk,
                "panel": "1",
                "semester": self.semester.pk,
                "scope": "ALL",
            },
            HTTP_HX_REQUEST="true",
        )
        self.assertIn("refresh-table", resp["HX-Trigger"])

    def test_the_sidebar_lights_up_group_progress_here_and_the_allocator_elsewhere(self):
        board = self._board()
        nav = board.context["nav"]
        self.assertEqual(nav, "allocation-progress")
        # The link is labelled "Group Progress", and it lives inside the
        # collapsible Allocation section rather than beside the allocator.
        self.assertContains(board, ">Group Progress<")
        self.assertContains(board, 'aria-label="Group Progress"')
        # The allocator link must not be the one that looks active.
        self.assertEqual(
            self.client.get("/allocation/").context["nav"], "allocation"
        )


class AllocationDefaultsToCurrentSemesterTests(AllocationTestCase):
    """A bare visit to an allocation page means the term the app is working in.

    These pages used to fall back to ``Semester.objects.all()[0]`` -- an
    unordered queryset, so which term the coordinator landed on was the
    database's row order rather than a decision. The dashboard already defaults
    to the current semester; the allocator and the progress board now do the
    same, so every page agrees about which term it is showing.

    ``_seed`` creates 2026/2027 semester 1 first and semester 2 second, so the
    *newest* term and the *first-created* term are different rows. Marking
    semester 2 current therefore distinguishes "honours the current semester"
    from the old "whichever row came out first", instead of passing by accident.
    """

    def setUp(self):
        super().setUp()
        # These are management pages, so the client has to be a staff member --
        # StaffAccessMiddleware 302s an anonymous one before the view runs, and
        # a test that never reaches the view verifies nothing.
        self.client.force_login(
            User.objects.create_user(
                "allocation-tester", password="pw", is_staff=True
            )
        )
        self._seed(requirements={"TUTORIAL": 1})
        self.tutorial = self._tutorial("MONDAY", 8, 9, self.hall)

    def test_the_progress_board_opens_on_the_current_semester(self):
        Semester.set_current(self.other_semester)
        self.assertEqual(
            self.client.get("/allocation/groups/").context["semester"],
            self.other_semester,
        )

    def test_the_allocator_opens_on_the_current_semester(self):
        Semester.set_current(self.other_semester)
        self.assertEqual(
            self.client.get("/allocation/").context["semester"],
            self.other_semester,
        )

    def test_a_group_page_opens_on_the_current_semester(self):
        Semester.set_current(self.other_semester)
        self.assertEqual(
            self.client.get(f"/allocation/group/{self.a1.pk}/").context["semester"],
            self.other_semester,
        )

    def test_the_slide_and_the_panel_agree_with_the_page(self):
        """A placement from a board row must not redraw against another term."""
        Semester.set_current(self.other_semester)
        for url in (
            f"/allocation/group/{self.a1.pk}/slide/",
            f"/allocation/group/{self.a1.pk}/panel/",
        ):
            with self.subTest(url=url):
                self.assertEqual(
                    self.client.get(url).context["semester"], self.other_semester
                )

    def test_an_explicit_picker_still_wins_over_the_current_semester(self):
        Semester.set_current(self.other_semester)
        self.assertEqual(
            self.client.get(
                f"/allocation/groups/?semester={self.semester.pk}"
            ).context["semester"],
            self.semester,
        )

    def test_with_no_current_semester_it_falls_back_to_the_newest(self):
        """The documented default, and now ordered rather than incidental."""
        Semester.set_current(None)
        self.assertEqual(
            self.client.get("/allocation/groups/").context["semester"],
            self.other_semester,
        )

    def test_a_current_semester_with_no_data_is_still_honoured(self):
        """Unlike the exports, which skip an empty term to avoid a blank sheet.

        An empty board is the honest answer for a term nothing has been
        imported into; quietly listing a *different* term's groups is not.
        """
        empty = Semester.objects.create(academic_year="2025/2026", semester=2)
        Semester.set_current(empty)
        resp = self.client.get("/allocation/groups/")
        self.assertEqual(resp.context["semester"], empty)
        self.assertEqual(resp.context["rows"], [])

    def test_a_database_with_no_semesters_at_all_is_still_graceful(self):
        Semester.objects.all().delete()
        for url in ("/allocation/", "/allocation/groups/"):
            with self.subTest(url=url):
                self.assertEqual(self.client.get(url).status_code, 200)


class AllocationScopePickerTests(AllocationTestCase):
    def setUp(self):
        super().setUp()
        self.client.force_login(
            User.objects.create_user(
                "scope-picker-tester", password="pw", is_staff=True
            )
        )
        self._seed()

    def test_the_widest_scope_option_is_just_all_activities(self):
        """It sat in a bare flex row, so the longest option set the page width."""
        for url in (
            "/allocation/",
            "/allocation/groups/",
            f"/allocation/group/{self.a1.pk}/",
        ):
            with self.subTest(url=url):
                resp = self.client.get(url)
                self.assertContains(resp, "All activities")
                self.assertNotContains(resp, "All configured activities")

    def test_the_navigation_bar_cannot_widen_the_page(self):
        """The nav bar's own component, and the one thing that actually bounds it.

        Its tracks are `minmax(0, 1fr)`, not `1fr`. A bare `1fr` in CSS grid is
        really `minmax(auto, 1fr)`, and the `auto` minimum is the content's
        min-content size -- a <select> wants to be as wide as its widest
        option, so that minimum sets the track's floor, the track outgrows the
        container and the document scrolls sideways. Zeroing the minimum is the
        fix; per-element `min-w-0` alone never was, because it bounds the
        control without bounding the track it sits in.

        Asserted on the partial, so the two pages cannot drift apart, plus on
        each served page, so the partial cannot be quietly dropped.
        """
        partial = (
            Path(settings.BASE_DIR) / "templates" / "core" / "_allocation_nav.html"
        ).read_text(encoding="utf-8")
        # The layout classes, not the whole file: the note above them explains
        # the rule in prose and would otherwise be counted as more tracks.
        form = re.search(r'<form method="get".*?>', partial, re.S).group(0)
        tight = re.search(r'class="([^"]*)"', form, re.S).group(1)
        tight = "".join(tight.split())
        self.assertIn("minmax(0,1fr)", tight)
        self.assertEqual(
            tight.count("1fr"),
            tight.count("minmax(0,1fr)"),
            "a grid track is not minmax(0, 1fr)",
        )

        for url in (
            "/allocation/groups/",
            f"/allocation/group/{self.a1.pk}/",
        ):
            with self.subTest(url=url):
                html = self.client.get(url).content.decode()
                self.assertIn("minmax(0,1fr)", html.replace(" ", ""))
                for name in ("scope", "semester"):
                    tag = re.search(
                        r"<select name=\"%s\".*?>" % name, html, re.S
                    ).group(0)
                    self.assertIn("w-full", tag)
                    self.assertIn("min-w-0", tag)

    def test_the_navigation_bar_is_its_own_component(self):
        """One partial, included by both pages, so the fix lives in one place."""
        for page in ("allocation_groups.html", "allocation_group.html"):
            with self.subTest(page=page):
                body = (
                    Path(settings.BASE_DIR) / "templates" / "core" / page
                ).read_text(encoding="utf-8")
                self.assertIn('_allocation_nav.html', body)
                # The pages must not carry their own copy of the pickers.
                self.assertNotIn('<select name="scope"', body)

    def test_the_navigation_button_is_the_same_height_as_the_pickers(self):
        """The bar read as a misaligned strip because the button was taller."""
        html = self.client.get("/allocation/groups/").content.decode()
        button = re.search(r"<a [^>]*?>\s*Run the allocator", html, re.S)
        self.assertIsNotNone(button, "the allocator button is missing")
        tag = re.search(r"<a [^>]*>", button.group(0), re.S).group(0)
        # Same vertical padding as the selects, or the row is a ragged edge.
        self.assertIn("py-2.5", tag)


class GroupPlacementPageTests(AllocationTestCase):
    def setUp(self):
        super().setUp()
        self._seed()
        self.monday = self._tutorial("MONDAY", 8, 9, self.hall)
        # NB102 seats 30 -- room for exactly one group.
        self.tuesday = self._tutorial("TUESDAY", 8, 9, self.small)
        self.tiny = self._tutorial("WEDNESDAY", 8, 9, self.no_capacity)

    def _page(self, group=None, **params):
        group = group or self.a1
        query = {"semester": self.semester.pk, "scope": "ALL", **params}
        return self.client.get(
            f"/allocation/group/{group.pk}/?", query
        )

    def test_the_page_lists_the_sessions_the_group_is_free_for(self):
        resp = self._page()
        self.assertEqual(resp.status_code, 200)
        entry = resp.context["status"].entries[0]
        free = [o.session for o in entry.free_options]
        self.assertEqual(free, [self.monday, self.tuesday])
        self.assertContains(resp, "Assign here")

    def test_a_room_that_is_already_full_is_refused_with_its_capacity(self):
        manual_assign(self.a1, self.tuesday)  # fills NB102's single seat
        resp = self._page(self.a2)
        reasons = " ".join(
            r.message for e in resp.context["status"].entries for o in e.options
            for r in o.reasons
        )
        self.assertIn("seats 30", reasons)
        self.assertContains(resp, "not possible")

    def test_a_room_with_no_recorded_capacity_is_refused_not_assumed(self):
        resp = self._page()
        self.assertEqual(
            [o.ok for o in resp.context["status"].entries[0].options
             if o.session == self.tiny],
            [False],
        )
        self.assertContains(resp, "no capacity recorded")

    def test_a_clash_is_reported_as_a_clash_not_as_capacity(self):
        manual_assign(self.a1, self.monday)
        clash = self._tutorial("MONDAY", 8, (9, 30), self.hall)
        entry = self._page().context["status"].entries[0]
        blocked = {o.session.pk: o for o in entry.options if not o.ok}
        self.assertIn(clash.pk, blocked)
        self.assertFalse(blocked[clash.pk].free)
        self.assertIn(
            "busy", " ".join(r.message for r in blocked[clash.pk].reasons).lower()
        )

    def test_assigning_from_the_page_creates_the_link(self):
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": self.tuesday.pk},
            HTTP_HX_REQUEST="true",
        )
        self.assertContains(resp, "assigned to")
        self.assertTrue(
            SessionGroup.objects.filter(
                session=self.tuesday, group=self.a1
            ).exists()
        )
        # And the page now shows it as done.
        entry = self._page().context["status"].entries[0]
        self.assertTrue(entry.is_met)
        self.assertEqual(entry.assigned, self.tuesday)

    def test_assigning_returns_the_fresh_panel_so_the_page_needs_no_reload(self):
        """The whole point: the response is the new status, not just a message.

        Without this the coordinator clicks "Assign here", sees a toast saying
        it worked, and is still looking at "Not yet placed" underneath.
        """
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": self.monday.pk, "panel": "1"},
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, 'id="allocation-requirements"')
        self.assertNotContains(resp, "Not yet placed")
        self.assertContains(resp, "Complete")  # this is the group's only one
        self.assertIn("allocation-toast", resp["HX-Trigger"])

    def test_removing_returns_the_fresh_panel_with_the_requirement_open_again(self):
        manual_assign(self.a1, self.monday)
        resp = self.client.post(
            "/allocation/unassign/",
            {"group": self.a1.pk, "session": self.monday.pk, "panel": "1"},
            HTTP_HX_REQUEST="true",
        )
        self.assertContains(resp, "Not yet placed")
        self.assertContains(resp, "1 still to place")
        self.assertContains(resp, "0/1")
        self.assertNotContains(resp, "Complete")

    def test_the_progress_bar_comes_back_updated_with_the_panel(self):
        self.maths.set_requirements({"TUTORIAL": 1, "PRACTICAL": 1})
        tutorial = self._tutorial("MONDAY", 8, 9, self.hall)
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": tutorial.pk, "panel": "1"},
            HTTP_HX_REQUEST="true",
        )
        self.assertContains(resp, "1/2")
        self.assertContains(resp, "1 still to place")

    def test_a_plain_post_from_the_group_page_goes_back_to_the_group_page(self):
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": self.monday.pk, "panel": "1"},
        )
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(
            resp.url, f"/allocation/group/{self.a1.pk}/?semester={self.semester.pk}"
        )

    def test_the_forms_carry_the_view_so_the_redraw_matches_the_page(self):
        """Two semesters exist; the panel must come back for the one on screen.

        Otherwise the list silently switches to the latest semester's sessions
        the instant a placement is made.
        """
        resp = self._page()
        self.assertContains(
            resp, f'name="semester" value="{self.semester.pk}"'
        )
        self.assertContains(resp, 'name="scope" value="ALL"')
        other = self._session(
            "MT161", ActivityType.TUTORIAL, "THURSDAY", 14, 15, self.hall
        )
        other.semester = self.other_semester
        other.save()
        resp = self.client.post(
            "/allocation/assign/",
            {
                "group": self.a1.pk,
                "session": self.monday.pk,
                "panel": "1",
                "semester": self.semester.pk,
                "scope": "ALL",
            },
            HTTP_HX_REQUEST="true",
        )
        self.assertNotContains(resp, "THURSDAY")
        self.assertNotContains(resp, "14:00")

    def test_a_panel_post_with_no_semester_uses_the_session_own(self):
        """The session names its semester outright, so nothing is guessed."""
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": self.monday.pk, "panel": "1"},
            HTTP_HX_REQUEST="true",
        )
        self.assertEqual(
            resp.context["semester"], self.semester
        )
        self.assertNotContains(resp, "Not yet placed")

    def test_a_post_without_the_panel_flag_still_gets_the_message_fragment(self):
        """The allocation page's own manual form has no panel to replace."""
        resp = self.client.post(
            "/allocation/assign/",
            {"group": self.a1.pk, "session": self.monday.pk},
            HTTP_HX_REQUEST="true",
        )
        self.assertContains(resp, "assigned to")
        self.assertNotContains(resp, 'id="allocation-requirements"')

    def test_the_panel_is_served_on_its_own(self):
        resp = self.client.get(
            f"/allocation/group/{self.a1.pk}/panel/",
            {"semester": self.semester.pk, "scope": "ALL"},
        )
        self.assertEqual(resp.status_code, 200)
        self.assertContains(resp, 'id="allocation-requirements"')
        self.assertContains(resp, "Not yet placed")
        self.assertEqual(self.client.get("/allocation/group/999999/panel/").status_code, 404)

    def test_a_placed_requirement_offers_a_move_and_a_remove(self):
        manual_assign(self.a1, self.monday)
        resp = self._page()
        self.assertContains(resp, "Move it: see the other sessions")
        self.assertContains(resp, "Remove")
        self.client.post(
            "/allocation/unassign/",
            {"group": self.a1.pk, "session": self.monday.pk},
            HTTP_HX_REQUEST="true",
        )
        self.assertFalse(
            SessionGroup.objects.filter(session=self.monday, group=self.a1).exists()
        )
        # With the link gone the requirement is outstanding again, and the page
        # says so rather than still claiming it is done.
        self.assertFalse(self._page().context["status"].entries[0].is_met)

    def test_a_group_with_nothing_required_is_said_so(self):
        self.maths.set_requirements({})
        resp = self._page()
        self.assertContains(resp, "no seminar, tutorial or practical requirement")
        self.assertEqual(resp.context["status"].total, 0)

    def test_a_group_whose_course_has_no_requirement_configured_is_flagged(self):
        self.maths.set_requirements({})
        resp = self._page()
        self.assertContains(resp, "no required activities configured")
        self.assertContains(resp, "MT161")

    def test_a_group_with_no_session_at_all_is_told_so_rather_than_shown_blank(self):
        Session.objects.all().delete()
        resp = self._page()
        entry = resp.context["status"].entries[0]
        self.assertEqual(entry.options, [])
        self.assertContains(resp, "No tutorial session exists")

    def test_an_unknown_group_is_a_404(self):
        self.assertEqual(
            self.client.get("/allocation/group/999999/").status_code, 404
        )

    def test_the_page_survives_a_database_with_no_semester(self):
        Semester.objects.all().delete()
        resp = self.client.get(f"/allocation/group/{self.a1.pk}/")
        self.assertEqual(resp.status_code, 200)
        self.assertIsNone(resp.context["semester"])
        self.assertIsNone(resp.context["status"])
        self.assertContains(resp, "no seminar, tutorial or practical requirement")


class CourseRequirementPageTests(TestCase):
    def setUp(self):
        self.course = Course.objects.create(code="MT161", name="Mathematics 1")
        self.prog = Programme.objects.create(code="CE", name="Civil Engineering")
        ProgrammeCourse.objects.create(
            programme=self.prog, course=self.course, semester=1
        )

    @staticmethod
    def _formset_payload(rows, total=None, initial=0):
        """POST payload for the inline requirement formset."""
        total = len(rows) if total is None else total
        payload = {
            "activity_requirements-TOTAL_FORMS": str(total),
            "activity_requirements-INITIAL_FORMS": str(initial),
            "activity_requirements-MIN_NUM_FORMS": "0",
            "activity_requirements-MAX_NUM_FORMS": "1000",
        }
        for index, row in enumerate(rows):
            payload[f"activity_requirements-{index}-id"] = row.get("id", "")
            payload[f"activity_requirements-{index}-activity_type"] = row["activity_type"]
            payload[f"activity_requirements-{index}-count"] = str(row["count"])
            payload[f"activity_requirements-{index}-DELETE"] = row.get("delete", "")
        return payload

    def test_the_list_shows_the_shared_requirements(self):
        self.course.set_requirements({"TUTORIAL": 1, "PRACTICAL": 1})
        resp = self.client.get("/course-requirements/")
        self.assertContains(resp, "MT161")
        self.assertContains(resp, "Tutorial; Practical")

    def test_the_list_flags_a_course_with_no_requirement(self):
        resp = self.client.get("/course-requirements/")
        self.assertContains(resp, "Not configured")

    def test_the_list_can_filter_to_unconfigured_courses(self):
        self.course.set_requirements({"TUTORIAL": 1})
        Course.objects.create(code="ZZ999", name="Unconfigured")
        resp = self.client.get("/course-requirements/?configured=no")
        self.assertContains(resp, "ZZ999")
        self.assertNotContains(resp, "MT161")

    def test_creating_a_course_saves_its_requirements(self):
        resp = self.client.post(
            "/course-requirements/create/",
            {
                "code": "EE131",
                "name": "Electronics",
                **self._formset_payload(
                    [{"activity_type": "SEMINAR", "count": 1}]
                ),
            },
        )
        self.assertEqual(resp.status_code, 302)
        course = Course.objects.get(code="EE131")
        self.assertEqual(course.required_activities(), ("SEMINAR",))

    def test_creating_a_course_with_no_requirements_is_allowed(self):
        resp = self.client.post(
            "/course-requirements/create/",
            {
                "code": "EE131",
                "name": "Electronics",
                **self._formset_payload([]),
            },
        )
        self.assertEqual(resp.status_code, 302)
        self.assertFalse(Course.objects.get(code="EE131").has_requirements())

    def test_editing_a_course_replaces_its_requirements(self):
        existing = self.course.set_requirements({"TUTORIAL": 1})
        row = self.course.activity_requirements.get(activity_type="TUTORIAL")
        resp = self.client.post(
            f"/course-requirements/{self.course.pk}/edit/",
            {
                "code": "MT161",
                "name": "Mathematics 1",
                **self._formset_payload(
                    [{"id": row.pk, "activity_type": "PRACTICAL", "count": 2}],
                    initial=1,
                ),
            },
        )
        self.assertEqual(resp.status_code, 302)
        self.course = Course.objects.get(pk=self.course.pk)
        self.assertEqual(self.course.required_activities(), ("PRACTICAL",))
        self.assertEqual(self.course.required_count("PRACTICAL"), 2)
        self.assertTrue(existing is not None)

    def test_the_requirement_form_never_offers_a_lecture(self):
        resp = self.client.get("/course-requirements/create/")
        self.assertContains(resp, "Tutorial")
        self.assertNotContains(resp, ">Lecture<")

    def test_the_detail_page_lists_the_programmes_and_name_variants(self):
        self.course.name_variants = json.dumps(["Mathematics 1A"])
        self.course.save()
        resp = self.client.get(f"/course-requirements/{self.course.pk}/")
        self.assertContains(resp, "CE")
        self.assertContains(resp, "Mathematics 1A")
        self.assertContains(resp, "other name")

    def test_deleting_a_course_also_removes_its_programme_links(self):
        resp = self.client.post(f"/course-requirements/{self.course.pk}/delete/")
        self.assertEqual(resp.status_code, 302, getattr(resp, "context", None))
        self.assertFalse(Course.objects.filter(code="MT161").exists())
        self.assertEqual(ProgrammeCourse.objects.count(), 0)

    def test_deleting_a_course_is_logged(self):
        self.client.post(f"/course-requirements/{self.course.pk}/delete/")
        self.assertTrue(
            ActivityLog.objects.filter(
                action=LogAction.DELETE, resource="Course"
            ).exists()
        )

    def test_the_programme_course_list_shows_the_shared_requirements(self):
        self.course.set_requirements({"PRACTICAL": 1})
        resp = self.client.get("/courses/")
        self.assertContains(resp, "Practical")

    def test_the_sidebar_links_to_the_shared_courses(self):
        resp = self.client.get("/course-requirements/")
        self.assertContains(resp, 'href="/course-requirements/"')


class DeriveRequirementsFromSessionsTests(TestCase):
    """The master-timetable import can fill in requirements from what it read."""

    def setUp(self):
        self.semester = Semester.objects.create(
            academic_year="2026/2027", semester=1
        )
        self.prog = Programme.objects.create(code="CE", name="Civil Engineering")
        self.group = StudentGroup.objects.create(programme=self.prog, code="A1")

    def _course(self, code, requirements=None, semester=1):
        course = Course.objects.create(code=code, name=code)
        ProgrammeCourse.objects.create(
            programme=self.prog, course=course, semester=semester
        )
        if requirements:
            course.set_requirements(requirements)
        return course

    def _session(self, code, activity):
        return Session.objects.create(
            semester=self.semester,
            course_code=code,
            activity_type=activity,
            day="MONDAY",
            start_time=time(8, 0),
            end_time=time(9, 0),
            venue=None,
        )

    def test_a_course_with_only_seminars_requires_a_seminar(self):
        course = self._course("CL111")
        self._session("CL111", ActivityType.LECTURE)
        self._session("CL111", ActivityType.SEMINAR)
        mapping, unresolved = derive_requirements_from_sessions(self.semester)
        self.assertEqual(mapping[course], {"SEMINAR": 1})
        self.assertNotIn(course, unresolved)

    def test_a_course_with_tutorials_and_practicals_requires_both(self):
        course = self._course("ME101")
        self._session("ME101", ActivityType.LECTURE)
        self._session("ME101", ActivityType.TUTORIAL)
        self._session("ME101", ActivityType.PRACTICAL)
        mapping, _ = derive_requirements_from_sessions(self.semester)
        self.assertEqual(mapping[course], {"TUTORIAL": 1, "PRACTICAL": 1})

    def test_lectures_alone_never_become_a_requirement(self):
        course = self._course("MT171")
        self._session("MT171", ActivityType.LECTURE)
        mapping, unresolved = derive_requirements_from_sessions(self.semester)
        self.assertNotIn(course, mapping)
        self.assertIn(course, unresolved)

    def test_workshops_never_become_a_requirement(self):
        course = self._course("WT107")
        self._session("WT107", ActivityType.WORKSHOP)
        mapping, unresolved = derive_requirements_from_sessions(self.semester)
        self.assertNotIn(course, mapping)
        self.assertIn(course, unresolved)

    def test_a_course_with_no_sessions_is_reported_not_assumed_empty(self):
        course = self._course("ZZ999")
        _, unresolved = derive_requirements_from_sessions(self.semester)
        self.assertIn(course, unresolved)

    def test_deriving_never_guesses_an_activity_with_no_session(self):
        # A course whose tutorials are not timetabled yet must not be given a
        # tutorial requirement nothing could satisfy.
        course = self._course("EE131")
        self._session("EE131", ActivityType.LECTURE)
        mapping, _ = derive_requirements_from_sessions(self.semester)
        self.assertNotIn(course, mapping)

    def test_the_import_sets_requirements_when_asked(self):
        self._course("CL111")
        self._session("CL111", ActivityType.SEMINAR)
        self._session("CL111", ActivityType.LECTURE)
        path = make_xlsx(
            [
                ["CL111", "SEMINAR", "MONDAY", "08:00", "09:00", "", ""],
                ["CL111", "LECTURE", "MONDAY", "10:00", "11:00", "", ""],
            ],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(
            path,
            semester_id=self.semester.pk,
            derive_requirements=True,
        )
        course = Course.objects.get(code="CL111")
        self.assertEqual(course.required_activities(), ("SEMINAR",))
        self.assertEqual(
            [e["code"] for e in result.derived_requirements], ["CL111"]
        )

    def test_the_import_never_overwrites_an_existing_requirement(self):
        course = self._course("MT161", requirements={"PRACTICAL": 2})
        self._session("MT161", ActivityType.TUTORIAL)
        self._session("MT161", ActivityType.LECTURE)
        path = make_xlsx(
            [["MT161", "TUTORIAL", "MONDAY", "08:00", "09:00", "", ""]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(
            path,
            semester_id=self.semester.pk,
            derive_requirements=True,
        )
        course.refresh_from_db()
        self.assertEqual(course.required_activities(), ("PRACTICAL",))
        self.assertEqual(course.required_count("PRACTICAL"), 2)
        self.assertTrue(
            any("already has" in line for line in result.derived_skipped),
            result.derived_skipped,
        )

    def test_the_import_reports_courses_it_could_not_decide(self):
        self._course("ZZ999")
        self._course("CL111")
        self._session("CL111", ActivityType.SEMINAR)
        path = make_xlsx(
            [["CL111", "SEMINAR", "MONDAY", "08:00", "09:00", "", ""]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(
            path,
            semester_id=self.semester.pk,
            derive_requirements=True,
        )
        self.assertTrue(
            any("ZZ999" in line for line in result.derived_unresolved),
            result.derived_unresolved,
        )

    def test_deriving_is_off_unless_asked_for(self):
        self._course("CL111")
        self._session("CL111", ActivityType.SEMINAR)
        self._session("CL111", ActivityType.LECTURE)
        path = make_xlsx(
            [["CL111", "SEMINAR", "MONDAY", "08:00", "09:00", "", ""]],
            MASTER_COLS,
        )
        import_master_timetable_from_excel(
            path, semester_id=self.semester.pk
        )
        self.assertFalse(Course.objects.get(code="CL111").has_requirements())

    def test_a_dry_run_never_derives(self):
        self._course("CL111")
        self._session("CL111", ActivityType.SEMINAR)
        self._session("CL111", ActivityType.LECTURE)
        path = make_xlsx(
            [["CL111", "SEMINAR", "MONDAY", "08:00", "09:00", "", ""]],
            MASTER_COLS,
        )
        result = import_master_timetable_from_excel(
            path, semester_id=self.semester.pk, dry_run=True,
            derive_requirements=True,
        )
        self.assertFalse(Course.objects.get(code="CL111").has_requirements())
        self.assertEqual(result.derived_requirements, [])


class AllocationShowsInTheTimetableTests(AllocationTestCase):
    """The point of the whole feature: an applied run must be visible in the
    master timetable, in the session detail, and in every PDF export.

    Two tutorial rooms, so the four groups are *split* across them. With one
    big room every group lands in the same session and every export correctly
    reads ALL, which would make these assertions vacuous.
    """

    def _applied(self, scope="ALL"):
        self._tutorial("MONDAY", 8, 9, self.hall)
        self._tutorial("TUESDAY", 8, 9, self.hall)
        plan = plan_allocation(self.semester, scope)
        return plan, apply_run(save_plan(plan, algorithm="smart"))

    def test_the_sessions_list_shows_the_assigned_groups(self):
        self._seed()
        self._applied()
        resp = self.client.get("/sessions/")
        self.assertContains(resp, "Assigned Groups")
        codes = {
            code
            for session in resp.context["items"]
            for code in [session.assigned_groups]
            if code != "—"
        }
        self.assertTrue(codes, "no session showed any assigned group")
        # Groups really were split, so real codes (not just ALL) are listed.
        self.assertTrue(
            any("," in text for text in codes), codes
        )

    def test_the_sessions_list_marks_which_sessions_are_allocatable(self):
        self._seed()
        self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 10, self.hall
        )
        self._tutorial("MONDAY", 14, 15, self.hall)
        resp = self.client.get("/sessions/")
        self.assertIn(
            {"label": "Allocatable", "key": "allocatable"},
            resp.context["detail_fields"],
        )
        # The detail panel is what renders that field.
        session = Session.objects.get(activity_type=ActivityType.LECTURE)
        panel = self.client.get(
            f"/sessions/{session.pk}/", HTTP_HX_REQUEST="true"
        )
        self.assertContains(panel, "Allocatable")
        self.assertContains(panel, "No")
        tutorial = Session.objects.get(activity_type=ActivityType.TUTORIAL)
        self.assertContains(
            self.client.get(f"/sessions/{tutorial.pk}/", HTTP_HX_REQUEST="true"),
            "Yes",
        )

    def test_the_sessions_list_counts_the_unassigned_small_group_sessions(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        resp = self.client.get("/sessions/")
        self.assertEqual(resp.context["unassigned_count"], 1)
        self.assertContains(resp, "no group assigned")

    def test_the_unassigned_filter_lists_only_empty_small_group_sessions(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        self._tutorial("TUESDAY", 8, 9, self.hall)
        # A session nothing can be placed in: no venue means no capacity, so it
        # stays empty whatever the allocator does.
        blocked = self._tutorial("WEDNESDAY", 8, 9, venue=None)
        self._session(
            "MT161", ActivityType.LECTURE, "THURSDAY", 8, 10, self.hall
        )
        self._applied()
        resp = self.client.get("/sessions/?assigned=no")
        pks = {s.pk for s in resp.context["items"]}
        self.assertIn(blocked.pk, pks)
        self.assertNotIn("MONDAY", {s.day for s in resp.context["items"]})
        lectures = Session.objects.filter(activity_type=ActivityType.LECTURE)
        self.assertFalse(pks & set(lectures.values_list("pk", flat=True)))

    def test_the_assigned_filter_lists_only_filled_sessions(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        self._tutorial("TUESDAY", 8, 9, self.hall)
        self._applied()
        resp = self.client.get("/sessions/?assigned=yes")
        self.assertTrue(resp.context["items"])
        for session in resp.context["items"]:
            self.assertNotEqual(session.assigned_groups, "—")

    def test_the_session_detail_page_lists_the_assigned_groups(self):
        self._seed()
        self._applied()
        session = Session.objects.get(
            activity_type=ActivityType.TUTORIAL, day="MONDAY"
        )
        codes = list(
            session.session_groups.values_list("group__code", flat=True)
        )
        self.assertTrue(codes)
        resp = self.client.get(f"/sessions/{session.pk}/")
        for code in codes:
            self.assertContains(resp, code)
        htmx = self.client.get(
            f"/sessions/{session.pk}/", HTTP_HX_REQUEST="true"
        )
        for code in codes:
            self.assertContains(htmx, code)

    def test_the_group_pdf_is_for_that_one_group(self):
        self._seed()
        self._applied()
        entries = collect_group_entries(self.a1, self.semester)
        labels = [e["label"] for e in entries if e["course_code"] == "MT161"]
        self.assertTrue(labels)
        self.assertIn("A1", labels[0])
        others = [
            g.code
            for g in StudentGroup.objects.exclude(pk=self.a1.pk)
        ]
        for code in others:
            self.assertNotIn(code, labels[0])

    def test_the_programme_pdf_names_that_programmes_groups_only(self):
        self._seed()
        self._applied()
        mine = set(
            StudentGroup.objects.filter(programme=self.ce).values_list(
                "code", flat=True
            )
        )
        entries = fold_for_display(
            collect_entries(self.ce, self.semester), mine
        )
        cell = next(e for e in entries if e["course_code"] == "MT161")
        listed = {code for code in mine if code in cell["label"]}
        self.assertTrue(listed, cell["label"])
        theirs = set(
            StudentGroup.objects.exclude(programme=self.ce).values_list(
                "code", flat=True
            )
        )
        self.assertFalse(
            theirs & {c for c in cell["label"].replace("ALL", " ").split(",")},
            cell["label"],
        )

    def test_the_all_programmes_pdf_names_every_attending_group(self):
        self._seed()
        self._applied()
        all_codes = set(StudentGroup.objects.values_list("code", flat=True))
        merged = _merge_master_entries(
            fold_for_display(
                collect_master_entries(self.semester), all_codes
            ),
            all_codes,
        )
        blocks = list(merged.values()) if isinstance(merged, dict) else list(merged)
        text = " ".join(_master_cell_text([b]) for b in blocks)
        for code in all_codes:
            self.assertIn(code, text)

    def test_the_all_programmes_pdf_export_renders_with_the_groups(self):
        self._seed()
        self._applied()
        response = self.client.get(
            f"/export/all-programmes/timetable.pdf/?semester={self.semester.pk}"
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "application/pdf")
        self.assertTrue(response.content.startswith(b"%PDF"))

    def test_the_group_export_draws_no_groups_line(self):
        # One group's own sheet: naming the group back to it — or printing
        # "ALL" because that one group is everybody — is noise.
        self._seed()
        self._applied()
        entries, _ = _collect_group_entries_and_rotations(self.a1, self.semester)
        folded = fold_for_display(
            entries, {self.a1.code}, show_groups=False
        )
        self.assertTrue(folded)
        for entry in folded:
            lines = [ln for ln in entry["label"].split("\n") if ln]
            self.assertNotIn("ALL", lines, entry["label"])
            for group in StudentGroup.objects.all():
                self.assertNotIn(group.code, lines, entry["label"])
        # ...and the data is still there for a caller that wants it.
        self.assertTrue(any(e.get("groups") for e in folded))

    def test_the_group_export_still_states_the_course_and_venue(self):
        self._seed()
        self._applied()
        entries, _ = _collect_group_entries_and_rotations(self.a1, self.semester)
        folded = fold_for_display(
            entries, {self.a1.code}, show_groups=False
        )
        cell = next(e for e in folded if e["course_code"] == "MT161")
        self.assertIn("MT161", cell["label"])
        self.assertIn("Tutorial", cell["label"])
        self.assertIn(self.hall.name, cell["label"])

    def test_the_other_exports_keep_their_groups_line(self):
        self._seed()
        self._applied()
        mine = {
            g.code
            for g in StudentGroup.objects.filter(programme=self.ce)
        }
        programme = fold_for_display(
            collect_entries(self.ce, self.semester), mine, show_groups=True
        )
        self.assertTrue(any(e.get("groups") for e in programme))
        all_codes = set(StudentGroup.objects.values_list("code", flat=True))
        master = fold_for_display(
            collect_master_entries(self.semester), all_codes, show_groups=True
        )
        self.assertTrue(any(e.get("groups") for e in master))

    def test_a_whole_cohort_lecture_also_loses_its_all_on_a_group_export(self):
        # CL111 lectures read ALL by design on a shared sheet; on a single
        # group's own sheet the line is gone entirely, not merely relabelled.
        self._seed()
        self.maths.set_requirements({"SEMINAR": 1})
        lecture = self._session(
            "MT161", ActivityType.LECTURE, "MONDAY", 8, 10, self.hall
        )
        SessionGroup.objects.create(session=lecture, group=self.a1)
        self._session("MT161", ActivityType.SEMINAR, "TUESDAY", 8, 9, self.hall)
        plan = plan_allocation(self.semester, "ALL")
        apply_run(save_plan(plan, algorithm="smart"))
        entries, _ = _collect_group_entries_and_rotations(self.a1, self.semester)
        with_groups = fold_for_display(entries, {self.a1.code})
        without = fold_for_display(
            entries, {self.a1.code}, show_groups=False
        )
        shared = next(e for e in with_groups if e["groups"] == "ALL")
        solo = next(e for e in without if e["course_code"] == shared["course_code"])
        self.assertNotIn("ALL", solo["label"])
        # The shared export is untouched.
        self.assertIn("ALL", shared["label"])

    def test_reverting_takes_the_groups_back_out_of_the_views_and_exports(self):
        self._seed()
        self._applied()
        self.assertTrue(SessionGroup.objects.exists())
        self.assertNotContains(self.client.get("/sessions/"), "A1, A2")
        run = AllocationRun.objects.get(status=AllocationStatus.APPLIED)
        result = revert_run(run)
        self.assertTrue(result["ok"])
        self.assertEqual(SessionGroup.objects.count(), 0)
        resp = self.client.get("/sessions/")
        self.assertContains(resp, "no group assigned")
        self.assertEqual(
            [
                e
                for e in collect_group_entries(self.a1, self.semester)
                if e.get("groups")
            ],
            [],
        )


class AllocationAppearsInExportsTests(AllocationTestCase):
    def test_an_applied_tutorial_shows_up_in_the_group_timetable(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.hall)
        plan = plan_allocation(self.semester, "ALL")
        apply_run(save_plan(plan, algorithm="smart"))
        entries = collect_group_entries(self.a1, self.semester)
        codes = {e["course_code"] for e in entries}
        self.assertIn("MT161", codes)
        self.assertIn("A1", {e.get("groups", "") for e in entries})

    def test_another_group_does_not_see_someone_elsses_tutorial(self):
        self._seed()
        # Only A1 is allocated; D1's requirement stays unresolved.
        self._tutorial("MONDAY", 8, 9, self.small)
        plan = plan_allocation(self.semester, "ALL")
        apply_run(save_plan(plan, algorithm="smart"))
        self.assertIn(
            "MT161", {e["course_code"] for e in collect_group_entries(self.a1, self.semester)}
        )
        self.assertNotIn(
            "MT161",
            {e["course_code"] for e in collect_group_entries(self.d1, self.semester)},
        )

    def test_the_group_pdf_export_includes_the_applied_session(self):
        self._seed()
        self._tutorial("MONDAY", 8, 9, self.big)
        plan = plan_allocation(self.semester, "ALL")
        apply_run(save_plan(plan, algorithm="smart"))
        out = io.BytesIO()
        render_group_timetable(self.a1, self.semester, out=out)
        self.assertGreater(len(out.getvalue()), 1000)

class CurrentSemesterTests(TestCase):
    """One current semester, set from the semesters list and used as the default.

    Staff were previously told to pick a term on every page, and the only
    persisted default lived in the student-portal settings, reachable only
    through the Django admin and only for superusers. These cover the value
    being single, being settable from the list, and being the fallback the
    dashboard, the exports and the portal actually read.
    """

    def setUp(self):
        self.staff = User.objects.create_user(
            "semester-tester", password="pw", is_staff=True
        )
        self.client = Client()
        self.client.force_login(self.staff)
        self.old = Semester.objects.create(academic_year="2025/2026", semester=2)
        self.new = Semester.objects.create(academic_year="2026/2027", semester=1)

    def test_no_semester_is_current_until_one_is_chosen(self):
        self.assertIsNone(Semester.current())
        Semester.set_current(self.new)
        self.assertEqual(Semester.current(), self.new)
        self.assertEqual(Semester.objects.filter(is_current=True).count(), 1)

    def test_setting_a_second_semester_never_leaves_two_current(self):
        Semester.set_current(self.new)
        Semester.set_current(self.old)
        self.assertEqual(Semester.current(), self.old)
        self.assertFalse(Semester.objects.get(pk=self.new.pk).is_current)

    def test_clearing_leaves_nothing_current_rather_than_guessing(self):
        Semester.set_current(self.new)
        Semester.set_current(None)
        self.assertIsNone(Semester.current())
        self.assertEqual(Semester.objects.filter(is_current=True).count(), 0)

    def test_the_database_refuses_a_second_current_semester(self):
        """The flag is not merely checked in Python: the constraint is the backstop.

        Every in-app writer goes through set_current, but a raw save() -- a
        management command, a fixture, a future view -- must not be able to
        leave a page reading two different "current" terms.
        """
        Semester.objects.filter(pk=self.new.pk).update(is_current=True)
        with self.assertRaises(IntegrityError), transaction.atomic():
            Semester.objects.filter(pk=self.old.pk).update(is_current=True)

    def _post(self, semester):
        return self.client.post(
            "/semesters/%d/current/" % semester.pk, HTTP_HX_REQUEST="true"
        )

    def test_the_list_offers_a_make_and_an_unset_action_per_row(self):
        html = self.client.get("/semesters/").content.decode()
        # The action states what it will do, so a second click cannot silently
        # clear the term the app is working in.
        self.assertIn("Make current", html)
        self.assertIn('hx-post="/semesters/%d/current/"' % self.new.pk, html)
        self.assertIn("No current semester set", html)

    def test_the_current_row_offers_unset_instead_of_make(self):
        Semester.set_current(self.new)
        html = self.client.get("/semesters/").content.decode()
        self.assertIn("Current semester: 2026/2027 - Semester 1", html)
        self.assertIn("Unset current", html)

    def test_the_action_sets_the_semester_and_logs_it(self):
        resp = self._post(self.new)
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(Semester.current(), self.new)
        self.assertEqual(resp["HX-Trigger"], "close-modal,refresh-table")
        self.assertTrue(
            ActivityLog.objects.filter(
                message__icontains="as the current semester"
            ).exists()
        )

    def test_the_action_twice_clears_the_flag(self):
        self._post(self.new)
        self._post(self.new)
        self.assertIsNone(Semester.current())
        self.assertTrue(
            ActivityLog.objects.filter(
                message__icontains="Unset the current semester"
            ).exists()
        )

    def test_a_get_only_bounces_to_the_list(self):
        resp = self.client.get("/semesters/%d/current/" % self.new.pk)
        self.assertEqual(resp.status_code, 302)
        self.assertIsNone(Semester.current())

    def test_the_dashboard_defaults_to_the_current_semester(self):
        Semester.set_current(self.old)
        resp = self.client.get("/staff/")
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(resp.context["selected_semester"], self.old)

    def test_the_dashboard_still_answers_to_an_explicit_picker(self):
        Semester.set_current(self.old)
        resp = self.client.get("/staff/", {"semester": self.new.pk})
        self.assertEqual(resp.context["selected_semester"], self.new)

    def test_the_import_form_preselects_the_current_semester(self):
        """The master-timetable import writes to exactly the semester named here,
        so the default has to be the term staff are working in, not row one."""
        Semester.set_current(self.old)
        html = self.client.get("/import/master-timetable/").content.decode()
        self.assertIn(
            '<option value="%d" selected' % self.old.pk, html
        )

    def test_an_export_falls_back_to_the_current_semester_when_it_has_data(self):
        Semester.set_current(self.new)
        venue = Venue.objects.create(name="LH1", capacity=80)
        Session.objects.create(
            semester=self.new,
            course_code="MT161",
            activity_type=ActivityType.LECTURE,
            day=Day.MONDAY,
            start_time=time(8, 0),
            end_time=time(10, 0),
            venue=venue,
        )
        self.assertEqual(_latest_semester_with_data(), self.new)

    def test_a_current_semester_with_no_data_does_not_export_a_blank_sheet(self):
        """Marking a term current before importing into it must not make the
        exports hand the reader an empty timetable: the newest semester that
        holds data still wins."""
        Session.objects.create(
            semester=self.old,
            course_code="MT161",
            activity_type=ActivityType.LECTURE,
            day=Day.MONDAY,
            start_time=time(8, 0),
            end_time=time(10, 0),
        )
        Semester.set_current(self.new)
        self.assertEqual(_latest_semester_with_data(), self.old)

    def test_the_portal_reads_the_app_wide_current_semester(self):
        from student_portal.views import current_semester

        Semester.set_current(self.old)
        self.assertEqual(current_semester(), self.old)

    def test_an_empty_portal_settings_row_does_not_mask_the_current_semester(self):
        """A settings row is a required, singleton override -- creating one is a
        deliberate act, so a fresh database leaves the portal on the app-wide
        current semester rather than on nothing."""
        from student_portal.models import PortalSettings
        from student_portal.views import current_semester

        self.assertEqual(PortalSettings.objects.count(), 0)
        Semester.set_current(self.old)
        self.assertEqual(current_semester(), self.old)

    def test_the_portal_override_still_wins_when_it_names_a_semester(self):
        """The portal can be pointed at a different term than the staff tools."""
        from student_portal.models import PortalSettings
        from student_portal.views import current_semester

        Semester.set_current(self.old)
        PortalSettings.objects.create(current_semester=self.new)
        self.assertEqual(current_semester(), self.new)
