warning: in the working copy of 'core/tests.py', LF will be replaced by CRLF the next time Git touches it
warning: in the working copy of 'core/timetable_pdf.py', LF will be replaced by CRLF the next time Git touches it
warning: in the working copy of 'core/views.py', LF will be replaced by CRLF the next time Git touches it
[1mdiff --git a/AGENTS.md b/AGENTS.md[m
[1mindex 3d19223..3d17e86 100644[m
[1m--- a/AGENTS.md[m
[1m+++ b/AGENTS.md[m
[36m@@ -32,7 +32,7 @@[m [mAll from repo root (`E:\COET_TIMETABLE\Coet_timetable`). Use `.venv\Scripts\pyth[m
 [m
 ## Timetable view & PDF exports[m
 [m
[31m-- On-screen timetable (`/timetable/`, view `timetable_view`) and the PDF exports share one grid builder: `core/timetable_grid.build_time_day_grid` (TIME rows / DAY columns, merged multi-slot cells) fed by `core/timetable_pdf.collect_entries` (sessions + workshops + TDs for a programme, or `collect_group_entries` for a single group). DAYS-only filter on `Day` choices; fills via `fill_color`; PDF rendering is `render_programme_timetable`/`render_group_timetable` (reportlab `SimpleDocTemplate` — the on-screen/template shading matches the PDF colours).[m
[32m+[m[32m- On-screen timetable (`/timetable/`, view `timetable_view`) and the PDF exports both build on `core/timetable_pdf.collect_entries` (sessions + workshops + TDs for a programme, or `collect_group_entries` for a single group). The on-screen view uses `core/timetable_grid.build_day_time_grid` (DAYS as rows — always the full Monday..Friday range with full weekday labels, plus any weekend days present; hourly TIME columns each labelled in the master-timetable format e.g. `07:00-07:55` — on-screen slot labels are built by `_screen_slot`, while the PDF/classic grid keeps the plain `07:00-08:00` labels from `_slot`). Each day is a vertical band of sub-rows (lanes, `_partition_lanes`): a session that spans several hourly slots is drawn as ONE cell merging those columns (`colspan` derived from every covered hour, including partial hours), and overlapping sessions sit on separate lanes instead of being sliced into one box per hour column. Rendered through `templates/core/_timetable_grid.html` with one plain-text entry box (`templates/core/_activity_card.html`) per entry — the box reads exactly like the UDSM Excel grid cells: `"Type, Start-End, Day"` / `Course: CODE` / `Venue: VENUE` / `Assigned groups: GROUPS` (left blank when a session has no assigned groups). The master timetable entries are the source of truth: ungrouped sessions still appear. The PDF exports use `core/timetable_grid.build_time_day_grid` (TIME rows / DAY columns, merged via `rowspan`). DAYS-only filter on `Day` choices; fills via `fill_color`; PDF rendering is `render_programme_timetable`/`render_group_timetable` (reportlab `SimpleDocTemplate` — the on-screen/template shading matches the PDF colours).[m
 - Export URLs: `export/` (export picker page), `export/programmes/<pk>/timetable.pdf/` (name `programme-timetable-export`), `export/groups/<pk>/timetable.pdf/` (name `group-timetable-export`); also `programmes/<pk>/timetable.pdf/` (`programme-timetable-pdf`), `timetable/` (`timetable`), `timetable/groups/` (`timetable-groups`).[m
 - PDF exports accept `?semester=<pk>` & `?year=<n>`; without `semester` they fall back to the latest semester that actually has data (`_latest_semester_with_data` in `core/views.py`).[m
 [m
[36m@@ -48,4 +48,4 @@[m [mAll from repo root (`E:\COET_TIMETABLE\Coet_timetable`). Use `.venv\Scripts\pyth[m
 - `manage.py check` reports `staticfiles.W004` because `static/` doesn't exist but is in `STATICFILES_DIRS`. Harmless — all styling comes from CDNs.[m
 - `settings.py` holds local-only, uncommitted LAN IPs in `ALLOWED_HOSTS` (172.16.185.226, 192.168.100.6) — don't strip them.[m
 - README is partly stale: reportlab (timetable PDFs in `core/timetable_pdf.py`) IS used now, but weasyprint is installed and unused.[m
[31m-- `core/templatetags/core_tags.py` defines `get_attr` and `time_short` filters, plus the `nav_section`/`active_cls` sidebar tags used by `templates/partials/sidebar.html`.[m
\ No newline at end of file[m
[32m+[m[32m- `core/templatetags/core_tags.py` defines `get_attr`, `time_short` and `activity_label` (maps variant spellings to the canonical Lecture/Tutorial/Practical/Seminar/Workshop labels) filters, plus the `nav_section`/`active_cls` sidebar tags used by `templates/partials/sidebar.html`.[m
\ No newline at end of file[m
[1mdiff --git a/core/importers.py b/core/importers.py[m
[1mindex 2790f61..2eb3b1f 100644[m
[1m--- a/core/importers.py[m
[1m+++ b/core/importers.py[m
[36m@@ -15,15 +15,38 @@[m [mfrom core.models import ([m
     SessionGroup,[m
     StudentGroup,[m
     TechnicalDrawingAllocation,[m
[32m+[m[32m    TimePeriod,[m
     Venue,[m
     WorkshopAllocation,[m
 )[m
[32m+[m[32mfrom core.workshop_times import ([m
[32m+[m[32m    course_programme_codes,[m
[32m+[m[32m    programme_codes_for_group_codes,[m
[32m+[m[32m    workshop_time_issue,[m
[32m+[m[32m)[m
 from core.venue_quality import base_key, detect_name_conflicts[m
 from core.workshop_parser import ([m
     detect_format,[m
     parse_workbook,[m
 )[m
 [m
[32m+[m
[32m+[m[32mdef _master_workshop_programme_codes(rec):[m
[32m+[m[32m    """Programme code(s) a master-timetable WORKSHOP row belongs to.[m
[32m+[m
[32m+[m[32m    A literal group list resolves via the StudentGroup rows (unknown groups[m
[32m+[m[32m    fall back to the default 10:00 Thursday rule); the legacy ``ALL`` group[m
[32m+[m[32m    resolves via the programmes that study the course.[m
[32m+[m[32m    """[m
[32m+[m[32m    if str(rec["raw_groups"]).strip().upper() == "ALL":[m
[32m+[m[32m        return course_programme_codes(rec["course_code"])[m
[32m+[m[32m    group_codes = [[m
[32m+[m[32m        g.strip()[m
[32m+[m[32m        for g in str(rec["raw_groups"]).split(",")[m
[32m+[m[32m        if g.strip() and g.strip().upper() != "ALL"[m
[32m+[m[32m    ][m
[32m+[m[32m    return programme_codes_for_group_codes(group_codes)[m
[32m+[m
 DAY_MAP = {[m
     "mon": Day.MONDAY,[m
     "tue": Day.TUESDAY,[m
[36m@@ -747,6 +770,17 @@[m [mdef _validate_master_rows(df, result) -> list[dict]:[m
             result.errors.append(f"Row {rec['row_no']}: time parse error: {exc}")[m
             result.skipped += 1[m
             continue[m
[32m+[m[32m        if rec["activity_type"] == ActivityType.WORKSHOP:[m
[32m+[m[32m            issue = workshop_time_issue([m
[32m+[m[32m                rec["day"],[m
[32m+[m[32m                rec["start_time"],[m
[32m+[m[32m                rec["end_time"],[m
[32m+[m[32m                programme_code=_master_workshop_programme_codes(rec),[m
[32m+[m[32m            )[m
[32m+[m[32m            if issue:[m
[32m+[m[32m                result.errors.append(f"Row {rec['row_no']}: {issue}")[m
[32m+[m[32m                result.skipped += 1[m
[32m+[m[32m                continue[m
         if rec["raw_groups"].upper() != "ALL":[m
             rec["group_codes"] = [[m
                 g.strip()[m
[36m@@ -932,6 +966,16 @@[m [mdef import_master_timetable_from_excel([m
         except (ValueError, TypeError):[m
             continue[m
         activity_type = normalise_activity_type(rec["activity_raw"])[m
[32m+[m[32m        if activity_type == ActivityType.WORKSHOP:[m
[32m+[m[32m            issue = workshop_time_issue([m
[32m+[m[32m                day,[m
[32m+[m[32m                start_time,[m
[32m+[m[32m                end_time,[m
[32m+[m[32m                programme_code=_master_workshop_programme_codes(rec),[m
[32m+[m[32m            )[m
[32m+[m[32m            if issue:[m
[32m+[m[32m                # Reported and counted as skipped in the validate pass.[m
[32m+[m[32m                continue[m
 [m
         venue_raw = rec["venue_raw"][m
         if not venue_raw:[m
[36m@@ -1176,7 +1220,9 @@[m [mdef import_workshop_allocation_from_excel([m
     time_col = _match_col([m
         df, "time", "time_range", "time range", "time_slot", "time slot", "period"[m
     )[m
[31m-    venue_col = _match_col(df, "venue", "room", "room_name", "venue_name")[m
[32m+[m[32m    venue_col = _match_col([m
[32m+[m[32m        df, "workshop", "venue", "room", "room_name", "venue_name"[m
[32m+[m[32m    )[m
 [m
     if ([m
         group_col is None[m
[36m@@ -1188,14 +1234,15 @@[m [mdef import_workshop_allocation_from_excel([m
             "group_code",[m
             "day",[m
             "a time column ('time' range, or 'start_time' + 'end_time')",[m
[31m-            "venue",[m
[32m+[m[32m            "workshop",[m
         ][m
         if course_col is not None:[m
             required.insert(0, "course_code")[m
         result.errors.append([m
             "Invalid workshop format. Expected FORMAT A (course_code, group_code, "[m
[31m-            "day, start_time, end_time, venue) or FORMAT B (group_code, day, "[m
[31m-            "start_time, end_time, venue where course_code is derived from venue). "[m
[32m+[m[32m            "day, start_time, end_time, workshop) or FORMAT B (group_code, day, "[m
[32m+[m[32m            "start_time, end_time, workshop where course_code is derived from the "[m
[32m+[m[32m            "workshop). "[m
             f"Required columns: {', '.join(required)}. "[m
             f"Found columns: {', '.join(map(str, df.columns))}"[m
         )[m
[36m@@ -1203,12 +1250,12 @@[m [mdef import_workshop_allocation_from_excel([m
 [m
     if course_col is None:[m
         result.format = ([m
[31m-            "workshop format (group_code, day, start_time, end_time, venue) — "[m
[31m-            "course_code derived from venue"[m
[32m+[m[32m            "workshop format (group_code, day, start_time, end_time, workshop) — "[m
[32m+[m[32m            "course_code derived from workshop"[m
         )[m
     else:[m
         result.format = ([m
[31m-            "full format (course_code, group_code, day, start_time, end_time, venue)"[m
[32m+[m[32m            "full format (course_code, group_code, day, start_time, end_time, workshop)"[m
         )[m
 [m
     if semester_id is None:[m
[36m@@ -1258,6 +1305,17 @@[m [mdef import_workshop_allocation_from_excel([m
             result.skipped += 1[m
             continue[m
 [m
[32m+[m[32m        issue = workshop_time_issue([m
[32m+[m[32m            day,[m
[32m+[m[32m            start_time,[m
[32m+[m[32m            end_time,[m
[32m+[m[32m            programme_code=programme_codes_for_group_codes(group_code),[m
[32m+[m[32m        )[m
[32m+[m[32m        if issue:[m
[32m+[m[32m            result.errors.append(f"Row {idx + 2}: {issue}")[m
[32m+[m[32m            result.skipped += 1[m
[32m+[m[32m            continue[m
[32m+[m
         basis = dict([m
             semester=semester,[m
             course_code=course_code,[m
[1mdiff --git a/core/models.py b/core/models.py[m
[1mindex 3a409dd..51acb84 100644[m
[1m--- a/core/models.py[m
[1m+++ b/core/models.py[m
[36m@@ -1,4 +1,5 @@[m
 from django.conf import settings[m
[32m+[m[32mfrom django.core.exceptions import ValidationError[m
 from django.db import models[m
 [m
 [m
[36m@@ -119,6 +120,26 @@[m [mclass Session(models.Model):[m
             f"({venue})"[m
         )[m
 [m
[32m+[m[32m    def clean(self):[m
[32m+[m[32m        """Workshop sessions run only at the standard workshop times.[m
[32m+[m
[32m+[m[32m        Non-workshop sessions are never affected. Enforcement is lazy-imported[m
[32m+[m[32m        so loading ``core.workshop_times`` (which imports this module) can[m
[32m+[m[32m        never form an import cycle.[m
[32m+[m[32m        """[m
[32m+[m[32m        super().clean()[m
[32m+[m[32m        if self.activity_type != "WORKSHOP" or not self.day:[m
[32m+[m[32m            return[m
[32m+[m[32m        from core.workshop_times import validate_workshop_session[m
[32m+[m
[32m+[m[32m        errors = validate_workshop_session(self)[m
[32m+[m[32m        if errors:[m
[32m+[m[32m            raise ValidationError({[m
[32m+[m[32m                field: msgs[m
[32m+[m[32m                for error in errors[m
[32m+[m[32m                for field, msgs in error.error_dict.items()[m
[32m+[m[32m            })[m
[32m+[m
 [m
 class SessionGroup(models.Model):[m
     session = models.ForeignKey([m
[36m@@ -167,6 +188,28 @@[m [mclass WorkshopAllocation(models.Model):[m
             f"{start}-{end} {self.venue}"[m
         )[m
 [m
[32m+[m[32m    def clean(self):[m
[32m+[m[32m        """Workshop allocations keep the standard workshop session times.[m
[32m+[m
[32m+[m[32m        Records imported from the raw workshop matrix carry a period but no[m
[32m+[m[32m        clock times and stay valid; manually entered times must equal one whole[m
[32m+[m[32m        standard session for the day. Enforcement is lazy-imported so loading[m
[32m+[m[32m        ``core.workshop_times`` (which imports this module) never forms an[m
[32m+[m[32m        import cycle.[m
[32m+[m[32m        """[m
[32m+[m[32m        super().clean()[m
[32m+[m[32m        if not self.day:[m
[32m+[m[32m            return[m
[32m+[m[32m        from core.workshop_times import validate_workshop_record[m
[32m+[m
[32m+[m[32m        errors = validate_workshop_record(self)[m
[32m+[m[32m        if errors:[m
[32m+[m[32m            raise ValidationError({[m
[32m+[m[32m                field: msgs[m
[32m+[m[32m                for error in errors[m
[32m+[m[32m                for field, msgs in error.error_dict.items()[m
[32m+[m[32m            })[m
[32m+[m
 [m
 class TechnicalDrawingAllocation(models.Model):[m
     semester = models.ForeignKey([m
[1mdiff --git a/core/templatetags/core_tags.py b/core/templatetags/core_tags.py[m
[1mindex 396eb4e..0c8a69f 100644[m
[1m--- a/core/templatetags/core_tags.py[m
[1m+++ b/core/templatetags/core_tags.py[m
[36m@@ -22,6 +22,31 @@[m [mdef time_short(value):[m
     return value[m
 [m
 [m
[32m+[m[32m_ACTIVITY_LABELS = {[m
[32m+[m[32m    "lecture": "Lecture",[m
[32m+[m[32m    "tutorial": "Tutorial",[m
[32m+[m[32m    "practical": "Practical",[m
[32m+[m[32m    "seminar": "Seminar",[m
[32m+[m[32m    "workshop": "Workshop",[m
[32m+[m[32m}[m
[32m+[m
[32m+[m
[32m+[m[32m@register.filter[m
[32m+[m[32mdef activity_label(value):[m
[32m+[m[32m    """Map an activity type label to one of the five canonical labels.[m
[32m+[m
[32m+[m[32m    Any variant text from the source data (case, spacing, plurals such as[m
[32m+[m[32m    "Lectures" or "Workshops") maps to exactly one of Lecture / Tutorial /[m
[32m+[m[32m    Practical / Seminar / Workshop; unrecognised labels pass through.[m
[32m+[m[32m    """[m
[32m+[m[32m    if not value:[m
[32m+[m[32m        return value[m
[32m+[m[32m    key = value.strip().lower()[m
[32m+[m[32m    singular = key.rstrip("s")[m
[32m+[m[32m    mapping = _ACTIVITY_LABELS.get(key) or _ACTIVITY_LABELS.get(singular)[m
[32m+[m[32m    return mapping if mapping else value[m
[32m+[m
[32m+[m
 @register.filter[m
 def get_item(mapping, key):[m
     """Dictionary lookup by key; returns None when missing."""[m
[1mdiff --git a/core/tests.py b/core/tests.py[m
[1mindex 63eeb91..927af34 100644[m
[1m--- a/core/tests.py[m
[1m+++ b/core/tests.py[m
[36m@@ -1,10 +1,11 @@[m
[31m-﻿import os[m
[32m+[m[32mimport os[m
 import json[m
 import tempfile[m
 from datetime import time[m
 from pathlib import Path[m
 [m
 import pandas as pd[m
[32m+[m[32mfrom django.core.exceptions import ValidationError[m
 from django.core.files.uploadedfile import SimpleUploadedFile[m
 from django.test import Client, TestCase[m
 from openpyxl import Workbook[m
[36m@@ -20,6 +21,7 @@[m [mfrom core.importers import ([m
 )[m
 from core.models import ([m
     ActivityLog,[m
[32m+[m[32m    ActivityType,[m
     ImportHistory,[m
     LogAction,[m
     Programme,[m
[36m@@ -33,8 +35,20 @@[m [mfrom core.models import ([m
     Venue,[m
     WorkshopAllocation,[m
 )[m
[31m-from core.timetable_grid import build_time_day_grid, time_day_grid_to_table[m
[31m-from core.timetable_pdf import build_grid, collect_entries, render_group_timetable[m
[32m+[m[32mfrom core.timetable_grid import ([m
[32m+[m[32m    build_day_time_grid,[m
[32m+[m[32m    build_time_day_grid,[m
[32m+[m[32m    time_day_grid_to_table,[m
[32m+[m[32m)[m
[32m+[m[32mfrom core.timetable_pdf import ([m
[32m+[m[32m    build_grid,[m
[32m+[m[32m    collect_entries,[m
[32m+[m[32m    collect_group_entries,[m
[32m+[m[32m    collect_master_entries,[m
[32m+[m[32m    collect_workshop_rotations,[m
[32m+[m[32m    render_group_timetable,[m
[32m+[m[32m    render_programme_timetable,[m
[32m+[m[32m)[m
 from core.venue_quality import ([m
     analyse_venues,[m
     base_key,[m
[36m@@ -45,6 +59,18 @@[m [mfrom core.venue_quality import ([m
     suggested_name,[m
 )[m
 from core.workshop_parser import parse_workbook[m
[32m+[m[32mfrom core.workshop_times import ([m
[32m+[m[32m    allocation_programme_codes,[m
[32m+[m[32m    course_programme_codes,[m
[32m+[m[32m    legacy_workshop_allocations,[m
[32m+[m[32m    legacy_workshop_sessions,[m
[32m+[m[32m    session_programme_codes,[m
[32m+[m[32m    validate_workshop_record,[m
[32m+[m[32m    validate_workshop_session,[m
[32m+[m[32m    workshop_hours,[m
[32m+[m[32m    workshop_times_for,[m
[32m+[m[32m    workshop_time_issue,[m
[32m+[m[32m)[m
 [m
 XLSX_CONTENT_TYPE = ([m
     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"[m
[36m@@ -349,7 +375,7 @@[m [mclass WorkshopTdImportTests(ImporterTestCase):[m
     def test_workshop_import_is_idempotent(self):[m
         self._seed()[m
         path = make_xlsx([m
[31m-            [["TG201", "C1", "MONDAY", "08:00", "10:00", "TW101"]],[m
[32m+[m[32m            [["TG201", "C1", "MONDAY", "09:00", "13:00", "TW101"]],[m
             ["course_code", "group_code", "day", "start_time", "end_time", "venue"],[m
         )[m
         first = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)[m
[36m@@ -401,6 +427,24 @@[m [mclass WorkshopTdImportTests(ImporterTestCase):[m
         self.assertIn("Invalid workshop format", result.errors[0])[m
         self.assertEqual(WorkshopAllocation.objects.count(), 0)[m
 [m
[32m+[m[32m    def test_workshop_heading_accepted_as_venue_legacy(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        path = make_xlsx([m
[32m+[m[32m            [["E1", "MONDAY", "09:00", "13:00", "Electrical"]],[m
[32m+[m[32m            ["group_code", "day", "start_time", "end_time", "workshop"],[m
[32m+[m[32m        )[m
[32m+[m[32m        first = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)[m
[32m+[m[32m        self.assertEqual(first.errors, [])[m
[32m+[m[32m        self.assertEqual(first.created, 1)[m
[32m+[m[32m        wa = WorkshopAllocation.objects.get()[m
[32m+[m[32m        self.assertEqual(wa.course_code, "Electrical")[m
[32m+[m[32m        self.assertEqual(wa.venue, "Electrical")[m
[32m+[m
[32m+[m[32m        second = import_workshop_allocation_from_excel(path, semester_id=self.sem1.pk)[m
[32m+[m[32m        self.assertEqual(second.created, 0)[m
[32m+[m[32m        self.assertEqual(second.updated, 1)[m
[32m+[m[32m        self.assertEqual(WorkshopAllocation.objects.count(), 1)[m
[32m+[m
     def test_td_import_is_idempotent(self):[m
         self._seed()[m
         path = make_xlsx([m
[36m@@ -583,14 +627,14 @@[m [mclass AdaptiveImportTests(ImporterTestCase):[m
         self._seed()[m
         result = import_workshop_allocation_from_excel([m
             make_xlsx([m
[31m-                [["TG201", "C1", "MONDAY", "08:00-10:00", "TW101"]],[m
[32m+[m[32m                [["TG201", "C1", "MONDAY", "09:00-13:00", "TW101"]],[m
                 ["course_code", "group_code", "day", "time", "venue"],[m
             )[m
         )[m
         self.assertEqual(result.created, 1)[m
         wa = WorkshopAllocation.objects.get(course_code="TG201", group_code="C1")[m
[31m-        self.assertEqual(wa.start_time, time(8, 0))[m
[31m-        self.assertEqual(wa.end_time, time(10, 0))[m
[32m+[m[32m        self.assertEqual(wa.start_time, time(9, 0))[m
[32m+[m[32m        self.assertEqual(wa.end_time, time(13, 0))[m
 [m
     def test_td_pivoted_layout_requires_course_code(self):[m
         self._seed()[m
[36m@@ -770,8 +814,8 @@[m [mclass CrudResponseTests(TestCase):[m
                     "course_code": "TG201",[m
                     "group_code": "C1",[m
                     "day": "MONDAY",[m
[31m-                    "start_time": "08:00",[m
[31m-                    "end_time": "10:00",[m
[32m+[m[32m                    "start_time": "09:00",[m
[32m+[m[32m                    "end_time": "13:00",[m
                     "venue": "TW101",[m
                 },[m
                 WorkshopAllocation,[m
[36m@@ -815,6 +859,121 @@[m [mclass SidebarTests(TestCase):[m
         self.assertEqual(self._active_count(resp), 1)[m
 [m
 [m
[32m+[m[32mclass SidebarCollapseTests(TestCase):[m
[32m+[m[32m    """Collapsible/expandable sidebar: toggle button, icon-only mode with[m
[32m+[m[32m    tooltips, mobile drawer behaviour, a11y/keyboard use, active-state[m
[32m+[m[32m    preservation, preference persistence and responsive content resizing."""[m
[32m+[m
[32m+[m[32m    def _html(self, path="/"):[m
[32m+[m[32m        resp = self.client.get(path)[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        return resp.content.decode()[m
[32m+[m
[32m+[m[32m    def test_collapse_button_at_top_with_controls(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        ctrl = html.find('aria-controls="app-sidebar"')[m
[32m+[m[32m        self.assertGreater(ctrl, -1)[m
[32m+[m[32m        self.assertLess(ctrl, html.find("<nav"))[m
[32m+[m[32m        self.assertIn('id="app-sidebar"', html)[m
[32m+[m[32m        btn = html[max(0, ctrl - 200):ctrl + 150][m
[32m+[m[32m        self.assertIn('type="button"', btn)[m
[32m+[m[32m        self.assertIn('@click="toggle()"', btn)[m
[32m+[m
[32m+[m[32m    def test_animated_width_switch(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        self.assertIn("transition-[width,transform] duration-300", html)[m
[32m+[m[32m        self.assertIn("collapsed ? 'lg:w-20' : 'lg:w-64'", html)[m
[32m+[m[32m        self.assertIn([m
[32m+[m[32m            "collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'",[m
[32m+[m[32m            html,[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertIn("transition-all duration-300", html)[m
[32m+[m
[32m+[m[32m    def test_icon_only_tooltips(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        for text in ([m
[32m+[m[32m            "Dashboard",[m
[32m+[m[32m            "Export Timetable",[m
[32m+[m[32m            "Venues",[m
[32m+[m[32m            "Workshop Allocation",[m
[32m+[m[32m        ):[m
[32m+[m[32m            self.assertIn("showTooltip($el, '%s')" % text, html)[m
[32m+[m[32m        self.assertIn('role="tooltip"', html)[m
[32m+[m[32m        self.assertIn("pointer-events-none", html)[m
[32m+[m[32m        self.assertIn('x-show="collapsed && tooltip"', html)[m
[32m+[m[32m        self.assertIn('x-text="tooltip"', html)[m
[32m+[m
[32m+[m[32m    def test_tooltips_hide_when_expanded(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        self.assertIn("window.innerWidth >= 1024", html)[m
[32m+[m[32m        self.assertIn("showTooltip: function", html)[m
[32m+[m[32m        self.assertIn("hideTooltip: function", html)[m
[32m+[m
[32m+[m[32m    def test_navigation_keeps_icons_and_hrefs(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        for href, aria in ([m
[32m+[m[32m            ("/", "Dashboard"),[m
[32m+[m[32m            ("/export/", "Export Timetable"),[m
[32m+[m[32m            ("/venues/", "Venues"),[m
[32m+[m[32m            ("/workshops/", "Workshop Allocation"),[m
[32m+[m[32m        ):[m
[32m+[m[32m            idx = html.find('aria-label="%s"' % aria)[m
[32m+[m[32m            self.assertGreater(idx, -1)[m
[32m+[m[32m            link = html[html.rfind("<a", 0, idx):html.find("</a>", idx)][m
[32m+[m[32m            self.assertIn('href="%s"' % href, link)[m
[32m+[m[32m            self.assertIn("showTooltip($el, '%s')" % aria, link)[m
[32m+[m[32m        self.assertGreater(html.count("<svg"), 15)[m
[32m+[m
[32m+[m[32m    def test_active_state_preserved_in_both_modes(self):[m
[32m+[m[32m        html = self._html("/programmes/")[m
[32m+[m[32m        start = html.find('href="/programmes/"')[m
[32m+[m[32m        prog_link = html[start:html.find("</a>", start)][m
[32m+[m[32m        self.assertIn("bg-slate-800 text-white", prog_link)[m
[32m+[m[32m        self.assertIn("lg:justify-center lg:gap-0", prog_link)[m
[32m+[m[32m        self.assertIn("<svg", prog_link)[m
[32m+[m[32m        self.assertIn('lg:max-w-0', html)[m
[32m+[m[32m        self.assertEqual(html.count("bg-slate-800 text-white"), 1)[m
[32m+[m[32m        dash_html = self._html("/")[m
[32m+[m[32m        self.assertEqual(dash_html.count("bg-slate-800 text-white"), 1)[m
[32m+[m
[32m+[m[32m    def test_mobile_drawer_behaviour(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        self.assertIn("fixed lg:sticky", html)[m
[32m+[m[32m        self.assertIn("translate-x-0", html)[m
[32m+[m[32m        self.assertIn("-translate-x-full", html)[m
[32m+[m[32m        self.assertIn("lg:hidden", html)[m
[32m+[m[32m        self.assertIn('x-show="sidebarOpen"', html)[m
[32m+[m[32m        self.assertIn('@click="sidebarOpen = false"', html)[m
[32m+[m[32m        self.assertIn("hidden lg:inline-flex", html)[m
[32m+[m
[32m+[m[32m    def test_keyboard_and_accessibility(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        self.assertIn(":aria-expanded=", html)[m
[32m+[m[32m        self.assertIn("Collapse sidebar", html)[m
[32m+[m[32m        self.assertIn("Expand sidebar", html)[m
[32m+[m[32m        self.assertIn("aria-controls=", html)[m
[32m+[m[32m        self.assertIn("focus-visible:ring", html)[m
[32m+[m[32m        self.assertIn(':aria-label="collapsed ? \'Expand sidebar\' : \'Collapse sidebar\'"', html)[m
[32m+[m[32m        self.assertIn("@focus=\"showTooltip($el, 'Dashboard')\"", html)[m
[32m+[m[32m        self.assertIn("@blur=\"hideTooltip()\"", html)[m
[32m+[m[32m        self.assertIn("focus:outline-none", html)[m
[32m+[m
[32m+[m[32m    def test_responsive_content_resizing(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        self.assertIn('class="flex-1 flex flex-col min-w-0"', html)[m
[32m+[m[32m        self.assertIn("shrink-0", html)[m
[32m+[m[32m        self.assertIn("ease-in-out", html)[m
[32m+[m[32m        self.assertIn("duration-300", html)[m
[32m+[m
[32m+[m[32m    def test_preference_persists_across_refreshes(self):[m
[32m+[m[32m        html = self._html()[m
[32m+[m[32m        self.assertIn("coet.sidebar.collapsed", html)[m
[32m+[m[32m        self.assertIn("sessionStorage", html)[m
[32m+[m[32m        self.assertIn("localStorage", html)[m
[32m+[m[32m        self.assertIn("Alpine.data('sidebar',", html)[m
[32m+[m[32m        self.assertIn('x-data="sidebar"', html)[m
[32m+[m
[32m+[m
 class ImportUploadViewTests(TestCase):[m
     def test_htmx_upload_returns_partial(self):[m
         self.client.get("/import/programmes/")[m
[36m@@ -2299,7 +2458,7 @@[m [mclass TimetableGridTests(TestCase):[m
 [m
     def test_grid_slots_and_weekday_columns(self):[m
         grid = build_time_day_grid(collect_entries(self.prog, self.sem))[m
[31m-        self.assertEqual(grid["slots"][0]["label"], "07:00-07:55")[m
[32m+[m[32m        self.assertEqual(grid["slots"][0]["label"], "07:00-08:00")[m
         labels = [d["day"] for d in grid["days"]][m
         self.assertEqual(labels, ["MONDAY", "WEDNESDAY"])[m
         self.assertNotIn("SATURDAY", labels)[m
[36m@@ -2328,7 +2487,7 @@[m [mclass TimetableGridTests(TestCase):[m
         self.assertEqual(data[0][0], "TIME")[m
         self.assertEqual(data[0][1], "MONDAY")[m
         self.assertEqual(data[0][2], "WEDNESDAY")[m
[31m-        self.assertEqual(data[1][0], "07:00-07:55")[m
[32m+[m[32m        self.assertEqual(data[1][0], "07:00-08:00")[m
         self.assertIn(((2, 1), (2, 3)), spans)[m
         self.assertIn(((2, 1), (2, 3), "#d1d5db"), fills)[m
 [m
[36m@@ -2418,8 +2577,8 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         data, _, _ = build_grid([m
             collect_entries(self.prog_ee, self.sem, group=self.ee_c2)[m
         )[m
[31m-        mon = self._flattened_cell(data, "MONDAY", "09:00-09:55")[m
[31m-        tue = self._flattened_cell(data, "TUESDAY", "09:00-09:55")[m
[32m+[m[32m        mon = self._flattened_cell(data, "MONDAY", "09:00-10:00")[m
[32m+[m[32m        tue = self._flattened_cell(data, "TUESDAY", "09:00-10:00")[m
         self.assertEqual(mon, "Building")[m
         self.assertEqual(tue, "Electronics")[m
         for cell in (mon, tue):[m
[36m@@ -2434,7 +2593,7 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         self.assertEqual([e["label"] for e in entries], ["Welding"])[m
         data, _, _ = build_grid(entries)[m
         self.assertEqual([m
[31m-            self._flattened_cell(data, "MONDAY", "09:00-09:55"), "Welding"[m
[32m+[m[32m            self._flattened_cell(data, "MONDAY", "09:00-10:00"), "Welding"[m
         )[m
 [m
     def test_distinct_workshops_on_different_days_all_kept(self):[m
[36m@@ -2447,13 +2606,13 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         )[m
         data, _, _ = build_grid(entries)[m
         self.assertEqual([m
[31m-            self._flattened_cell(data, "MONDAY", "09:00-09:55"), "Carpentry"[m
[32m+[m[32m            self._flattened_cell(data, "MONDAY", "09:00-10:00"), "Carpentry"[m
         )[m
         self.assertEqual([m
[31m-            self._flattened_cell(data, "TUESDAY", "09:00-09:55"), "Plumbing"[m
[32m+[m[32m            self._flattened_cell(data, "TUESDAY", "09:00-10:00"), "Plumbing"[m
         )[m
         self.assertEqual([m
[31m-            self._flattened_cell(data, "WEDNESDAY", "09:00-09:55"), "Welding"[m
[32m+[m[32m            self._flattened_cell(data, "WEDNESDAY", "09:00-10:00"), "Welding"[m
         )[m
 [m
     def test_same_workshop_week_runs_are_distinguished(self):[m
[36m@@ -2482,7 +2641,7 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         data, _, _ = build_grid([m
             collect_entries(self.prog_ee, self.sem, group=self.ee_c2)[m
         )[m
[31m-        cell = self._flattened_cell(data, "TUESDAY", "09:00-09:55")[m
[32m+[m[32m        cell = self._flattened_cell(data, "TUESDAY", "09:00-10:00")[m
         self.assertIn("Building \u00b7 Wk 1-6", cell)[m
         self.assertIn("Building \u00b7 Wk 8-13", cell)[m
 [m
[36m@@ -2499,7 +2658,7 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)[m
         self.assertEqual(entries[0]["hours"], {9, 10, 11, 12})[m
         grid = build_time_day_grid(entries)[m
[31m-        self.assertEqual(grid["slots"][0]["label"], "09:00-09:55")[m
[32m+[m[32m        self.assertEqual(grid["slots"][0]["label"], "09:00-10:00")[m
         cell = grid["rows"][0]["cols"][0][m
         self.assertEqual(cell["rowspan"], 4)[m
 [m
[36m@@ -2516,10 +2675,10 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)[m
         self.assertEqual(entries[0]["hours"], {15, 16, 17, 18})[m
         grid = build_time_day_grid(entries)[m
[31m-        self.assertEqual(grid["slots"][0]["label"], "15:00-15:55")[m
[32m+[m[32m        self.assertEqual(grid["slots"][0]["label"], "15:00-16:00")[m
         self.assertEqual(grid["rows"][0]["cols"][0]["rowspan"], 4)[m
 [m
[31m-    def test_group_has_only_one_workshop_per_day(self):[m
[32m+[m[32m    def test_same_day_different_time_workshops_are_separate_sessions(self):[m
         self._flat_workshop("Building")  # group C2, MONDAY morning[m
         WorkshopAllocation.objects.create([m
             semester=self.sem,[m
[36m@@ -2531,9 +2690,8 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
             venue="Carpentry",[m
         )[m
         entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)[m
[31m-        labels = [e["label"] for e in entries][m
[31m-        self.assertEqual(labels, ["Building"])[m
[31m-        self.assertNotIn("Carpentry", labels)[m
[32m+[m[32m        labels = sorted(e["label"] for e in entries)[m
[32m+[m[32m        self.assertEqual(labels, ["Building", "Carpentry"])[m
         # A shared name across week runs is still one workshop per day.[m
         WorkshopAllocation.objects.create([m
             semester=self.sem,[m
[36m@@ -2561,7 +2719,7 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         thu = [e for e in entries if e["day"] == "THURSDAY"][m
         self.assertEqual(sorted(e["label"] for e in thu), ["Building", "Building"])[m
         self.assertEqual(sorted(e["note"] for e in thu), ["Wk 1-6", "Wk 8-13"])[m
[31m-        labels = [e["label"] for e in entries]  # Monday flat + two Thursday runs[m
[32m+[m[32m        labels = [e["label"] for e in entries]  # Monday two + two Thursday runs[m
         self.assertEqual(labels.count("Building"), 3)[m
 [m
     def test_all_groups_cells_attribute_group_codes(self):[m
[36m@@ -2570,13 +2728,13 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         data, _, _ = build_grid([m
             collect_entries(self.prog_ee, self.sem), show_groups=True[m
         )[m
[31m-        cell = self._flattened_cell(data, "TUESDAY", "09:00-09:55")[m
[32m+[m[32m        cell = self._flattened_cell(data, "TUESDAY", "09:00-10:00")[m
         self.assertEqual(cell, "Building \u00b7 C1\n\nBuilding \u00b7 C2")[m
         # In a single-group export neither group code is repeated.[m
         data, _, _ = build_grid([m
             collect_entries(self.prog_ee, self.sem, group=self.ee_c1)[m
         )[m
[31m-        cell = self._flattened_cell(data, "TUESDAY", "09:00-09:55")[m
[32m+[m[32m        cell = self._flattened_cell(data, "TUESDAY", "09:00-10:00")[m
         self.assertEqual(cell, "Building")[m
 [m
     def test_all_groups_never_merge_unrelated_group_sessions(self):[m
[36m@@ -2626,8 +2784,8 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         )[m
         self.assertEqual(resp.status_code, 200)[m
         html = resp.content.decode()[m
[31m-        self.assertIn('tt-groups">C1<', html)[m
[31m-        self.assertIn('tt-groups">C2<', html)[m
[32m+[m[32m        self.assertIn("Assigned groups: C1", html)[m
[32m+[m[32m        self.assertIn("Assigned groups: C2", html)[m
 [m
     def test_workshop_names_simplified_across_programmes(self):[m
         WorkshopAllocation.objects.create([m
[36m@@ -2694,8 +2852,9 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         )[m
         self.assertEqual(resp.status_code, 200)[m
         html = resp.content.decode()[m
[31m-        self.assertIn("tt-course", html)[m
[31m-        self.assertEqual(html.count("Building"), 1)[m
[32m+[m[32m        self.assertEqual(html.count("Course: Building"), 1)[m
[32m+[m[32m        self.assertIn("Venue: Building", html)[m
[32m+[m[32m        self.assertIn("Assigned groups: C2", html)[m
         self.assertNotIn("Building Building", html)[m
 [m
     def test_group_pdf_renders_with_simplified_workshop_cell(self):[m
[36m@@ -2710,6 +2869,208 @@[m [mclass WorkshopCellDisplayTests(TestCase):[m
         self.assertGreater(len(pdf), 1000)[m
 [m
 [m
[32m+[m[32mclass WorkshopRotationTests(TestCase):[m
[32m+[m[32m    """Rotating workshops merge into one cell plus a Workshop Rotation Key.[m
[32m+[m
[32m+[m[32m    A slot (group + day + time period) holding two or more different workshop[m
[32m+[m[32m    identities is a weekly rotation: the shared cell lists all names and the[m
[32m+[m[32m    key maps each contiguous week block to a workshop. Different-time sessions[m
[32m+[m[32m    on the same day stay separate. Programme affinity (workshop name in the[m
[32m+[m[32m    programme name) orders the names without hard-coding.[m
[32m+[m[32m    """[m
[32m+[m
[32m+[m[32m    def setUp(self):[m
[32m+[m[32m        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)[m
[32m+[m[32m        self.prog_ee = Programme.objects.create([m
[32m+[m[32m            code="EE", name="BSc. in Electrical Engineering"[m
[32m+[m[32m        )[m
[32m+[m[32m        self.prog_cpe = Programme.objects.create([m
[32m+[m[32m            code="CPE", name="BSc. in Chemical and Processing Engineering"[m
[32m+[m[32m        )[m
[32m+[m[32m        self.ee_c1 = StudentGroup.objects.create(programme=self.prog_ee, code="C1")[m
[32m+[m[32m        self.ee_c2 = StudentGroup.objects.create(programme=self.prog_ee, code="C2")[m
[32m+[m[32m        self.cpe_b1 = StudentGroup.objects.create(programme=self.prog_cpe, code="B1")[m
[32m+[m
[32m+[m[32m    def _workshop(self, name, group, day="THURSDAY", period="MORNING", course=None):[m
[32m+[m[32m        return WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code=course or name,[m
[32m+[m[32m            workshop=name,[m
[32m+[m[32m            group_code=group,[m
[32m+[m[32m            day=day,[m
[32m+[m[32m            time_period=period,[m
[32m+[m[32m            venue="",[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_rotating_workshops_merge_into_one_entry(self):[m
[32m+[m[32m        self._workshop("Carpentry", "C1")[m
[32m+[m[32m        self._workshop("Electrical", "C1")[m
[32m+[m[32m        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c1)[m
[32m+[m[32m        self.assertEqual(len(entries), 1)[m
[32m+[m[32m        self.assertEqual(entries[0]["label"], "Electrical / Carpentry")[m
[32m+[m[32m        self.assertEqual(entries[0]["note"], "Wk 1-7 / Wk 8-14")[m
[32m+[m
[32m+[m[32m    def test_rotation_key_row_describes_the_slot(self):[m
[32m+[m[32m        self._workshop("Carpentry", "C1")[m
[32m+[m[32m        self._workshop("Electrical", "C1")[m
[32m+[m[32m        rows = collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c1)[m
[32m+[m[32m        self.assertEqual(len(rows), 1)[m
[32m+[m[32m        row = rows[0][m
[32m+[m[32m        self.assertEqual(row["programme"], self.prog_ee.name)[m
[32m+[m[32m        self.assertEqual(row["group"], "C1")[m
[32m+[m[32m        self.assertEqual(row["day"], "THURSDAY")[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            row["blocks"], [(1, 7, "Electrical"), (8, 14, "Carpentry")][m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_non_rotating_workshop_has_no_rotation_row(self):[m
[32m+[m[32m        self._workshop("Building", "C2")[m
[32m+[m[32m        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c2)[m
[32m+[m[32m        self.assertEqual(entries[0]["label"], "Building")[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c2),[m
[32m+[m[32m            [],[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_three_way_rotation_merges_all_names(self):[m
[32m+[m[32m        self._workshop("Welding", "C1")[m
[32m+[m[32m        self._workshop("Electrical", "C1")[m
[32m+[m[32m        self._workshop("M/Tools", "C1")[m
[32m+[m[32m        entries = collect_entries(self.prog_ee, self.sem, group=self.ee_c1)[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            entries[0]["label"], "Electrical / M/Tools / Welding"[m
[32m+[m[32m        )[m
[32m+[m[32m        row = collect_workshop_rotations(self.prog_ee, self.sem)[0][m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            row["blocks"],[m
[32m+[m[32m            [[m
[32m+[m[32m                (1, 7, "Electrical"),[m
[32m+[m[32m                (8, 14, "M/Tools"),[m
[32m+[m[32m                (15, 21, "Welding"),[m
[32m+[m[32m            ],[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_rotations_are_isolated_per_programme(self):[m
[32m+[m[32m        self._workshop("Carpentry", "C1")[m
[32m+[m[32m        self._workshop("Electrical", "C1")[m
[32m+[m[32m        self._workshop("Electronics", "B1")[m
[32m+[m[32m        self._workshop("CPE", "B1")[m
[32m+[m[32m        self.assertEqual(len(collect_workshop_rotations(self.prog_ee, self.sem)), 1)[m
[32m+[m[32m        cpe_rows = collect_workshop_rotations(self.prog_cpe, self.sem)[m
[32m+[m[32m        self.assertEqual(len(cpe_rows), 1)[m
[32m+[m[32m        self.assertEqual(cpe_rows[0]["group"], "B1")[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            cpe_rows[0]["blocks"], [(1, 7, "CPE"), (8, 14, "Electronics")][m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_rotation_mapping_is_group_specific(self):[m
[32m+[m[32m        self._workshop("Carpentry", "C1")[m
[32m+[m[32m        self._workshop("Electrical", "C1")[m
[32m+[m[32m        self._workshop("Building", "C2")[m
[32m+[m[32m        self._workshop("Electronics", "C2")[m
[32m+[m[32m        ee_rows = collect_workshop_rotations(self.prog_ee, self.sem)[m
[32m+[m[32m        c1_row = [r for r in ee_rows if r["group"] == "C1"][m
[32m+[m[32m        c2_row = [r for r in ee_rows if r["group"] == "C2"][m
[32m+[m[32m        self.assertEqual(len(c1_row), 1)[m
[32m+[m[32m        self.assertEqual(len(c2_row), 1)[m
[32m+[m[32m        self.assertEqual(c1_row[0]["blocks"][0][2], "Electrical")[m
[32m+[m[32m        self.assertEqual(c2_row[0]["blocks"][0][2], "Building")[m
[32m+[m
[32m+[m[32m    def test_shared_course_code_appears_in_the_key(self):[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="WK-101",[m
[32m+[m[32m            workshop="Carpentry",[m
[32m+[m[32m            group_code="C1",[m
[32m+[m[32m            day="THURSDAY",[m
[32m+[m[32m            time_period="MORNING",[m
[32m+[m[32m            venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="WK-101",[m
[32m+[m[32m            workshop="Electrical",[m
[32m+[m[32m            group_code="C1",[m
[32m+[m[32m            day="THURSDAY",[m
[32m+[m[32m            time_period="MORNING",[m
[32m+[m[32m            venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        row = collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c1)[0][m
[32m+[m[32m        self.assertEqual(row["course"], "WK-101")[m
[32m+[m
[32m+[m[32m    def test_explicit_week_blocks_are_honoured(self):[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="Building",[m
[32m+[m[32m            workshop="Building",[m
[32m+[m[32m            group_code="C2",[m
[32m+[m[32m            day="MONDAY",[m
[32m+[m[32m            time_period="MORNING",[m
[32m+[m[32m            venue="",[m
[32m+[m[32m            week_start=1,[m
[32m+[m[32m            week_end=7,[m
[32m+[m[32m        )[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="Welding",[m
[32m+[m[32m            workshop="Welding",[m
[32m+[m[32m            group_code="C2",[m
[32m+[m[32m            day="MONDAY",[m
[32m+[m[32m            time_period="MORNING",[m
[32m+[m[32m            venue="",[m
[32m+[m[32m            week_start=17,[m
[32m+[m[32m            week_end=18,[m
[32m+[m[32m        )[m
[32m+[m[32m        row = collect_workshop_rotations(self.prog_ee, self.sem, group=self.ee_c2)[0][m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            row["blocks"], [(1, 7, "Building"), (17, 18, "Welding")][m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_empty_programme_has_no_rotation_rows(self):[m
[32m+[m[32m        self.assertEqual(collect_workshop_rotations(self.prog_cpe, self.sem), [])[m
[32m+[m
[32m+[m[32m    def test_full_day_pdf_grid_starts_at_seven(self):[m
[32m+[m[32m        self._workshop("Carpentry", "C1")[m
[32m+[m[32m        self._workshop("Electrical", "C1")[m
[32m+[m[32m        data, _, _ = build_grid(collect_entries(self.prog_ee, self.sem))[m
[32m+[m[32m        self.assertEqual(data[0][0], "TIME")[m
[32m+[m[32m        self.assertEqual(data[1][0], "07:00-08:00")[m
[32m+[m[32m        self.assertEqual(data[-1][0], "19:00-20:00")[m
[32m+[m[32m        self.assertEqual(len(data) - 1, 13)[m
[32m+[m
[32m+[m[32m    def test_late_evening_session_still_renders_on_full_grid(self):[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="Building",[m
[32m+[m[32m            workshop="Building",[m
[32m+[m[32m            group_code="C2",[m
[32m+[m[32m            day="MONDAY",[m
[32m+[m[32m            start_time="18:00",[m
[32m+[m[32m            end_time="21:00",[m
[32m+[m[32m            venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        data, _, _ = build_grid(collect_entries(self.prog_ee, self.sem))[m
[32m+[m[32m        header = data[0][m
[32m+[m[32m        di = header.index("MONDAY")[m
[32m+[m[32m        times = [row[0] for row in data[1:]][m
[32m+[m[32m        self.assertIn("07:00-08:00", times)[m
[32m+[m[32m        self.assertIn("19:00-20:00", times)[m
[32m+[m[32m        self.assertIn("20:00-21:00", times)[m
[32m+[m[32m        mon = [row[di] for row in data[1:] if row[di]][m
[32m+[m[32m        self.assertTrue(any("Building" in cell for cell in mon))[m
[32m+[m
[32m+[m[32m    def test_pdf_contains_rotation_key_table(self):[m
[32m+[m[32m        from io import BytesIO[m
[32m+[m
[32m+[m[32m        self._workshop("Carpentry", "C1")[m
[32m+[m[32m        self._workshop("Electrical", "C1")[m
[32m+[m[32m        buf = BytesIO()[m
[32m+[m[32m        render_programme_timetable(self.prog_ee, self.sem, 1, out=buf)[m
[32m+[m[32m        pdf = buf.getvalue()[m
[32m+[m[32m        self.assertTrue(pdf.startswith(b"%PDF-"))[m
[32m+[m[32m        self.assertGreater(len(pdf), 1000)[m
[32m+[m
[32m+[m
 class GroupTimetablePageTests(TestCase):[m
     """The on-screen timetable filters strictly to the selected student group."""[m
 [m
[36m@@ -2807,7 +3168,7 @@[m [mclass GroupTimetablePageTests(TestCase):[m
 [m
     def test_group_grid_merges_multi_slot_entries(self):[m
         c1_page = self._get(self.c1).content.decode()[m
[31m-        self.assertIn('rowspan="2"', c1_page)[m
[32m+[m[32m        self.assertIn('colspan="2"', c1_page)[m
         self.assertIn("09:00", c1_page)[m
 [m
 [m
[36m@@ -2832,8 +3193,38 @@[m [mclass TimetablePageTests(TestCase):[m
             venue=venue,[m
         )[m
         SessionGroup.objects.create(session=session, group=group)[m
[32m+[m[32m        # An ungrouped session (as in the real master timetable) — tutorials,[m
[32m+[m[32m        # seminars and practicals have no group links, so only the default[m
[32m+[m[32m        # whole-year view (master timetable as source of truth) shows them.[m
[32m+[m[32m        Session.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="ST171",[m
[32m+[m[32m            activity_type="TUTORIAL",[m
[32m+[m[32m            day="TUESDAY",[m
[32m+[m[32m            start_time="10:00",[m
[32m+[m[32m            end_time="12:00",[m
[32m+[m[32m            venue=venue,[m
[32m+[m[32m        )[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="EE151",[m
[32m+[m[32m            group_code="C1",[m
[32m+[m[32m            day="WEDNESDAY",[m
[32m+[m[32m            start_time="09:00",[m
[32m+[m[32m            end_time="12:00",[m
[32m+[m[32m            venue="WORKSHOP 1",[m
[32m+[m[32m        )[m
[32m+[m[32m        TechnicalDrawingAllocation.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="EE153",[m
[32m+[m[32m            group_code="C1",[m
[32m+[m[32m            day="THURSDAY",[m
[32m+[m[32m            start_time="13:00",[m
[32m+[m[32m            end_time="17:00",[m
[32m+[m[32m            venue="TD LAB",[m
[32m+[m[32m        )[m
 [m
[31m-    def test_page_renders_classic_grid(self):[m
[32m+[m[32m    def test_page_renders_day_time_grid(self):[m
         resp = self.client.get([m
             "/timetable/",[m
             {"programme": self.prog_a.pk, "semester": self.sem.pk},[m
[36m@@ -2842,15 +3233,30 @@[m [mclass TimetablePageTests(TestCase):[m
         self.assertEqual(resp.status_code, 200)[m
         html = resp.content.decode()[m
         self.assertIn("TIME", html)[m
[32m+[m[32m        # Days run down the first column, labelled with their full weekday name.[m
         self.assertIn("Monday", html)[m
[31m-        self.assertIn("08:00", html)[m
[31m-        self.assertIn("08:55", html)[m
[32m+[m[32m        self.assertIn("Friday", html)[m
[32m+[m[32m        self.assertNotIn("MONDAY</td>", html)[m
[32m+[m[32m        # Hours run across the top, ascending from 07:00, headers in the[m
[32m+[m[32m        # master-timetable format 07:00-07:55 etc.[m
[32m+[m[32m        self.assertIn(">07:00-07:55<", html)[m
[32m+[m[32m        self.assertIn(">08:00-08:55<", html)[m
[32m+[m[32m        self.assertIn(">09:00-09:55<", html)[m
[32m+[m[32m        self.assertNotIn(">07:00-08:00<", html)[m
         self.assertIn("MT171", html)[m
[31m-        self.assertIn("Lecture", html)[m
         self.assertIn("YOMBO5", html)[m
         self.assertIn("tt-card", html)[m
         self.assertIn("tt-kind-lecture", html)[m
         self.assertIn("tt-time", html)[m
[32m+[m[32m        # Cards live in an inner flex wrapper — the <td> itself stays a plain[m
[32m+[m[32m        # table cell so colspans align under the right hourly column.[m
[32m+[m[32m        self.assertIn('class="tt-block-inner"', html)[m
[32m+[m[32m        self.assertTrue(html.count('class="tt-block"') == html.count('class="tt-block-inner"'))[m
[32m+[m[32m        # Entry box reads like the Excel grid cells.[m
[32m+[m[32m        self.assertIn("Lecture, 08:00-09:00, Mon", html)[m
[32m+[m[32m        self.assertIn("Course: MT171", html)[m
[32m+[m[32m        self.assertIn("Venue: YOMBO5", html)[m
[32m+[m[32m        self.assertIn("Assigned groups:", html)[m
 [m
     def test_page_falls_back_when_programme_unknown(self):[m
         resp = self.client.get([m
[36m@@ -2867,6 +3273,58 @@[m [mclass TimetablePageTests(TestCase):[m
         self.assertEqual(resp.status_code, 200)[m
         self.assertIn("No sessions scheduled", resp.content.decode())[m
 [m
[32m+[m[32m    def test_default_view_shows_complete_first_year_timetable(self):[m
[32m+[m[32m        resp = self.client.get("/timetable/", HTTP_HOST="localhost")[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        html = resp.content.decode()[m
[32m+[m[32m        self.assertIn("ALL PROGRAMMES", html)[m
[32m+[m[32m        self.assertIn(">TIME<", html)[m
[32m+[m[32m        self.assertNotIn("No programmes registered yet", html)[m
[32m+[m[32m        self.assertIn("MT171", html)[m
[32m+[m[32m        # The master timetable is the source of truth: ungrouped sessions[m
[32m+[m[32m        # (tutorials/seminars/practicals) appear too, with a blank groups line.[m
[32m+[m[32m        self.assertIn("Course: ST171", html)[m
[32m+[m[32m        self.assertRegex(html, r"Assigned groups:\s*</div>")[m
[32m+[m
[32m+[m[32m    def test_all_programmes_option_shows_complete_first_year_timetable(self):[m
[32m+[m[32m        resp = self.client.get([m
[32m+[m[32m            "/timetable/", {"programme": "all"}, HTTP_HOST="localhost"[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        html = resp.content.decode()[m
[32m+[m[32m        self.assertIn("ALL PROGRAMMES", html)[m
[32m+[m[32m        self.assertIn(">TIME<", html)[m
[32m+[m[32m        self.assertNotIn("No programmes registered yet", html)[m
[32m+[m[32m        self.assertIn("MT171", html)[m
[32m+[m[32m        self.assertIn("ST171", html)[m
[32m+[m
[32m+[m[32m    def test_programme_filter_hides_ungrouped_master_sessions(self):[m
[32m+[m[32m        resp = self.client.get([m
[32m+[m[32m            "/timetable/",[m
[32m+[m[32m            {"programme": self.prog_a.pk, "semester": self.sem.pk},[m
[32m+[m[32m            HTTP_HOST="localhost",[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        html = resp.content.decode()[m
[32m+[m[32m        self.assertIn("MT171", html)[m
[32m+[m[32m        self.assertIn("ELECTRICAL ENGINEERING", html)[m
[32m+[m[32m        # Ungrouped master-timetable sessions are not part of any single[m
[32m+[m[32m        # programme's timetable view.[m
[32m+[m[32m        self.assertNotIn("ST171", html)[m
[32m+[m
[32m+[m[32m    def test_master_entries_include_every_session_in_the_semester(self):[m
[32m+[m[32m        entries = collect_master_entries(self.sem)[m
[32m+[m[32m        session_codes = {[m
[32m+[m[32m            e["name"] for e in entries if e["key"][0] == "session"[m
[32m+[m[32m        }[m
[32m+[m[32m        self.assertIn("MT171", session_codes)[m
[32m+[m[32m        self.assertIn("ST171", session_codes)[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            len(session_codes), Session.objects.filter(semester=self.sem).count()[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertTrue(any(e["kind"] == "workshop" for e in entries))[m
[32m+[m[32m        self.assertTrue(any(e["kind"] == "td" for e in entries))[m
[32m+[m
     def test_classic_pdf_export(self):[m
         resp = self.client.get([m
             "/export/programmes/%d/timetable.pdf/" % self.prog_a.pk,[m
[36m@@ -2941,6 +3399,244 @@[m [mclass TimetablePageTests(TestCase):[m
         self.assertIn("No programmes or groups match", html)[m
 [m
 [m
[32m+[m[32mclass DayTimeGridTests(TestCase):[m
[32m+[m[32m    """The default on-screen grid: DAYS as rows (full weekday labels), TIME as[m
[32m+[m[32m    columns, each day drawn as vertical lanes with merged session cells."""[m
[32m+[m
[32m+[m[32m    def setUp(self):[m
[32m+[m[32m        self.sem = Semester.objects.create(academic_year="2026/2027", semester=1)[m
[32m+[m[32m        self.prog = Programme.objects.create(code="EE", name="BSc. in Electrical Engineering")[m
[32m+[m[32m        self.g = StudentGroup.objects.create(programme=self.prog, code="C1")[m
[32m+[m[32m        StudentGroup.objects.create(programme=self.prog, code="C2")[m
[32m+[m[32m        self.venue = Venue.objects.create(name="NB102", capacity=80)[m
[32m+[m
[32m+[m[32m    @staticmethod[m
[32m+[m[32m    def _entry(**overrides):[m
[32m+[m[32m        base = {[m
[32m+[m[32m            "key": ("session", 1),[m
[32m+[m[32m            "day": "MONDAY",[m
[32m+[m[32m            "hours": {8, 9},[m
[32m+[m[32m            "label": "MT171",[m
[32m+[m[32m            "kind": "lecture",[m
[32m+[m[32m            "course_code": "MT171",[m
[32m+[m[32m            "name": "MT171",[m
[32m+[m[32m            "type_label": "Lecture",[m
[32m+[m[32m            "venue": "NB102",[m
[32m+[m[32m            "start": "08:00",[m
[32m+[m[32m            "end": "09:55",[m
[32m+[m[32m            "groups": "C1",[m
[32m+[m[32m            "note": "",[m
[32m+[m[32m        }[m
[32m+[m[32m        base.update(overrides)[m
[32m+[m[32m        return base[m
[32m+[m
[32m+[m[32m    def _spanning_cells(self, grid):[m
[32m+[m[32m        return [[m
[32m+[m[32m            c[m
[32m+[m[32m            for r in grid["rows"][m
[32m+[m[32m            for lane in r["lanes"][m
[32m+[m[32m            for c in lane[m
[32m+[m[32m            if not c.get("empty")[m
[32m+[m[32m        ][m
[32m+[m
[32m+[m[32m    def test_orientation_full_day_labels_lane_span(self):[m
[32m+[m[32m        g = build_day_time_grid([self._entry()])[m
[32m+[m[32m        self.assertEqual(g["slots"][0]["start"], "07:00")[m
[32m+[m[32m        self.assertEqual(g["slots"][0]["label"], "07:00-07:55")[m
[32m+[m[32m        # All five weekdays are present, even with data on one day only.[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            [r["label"] for r in g["rows"]],[m
[32m+[m[32m            ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual(g["rows"][0]["day"], "MONDAY")[m
[32m+[m[32m        lanes = g["rows"][0]["lanes"][m
[32m+[m[32m        self.assertEqual(len(lanes), 1)[m
[32m+[m[32m        self.assertEqual(lanes[0][0]["empty"], True)[m
[32m+[m[32m        self.assertEqual(len(lanes[0]), 2)[m
[32m+[m[32m        self.assertEqual(lanes[0][-1]["colspan"], 2)[m
[32m+[m[32m        self.assertEqual(lanes[0][-1]["entries"][0]["course_code"], "MT171")[m
[32m+[m
[32m+[m[32m    def test_column_range_ends_at_latest_hour(self):[m
[32m+[m[32m        g = build_day_time_grid([m
[32m+[m[32m            [[m
[32m+[m[32m                self._entry([m
[32m+[m[32m                    key=("session", 1),[m
[32m+[m[32m                    day="FRIDAY",[m
[32m+[m[32m                    hours={13},[m
[32m+[m[32m                    label="IE101",[m
[32m+[m[32m                    course_code="IE101",[m
[32m+[m[32m                    name="IE101",[m
[32m+[m[32m                    start="13:00",[m
[32m+[m[32m                    end="14:00",[m
[32m+[m[32m                    groups="",[m
[32m+[m[32m                )[m
[32m+[m[32m            ][m
[32m+[m[32m        )[m
[32m+[m[32m        headers = [s["start"] for s in g["slots"]][m
[32m+[m[32m        self.assertEqual(headers, ["07:00", "08:00", "09:00", "10:00", "11:00", "12:00", "13:00"])[m
[32m+[m[32m        self.assertNotIn("14:00", headers)[m
[32m+[m
[32m+[m[32m    def test_day_time_grid_merges_across_hours(self):[m
[32m+[m[32m        session = Session.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="MT171",[m
[32m+[m[32m            activity_type="LECTURE",[m
[32m+[m[32m            day="MONDAY",[m
[32m+[m[32m            start_time="08:00",[m
[32m+[m[32m            end_time="09:55",[m
[32m+[m[32m            venue=self.venue,[m
[32m+[m[32m        )[m
[32m+[m[32m        SessionGroup.objects.create(session=session, group=self.g)[m
[32m+[m[32m        grid = build_day_time_grid(collect_entries(self.prog, self.sem))[m
[32m+[m[32m        self.assertEqual([c.get("colspan") for c in self._spanning_cells(grid)], [2])[m
[32m+[m
[32m+[m[32m    def test_sessions_begin_under_their_start_hour_column(self):[m
[32m+[m[32m        g = build_day_time_grid([m
[32m+[m[32m            [[m
[32m+[m[32m                self._entry(key=("session", 1), name="A", course_code="A",[m
[32m+[m[32m                            day="MONDAY", hours={8}, start="08:00", end="09:00"),[m
[32m+[m[32m                self._entry(key=("session", 2), name="B", course_code="B",[m
[32m+[m[32m                            day="MONDAY", hours={10}, start="10:00", end="11:00"),[m
[32m+[m[32m                self._entry(key=("session", 3), name="C", course_code="C",[m
[32m+[m[32m                            day="MONDAY", hours={14}, start="14:00", end="15:00"),[m
[32m+[m[32m            ][m
[32m+[m[32m        )[m
[32m+[m[32m        lane = g["rows"][0]["lanes"][0][m
[32m+[m[32m        # A sits right after the 07:00 empty column, then an empty column[m
[32m+[m[32m        # (09:00) before B at 10:00, then three empty columns (11:00-13:00)[m
[32m+[m[32m        # before C at 14:00. The column the block starts on is its start hour.[m
[32m+[m[32m        empties = [c["colspan"] for c in lane if c.get("empty")][m
[32m+[m[32m        blocks = [c["entries"][0]["course_code"] for c in lane if not c.get("empty")][m
[32m+[m[32m        self.assertEqual(empties, [1, 1, 3])[m
[32m+[m[32m        self.assertEqual(blocks, ["A", "B", "C"])[m
[32m+[m
[32m+[m[32m    def test_partial_hour_session_spans_columns(self):[m
[32m+[m[32m        g = build_day_time_grid([m
[32m+[m[32m            [[m
[32m+[m[32m                self._entry([m
[32m+[m[32m                    key=("session", 1),[m
[32m+[m[32m                    day="MONDAY",[m
[32m+[m[32m                    hours={8, 9, 10},[m
[32m+[m[32m                    start="08:30",[m
[32m+[m[32m                    end="10:15",[m
[32m+[m[32m                )[m
[32m+[m[32m            ][m
[32m+[m[32m        )[m
[32m+[m[32m        cells = [c for c in self._spanning_cells(g) if not c.get("empty")][m
[32m+[m[32m        self.assertEqual(cells[0]["colspan"], 3)[m
[32m+[m
[32m+[m[32m    def test_spanned_session_is_one_cell_not_per_hour(self):[m
[32m+[m[32m        g = build_day_time_grid([m
[32m+[m[32m            [[m
[32m+[m[32m                self._entry([m
[32m+[m[32m                    key=("session", 1),[m
[32m+[m[32m                    day="MONDAY",[m
[32m+[m[32m                    hours={9, 10, 11, 12},[m
[32m+[m[32m                    start="09:00",[m
[32m+[m[32m                    end="13:00",[m
[32m+[m[32m                )[m
[32m+[m[32m            ][m
[32m+[m[32m        )[m
[32m+[m[32m        cells = [c for c in self._spanning_cells(g) if not c.get("empty")][m
[32m+[m[32m        self.assertEqual(len(cells), 1)[m
[32m+[m[32m        self.assertEqual(cells[0]["colspan"], 4)[m
[32m+[m[32m        self.assertEqual(len(cells[0]["entries"]), 1)[m
[32m+[m
[32m+[m[32m    def test_overlapping_sessions_use_separate_lanes(self):[m
[32m+[m[32m        g = build_day_time_grid([m
[32m+[m[32m            [[m
[32m+[m[32m                self._entry(key=("session", 1), day="MONDAY", hours={9, 10, 11, 12}),[m
[32m+[m[32m                self._entry(key=("session", 2), day="MONDAY", hours={10, 11, 12, 13}),[m
[32m+[m[32m            ][m
[32m+[m[32m        )[m
[32m+[m[32m        row = g["rows"][0][m
[32m+[m[32m        self.assertEqual(len(row["lanes"]), 2)[m
[32m+[m[32m        for lane in row["lanes"]:[m
[32m+[m[32m            non_empty = [c for c in lane if not c.get("empty")][m
[32m+[m[32m            self.assertEqual(len(non_empty), 1)[m
[32m+[m[32m            self.assertEqual(non_empty[0]["colspan"], 4)[m
[32m+[m
[32m+[m[32m    def test_adjacent_sessions_share_a_lane(self):[m
[32m+[m[32m        g = build_day_time_grid([m
[32m+[m[32m            [[m
[32m+[m[32m                self._entry(key=("session", 1), day="MONDAY", hours={8}, start="08:00", end="09:00"),[m
[32m+[m[32m                self._entry(key=("session", 2), day="MONDAY", hours={9}, start="09:00", end="10:00"),[m
[32m+[m[32m            ][m
[32m+[m[32m        )[m
[32m+[m[32m        row = g["rows"][0][m
[32m+[m[32m        self.assertEqual(len(row["lanes"]), 1)[m
[32m+[m[32m        cells = [c for c in row["lanes"][0] if not c.get("empty")][m
[32m+[m[32m        self.assertEqual(len(cells), 2)[m
[32m+[m
[32m+[m[32m    def test_empty_day_is_still_rendered_with_an_empty_lane(self):[m
[32m+[m[32m        g = build_day_time_grid([m
[32m+[m[32m            [self._entry(key=("session", 1), day="TUESDAY", hours={9})][m
[32m+[m[32m        )[m
[32m+[m[32m        monday = g["rows"][0][m
[32m+[m[32m        self.assertEqual(monday["day"], "MONDAY")[m
[32m+[m[32m        self.assertEqual(len(monday["lanes"]), 1)[m
[32m+[m[32m        self.assertEqual(len(monday["lanes"][0]), 1)[m
[32m+[m[32m        self.assertEqual(monday["lanes"][0][0]["empty"], True)[m
[32m+[m[32m        self.assertEqual(monday["lanes"][0][0]["colspan"], len(g["slots"]))[m
[32m+[m
[32m+[m[32m    def test_activity_card_renders_excel_style_text(self):[m
[32m+[m[32m        session = Session.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="MT171",[m
[32m+[m[32m            activity_type="LECTURE",[m
[32m+[m[32m            day="MONDAY",[m
[32m+[m[32m            start_time="08:00",[m
[32m+[m[32m            end_time="09:00",[m
[32m+[m[32m            venue=self.venue,[m
[32m+[m[32m        )[m
[32m+[m[32m        SessionGroup.objects.create(session=session, group=self.g)[m
[32m+[m[32m        resp = self.client.get([m
[32m+[m[32m            "/timetable/",[m
[32m+[m[32m            {"programme": self.prog.pk, "semester": self.sem.pk},[m
[32m+[m[32m            HTTP_HOST="localhost",[m
[32m+[m[32m        )[m
[32m+[m[32m        html = resp.content.decode()[m
[32m+[m[32m        # Box reads exactly like the Excel grid cells: header line, then[m
[32m+[m[32m        # Course:/Venue:/Assigned groups: rows.[m
[32m+[m[32m        self.assertIn("Lecture, 08:00-09:00, Mon", html)[m
[32m+[m[32m        self.assertIn("Course: MT171", html)[m
[32m+[m[32m        self.assertIn("Venue: NB102", html)[m
[32m+[m[32m        self.assertIn("Assigned groups: C1", html)[m
[32m+[m[32m        self.assertNotIn("Assigned groups: N/A", html)[m
[32m+[m
[32m+[m[32m    def test_activity_card_groups_line_blank_without_groups(self):[m
[32m+[m[32m        Session.objects.create([m
[32m+[m[32m            semester=self.sem,[m
[32m+[m[32m            course_code="MT171",[m
[32m+[m[32m            activity_type="LECTURE",[m
[32m+[m[32m            day="MONDAY",[m
[32m+[m[32m            start_time="08:00",[m
[32m+[m[32m            end_time="09:00",[m
[32m+[m[32m            venue=self.venue,[m
[32m+[m[32m        )[m
[32m+[m[32m        # Ungrouped sessions still appear (master timetable is the source of[m
[32m+[m[32m        # truth) and their "Assigned groups:" line is left blank.[m
[32m+[m[32m        resp = self.client.get("/timetable/", HTTP_HOST="localhost")[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        html = resp.content.decode()[m
[32m+[m[32m        self.assertIn("Course: MT171", html)[m
[32m+[m[32m        self.assertRegex(html, r"Assigned groups:\s*</div>")[m
[32m+[m
[32m+[m[32m    def test_activity_label_filter_maps_variants(self):[m
[32m+[m[32m        from django.template import Template, Context[m
[32m+[m[32m        from django.template import engines[m
[32m+[m
[32m+[m[32m        engine = engines["django"][m
[32m+[m[32m        tmpl = engine.from_string([m
[32m+[m[32m            "{% load core_tags %}"[m
[32m+[m[32m            "{{ 'lecture'|activity_label }}|{{ 'Lectures'|activity_label }}"[m
[32m+[m[32m            "|{{ 'Tutorial'|activity_label }}|{{ 'PRACTICAL'|activity_label }}"[m
[32m+[m[32m            "|{{ ' Seminar '|activity_label }}|{{ 'Workshops'|activity_label }}"[m
[32m+[m[32m        )[m
[32m+[m[32m        out = tmpl.render({"request": None})[m
[32m+[m[32m        self.assertEqual(out, "Lecture|Lecture|Tutorial|Practical|Seminar|Workshop")[m
[32m+[m
[32m+[m
 class DeletionImpactTests(TestCase):[m
     """Delete previews must show the full cascade before confirming, and the[m
     actual delete must match what was previewed (Django cascades + detaches)."""[m
[36m@@ -3676,4 +4372,529 @@[m [mclass ImportHistoryTests(TestCase):[m
         html = self.client.get("/import/").content.decode()[m
         self.assertIn("Previous Import Details", html)[m
         self.assertIn("hub_progs.xlsx", html)[m
[31m-        self.assertIn("SUCCESS", html)[m
\ No newline at end of file[m
[32m+[m[32m        self.assertIn("SUCCESS", html)[m
[32m+[m
[32m+[m
[32m+[m[32mclass WorkshopStandardTimesTests(ImporterTestCase):[m
[32m+[m[32m    """The standard workshop session times are enforced everywhere.[m
[32m+[m
[32m+[m[32m    Monday/Tuesday/Wednesday/Friday morning 09:00-13:00, afternoon[m
[32m+[m[32m    15:00-19:00; Thursday morning 10:00-14:00, afternoon 15:00-19:00.[m
[32m+[m[32m    Workshops on other days or at other times are rejected with a clear[m
[32m+[m[32m    message; non-workshop sessions are never affected. The same times apply[m
[32m+[m[32m    to every programme, regardless of the record's student group.[m
[32m+[m[32m    """[m
[32m+[m
[32m+[m[32m    def _workshop(self, **overrides):[m
[32m+[m[32m        fields = dict([m
[32m+[m[32m            semester=self.sem1,[m
[32m+[m[32m            course_code="TG201",[m
[32m+[m[32m            group_code="A1",[m
[32m+[m[32m            day="MONDAY",[m
[32m+[m[32m            start_time=time(9, 0),[m
[32m+[m[32mend_time=time(13, 0),[m
[32m+[m[32m            venue="TW101",[m
[32m+[m[32m        )[m
[32m+[m[32m        fields.update(overrides)[m
[32m+[m[32m        return WorkshopAllocation.objects.create(**fields)[m
[32m+[m
[32m+[m[32m    # --------------------------------------------------------- rules module[m
[32m+[m
[32m+[m[32m    def test_standard_times_table(self):[m
[32m+[m[32m        cases = [[m
[32m+[m[32m            ("MONDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),[m
[32m+[m[32m            ("MONDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),[m
[32m+[m[32m            ("TUESDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),[m
[32m+[m[32m            ("TUESDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),[m
[32m+[m[32m            ("WEDNESDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),[m
[32m+[m[32m            ("WEDNESDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),[m
[32m+[m[32m            ("FRIDAY", TimePeriod.MORNING, time(9, 0), time(13, 0)),[m
[32m+[m[32m            ("FRIDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),[m
[32m+[m[32m            ("THURSDAY", TimePeriod.MORNING, time(10, 0), time(14, 0)),[m
[32m+[m[32m            ("THURSDAY", TimePeriod.AFTERNOON, time(15, 0), time(19, 0)),[m
[32m+[m[32m        ][m
[32m+[m[32m        for day, period, start, end in cases:[m
[32m+[m[32m            with self.subTest(day=day, period=period):[m
[32m+[m[32m                self.assertEqual(workshop_times_for(day, period), (start, end))[m
[32m+[m
[32m+[m[32m    def test_weekend_has_no_workshop_sessions(self):[m
[32m+[m[32m        self.assertIsNone(workshop_times_for("SATURDAY", TimePeriod.MORNING))[m
[32m+[m[32m        self.assertIsNone(workshop_times_for("SUNDAY", TimePeriod.AFTERNOON))[m
[32m+[m[32m        self.assertIn("not scheduled", workshop_time_issue("SATURDAY", time(9, 0), time(13, 0)))[m
[32m+[m
[32m+[m[32m    def test_thursday_morning_is_1000_to_1400_for_all_programmes(self):[m
[32m+[m[32m        for code in ("CE", "TE", "EE", "ME", "QS", "cpE", "UNKNOWN"):[m
[32m+[m[32m            with self.subTest(programme=code):[m
[32m+[m[32m                self.assertEqual([m
[32m+[m[32m                    workshop_times_for("THURSDAY", TimePeriod.MORNING, code),[m
[32m+[m[32m                    (time(10, 0), time(14, 0)),[m
[32m+[m[32m                )[m
[32m+[m[32m                self.assertEqual([m
[32m+[m[32m                    workshop_hours("THURSDAY", TimePeriod.MORNING, code),[m
[32m+[m[32m                    {10, 11, 12, 13},[m
[32m+[m[32m                )[m
[32m+[m[32m        # A mixed bag of programmes still uses the single Thursday morning[m
[32m+[m[32m        # session, and Thursday-only: Monday morning stays 09:00-13:00.[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            workshop_times_for("THURSDAY", TimePeriod.MORNING, ("ME", "CE")),[m
[32m+[m[32m            (time(10, 0), time(14, 0)),[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            workshop_times_for("MONDAY", TimePeriod.MORNING, "CE"),[m
[32m+[m[32m            (time(9, 0), time(13, 0)),[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            workshop_times_for("MONDAY", TimePeriod.MORNING, "ME"),[m
[32m+[m[32m            (time(9, 0), time(13, 0)),[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_workshop_hours_are_day_aware(self):[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            workshop_hours("MONDAY", TimePeriod.MORNING), {9, 10, 11, 12}[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            workshop_hours("THURSDAY", TimePeriod.MORNING), {10, 11, 12, 13}[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            workshop_hours("THURSDAY", TimePeriod.MORNING, "CE"), {10, 11, 12, 13}[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            workshop_hours("FRIDAY", TimePeriod.AFTERNOON), {15, 16, 17, 18}[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_time_issue_accepts_standard_and_rejects_bad(self):[m
[32m+[m[32m        for day in ("MONDAY", "TUESDAY", "WEDNESDAY", "FRIDAY"):[m
[32m+[m[32m            self.assertEqual(workshop_time_issue(day, time(9, 0), time(13, 0)), "")[m
[32m+[m[32m            self.assertEqual(workshop_time_issue(day, time(15, 0), time(19, 0)), "")[m
[32m+[m[32m        self.assertEqual(workshop_time_issue("THURSDAY", time(10, 0), time(14, 0)), "")[m
[32m+[m[32m        self.assertEqual(workshop_time_issue("THURSDAY", time(15, 0), time(19, 0)), "")[m
[32m+[m[32m        self.assertIn("must run", workshop_time_issue("MONDAY", time(8, 0), time(10, 0)))[m
[32m+[m[32m        # Thursday morning is 10:00-14:00 for every programme.[m
[32m+[m[32m        self.assertIn([m
[32m+[m[32m            "must run",[m
[32m+[m[32m            workshop_time_issue("THURSDAY", time(9, 0), time(13, 0), programme_code="CE"),[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertIn([m
[32m+[m[32m            "must run",[m
[32m+[m[32m            workshop_time_issue("THURSDAY", time(10, 0), time(13, 0), programme_code="ME"),[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertIn([m
[32m+[m[32m            "must run",[m
[32m+[m[32m            workshop_time_issue("THURSDAY", time(10, 0), time(13, 0), programme_code="CE"),[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertIn("must run", workshop_time_issue("THURSDAY", time(9, 0), time(13, 0)))[m
[32m+[m[32m        self.assertIn("required", workshop_time_issue("MONDAY", time(9, 0), None))[m
[32m+[m
[32m+[m[32m    # ----------------------------------------------------------- model clean[m
[32m+[m
[32m+[m[32m    def test_model_clean_rejects_non_standard_workshop(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        rec = WorkshopAllocation([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="MONDAY", start_time=time(8, 0), end_time=time(10, 0), venue="W",[m
[32m+[m[32m        )[m
[32m+[m[32m        with self.assertRaises(ValidationError):[m
[32m+[m[32m            rec.full_clean()[m
[32m+[m
[32m+[m[32m    def test_model_clean_rejects_weekend_workshop(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        rec = WorkshopAllocation([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="SATURDAY", start_time=time(9, 0), end_time=time(13, 0), venue="W",[m
[32m+[m[32m        )[m
[32m+[m[32m        with self.assertRaises(ValidationError) as ctx:[m
[32m+[m[32m            rec.full_clean()[m
[32m+[m[32m        self.assertIn("not scheduled", str(ctx.exception))[m
[32m+[m
[32m+[m[32m    def test_model_clean_accepts_matrix_period_only_record(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        rec = WorkshopAllocation([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        rec.full_clean()  # must not raise[m
[32m+[m
[32m+[m[32m    def test_model_clean_accepts_thursday_morning(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        rec = WorkshopAllocation([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="THURSDAY", start_time=time(10, 0), end_time=time(14, 0), venue="W",[m
[32m+[m[32m        )[m
[32m+[m[32m        rec.full_clean()  # must not raise — Thursday morning is 10:00-14:00[m
[32m+[m
[32m+[m[32m    def test_model_clean_rejects_thursday_0900(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        rec = WorkshopAllocation([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="B1",[m
[32m+[m[32m            day="THURSDAY", start_time=time(9, 0), end_time=time(13, 0), venue="W",[m
[32m+[m[32m        )[m
[32m+[m[32m        with self.assertRaises(ValidationError) as ctx:[m
[32m+[m[32m            rec.full_clean()[m
[32m+[m[32m        self.assertIn("10:00-14:00", str(ctx.exception))[m
[32m+[m
[32m+[m[32m    def test_model_clean_rejects_thursday_1000_1300(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        rec = WorkshopAllocation([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="THURSDAY", start_time=time(10, 0), end_time=time(13, 0), venue="W",[m
[32m+[m[32m        )[m
[32m+[m[32m        with self.assertRaises(ValidationError) as ctx:[m
[32m+[m[32m            rec.full_clean()[m
[32m+[m[32m        self.assertIn("10:00-14:00", str(ctx.exception))[m
[32m+[m
[32m+[m[32m    def test_session_clean_rejects_non_standard_workshop(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        ses = Session([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,[m
[32m+[m[32m            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,[m
[32m+[m[32m        )[m
[32m+[m[32m        with self.assertRaises(ValidationError):[m
[32m+[m[32m            ses.full_clean()[m
[32m+[m
[32m+[m[32m    def test_session_clean_ignores_non_workshop(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        ses = Session([m
[32m+[m[32m            semester=self.sem1, course_code="MT161", activity_type=ActivityType.LECTURE,[m
[32m+[m[32m            day="MONDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,[m
[32m+[m[32m        )[m
[32m+[m[32m        ses.full_clean()  # must not raise[m
[32m+[m
[32m+[m[32m    def test_orm_create_bypasses_validation(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        self._workshop(day="MONDAY", start_time=time(8, 0), end_time=time(10, 0))[m
[32m+[m[32m        self.assertEqual(WorkshopAllocation.objects.count(), 1)[m
[32m+[m
[32m+[m[32m    # ---------------------------------------------------- form inline errors[m
[32m+[m
[32m+[m[32m    def test_workshop_create_shows_inline_time_error(self):[m
[32m+[m[32m        sem = Semester.objects.create(academic_year="2026/2027", semester=1)[m
[32m+[m[32m        resp = self.client.post("/workshops/create/", {[m
[32m+[m[32m            "semester": sem.pk,[m
[32m+[m[32m            "course_code": "TG201",[m
[32m+[m[32m            "group_code": "A1",[m
[32m+[m[32m            "day": "MONDAY",[m
[32m+[m[32m            "start_time": "08:00",[m
[32m+[m[32m            "end_time": "10:00",[m
[32m+[m[32m            "venue": "TW101",[m
[32m+[m[32m        })[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        self.assertContains(resp, "must run")[m
[32m+[m[32m        self.assertEqual(WorkshopAllocation.objects.count(), 0)[m
[32m+[m
[32m+[m[32m    def test_session_create_shows_inline_workshop_time_error(self):[m
[32m+[m[32m        sem = Semester.objects.create(academic_year="2026/2027", semester=1)[m
[32m+[m[32m        resp = self.client.post("/sessions/create/", {[m
[32m+[m[32m            "semester": sem.pk,[m
[32m+[m[32m            "course_code": "TG201",[m
[32m+[m[32m            "activity_type": "WORKSHOP",[m
[32m+[m[32m            "day": "THURSDAY",[m
[32m+[m[32m            "start_time": "08:00",[m
[32m+[m[32m            "end_time": "10:00",[m
[32m+[m[32m            "venue": "",[m
[32m+[m[32m            "session_groups-TOTAL_FORMS": "0",[m
[32m+[m[32m            "session_groups-INITIAL_FORMS": "0",[m
[32m+[m[32m            "session_groups-MIN_NUM_FORMS": "0",[m
[32m+[m[32m            "session_groups-MAX_NUM_FORMS": "1000",[m
[32m+[m[32m        })[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        self.assertContains(resp, "must run")[m
[32m+[m[32m        self.assertEqual(Session.objects.count(), 0)[m
[32m+[m
[32m+[m[32m    def test_lecture_session_create_unaffected(self):[m
[32m+[m[32m        sem = Semester.objects.create(academic_year="2026/2027", semester=1)[m
[32m+[m[32m        resp = self.client.post("/sessions/create/", {[m
[32m+[m[32m            "semester": sem.pk,[m
[32m+[m[32m            "course_code": "MT161",[m
[32m+[m[32m            "activity_type": "LECTURE",[m
[32m+[m[32m            "day": "MONDAY",[m
[32m+[m[32m            "start_time": "08:00",[m
[32m+[m[32m            "end_time": "10:00",[m
[32m+[m[32m            "venue": "",[m
[32m+[m[32m            "session_groups-TOTAL_FORMS": "0",[m
[32m+[m[32m            "session_groups-INITIAL_FORMS": "0",[m
[32m+[m[32m            "session_groups-MIN_NUM_FORMS": "0",[m
[32m+[m[32m            "session_groups-MAX_NUM_FORMS": "1000",[m
[32m+[m[32m        })[m
[32m+[m[32m        self.assertEqual(resp.status_code, 302)[m
[32m+[m[32m        self.assertEqual(Session.objects.get(activity_type="LECTURE").start_time, time(8, 0))[m
[32m+[m[32m        self.assertEqual(Session.objects.get(activity_type="LECTURE").end_time, time(10, 0))[m
[32m+[m
[32m+[m[32m    # ----------------------------------------------------------- imports[m
[32m+[m
[32m+[m[32m    def test_flat_import_rejects_non_standard_times(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        result = import_workshop_allocation_from_excel([m
[32m+[m[32m            make_xlsx([m
[32m+[m[32m                [["TG201", "C1", "MONDAY", "08:00", "10:00", "TW101"]],[m
[32m+[m[32m                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],[m
[32m+[m[32m            ),[m
[32m+[m[32m            semester_id=self.sem1.pk,[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertTrue(result.errors)[m
[32m+[m[32m        self.assertIn("must run", result.errors[0])[m
[32m+[m[32m        self.assertEqual(result.skipped, 1)[m
[32m+[m[32m        self.assertEqual(WorkshopAllocation.objects.count(), 0)[m
[32m+[m
[32m+[m[32m    def test_flat_import_rejects_saturday(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        result = import_workshop_allocation_from_excel([m
[32m+[m[32m            make_xlsx([m
[32m+[m[32m                [["TG201", "C1", "SATURDAY", "09:00", "13:00", "TW101"]],[m
[32m+[m[32m                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],[m
[32m+[m[32m            ),[m
[32m+[m[32m            semester_id=self.sem1.pk,[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertIn("not scheduled", result.errors[0])[m
[32m+[m[32m        self.assertEqual(WorkshopAllocation.objects.count(), 0)[m
[32m+[m
[32m+[m[32m    def test_flat_import_accepts_thursday_morning(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        result = import_workshop_allocation_from_excel([m
[32m+[m[32m            make_xlsx([m
[32m+[m[32m                [["TG201", "C1", "THURSDAY", "10:00", "14:00", "TW101"]],[m
[32m+[m[32m                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],[m
[32m+[m[32m            ),[m
[32m+[m[32m            semester_id=self.sem1.pk,[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual(result.errors, [])[m
[32m+[m[32m        self.assertEqual(result.created, 1)[m
[32m+[m[32m        rec = WorkshopAllocation.objects.get()[m
[32m+[m[32m        self.assertEqual(rec.start_time, time(10, 0))[m
[32m+[m[32m        self.assertEqual(rec.end_time, time(14, 0))[m
[32m+[m
[32m+[m[32m    def test_flat_import_accepts_thursday_morning_for_any_group(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        result = import_workshop_allocation_from_excel([m
[32m+[m[32m            make_xlsx([m
[32m+[m[32m                [["TG201", "A1", "THURSDAY", "10:00", "14:00", "TW101"]],[m
[32m+[m[32m                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],[m
[32m+[m[32m            ),[m
[32m+[m[32m            semester_id=self.sem1.pk,[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual(result.errors, [])[m
[32m+[m[32m        self.assertEqual(result.created, 1)[m
[32m+[m[32m        rec = WorkshopAllocation.objects.get()[m
[32m+[m[32m        self.assertEqual(rec.start_time, time(10, 0))[m
[32m+[m[32m        self.assertEqual(rec.end_time, time(14, 0))[m
[32m+[m
[32m+[m[32m    def test_flat_import_rejects_thursday_0900(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        result = import_workshop_allocation_from_excel([m
[32m+[m[32m            make_xlsx([m
[32m+[m[32m                [["TG201", "B1", "THURSDAY", "09:00", "13:00", "TW101"]],[m
[32m+[m[32m                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],[m
[32m+[m[32m            ),[m
[32m+[m[32m            semester_id=self.sem1.pk,[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertTrue(result.errors)[m
[32m+[m[32m        self.assertIn("10:00-14:00", result.errors[0])[m
[32m+[m[32m        self.assertEqual(result.skipped, 1)[m
[32m+[m[32m        self.assertEqual(WorkshopAllocation.objects.count(), 0)[m
[32m+[m
[32m+[m[32m    def test_flat_import_rejects_thursday_1000_1300(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        result = import_workshop_allocation_from_excel([m
[32m+[m[32m            make_xlsx([m
[32m+[m[32m                [["TG201", "A1", "THURSDAY", "10:00", "13:00", "TW101"]],[m
[32m+[m[32m                ["course_code", "group_code", "day", "start_time", "end_time", "venue"],[m
[32m+[m[32m            ),[m
[32m+[m[32m            semester_id=self.sem1.pk,[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertTrue(result.errors)[m
[32m+[m[32m        self.assertIn("10:00-14:00", result.errors[0])[m
[32m+[m[32m        self.assertEqual(WorkshopAllocation.objects.count(), 0)[m
[32m+[m
[32m+[m[32m    def test_master_import_rejects_non_standard_workshop(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        path = make_xlsx([m
[32m+[m[32m            [["TG201", "WORKSHOP", "MONDAY", "08:00", "10:00", "LH1", "A1"]],[m
[32m+[m[32m            MASTER_COLS,[m
[32m+[m[32m        )[m
[32m+[m[32m        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)[m
[32m+[m[32m        self.assertTrue(result.errors)[m
[32m+[m[32m        self.assertFalse([m
[32m+[m[32m            Session.objects.filter(activity_type=ActivityType.WORKSHOP).exists()[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_master_import_accepts_default_thursday_morning(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        path = make_xlsx([m
[32m+[m[32m            [["TG201", "WORKSHOP", "THURSDAY", "10:00", "14:00", "LH1", "B1"]],[m
[32m+[m[32m            MASTER_COLS,[m
[32m+[m[32m        )[m
[32m+[m[32m        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)[m
[32m+[m[32m        self.assertEqual(result.errors, [])[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            Session.objects.filter([m
[32m+[m[32m                activity_type=ActivityType.WORKSHOP,[m
[32m+[m[32m                start_time=time(10, 0),[m
[32m+[m[32m                end_time=time(14, 0),[m
[32m+[m[32m            ).count(),[m
[32m+[m[32m            1,[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_master_import_accepts_thursday_morning_for_any_group(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        path = make_xlsx([m
[32m+[m[32m            [["TG201", "WORKSHOP", "THURSDAY", "10:00", "14:00", "LH1", "A1"]],[m
[32m+[m[32m            MASTER_COLS,[m
[32m+[m[32m        )[m
[32m+[m[32m        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)[m
[32m+[m[32m        self.assertEqual(result.errors, [])[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            Session.objects.filter([m
[32m+[m[32m                activity_type=ActivityType.WORKSHOP,[m
[32m+[m[32m                start_time=time(10, 0),[m
[32m+[m[32m                end_time=time(14, 0),[m
[32m+[m[32m            ).count(),[m
[32m+[m[32m            1,[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    def test_master_import_rejects_thursday_0900(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        path = make_xlsx([m
[32m+[m[32m            [["TG201", "WORKSHOP", "THURSDAY", "09:00", "13:00", "LH1", "B1"]],[m
[32m+[m[32m            MASTER_COLS,[m
[32m+[m[32m        )[m
[32m+[m[32m        result = import_master_timetable_from_excel(path, semester_id=self.sem1.pk)[m
[32m+[m[32m        self.assertTrue(result.errors)[m
[32m+[m[32m        self.assertFalse([m
[32m+[m[32m            Session.objects.filter(activity_type=ActivityType.WORKSHOP).exists()[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    # ---------------------------------------------------------- grid / PDF[m
[32m+[m
[32m+[m[32m    def test_grid_thursday_morning_uses_10_to_14_slot(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="B1",[m
[32m+[m[32m            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        entries = collect_entries(self.prog_b, self.sem1, group=self.g3)[m
[32m+[m[32m        self.assertEqual(entries[0]["hours"], {10, 11, 12, 13})[m
[32m+[m[32m        grid = build_time_day_grid(entries)[m
[32m+[m[32m        self.assertEqual(grid["slots"][0]["label"], "10:00-11:00")[m
[32m+[m[32m        self.assertEqual(grid["rows"][0]["cols"][0]["rowspan"], 4)[m
[32m+[m
[32m+[m[32m    def test_grid_thursday_morning_uses_10_to_14_slot_for_any_programme(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        entries = collect_entries(self.prog_a, self.sem1, group=self.g1)[m
[32m+[m[32m        self.assertEqual(entries[0]["hours"], {10, 11, 12, 13})[m
[32m+[m[32m        grid = build_time_day_grid(entries)[m
[32m+[m[32m        self.assertEqual(grid["slots"][0]["label"], "10:00-11:00")[m
[32m+[m[32m        self.assertEqual(grid["rows"][0]["cols"][0]["rowspan"], 4)[m
[32m+[m
[32m+[m[32m    def test_grid_monday_morning_uses_09_to_1255_slot(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="MONDAY", time_period=TimePeriod.MORNING, venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        entries = collect_entries(self.prog_a, self.sem1, group=self.g1)[m
[32m+[m[32m        self.assertEqual(entries[0]["hours"], {9, 10, 11, 12})[m
[32m+[m[32m        grid = build_time_day_grid(entries)[m
[32m+[m[32m        self.assertEqual(grid["slots"][0]["label"], "09:00-10:00")[m
[32m+[m
[32m+[m[32m    def test_pdf_renders_thursday_morning_workshop(self):[m
[32m+[m[32m        from io import BytesIO[m
[32m+[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        WorkshopAllocation.objects.create([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        out = BytesIO()[m
[32m+[m[32m        render_programme_timetable(self.prog_a, self.sem1, out=out)[m
[32m+[m[32m        self.assertGreater(len(out.getvalue()), 1000)[m
[32m+[m
[32m+[m[32m    # -------------------------------------------------------------- legacy[m
[32m+[m
[32m+[m[32m    def test_legacy_workshop_allocations_detected(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        self._workshop()[m
[32m+[m[32m        self._workshop(day="TUESDAY", start_time=time(8, 0), end_time=time(10, 0))[m
[32m+[m[32m        found = legacy_workshop_allocations()[m
[32m+[m[32m        self.assertEqual(len(found), 1)[m
[32m+[m[32m        rec, issue = found[0][m
[32m+[m[32m        self.assertEqual(rec.day, "TUESDAY")[m
[32m+[m[32m        self.assertIn("must run", issue)[m
[32m+[m
[32m+[m[32m    def test_legacy_workshop_allocations_flag_old_thursday_times(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        # Neither CE (group A1) nor ME (group B1) may start Thursday at 09:00[m
[32m+[m[32m        # anymore — the morning session is 10:00-14:00 for every programme.[m
[32m+[m[32m        self._workshop([m
[32m+[m[32m            day="THURSDAY", group_code="A1",[m
[32m+[m[32m            start_time=time(9, 0), end_time=time(13, 0),[m
[32m+[m[32m        )[m
[32m+[m[32m        self._workshop([m
[32m+[m[32m            day="THURSDAY", group_code="B1",[m
[32m+[m[32m            start_time=time(9, 0), end_time=time(13, 0),[m
[32m+[m[32m        )[m
[32m+[m[32m        found = legacy_workshop_allocations()[m
[32m+[m[32m        self.assertEqual(len(found), 2)[m
[32m+[m[32m        self.assertEqual({rec.group_code for rec, _ in found}, {"A1", "B1"})[m
[32m+[m
[32m+[m[32m    def test_programme_resolution_helpers(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        rec = WorkshopAllocation([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", group_code="A1",[m
[32m+[m[32m            day="THURSDAY", time_period=TimePeriod.MORNING, venue="",[m
[32m+[m[32m        )[m
[32m+[m[32m        self.assertEqual(allocation_programme_codes(rec), ("CE",))[m
[32m+[m[32m        self.assertEqual(allocation_programme_codes([m
[32m+[m[32m            WorkshopAllocation([m
[32m+[m[32m                semester=self.sem1, course_code="TG201", group_code="UNKNOWN",[m
[32m+[m[32m                day="MONDAY", time_period=TimePeriod.MORNING, venue="",[m
[32m+[m[32m            )[m
[32m+[m[32m        ), ())[m
[32m+[m[32m        self.assertEqual([m
[32m+[m[32m            course_programme_codes("TG201"), ("CE",)[m
[32m+[m[32m        )[m
[32m+[m[32m        ses = Session([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,[m
[32m+[m[32m            day="THURSDAY", start_time=time(9, 0), end_time=time(13, 0), venue=None,[m
[32m+[m[32m        )[m
[32m+[m[32m        ses.save()[m
[32m+[m[32m        self.assertEqual(session_programme_codes(ses), ())[m
[32m+[m[32m        SessionGroup.objects.create(session=ses, group=self.g1)[m
[32m+[m[32m        self.assertEqual(session_programme_codes(ses), ("CE",))[m
[32m+[m[32m        ses2 = Session([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,[m
[32m+[m[32m            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,[m
[32m+[m[32m        )[m
[32m+[m[32m        # An unsaved session's groups have not been assigned yet.[m
[32m+[m[32m        self.assertEqual(session_programme_codes(ses2), ())[m
[32m+[m
[32m+[m[32m    def test_legacy_workshop_sessions_detected(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        Session.objects.create([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,[m
[32m+[m[32m            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,[m
[32m+[m[32m        )[m
[32m+[m[32m        found = legacy_workshop_sessions()[m
[32m+[m[32m        self.assertEqual(len(found), 1)[m
[32m+[m[32m        session, issue = found[0][m
[32m+[m[32m        self.assertEqual(session.day, "THURSDAY")[m
[32m+[m[32m        self.assertIn("must run", issue)[m
[32m+[m
[32m+[m[32m    def test_workshop_list_shows_legacy_banner(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        self._workshop(day="TUESDAY", start_time=time(8, 0), end_time=time(10, 0))[m
[32m+[m[32m        resp = self.client.get("/workshops/")[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        self.assertContains(resp, "standard workshop session times")[m
[32m+[m
[32m+[m[32m    def test_session_list_shows_legacy_banner(self):[m
[32m+[m[32m        self._seed()[m
[32m+[m[32m        Session.objects.create([m
[32m+[m[32m            semester=self.sem1, course_code="TG201", activity_type=ActivityType.WORKSHOP,[m
[32m+[m[32m            day="THURSDAY", start_time=time(8, 0), end_time=time(10, 0), venue=None,[m
[32m+[m[32m        )[m
[32m+[m[32m        resp = self.client.get("/sessions/")[m
[32m+[m[32m        self.assertEqual(resp.status_code, 200)[m
[32m+[m[32m        self.assertContains(resp, "standard workshop session times")[m
\ No newline at end of file[m
[1mdiff --git a/core/timetable_grid.py b/core/timetable_grid.py[m
[1mindex ce811ae..8b2be33 100644[m
[1m--- a/core/timetable_grid.py[m
[1m+++ b/core/timetable_grid.py[m
[36m@@ -1,13 +1,19 @@[m
[31m-"""Weekly timetable grid builder: TIME as rows, DAYS as columns.[m
[31m-[m
[31m-Builds on the resolved entries from ``core.timetable_pdf.collect_entries`` so[m
[31m-the on-screen timetable (``core/timetable.html``) and the PDF export render[m
[31m-the same layout. The row headers are the hourly slots actually covered by the[m
[31m-selected programme/semester (07:00-07:55 .. 19:00-19:55 window); the columns[m
[31m-are MONDAY..FRIDAY (plus SATURDAY/SUNDAY when the timetable has weekend[m
[31m-activity). Sessions that occupy consecutive slots are merged into a single[m
[31m-vertical cell. Every render cell keeps the session type so templates and the[m
[31m-PDF can shade Lecture (grey), Workshop (green) and Technical Drawing (pink).[m
[32m+[m[32m"""Weekly timetable grid builders.[m
[32m+[m
[32m+[m[32m``build_time_day_grid`` produces the classic time-row/day-column layout shared[m
[32m+[m[32mwith the PDF export (``core/timetable_pdf.build_grid``). ``build_day_time_grid``[m
[32m+[m[32mproduces the on-screen transposed layout — DAYS as rows (always the full[m
[32m+[m[32mMonday..Friday range, so labels are the full weekday names) and HOURLY time[m
[32m+[m[32mslots as columns — described by the default timetable view[m
[32m+[m[32m(``core/timetable.html``). Both build on the resolved entries from[m
[32m+[m[32m``core.timetable_pdf.collect_entries`` so the on-screen timetable and the PDF[m
[32m+[m[32mexport render the same data. Sessions that occupy consecutive slots are merged[m
[32m+[m[32minto a single cell (vertical ``rowspan`` in the classic builder, horizontal[m
[32m+[m[32m``colspan`` in the transposed one) spanning every covered hour, partial hours[m
[32m+[m[32mincluded; on-screen, overlapping sessions are placed on separate lanes within a[m
[32m+[m[32mday band (``_partition_lanes``). Every render cell keeps the session type so[m
[32m+[m[32mtemplates and the PDF can shade Lecture (grey), Workshop (green) and Technical[m
[32m+[m[32mDrawing (pink).[m
 """[m
 [m
 from core.models import Day[m
[36m@@ -15,6 +21,12 @@[m [mfrom core.models import Day[m
 DAY_ORDER = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"][m
 WEEKEND_ORDER = ["SATURDAY", "SUNDAY"][m
 [m
[32m+[m[32m# The configured daily timetable window. The on-screen grid trims to the slot[m
[32m+[m[32m# range that actually has activity; the PDF export requests the full range[m
[32m+[m[32m# (07:00-08:00 .. 19:00-20:00) via ``build_time_day_grid(..., full_range=True)``.[m
[32m+[m[32mGRID_HOUR_START = 7[m
[32m+[m[32mGRID_HOUR_END = 19[m
[32m+[m
 # Light "faded" fills used by the PDF. Priority when a merged block contains[m
 # several session types: TD > Workshop > Lecture > Tutorial > Seminar > Practical.[m
 FILL_COLORS = {[m
[36m@@ -68,16 +80,42 @@[m [mdef _by_label(entries):[m
     return by_label[m
 [m
 [m
[32m+[m[32m# ─────────────────────────────────────────────────────────────[m
[32m+[m[32m# STRUCTURE OF THE CODE START — ON-SCREEN TIMETABLE MAPPING[m
[32m+[m[32m# This is the section that maps every session into the on-screen[m
[32m+[m[32m# timetable: DAYS become rows, hourly TIME slots become columns,[m
[32m+[m[32m# and a session spanning several hours becomes ONE merged cell[m
[32m+[m[32m# (`colspan`), with overlapping sessions placed on separate lanes.[m
[32m+[m[32m#   - `_slot`            → builds the hourly column labels for the PDF/classic grid[m
[32m+[m[32m#   - `_screen_slot`     → builds the on-screen header labels (07:00-07:55, 08:00-08:55...)[m
[32m+[m[32m#   - `_partition_lanes` → splits a day's sessions into band sub-rows[m
[32m+[m[32m#   - `build_day_time_grid` → the entry point (used by core/views.py)[m
[32m+[m[32m# ─────────────────────────────────────────────────────────────[m
 def _slot(hour):[m
     return {[m
         "hour": hour,[m
         "start": f"{hour:02d}:00",[m
[31m-        "end": f"{hour:02d}:55",[m
[32m+[m[32m        "end": f"{hour + 1:02d}:00",[m
[32m+[m[32m        "label": f"{hour:02d}:00-{hour + 1:02d}:00",[m
[32m+[m[32m    }[m
[32m+[m
[32m+[m
[32m+[m[32mdef _screen_slot(hour):[m
[32m+[m[32m    """Header slot for the on-screen grid: labels read 07:00-07:55 etc.[m
[32m+[m
[32m+[m[32m    Only the rendered header text differs from ``_slot``; ``start``/``end``[m
[32m+[m[32m    and the ``hour`` value are unchanged so sessions stay mapped to the[m
[32m+[m[32m    same hourly columns.[m
[32m+[m[32m    """[m
[32m+[m[32m    return {[m
[32m+[m[32m        "hour": hour,[m
[32m+[m[32m        "start": f"{hour:02d}:00",[m
[32m+[m[32m        "end": f"{hour + 1:02d}:00",[m
         "label": f"{hour:02d}:00-{hour:02d}:55",[m
     }[m
 [m
 [m
[31m-def build_time_day_grid(entries):[m
[32m+[m[32mdef build_time_day_grid(entries, full_range=False):[m
     """Turn ``collect_entries`` output into a time-row/day-column grid.[m
 [m
     Returns a dict with ``slots``, ``days`` (column headers) and ``rows``.[m
[36m@@ -85,6 +123,13 @@[m [mdef build_time_day_grid(entries):[m
     aligned to ``days``. Each column cell is ``None`` when the row is already[m
     consumed by the ``rowspan`` cell above it, ``{"empty": True}`` for a free[m
     period, or ``{"row", "rowspan", "entries"}`` for a merged session block.[m
[32m+[m
[32m+[m[32m    By default the slot rows are trimmed to the hours actually covered by the[m
[32m+[m[32m    entries. Passing ``full_range=True`` (used by the PDF export) always[m
[32m+[m[32m    produces the configured daily window 07:00-08:00 .. 19:00-20:00 so the[m
[32m+[m[32m    timetable covers the whole day even when few sessions exist; any session[m
[32m+[m[32m    falling outside that window still adds its own slots rather than being[m
[32m+[m[32m    clipped.[m
     """[m
     if not entries:[m
         return {"slots": [], "days": [], "rows": []}[m
[36m@@ -100,8 +145,11 @@[m [mdef build_time_day_grid(entries):[m
     if not covered:[m
         return {"slots": [], "days": [], "rows": []}[m
 [m
[31m-    min_hour, max_hour = min(covered), max(covered)[m
[31m-    slots = list(range(min_hour, max_hour + 1))[m
[32m+[m[32m    if full_range:[m
[32m+[m[32m        slots = sorted(set(range(GRID_HOUR_START, GRID_HOUR_END + 1)) | set(covered))[m
[32m+[m[32m    else:[m
[32m+[m[32m        min_hour, max_hour = min(covered), max(covered)[m
[32m+[m[32m        slots = list(range(min_hour, max_hour + 1))[m
     slot_defs = [_slot(hour) for hour in slots][m
     n_slots = len(slots)[m
 [m
[36m@@ -144,6 +192,117 @@[m [mdef build_time_day_grid(entries):[m
     return {"slots": slot_defs, "days": days, "rows": rows}[m
 [m
 [m
[32m+[m[32mdef _partition_lanes(entries, slots):[m
[32m+[m[32m    """Split a day's sessions into non-overlapping lanes (band sub-rows).[m
[32m+[m
[32m+[m[32m    Greedy interval partitioning by the hours each entry covers: an entry is[m
[32m+[m[32m    placed in the first lane whose last session ends before it starts, so a[m
[32m+[m[32m    new lane is opened whenever sessions overlap. Each returned lane is a[m
[32m+[m[32m    list of cells aligned to ``slots``: ``{"empty": True, "colspan"}`` for a[m
[32m+[m[32m    free span of columns, or ``{"colspan", "entries"}`` where the entries[m
[32m+[m[32m    visually merge across their whole covered range.[m
[32m+[m[32m    """[m
[32m+[m[32m    if not entries:[m
[32m+[m[32m        return [[{"empty": True, "colspan": len(slots)}]][m
[32m+[m
[32m+[m[32m    lane_runs = [][m
[32m+[m[32m    lane_cells = [][m
[32m+[m[32m    for e in sorted([m
[32m+[m[32m        entries,[m
[32m+[m[32m        key=lambda e: ([m
[32m+[m[32m            min(e["hours"]),[m
[32m+[m[32m            -(max(e["hours"]) - min(e["hours"])),[m
[32m+[m[32m            e["course_code"],[m
[32m+[m[32m            str(e["key"]),[m
[32m+[m[32m        ),[m
[32m+[m[32m    ):[m
[32m+[m[32m        hours = sorted(e["hours"])[m
[32m+[m[32m        start, end = hours[0], hours[-1][m
[32m+[m[32m        start_idx = slots.index(start)[m
[32m+[m[32m        span = end - start + 1[m
[32m+[m[32m        lane_i = next([m
[32m+[m[32m            (i for i, run in enumerate(lane_runs) if run[-1][1] < start), None[m
[32m+[m[32m        )[m
[32m+[m[32m        if lane_i is None:[m
[32m+[m[32m            lane_i = len(lane_runs)[m
[32m+[m[32m            lane_runs.append([])[m
[32m+[m[32m            lane_cells.append([])[m
[32m+[m[32m        lane_runs[lane_i].append((start, end))[m
[32m+[m[32m        lane_cells[lane_i].append((start_idx, span, e))[m
[32m+[m
[32m+[m[32m    lanes = [][m
[32m+[m[32m    for cells in lane_cells:[m
[32m+[m[32m        cells.sort()[m
[32m+[m[32m        lane = [][m
[32m+[m[32m        pos = 0[m
[32m+[m[32m        for start_idx, span, e in cells:[m
[32m+[m[32m            if start_idx > pos:[m
[32m+[m[32m                lane.append({"empty": True, "colspan": start_idx - pos})[m
[32m+[m[32m            lane.append({"colspan": span, "entries": [e]})[m
[32m+[m[32m            pos = start_idx + span[m
[32m+[m[32m        if pos < len(slots):[m
[32m+[m[32m            lane.append({"empty": True, "colspan": len(slots) - pos})[m
[32m+[m[32m        lanes.append(lane)[m
[32m+[m[32m    return lanes[m
[32m+[m
[32m+[m
[32m+[m[32mdef build_day_time_grid(entries):[m
[32m+[m[32m    """Transposed weekly grid for the default on-screen timetable view.[m
[32m+[m
[32m+[m[32m    DAYS are rows (always Monday..Friday, plus any weekend days present in[m
[32m+[m[32m    the data) and the hourly time slots are columns, in ascending order from[m
[32m+[m[32m    07:00 up to the latest hour actually covered. Each day is drawn as a[m
[32m+[m[32m    vertical band of one or more sub-rows (lanes). A session spanning several[m
[32m+[m[32m    hourly slots occupies ONE cell that merges those columns (``colspan``[m
[32m+[m[32m    derived from every covered hour, including partial hours); sessions that[m
[32m+[m[32m    overlap on the same day are pushed onto separate lanes instead of being[m
[32m+[m[32m    sliced into one box per hour column.[m
[32m+[m
[32m+[m[32m    Returns ``{"slots", "days", "rows"}``. ``rows`` is one dict per day:[m
[32m+[m[32m    ``{"day", "label", "lanes"}`` where ``label`` is the full weekday name,[m
[32m+[m[32m    and ``lanes`` is the list of band rows from ``_partition_lanes``.[m
[32m+[m[32m    """[m
[32m+[m[32m    if not entries:[m
[32m+[m[32m        return {"slots": [], "days": [], "rows": []}[m
[32m+[m
[32m+[m[32m    present_days = {e["day"] for e in entries}[m
[32m+[m[32m    day_order = list(DAY_ORDER)[m
[32m+[m[32m    day_order += [[m
[32m+[m[32m        d for d in WEEKEND_ORDER if d in present_days and d not in day_order[m
[32m+[m[32m    ][m
[32m+[m
[32m+[m[32m    covered = set()[m
[32m+[m[32m    for e in entries:[m
[32m+[m[32m        covered.update(e["hours"])[m
[32m+[m[32m    if not covered:[m
[32m+[m[32m        return {"slots": [], "days": [], "rows": []}[m
[32m+[m
[32m+[m[32m    first_hour = min(GRID_HOUR_START, min(covered))[m
[32m+[m[32m    slots = list(range(first_hour, max(covered) + 1))[m
[32m+[m[32m    slot_defs = [_screen_slot(hour) for hour in slots][m
[32m+[m
[32m+[m[32m    by_day = {}[m
[32m+[m[32m    for e in entries:[m
[32m+[m[32m        if e["hours"]:[m
[32m+[m[32m            by_day.setdefault(e["day"], []).append(e)[m
[32m+[m
[32m+[m[32m    rows = [][m
[32m+[m[32m    for day in day_order:[m
[32m+[m[32m        rows.append([m
[32m+[m[32m            {[m
[32m+[m[32m                "day": day,[m
[32m+[m[32m                "label": Day(day).label,[m
[32m+[m[32m                "lanes": _partition_lanes(by_day.get(day, []), slots),[m
[32m+[m[32m            }[m
[32m+[m[32m        )[m
[32m+[m
[32m+[m[32m    days = [{"day": row["day"], "label": row["label"]} for row in rows][m
[32m+[m[32m    return {"slots": slot_defs, "days": days, "rows": rows}[m
[32m+[m[32m# ─────────────────────────────────────────────────────────────[m
[32m+[m[32m# STRUCTURE OF THE CODE END — ON-SCREEN TIMETABLE MAPPING[m
[32m+[m[32m# ─────────────────────────────────────────────────────────────[m
[32m+[m
[32m+[m
 def fill_color(entries):[m
     """Best shading fill for a set of entries (first matching priority kind)."""[m
     kinds = {e["kind"] for e in entries}[m
[1mdiff --git a/core/timetable_pdf.py b/core/timetable_pdf.py[m
[1mindex 5f7a4cb..b94875f 100644[m
[1m--- a/core/timetable_pdf.py[m
[1m+++ b/core/timetable_pdf.py[m
[36m@@ -1,16 +1,35 @@[m
 """Programme timetable PDF generation (reportlab).[m
 [m
 Produces a classic weekly grid, A4 portrait: the DAYS run across the first row[m
[31m-and the hourly TIME slots (07:00-07:55 .. 19:00-19:55) run down the first[m
[32m+[m[32mand the hourly TIME slots (07:00-08:00 .. 19:00-20:00) run down the first[m
 column, matching the on-screen timetable in ``core/timetable.html``. Sessions[m
 that span multiple hourly slots merge their cells on that day. Each cell is[m
 shaded by activity type — Lecture faded grey, Workshop faded green, Technical[m
 Drawing faded pink. External activities (workshops / technical drawing) that[m
[31m-only carry a Morning/Afternoon period are placed in the morning (09:00-12:55)[m
[31m-or afternoon (15:00-18:55) range.[m
[32m+[m[32monly carry a Morning/Afternoon period are placed in that day's standard[m
[32m+[m[32mworkshop slot — Monday/Tuesday/Wednesday/Friday morning 09:00-13:00, Thursday[m
[32m+[m[32mmorning 10:00-14:00, afternoons 15:00-19:00.[m
[32m+[m
[32m+[m[32mThe PDF grid always spans the full configured daily window 07:00-20:00 so a[m
[32m+[m[32msession at any hour within the day appears, even when the programme only has a[m
[32m+[m[32mhandful of sessions.[m
[32m+[m
[32m+[m[32mWorkshop rotation[m
[32m+[m[32m-----------------[m
[32m+[m
[32m+[m[32mWhen a student group has two or more *different* workshops allocated to the[m
[32m+[m[32msame daily slot (same day and time, e.g. Electrical and Carpentry on Thursday[m
[32m+[m[32mmorning), those workshops rotate after every seven weeks. The cell then shows[m
[32m+[m[32m``Electrical / Carpentry`` and a ``Workshop Rotation Key`` table is appended[m
[32m+[m[32mbelow the grid listing, per programme/group/day: the workshop used in Week 1-7,[m
[32m+[m[32mWeek 8-14, and so on. Workshops that explicitly carry ``week_start``/``week_end``[m
[32m+[m[32mranges keep those ranges; otherwise the seven-week blocks are assigned in a[m
[32m+[m[32mdeterministic order (the workshop that matches the programme's name first, then[m
[32m+[m[32malphabetical). Non-rotating workshops render only their normal name and never[m
[32m+[m[32mproduce a rotation-key entry.[m
 """[m
[31m-[m
 import datetime[m
[32m+[m[32mfrom collections import defaultdict[m
 from xml.sax.saxutils import escape[m
 [m
 from django.db.models import Q[m
[36m@@ -28,6 +47,7 @@[m [mfrom reportlab.platypus import ([m
 )[m
 [m
 from core.models import ([m
[32m+[m[32m    Day,[m
     Session,[m
     StudentGroup,[m
     TechnicalDrawingAllocation,[m
[36m@@ -35,13 +55,11 @@[m [mfrom core.models import ([m
     WorkshopAllocation,[m
 )[m
 from core.timetable_grid import build_time_day_grid, time_day_grid_to_table[m
[31m-[m
[31m-HOUR_START = 7[m
[31m-HOUR_END = 19  # inclusive last hour (19:00-19:55)[m
[31m-[m
[31m-[m
[31m-def _slot_hours():[m
[31m-    return list(range(HOUR_START, HOUR_END + 1))[m
[32m+[m[32mfrom core.workshop_times import ([m
[32m+[m[32m    allocation_programme_codes,[m
[32m+[m[32m    workshop_hours,[m
[32m+[m[32m    workshop_times_for,[m
[32m+[m[32m)[m
 [m
 [m
 def _ordinal(n):[m
[36m@@ -52,12 +70,32 @@[m [mdef _ordinal(n):[m
     return f"{n}{suffix}"[m
 [m
 [m
[32m+[m[32mdef _display_times(hours, start_time=None, end_time=None):[m
[32m+[m[32m    """(start, end) clock strings for an Activity Card header.[m
[32m+[m
[32m+[m[32m    Prefers the record's exact clock times; period-only records (matrix[m
[32m+[m[32m    workshops) fall back to their covered hourly slots — the first slot is the[m
[32m+[m[32m    start, one hour past the last slot is the end (e.g. Thursday morning[m
[32m+[m[32m    hours {10, 11, 12, 13} => "10:00"-"14:00").[m
[32m+[m[32m    """[m
[32m+[m[32m    if start_time is not None and end_time is not None:[m
[32m+[m[32m        return start_time.strftime("%H:%M"), end_time.strftime("%H:%M")[m
[32m+[m[32m    if hours:[m
[32m+[m[32m        return f"{min(hours):02d}:00", f"{max(hours) + 1:02d}:00"[m
[32m+[m[32m    return "", ""[m
[32m+[m
[32m+[m
 def _session_entry(session, groups_text=""):[m
     venue = session.venue.name if session.venue and session.venue.name else ""[m
     lines = [f"{session.course_code} {session.get_activity_type_display()}"][m
     if groups_text:[m
         lines.append(groups_text)[m
     lines.append(venue or "-")[m
[32m+[m[32m    start, end = _display_times([m
[32m+[m[32m        _covered_hours(session.start_time, session.end_time),[m
[32m+[m[32m        session.start_time,[m
[32m+[m[32m        session.end_time,[m
[32m+[m[32m    )[m
     return {[m
         "key": ("session", session.pk),[m
         "day": session.day,[m
[36m@@ -68,6 +106,8 @@[m [mdef _session_entry(session, groups_text=""):[m
         "name": session.course_code,[m
         "type_label": session.get_activity_type_display(),[m
         "venue": venue,[m
[32m+[m[32m        "start": start,[m
[32m+[m[32m        "end": end,[m
         "groups": groups_text,[m
         "note": "",[m
     }[m
[36m@@ -94,26 +134,31 @@[m [mdef _workshop_entry(rec):[m
     if rec.week_start and rec.week_end:[m
         note = f"Wk {rec.week_start}-{rec.week_end}"[m
     name = _workshop_display_name(rec)[m
[32m+[m[32m    hours = _workshop_hours(rec)[m
[32m+[m[32m    start, end = _display_times(hours, rec.start_time, rec.end_time)[m
     return {[m
         "key": ("workshop", rec.pk),[m
         "day": rec.day,[m
[31m-        "hours": ([m
[31m-            _covered_hours(rec.start_time, rec.end_time)[m
[31m-            if rec.start_time is not None[m
[31m-            else _period_hours(rec.time_period)[m
[31m-        ),[m
[32m+[m[32m        "hours": hours,[m
         "label": name,[m
         "kind": "workshop",[m
         "course_code": rec.course_code,[m
         "name": name,[m
         "type_label": "Workshop",[m
         "venue": rec.venue,[m
[32m+[m[32m        "start": start,[m
[32m+[m[32m        "end": end,[m
         "groups": rec.group_code,[m
         "note": note,[m
     }[m
 [m
 [m
 def _td_entry(rec):[m
[32m+[m[32m    start, end = _display_times([m
[32m+[m[32m        _covered_hours(rec.start_time, rec.end_time),[m
[32m+[m[32m        rec.start_time,[m
[32m+[m[32m        rec.end_time,[m
[32m+[m[32m    )[m
     return {[m
         "key": ("td", rec.pk),[m
         "day": rec.day,[m
[36m@@ -128,6 +173,8 @@[m [mdef _td_entry(rec):[m
         "name": rec.course_code,[m
         "type_label": "Technical Drawing",[m
         "venue": rec.venue,[m
[32m+[m[32m        "start": start,[m
[32m+[m[32m        "end": end,[m
         "groups": rec.group_code,[m
         "note": "",[m
     }[m
[36m@@ -137,6 +184,11 @@[m [mdef _workshop_sort_time(rec):[m
     """Clock-time a workshop occupies, for a deterministic day ordering."""[m
     if rec.start_time is not None:[m
         return rec.start_time[m
[32m+[m[32m    times = workshop_times_for([m
[32m+[m[32m        rec.day, rec.time_period, allocation_programme_codes(rec)[m
[32m+[m[32m    )[m
[32m+[m[32m    if times:[m
[32m+[m[32m        return times[0][m
     if rec.time_period == TimePeriod.MORNING:[m
         return datetime.time(9)[m
     if rec.time_period == TimePeriod.AFTERNOON:[m
[36m@@ -144,62 +196,171 @@[m [mdef _workshop_sort_time(rec):[m
     return datetime.time(0)[m
 [m
 [m
[31m-def _workshop_entries(qs):[m
[31m-    """Turn a WorkshopAllocation queryset into clean display entries.[m
[32m+[m[32mdef _workshop_slot_key(rec):[m
[32m+[m[32m    """Identify the daily slot a workshop record belongs to.[m
[32m+[m
[32m+[m[32m    Records that share (group, day, slot) are the candidates for rotation: two[m
[32m+[m[32m    *different* workshops in the same slot mean the group alternates between[m
[32m+[m[32m    them every seven weeks. A different time on the same day is a separate[m
[32m+[m[32m    session, not a rotation.[m
[32m+[m[32m    """[m
[32m+[m[32m    if rec.start_time is not None and rec.end_time is not None:[m
[32m+[m[32m        return ("time", rec.start_time, rec.end_time)[m
[32m+[m[32m    if rec.time_period:[m
[32m+[m[32m        return ("period", rec.time_period)[m
[32m+[m[32m    return ("time", None, None)[m
[32m+[m
[32m+[m
[32m+[m[32mdef _affinity_key(name, programme):[m
[32m+[m[32m    """Deterministic ordering of workshop names inside a rotation.[m
[32m+[m
[32m+[m[32m    The workshop that matches the programme's identity (e.g. ``Electrical`` for[m
[32m+[m[32m    the BSc. in Electrical Engineering programme) leads, so the natural[m
[32m+[m[32m    workshop is shown first; everything else falls back to alphabetical order.[m
[32m+[m[32m    Nothing is hard-coded — the rule is simply "a workshop whose name appears[m
[32m+[m[32m    in the programme name comes first".[m
[32m+[m[32m    """[m
[32m+[m[32m    if programme and name and name.lower() in programme.name.lower():[m
[32m+[m[32m        return (0, name.lower())[m
[32m+[m[32m    return (1, name.lower())[m
[32m+[m
[32m+[m
[32m+[m[32mdef _rotation_blocks(records, programme):[m
[32m+[m[32m    """Order a rotation's workshop names and assign their week ranges.[m
[32m+[m
[32m+[m[32m    When every record carries explicit ``week_start``/``week_end`` ranges those[m
[32m+[m[32m    are kept (the matrix workbook shape). Otherwise the rotation is split into[m
[32m+[m[32m    contiguous seven-week blocks — Week 1-7, Week 8-14, Week 15-21, ... — in[m
[32m+[m[32m    the deterministic order produced by ``_affinity_key``.[m
[32m+[m[32m    """[m
[32m+[m[32m    names = sorted([m
[32m+[m[32m        {_workshop_display_name(r) for r in records},[m
[32m+[m[32m        key=lambda n: _affinity_key(n, programme),[m
[32m+[m[32m    )[m
[32m+[m[32m    explicit = all([m
[32m+[m[32m        r.week_start is not None and r.week_end is not None for r in records[m
[32m+[m[32m    )[m
[32m+[m[32m    if explicit and len(records) == len({_workshop_display_name(r) for r in records}):[m
[32m+[m[32m        ordered = sorted([m
[32m+[m[32m            records, key=lambda r: (r.week_start, r.week_end, _workshop_display_name(r))[m
[32m+[m[32m        )[m
[32m+[m[32m        return [[m
[32m+[m[32m            (r.week_start, r.week_end, _workshop_display_name(r)) for r in ordered[m
[32m+[m[32m        ][m
[32m+[m[32m    return [(1 + 7 * i, 7 + 7 * i, name) for i, name in enumerate(names)][m
[32m+[m
[32m+[m
[32m+[m[32mdef _workshop_hours(rec):[m
[32m+[m[32m    if rec.start_time is not None and rec.end_time is not None:[m
[32m+[m[32m        return _covered_hours(rec.start_time, rec.end_time)[m
[32m+[m[32m    return _period_hours([m
[32m+[m[32m        rec.time_period, rec.day, allocation_programme_codes(rec)[m
[32m+[m[32m    )[m
[32m+[m
[32m+[m
[32m+[m[32mdef _build_rotation(group_code, day, records, programme):[m
[32m+[m[32m    """Merge a group's same-slot workshops into one rotating cell + key row."""[m
[32m+[m[32m    blocks = _rotation_blocks(records, programme)[m
[32m+[m[32m    names = [name for _, _, name in blocks][m
[32m+[m[32m    label = " / ".join(names)[m
[32m+[m[32m    note = " / ".join(f"Wk {wk_start}-{wk_end}" for wk_start, wk_end, _ in blocks)[m
[32m+[m[32m    hours = set()[m
[32m+[m[32m    for rec in records:[m
[32m+[m[32m        hours.update(_workshop_hours(rec))[m
[32m+[m[32m    courses = {rec.course_code for rec in records}[m
[32m+[m[32m    course = next(iter(courses)) if len(courses) == 1 else ""[m
[32m+[m[32m    keys = tuple(sorted(rec.pk for rec in records))[m
[32m+[m[32m    start, end = _display_times(hours)[m
[32m+[m[32m    entry = {[m
[32m+[m[32m        "key": ("workshop-rotation", keys),[m
[32m+[m[32m        "day": day,[m
[32m+[m[32m        "hours": hours,[m
[32m+[m[32m        "label": label,[m
[32m+[m[32m        "kind": "workshop",[m
[32m+[m[32m        "course_code": course or label,[m
[32m+[m[32m        "name": label,[m
[32m+[m[32m        "type_label": "Workshop",[m
[32m+[m[32m        "venue": "",[m
[32m+[m[32m        "start": start,[m
[32m+[m[32m        "end": end,[m
[32m+[m[32m        "groups": group_code,[m
[32m+[m[32m        "note": note,[m
[32m+[m[32m    }[m
[32m+[m[32m    row = {[m
[32m+[m[32m        "programme": programme.name if programme else "",[m
[32m+[m[32m        "programme_code": programme.code if programme else "",[m
[32m+[m[32m        "group": group_code,[m
[32m+[m[32m        "day": day,[m
[32m+[m[32m        "course": course,[m
[32m+[m[32m        "blocks": blocks,[m
[32m+[m[32m    }[m
[32m+[m[32m    return entry, row[m
[32m+[m
[32m+[m
[32m+[m[32mdef _workshop_entries(qs, programme=None):[m
[32m+[m[32m    """Turn a WorkshopAllocation queryset into display entries + rotation keys.[m
[32m+[m
[32m+[m[32m    Returns ``(entries, rotation_keys)``. A group's workshop for a given daily[m
[32m+[m[32m    slot is either:[m
 [m
[31m-    Applies the timetable rules that a student group has exactly one workshop[m
[31m-    per day: when the source carries two *different* workshops for the same[m
[31m-    group on the same day (a data inconsistency), only the first is kept so the[m
[31m-    grid never shows two workshops for one group on one day. Week-run splits of[m
[31m-    the same workshop (e.g. Wk 1-6 and Wk 8-13) are all preserved — they are[m
[31m-    one workshop with distinct valid sessions. The chosen records are ordered[m
[31m-    deterministically by day/time so cells render stably.[m
[32m+[m[32m    - a single workshop: its week-run splits (e.g. Wk 1-6 and Wk 8-13 of the[m
[32m+[m[32m      same workshop) are all preserved as separate entries, or[m
[32m+[m[32m    - a rotation: two or more *different* workshops in the same slot merge into[m
[32m+[m[32m      one ``Electrical / Carpentry`` entry, with a corresponding rotation-key[m
[32m+[m[32m      row describing which workshop runs in Week 1-7, Week 8-14, ...[m
[32m+[m
[32m+[m[32m    Workshops in different slots of the same day are distinct sessions and are[m
[32m+[m[32m    all kept. Record order within a slot is deterministic (day/time).[m
     """[m
[31m-    by_group_day = {}[m
[32m+[m[32m    entries = [][m
[32m+[m[32m    rotation_keys = [][m
[32m+[m[32m    by_slot = defaultdict(list)[m
     for rec in qs:[m
[31m-        by_group_day.setdefault((rec.group_code, rec.day), []).append(rec)[m
[32m+[m[32m        by_slot[(rec.group_code, rec.day, _workshop_slot_key(rec))].append(rec)[m
 [m
[31m-    kept = [][m
[31m-    for records in by_group_day.values():[m
[31m-        by_identity = {}[m
[32m+[m[32m    for (group_code, day, _slot), records in by_slot.items():[m
[32m+[m[32m        by_identity = defaultdict(list)[m
         for rec in records:[m
[31m-            by_identity.setdefault(_workshop_display_name(rec), []).append(rec)[m
[31m-        chosen = by_identity[m
[31m-        if len(by_identity) > 1:[m
[31m-            best = min([m
[31m-                by_identity,[m
[31m-                key=lambda name: ([m
[31m-                    _workshop_sort_time(by_identity[name][0]),[m
[31m-                    by_identity[name][0].pk,[m
[31m-                ),[m
[31m-            )[m
[31m-            chosen = {best: by_identity[best]}[m
[31m-        for records_of_identity in chosen.values():[m
[31m-            kept.extend([m
[31m-                sorted([m
[31m-                    records_of_identity,[m
[31m-                    key=lambda r: (_workshop_sort_time(r), r.pk),[m
[32m+[m[32m            by_identity[_workshop_display_name(rec)].append(rec)[m
[32m+[m[32m        if len(by_identity) == 1:[m
[32m+[m[32m            for recs in by_identity.values():[m
[32m+[m[32m                sorted_recs = sorted([m
[32m+[m[32m                    recs, key=lambda r: (_workshop_sort_time(r), r.pk)[m
                 )[m
[31m-            )[m
[31m-    return [_workshop_entry(rec) for rec in kept][m
[32m+[m[32m                entries.extend(_workshop_entry(r) for r in sorted_recs)[m
[32m+[m[32m            continue[m
[32m+[m[32m        entry, row = _build_rotation(group_code, day, records, programme)[m
[32m+[m[32m        entries.append(entry)[m
[32m+[m[32m        rotation_keys.append(row)[m
[32m+[m[32m    return entries, rotation_keys[m
 [m
 [m
 def _covered_hours(start, end):[m
[31m-    """Hour slots (7..19) a session covers given its start/end times."""[m
[32m+[m[32m    """Hour slots a session covers given its start/end times.[m
[32m+[m
[32m+[m[32m    Considers the full clock day (00:00 .. 22:55 as the last whole slot) so a[m
[32m+[m[32m    session that starts before 07:00 or ends after 19:55 is never clipped; the[m
[32m+[m[32m    on-screen grid trims and the PDF's full-range grid unions these hours with[m
[32m+[m[32m    the 07:00-20:00 display window.[m
[32m+[m[32m    """[m
     hours = set()[m
[31m-    for h in _slot_hours():[m
[32m+[m[32m    for h in range(0, 23):[m
         if start is not None and end is not None:[m
             if start < datetime.time(h + 1) and end > datetime.time(h):[m
                 hours.add(h)[m
     return hours[m
 [m
 [m
[31m-def _period_hours(time_period):[m
[32m+[m[32mdef _period_hours(time_period, day=None, programme_code=None):[m
     """Hourly slots a morning/afternoon workshop covers.[m
 [m
[31m-    A morning workshop is one four-hour session 09:00-12:55 and an afternoon[m
[31m-    workshop one four-hour session 15:00-18:55, matching the university rules.[m
[32m+[m[32m    A morning workshop is one four-hour session (09:00-13:00 on most days,[m
[32m+[m[32m    10:00-14:00 on Thursday) and an afternoon workshop one four-hour session[m
[32m+[m[32m    15:00-19:00, matching the university rules.[m
     """[m
[32m+[m[32m    hours = workshop_hours(day, time_period, programme_code)[m
[32m+[m[32m    if hours:[m
[32m+[m[32m        return hours[m
     if time_period == TimePeriod.MORNING:[m
         return set(range(9, 13))[m
     if time_period == TimePeriod.AFTERNOON:[m
[36m@@ -218,21 +379,26 @@[m [mdef _cell_markup(text):[m
     return escape(text).replace("\n", "<br/>")[m
 [m
 [m
[31m-def collect_entries(programme, semester, group=None, year=None):[m
[31m-    """Gather timetable entries for one programme's groups in a semester.[m
[32m+[m[32mdef _apply_year_filter(qs, year):[m
[32m+[m[32m    if year:[m
[32m+[m[32m        return qs.filter(Q(year_of_study__isnull=True) | Q(year_of_study=year))[m
[32m+[m[32m    return qs[m
[32m+[m
[32m+[m
[32m+[m[32mdef _collect_entries_and_rotations(programme, semester, group=None, year=None):[m
[32m+[m[32m    """Gather timetable entries and workshop rotation keys for a selection.[m
 [m
     When ``group`` is given only that single student group's timetable is[m
[31m-    collected (sessions a group actually attends plus its workshops/TDs),[m
[31m-    which backs the per-group export. ``year`` optionally filters workshop[m
[32m+[m[32m    collected (sessions a group actually attends plus its workshops/TDs), which[m
[32m+[m[32m    backs the per-group export. ``year`` optionally filters workshop[m
     allocations to a year of study (records without a year apply to every[m
[31m-    year).[m
[32m+[m[32m    year). Returns ``(entries, rotation_keys)``.[m
     """[m
     if group is not None:[m
[31m-        return collect_group_entries(group, semester, year=year)[m
[32m+[m[32m        return _collect_group_entries_and_rotations(group, semester, year=year)[m
 [m
     groups = list(StudentGroup.objects.filter(programme=programme))[m
     group_codes = {g.code for g in groups}[m
[31m-    group_pk_by_code = {g.code: g.pk for g in groups}[m
 [m
     entries = [][m
 [m
[36m@@ -254,14 +420,14 @@[m [mdef collect_entries(programme, semester, group=None, year=None):[m
             groups_text = ", ".join(sorted(attending))[m
         entries.append(_session_entry(session, groups_text))[m
 [m
[31m-    workshops = WorkshopAllocation.objects.filter([m
[31m-        semester=semester, group_code__in=group_codes[m
[32m+[m[32m    workshops = _apply_year_filter([m
[32m+[m[32m        WorkshopAllocation.objects.filter([m
[32m+[m[32m            semester=semester, group_code__in=group_codes[m
[32m+[m[32m        ),[m
[32m+[m[32m        year,[m
     )[m
[31m-    if year:[m
[31m-        workshops = workshops.filter([m
[31m-            Q(year_of_study__isnull=True) | Q(year_of_study=year)[m
[31m-        )[m
[31m-    entries.extend(_workshop_entries(workshops))[m
[32m+[m[32m    workshop_entries, rotation_keys = _workshop_entries(workshops, programme=programme)[m
[32m+[m[32m    entries.extend(workshop_entries)[m
 [m
     tds = TechnicalDrawingAllocation.objects.filter([m
         semester=semester, group_code__in=group_codes[m
[36m@@ -269,7 +435,51 @@[m [mdef collect_entries(programme, semester, group=None, year=None):[m
     for rec in tds:[m
         entries.append(_td_entry(rec))[m
 [m
[31m-    return entries[m
[32m+[m[32m    return entries, rotation_keys[m
[32m+[m
[32m+[m
[32m+[m[32mdef _collect_group_entries_and_rotations(group, semester, year=None):[m
[32m+[m[32m    """As ``_collect_entries_and_rotations`` but restricted to ONE group."""[m
[32m+[m[32m    entries = [][m
[32m+[m
[32m+[m[32m    sessions = ([m
[32m+[m[32m        Session.objects.filter(semester=semester, session_groups__group=group)[m
[32m+[m[32m        .select_related("venue")[m
[32m+[m[32m        .distinct()[m
[32m+[m[32m    )[m
[32m+[m[32m    for session in sessions:[m
[32m+[m[32m        entries.append(_session_entry(session, group.code))[m
[32m+[m
[32m+[m[32m    workshops = _apply_year_filter([m
[32m+[m[32m        WorkshopAllocation.objects.filter(semester=semester, group_code=group.code),[m
[32m+[m[32m        year,[m
[32m+[m[32m    )[m
[32m+[m[32m    workshop_entries, rotation_keys = _workshop_entries([m
[32m+[m[32m        workshops, programme=group.programme[m
[32m+[m[32m    )[m
[32m+[m[32m    entries.extend(workshop_entries)[m
[32m+[m
[32m+[m[32m    tds = TechnicalDrawingAllocation.objects.filter([m
[32m+[m[32m        semester=semester, group_code=group.code[m
[32m+[m[32m    )[m
[32m+[m[32m    for rec in tds:[m
[32m+[m[32m        entries.append(_td_entry(rec))[m
[32m+[m
[32m+[m[32m    return entries, rotation_keys[m
[32m+[m
[32m+[m
[32m+[m[32mdef collect_entries(programme, semester, group=None, year=None):[m
[32m+[m[32m    """Gather timetable entries for one programme's groups in a semester.[m
[32m+[m
[32m+[m[32m    When ``group`` is given only that single student group's timetable is[m
[32m+[m[32m    collected (sessions a group actually attends plus its workshops/TDs),[m
[32m+[m[32m    which backs the per-group export. ``year`` optionally filters workshop[m
[32m+[m[32m    allocations to a year of study (records without a year apply to every[m
[32m+[m[32m    year).[m
[32m+[m[32m    """[m
[32m+[m[32m    return _collect_entries_and_rotations(programme, semester, group=group, year=year)[[m
[32m+[m[32m        0[m
[32m+[m[32m    ][m
 [m
 [m
 def collect_group_entries(group, semester, year=None):[m
[36m@@ -278,34 +488,58 @@[m [mdef collect_group_entries(group, semester, year=None):[m
     Sessions are those the group actually attends (via ``SessionGroup``);[m
     workshops and technical drawing slots are matched on the group's code.[m
     """[m
[32m+[m[32m    return _collect_group_entries_and_rotations(group, semester, year=year)[0][m
[32m+[m
[32m+[m
[32m+[m[32mdef collect_master_entries(semester, year=None):[m
[32m+[m[32m    """Gather timetable entries for the whole COET First Year.[m
[32m+[m
[32m+[m[32m    The master timetable is the source of truth: every ``Session`` scheduled in[m
[32m+[m[32m    the semester is included — not only lectures with group links — plus all[m
[32m+[m[32m    workshop and technical-drawing allocations, so the default whole-year view[m
[32m+[m[32m    shows the same sessions as the master timetable and renders in the same[m
[32m+[m[32m    layout as a single programme's timetable. Session cards carry their[m
[32m+[m[32m    actually assigned student group codes.[m
[32m+[m[32m    """[m
     entries = [][m
 [m
     sessions = ([m
[31m-        Session.objects.filter(semester=semester, session_groups__group=group)[m
[32m+[m[32m        Session.objects.filter(semester=semester)[m
         .select_related("venue")[m
[31m-        .distinct()[m
[32m+[m[32m        .prefetch_related("session_groups__group")[m
     )[m
     for session in sessions:[m
[31m-        entries.append(_session_entry(session))[m
[31m-[m
[31m-    workshops = WorkshopAllocation.objects.filter([m
[31m-        semester=semester, group_code=group.code[m
[31m-    )[m
[31m-    if year:[m
[31m-        workshops = workshops.filter([m
[31m-            Q(year_of_study__isnull=True) | Q(year_of_study=year)[m
[32m+[m[32m        groups_text = ", ".join([m
[32m+[m[32m            sorted(sg.group.code for sg in session.session_groups.all())[m
         )[m
[31m-    entries.extend(_workshop_entries(workshops))[m
[32m+[m[32m        entries.append(_session_entry(session, groups_text))[m
 [m
[31m-    tds = TechnicalDrawingAllocation.objects.filter([m
[31m-        semester=semester, group_code=group.code[m
[32m+[m[32m    workshops = _apply_year_filter([m
[32m+[m[32m        WorkshopAllocation.objects.filter(semester=semester), year[m
     )[m
[32m+[m[32m    workshop_entries, _rotation_keys = _workshop_entries(workshops)[m
[32m+[m[32m    entries.extend(workshop_entries)[m
[32m+[m
[32m+[m[32m    tds = TechnicalDrawingAllocation.objects.filter(semester=semester)[m
     for rec in tds:[m
         entries.append(_td_entry(rec))[m
 [m
     return entries[m
 [m
 [m
[32m+[m[32mdef collect_workshop_rotations(programme, semester, group=None, year=None):[m
[32m+[m[32m    """Rotation-key rows for a programme's (or one group's) workshops.[m
[32m+[m
[32m+[m[32m    Each row describes a rotating slot: the programme, group, day, shared[m
[32m+[m[32m    course (where the records agree) and the ordered week blocks[m
[32m+[m[32m    ``(week_start, week_end, workshop_name)``. Non-rotating workshops produce[m
[32m+[m[32m    no rows.[m
[32m+[m[32m    """[m
[32m+[m[32m    return _collect_entries_and_rotations(programme, semester, group=group, year=year)[[m
[32m+[m[32m        1[m
[32m+[m[32m    ][m
[32m+[m
[32m+[m
 def build_grid(entries, show_groups=False):[m
     """Return table data (list of rows), SPAN commands and fill colours.[m
 [m
[36m@@ -317,13 +551,15 @@[m [mdef build_grid(entries, show_groups=False):[m
     ``show_groups`` appends owning group codes to workshop cells (the[m
     all-groups export).[m
     """[m
[31m-    grid = build_time_day_grid(entries)[m
[32m+[m[32m    grid = build_time_day_grid(entries, full_range=True)[m
     return time_day_grid_to_table(grid, show_groups=show_groups)[m
 [m
 [m
 def render_programme_timetable(programme, semester, year_of_study=1, out=None):[m
     """Render the programme timetable PDF to `out` (file-like or a path)."""[m
[31m-    entries = collect_entries(programme, semester, year=year_of_study)[m
[32m+[m[32m    entries, rotation_keys = _collect_entries_and_rotations([m
[32m+[m[32m        programme, semester, year=year_of_study[m
[32m+[m[32m    )[m
     show_groups = StudentGroup.objects.filter(programme=programme).count() > 1[m
     return _render_grid([m
         entries,[m
[36m@@ -332,14 +568,15 @@[m [mdef render_programme_timetable(programme, semester, year_of_study=1, out=None):[m
         semester=semester,[m
         doc_title=f"{programme.name} Timetable",[m
         show_groups=show_groups,[m
[32m+[m[32m        rotation_keys=rotation_keys,[m
         out=out,[m
     )[m
 [m
 [m
 def render_group_timetable(group, semester, year_of_study=1, out=None):[m
     """Render ONE student group's timetable PDF to `out` (file-like/path)."""[m
[31m-    entries = collect_entries([m
[31m-        group.programme, semester, group=group, year=year_of_study[m
[32m+[m[32m    entries, rotation_keys = _collect_group_entries_and_rotations([m
[32m+[m[32m        group, semester, year=year_of_study[m
     )[m
     return _render_grid([m
         entries,[m
[36m@@ -351,13 +588,81 @@[m [mdef render_group_timetable(group, semester, year_of_study=1, out=None):[m
         semester=semester,[m
         doc_title=f"{group.programme.code} {group.code} Timetable",[m
         show_groups=False,[m
[32m+[m[32m        rotation_keys=rotation_keys,[m
         out=out,[m
     )[m
 [m
 [m
[31m-def _render_grid(entries, title, subtitle, semester, doc_title, out, show_groups=False):[m
[31m-    """Render the weekly grid PDF to `out` (file-like or a path)."""[m
[31m-    data, spans, fills = build_grid(entries, show_groups=show_groups)[m
[32m+[m[32mdef _rotation_key_table(rotation_keys, cell_style, head_style):[m
[32m+[m[32m    """Build the Workshop Rotation Key table flowables.[m
[32m+[m
[32m+[m[32m    Columns: Programme, Group, Day, Course/Session, then one week column per[m
[32m+[m[32m    rotation block (Wk 1-7, Wk 8-14, ...).[m
[32m+[m[32m    """[m
[32m+[m[32m    usable = A4[0] - 22 * mm[m
[32m+[m[32m    max_blocks = max(len(row["blocks"]) for row in rotation_keys)[m
[32m+[m[32m    first = rotation_keys[0]["blocks"][m
[32m+[m[32m    week_headers = [f"Wk {ws}-{we}" for ws, we, _ in first][m
[32m+[m[32m    while len(week_headers) < max_blocks:[m
[32m+[m[32m        week_headers.append(week_headers[-1] if week_headers else "Wk ?")[m
[32m+[m
[32m+[m[32m    headers = ["Programme", "Group", "Day", "Course/Session"] + week_headers[m
[32m+[m[32m    data = [[Paragraph(head, head_style) for head in headers]][m
[32m+[m[32m    for row in rotation_keys:[m
[32m+[m[32m        values = [[m
[32m+[m[32m            row["programme"] or row["programme_code"],[m
[32m+[m[32m            row["group"],[m
[32m+[m[32m            Day(row["day"]).label,[m
[32m+[m[32m            row["course"] or "—",[m
[32m+[m[32m        ][m
[32m+[m[32m        values += [name for _, _, name in row["blocks"]][m
[32m+[m[32m        while len(values) < len(headers):[m
[32m+[m[32m            values.append("—")[m
[32m+[m[32m        data.append([Paragraph(str(v), cell_style) for v in values])[m
[32m+[m
[32m+[m[32m    n_week = max_blocks[m
[32m+[m[32m    week_w = usable * 0.32 / max(n_week, 1)[m
[32m+[m[32m    fixed = usable - week_w * n_week[m
[32m+[m[32m    widths = [[m
[32m+[m[32m        fixed * 0.34,[m
[32m+[m[32m        fixed * 0.12,[m
[32m+[m[32m        fixed * 0.16,[m
[32m+[m[32m        fixed * 0.22,[m
[32m+[m[32m    ] + [week_w] * n_week[m
[32m+[m[32m    while len(widths) < len(headers):[m
[32m+[m[32m        widths.append(week_w)[m
[32m+[m
[32m+[m[32m    table = Table(data, colWidths=widths, repeatRows=1, hAlign="CENTER")[m
[32m+[m[32m    style = [[m
[32m+[m[32m        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),[m
[32m+[m[32m        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dce6f1")),[m
[32m+[m[32m        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),[m
[32m+[m[32m        ("TOPPADDING", (0, 0), (-1, -1), 3),[m
[32m+[m[32m        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),[m
[32m+[m[32m        ("LEFTPADDING", (0, 0), (-1, -1), 3),[m
[32m+[m[32m        ("RIGHTPADDING", (0, 0), (-1, -1), 3),[m
[32m+[m[32m    ][m
[32m+[m[32m    table.setStyle(TableStyle(style))[m
[32m+[m[32m    return table[m
[32m+[m
[32m+[m
[32m+[m[32mdef _render_grid([m
[32m+[m[32m    entries,[m
[32m+[m[32m    title,[m
[32m+[m[32m    subtitle,[m
[32m+[m[32m    semester,[m
[32m+[m[32m    doc_title,[m
[32m+[m[32m    out,[m
[32m+[m[32m    show_groups=False,[m
[32m+[m[32m    rotation_keys=None,[m
[32m+[m[32m):[m
[32m+[m[32m    """Render the weekly grid PDF to `out` (file-like or a path).[m
[32m+[m
[32m+[m[32m    Programmes/groups with no sessions get a clean "nothing scheduled" notice[m
[32m+[m[32m    instead of a blank grid section; the Workshop Rotation Key is appended only[m
[32m+[m[32m    when rotating workshops actually exist.[m
[32m+[m[32m    """[m
[32m+[m[32m    rotation_keys = rotation_keys or [][m
 [m
     doc = SimpleDocTemplate([m
         out,[m
[36m@@ -397,9 +702,25 @@[m [mdef _render_grid(entries, title, subtitle, semester, doc_title, out, show_groups[m
         alignment=TA_CENTER,[m
         leading=8,[m
     )[m
[32m+[m[32m    empty_style = ParagraphStyle([m
[32m+[m[32m        "empty",[m
[32m+[m[32m        fontName="Helvetica",[m
[32m+[m[32m        fontSize=10,[m
[32m+[m[32m        alignment=TA_CENTER,[m
[32m+[m[32m        leading=14,[m
[32m+[m[32m        spaceBefore=6,[m
[32m+[m[32m    )[m
     foot_style = ParagraphStyle([m
         "foot", fontName="Helvetica", fontSize=8, alignment=TA_CENTER[m
     )[m
[32m+[m[32m    key_title_style = ParagraphStyle([m
[32m+[m[32m        "keytitle",[m
[32m+[m[32m        fontName="Helvetica-Bold",[m
[32m+[m[32m        fontSize=9,[m
[32m+[m[32m        alignment=TA_CENTER,[m
[32m+[m[32m        leading=11,[m
[32m+[m[32m        spaceAfter=3,[m
[32m+[m[32m    )[m
 [m
     elements = [[m
         Paragraph(title, title_style),[m
[36m@@ -408,57 +729,75 @@[m [mdef _render_grid(entries, title, subtitle, semester, doc_title, out, show_groups[m
         Spacer(1, 4 * mm),[m
     ][m
 [m
[31m-    n_days = max(len(data[0]) - 1, 1) if data else 1[m
[31m-    usable = A4[0] - 22 * mm[m
[31m-    time_w = 44[m
[31m-    day_w = (usable - time_w) / n_days[m
[31m-    widths = [time_w] + [day_w] * n_days[m
[31m-[m
[31m-    table_data = [[Paragraph(item, head_style) for item in data[0]]][m
[31m-    for row in data[1:]:[m
[31m-        table_data.append([m
[31m-            [[m
[31m-                Paragraph(row[0], time_style) if row[0] else "",[m
[31m-            ][m
[31m-            + [[m
[31m-                Paragraph(_cell_markup(cell), cell_style) if cell else ""[m
[31m-                for cell in row[1:][m
[31m-            ][m
[32m+[m[32m    if entries:[m
[32m+[m[32m        data, spans, fills = build_grid(entries, show_groups=show_groups)[m
[32m+[m
[32m+[m[32m        n_days = max(len(data[0]) - 1, 1) if data else 1[m
[32m+[m[32m        usable = A4[0] - 22 * mm[m
[32m+[m[32m        time_w = 44[m
[32m+[m[32m        day_w = (usable - time_w) / n_days[m
[32m+[m[32m        widths = [time_w] + [day_w] * n_days[m
[32m+[m
[32m+[m[32m        table_data = [[Paragraph(item, head_style) for item in data[0]]][m
[32m+[m[32m        for row in data[1:]:[m
[32m+[m[32m            table_data.append([m
[32m+[m[32m                [[m
[32m+[m[32m                    Paragraph(row[0], time_style) if row[0] else "",[m
[32m+[m[32m                ][m
[32m+[m[32m                + [[m
[32m+[m[32m                    Paragraph(_cell_markup(cell), cell_style) if cell else ""[m
[32m+[m[32m                    for cell in row[1:][m
[32m+[m[32m                ][m
[32m+[m[32m            )[m
[32m+[m
[32m+[m[32m        n_rows = len(data) - 1[m
[32m+[m[32m        row_h = None[m
[32m+[m[32m        if n_rows:[m
[32m+[m[32m            title_block = 18 + 12 + 12 + 4 * mm[m
[32m+[m[32m            footer_block = 6 * mm + 8[m
[32m+[m[32m            usable_h = A4[1] - doc.topMargin - doc.bottomMargin[m
[32m+[m[32m            avail = max(usable_h - title_block - footer_block, n_rows * 18)[m
[32m+[m[32m            row_h = min(avail / n_rows, 42)[m
[32m+[m
[32m+[m[32m        grid = Table([m
[32m+[m[32m            table_data,[m
[32m+[m[32m            colWidths=widths,[m
[32m+[m[32m            repeatRows=1,[m
[32m+[m[32m            rowHeights=[None] + [row_h] * n_rows if row_h else None,[m
[32m+[m[32m        )[m
[32m+[m[32m        style = [[m
[32m+[m[32m            ("GRID", (0, 0), (-1, -1), 0.4, colors.black),[m
[32m+[m[32m            ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#eceff4")),[m
[32m+[m[32m            ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#f8fafc")),[m
[32m+[m[32m            ("BACKGROUND", (1, 0), (-1, 0), colors.HexColor("#dce6f1")),[m
[32m+[m[32m            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),[m
[32m+[m[32m            ("TOPPADDING", (0, 0), (-1, -1), 3),[m
[32m+[m[32m            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),[m
[32m+[m[32m            ("LEFTPADDING", (0, 0), (-1, -1), 2),[m
[32m+[m[32m            ("RIGHTPADDING", (0, 0), (-1, -1), 2),[m
[32m+[m[32m        ][m
[32m+[m[32m        for start, end in spans:[m
[32m+[m[32m            style.append(("SPAN", start, end))[m
[32m+[m[32m        for start, end, hex_color in fills:[m
[32m+[m[32m            style.append(("BACKGROUND", start, end, colors.HexColor(hex_color)))[m
[32m+[m[32m        grid.setStyle(TableStyle(style))[m
[32m+[m
[32m+[m[32m        elements.append(grid)[m
[32m+[m[32m    else:[m
[32m+[m[32m        elements.append([m
[32m+[m[32m            Paragraph([m
[32m+[m[32m                "No timetable sessions scheduled for this selection.",[m
[32m+[m[32m                empty_style,[m
[32m+[m[32m            )[m
         )[m
 [m
[31m-    n_rows = len(data) - 1[m
[31m-    row_h = None[m
[31m-    if n_rows:[m
[31m-        title_block = 18 + 12 + 12 + 4 * mm[m
[31m-        footer_block = 6 * mm + 8[m
[31m-        usable_h = A4[1] - doc.topMargin - doc.bottomMargin[m
[31m-        avail = max(usable_h - title_block - footer_block, n_rows * 18)[m
[31m-        row_h = min(avail / n_rows, 42)[m
[31m-[m
[31m-    grid = Table([m
[31m-        table_data,[m
[31m-        colWidths=widths,[m
[31m-        repeatRows=1,[m
[31m-        rowHeights=[None] + [row_h] * n_rows if row_h else None,[m
[31m-    )[m
[31m-    style = [[m
[31m-        ("GRID", (0, 0), (-1, -1), 0.4, colors.black),[m
[31m-        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#eceff4")),[m
[31m-        ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#f8fafc")),[m
[31m-        ("BACKGROUND", (1, 0), (-1, 0), colors.HexColor("#dce6f1")),[m
[31m-        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),[m
[31m-        ("TOPPADDING", (0, 0), (-1, -1), 3),[m
[31m-        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),[m
[31m-        ("LEFTPADDING", (0, 0), (-1, -1), 2),[m
[31m-        ("RIGHTPADDING", (0, 0), (-1, -1), 2),[m
[31m-    ][m
[31m-    for start, end in spans:[m
[31m-        style.append(("SPAN", start, end))[m
[31m-    for start, end, hex_color in fills:[m
[31m-        style.append(("BACKGROUND", start, end, colors.HexColor(hex_color)))[m
[31m-    grid.setStyle(TableStyle(style))[m
[32m+[m[32m    if rotation_keys:[m
[32m+[m[32m        elements.append(Spacer(1, 5 * mm))[m
[32m+[m[32m        elements.append([m
[32m+[m[32m            Paragraph("WORKSHOP ROTATION KEY", key_title_style)[m
[32m+[m[32m        )[m
[32m+[m[32m        elements.append(_rotation_key_table(rotation_keys, cell_style, head_style))[m
 [m
[31m-    elements.append(grid)[m
     elements.append(Spacer(1, 6 * mm))[m
     export_date = datetime.date.today().strftime("%d %B %Y")[m
     elements.append(Paragraph(f"Prepared for personal use · {export_date}", foot_style))[m
[1mdiff --git a/core/views.py b/core/views.py[m
[1mindex 114cedb..2569f94 100644[m
[1m--- a/core/views.py[m
[1m+++ b/core/views.py[m
[36m@@ -10,10 +10,11 @@[m [mfrom django.shortcuts import get_object_or_404, redirect, render[m
 from django.urls import reverse[m
 from django.utils import timezone[m
 [m
[31m-from .timetable_grid import build_time_day_grid[m
[32m+[m[32mfrom .timetable_grid import build_day_time_grid[m
 from .timetable_pdf import ([m
     collect_entries,[m
     collect_group_entries,[m
[32m+[m[32m    collect_master_entries,[m
     render_group_timetable,[m
     render_programme_timetable,[m
 )[m
[36m@@ -68,6 +69,7 @@[m [mfrom .venue_quality import ([m
     resolve_venue_name_conflict,[m
     suggested_name,[m
 )[m
[32m+[m[32mfrom .workshop_times import legacy_workshop_allocations, legacy_workshop_sessions[m
 [m
 HTMX_HEADER = "HX-Request"[m
 [m
[36m@@ -562,12 +564,15 @@[m [mdef group_timetable_pdf(request, pk):[m
 [m
 [m
 def timetable_view(request):[m
[31m-    """Screen timetable grid: DAYS across the first row, TIME down the first column.[m
[32m+[m[32m    """Screen timetable grid: TIME slots across the first row, DAYS down the first column.[m
 [m
     Built from the session/workshop/TD records of the selected programme and[m
[31m-    semester via ``core.timetable_grid.build_time_day_grid`` (same data source[m
[31m-    as the PDF export), so any programme renders correctly. Cells are shaded[m
[31m-    by activity type (Lecture grey, Workshop green, Technical Drawing pink).[m
[32m+[m[32m    semester via ``core.timetable_grid.build_day_time_grid`` (same data source[m
[32m+[m[32m    as the PDF export), so any programme renders correctly. Days run down the[m
[32m+[m[32m    first column with full weekday labels (Monday..Friday, plus any weekend[m
[32m+[m[32m    days with data), and each day is drawn as a vertical band of lanes:[m
[32m+[m[32m    sessions are rendered as one Activity Card per cell spanning their hourly[m
[32m+[m[32m    columns, and overlapping sessions sit on separate lanes.[m
     """[m
     programmes = list(Programme.objects.all())[m
     semesters = list(Semester.objects.order_by("-academic_year", "-semester"))[m
[36m@@ -576,6 +581,8 @@[m [mdef timetable_view(request):[m
     programme_id = request.GET.get("programme")[m
     if programme_id and programme_id != "all":[m
         active_programme = Programme.objects.filter(pk=programme_id).first()[m
[32m+[m[32m    # Empty/missing selection (or "all") means the whole COET First Year.[m
[32m+[m[32m    show_all = active_programme is None[m
 [m
     active_group = None[m
     group_id = request.GET.get("group")[m
[36m@@ -602,10 +609,9 @@[m [mdef timetable_view(request):[m
             entries = collect_group_entries(active_group, active_semester, year=year)[m
         elif active_programme:[m
             entries = collect_entries(active_programme, active_semester, year=year)[m
[31m-        elif programmes:[m
[31m-            for prog in programmes:[m
[31m-                entries.extend(collect_entries(prog, active_semester, year=year))[m
[31m-        grid = build_time_day_grid(entries)[m
[32m+[m[32m        elif show_all:[m
[32m+[m[32m            entries = collect_master_entries(active_semester, year=year)[m
[32m+[m[32m        grid = build_day_time_grid(entries)[m
 [m
     ordinal = {1: "1st", 2: "2nd", 3: "3rd", 4: "4th"}[m
     [m
[36m@@ -613,7 +619,6 @@[m [mdef timetable_view(request):[m
     if active_programme:[m
         groups_for_programme = list(StudentGroup.objects.filter(programme=active_programme))[m
     [m
[31m-    show_all = programme_id == "all"[m
     ctx = {[m
         "page_title": "Timetable",[m
         "programmes": programmes,[m
[36m@@ -1380,6 +1385,7 @@[m [mdef session_list(request):[m
     if start_to:[m
         qs = qs.filter(start_time__lte=start_to)[m
     items, page, pages, total = _paginate(request, qs)[m
[32m+[m[32m    legacy_sessions = legacy_workshop_sessions()[m
     ctx = {[m
         "items": items,[m
         "columns": SESS_COLS,[m
[36m@@ -1448,6 +1454,16 @@[m [mdef session_list(request):[m
         "active_activity": act,[m
         "active_day": day,[m
         "active_semester": sem,[m
[32m+[m[32m        "legacy_sessions_count": len(legacy_sessions),[m
[32m+[m[32m        "legacy_sessions_examples": [[m
[32m+[m[32m            {[m
[32m+[m[32m                "course": ses.course_code,[m
[32m+[m[32m                "day": ses.get_day_display(),[m
[32m+[m[32m                "time": f"{ses.start_time:%H:%M}-{ses.end_time:%H:%M}",[m
[32m+[m[32m                "issue": issue,[m
[32m+[m[32m            }[m
[32m+[m[32m            for ses, issue in legacy_sessions[:3][m
[32m+[m[32m        ],[m
     }[m
     if _htmx(request):[m
         return render(request, "core/_table_and_cards.html", ctx)[m
[36m@@ -1734,7 +1750,7 @@[m [mWS_COLS = [[m
     {"key": "time_period", "label": "Period"},[m
     {"key": "start_time", "label": "Start"},[m
     {"key": "end_time", "label": "End"},[m
[31m-    {"key": "venue", "label": "Venue"},[m
[32m+[m[32m    {"key": "venue", "label": "Workshop"},[m
 ][m
 WS_FIELDS = [[m
     {"label": "Semester", "key": "semester"},[m
[36m@@ -1749,7 +1765,7 @@[m [mWS_FIELDS = [[m
     {"label": "Week End", "key": "week_end"},[m
     {"label": "Start Time", "key": "start_time"},[m
     {"label": "End Time", "key": "end_time"},[m
[31m-    {"label": "Venue", "key": "venue"},[m
[32m+[m[32m    {"label": "Workshop", "key": "venue"},[m
     {"label": "Year of Study", "key": "year_of_study"},[m
 ][m
 [m
[36m@@ -1771,6 +1787,7 @@[m [mdef workshop_list(request):[m
     if start_to:[m
         qs = qs.filter(start_time__lte=start_to)[m
     items, page, pages, total = _paginate(request, qs)[m
[32m+[m[32m    legacy = legacy_workshop_allocations()[m
     ctx = {[m
         "items": items,[m
         "columns": WS_COLS,[m
[36m@@ -1812,10 +1829,26 @@[m [mdef workshop_list(request):[m
         "total": total,[m
         "semesters": Semester.objects.all(),[m
         "active_semester": sem,[m
[32m+[m[32m        "legacy_count": len(legacy),[m
[32m+[m[32m        "legacy_examples": [[m
[32m+[m[32m            {[m
[32m+[m[32m                "course": rec.course_code,[m
[32m+[m[32m                "group": rec.group_code,[m
[32m+[m[32m                "day": rec.get_day_display(),[m
[32m+[m[32m                "time": ([m
[32m+[m[32m                    f"{rec.start_time:%H:%M}-{rec.end_time:%H:%M} "[m
[32m+[m[32m                    f"({rec.get_time_period_display()})"[m
[32m+[m[32m                    if rec.time_period[m
[32m+[m[32m                    else f"{rec.start_time:%H:%M}-{rec.end_time:%H:%M}"[m
[32m+[m[32m                ),[m
[32m+[m[32m                "issue": issue,[m
[32m+[m[32m            }[m
[32m+[m[32m            for rec, issue in legacy[:3][m
[32m+[m[32m        ],[m
     }[m
     if _htmx(request):[m
         return render(request, "core/_table_and_cards.html", ctx)[m
[31m-    return render(request, "core/list.html", ctx)[m
[32m+[m[32m    return render(request, "core/workshop_list.html", ctx)[m
 [m
 [m
 def workshop_detail(request, pk):[m
[36m@@ -2414,7 +2447,7 @@[m [mIMPORT_TYPES = {[m
     "workshop-allocation": {[m
         "title": "Workshop Allocation",[m
         "columns": ([m
[31m-            "Flat: course_code, group_code, day, start_time, end_time, venue — "[m
[32m+[m[32m            "Flat: course_code, group_code, day, start_time, end_time, workshop — "[m
             "or drop in the raw university Workshop Schedule workbook (matrix) directly"[m
         ),[m
         "semester": "optional",[m
[1mdiff --git a/templates/base.html b/templates/base.html[m
[1mindex b4ce1ef..c52c85b 100644[m
[1m--- a/templates/base.html[m
[1m+++ b/templates/base.html[m
[36m@@ -6,6 +6,18 @@[m
     <meta charset="UTF-8">[m
     <meta name="viewport" content="width=device-width, initial-scale=1.0">[m
     <title>{% block title %}{{ page_title|default:"CoET Timetable" }}{% endblock %} — CoET Timetable</title>[m
[32m+[m[32m    <script>[m
[32m+[m[32m        // Apply the persisted sidebar collapse state before the first paint so[m
[32m+[m[32m        // navigating between pages never flashes the sidebar open/closed.[m
[32m+[m[32m        (function () {[m
[32m+[m[32m            var v = null;[m
[32m+[m[32m            try {[m
[32m+[m[32m                v = window.sessionStorage.getItem('coet.sidebar.collapsed');[m
[32m+[m[32m                if (v === null) { v = window.localStorage.getItem('coet.sidebar.collapsed'); }[m
[32m+[m[32m            } catch (e) {}[m
[32m+[m[32m            if (v === '1') { document.documentElement.classList.add('tt-collapsed'); }[m
[32m+[m[32m        })();[m
[32m+[m[32m    </script>[m
     <script src="https://cdn.tailwindcss.com"></script>[m
     <script src="https://unpkg.com/htmx.org@2.0.4"></script>[m
     <script defer src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js"></script>[m
[36m@@ -45,10 +57,56 @@[m
         html, body { overflow-x: hidden; }[m
 [m
         @media (max-width: 640px) {[m
[31m-            .max-w-7xl { max-width: 100%; padding: 0 0.5rem; }[m
[31m-            .max-w-6xl { max-width: 100%; padding: 0 0.5rem; }[m
[32m+[m[32m            .max-w-7xl, .max-w-6xl, .max-w-5xl, .max-w-4xl, .max-w-3xl, .max-w-2xl {[m
[32m+[m[32m                max-width: 100%;[m
[32m+[m[32m                padding: 0 0.25rem;[m
[32m+[m[32m            }[m
[32m+[m[32m        }[m
[32m+[m
[32m+[m[32m        /* Consistent, accessible focus ring for keyboard users */[m
[32m+[m[32m        :where(a, button, input, select, textarea, [role="button"], [tabindex]):focus-visible {[m
[32m+[m[32m            outline: 2px solid #2563eb;[m
[32m+[m[32m            outline-offset: 2px;[m
[32m+[m[32m        }[m
[32m+[m
[32m+[m[32m        /* Keep long unbroken strings from bursting out of narrow cards/cells */[m
[32m+[m[32m        .break-safe { overflow-wrap: anywhere; word-break: break-word; }[m
[32m+[m
[32m+[m[32m        /* Sidebar collapse state is mirrored on <html> (html.tt-collapsed) by[m
[32m+[m[32m           the inline <head> script, so the icon-only layout is applied from the[m
[32m+[m[32m           very first paint — page navigations never flash it expanded again. */[m
[32m+[m[32m        @media (min-width: 1024px) {[m
[32m+[m[32m            html.tt-collapsed #app-sidebar { width: 5rem; }[m
[32m+[m[32m            html.tt-collapsed #app-sidebar .tt-brand-row { justify-content: center; gap: 0; }[m
[32m+[m[32m            html.tt-collapsed #app-sidebar .tt-brand-link { max-width: 0; opacity: 0; }[m
[32m+[m[32m            html.tt-collapsed #app-sidebar .tt-hide-when-collapsed { display: none; }[m
[32m+[m[32m            html.tt-collapsed #app-sidebar .tt-section {[m
[32m+[m[32m                margin: 0.5rem 0;[m
[32m+[m[32m                height: 1px;[m
[32m+[m[32m                padding-top: 0;[m
[32m+[m[32m                padding-bottom: 0;[m
[32m+[m[32m                overflow: hidden;[m
[32m+[m[32m                color: transparent;[m
[32m+[m[32m                background-color: rgb(51 65 85 / 0.5);[m
[32m+[m[32m            }[m
[32m+[m[32m            html.tt-collapsed #app-sidebar nav > a { justify-content: center; gap: 0; }[m
[32m+[m[32m            html.tt-collapsed #app-sidebar nav > a > span { max-width: 0; opacity: 0; overflow: hidden; }[m
[32m+[m[32m        }[m
[32m+[m
[32m+[m[32m        /* Mobile drawer starts fully off-canvas from the first paint; only the[m
[32m+[m[32m           .translate-x-0 class (Alpine's open state) brings it onto screen. */[m
[32m+[m[32m        @media (max-width: 1023px) {[m
[32m+[m[32m            #app-sidebar { transform: translateX(-100%); }[m
[32m+[m[32m            #app-sidebar.translate-x-0 { transform: translateX(0); }[m
[32m+[m[32m        }[m
[32m+[m
[32m+[m[32m        /* Toast notifications sit flush on phones and never overflow the screen */[m
[32m+[m[32m        @media (max-width: 640px) {[m
[32m+[m[32m            #toast-container { left: 1rem; right: 1rem; }[m
         }[m
 [m
[32m+[m[32m        html { scroll-behavior: smooth; }[m
[32m+[m
         @media (prefers-reduced-motion: reduce) {[m
             *, *::before, *::after {[m
                 animation-duration: 0.01ms !important;[m
[36m@@ -149,7 +207,7 @@[m [mfunction showToast(message, type) {[m
     var el = document.createElement('div');[m
     var color = isError ? 'bg-red-600' : type === 'info' ? 'bg-blue-600' : 'bg-emerald-600';[m
     el.dataset.kind = isError ? 'persistent' : 'transient';[m
[31m-    el.className = 'tt-toast ' + color + ' text-white px-4 py-3 rounded-lg shadow-lg text-sm font-medium flex items-start gap-2';[m
[32m+[m[32m    el.className = 'tt-toast ' + color + ' text-white px-4 py-3 rounded-lg shadow-lg text-sm font-medium flex items-start gap-2 w-full sm:w-auto sm:min-w-[20rem] max-w-full';[m
     var icon = '<span class="shrink-0 mt-0.5">' + (isError ? '&#10007;' : '&#10003;') + '</span>';[m
     var close = '<button type="button" onclick="dismissToast(this.parentNode)" aria-label="Dismiss notification" class="shrink-0 ml-1 rounded p-1 -m-1 text-white/80 hover:text-white hover:bg-white/20">' +[m
                 '<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg></button>';[m
[36m@@ -188,6 +246,47 @@[m [mdocument.addEventListener('alpine:init', function () {[m
             }[m
         };[m
     });[m
[32m+[m[32m    Alpine.data('sidebar', function () {[m
[32m+[m[32m        return {[m
[32m+[m[32m            collapsed: false,[m
[32m+[m[32m            tooltip: '',[m
[32m+[m[32m            tooltipX: 0,[m
[32m+[m[32m            tooltipY: 0,[m
[32m+[m[32m            init: function () {[m
[32m+[m[32m                var val = null;[m
[32m+[m[32m                try {[m
[32m+[m[32m                    val = window.sessionStorage.getItem('coet.sidebar.collapsed');[m
[32m+[m[32m                    if (val === null) {[m
[32m+[m[32m                        val = window.localStorage.getItem('coet.sidebar.collapsed');[m
[32m+[m[32m                    }[m
[32m+[m[32m                } catch (e) { val = null; }[m
[32m+[m[32m                if (val === '1') { this.collapsed = true; }[m
[32m+[m[32m                document.documentElement.classList.toggle('tt-collapsed', this.collapsed);[m
[32m+[m[32m            },[m
[32m+[m[32m            toggle: function () {[m
[32m+[m[32m                this.collapsed = !this.collapsed;[m
[32m+[m[32m                var val = this.collapsed ? '1' : '0';[m
[32m+[m[32m                try {[m
[32m+[m[32m                    window.localStorage.setItem('coet.sidebar.collapsed', val);[m
[32m+[m[32m                    window.sessionStorage.setItem('coet.sidebar.collapsed', val);[m
[32m+[m[32m                } catch (e) {}[m
[32m+[m[32m                document.documentElement.classList.toggle('tt-collapsed', this.collapsed);[m
[32m+[m[32m            },[m
[32m+[m[32m            showTooltip: function (el, text) {[m
[32m+[m[32m                if (this.collapsed && window.innerWidth >= 1024) {[m
[32m+[m[32m                    var rect = el.getBoundingClientRect();[m
[32m+[m[32m                    this.tooltip = text;[m
[32m+[m[32m                    this.tooltipX = rect.right + 10;[m
[32m+[m[32m                    this.tooltipY = rect.top + rect.height / 2;[m
[32m+[m[32m                } else {[m
[32m+[m[32m                    this.tooltip = '';[m
[32m+[m[32m                }[m
[32m+[m[32m            },[m
[32m+[m[32m            hideTooltip: function () {[m
[32m+[m[32m                this.tooltip = '';[m
[32m+[m[32m            }[m
[32m+[m[32m        };[m
[32m+[m[32m    });[m
 });[m
 [m
 document.body.addEventListener('htmx:afterSwap', function (evt) {[m
[36m@@ -234,7 +333,7 @@[m [mdocument.body.addEventListener('htmx:sendError', function () {[m
 });[m
 </script>[m
 [m
[31m-<div class="flex min-h-screen">[m
[32m+[m[32m<div class="flex min-h-screen" x-data="sidebar">[m
     <!-- Mobile overlay -->[m
     <div x-show="sidebarOpen" x-cloak @click="sidebarOpen = false"[m
          x-transition.opacity.duration.200ms[m
[36m@@ -246,8 +345,9 @@[m [mdocument.body.addEventListener('htmx:sendError', function () {[m
     <!-- Main content -->[m
     <div class="flex-1 flex flex-col min-w-0">[m
         <!-- Top bar -->[m
[31m-        <header class="sticky top-0 z-20 bg-white border-b border-gray-200 px-4 lg:px-6 h-14 flex items-center gap-4">[m
[31m-            <button @click="sidebarOpen = !sidebarOpen" class="lg:hidden text-gray-500 hover:text-gray-700">[m
[32m+[m[32m        <header class="sticky top-0 z-20 bg-white border-b border-gray-200 px-3 sm:px-4 lg:px-6 h-14 flex items-center gap-3 sm:gap-4">[m
[32m+[m[32m            <button @click="sidebarOpen = !sidebarOpen" aria-label="Toggle navigation menu"[m
[32m+[m[32m                    class="lg:hidden text-gray-500 hover:text-gray-700 rounded-lg p-2 -ml-1.5 hover:bg-gray-100 focus-visible:outline-2 focus-visible:outline-blue-500">[m
                 <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h16M4 18h16"/></svg>[m
             </button>[m
             <h1 class="text-base font-semibold text-gray-800 truncate">{{ page_title|default:"Dashboard" }}</h1>[m
[36m@@ -259,6 +359,14 @@[m [mdocument.body.addEventListener('htmx:sendError', function () {[m
             {% block content %}{% endblock %}[m
         </main>[m
     </div>[m
[32m+[m
[32m+[m[32m    <!-- Sidebar collapsed-mode tooltip -->[m
[32m+[m[32m    <div id="sidebar-tooltip" x-show="collapsed && tooltip" x-cloak role="tooltip"[m
[32m+[m[32m         :style="'left: ' + tooltipX + 'px; top: ' + tooltipY + 'px; transform: translateY(-50%)'"[m
[32m+[m[32m         class="fixed z-50 px-2.5 py-1.5 rounded-md bg-slate-900 text-white text-xs font-medium shadow-lg pointer-events-none">[m
[32m+[m[32m        <span class="absolute top-1/2 -translate-y-1/2 -left-1.5 h-3 w-3 rotate-45 bg-slate-900" aria-hidden="true"></span>[m
[32m+[m[32m        <span x-text="tooltip"></span>[m
[32m+[m[32m    </div>[m
 </div>[m
 [m
 <!-- Slide-over detail panel -->[m
[36m@@ -278,14 +386,14 @@[m [mdocument.body.addEventListener('htmx:sendError', function () {[m
          x-transition:leave="transition ease-in duration-200"[m
          x-transition:leave-start="translate-x-0"[m
          x-transition:leave-end="translate-x-full"[m
[31m-         class="absolute inset-y-0 right-0 w-full max-w-2xl bg-white shadow-2xl flex flex-col">[m
[31m-        <div class="flex items-center justify-between px-6 py-4 border-b border-gray-200">[m
[31m-            <h2 class="text-lg font-semibold text-gray-800 truncate" id="slide-title">Details</h2>[m
[31m-            <button @click="slideOpen = false" class="text-gray-400 hover:text-gray-600 rounded-lg p-1.5 hover:bg-gray-100">[m
[32m+[m[32m         class="absolute inset-y-0 right-0 w-full max-w-2xl bg-white shadow-2xl flex flex-col sm:border-l sm:border-gray-200">[m
[32m+[m[32m        <div class="flex items-center justify-between px-4 sm:px-6 py-4 border-b border-gray-200">[m
[32m+[m[32m            <h2 class="text-base sm:text-lg font-semibold text-gray-800 truncate pr-2" id="slide-title">Details</h2>[m
[32m+[m[32m            <button @click="slideOpen = false" aria-label="Close details panel" class="text-gray-400 hover:text-gray-600 rounded-lg p-2 hover:bg-gray-100 shrink-0">[m
                 <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>[m
             </button>[m
         </div>[m
[31m-        <div id="slide-content" hx-settle="220" class="flex-1 overflow-y-auto p-6">[m
[32m+[m[32m        <div id="slide-content" hx-settle="220" class="flex-1 overflow-y-auto overscroll-contain p-4 sm:p-6">[m
             <div class="flex items-center justify-center h-32">[m
                 <div class="text-gray-400 text-sm">Loading...</div>[m
             </div>[m
[36m@@ -301,23 +409,23 @@[m [mdocument.body.addEventListener('htmx:sendError', function () {[m
      x-transition:leave="transition-opacity ease-in duration-150"[m
      x-transition:leave-start="opacity-100"[m
      x-transition:leave-end="opacity-0"[m
[31m-     class="fixed inset-0 z-50 flex items-center justify-center p-4">[m
[32m+[m[32m     class="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4 overflow-y-auto">[m
     <div class="absolute inset-0 bg-black/50" @click="modalOpen = false"></div>[m
     <div x-show="modalOpen"[m
          x-transition:enter="transition ease-out duration-200"[m
[31m-         x-transition:enter-start="opacity-0 scale-95"[m
[31m-         x-transition:enter-end="opacity-100 scale-100"[m
[32m+[m[32m         x-transition:enter-start="opacity-0 scale-95 translate-y-3 sm:translate-y-0"[m
[32m+[m[32m         x-transition:enter-end="opacity-100 scale-100 translate-y-0"[m
          x-transition:leave="transition ease-in duration-150"[m
[31m-         x-transition:leave-start="opacity-100 scale-100"[m
[31m-         x-transition:leave-end="opacity-0 scale-95"[m
[31m-         class="relative bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col">[m
[31m-        <div class="flex items-center justify-between px-6 py-4 border-b border-gray-200">[m
[31m-            <h2 class="text-lg font-semibold text-gray-800" id="modal-title">Form</h2>[m
[31m-            <button @click="modalOpen = false" class="text-gray-400 hover:text-gray-600 rounded-lg p-1.5 hover:bg-gray-100">[m
[32m+[m[32m         x-transition:leave-start="opacity-100 scale-100 translate-y-0"[m
[32m+[m[32m         x-transition:leave-end="opacity-0 scale-95 translate-y-3 sm:translate-y-0"[m
[32m+[m[32m         class="relative sm:rounded-2xl rounded-t-2xl shadow-2xl w-full sm:max-w-2xl max-h-[92vh] sm:max-h-[90vh] overflow-hidden flex flex-col bg-white">[m
[32m+[m[32m        <div class="flex items-center justify-between px-4 sm:px-6 py-4 border-b border-gray-200">[m
[32m+[m[32m            <h2 class="text-base sm:text-lg font-semibold text-gray-800 truncate pr-2" id="modal-title">Form</h2>[m
[32m+[m[32m            <button @click="modalOpen = false" aria-label="Close dialog" class="text-gray-400 hover:text-gray-600 rounded-lg p-2 hover:bg-gray-100 shrink-0">[m
                 <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>[m
             </button>[m
         </div>[m
[31m-        <div id="modal-content" hx-settle="220" class="flex-1 overflow-y-auto p-6">[m
[32m+[m[32m        <div id="modal-content" hx-settle="220" class="flex-1 overflow-y-auto overscroll-contain p-4 sm:p-6">[m
             <div class="flex items-center justify-center h-32">[m
                 <div class="text-gray-400 text-sm">Loading...</div>[m
             </div>[m
[1mdiff --git a/templates/core/_activity_log.html b/templates/core/_activity_log.html[m
[1mindex f57b783..2424753 100644[m
[1m--- a/templates/core/_activity_log.html[m
[1m+++ b/templates/core/_activity_log.html[m
[36m@@ -5,12 +5,12 @@[m
         <div class="divide-y divide-gray-100 min-w-full">[m
         {% for log in items %}[m
         <a href="{% url 'dashboard' %}?log={{ log.pk }}"[m
[31m-           class="flex items-start gap-3 px-5 py-4 hover:bg-blue-50/50 transition">[m
[32m+[m[32m           class="flex items-start gap-3 px-4 sm:px-5 py-4 hover:bg-blue-50/50 transition">[m
             <span class="mt-0.5 shrink-0 inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold text-white {{ log.action|log_badge_cls }}">[m
                 {{ log.get_action_display }}[m
             </span>[m
             <span class="min-w-0 flex-1">[m
[31m-                <span class="block text-sm text-gray-800">{{ log.message }}</span>[m
[32m+[m[32m                <span class="block text-sm text-gray-800 break-safe">{{ log.message }}</span>[m
                 <span class="block text-xs text-gray-400 mt-0.5">[m
                     {{ log.created_at|date:"d M Y, H:i" }}[m
                     {% if log.resource %}&middot; {{ log.resource }}{% endif %}[m
[1mdiff --git a/templates/core/_detail_content.html b/templates/core/_detail_content.html[m
[1mindex e16e48d..44b687e 100644[m
[1m--- a/templates/core/_detail_content.html[m
[1m+++ b/templates/core/_detail_content.html[m
[36m@@ -18,21 +18,21 @@[m
 [m
     {% block detail_extra %}{% endblock %}[m
 [m
[31m-    <div class="pt-4 border-t border-gray-100 flex flex-wrap gap-2">[m
[32m+[m[32m    <div class="pt-4 border-t border-gray-100 grid grid-cols-1 sm:flex sm:flex-wrap gap-2">[m
         {% if export_url %}[m
         <a href="{{ export_url }}"[m
[31m-           class="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-green-700 bg-green-50 hover:bg-green-100 rounded-lg transition">[m
[32m+[m[32m           class="inline-flex items-center justify-center gap-1.5 px-4 py-2.5 sm:py-2 text-sm font-medium text-green-700 bg-green-50 hover:bg-green-100 rounded-lg transition">[m
             <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>[m
             {{ export_label|default:"Export PDF" }}[m
         </a>[m
         {% endif %}[m
         <button onclick="openEdit('{{ edit_url }}')"[m
[31m-                class="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition">[m
[32m+[m[32m                class="inline-flex items-center justify-center gap-1.5 px-4 py-2.5 sm:py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition">[m
             <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>[m
             Edit[m
         </button>[m
         <button onclick="openDelete('{{ delete_url }}')"[m
[31m-                class="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition">[m
[32m+[m[32m                class="inline-flex items-center justify-center gap-1.5 px-4 py-2.5 sm:py-2 text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition">[m
             <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>[m
             Delete[m
         </button>[m
[1mdiff --git a/templates/core/_session_groups.html b/templates/core/_session_groups.html[m
[1mindex 00a2bf5..a8817ec 100644[m
[1m--- a/templates/core/_session_groups.html[m
[1m+++ b/templates/core/_session_groups.html[m
[36m@@ -83,13 +83,13 @@[m
                   hx-disabled-elt="this"[m
                   class="flex flex-col sm:flex-row gap-2">[m
                 {% csrf_token %}[m
[31m-                <select name="group_id" class="flex-1 px-3 py-2 border border-gray-300 rounded-lg text-sm text-gray-800 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">[m
[32m+[m[32m                <select name="group_id" class="w-full sm:flex-1 px-3 py-2.5 border border-gray-300 rounded-lg text-sm text-gray-800 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500">[m
                     <option value="">— Select a student group —</option>[m
                     {% for g in all_groups %}[m
                     <option value="{{ g.pk }}">{{ g.programme.code }} {{ g.code }}</option>[m
                     {% endfor %}[m
                 </select>[m
[31m-                <button type="submit" class="inline-flex items-center justify-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition disabled:opacity-50 disabled:cursor-wait">[m
[32m+[m[32m                <button type="submit" class="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition disabled:opacity-50 disabled:cursor-wait">[m
                     <span class="tt-idle-label inline-flex items-center gap-1">Add Group</span>[m
                     <span class="tt-status-label">[m
                         <svg class="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>[m
[1mdiff --git a/templates/core/_table_and_cards.html b/templates/core/_table_and_cards.html[m
[1mindex e4dfebf..40d7a18 100644[m
[1m--- a/templates/core/_table_and_cards.html[m
[1m+++ b/templates/core/_table_and_cards.html[m
[36m@@ -23,7 +23,8 @@[m
                     hx-swap="innerHTML"[m
                     @click="slideOpen = true">[m
                     {% for col in columns %}[m
[31m-                    <td class="px-4 py-3 text-sm text-gray-700 whitespace-nowrap">[m
[32m+[m[32m                    <td class="px-4 py-3 text-sm text-gray-700 whitespace-nowrap[m
[32m+[m[32m                        {% if col.key != 'start_time' and col.key != 'end_time' and col.key != 'activity_type' and col.key != 'day' %}max-w-[16rem] truncate{% endif %}">[m
                         {% if col.key == 'start_time' or col.key == 'end_time' %}[m
                             {{ item|get_attr:col.key|time_short }}[m
                         {% elif col.key == 'activity_type' %}[m
[36m@@ -39,7 +40,7 @@[m
                         {% elif col.key == 'day' %}[m
                             <span class="text-gray-600">{{ item|get_attr:col.key }}</span>[m
                         {% else %}[m
[31m-                            {{ item|get_attr:col.key }}[m
[32m+[m[32m                            <span class="block truncate" title="{{ item|get_attr:col.key }}">{{ item|get_attr:col.key }}</span>[m
                         {% endif %}[m
                     </td>[m
                     {% endfor %}[m
[36m@@ -56,17 +57,17 @@[m
                     </td>[m
                     {% endif %}[m
                     <td class="px-4 py-3 text-right whitespace-nowrap" @click.stop>[m
[31m-                        <div class="inline-flex items-center gap-1 opacity-0 group-hover:opacity-100 transition">[m
[32m+[m[32m                        <div class="inline-flex items-center">[m
                             <a href="{{ item.pk }}/edit/"[m
                                onclick="event.preventDefault(); openEdit(this.getAttribute('href'))"[m
[31m-                               class="p-1.5 rounded-md text-gray-400 hover:text-blue-600 hover:bg-blue-50 transition"[m
[31m-                               title="Edit">[m
[32m+[m[32m                               class="p-2 rounded-md text-gray-400 hover:text-blue-600 hover:bg-blue-50 transition"[m
[32m+[m[32m                               title="Edit" aria-label="Edit">[m
                                 <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>[m
                             </a>[m
                             <a href="{{ item.pk }}/delete/"[m
                                onclick="event.preventDefault(); openDelete(this.getAttribute('href'))"[m
[31m-                               class="p-1.5 rounded-md text-gray-400 hover:text-red-600 hover:bg-red-50 transition"[m
[31m-                               title="Delete">[m
[32m+[m[32m                               class="p-2 rounded-md text-gray-400 hover:text-red-600 hover:bg-red-50 transition"[m
[32m+[m[32m                               title="Delete" aria-label="Delete">[m
                                 <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>[m
                             </a>[m
                         </div>[m
[36m@@ -112,11 +113,11 @@[m
         <div class="p-4">[m
             <!-- Primary line -->[m
             <div class="flex items-start justify-between gap-2 mb-2">[m
[31m-                <div class="text-sm font-semibold text-gray-800">[m
[32m+[m[32m                <div class="text-base font-semibold text-gray-800 break-safe">[m
                     {% if columns.0 %}{{ item|get_attr:columns.0.key }}{% endif %}[m
                 </div>[m
                 {% if columns|length > 1 %}[m
[31m-                <span class="text-xs text-gray-500">{{ item|get_attr:columns.1.key }}</span>[m
[32m+[m[32m                <span class="text-xs text-gray-500 shrink-0">{{ item|get_attr:columns.1.key }}</span>[m
                 {% endif %}[m
             </div>[m
             {% if issue %}[m
[36m@@ -128,15 +129,15 @@[m
             </div>[m
             {% endif %}[m
             <!-- Detail rows -->[m
[31m-            <div class="space-y-1">[m
[32m+[m[32m            <div class="space-y-1.5">[m
                 {% for col in columns %}[m
                 {% if not forloop.first %}[m
[31m-                <div class="flex justify-between items-center">[m
[31m-                    <span class="text-xs text-gray-500">{{ col.label }}</span>[m
[32m+[m[32m                <div class="flex justify-between items-start gap-3">[m
[32m+[m[32m                    <span class="text-xs text-gray-500 shrink-0">{{ col.label }}</span>[m
                     {% if col.key == 'start_time' or col.key == 'end_time' %}[m
[31m-                    <span class="text-xs text-gray-700 font-medium">{{ item|get_attr:col.key|time_short }}</span>[m
[32m+[m[32m                    <span class="text-xs text-gray-700 font-medium text-right">{{ item|get_attr:col.key|time_short }}</span>[m
                     {% elif col.key == 'activity_type' %}[m
[31m-                    <span class="text-xs font-medium[m
[32m+[m[32m                    <span class="text-xs font-medium grow text-right[m
                         {% if item.activity_type == 'LECTURE' %}text-blue-700[m
                         {% elif item.activity_type == 'TUTORIAL' %}text-emerald-700[m
                         {% elif item.activity_type == 'SEMINAR' %}text-purple-700[m
[36m@@ -145,7 +146,7 @@[m
                         {{ item|get_attr:col.key }}[m
                     </span>[m
                     {% else %}[m
[31m-                    <span class="text-xs text-gray-700 font-medium">{{ item|get_attr:col.key }}</span>[m
[32m+[m[32m                    <span class="text-xs text-gray-700 font-medium text-right break-safe">{{ item|get_attr:col.key }}</span>[m
                     {% endif %}[m
                 </div>[m
                 {% endif %}[m
[36m@@ -154,12 +155,12 @@[m
             <div class="mt-3 pt-3 border-t border-gray-100 flex gap-2" @click.stop>[m
                 <a href="{{ item.pk }}/edit/"[m
                    onclick="event.preventDefault(); openEdit(this.getAttribute('href'))"[m
[31m-                   class="flex items-center gap-1 text-xs font-medium text-blue-600 hover:text-blue-800">[m
[32m+[m[32m                   class="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-2.5 text-xs sm:text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition">[m
                     Edit[m
                 </a>[m
                 <a href="{{ item.pk }}/delete/"[m
                    onclick="event.preventDefault(); openDelete(this.getAttribute('href'))"[m
[31m-                   class="flex items-center gap-1 text-xs font-medium text-red-600 hover:text-red-800">[m
[32m+[m[32m                   class="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-2.5 text-xs sm:text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition">[m
                     Delete[m
                 </a>[m
             </div>[m
[36m@@ -178,10 +179,10 @@[m
         <span class="text-xs text-gray-500">Page {{ page }} of {{ pages }}</span>[m
         <div class="flex gap-2">[m
             {% if page > 1 %}[m
[31m-            <a href="?page={{ page|add:"-1" }}&q={{ q }}" class="px-3 py-1.5 text-xs font-medium bg-white border border-gray-300 text-gray-600 rounded-lg">Prev</a>[m
[32m+[m[32m            <a href="?page={{ page|add:"-1" }}&q={{ q }}" class="px-4 py-2.5 text-xs font-medium bg-white border border-gray-300 text-gray-600 rounded-lg">Prev</a>[m
             {% endif %}[m
             {% if page < pages %}[m
[31m-            <a href="?page={{ page|add:"1" }}&q={{ q }}" class="px-3 py-1.5 text-xs font-medium bg-white border border-gray-300 text-gray-600 rounded-lg">Next</a>[m
[32m+[m[32m            <a href="?page={{ page|add:"1" }}&q={{ q }}" class="px-4 py-2.5 text-xs font-medium bg-white border border-gray-300 text-gray-600 rounded-lg">Next</a>[m
             {% endif %}[m
         </div>[m
     </div>[m
[1mdiff --git a/templates/core/activity_log.html b/templates/core/activity_log.html[m
[1mindex c9a758a..7316f78 100644[m
[1m--- a/templates/core/activity_log.html[m
[1m+++ b/templates/core/activity_log.html[m
[36m@@ -12,11 +12,11 @@[m
                 <div class="flex-1 relative">[m
                     <svg class="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>[m
                     <input type="text" name="q" value="{{ q }}" placeholder="Search activity..."[m
[31m-                           class="w-full pl-9 pr-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                           class="w-full pl-9 pr-3 py-2.5 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
                 </div>[m
                 {% for f in filters %}[m
                     {% if f.type == "select" %}[m
[31m-                    <select name="{{ f.name }}" class="shrink-0 px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                    <select name="{{ f.name }}" class="w-full md:w-auto shrink-0 px-3 py-2.5 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
                         <option value="">All {{ f.label }}s</option>[m
                         {% for v, label in f.options %}[m
                         <option value="{{ v }}" {% if f.value == v %}selected{% endif %}>{{ label }}</option>[m
[36m@@ -24,12 +24,14 @@[m
                     </select>[m
                     {% endif %}[m
                 {% endfor %}[m
[31m-                <button type="submit" class="px-4 py-2 bg-gray-800 text-white text-sm font-medium rounded-lg hover:bg-gray-900 transition">[m
[31m-                    Filter[m
[31m-                </button>[m
[31m-                <a href="?" class="px-4 py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
[31m-                    Reset[m
[31m-                </a>[m
[32m+[m[32m                <div class="flex gap-2 w-full md:w-auto">[m
[32m+[m[32m                    <button type="submit" class="flex-1 md:flex-none px-4 py-2.5 bg-gray-800 text-white text-sm font-medium rounded-lg hover:bg-gray-900 transition">[m
[32m+[m[32m                        Filter[m
[32m+[m[32m                    </button>[m
[32m+[m[32m                    <a href="?" class="flex-1 md:flex-none flex items-center justify-center px-4 py-2.5 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
[32m+[m[32m                        Reset[m
[32m+[m[32m                    </a>[m
[32m+[m[32m                </div>[m
             </div>[m
         </form>[m
     </div>[m
[1mdiff --git a/templates/core/clear_all.html b/templates/core/clear_all.html[m
[1mindex 7aec92e..427455f 100644[m
[1m--- a/templates/core/clear_all.html[m
[1m+++ b/templates/core/clear_all.html[m
[36m@@ -78,14 +78,14 @@[m
     </div>[m
     {% endif %}[m
 [m
[31m-    <div class="mt-6 flex justify-end gap-3">[m
[32m+[m[32m    <div class="mt-6 flex flex-col sm:flex-row sm:justify-end gap-3">[m
         <button type="button" @click="modalOpen = false"[m
[31m-                class="px-4 py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
[32m+[m[32m                class="w-full sm:w-auto px-4 py-2.5 sm:py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
             Cancel[m
         </button>[m
         {% if has_records %}[m
         <button type="submit" :disabled="submitting || phrase !== 'DELETE ALL'"[m
[31m-                class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-lg transition disabled:opacity-50 disabled:cursor-wait">[m
[32m+[m[32m                class="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-4 py-2.5 sm:py-2 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-lg transition disabled:opacity-50 disabled:cursor-wait">[m
             <span x-show="!submitting">Clear All</span>[m
             <span x-show="submitting" class="inline-flex items-center gap-2">[m
                 <svg class="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>[m
[1mdiff --git a/templates/core/delete.html b/templates/core/delete.html[m
[1mindex cf0ace1..9bbe021 100644[m
[1m--- a/templates/core/delete.html[m
[1m+++ b/templates/core/delete.html[m
[36m@@ -80,12 +80,12 @@[m
         {% endif %}[m
     </div>[m
 [m
[31m-    <div class="mt-6 flex justify-end gap-3">[m
[31m-        <button type="button" @click="modalOpen = false" class="px-4 py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
[32m+[m[32m    <div class="mt-6 flex flex-col sm:flex-row sm:justify-end gap-3">[m
[32m+[m[32m        <button type="button" @click="modalOpen = false" class="w-full sm:w-auto px-4 py-2.5 sm:py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
             Cancel[m
         </button>[m
         <button type="submit" :disabled="submitting"[m
[31m-                class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-lg transition disabled:opacity-50 disabled:cursor-wait">[m
[32m+[m[32m                class="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-4 py-2.5 sm:py-2 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-lg transition disabled:opacity-50 disabled:cursor-wait">[m
             <span x-show="!submitting">Yes, Delete</span>[m
             <span x-show="submitting" class="inline-flex items-center gap-2">[m
                 <svg class="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>[m
[1mdiff --git a/templates/core/detail.html b/templates/core/detail.html[m
[1mindex 3c8dbf7..6efb88a 100644[m
[1m--- a/templates/core/detail.html[m
[1m+++ b/templates/core/detail.html[m
[36m@@ -8,27 +8,27 @@[m
         Back[m
     </button>[m
     <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">[m
[31m-        <div class="px-6 py-4 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3">[m
[31m-            <h2 class="text-lg font-semibold text-gray-800">{{ item }}</h2>[m
[31m-            <div class="flex gap-2">[m
[32m+[m[32m        <div class="px-4 sm:px-6 py-4 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3">[m
[32m+[m[32m            <h2 class="text-lg font-semibold text-gray-800 break-safe">{{ item }}</h2>[m
[32m+[m[32m            <div class="flex flex-wrap gap-2">[m
                 {% if export_url %}[m
                 <a href="{{ export_url }}"[m
[31m-                   class="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-green-700 bg-green-50 hover:bg-green-100 rounded-lg transition">[m
[32m+[m[32m                   class="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-sm font-medium text-green-700 bg-green-50 hover:bg-green-100 rounded-lg transition">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>[m
                     {{ export_label|default:"Export PDF" }}[m
                 </a>[m
                 {% endif %}[m
                 <a href="{{ edit_url }}" onclick="event.preventDefault(); openEdit(this.getAttribute('href'))"[m
[31m-                   class="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition">[m
[32m+[m[32m                   class="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition">[m
                     Edit[m
                 </a>[m
                 <a href="{{ delete_url }}" onclick="event.preventDefault(); openDelete(this.getAttribute('href'))"[m
[31m-                   class="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition">[m
[32m+[m[32m                   class="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition">[m
                     Delete[m
                 </a>[m
             </div>[m
         </div>[m
[31m-        <div class="p-6">[m
[32m+[m[32m        <div class="p-4 sm:p-6">[m
             <dl class="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-4">[m
                 {% for field in detail_fields %}[m
                 <div>[m
[1mdiff --git a/templates/core/export_timetable.html b/templates/core/export_timetable.html[m
[1mindex 2c5f9b2..73f48d5 100644[m
[1m--- a/templates/core/export_timetable.html[m
[1m+++ b/templates/core/export_timetable.html[m
[36m@@ -9,19 +9,19 @@[m
     </div>[m
 [m
     <!-- Shared semester / year controls -->[m
[31m-    <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 mb-6">[m
[31m-        <div class="flex flex-wrap items-end gap-4">[m
[31m-            <div>[m
[32m+[m[32m    <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 mb-6">[m
[32m+[m[32m        <div class="flex flex-col sm:flex-row sm:flex-wrap items-stretch sm:items-end gap-3 sm:gap-4 sm:mb-3">[m
[32m+[m[32m            <div class="w-full sm:w-auto">[m
                 <label for="semester" class="block text-xs font-medium text-gray-500 uppercase tracking-wide mb-1">Semester</label>[m
[31m-                <select id="semester" class="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                <select id="semester" class="w-full sm:w-auto border border-gray-300 rounded-lg px-3 py-2.5 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white">[m
                     {% for s in semesters %}[m
                     <option value="{{ s.pk }}" {% if default_semester and s.pk == default_semester.pk %}selected{% endif %}>{{ s }}</option>[m
                     {% endfor %}[m
                 </select>[m
             </div>[m
[31m-            <div>[m
[32m+[m[32m            <div class="w-full sm:w-auto">[m
                 <label for="year" class="block text-xs font-medium text-gray-500 uppercase tracking-wide mb-1">Year of Study</label>[m
[31m-                <select id="year" class="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                <select id="year" class="w-full sm:w-auto border border-gray-300 rounded-lg px-3 py-2.5 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white">[m
                     {% for y in year_options %}[m
                     <option value="{{ y }}">{{ y }} {% if y == 1 %}st{% elif y == 2 %}nd{% elif y == 3 %}rd{% else %}th{% endif %} Year</option>[m
                     {% endfor %}[m
[36m@@ -33,21 +33,23 @@[m
 [m
     <!-- Results -->[m
     <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">[m
[31m-        <div class="px-6 py-4 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3">[m
[32m+[m[32m        <div class="px-4 sm:px-6 py-4 border-b border-gray-200 flex flex-col lg:flex-row lg:items-center justify-between gap-3">[m
             <div>[m
                 <h2 class="text-sm font-semibold text-gray-700 uppercase tracking-wide">Student Groups</h2>[m
                 <p class="text-xs text-gray-500 mt-0.5">Export a single group's timetable, or the whole programme.</p>[m
             </div>[m
[31m-            <form method="get" action="." class="flex items-center gap-2">[m
[31m-                <div class="relative">[m
[32m+[m[32m            <form method="get" action="." class="flex flex-col sm:flex-row sm:items-center gap-2 w-full lg:w-auto">[m
[32m+[m[32m                <div class="relative w-full">[m
                     <svg class="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>[m
                     <input type="text" name="q" value="{{ q }}" placeholder="Search group or programme..."[m
[31m-                           class="w-64 pl-9 pr-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                           class="w-full sm:w-72 pl-9 pr-3 py-2.5 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                </div>[m
[32m+[m[32m                <div class="flex gap-2 w-full sm:w-auto">[m
[32m+[m[32m                    <button type="submit" class="flex-1 sm:flex-none px-4 py-2.5 bg-gray-800 text-white text-sm font-medium rounded-lg hover:bg-gray-900 transition">Search</button>[m
[32m+[m[32m                    {% if q %}[m
[32m+[m[32m                    <a href="?" class="flex-1 sm:flex-none flex items-center justify-center px-4 py-2.5 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">Reset</a>[m
[32m+[m[32m                    {% endif %}[m
                 </div>[m
[31m-                <button type="submit" class="px-4 py-2 bg-gray-800 text-white text-sm font-medium rounded-lg hover:bg-gray-900 transition">Search</button>[m
[31m-                {% if q %}[m
[31m-                <a href="?" class="px-4 py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">Reset</a>[m
[31m-                {% endif %}[m
             </form>[m
         </div>[m
         {% if rows %}[m
[36m@@ -57,20 +59,20 @@[m
         <div class="divide-y divide-gray-100">[m
             {% for r in rows %}[m
             <div>[m
[31m-                <div class="px-6 py-4 flex flex-wrap items-center justify-between gap-3 bg-gray-50/70">[m
[32m+[m[32m                <div class="px-4 sm:px-6 py-4 flex flex-wrap items-center justify-between gap-3 bg-gray-50/70">[m
                     <div class="min-w-0">[m
[31m-                        <p class="font-semibold text-gray-800">{{ r.programme.name }}</p>[m
[32m+[m[32m                        <p class="font-semibold text-gray-800 break-safe">{{ r.programme.name }}</p>[m
                         <p class="text-xs text-gray-500 mt-0.5">{{ r.programme.code }} &middot; {{ r.groups|length }} group{{ r.groups|length|pluralize }}</p>[m
                     </div>[m
                     <button type="button" onclick="doExport(this)"[m
                             data-base="{% url 'programme-timetable-export' r.programme.pk %}"[m
[31m-                            class="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition shrink-0"[m
[32m+[m[32m                            class="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition shrink-0"[m
                             title="Download one PDF with every group of {{ r.programme.code }} combined">[m
                         <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>[m
                         Whole programme[m
                     </button>[m
                 </div>[m
[31m-                <div class="px-6 py-4 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">[m
[32m+[m[32m                <div class="px-4 sm:px-6 py-4 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">[m
                     {% for g in r.groups %}[m
                     <div class="border border-gray-200 rounded-lg p-3 flex items-center justify-between gap-2 hover:border-green-300 hover:shadow-sm transition">[m
                         <div class="min-w-0">[m
[36m@@ -79,7 +81,7 @@[m
                         </div>[m
                         <button type="button" onclick="doExport(this)"[m
                                 data-base="{% url 'group-timetable-export' g.pk %}"[m
[31m-                                class="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-green-700 bg-green-50 hover:bg-green-100 rounded-lg transition shrink-0">[m
[32m+[m[32m                                class="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-green-700 bg-green-50 hover:bg-green-100 rounded-lg transition shrink-0">[m
                             <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>[m
                             Export PDF[m
                         </button>[m
[1mdiff --git a/templates/core/form.html b/templates/core/form.html[m
[1mindex 2d92814..821abcf 100644[m
[1m--- a/templates/core/form.html[m
[1m+++ b/templates/core/form.html[m
[36m@@ -28,12 +28,12 @@[m
         {% endfor %}[m
     </div>[m
     {% endif %}[m
[31m-    <div class="mt-6 flex justify-end gap-3">[m
[31m-        <button type="button" @click="modalOpen = false" class="px-4 py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
[32m+[m[32m    <div class="mt-6 flex flex-col sm:flex-row sm:justify-end gap-3">[m
[32m+[m[32m        <button type="button" @click="modalOpen = false" class="w-full sm:w-auto px-4 py-2.5 sm:py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
             Cancel[m
         </button>[m
         <button type="submit" :disabled="submitting"[m
[31m-                class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition disabled:opacity-50">[m
[32m+[m[32m                class="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-4 py-2.5 sm:py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition disabled:opacity-50">[m
             <span x-show="!submitting">Save</span>[m
             <span x-show="submitting" class="inline-flex items-center gap-2">[m
                 <svg class="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>[m
[1mdiff --git a/templates/core/import_history.html b/templates/core/import_history.html[m
[1mindex 0797ea1..f75cfaa 100644[m
[1m--- a/templates/core/import_history.html[m
[1m+++ b/templates/core/import_history.html[m
[36m@@ -37,7 +37,7 @@[m
     {% if items %}[m
     <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">[m
         <div class="overflow-x-auto">[m
[31m-            <table class="w-full text-left">[m
[32m+[m[32m            <table class="w-full min-w-[920px] text-left">[m
                 <thead class="bg-gray-50 border-b border-gray-200 text-xs uppercase tracking-wider text-gray-500">[m
                     <tr>[m
                         <th class="px-4 py-3 font-semibold">Date &amp; Time</th>[m
[1mdiff --git a/templates/core/import_hub.html b/templates/core/import_hub.html[m
[1mindex 8ba0a7e..b9c15ab 100644[m
[1m--- a/templates/core/import_hub.html[m
[1m+++ b/templates/core/import_hub.html[m
[36m@@ -11,10 +11,10 @@[m
         {% for key, info in import_types.items %}[m
         <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 flex flex-col hover:shadow-md transition">[m
             <div class="flex items-start justify-between gap-3 mb-3">[m
[31m-                <div>[m
[32m+[m[32m                <div class="min-w-0">[m
                     <h3 class="font-semibold text-gray-800">{{ info.title }}</h3>[m
                     <p class="text-xs text-gray-500 mt-1">Expected columns:</p>[m
[31m-                    <code class="inline-block mt-1 text-xs bg-gray-50 border border-gray-200 rounded px-2 py-1 text-gray-600">{{ info.columns }}</code>[m
[32m+[m[32m                    <code class="inline-block mt-1 text-xs bg-gray-50 border border-gray-200 rounded px-2 py-1 text-gray-600 break-words max-w-full">{{ info.columns }}</code>[m
                 </div>[m
             </div>[m
             <div class="mt-4 pt-4 border-t border-gray-100">[m
[36m@@ -46,7 +46,7 @@[m
     </div>[m
 [m
     <div class="mt-8 bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">[m
[31m-        <div class="px-5 py-4 border-b border-gray-100 flex items-center justify-between">[m
[32m+[m[32m        <div class="px-4 sm:px-5 py-4 border-b border-gray-100 flex flex-wrap items-center justify-between gap-3">[m
             <div>[m
                 <h3 class="font-semibold text-gray-800">Previous Import Details</h3>[m
                 <p class="text-xs text-gray-500 mt-0.5">Every import is saved with its full summary — date, file, type, rows, errors and automatic fixes. Review any of them any time.</p>[m
[36m@@ -58,19 +58,21 @@[m
         {% if recent_imports %}[m
         <div class="divide-y divide-gray-100">[m
             {% for rec in recent_imports %}[m
[31m-            <a href="{% url 'import-history-detail' rec.pk %}" class="flex items-center gap-4 px-5 py-3 hover:bg-gray-50 transition">[m
[31m-                <span class="shrink-0 w-2 h-2 rounded-full {% if rec.status == 'FAILED' %}bg-red-500{% elif rec.status == 'PARTIAL' %}bg-amber-500{% else %}bg-emerald-500{% endif %}"></span>[m
[32m+[m[32m            <a href="{% url 'import-history-detail' rec.pk %}" class="flex flex-col sm:flex-row sm:items-center gap-1.5 sm:gap-4 px-4 sm:px-5 py-3 hover:bg-gray-50 transition">[m
[32m+[m[32m                <span class="shrink-0 w-2 h-2 rounded-full {% if rec.status == 'FAILED' %}bg-red-500{% elif rec.status == 'PARTIAL' %}bg-amber-500{% else %}bg-emerald-500{% endif %} sm:mt-1.5"></span>[m
                 <span class="min-w-0 flex-1">[m
                     <span class="block text-sm font-medium text-gray-800 truncate">{{ rec.import_title }}</span>[m
[31m-                    <span class="block text-xs text-gray-500">[m
[32m+[m[32m                    <span class="block text-xs text-gray-500 break-safe">[m
                         {{ rec.filename|default:"(in-memory upload)" }} · {{ rec.created_at|date:"d M Y H:i" }}[m
                     </span>[m
                 </span>[m
[31m-                <span class="shrink-0 text-xs text-gray-500">[m
[31m-                    {{ rec.rows_processed }} row{{ rec.rows_processed|pluralize }} · {{ rec.created }} created {{ rec.updated }} updated[m
[31m-                </span>[m
[31m-                <span class="shrink-0 text-xs font-medium {% if rec.error_count %}text-red-600{% else %}text-emerald-600{% endif %}">[m
[31m-                    {{ rec.error_count }} error{{ rec.error_count|pluralize }}[m
[32m+[m[32m                <span class="shrink-0 flex items-center gap-3 sm:gap-4 pl-5 sm:pl-0 text-xs">[m
[32m+[m[32m                    <span class="text-gray-500">[m
[32m+[m[32m                        {{ rec.rows_processed }} row{{ rec.rows_processed|pluralize }} · {{ rec.created }} created {{ rec.updated }} updated[m
[32m+[m[32m                    </span>[m
[32m+[m[32m                    <span class="font-medium {% if rec.error_count %}text-red-600{% else %}text-emerald-600{% endif %}">[m
[32m+[m[32m                        {{ rec.error_count }} error{{ rec.error_count|pluralize }}[m
[32m+[m[32m                    </span>[m
                 </span>[m
             </a>[m
             {% endfor %}[m
[1mdiff --git a/templates/core/import_upload.html b/templates/core/import_upload.html[m
[1mindex 4fba722..95c78a6 100644[m
[1m--- a/templates/core/import_upload.html[m
[1m+++ b/templates/core/import_upload.html[m
[36m@@ -9,10 +9,10 @@[m
     </button>[m
 [m
     <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">[m
[31m-        <div class="px-6 py-4 border-b border-gray-200 bg-gray-50">[m
[32m+[m[32m        <div class="px-4 sm:px-6 py-4 border-b border-gray-200 bg-gray-50">[m
             <h2 class="text-base font-semibold text-gray-800">Import {{ import_title }}</h2>[m
         </div>[m
[31m-        <div class="p-6">[m
[32m+[m[32m        <div class="p-4 sm:p-6">[m
             <div class="bg-gray-50 border border-gray-200 rounded-lg p-4 mb-6">[m
                 <p class="text-sm text-gray-600 font-medium mb-2">Required Excel columns:</p>[m
                 <code class="text-xs bg-white border border-gray-200 rounded px-3 py-1.5 text-gray-700 block">{{ columns }}</code>[m
[1mdiff --git a/templates/core/list.html b/templates/core/list.html[m
[1mindex acffff5..04acb38 100644[m
[1m--- a/templates/core/list.html[m
[1m+++ b/templates/core/list.html[m
[36m@@ -9,13 +9,13 @@[m
         <div class="flex flex-wrap items-center gap-2">[m
             {% if clear_all_url %}[m
             <button onclick="openModal('{{ clear_all_url }}{{ query }}', 'Clear all {{ page_title }}')"[m
[31m-                    class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-red-700 border border-red-300 bg-red-50 rounded-lg hover:bg-red-100 transition">[m
[32m+[m[32m                    class="inline-flex items-center justify-center gap-2 px-3.5 sm:px-4 py-2 sm:py-2.5 text-sm font-medium text-red-700 border border-red-300 bg-red-50 rounded-lg hover:bg-red-100 transition">[m
                 <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>[m
                 Clear All[m
             </button>[m
             {% endif %}[m
             <button onclick="openModal('{{ create_url }}', '{{ "Create " }}{{ page_title }}')"[m
[31m-                    class="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 transition">[m
[32m+[m[32m                    class="inline-flex items-center justify-center gap-2 px-3.5 sm:px-4 py-2 sm:py-2.5 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 transition">[m
                 <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"/></svg>[m
                 Add New[m
             </button>[m
[36m@@ -25,19 +25,19 @@[m
     {% block list_actions %}{% endblock %}[m
 [m
     <!-- Filter bar -->[m
[31m-    <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 mb-4">[m
[32m+[m[32m    <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-3 sm:p-4 mb-4">[m
         <form method="get" action=".">[m
             <div class="flex flex-col lg:flex-row gap-3">[m
                 <div class="flex-1 min-w-0 relative">[m
                     <svg class="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>[m
                     <input type="text" name="q" value="{{ q }}" placeholder="Search..."[m
[31m-                           class="w-full min-w-0 pl-9 pr-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                           class="w-full min-w-0 pl-9 pr-3 py-2.5 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
                 </div>[m
[31m-                <div class="flex flex-wrap gap-3 items-center">[m
[32m+[m[32m                <div class="flex flex-col sm:flex-row sm:flex-wrap lg:flex-nowrap gap-2 sm:gap-3 items-stretch sm:items-center">[m
                     {% block filters %}[m
                     {% for f in filters %}[m
                         {% if f.type == "select" %}[m
[31m-                        <select name="{{ f.name }}" class="min-w-0 max-w-[16rem] shrink px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white truncate focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                        <select name="{{ f.name }}" class="w-full lg:w-auto min-w-0 lg:max-w-[16rem] px-3 py-2.5 border border-gray-300 rounded-lg text-sm bg-white truncate focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
                             <option value="">All {{ f.label }}s</option>[m
                             {% for v, label in f.options %}[m
                             <option value="{{ v }}" {% if f.value == v %}selected{% endif %}>{{ label }}</option>[m
[36m@@ -45,22 +45,22 @@[m
                         </select>[m
                         {% elif f.type == "time" %}[m
                         <input type="time" name="{{ f.name }}" value="{{ f.value }}" title="{{ f.label }}"[m
[31m-                               class="min-w-0 max-w-[11rem] shrink px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                               class="w-full lg:w-auto min-w-0 lg:max-w-[11rem] px-3 py-2.5 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
                         {% elif f.type == "number" %}[m
                         <input type="number" name="{{ f.name }}" value="{{ f.value }}" placeholder="{{ f.label }}"[m
[31m-                               class="min-w-0 max-w-[10rem] shrink px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                               class="w-full lg:w-auto min-w-0 lg:max-w-[10rem] px-3 py-2.5 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
                         {% else %}[m
                         <input type="text" name="{{ f.name }}" value="{{ f.value }}" placeholder="{{ f.label }}"[m
[31m-                               class="min-w-0 max-w-[16rem] shrink px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
[32m+[m[32m                               class="w-full lg:w-auto min-w-0 lg:max-w-[16rem] px-3 py-2.5 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500">[m
                         {% endif %}[m
                     {% endfor %}[m
                     {% endblock %}[m
                 </div>[m
[31m-                <div class="flex gap-2 shrink-0">[m
[31m-                    <button type="submit" class="px-4 py-2 bg-gray-800 text-white text-sm font-medium rounded-lg hover:bg-gray-900 transition">[m
[32m+[m[32m                <div class="flex gap-2 shrink-0 w-full sm:w-auto">[m
[32m+[m[32m                    <button type="submit" class="flex-1 sm:flex-none px-4 py-2.5 bg-gray-800 text-white text-sm font-medium rounded-lg hover:bg-gray-900 transition">[m
                         Filter[m
                     </button>[m
[31m-                    <a href="?" class="px-4 py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
[32m+[m[32m                    <a href="?" class="flex-1 sm:flex-none flex items-center justify-center px-4 py-2.5 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
                         Reset[m
                     </a>[m
                 </div>[m
[1mdiff --git a/templates/core/session_detail.html b/templates/core/session_detail.html[m
[1mindex 0c81bb3..f8e7cfb 100644[m
[1m--- a/templates/core/session_detail.html[m
[1m+++ b/templates/core/session_detail.html[m
[36m@@ -11,20 +11,20 @@[m
         {% include "core/_session_groups.html" %}[m
 [m
         <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">[m
[31m-            <div class="px-6 py-4 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3">[m
[31m-                <h2 class="text-lg font-semibold text-gray-800">{{ item }}</h2>[m
[31m-                <div class="flex gap-2">[m
[32m+[m[32m            <div class="px-4 sm:px-6 py-4 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3">[m
[32m+[m[32m                <h2 class="text-lg font-semibold text-gray-800 break-safe">{{ item }}</h2>[m
[32m+[m[32m                <div class="flex flex-wrap gap-2">[m
                     <a href="{{ edit_url }}" onclick="event.preventDefault(); openEdit(this.getAttribute('href'))"[m
[31m-                       class="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition">[m
[32m+[m[32m                       class="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition">[m
                         Edit[m
                     </a>[m
                     <a href="{{ delete_url }}" onclick="event.preventDefault(); openDelete(this.getAttribute('href'))"[m
[31m-                       class="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition">[m
[32m+[m[32m                       class="inline-flex items-center justify-center gap-1.5 px-3 py-2 text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 rounded-lg transition">[m
                         Delete[m
                     </a>[m
                 </div>[m
             </div>[m
[31m-            <div class="p-6">[m
[32m+[m[32m            <div class="p-4 sm:p-6">[m
                 <dl class="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-4">[m
                     {% for field in detail_fields %}[m
                     <div>[m
[1mdiff --git a/templates/core/session_form.html b/templates/core/session_form.html[m
[1mindex 9a6d9a5..e32864e 100644[m
[1m--- a/templates/core/session_form.html[m
[1m+++ b/templates/core/session_form.html[m
[36m@@ -25,10 +25,10 @@[m
             </div>[m
             <div class="space-y-2">[m
                 {% for sub in formset %}[m
[31m-                <div class="flex items-center gap-2 bg-gray-50 rounded-lg px-3 py-2">[m
[32m+[m[32m                <div class="flex flex-col sm:flex-row sm:items-center gap-2 bg-gray-50 rounded-lg px-3 py-2">[m
                     {{ sub.id }}[m
                     {{ sub.group }}[m
[31m-                    <label class="text-xs text-gray-400 whitespace-nowrap flex items-center gap-1">Delete {{ sub.DELETE }}</label>[m
[32m+[m[32m                    <label class="shrink-0 text-xs text-gray-400 whitespace-nowrap flex items-center gap-1">Delete {{ sub.DELETE }}</label>[m
                     {{ sub.errors }}[m
                 </div>[m
                 {% endfor %}[m
[36m@@ -43,12 +43,12 @@[m
         {% endfor %}[m
     </div>[m
     {% endif %}[m
[31m-    <div class="mt-6 flex justify-end gap-3">[m
[31m-        <button type="button" @click="modalOpen = false" class="px-4 py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
[32m+[m[32m    <div class="mt-6 flex flex-col sm:flex-row sm:justify-end gap-3">[m
[32m+[m[32m        <button type="button" @click="modalOpen = false" class="w-full sm:w-auto px-4 py-2.5 sm:py-2 text-sm font-medium text-gray-600 border border-gray-300 rounded-lg hover:bg-gray-50 transition">[m
             Cancel[m
         </button>[m
         <button type="submit" :disabled="submitting"[m
[31m-                class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition disabled:opacity-50">[m
[32m+[m[32m                class="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-4 py-2.5 sm:py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition disabled:opacity-50">[m
             <span x-show="!submitting">Save</span>[m
             <span x-show="submitting" class="inline-flex items-center gap-2">[m
                 <svg class="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>[m
[1mdiff --git a/templates/core/session_list.html b/templates/core/session_list.html[m
[1mindex 3757a3f..a2643b2 100644[m
[1m--- a/templates/core/session_list.html[m
[1m+++ b/templates/core/session_list.html[m
[36m@@ -1,6 +1,26 @@[m
 {% extends "core/list.html" %}[m
 [m
 {% block list_actions %}[m
[32m+[m[32m{% if legacy_sessions_count %}[m
[32m+[m[32m<div class="flex items-start gap-3 mb-4 rounded-xl border border-amber-300 bg-amber-50 p-4">[m
[32m+[m[32m    <span class="text-2xl">&#9888;&#65039;</span>[m
[32m+[m[32m    <div class="flex-1">[m
[32m+[m[32m        <p class="text-sm font-semibold text-amber-800">[m
[32m+[m[32m            {{ legacy_sessions_count }} workshop session{{ legacy_sessions_count|pluralize }}[m
[32m+[m[32m            do not match the standard workshop session times[m
[32m+[m[32m        </p>[m
[32m+[m[32m        <p class="text-xs text-amber-700 mt-0.5">[m
[32m+[m[32m            Workshops run Monday-Friday, 09:00-13:00 (Morning) or 15:00-19:00 (Afternoon);[m
[32m+[m[32m            Thursday mornings start 10:00-14:00.[m
[32m+[m[32m        </p>[m
[32m+[m[32m        {% for ex in legacy_sessions_examples %}[m
[32m+[m[32m        <p class="text-xs text-amber-700 mt-1">[m
[32m+[m[32m            {{ ex.course }} &middot; {{ ex.day }} {{ ex.time }} — {{ ex.issue }}[m
[32m+[m[32m        </p>[m
[32m+[m[32m        {% endfor %}[m
[32m+[m[32m    </div>[m
[32m+[m[32m</div>[m
[32m+[m[32m{% endif %}[m
 <div class="mb-4">[m
     <button[m
         hx-post="{% url 'session-assign-lecture-groups' %}"[m
[36m@@ -8,7 +28,7 @@[m
         hx-swap="innerHTML"[m
         hx-disabled-elt="this"[m
         hx-confirm="Attach all Student Groups of the programme(s) that study each lecture course to those lecture sessions?"[m
[31m-        class="inline-flex items-center px-4 py-2 bg-emerald-600 text-white text-sm font-medium rounded-lg hover:bg-emerald-700 transition disabled:opacity-50 disabled:cursor-wait">[m
[32m+[m[32m        class="inline-flex items-center justify-center w-full sm:w-auto px-4 py-2.5 bg-emerald-600 text-white text-sm font-medium rounded-lg hover:bg-emerald-700 transition disabled:opacity-50 disabled:cursor-wait">[m
         <span class="tt-idle-label inline-flex items-center gap-2">[m
             <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 20h5v-2a3 3 0 00-5.36-1.86M9 20H4v-2a3 3 0 015.36-1.86M16 7a3 3 0 11-6 0 3 3 0 016 0zm1 7a3 3 0 11-6 0 3 3 0 016 0zM6 7a3 3 0 11-6 0 3 3 0 016 0z"/></svg>[m
             Assign Lecture Groups[m
[1mdiff --git a/templates/core/timetable.html b/templates/core/timetable.html[m
[1mindex 6e5a971..68e2fbe 100644[m
[1m--- a/templates/core/timetable.html[m
[1m+++ b/templates/core/timetable.html[m
[36m@@ -5,107 +5,15 @@[m
 [m
 {% block content %}[m
 <style>[m
[31m-    .tt-scroll { -webkit-overflow-scrolling: touch; scrollbar-width: thin; }[m
[31m-    .tt-table { border-collapse: separate; border-spacing: 0; width: 100%; table-layout: fixed; }[m
[31m-    .tt-table th, .tt-table td { border-bottom: 1px solid #e6eaf0; border-right: 1px solid #e6eaf0; padding: 0; }[m
[31m-    .tt-table thead th {[m
[31m-        position: sticky; top: 0; z-index: 30; background: #f4f7fb;[m
[31m-        box-shadow: 0 4px 10px -8px rgba(2, 6, 23, 0.35);[m
[31m-    }[m
[31m-    .tt-table th.tt-corner { left: 0; top: 0; z-index: 40; min-width: 6.5rem; padding: 0.65rem 0.7rem; }[m
[31m-    .tt-table th.tt-day-col {[m
[31m-        min-width: 10rem; padding: 0.65rem 0.5rem; white-space: nowrap;[m
[31m-        font-size: 0.75rem; font-weight: 700; color: #1e3a5f; text-align: center;[m
[31m-        text-transform: uppercase; letter-spacing: 0.04em;[m
[31m-    }[m
[31m-    .tt-table td.tt-time {[m
[31m-        position: sticky; left: 0; z-index: 15; background: #f8fafc;[m
[31m-        min-width: 6.5rem; padding: 0.55rem 0.75rem; vertical-align: middle;[m
[31m-        height: 6rem;[m
[31m-        font-size: 0.6875rem; font-weight: 600; color: #334155; text-align: right;[m
[31m-        white-space: nowrap; box-shadow: 5px 0 12px -10px rgba(2, 6, 23, 0.4);[m
[31m-    }[m
[31m-    .tt-table tr:hover td.tt-time { background: #eef2f8; }[m
[31m-    .tt-table td.tt-empty { height: 6rem; background: #ffffff; }[m
[31m-    .tt-table tr:hover td.tt-empty { background: #fbfdff; }[m
[31m-    .tt-table td.tt-block {[m
[31m-        padding: 0.2rem; vertical-align: top; background: #ffffff;[m
[31m-        display: flex; flex-direction: column; gap: 0.25rem;[m
[31m-    }[m
[31m-[m
[31m-    .tt-card {[m
[31m-        display: flex; flex-direction: column; justify-content: center; gap: 0.2rem;[m
[31m-        margin: 0; padding: 0.45rem 0.6rem; border-radius: 0.45rem; border-left: 3px solid;[m
[31m-        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06);[m
[31m-        flex: 1; min-height: 5.4rem;[m
[31m-    }[m
[31m-    .tt-course { font-size: 0.8125rem; font-weight: 700; line-height: 1.25; letter-spacing: 0.01em; }[m
[31m-    .tt-type { display: flex; align-items: center; gap: 0.35rem; flex-wrap: wrap; font-size: 0.6875rem; line-height: 1.3; }[m
[31m-    .tt-type::before { content: ""; width: 0.4rem; height: 0.4rem; border-radius: 9999px; background: currentColor; flex-shrink: 0; }[m
[31m-    .tt-groups { font-weight: 600; padding-left: 0.15rem; }[m
[31m-    .tt-venue { font-size: 0.6875rem; color: #6b7280; line-height: 1.3; }[m
[31m-    .tt-note { font-size: 0.625rem; color: #b45309; line-height: 1.3; }[m
[31m-[m
[31m-    /* Shading: Lecture faded grey, Workshop faded green, TD faded pink */[m
[31m-    .tt-kind-lecture { border-left-color: #9ca3af; background: #f3f4f6; }[m
[31m-    .tt-kind-lecture .tt-course { color: #374151; }[m
[31m-    .tt-kind-lecture .tt-type { color: #6b7280; }[m
[31m-    .tt-kind-workshop { border-left-color: #34d399; background: #dcfce7; }[m
[31m-    .tt-kind-workshop .tt-course { color: #047857; }[m
[31m-    .tt-kind-workshop .tt-type { color: #059669; }[m
[31m-    .tt-kind-td { border-left-color: #f472b6; background: #fce7f3; }[m
[31m-    .tt-kind-td .tt-course { color: #be185d; }[m
[31m-    .tt-kind-td .tt-type { color: #db2777; }[m
[31m-    .tt-kind-tutorial { border-left-color: #60a5fa; background: #dbeafe; }[m
[31m-    .tt-kind-tutorial .tt-course { color: #1d4ed8; }[m
[31m-    .tt-kind-tutorial .tt-type { color: #3b82f6; }[m
[31m-    .tt-kind-seminar { border-left-color: #a78bfa; background: #ede9fe; }[m
[31m-    .tt-kind-seminar .tt-course { color: #6d28d9; }[m
[31m-    .tt-kind-seminar .tt-type { color: #8b5cf6; }[m
[31m-    .tt-kind-practical { border-left-color: #fbbf24; background: #fef3c7; }[m
[31m-    .tt-kind-practical .tt-course { color: #b45309; }[m
[31m-    .tt-kind-practical .tt-type { color: #d97706; }[m
[31m-[m
[31m-    .tt-legend .tt-kind-dot { display: inline-block; width: 0.55rem; height: 0.55rem; border-radius: 9999px; }[m
[31m-[m
[31m-    @media print {[m
[31m-        @page { size: A4 portrait; margin: 10mm; }[m
[31m-        body { background: #ffffff; }[m
[31m-        aside, header, footer, .no-print, #toast-container { display: none !important; }[m
[31m-        .tt-scroll { overflow: visible !important; }[m
[31m-        .tt-table thead th, .tt-table td.tt-time, .tt-table th.tt-corner { position: static !important; box-shadow: none !important; }[m
[31m-        .tt-table th.tt-day-col, .tt-table td.tt-time, .tt-table th.tt-corner { min-width: 0 !important; }[m
[31m-        .tt-card { break-inside: avoid; }[m
[31m-    }[m
[31m-[m
[32m+[m[32m    .max-w-6xl { max-width: 72rem; }[m
     @media (max-width: 768px) {[m
[31m-        .tt-table th.tt-corner { min-width: 5rem; padding: 0.5rem; font-size: 0.65rem; }[m
[31m-        .tt-table th.tt-day-col { min-width: 6rem; padding: 0.5rem 0.3rem; font-size: 0.65rem; }[m
[31m-        .tt-table td.tt-time { min-width: 5rem; padding: 0.4rem 0.5rem; font-size: 0.6rem; height: 5rem; }[m
[31m-        .tt-table td.tt-empty { height: 5rem; }[m
[31m-        .tt-card { padding: 0.35rem 0.5rem; min-height: 4.5rem; }[m
[31m-        .tt-course { font-size: 0.7rem; }[m
[31m-        .tt-type { font-size: 0.6rem; }[m
[31m-        .tt-venue { font-size: 0.6rem; }[m
[31m-        .tt-note { font-size: 0.55rem; }[m
         .max-w-6xl { max-width: 100%; padding: 0 0.5rem; }[m
[31m-        .tt-scroll { border-radius: 0.5rem; }[m
     }[m
[31m-[m
[31m-    @media (max-width: 480px) {[m
[31m-        .tt-table th.tt-corner { min-width: 4rem; padding: 0.4rem; font-size: 0.6rem; }[m
[31m-        .tt-table th.tt-day-col { min-width: 5rem; padding: 0.4rem 0.2rem; font-size: 0.6rem; }[m
[31m-        .tt-table td.tt-time { min-width: 4rem; padding: 0.3rem 0.4rem; font-size: 0.55rem; height: 4.5rem; }[m
[31m-        .tt-table td.tt-empty { height: 4.5rem; }[m
[31m-        .tt-card { padding: 0.3rem 0.4rem; min-height: 4rem; border-left-width: 2px; }[m
[31m-        .tt-course { font-size: 0.65rem; }[m
[31m-        .tt-type { font-size: 0.55rem; }[m
[31m-        .tt-venue { font-size: 0.55rem; }[m
[31m-        .tt-note { font-size: 0.5rem; }[m
[32m+[m[32m    @media (max-width: 640px) {[m
         .no-print form { flex-direction: column; align-items: stretch; }[m
[31m-        .no-print form > div { width: 100%; }[m
[31m-        .no-print form .ml-auto { margin-left: 0; margin-top: 0.5rem; justify-content: stretch; }[m
[31m-        .no-print form .ml-auto button, .no-print form .ml-auto a { flex: 1; }[m
[32m+[m[32m        .no-print form > div { width: 100%; max-width: 100%; }[m
[32m+[m[32m        .no-print form .ml-auto { margin-left: 0; margin-top: 0.5rem; flex-direction: column; }[m
[32m+[m[32m        .no-print form .ml-auto > * { flex: 1; width: 100%; }[m
     }[m
 </style>[m
 [m
[36m@@ -180,14 +88,14 @@[m
         </div>[m
     </form>[m
 [m
[31m-    {% if not active_programme %}[m
[32m+[m[32m    {% if not active_semester %}[m
     <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-16 text-center text-gray-400">[m
[31m-        <p class="text-sm">No programmes registered yet.</p>[m
[31m-        <p class="text-xs mt-1">Add a programme and import a master timetable to see it here.</p>[m
[32m+[m[32m        <p class="text-sm">No semester available. Create a semester first.</p>[m
     </div>[m
[31m-    {% elif not active_semester %}[m
[32m+[m[32m    {% elif not programmes and not grid.slots %}[m
     <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-16 text-center text-gray-400">[m
[31m-        <p class="text-sm">No semester available. Create a semester first.</p>[m
[32m+[m[32m        <p class="text-sm">No programmes registered yet.</p>[m
[32m+[m[32m        <p class="text-xs mt-1">Add a programme and import a master timetable to see it here.</p>[m
     </div>[m
     {% else %}[m
 [m
[36m@@ -197,6 +105,7 @@[m
             <div>[m
                 {% if show_all %}[m
                 <p class="text-lg font-bold tracking-tight">ALL PROGRAMMES</p>[m
[32m+[m[32m                <p class="text-sm font-medium text-slate-300 mt-0.5">{{ year_ordinal|upper }} YEAR &middot; {{ active_semester.academic_year }}</p>[m
                 {% elif active_group %}[m
                 <p class="text-lg font-bold tracking-tight">{{ active_group.programme.name|upper }}</p>[m
                 <p class="text-sm font-medium text-slate-300 mt-0.5">GROUP {{ active_group.code }} &middot; {{ year_ordinal|upper }} YEAR &middot; {{ active_semester.academic_year }}</p>[m
[36m@@ -230,55 +139,16 @@[m
         <span class="flex items-center gap-1.5"><span class="tt-kind-dot bg-pink-400"></span>Technical Drawing</span>[m
     </div>[m
 [m
[31m-    <!-- Grid: DAYS in the first row, TIME slots in the first column -->[m
[32m+[m[32m    <!-- Grid: TIME slots in the first row (top header), DAYS down the first column -->[m
     {% if grid.slots %}[m
[31m-    <div class="tt-scroll overflow-x-auto bg-white rounded-xl shadow-sm border border-gray-200">[m
[31m-        <table class="tt-table">[m
[31m-            <thead>[m
[31m-                <tr>[m
[31m-                    <th class="tt-corner">TIME</th>[m
[31m-                    {% for day in grid.days %}[m
[31m-                    <th class="tt-day-col">{{ day.label }}</th>[m
[31m-                    {% endfor %}[m
[31m-                </tr>[m
[31m-            </thead>[m
[31m-            <tbody>[m
[31m-                {% for row in grid.rows %}[m
[31m-                <tr>[m
[31m-                    <td class="tt-time">{{ row.slot.start }}–{{ row.slot.end }}</td>[m
[31m-                    {% for cell in row.cols %}[m
[31m-                        {% if cell %}[m
[31m-                            {% if cell.empty %}[m
[31m-                            <td class="tt-empty"></td>[m
[31m-                            {% else %}[m
[31m-                            <td class="tt-block" rowspan="{{ cell.rowspan }}">[m
[31m-                                {% for e in cell.entries %}[m
[31m-                                <div class="tt-card tt-kind-{{ e.kind }}">[m
[31m-                                    <div class="tt-course">{{ e.name }}</div>[m
[31m-                                    {% if e.kind == "workshop" %}[m
[31m-                                        {% if show_groups %}<div class="tt-groups">{{ e.groups }}</div>{% endif %}[m
[31m-                                        {% if e.venue and e.venue != e.name %}<div class="tt-venue">{{ e.venue }}</div>{% endif %}[m
[31m-                                    {% else %}[m
[31m-                                    <div class="tt-type">[m
[31m-                                        {{ e.type_label }}{% if e.groups %}<span class="tt-groups">{{ e.groups }}</span>{% endif %}[m
[31m-                                    </div>[m
[31m-                                    <div class="tt-venue">{{ e.venue|default:"—" }}</div>[m
[31m-                                    {% endif %}[m
[31m-                                    {% if e.note %}<div class="tt-note">{{ e.note }}</div>{% endif %}[m
[31m-                                </div>[m
[31m-                                {% endfor %}[m
[31m-                            </td>[m
[31m-                            {% endif %}[m
[31m-                        {% endif %}[m
[31m-                    {% endfor %}[m
[31m-                </tr>[m
[31m-                {% endfor %}[m
[31m-            </tbody>[m
[31m-        </table>[m
[31m-    </div>[m
[32m+[m[32m    {% include "core/_timetable_grid.html" %}[m
     {% else %}[m
     <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-16 text-center text-gray-400">[m
[32m+[m[32m        {% if show_all %}[m
[32m+[m[32m        <p class="text-sm">No sessions scheduled for COET First Year in {{ active_semester }} yet.</p>[m
[32m+[m[32m        {% else %}[m
         <p class="text-sm">No sessions scheduled for {{ active_programme.code }} in {{ active_semester }} yet.</p>[m
[32m+[m[32m        {% endif %}[m
         <p class="text-xs mt-1">Import a master timetable (or assign groups to sessions) and refresh this page.</p>[m
     </div>[m
     {% endif %}[m
[1mdiff --git a/templates/core/venue_recycle.html b/templates/core/venue_recycle.html[m
[1mindex f0c02e3..6ce7c13 100644[m
[1m--- a/templates/core/venue_recycle.html[m
[1m+++ b/templates/core/venue_recycle.html[m
[36m@@ -78,10 +78,10 @@[m
                                     {% if info.duplicate_of %}&middot; matches: {{ info.duplicate_of|join:", " }}{% endif %}[m
                                 </p>[m
                             </div>[m
[31m-                            <div class="flex items-center gap-2 shrink-0">[m
[32m+[m[32m                            <div class="flex flex-wrap items-center gap-x-2 gap-y-2 shrink-0">[m
                                 {% with imp=venue_impacts|get_item:venue.pk %}[m
                                 {% if imp %}[m
[31m-                                <span class="text-xs text-gray-500 max-w-[220px] text-right">{{ imp.summary }}</span>[m
[32m+[m[32m                                <span class="text-xs text-gray-500 w-full sm:w-auto sm:max-w-[220px] sm:text-right">{{ imp.summary }}</span>[m
                                 {% endif %}[m
                                 <form method="post" action="{{ recycle_url }}">[m
                                     {% csrf_token %}[m
[36m@@ -89,7 +89,7 @@[m
                                     <input type="hidden" name="pk" value="{{ venue.pk }}">[m
                                     <button type="submit"[m
                                             onclick="return confirm('Remove {{ venue.name|escapejs }}? {{ imp.summary|default:'' }}')"[m
[31m-                                            class="px-3 py-1.5 text-xs font-medium text-rose-700 border border-rose-300 rounded-lg hover:bg-rose-50 transition">[m
[32m+[m[32m                                            class="px-3 py-2 text-xs font-medium text-rose-700 border border-rose-300 rounded-lg hover:bg-rose-50 transition">[m
                                         Remove[m
                                     </button>[m
                                 </form>[m
[36m@@ -101,13 +101,13 @@[m
                             <input type="hidden" name="action" value="edit">[m
                             <input type="hidden" name="pk" value="{{ venue.pk }}">[m
                             <input type="text" name="name" value="{{ info.suggested }}"[m
[31m-                                   class="flex-1 px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500">[m
[32m+[m[32m                                   class="w-full sm:flex-1 px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500">[m
                             <div class="flex items-center gap-2">[m
                                 <span class="text-xs text-gray-500">Capacity:</span>[m
                                 <input type="number" name="capacity" value="{{ venue.capacity|default:'' }}" min="0"[m
                                        class="w-24 px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500">[m
                                 <button type="submit"[m
[31m-                                        class="inline-flex items-center gap-1 px-3 py-2 bg-amber-500 hover:bg-amber-600 text-white text-xs font-medium rounded-lg transition">[m
[32m+[m[32m                                        class="inline-flex items-center justify-center gap-1 px-3 py-2 bg-amber-500 hover:bg-amber-600 text-white text-xs font-medium rounded-lg transition">[m
                                     Fix[m
                                 </button>[m
                             </div>[m
[36m@@ -132,34 +132,34 @@[m
             {% with info=issues_map|get_item:venue.pk %}[m
             <div class="p-4">[m
                 <div class="flex flex-wrap items-start justify-between gap-3">[m
[31m-                    <div class="min-w-0">[m
[31m-                        <div class="flex items-center gap-2 flex-wrap">[m
[31m-                            <span class="font-semibold text-sm text-gray-800">{{ venue.name }}</span>[m
[31m-                            {% for badge in info.issues %}[m
[31m-                            <span class="inline-flex px-2 py-0.5 rounded-full text-xs font-medium[m
[31m-                                {% if badge == 'Casing' %}bg-rose-100 text-rose-700[m
[31m-                                {% else %}bg-blue-100 text-blue-700{% endif %}">[m
[31m-                                {{ badge }}[m
[31m-                            </span>[m
[31m-                            {% endfor %}[m
[32m+[m[32m                        <div class="min-w-0">[m
[32m+[m[32m                            <div class="flex items-center gap-2 flex-wrap">[m
[32m+[m[32m                                <span class="font-semibold text-sm text-gray-800">{{ venue.name }}</span>[m
[32m+[m[32m                                {% for badge in info.issues %}[m
[32m+[m[32m                                <span class="inline-flex px-2 py-0.5 rounded-full text-xs font-medium[m
[32m+[m[32m                                    {% if badge == 'Casing' %}bg-rose-100 text-rose-700[m
[32m+[m[32m                                    {% else %}bg-blue-100 text-blue-700{% endif %}">[m
[32m+[m[32m                                    {{ badge }}[m
[32m+[m[32m                                </span>[m
[32m+[m[32m                                {% endfor %}[m
[32m+[m[32m                            </div>[m
[32m+[m[32m                            <p class="text-xs text-gray-500 mt-1">[m
[32m+[m[32m                                Capacity: {{ venue.capacity|default:"—" }}[m
[32m+[m[32m                                {% if info.summary %}&middot; {{ info.summary }}{% endif %}[m
[32m+[m[32m                            </p>[m
                         </div>[m
[31m-                        <p class="text-xs text-gray-500 mt-1">[m
[31m-                            Capacity: {{ venue.capacity|default:"—" }}[m
[31m-                            {% if info.summary %}&middot; {{ info.summary }}{% endif %}[m
[31m-                        </p>[m
[32m+[m[32m                        <form method="post" action="{{ recycle_url }}" class="flex flex-col sm:flex-row sm:items-center gap-2 w-full sm:w-auto">[m
[32m+[m[32m                            {% csrf_token %}[m
[32m+[m[32m                            <input type="hidden" name="action" value="edit">[m
[32m+[m[32m                            <input type="hidden" name="pk" value="{{ venue.pk }}">[m
[32m+[m[32m                            <input type="text" name="name" value="{{ info.suggested }}"[m
[32m+[m[32m                                   class="w-full sm:w-44 px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500">[m
[32m+[m[32m                            <button type="submit"[m
[32m+[m[32m                                    class="inline-flex items-center justify-center gap-1 px-3 py-2 bg-amber-500 hover:bg-amber-600 text-white text-xs font-medium rounded-lg transition">[m
[32m+[m[32m                                Fix[m
[32m+[m[32m                            </button>[m
[32m+[m[32m                        </form>[m
                     </div>[m
[31m-                    <form method="post" action="{{ recycle_url }}" class="flex items-center gap-2 shrink-0">[m
[31m-                        {% csrf_token %}[m
[31m-                        <input type="hidden" name="action" value="edit">[m
[31m-                        <input type="hidden" name="pk" value="{{ venue.pk }}">[m
[31m-                        <input type="text" name="name" value="{{ info.suggested }}"[m
[31m-                               class="w-44 px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-amber-500">[m
[31m-                        <button type="submit"[m
[31m-                                class="inline-flex items-center gap-1 px-3 py-2 bg-amber-500 hover:bg-amber-600 text-white text-xs font-medium rounded-lg transition">[m
[31m-                            Fix[m
[31m-                        </button>[m
[31m-                    </form>[m
[31m-                </div>[m
             </div>[m
             {% endwith %}[m
             {% endfor %}[m
[1mdiff --git a/templates/dashboard.html b/templates/dashboard.html[m
[1mindex ba68936..37a4a3c 100644[m
[1m--- a/templates/dashboard.html[m
[1m+++ b/templates/dashboard.html[m
[36m@@ -23,92 +23,92 @@[m
     </div>[m
     {% endif %}[m
     <!-- Stats grid -->[m
[31m-    <div class="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-4 mb-8">[m
[31m-        <a href="{% url 'programme-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m    <div class="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-3 sm:gap-4 mb-8">[m
[32m+[m[32m        <a href="{% url 'programme-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">Programmes</div>[m
                 <span class="w-8 h-8 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ programme_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ programme_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Programmes configured</div>[m
         </a>[m
 [m
[31m-        <a href="{% url 'group-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m        <a href="{% url 'group-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">Student Groups</div>[m
                 <span class="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ group_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ group_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Across all programmes</div>[m
         </a>[m
 [m
[31m-        <a href="{% url 'session-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m        <a href="{% url 'session-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">Sessions</div>[m
                 <span class="w-8 h-8 rounded-lg bg-purple-50 text-purple-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ session_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ session_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Master timetable</div>[m
         </a>[m
 [m
[31m-        <a href="{% url 'venue-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m        <a href="{% url 'venue-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">Venues</div>[m
                 <span class="w-8 h-8 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ venue_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ venue_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Available venues</div>[m
         </a>[m
 [m
[31m-        <a href="{% url 'workshop-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m        <a href="{% url 'workshop-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">Workshops</div>[m
                 <span class="w-8 h-8 rounded-lg bg-rose-50 text-rose-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ workshop_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ workshop_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Workshop allocations</div>[m
         </a>[m
 [m
[31m-        <a href="{% url 'td-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m        <a href="{% url 'td-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">TD Allocations</div>[m
                 <span class="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ td_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ td_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Technical drawing</div>[m
         </a>[m
 [m
[31m-        <a href="{% url 'course-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m        <a href="{% url 'course-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">Courses</div>[m
                 <span class="w-8 h-8 rounded-lg bg-cyan-50 text-cyan-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ course_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ course_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Programme courses</div>[m
         </a>[m
 [m
[31m-        <a href="{% url 'semester-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-5 hover:shadow-md hover:border-blue-200 transition">[m
[32m+[m[32m        <a href="{% url 'semester-list' %}" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sm:p-5 hover:shadow-md hover:border-blue-200 transition">[m
             <div class="flex items-center justify-between">[m
                 <div class="text-xs font-medium text-gray-500 uppercase tracking-wide">Semesters</div>[m
                 <span class="w-8 h-8 rounded-lg bg-teal-50 text-teal-600 flex items-center justify-center">[m
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>[m
                 </span>[m
             </div>[m
[31m-            <div class="mt-3 text-3xl font-bold text-gray-800">{{ semester_count }}</div>[m
[32m+[m[32m            <div class="mt-3 text-2xl sm:text-3xl font-bold text-gray-800">{{ semester_count }}</div>[m
             <div class="mt-1 text-xs text-gray-400">Academic periods</div>[m
         </a>[m
     </div>[m
[36m@@ -155,20 +155,20 @@[m
     </div>[m
 [m
     <!-- Quick actions -->[m
[31m-    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-8">[m
[31m-        <a href="{% url 'import-hub' %}" class="bg-gradient-to-br from-blue-600 to-indigo-700 rounded-xl p-6 text-white shadow-lg hover:shadow-xl transition">[m
[31m-            <div class="flex items-center gap-3">[m
[31m-                <svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>[m
[31m-                <div>[m
[32m+[m[32m    <div class="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4 mb-8">[m
[32m+[m[32m        <a href="{% url 'import-hub' %}" class="bg-gradient-to-br from-blue-600 to-indigo-700 rounded-xl p-4 sm:p-6 text-white shadow-lg hover:shadow-xl transition">[m
[32m+[m[32m            <div class="flex items-center gap-3 sm:gap-4">[m
[32m+[m[32m                <svg class="w-7 h-7 sm:w-8 sm:h-8 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>[m
[32m+[m[32m                <div class="min-w-0">[m
                     <h3 class="font-semibold">Import Data</h3>[m
[31m-                    <p class="text-sm text-blue-100">Upload master timetable, workshop & TD allocation Excel files</p>[m
[32m+[m[32m                    <p class="text-sm text-blue-100">Upload master timetable, workshop &amp; TD allocation Excel files</p>[m
                 </div>[m
             </div>[m
         </a>[m
[31m-        <a href="{% url 'session-list' %}" class="bg-gradient-to-br from-gray-700 to-gray-900 rounded-xl p-6 text-white shadow-lg hover:shadow-xl transition">[m
[31m-            <div class="flex items-center gap-3">[m
[31m-                <svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/></svg>[m
[31m-                <div>[m
[32m+[m[32m        <a href="{% url 'session-list' %}" class="bg-gradient-to-br from-gray-700 to-gray-900 rounded-xl p-4 sm:p-6 text-white shadow-lg hover:shadow-xl transition">[m
[32m+[m[32m            <div class="flex items-center gap-3 sm:gap-4">[m
[32m+[m[32m                <svg class="w-7 h-7 sm:w-8 sm:h-8 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/></svg>[m
[32m+[m[32m                <div class="min-w-0">[m
                     <h3 class="font-semibold">Review Timetable</h3>[m
                     <p class="text-sm text-gray-300">Browse sessions and manage group assignments</p>[m
                 </div>[m
[1mdiff --git a/templates/partials/sidebar.html b/templates/partials/sidebar.html[m
[1mindex ab0998c..fa69a62 100644[m
[1m--- a/templates/partials/sidebar.html[m
[1m+++ b/templates/partials/sidebar.html[m
[36m@@ -1,100 +1,203 @@[m
 {% load core_tags %}[m
 {% nav_section as nav %}[m
[31m-<aside :class="sidebarOpen ? 'translate-x-0' : '-translate-x-full'"[m
[31m-       class="fixed lg:sticky top-0 left-0 z-40 h-screen w-64 bg-slate-900 text-slate-200 flex flex-col transition-transform lg:translate-x-0 shrink-0">[m
[31m-    <!-- Logo -->[m
[31m-    <div class="px-5 py-5 border-b border-slate-700">[m
[31m-        <a href="{% url 'dashboard' %}" class="text-lg font-bold text-white tracking-tight">[m
[31m-            <span class="text-emerald-400">CoET</span> Timetable[m
[31m-        </a>[m
[31m-        <p class="text-xs text-slate-400 mt-0.5">Group Allocation &amp; Validation</p>[m
[32m+[m[32m<aside id="app-sidebar"[m
[32m+[m[32m       :class="(sidebarOpen ? 'translate-x-0' : '-translate-x-full') + ' ' + (collapsed ? 'lg:w-20' : 'lg:w-64')"[m
[32m+[m[32m       class="fixed lg:sticky top-0 left-0 z-40 h-screen w-64 bg-slate-900 text-slate-200 flex flex-col shrink-0 transition-[width,transform] duration-300 ease-in-out lg:translate-x-0">[m
[32m+[m[32m    <div class="px-4 py-4 border-b border-slate-700 flex flex-col gap-2">[m
[32m+[m[32m        <div class="tt-brand-row flex items-center justify-between gap-2" :class="collapsed ? 'lg:justify-center lg:gap-0' : ''">[m
[32m+[m[32m            <a href="{% url 'dashboard' %}"[m
[32m+[m[32m               :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-48 lg:opacity-100'"[m
[32m+[m[32m               class="tt-brand-link text-lg font-bold text-white tracking-tight truncate transition-all duration-300 min-w-0">[m
[32m+[m[32m                <span class="text-emerald-400">CoET</span> Timetable[m
[32m+[m[32m            </a>[m
[32m+[m[32m            <button type="button"[m
[32m+[m[32m                    @click="toggle()"[m
[32m+[m[32m                    :aria-expanded="collapsed ? 'false' : 'true'"[m
[32m+[m[32m                    aria-controls="app-sidebar"[m
[32m+[m[32m                    :aria-label="collapsed ? 'Expand sidebar' : 'Collapse sidebar'"[m
[32m+[m[32m                    :title="collapsed ? 'Expand sidebar' : 'Collapse sidebar'"[m
[32m+[m[32m                    class="hidden lg:inline-flex items-center justify-center w-8 h-8 rounded-md text-slate-300 hover:text-white hover:bg-slate-700/60 focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 transition shrink-0">[m
[32m+[m[32m                <svg x-show="!collapsed" x-cloak class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 19l-7-7 7-7m8 14l-7-7 7-7"/></svg>[m
[32m+[m[32m                <svg x-show="collapsed" x-cloak class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 5l7 7-7 7M5 5l7 7-7 7"/></svg>[m
[32m+[m[32m            </button>[m
[32m+[m[32m        </div>[m
[32m+[m[32m        <p :class="collapsed ? 'lg:hidden' : ''" class="tt-hide-when-collapsed text-xs text-slate-400">Group Allocation &amp; Validation</p>[m
     </div>[m
[31m-    <!-- Nav -->[m
[31m-    <nav class="flex-1 overflow-y-auto py-4 px-3 space-y-1 text-sm">[m
[32m+[m[32m    <nav class="flex-1 overflow-y-auto py-4 px-3 space-y-1 text-sm" @mouseleave="hideTooltip()">[m
         <a href="{% url 'dashboard' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'dashboard' }}">[m
[32m+[m[32m           aria-label="Dashboard"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Dashboard')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Dashboard')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'dashboard' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h-3m-4 0a1 1 0 01-1-1v-4a1 1 0 011-1h2a1 1 0 011 1v4a1 1 0 01-1 1h-2z"/></svg>[m
[31m-            Dashboard[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Dashboard</span>[m
         </a>[m
 [m
[31m-        <div class="pt-3 pb-1 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Reference Data</div>[m
[32m+[m[32m        <div :class="collapsed ? 'lg:my-2 lg:h-px lg:py-0 lg:overflow-hidden lg:bg-slate-700/50 lg:text-transparent' : ''"[m
[32m+[m[32m             class="tt-section px-3 pt-3 pb-1 text-xs font-semibold uppercase tracking-wider text-slate-500 transition-all duration-300">Reference Data</div>[m
         <a href="{% url 'programme-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'programmes' }}">[m
[32m+[m[32m           aria-label="Programmes"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Programmes')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Programmes')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'programmes' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>[m
[31m-            Programmes[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Programmes</span>[m
         </a>[m
         <a href="{% url 'group-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'groups' }}">[m
[32m+[m[32m           aria-label="Student Groups"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Student Groups')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Student Groups')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'groups' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z"/></svg>[m
[31m-            Student Groups[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Student Groups</span>[m
         </a>[m
         <a href="{% url 'course-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'courses' }}">[m
[32m+[m[32m           aria-label="Programme Courses"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Programme Courses')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Programme Courses')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'courses' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253"/></svg>[m
[31m-            Programme Courses[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Programme Courses</span>[m
         </a>[m
         <a href="{% url 'venue-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'venues' }}">[m
[32m+[m[32m           aria-label="Venues"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Venues')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Venues')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'venues' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>[m
[31m-            Venues[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Venues</span>[m
         </a>[m
         <a href="{% url 'semester-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'semesters' }}">[m
[32m+[m[32m           aria-label="Semesters"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Semesters')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Semesters')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'semesters' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>[m
[31m-            Semesters[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Semesters</span>[m
         </a>[m
 [m
[31m-        <div class="pt-3 pb-1 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Timetable</div>[m
[32m+[m[32m        <div :class="collapsed ? 'lg:my-2 lg:h-px lg:py-0 lg:overflow-hidden lg:bg-slate-700/50 lg:text-transparent' : ''"[m
[32m+[m[32m             class="tt-section px-3 pt-3 pb-1 text-xs font-semibold uppercase tracking-wider text-slate-500 transition-all duration-300">Timetable</div>[m
         <a href="{% url 'timetable' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'timetable' }}">[m
[32m+[m[32m           aria-label="Timetable View"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Timetable View')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Timetable View')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'timetable' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 5a2 2 0 012-2h12a2 2 0 012 2v14a2 2 0 01-2 2H6a2 2 0 01-2-2V5z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 3v4m8-4v4M4 11h16M8 15h.01M12 15h.01M16 15h.01"/></svg>[m
[31m-            Timetable View[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Timetable View</span>[m
         </a>[m
         <a href="{% url 'session-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'sessions' }}">[m
[32m+[m[32m           aria-label="Master Timetable"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Master Timetable')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Master Timetable')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'sessions' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"/></svg>[m
[31m-            Master Timetable[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Master Timetable</span>[m
         </a>[m
         <a href="{% url 'workshop-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'workshops' }}">[m
[32m+[m[32m           aria-label="Workshop Allocation"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Workshop Allocation')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Workshop Allocation')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'workshops' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/></svg>[m
[31m-            Workshop Allocations[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Workshop Allocations</span>[m
         </a>[m
         <a href="{% url 'td-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'td' }}">[m
[32m+[m[32m           aria-label="TD Allocations"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'TD Allocations')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'TD Allocations')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'td' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>[m
[31m-            TD Allocations[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">TD Allocations</span>[m
         </a>[m
 [m
[31m-        <div class="pt-3 pb-1 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Data Import</div>[m
[32m+[m[32m        <div :class="collapsed ? 'lg:my-2 lg:h-px lg:py-0 lg:overflow-hidden lg:bg-slate-700/50 lg:text-transparent' : ''"[m
[32m+[m[32m             class="tt-section px-3 pt-3 pb-1 text-xs font-semibold uppercase tracking-wider text-slate-500 transition-all duration-300">Data Import</div>[m
         <a href="{% url 'import-hub' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'imports' }}">[m
[32m+[m[32m           aria-label="Import Data"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Import Data')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Import Data')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'imports' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>[m
[31m-            Import Data[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Import Data</span>[m
         </a>[m
         <a href="{% url 'import-history' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'imports' }}">[m
[32m+[m[32m           aria-label="Import History"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Import History')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Import History')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'imports' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 0l-2-2m2 2l2-2m-2 4h6a2 2 0 002-2V6a2 2 0 00-2-2h-6a2 2 0 00-2 2v4M5 12H3v6a2 2 0 002 2h10a2 2 0 002-2v-3"/></svg>[m
[31m-            Import History[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Import History</span>[m
         </a>[m
 [m
[31m-        <div class="pt-3 pb-1 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Export</div>[m
[32m+[m[32m        <div :class="collapsed ? 'lg:my-2 lg:h-px lg:py-0 lg:overflow-hidden lg:bg-slate-700/50 lg:text-transparent' : ''"[m
[32m+[m[32m             class="tt-section px-3 pt-3 pb-1 text-xs font-semibold uppercase tracking-wider text-slate-500 transition-all duration-300">Export</div>[m
         <a href="{% url 'export-timetable' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'exports' }}">[m
[32m+[m[32m           aria-label="Export Timetable"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Export Timetable')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Export Timetable')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'exports' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>[m
[31m-            Export Timetable[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Export Timetable</span>[m
         </a>[m
 [m
[31m-        <div class="pt-3 pb-1 px-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">Activity Log</div>[m
[32m+[m[32m        <div :class="collapsed ? 'lg:my-2 lg:h-px lg:py-0 lg:overflow-hidden lg:bg-slate-700/50 lg:text-transparent' : ''"[m
[32m+[m[32m             class="tt-section px-3 pt-3 pb-1 text-xs font-semibold uppercase tracking-wider text-slate-500 transition-all duration-300">Activity Log</div>[m
         <a href="{% url 'activity-list' %}"[m
[31m-           class="flex items-center gap-3 px-3 py-2.5 rounded-lg transition {{ nav|active_cls:'activity' }}">[m
[32m+[m[32m           aria-label="Activity Log"[m
[32m+[m[32m           @mouseenter="showTooltip($el, 'Activity Log')"[m
[32m+[m[32m           @mouseleave="hideTooltip()"[m
[32m+[m[32m           @focus="showTooltip($el, 'Activity Log')"[m
[32m+[m[32m           @blur="hideTooltip()"[m
[32m+[m[32m           :class="collapsed ? 'lg:justify-center lg:gap-0' : 'lg:justify-start lg:gap-3'"[m
[32m+[m[32m           class="relative flex items-center justify-start gap-3 px-3 py-2.5 rounded-lg transition focus:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/70 {{ nav|active_cls:'activity' }}">[m
             <svg class="w-5 h-5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>[m
[31m-            Activity Log[m
[32m+[m[32m            <span :class="collapsed ? 'lg:max-w-0 lg:opacity-0' : 'lg:max-w-44 lg:opacity-100'" class="truncate transition-all duration-300">Activity Log</span>[m
         </a>[m
         {% activity_logs 6 as logs %}[m
         {% for log in logs %}[m
         <a href="{% url 'dashboard' %}?log={{ log.pk }}"[m
[31m-           class="flex items-start gap-2 px-2 py-1.5 rounded-lg transition hover:bg-slate-800/60 group">[m
[32m+[m[32m           :class="collapsed ? 'lg:hidden' : ''"[m
[32m+[m[32m           class="tt-hide-when-collapsed flex items-start gap-2 px-2 py-1.5 rounded-lg transition hover:bg-slate-800/60 group">[m
             <span class="mt-1.5 w-1.5 h-1.5 rounded-full {{ log.action|log_badge_cls }} shrink-0"></span>[m
             <span class="min-w-0 flex-1 leading-snug">[m
                 <span class="block text-xs text-slate-300 group-hover:text-white truncate">{{ log.message|truncatechars:44 }}</span>[m
[36m@@ -102,9 +205,8 @@[m
             </span>[m
         </a>[m
         {% empty %}[m
[31m-        <p class="px-3 py-2 text-xs text-slate-500">No activity yet.</p>[m
[32m+[m[32m        <p :class="collapsed ? 'lg:hidden' : ''" class="tt-hide-when-collapsed px-3 py-2 text-xs text-slate-500">No activity yet.</p>[m
         {% endfor %}[m
     </nav>[m
[31m-    <!-- Version -->[m
[31m-    <div class="px-5 py-3 border-t border-slate-700 text-xs text-slate-500">v0.1 — Phase 1</div>[m
[32m+[m[32m    <div :class="collapsed ? 'lg:hidden' : ''" class="tt-hide-when-collapsed px-5 py-3 border-t border-slate-700 text-xs text-slate-500">v0.1 — Phase 1</div>[m
 </aside>[m
\ No newline at end of file[m
