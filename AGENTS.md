# AGENTS.md

Django 5.2 (Python 3.13) timetable app for COET. The only active app is `core`; `allocation` is a stub (empty models, no migration). Project templates live in `templates/` (global `DIRS`) with shared partials under `templates/core/`.

## Commands

All from repo root (`E:\COET_TIMETABLE\Coet_timetable`). Use `.venv\Scripts\python.exe` — a second venv `venv/` also exists, both Python 3.13.

- Dev server: `.venv\Scripts\python.exe manage.py runserver`
- Sample data import requires `--file` (README omits the flags), and order matters:
  ```
  manage.py import_semesters --file sample_data/semesters.xlsx
  manage.py import_programmes --file sample_data/programmes.xlsx
  manage.py import_student_groups --file sample_data/student_groups.xlsx
  manage.py import_programme_courses --file sample_data/programme_courses.xlsx
  manage.py import_venues --file sample_data/venues.xlsx
  manage.py import_master_timetable --file sample_data/master_timetable.xlsx --semester 1
  manage.py import_workshop_allocation --file sample_data/workshop_allocation.xlsx --semester 1
  manage.py import_td_allocation --file sample_data/td_allocation.xlsx --semester 1
  ```
  `create_sample_data.py` regenerates those Excel files and prints the full command sequence. Import commands `exit(1)` on any row error. Master-timetable imports support a `--dry-run` flag (reconcilies/validates but writes nothing; also prints `DRY RUN — no records were written.`).
- Test suite: `.venv\Scripts\python.exe manage.py test` runs ~25 tests in `core/tests.py` (import reconciliation/idempotency, CSRF-enforced CRUD responses for htmx and plain posts, sidebar active-nav, upload view). Uses an isolated test DB — `db.sqlite3` is never touched. The first run historically failed on `set - dict` / `ImporterTestCase` misuse; keep uses of `make_xlsx` (module function) rather than a per-instance `_xlsx`.

## Conventions when adding features

- CRUD is hand-rolled and repetitive in `core/views.py`: one resource = list + create + edit + delete + detail views, 5 URL patterns (`core/urls.py`), and templates. Copy an existing block (e.g. venues) instead of writing class-based views.
- List views filter on `?q=` plus resource-specific filters, and render the partial `core/_table_and_cards.html` when the `HX-Request` header is present, otherwise the full `core/list.html`.
- Write views return `HttpResponse("")` with an `HX-Trigger` header on success (`close-modal,refresh-table`, plus `refresh-detail` for edits) via the shared `_write_response(request, trigger, redirect_name, *args)` helper, which also does a plain `redirect()` for non-htmx POSTs so saves never appear to hang. The browser modal/slide-panel JS in `templates/base.html` is wired to those triggers and auto-refreshes — keep using them.
- Frontend is htmx (2.0.4) + Alpine.js + Tailwind, all loaded from CDNs (requires internet, no bundler). Form templates post with `hx-target="#modal-content"`. After innerHTML swaps, `base.html` runs `Alpine.initTree` on the swapped element (via `htmx:afterSwap`) — page-level Alpine state lives on `<body x-data>` so it must be hoisted there, not on a swappable child. Sidebar active state comes from the `nav_section`/`active_cls` tags in `core_tags.py` (derived from `request.resolver_match.url_name`); per-view `NAV` context keys no longer exist.
- Style form widgets with the existing `INPUT_CLS` / `SELECT_CLS` constants in `core/forms.py`. Session create/edit uses the inline `SessionGroupFormSet`.

## Data model / importers

- `core/importers.py` returns an `ImportResult` (created/updated/skipped/errors/conflicts, plus reconciliation fields). Master-timetable import (`reconcile_master_timetable`, `import_master_timetable_from_excel`) is idempotent, requires a valid `Semester` (blocks otherwise), validates all `MASTER_REQUIRED` columns, auto-creates missing venues with capacity 0 on real imports (not dry-run), expands the legacy group code `ALL` via `ProgrammeCourse`→programme→`StudentGroup` (unresolvable ALL rows become conflicts, never silently dropped), and reports missing courses/groups/venues. Workshop/TD imports use `update_or_create` on a full natural key with `defaults={}` so re-imports are no-ops. All session/workshop/TD imports target a `Semester` by PK (`--semester`); for workshop, `None` means auto-detect from the workbook title (flat files fall back to semester 1).
- Workshop import accepts BOTH formats based on a content sniff (`core/workshop_parser.py`): the legacy flat columns (`course_code, group_code, day, start_time, end_time, venue`) and the raw university Workshop Schedule matrix workbook (GROUPS/POSITION/SCHEDULE/KEY layout) parsed by `parse_workbook` → `WorkshopRecord` intermediates (`reconcile_workshop_workbook` = dry preview; `import_university_workshop_workbook` = write, `--dry-run` supported). The matrix read derives course_code/workshop from the workshop category, day+`time_period` from the KEY legend positions (no invented clock times → `start_time`/`end_time` are NULL), week_start/week_end from the contiguous week runs per section, and blocks if the referenced `Semester` is missing; group codes without a matching `StudentGroup` are reported, never created. Raw stray cells (e.g. `Q18=4`) land in `result.unrecognized_cells`.
- `core/models.py`: Semester (unique year+sem), Programme, StudentGroup, ProgrammeCourse, Venue, Session, SessionGroup (through M2M), WorkshopAllocation, TechnicalDrawingAllocation. Activity/day are stored as UPPERCASE `TextChoices` (`LECTURE`, `MONDAY`, ...).
- SQLite (`db.sqlite3`, gitignored). After model changes run `makemigrations core` and `migrate`.

## Gotchas

- `manage.py check` reports `staticfiles.W004` because `static/` doesn't exist but is in `STATICFILES_DIRS`. Harmless — all styling comes from CDNs.
- `settings.py` holds local-only, uncommitted LAN IPs in `ALLOWED_HOSTS` (172.16.185.226, 192.168.100.6) — don't strip them.
- README is stale: PDF report generation (reportlab/weasyprint) is claimed but no report code exists — those packages are installed but unused.
- `core/templatetags/core_tags.py` defines `get_attr` and `time_short` filters, plus the `nav_section`/`active_cls` sidebar tags used by `templates/partials/sidebar.html`.