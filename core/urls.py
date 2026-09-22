from django.urls import path

from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    # Programme
    path("programmes/", views.programme_list, name="programme-list"),
    path(
        "programmes/clear-all/",
        views.programme_clear_all,
        name="programme-clear-all",
    ),
    path("programmes/create/", views.programme_create, name="programme-create"),
    path("programmes/<int:pk>/", views.programme_detail, name="programme-detail"),
    path("programmes/<int:pk>/edit/", views.programme_edit, name="programme-edit"),
    path("programmes/<int:pk>/delete/", views.programme_delete, name="programme-delete"),
    path(
        "programmes/<int:pk>/timetable.pdf/",
        views.programme_timetable_pdf,
        name="programme-timetable-pdf",
    ),
    # Student Groups
    path("groups/", views.studentgroup_list, name="group-list"),
    path("groups/clear-all/", views.group_clear_all, name="group-clear-all"),
    path("groups/create/", views.studentgroup_create, name="group-create"),
    path("groups/<int:pk>/", views.studentgroup_detail, name="group-detail"),
    path("groups/<int:pk>/edit/", views.studentgroup_edit, name="group-edit"),
    path("groups/<int:pk>/delete/", views.studentgroup_delete, name="group-delete"),
    # Venues
    path("venues/", views.venue_list, name="venue-list"),
    path("venues/clear-all/", views.venue_clear_all, name="venue-clear-all"),
    path("venues/create/", views.venue_create, name="venue-create"),
    path("venues/recycle/", views.venue_recycle, name="venue-recycle"),
    path("venues/fix-conflict/", views.venue_fix_conflict, name="venue-fix-conflict"),
    path("venues/<int:pk>/", views.venue_detail, name="venue-detail"),
    path("venues/<int:pk>/edit/", views.venue_edit, name="venue-edit"),
    path("venues/<int:pk>/delete/", views.venue_delete, name="venue-delete"),
    # Semesters
    path("semesters/", views.semester_list, name="semester-list"),
    path("semesters/create/", views.semester_create, name="semester-create"),
    path("semesters/<int:pk>/", views.semester_detail, name="semester-detail"),
    path("semesters/<int:pk>/edit/", views.semester_edit, name="semester-edit"),
    path("semesters/<int:pk>/delete/", views.semester_delete, name="semester-delete"),
    # Sessions (Master Timetable)
    path("sessions/", views.session_list, name="session-list"),
    path("sessions/clear-all/", views.session_clear_all, name="session-clear-all"),
    path(
        "sessions/assign-lecture-groups/",
        views.session_assign_lecture_groups,
        name="session-assign-lecture-groups",
    ),
    path("sessions/create/", views.session_create, name="session-create"),
    path("sessions/<int:pk>/", views.session_detail, name="session-detail"),
    path("sessions/<int:pk>/edit/", views.session_edit, name="session-edit"),
    path("sessions/<int:pk>/delete/", views.session_delete, name="session-delete"),
    path(
        "sessions/<int:pk>/add-group/",
        views.session_add_group,
        name="session-add-group",
    ),
    path(
        "sessions/<int:pk>/remove-group/<int:group_pk>/",
        views.session_remove_group,
        name="session-remove-group",
    ),
    path(
        "sessions/<int:pk>/remove-programme-groups/<int:programme_pk>/",
        views.session_remove_programme_groups,
        name="session-remove-programme-groups",
    ),
    path(
        "sessions/<int:pk>/clear-groups/",
        views.session_clear_groups,
        name="session-clear-groups",
    ),
    # Workshop Allocations
    path("workshops/", views.workshop_list, name="workshop-list"),
    path(
        "workshops/clear-all/",
        views.workshop_clear_all,
        name="workshop-clear-all",
    ),
    path("workshops/create/", views.workshop_create, name="workshop-create"),
    path("workshops/<int:pk>/", views.workshop_detail, name="workshop-detail"),
    path("workshops/<int:pk>/edit/", views.workshop_edit, name="workshop-edit"),
    path(
        "workshops/<int:pk>/delete/",
        views.workshop_delete,
        name="workshop-delete",
    ),
    # Technical Drawing Allocations
    path("td/", views.td_list, name="td-list"),
    path("td/clear-all/", views.td_clear_all, name="td-clear-all"),
    path("td/create/", views.td_create, name="td-create"),
    path("td/<int:pk>/", views.td_detail, name="td-detail"),
    path("td/<int:pk>/edit/", views.td_edit, name="td-edit"),
    path("td/<int:pk>/delete/", views.td_delete, name="td-delete"),
    # Programme Courses
    path("courses/", views.course_list, name="course-list"),
    path("courses/clear-all/", views.course_clear_all, name="course-clear-all"),
    path("courses/create/", views.course_create, name="course-create"),
    path("courses/<int:pk>/", views.course_detail, name="course-detail"),
    path("courses/<int:pk>/edit/", views.course_edit, name="course-edit"),
    path("courses/<int:pk>/delete/", views.course_delete, name="course-delete"),
    # Imports
    path("import/", views.import_hub, name="import-hub"),
    path("import/history/", views.import_history, name="import-history"),
    path(
        "import/history/<int:pk>/",
        views.import_history_detail,
        name="import-history-detail",
    ),
    path(
        "import/history/<slug:import_type>/",
        views.import_history,
        name="import-history-type",
    ),
    path("import/<slug:import_type>/", views.import_upload, name="import-upload"),
    # Exports
    path("export/", views.export_timetable, name="export-timetable"),
    path(
        "export/programmes/<int:pk>/timetable.pdf/",
        views.programme_timetable_pdf,
        name="programme-timetable-export",
    ),
    path(
        "export/groups/<int:pk>/timetable.pdf/",
        views.group_timetable_pdf,
        name="group-timetable-export",
    ),
    # Timetable display
    path("timetable/", views.timetable_view, name="timetable"),
    path("timetable/groups/", views.timetable_groups_json, name="timetable-groups"),
    # Activity Log
    path("activity/", views.activity_list, name="activity-list"),
]
