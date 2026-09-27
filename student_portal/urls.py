from django.urls import path

from . import views

urlpatterns = [
    path("", views.student_home, name="student-home"),
    path("student/timetable.pdf/", views.student_timetable_pdf, name="student-timetable-pdf"),
    path("reports/collision/", views.collision_report, name="collision-report"),
    path("reports/collision/thanks/<uuid:reference>/", views.collision_report_thanks, name="collision-report-thanks"),
    path("staff/collision-reports/", views.collision_report_list, name="collision-report-list"),
    path("staff/collision-reports/<uuid:reference>/", views.collision_report_detail, name="collision-report-detail"),
]
