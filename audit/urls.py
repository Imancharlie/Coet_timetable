"""Audit URLs.

Mounted at the project root, alongside ``core.urls`` and ``student_portal``:
an audit report belongs to no single resource, and the URL names are prefixed
with ``audit-`` so ``nav_section`` can claim them for the Allocation section
without colliding with anything in ``core``.
"""

from django.urls import path

from audit import views

urlpatterns = [
    # Generate. POST only — it writes a row and prints a PDF.
    path(
        "allocation-runs/<int:run_pk>/audit/",
        views.run_audit,
        name="allocation-run-audit",
    ),
    # List every report, and read/export one.
    # NOTE: "audit-reports/" is registered before the "<int:pk>/" pattern only
    # for readability; the int converter already disambiguates them.
    path("audit-reports/", views.report_index, name="audit-reports"),
    path("audit-reports/<int:pk>/", views.report_detail, name="audit-report-detail"),
    path("audit-reports/<int:pk>/pdf/", views.report_pdf, name="audit-report-pdf"),
    path("audit-reports/<int:pk>/json/", views.report_json, name="audit-report-json"),
]
