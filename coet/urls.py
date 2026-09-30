from django.contrib import admin
from django.urls import include, path
from core import views as core_views

from django.contrib.auth.views import LogoutView
from student_portal.auth import StaffLoginView

urlpatterns = [
    path("admin/", admin.site.urls),
    path("staff/login/", StaffLoginView.as_view(), name="staff-login"),
    path("staff/logout/", LogoutView.as_view(), name="staff-logout"),
    path("", include("student_portal.urls")),
    path("", include("core.urls")),
    # The audit app owns its own URLs rather than nesting under core: its URLs
    # are named "audit-*" and are claimed by nav_section for the Allocation
    # section, which would be a lie if they lived in core/urls.py.
    path("", include("audit.urls")),
    # The management dashboard moves to /staff/; reverse('dashboard') follows
    # this alias while the public student portal owns the site root.
    path("staff/", core_views.dashboard, name="dashboard"),
]
