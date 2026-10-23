from django.contrib import admin
from django.urls import include, path
from core import views as core_views

from django.contrib.auth.views import (LogoutView, PasswordResetView,
    PasswordResetDoneView, PasswordResetConfirmView, PasswordResetCompleteView)
from student_portal.auth import StaffLoginView

urlpatterns = [
    path("api/", include("api.urls")),
    path("admin/", admin.site.urls),
    path("staff/login/", StaffLoginView.as_view(), name="staff-login"),
    path("staff/logout/", LogoutView.as_view(), name="staff-logout"),
    path("staff/password-reset/", PasswordResetView.as_view(template_name="registration/password_reset_form.html"), name="password_reset"),
    path("staff/password-reset/done/", PasswordResetDoneView.as_view(), name="password_reset_done"),
    path("staff/password-reset-confirm/<uidb64>/<token>/", PasswordResetConfirmView.as_view(), name="password_reset_confirm"),
    path("staff/password-reset/complete/", PasswordResetCompleteView.as_view(), name="password_reset_complete"),
    path("", include("student_portal.urls")),
    path("", include("core.urls")),
    # The management dashboard moves to /staff/; reverse('dashboard') follows
    # this alias while the public student portal owns the site root.
    path("staff/", core_views.dashboard, name="dashboard"),
]
