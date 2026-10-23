from django.conf import settings
from django.contrib.auth import logout
from django.shortcuts import redirect
from django.utils import timezone


PUBLIC_PATHS = {
    "/",
    "/reports/collision/",
    "/reports/collision/thanks/",
    "/student/timetable.pdf/",
    settings.LOGIN_URL,
    settings.LOGOUT_REDIRECT_URL,
    "/staff/password-reset/",
    "/staff/password-reset/done/",
    "/staff/password-reset-confirm/",
    "/staff/password-reset/complete/",
}


def _is_admin_path(path):
    """Django admin keeps its own login page and its own permission check."""
    return path == "/admin" or path.startswith("/admin/")


class StaffAccessMiddleware:
    """Keep management views private while allowing explicitly public student pages."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        path = request.path_info
        if _is_admin_path(path) or path.startswith("/api/"):
            return self.get_response(request)
        is_public = (
            path in PUBLIC_PATHS
            or path.startswith("/static/")
            or path.startswith("/reports/collision/thanks/")
        )
        if is_public:
            return self.get_response(request)

        user = request.user
        if not user.is_authenticated:
            return redirect(settings.LOGIN_URL)
        if not user.is_staff:
            return redirect("/")
        return self.get_response(request)


class StaffIdleTimeoutMiddleware:
    """Expire staff sessions after inactivity, enforced on the server."""

    SESSION_KEY = "staff_last_activity"

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        user = request.user
        if user.is_authenticated and user.is_staff:
            now = timezone.now().timestamp()
            previous = request.session.get(self.SESSION_KEY)
            timeout = getattr(settings, "STAFF_IDLE_TIMEOUT_SECONDS", 900)
            try:
                expired = previous is not None and now - float(previous) > timeout
            except (TypeError, ValueError):
                expired = True
            if expired:
                # Sign them out, then let the page they opened render. Sending
                # this request to the staff login page is what replaced the
                # page they came back to.
                logout(request)
            else:
                request.session[self.SESSION_KEY] = now
        return self.get_response(request)
