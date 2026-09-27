from urllib.parse import urlencode

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
}


class StaffAccessMiddleware:
    """Keep management views private while allowing explicitly public student pages."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        path = request.path_info
        is_public = (
            path in PUBLIC_PATHS
            or path.startswith("/static/")
            or path.startswith("/reports/collision/thanks/")
        )
        if is_public:
            return self.get_response(request)

        user = request.user
        if not user.is_authenticated:
            query = urlencode({"next": request.get_full_path()})
            return redirect(f"{settings.LOGIN_URL}?{query}")
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
                logout(request)
                query = urlencode({"expired": "1", "next": request.get_full_path()})
                return redirect(f"{settings.LOGIN_URL}?{query}")
            request.session[self.SESSION_KEY] = now
        return self.get_response(request)
