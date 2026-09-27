from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.views import LoginView
from django.core.cache import cache
from django.core.exceptions import ValidationError


class StaffAuthenticationForm(AuthenticationForm):
    """Add a temporary IP-based throttle to Django's standard login form."""

    MAX_FAILURES = 8
    WINDOW_SECONDS = 15 * 60

    def __init__(self, request=None, *args, **kwargs):
        super().__init__(request, *args, **kwargs)
        self.fields["username"].widget.attrs.update({
            "class": "field-control", "autocomplete": "username",
        })
        self.fields["password"].widget.attrs.update({
            "class": "field-control", "autocomplete": "current-password",
        })

    def clean(self):
        address = self.request.META.get("REMOTE_ADDR", "unknown") if self.request else "unknown"
        key = f"staff-login-failures:{address}"
        if cache.get(key, 0) >= self.MAX_FAILURES:
            raise ValidationError("Too many sign-in attempts. Wait 15 minutes and try again.")
        try:
            cleaned = super().clean()
        except ValidationError:
            cache.add(key, 0, self.WINDOW_SECONDS)
            cache.incr(key)
            raise
        cache.delete(key)
        return cleaned


class StaffLoginView(LoginView):
    template_name = "student_portal/login.html"
    authentication_form = StaffAuthenticationForm
    redirect_authenticated_user = True

    def form_valid(self, form):
        if not form.get_user().is_staff:
            form.add_error(None, "This account is not authorized for staff access.")
            return self.form_invalid(form)
        return super().form_valid(form)
