import json
import secrets

from django.conf import settings
from django.http import JsonResponse


def api_key_ok(request):
    expected = getattr(settings, "API_KEY", "") or ""
    provided = request.headers.get("X-API-Key", "")
    if not expected or not provided:
        return False
    return secrets.compare_digest(provided, expected)


def require_api_key(view):
    def wrapped(request, *args, **kwargs):
        if not api_key_ok(request):
            response = JsonResponse({"detail": "A valid API key is required."}, status=401)
            response["WWW-Authenticate"] = "Api-Key"
            return response
        return view(request, *args, **kwargs)

    return wrapped
