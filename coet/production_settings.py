"""Production settings for COET Timetable.

Kept as a separate module so ``coet/settings.py`` stays exactly as committed and
``git pull`` never conflicts. Everything dev-specific in the committed settings
(SQLite, no STATIC_ROOT, ALLOWED_HOSTS=['*']) is overridden here.

Configured entirely from the environment so no secret is ever written into the
repo. systemd loads the values from /var/www/coet-timetable/.env
(EnvironmentFile=), which is chmod 600.

    DJANGO_SECRET_KEY             required when DEBUG is off
    DJANGO_DEBUG                  default False in production
    DJANGO_ALLOWED_HOSTS          comma separated; no '*' allowed
    DJANGO_CSRF_TRUSTED_ORIGINS   comma separated, scheme included
    COET_DB_NAME / _USER / _PASSWORD / _HOST / _PORT
"""
import os
from pathlib import Path

from django.core.exceptions import ImproperlyConfigured

from .settings import *  # noqa: F401,F403  (base settings, overridden below)

BASE_DIR = Path(__file__).resolve().parent.parent


def _require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise ImproperlyConfigured(f"{name} must be set in the environment.")
    return value


# --- Security -------------------------------------------------------------
DEBUG = os.environ.get("DJANGO_DEBUG", "False").strip().lower() in {"1", "true", "yes"}

if not DEBUG:
    SECRET_KEY = _require("DJANGO_SECRET_KEY")
    hosts = [h.strip() for h in os.environ.get("DJANGO_ALLOWED_HOSTS", "").split(",") if h.strip()]
    if not hosts or "*" in hosts:
        raise ImproperlyConfigured("DJANGO_ALLOWED_HOSTS must list explicit hostnames.")
    ALLOWED_HOSTS = hosts

    origins = [o.strip() for o in os.environ.get("DJANGO_CSRF_TRUSTED_ORIGINS", "").split(",") if o.strip()]
    if not origins:
        raise ImproperlyConfigured("DJANGO_CSRF_TRUSTED_ORIGINS must be set.")
    CSRF_TRUSTED_ORIGINS = origins

    SECURE_SSL_REDIRECT = True
    SECURE_HSTS_SECONDS = 31536000
    SECURE_HSTS_INCLUDE_SUBDOMAINS = True
    SECURE_HSTS_PRELOAD = True
    SESSION_COOKIE_SECURE = True
    CSRF_COOKIE_SECURE = True


# --- Database -------------------------------------------------------------
# PostgreSQL in production: sqlite on a shared host serialises writes and loses
# data if the process is killed mid-transaction.
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.postgresql",
        "NAME": _require("COET_DB_NAME"),
        "USER": _require("COET_DB_USER"),
        "PASSWORD": _require("COET_DB_PASSWORD"),
        "HOST": os.environ.get("COET_DB_HOST", "127.0.0.1"),
        "PORT": os.environ.get("COET_DB_PORT", "5432"),
        "CONN_MAX_AGE": 60,
    }
}


# --- Static & media -------------------------------------------------------
# The committed settings points STATICFILES_DIRS at BASE_DIR/'static', which is
# not in the repo; collectstatic errors on a missing directory, so it is
# dropped here and everything is collected into staticfiles/.
STATIC_URL = "/static/"
STATIC_ROOT = BASE_DIR / "staticfiles"
STATICFILES_DIRS = []

MEDIA_URL = "/media/"
MEDIA_ROOT = BASE_DIR / "media"

STORAGES = {
    "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
    "staticfiles": {"BACKEND": "django.contrib.staticfiles.storage.StaticFilesStorage"},
}