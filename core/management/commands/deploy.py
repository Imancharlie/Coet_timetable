"""Auto-deploy pipeline for COET Timetable.

Polls origin/<branch>, fast-forwards the working tree, syncs dependencies,
applies migrations, collects static files, restarts the gunicorn systemd
service and verifies health. Any failure rolls back automatically:

  * ``git reset --hard`` to the previously deployed commit
  * restore the pre-migrate ``pg_dump`` snapshot (PostgreSQL)
  * restart the service

Every run appends to a durable plain-text audit log (``DEPLOY_LOG_FILE``,
default ``/var/log/coet-deploy.log``) so even a rolled-back run stays
traceable after the database is restored.

Driven by ``coet-deploy.timer`` (every minute) or run by hand::

    venv/bin/python manage.py deploy --branch deploy --trigger manual
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError

# The settings module gunicorn actually boots with. Every deploy-time
# manage.py call must use it, otherwise migrations and system checks would run
# against settings that differ from the app serving traffic.
SETTINGS_MODULE = os.getenv("DEPLOY_SETTINGS_MODULE", "coet.production_settings")


def _run(cmd, log, cwd=None, env=None):
    """Run a command, echoing output, raising RuntimeError on failure."""
    proc = subprocess.run(
        cmd, cwd=cwd, env=env, capture_output=True, text=True,
    )
    for line in (proc.stdout + proc.stderr).strip().splitlines():
        log(line)
    if proc.returncode != 0:
        raise RuntimeError(f"{' '.join(cmd[:3])} failed ({proc.returncode})")
    return proc.stdout.strip()


def _split_log(args):
    """Callers pass the ``log`` callable as the final positional argument."""
    return args[:-1], args[-1]


def _git(base_dir, *args):
    argv, log = _split_log(args)
    return _run(["git", *argv], log, cwd=base_dir)


def _manage(base_dir, *args):
    argv, log = _split_log(args)
    env = dict(os.environ)
    env["DJANGO_SETTINGS_MODULE"] = SETTINGS_MODULE
    return _run([sys.executable, "manage.py", *argv], log, cwd=base_dir, env=env)


def _pg_dump(db, dest, log):
    """Write a plain-SQL snapshot of the live database to ``dest``.

    Uses ``--clean --if-exists`` so the file can be fed straight back to psql
    to drop and recreate the objects a failed migration left behind.
    """
    cmd = [
        "pg_dump",
        "-h", db.get("HOST") or "127.0.0.1",
        "-p", str(db.get("PORT") or 5432),
        "-U", db["USER"],
        "-d", db["NAME"],
        "--clean", "--if-exists",
    ]
    # postgres wants the password on stdin rather than in argv or the env of a
    # long-lived process.
    proc = subprocess.run(
        cmd, input=db["PASSWORD"] + "\n", capture_output=True, text=True,
        env=dict(os.environ, PGPASSWORD=db["PASSWORD"]),
    )
    for line in proc.stderr.strip().splitlines():
        log(line)
    if proc.returncode != 0:
        raise RuntimeError(f"pg_dump failed ({proc.returncode})")
    dest.write_text(proc.stdout, encoding="utf-8")
    if not dest.stat().st_size:
        raise RuntimeError("pg_dump produced an empty snapshot")


def _restart(service, log):
    try:
        _run(["systemctl", "restart", service], log)
    except FileNotFoundError:
        log("systemctl not found; skipping restart (local dev)")


def _wait_healthy(url, timeout, log):
    """Poll until the URL answers with a non-5xx status."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=10) as resp:
                if resp.status < 500:
                    log(f"health OK ({resp.status})")
                    return True
        except (urllib.error.URLError, OSError):
            pass
        time.sleep(3)
    return False


class Command(BaseCommand):
    help = "Pull origin/<branch> and deploy COET Timetable with automatic rollback."

    def add_arguments(self, parser):
        parser.add_argument("--branch", default=os.getenv("DEPLOY_BRANCH", "deploy"))
        parser.add_argument(
            "--trigger", default="timer", choices=["timer", "manual"],
            help="Where this run was initiated from (recorded in the audit log).",
        )
        parser.add_argument("--health-url", default="")
        parser.add_argument("--health-timeout", type=int, default=120)
        parser.add_argument(
            "--skip-deps", action="store_true",
            help="Do not pip install -r requirements.txt",
        )
        parser.add_argument(
            "--skip-restart", action="store_true",
            help="Do not restart the service or health check (CI runs).",
        )
        parser.add_argument(
            "--force", action="store_true",
            help="Deploy even when the branch is already checked out.",
        )

    def handle(self, *args, **options):  # noqa: C901 - linear pipeline
        base_dir = str(settings.BASE_DIR)
        branch = options["branch"]
        service = os.getenv("DEPLOY_SERVICE_NAME", "gunicorn-coet-timetable.service")
        health_url = options["health_url"] or os.getenv(
            "DEPLOY_HEALTH_URL", "https://coet-timetable.kodin.co.tz/"
        )
        audit_path = Path(os.getenv("DEPLOY_LOG_FILE", "/var/log/coet-deploy.log"))
        snapshot_dir = Path(
            os.getenv("DEPLOY_SNAPSHOT_DIR", str(Path(base_dir) / "deploy_snapshots"))
        )

        def log(line=""):
            self.stdout.write(str(line))

        def durable(line):
            try:
                audit_path.parent.mkdir(parents=True, exist_ok=True)
                with audit_path.open("a", encoding="utf-8") as fh:
                    fh.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {line}\n")
            except OSError:
                pass

        try:
            _git(base_dir, "fetch", "origin", branch, log)

            # The tree may sit on a detached HEAD (the initial deploy checked
            # out a raw commit), which would leave the branch unpinned and make
            # every later `git reset` land on an orphaned commit. Put the
            # working tree on the tracked branch before comparing SHAs.
            current_branch = subprocess.run(
                ["git", "rev-parse", "--abbrev-ref", "HEAD"],
                cwd=base_dir, capture_output=True, text=True,
            ).stdout.strip()
            if current_branch != branch:
                _git(base_dir, "checkout", "-B", branch, f"origin/{branch}", log)

            before = _git(base_dir, "rev-parse", "HEAD", log)
            target = _git(base_dir, "rev-parse", f"origin/{branch}", log)
        except RuntimeError as exc:
            raise CommandError(f"git fetch/rev-parse failed: {exc}") from exc

        if before == target and not options["force"]:
            log(f"Already up to date at {before[:12]} ({branch}); nothing to deploy")
            return

        started = time.time()
        durable(
            f"START {options['trigger']} deploy {branch} "
            f"{before[:12]} -> {target[:12]}"
        )

        snapshot = None
        migrations_started = False
        try:
            # pg_dump before touching migrations. A broken migration is the one
            # failure code rollback alone cannot fix, since git cannot undo a
            # schema change that already committed.
            snapshot_dir.mkdir(parents=True, exist_ok=True)
            snapshot = snapshot_dir / f"predeploy-{target[:12]}-{int(started)}.sql"
            db = settings.DATABASES["default"]
            log(f"Snapshotting database to {snapshot.name}...")
            _pg_dump(db, snapshot, log)
            log(f"Database snapshot written ({snapshot.stat().st_size} bytes)")

            _git(base_dir, "merge", "--ff-only", f"origin/{branch}", log)
            after = _git(base_dir, "rev-parse", "HEAD", log)
            log(f"Fast-forwarded to {after[:12]}")

            if not options["skip_deps"]:
                log("Syncing dependencies...")
                _run(
                    [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"],
                    log, cwd=base_dir,
                )

            log("Running system check...")
            _manage(base_dir, "check", log)

            log("Applying migrations...")
            migrations_started = True
            _manage(base_dir, "migrate", "--no-input", log)

            log("Collecting static files...")
            _manage(base_dir, "collectstatic", "--no-input", log)

            if not options["skip_restart"]:
                log(f"Restarting {service}...")
                _restart(service, log)
                log(f"Waiting for health at {health_url}...")
                if not _wait_healthy(health_url, options["health_timeout"], log):
                    raise RuntimeError("health check did not pass after restart")

            durable(f"SUCCESS {branch} {after[:12]} in {time.time() - started:.1f}s")
        except Exception as exc:  # noqa: BLE001 - any failure must roll back
            log(f"ERROR: {exc}")
            durable(f"ERROR {branch} {target[:12]}: {exc}")

            try:
                _git(base_dir, "reset", "--hard", before, log)
                log(f"Rolled code back to {before[:12]}")

                # Only restore the schema when migrations actually ran;
                # otherwise the snapshot would needlessly discard live writes
                # made since the snapshot was taken.
                if snapshot and snapshot.exists() and migrations_started:
                    db = settings.DATABASES["default"]
                    log(f"Restoring database from {snapshot.name}...")
                    proc = subprocess.run(
                        [
                            "psql", "-h", db.get("HOST") or "127.0.0.1",
                            "-p", str(db.get("PORT") or 5432),
                            "-U", db["USER"], "-d", db["NAME"],
                            "-v", "ON_ERROR_STOP=1", "-f", str(snapshot),
                        ],
                        input=db["PASSWORD"] + "\n", capture_output=True, text=True,
                        env=dict(os.environ, PGPASSWORD=db["PASSWORD"]),
                    )
                    if proc.returncode == 0:
                        log("Database restored")
                    else:
                        log(f"WARNING: database restore failed: {proc.stderr[:300]}")

                if not options["skip_restart"]:
                    _restart(service, log)
                    _wait_healthy(health_url, options["health_timeout"], log)
                durable(f"ROLLED BACK to {before[:12]}")
            except Exception as rb:  # noqa: BLE001
                log(f"ROLLBACK PROBLEMS: {rb}")
                durable(f"ROLLBACK PROBLEMS: {rb}")
                raise CommandError(f"deploy failed ({exc}) and rollback failed ({rb})") from rb

            raise CommandError(f"deploy failed and was rolled back: {exc}") from exc

        # Keep a few snapshots so post-mortem comparison is possible without
        # letting deploy_snapshots/ grow without bound on a small disk.
        try:
            snaps = sorted(
                snapshot_dir.glob("predeploy-*.sql"), key=lambda p: p.stat().st_mtime
            )
            for old in snaps[:-5]:
                old.unlink()
                log(f"Pruned old snapshot {old.name}")
        except OSError as exc:
            log(f"WARNING: snapshot pruning skipped: {exc}")