"""Exercise the deploy command's rollback path with no network and no service.

Builds a throwaway git repo whose origin/deploy points at a commit containing a
syntax error while the working tree stays on the previous good commit, then runs
the real Command with pg_dump/psql, the service restart and the health poll
stubbed.

Asserts the properties that actually matter for a live site:

  1. a commit with a syntax error is rejected *before* the service is restarted,
     so the running site is never handed broken code
  2. the working tree is reset to the previously deployed commit
  3. the audit log records the failure and the rollback
  4. a database snapshot is taken before migrating
  5. a good commit deploys cleanly and does restart the service
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path
from unittest import mock

# Runnable both as `python core/tests_deploy.py` and as `manage.py test`, so the
# project root goes on sys.path before Django settings are imported.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "coet.settings")
os.environ.setdefault("DJANGO_SECRET_KEY", "test-only-not-a-real-secret")
os.environ.setdefault("DJANGO_ALLOWED_HOSTS", "localhost")
os.environ.setdefault("DJANGO_CSRF_TRUSTED_ORIGINS", "http://localhost")

import django  # noqa: E402

django.setup()

from core.management.commands import deploy as d  # noqa: E402

GIT_ENV = {
    **os.environ,
    "GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "t@t",
    "GIT_COMMITNER_NAME": "t", "GIT_COMMITNER_EMAIL": "t@t",
}


def sh(*args, cwd):
    return subprocess.run(
        args, cwd=cwd, env=GIT_ENV, check=True, capture_output=True, text=True
    ).stdout.strip()


def build_repo(tmp: Path):
    """origin/deploy = broken commit, working tree = the good commit."""
    origin = tmp / "origin.git"
    repo = tmp / "repo"
    sh("git", "init", "-q", "--bare", str(origin), cwd=tmp)
    repo.mkdir()
    sh("git", "init", "-q", "-b", "deploy", ".", cwd=repo)
    sh("git", "remote", "add", "origin", str(origin), cwd=repo)

    (repo / "app.py").write_text("VALUE = 'good'\n", encoding="utf-8")
    sh("git", "add", "-A", cwd=repo)
    sh("git", "commit", "-q", "-m", "good", cwd=repo)
    good_sha = sh("git", "rev-parse", "HEAD", cwd=repo)
    sh("git", "push", "-q", "origin", "deploy", cwd=repo)

    # Break the branch tip on origin only; leave the working tree on `good`.
    sh("git", "checkout", "-q", "-b", "broken", cwd=repo)
    (repo / "broken.py").write_text("this is not valid python !!!\n", encoding="utf-8")
    sh("git", "add", "-A", cwd=repo)
    sh("git", "commit", "-q", "-m", "broken", cwd=repo)
    # Push the broken tip onto origin/deploy while the working tree stays on
    # `good`. (Pushing "broken:deploy", not "deploy:deploy" -- the local
    # `deploy` ref still points at the good commit here.)
    sh("git", "push", "-q", "origin", "broken:deploy", cwd=repo)
    sh("git", "checkout", "-q", "--force", "deploy", cwd=repo)
    sh("git", "reset", "--hard", good_sha, cwd=repo)
    sh("git", "clean", "-fdq", cwd=repo)
    return repo, good_sha


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="coet-deploy-test-"))
    repo, good_sha = build_repo(tmp)
    good_short = good_sha[:7]

    audit = tmp / "audit.log"
    snaps = tmp / "snapshots"
    lock = tmp / "deploy.lock"

    restarts = []

    def fake_pg_dump(db, dest, log):
        dest.write_text("-- fake dump\nCREATE TABLE t (id int);\n", encoding="utf-8")

    def fake_restart(service, log):
        restarts.append(service)
        log(f"systemctl restart {service}")

    def fake_manage(base_dir, *args):
        log = args[-1]
        log(f"manage.py {' '.join(args[:-1])}")
        return "ok"

    def fake_health(url, timeout, log):
        log("health OK (200) [stub]")
        return True

    captured = []

    def attempt():
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with mock.patch.object(d, "_pg_dump", fake_pg_dump), \
             mock.patch.object(d, "_restart", fake_restart), \
             mock.patch.object(d, "_manage", fake_manage), \
             mock.patch.object(d, "_wait_healthy", fake_health), \
             mock.patch.object(d, "LOCK_PATH", lock), \
             mock.patch.dict(os.environ, {
                 "DEPLOY_LOG_FILE": str(audit),
                 "DEPLOY_SNAPSHOT_DIR": str(snaps),
             }), \
             mock.patch("django.conf.settings.BASE_DIR", repo):
            try:
                with redirect_stdout(buf):
                    d.Command().handle(
                        branch="deploy", trigger="timer", health_url="",
                        health_timeout=2, skip_deps=True, skip_restart=False,
                        force=False,
                    )
                captured.append(buf.getvalue())
                return None
            except Exception as exc:  # noqa: BLE001
                captured.append(buf.getvalue())
                return exc

    failures = []

    # ---------- broken commit ----------
    exc = attempt()
    text = audit.read_text(encoding="utf-8") if audit.exists() else ""

    if exc is None:
        failures.append("deploy of a broken commit should have raised")
    else:
        print(f"[ok] broken commit rejected: {type(exc).__name__}")

    # The only restart allowed is the one *after* the rollback. A restart before
    # validation would have handed the broken code to the live workers.
    if restarts:
        before_rollback = text.split("Rolled code back")[0]
        if "systemctl restart" in before_rollback:
            failures.append("service was restarted BEFORE the failure was detected")
        else:
            print("[ok] no restart before validation; restart only happened as part of rollback")
    else:
        print("[ok] service was never restarted")

    head = sh("git", "rev-parse", "HEAD", cwd=repo)
    if head != good_sha:
        failures.append(f"HEAD is {head[:7]}, expected the good commit {good_short}")
    else:
        print(f"[ok] code rolled back to {good_short}")

    if "ERROR" not in text or "ROLLED BACK" not in text:
        failures.append(f"audit log incomplete:\n{text}")
    else:
        print("[ok] audit log recorded ERROR + ROLLED BACK")

    if not list(snaps.glob("predeploy-*.sql")):
        failures.append("no pg_dump snapshot was taken before migrating")
    else:
        print("[ok] database snapshot taken before migrating")

    if "Byte-compiling" not in captured[0]:
        failures.append("byte-compile validation step did not run")
    else:
        print("[ok] byte-compile validation ran before the restart")

    # ---------- good commit ----------
    # Publish a good commit as the branch tip, then put the working tree back on
    # the old good_sha. That reproduces the real situation: origin has moved
    # ahead, so the next deploy has something to fast-forward onto. Force is
    # needed because origin/deploy still points at the broken commit the deploy
    # command rolled away from.
    (repo / "app.py").write_text("VALUE = 'better'\n", encoding="utf-8")
    sh("git", "add", "-A", cwd=repo)
    sh("git", "commit", "-q", "-m", "better", cwd=repo)
    sh("git", "push", "-q", "--force", "origin", "deploy", cwd=repo)
    sh("git", "reset", "--hard", good_sha, cwd=repo)
    sh("git", "clean", "-fdq", cwd=repo)

    restarts.clear()
    exc = attempt()
    good_out = captured[-1]
    if exc is not None:
        failures.append(f"good commit should have deployed, got {exc}")
    else:
        print("[ok] good commit deployed cleanly")

    if not restarts:
        failures.append(
            "service was not restarted after a successful deploy; output was:\n"
            + good_out
        )
    else:
        print(f"[ok] service restarted after a clean deploy: {restarts}")

    if "Byte-compiling" not in good_out:
        failures.append("byte-compile step missing on the good path")
    else:
        print("[ok] byte-compile ran on the good path too")

    if "SUCCESS" not in audit.read_text(encoding="utf-8"):
        failures.append("audit log missing SUCCESS")
    else:
        print("[ok] audit log recorded SUCCESS")

    print()
    if failures:
        for f in failures:
            print(f"[FAIL] {f}")
        return 1
    print("all deploy checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())