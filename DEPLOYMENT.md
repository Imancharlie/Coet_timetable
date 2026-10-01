# Deployment

Production runs on the VPS at `https://coet-timetable.kodin.co.tz` and tracks
the **`deploy` branch** automatically. Push to that branch and the server
pulls, migrates and restarts within a minute.

## How it works

A systemd timer (`coet-deploy.timer`) fires `coet-deploy.service` every 60
seconds. That runs:

```
venv/bin/python manage.py deploy --branch deploy --trigger timer
```

The command compares local `HEAD` against `origin/deploy`. If they match it
exits immediately, so an idle timer costs one `git fetch` per minute and
nothing else. When a new commit appears it:

1. Snapshots the database (`pg_dump --clean --if-exists`) to
   `deploy_snapshots/`
2. Fast-forwards the working tree (`git merge --ff-only`)
3. `pip install -r requirements.txt`
4. `manage.py check`
5. `manage.py migrate`
6. `manage.py collectstatic`
7. Restarts `gunicorn-coet-timetable.service`
8. Polls the site until it answers

## Automatic rollback

Any failure in steps 1–8 triggers a rollback:

- `git reset --hard` back to the previously deployed commit
- restore the `pg_dump` snapshot (only if migrations had started, so live
  writes are not discarded when the failure was elsewhere)
- restart the service and wait for health again

The site keeps serving the last good version. Both success and failure are
appended to `/var/log/coet-deploy.log`, which survives a database restore.

## Pushing a change

```bash
git checkout deploy
git pull origin deploy
# ... make your change, test locally ...
git add -A && git commit -m "your message"
git push origin deploy
```

Check the result within a minute:

```bash
ssh root@199.192.23.102 'tail -30 /var/log/coet-deploy.log'
```

## Deploying by hand

Useful when you want to see the output rather than read the log:

```bash
cd /var/www/coet-timetable
set -a; . ./.env; set +a
DJANGO_SETTINGS_MODULE=coet.production_settings \
  venv/bin/python manage.py deploy --branch deploy --trigger manual
```

Useful flags:

| Flag | Effect |
| --- | --- |
| `--force` | Redeploy even when already up to date |
| `--skip-deps` | Skip `pip install` |
| `--skip-restart` | Run checks/migrations without restarting (CI) |
| `--health-url URL` | Override the health endpoint |
| `--health-timeout N` | Seconds to wait for health (default 120) |

## Configuration

Secrets live in `/var/www/coet-timetable/.env` (chmod 600, owned by root) and
are injected by systemd:

| Variable | Purpose |
| --- | --- |
| `DJANGO_SECRET_KEY` | Django secret key |
| `DJANGO_DEBUG` | `False` |
| `DJANGO_ALLOWED_HOSTS` | Comma separated; `*` is rejected when DEBUG is off |
| `DJANGO_CSRF_TRUSTED_ORIGINS` | `https://coet-timetable.kodin.co.tz` |
| `COET_DB_*` | PostgreSQL connection |

The deploy command reads its own settings from:

| Variable | Default |
| --- | --- |
| `DEPLOY_BRANCH` | `deploy` |
| `DEPLOY_SERVICE_NAME` | `gunicorn-coet-timetable.service` |
| `DEPLOY_HEALTH_URL` | `https://coet-timetable.kodin.co.tz/` |
| `DEPLOY_LOG_FILE` | `/var/log/coet-deploy.log` |
| `DEPLOY_SNAPSHOT_DIR` | `<app>/deploy_snapshots` |
| `DEPLOY_SETTINGS_MODULE` | `coet.production_settings` |

## Service management

```bash
systemctl status gunicorn-coet-timetable
systemctl restart gunicorn-coet-timetable
journalctl -u gunicorn-coet-timetable -f

systemctl list-timers coet-deploy.timer
systemctl stop coet-deploy.timer      # pause auto-deploys
systemctl start coet-deploy.timer
```

## Notes

`manage.py check` runs *before* `migrate`, so a change that breaks settings or
models is rejected without touching the database.

Only the 5 most recent `pg_dump` snapshots are kept in `deploy_snapshots/`.

Never put a secret in a committed file: `coet/settings.py` stays dev-only
(SQLite, `DEBUG=True`) and `coet/production_settings.py` reads everything from
the environment.

The first deploy after cloning needs the `.env` file in place before the timer
is enabled, or `production_settings` will refuse to boot.