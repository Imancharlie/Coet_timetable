"""Render an audit payload to paged HTML, and that HTML to a real PDF.

Two functions, deliberately separate:

* :func:`build_payload` turns an :class:`~audit.analytics.Analysis` plus its
  recommendations into a plain JSON-safe dict. That dict is what gets stored on
  ``AuditReport.summary_json`` and what the on-screen report page renders, so
  the page and the PDF can never disagree about what the report said.
* :func:`html_to_pdf` drives headless Chrome. WeasyPrint is installed and
  unusable on Windows (``cannot load library 'libgobject-2.0-0'``), and
  reportlab cannot lay out HTML at all, so Chrome is the only renderer here
  that gets paged media, CSS counters and selectable text right.

The payload is assembled with explicit key names rather than ``asdict`` because
the analysis holds dataclasses with sets in them (``gap_bands``, ``clash_ids``)
that are not JSON-serialisable, and because most of the report is a *choice* of
what to show, not a faithful dump.
"""

from __future__ import annotations

import html
import json
import shutil
import subprocess
import tempfile
from pathlib import Path

from django.conf import settings
from django.utils import timezone

# Chrome is looked for in a fixed list rather than trusted from PATH: the
# program is a desktop install, not a Python dependency, and its location on
# Windows is stable while the registry is not worth a dependency.
CHROME_CANDIDATES = (
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
)

#: Bumped whenever the *shape* of the stored payload changes, not when a number
#: does. A stored report is a version of what was said, so it is never rewritten
#: in place — but a payload written by an older build of this app may be
#: unrenderable by the current templates. ``audit.views.report_detail`` checks
#: this and regenerates into a new version rather than showing a broken page.
PAYLOAD_VERSION = 2

PAGE_CSS = """
@page { size: A4 portrait; margin: 18mm 14mm 20mm; }
* { box-sizing: border-box; }
body { font: 9.5pt/1.45 "Segoe UI", system-ui, sans-serif; color: #0f172a;
       margin: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
h1 { font-size: 20pt; margin: 0 0 2mm; }
h2 { font-size: 12pt; margin: 7mm 0 2mm; padding-bottom: 1.2mm;
     border-bottom: .6pt solid #cbd5e1; break-after: avoid; }
h3 { font-size: 10pt; margin: 4mm 0 1.5mm; break-after: avoid; }
p  { margin: 0 0 2mm; }
ul, ol { margin: 0 0 2mm; padding-left: 5mm; }
table { width: 100%; border-collapse: collapse; margin: 0 0 3mm;
        font-size: 8.5pt; }
th, td { text-align: left; padding: 1.4mm 2mm; border-bottom: .4pt solid #e2e8f0;
         vertical-align: top; }
th { background: #f1f5f9; font-weight: 600; }
tbody tr { break-inside: avoid; }
code { font-family: Consolas, monospace; font-size: 8pt; }
.cover { border-top: 2.5mm solid #0b4f9e; padding-top: 4mm; margin-bottom: 5mm; }
.sub { color: #475569; font-size: 10pt; }
.verdict { display: inline-block; padding: 2mm 4mm; border-radius: 2mm;
           color: #fff; font-weight: 700; font-size: 13pt; margin: 3mm 0; }
.grid { display: flex; flex-wrap: wrap; gap: 3mm; margin-bottom: 4mm; }
.kpi { flex: 1 1 30mm; border: .5pt solid #cbd5e1; border-radius: 2mm;
       padding: 2.5mm 3mm; }
.kpi .n { font-size: 15pt; font-weight: 700; display: block; }
.kpi .l { font-size: 7.5pt; color: #475569; text-transform: uppercase;
          letter-spacing: .3pt; }
.rec { border-left: 2mm solid #0b4f9e; padding: 0 0 0 3mm; margin: 0 0 4mm;
       break-inside: avoid; }
.rec.critical { border-left-color: #b91c1c; }
.rec.high     { border-left-color: #c2410c; }
.rec.medium   { border-left-color: #b45309; }
.rec.low      { border-left-color: #0b4f9e; }
.rank { font-weight: 700; }
.meta { color: #64748b; font-size: 8pt; }
.ev { color: #334155; font-size: 8pt; margin: 1mm 0 0; }
.bar { background: #e2e8f0; height: 2.6mm; border-radius: 1mm; overflow: hidden; }
.bar > span { display: block; height: 100%; background: #0b4f9e; }
td.num { text-align: right; font-variant-numeric: tabular-nums; }
.appendix { color: #334155; font-size: 8pt; }
.appendix pre { white-space: pre-wrap; word-break: break-word; margin: 0; }
.foot { position: fixed; bottom: -12mm; left: 0; right: 0;
        border-top: .4pt solid #cbd5e1; padding-top: 1.2mm;
        font-size: 7.5pt; color: #64748b;
        display: flex; justify-content: space-between; }
"""


def find_chrome() -> str | None:
    """First Chrome/Chromium on this machine, or ``None``."""
    for candidate in CHROME_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    for name in ("chrome", "google-chrome", "chromium", "chromium-browser"):
        found = shutil.which(name)
        if found:
            return found
    return None


def e(value) -> str:
    """HTML-escape, rendering anything missing as an explicit dash.

    ``Not recorded`` is load-bearing in this report: a blank cell reads as
    "zero" or "fine", and the difference matters wherever a metric is absent.
    """
    if value is None or value == "":
        return '<span class="meta">&mdash;</span>'
    return html.escape(str(value))


def _num(value, suffix="", digits=1):
    if value is None:
        return None
    return f"{float(value):,.{digits}f}{suffix}"


def _pct(value, digits=1):
    if value is None:
        return None
    return f"{float(value) * 100:.{digits}f}%"


# ---------------------------------------------------------------------------
# payload
# ---------------------------------------------------------------------------


def build_payload(analysis, recommendations, run=None) -> dict:
    """The JSON-safe body of one report version."""
    cfg = analysis.config
    kpis = dict(analysis.kpis)
    validity = [i for i in analysis.issues if i.affects_validity]
    day_pressure = _day_pressure_rows(analysis)

    return {
        "payload_version": PAYLOAD_VERSION,
        "generated_at": timezone.localtime().strftime("%d %b %Y, %H:%M"),
        "verdict": {
            "label": analysis.verdict_label,
            "colour": analysis.verdict_colour,
        },
        "run": {
            "id": run.pk if run else None,
            "semester": str(run.semester) if run else "",
            "scope": getattr(run, "scope", "") or "",
            "status": getattr(run, "status", "") or "",
            "assigned": getattr(run, "assigned", None),
            "unresolved": getattr(run, "unresolved", None),
        } if run else {},
        "kpis": kpis,
        "data_hash": _data_hash(analysis.dataset),
        "counts": {
            "groups": len(analysis.groups),
            "unscored": len(analysis.unscored_groups),
            "issues": len(analysis.issues),
            "validity": len(validity),
            "students": sum(g.size for g in analysis.groups),
        },
        "day_pressure": day_pressure,
        "day_peak": max((d["student_hours"] for d in day_pressure), default=0) or 1,
        "tfi": {
            "mean": kpis.get("campus_tfi"),
            "median": kpis.get("campus_median_tfi"),
            "worst": min(
                (g.tfi["score"] for g in analysis.groups if g.tfi), default=None
            ),
            "completion": _as_pct(kpis.get("completion_pct")),
            "dead_time": kpis.get("dead_time_pct"),
            "groups_at_de": kpis.get("groups_at_de"),
            "seat_util": kpis.get("venue_seat_util"),
        },
        "allocation": analysis.allocation,
        "performance": analysis.performance,
        "venues": _venue_rows(analysis),
        "recommendations": _rec_rows(recommendations),
        "issues": _issue_rows(analysis, cfg),
        "pareto": [
            {
                "count": row.get("count"),
                "students": row.get("students"),
                "categories": _as_text(row.get("categories")),
                "cumulative_pct": row.get("cumulative_pct"),
            }
            for row in (analysis.pareto or [])[:15]
        ],
        "observations": [dict(o) if isinstance(o, dict) else {"text": str(o)}
                         for o in (analysis.observations or [])],
        "groups": _group_rows(analysis),
        "unscored": [
            {"code": g.code, "programme": g.programme}
            for g in (analysis.unscored_groups or [])
        ],
        "config_yaml": _config_yaml(cfg),
        "config_path": getattr(cfg, "source_path", "") or "",
    }


def _data_hash(dataset) -> str:
    """``data_hash`` is a method on the dataset, not a cached attribute."""
    value = getattr(dataset, "data_hash", "")
    try:
        return value() if callable(value) else (value or "")
    except Exception:  # noqa: BLE001 - lineage is useful, not required
        return ""


def _as_pct(value):
    """The analysis stores 0-100 for percentages; the report shows 0-100 too."""
    return None if value is None else round(float(value), 1)


def _as_text(value):
    """``categories`` arrives as a set; JSON needs a string."""
    if isinstance(value, (set, frozenset, list, tuple)):
        return ", ".join(sorted(str(v) for v in value))
    return value


def _day_pressure_rows(analysis) -> list:
    rows = []
    for row in analysis.rollup.get("day_pressure") or []:
        rows.append(
            {
                "day": row.get("label"),
                "student_hours": row.get("hours"),
                "sessions": row.get("sessions"),
            }
        )
    return rows


def _venue_rows(analysis) -> list:
    rows = []
    for name, data in sorted((analysis.venues or {}).items()):
        rows.append(
            {
                "venue": name,
                "capacity": data.get("capacity"),
                "has_venue": data.get("has_venue"),
                "sessions": data.get("sessions"),
                "students": data.get("students"),
                "booked_hours": data.get("booked_hours"),
                "seat_util": data.get("seat_util_avg"),
                "seat_util_max": data.get("seat_util_max"),
                "time_util": data.get("time_util"),
                # ``overflow_sessions`` is a count in the analysis, not a list.
                # It is kept as a count here so both renderers can say "N over
                # capacity" without having to guess which they were handed.
                "over_capacity": int(data.get("overflow_sessions") or 0),
            }
        )
    return rows


#: Weekday labels for the integer day values the analysis carries. A
#: presentation concern, so it lives with the presentation.
DAY_LABELS = (
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
    "Saturday", "Sunday",
)


def _clock(minutes):
    """540 -> '09:00'. The analysis stores minutes from midnight."""
    if minutes is None:
        return ""
    minutes = int(minutes)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def _time_span(start, end):
    if start is None:
        return ""
    return f"{_clock(start)}-{_clock(end)}" if end is not None else _clock(start)


def _evidence_rows(rec) -> list:
    """Evidence, with both keys always present.

    The recommendation builders write evidence as free-form dicts and use
    different key names. A template reading ``{{ item.detail }}`` on a row that
    only has ``note`` raises VariableDoesNotExist and takes the whole report
    page down, so the two names are folded into one here — once — rather than
    defended against at every read site.
    """
    out = []
    for item in rec.evidence or []:
        if not isinstance(item, dict):
            out.append({"label": "", "detail": str(item), "groups": []})
            continue
        out.append(
            {
                "label": item.get("label") or item.get("metric") or "",
                "detail": item.get("detail") or item.get("note") or "",
                "groups": [str(g) for g in (item.get("groups") or [])][:12],
            }
        )
    return out


def _alternative_rows(rec) -> list:
    """Flatten one recommendation's alternatives into printable rows.

    An alternative for the equity recommendation is a *group* carrying a nested
    ``options`` list of candidate sessions, not a flat session record, so the
    nesting is resolved here. The alternative is the best-scoring option, and
    every option is kept in ``all`` because "here is the one I would move it to"
    and "here is what else it could move to" are different questions.
    """
    out = []
    for alt in rec.alternatives or []:
        if not isinstance(alt, dict):
            out.append({"group": str(alt), "problem": "", "options": [], "all": []})
            continue
        options = []
        for opt in alt.get("options") or []:
            options.append(
                {
                    "session": (
                        f"{opt.get('course', '')} {opt.get('activity', '')}".strip()
                    ),
                    "day": opt.get("day_label") or _day_label(opt.get("day")),
                    "time": _time_span(opt.get("start"), opt.get("end")),
                    "venue": opt.get("venue"),
                    "why": opt.get("why") or "",
                    "improves": bool(opt.get("improves")),
                    "session_pk": opt.get("session_pk"),
                }
            )
        best = options[0] if options else {}
        out.append(
            {
                "group": alt.get("group") or "",
                "programme": alt.get("programme") or "",
                "score": alt.get("score"),
                "grade": alt.get("grade") or "",
                "dead_min": alt.get("dead_min"),
                "longest_gap": alt.get("longest_gap"),
                "busiest_day": _day_label(alt.get("busiest_day")),
                "problem": _problem_text(alt, best, options),
                "session": best.get("session", ""),
                "day": best.get("day", ""),
                "time": best.get("time", ""),
                "venue": best.get("venue"),
                "why": best.get("why", ""),
                "options": options,
                "all": options,
            }
        )
    return out


def _day_label(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    index = int(value)
    return DAY_LABELS[index] if 0 <= index < len(DAY_LABELS) else str(value)


def _problem_text(alt, best, options) -> str:
    """Why this group is on the list — the number that put it there."""
    bits = []
    if alt.get("score") is not None:
        bits.append(f"TFI {alt['score']} (grade {alt.get('grade', '?')})")
    if alt.get("longest_gap"):
        bits.append(f"{alt['longest_gap']}-minute gap")
    if alt.get("busiest_day") is not None:
        bits.append(f"heaviest day {_day_label(alt['busiest_day'])}")
    if not options:
        bits.append("no existing session it could be moved to")
    return "; ".join(bits) or "flagged"


def _rec_rows(recommendations) -> list:
    out = []
    for rec in getattr(recommendations, "all", recommendations) or []:
        out.append(
            {
                "rank": rec.rank,
                # ``key`` is the stable slug ('venue-overflow'); there is no
                # ``id`` on a recommendation — a report never mints one.
                "key": rec.key,
                "category": rec.category,
                "title": rec.title,
                "priority": rec.priority,
                "presentation": rec.presentation,
                "headline": rec.headline,
                "action": rec.action,
                "impact": rec.impact,
                "groups": len(rec.groups),
                "group_codes": rec.groups,
                "students": rec.students,
                "hours": rec.hours,
                "issue_ids": rec.issue_ids,
                "evidence": _evidence_rows(rec),
                "alternatives": _alternative_rows(rec),
            }
        )
    return out


def _issue_rows(analysis, cfg) -> list:
    limit = int(cfg.get_path("report.issue_table_limit", 400) or 400)
    return [
        {
            "id": i.id,
            "category": i.category,
            "severity": i.severity,
            "affects_validity": i.affects_validity,
            "description": i.description,
            "action": i.action,
            "groups": ", ".join(i.groups[:12]) + ("…" if len(i.groups) > 12 else ""),
            "group_count": len(i.groups),
            "students": i.students,
            "detection": i.detection,
        }
        for i in (analysis.issues or [])[:limit]
    ]


def _group_rows(analysis) -> list:
    rows = []
    for g in sorted(
        analysis.groups,
        key=lambda x: (x.tfi.get("score", 999) if x.tfi else 999),
    ):
        tfi = g.tfi or {}
        longest = tfi.get("longest_gap")
        rows.append(
            {
                "code": g.code,
                "programme": g.programme,
                "size": g.size,
                "sessions": len(g.sessions),
                "score": tfi.get("score"),
                "grade": tfi.get("grade"),
                "dead_ratio": tfi.get("dead_ratio"),
                "longest_gap": None if longest is None else int(longest),
                "days_used": tfi.get("days_used"),
                "pct_campus": g.percentile_campus,
                # A flag is a structured record ({'metric', 'value', 'grade',
                # 'why', 'z', 'note'}), not a label. The ``note`` is the half a
                # human wrote for a human, so that is what is reported; the
                # numbers behind it stay in the analysis and the register.
                "flags": [_flag_text(f) for f in g.flags],
            }
        )
    return rows


def _flag_text(flag) -> str:
    if isinstance(flag, dict):
        return flag.get("note") or flag.get("metric") or str(flag)
    return str(flag)


def _config_yaml(cfg) -> str:
    """The config exactly as it was applied, for Appendix C.

    Dumped from the object rather than re-read from disk: the report has to
    state the thresholds *this* run judged against, and a file edited since
    would otherwise print values the numbers were never compared to.
    """
    import yaml

    from audit.config import as_plain_dict

    try:
        return yaml.safe_dump(as_plain_dict(cfg), sort_keys=False, default_flow_style=False)
    except Exception:  # pragma: no cover - config always loads in practice
        return "{}"


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------

VERDICT_COLOURS = {
    "green": "#15803d",
    "amber": "#b45309",
    "red": "#b91c1c",
}


def html_report(payload: dict, title: str) -> str:
    """The paged document. Text is selectable, so the PDF is searchable."""
    p = payload
    verdict = p.get("verdict") or {}
    colour = VERDICT_COLOURS.get(verdict.get("colour"), "#334155")
    foot = (
        '<div class="foot">'
        f'<span>{e(title)}</span>'
        '<span>Page <span class="pageNumber"></span> of '
        '<span class="totalPages"></span></span>'
        "</div>"
    )

    parts = [
        "<!doctype html><html lang='en'><head><meta charset='utf-8'>",
        f"<title>{e(title)}</title><style>{PAGE_CSS}</style></head><body>",
        _cover(p, verdict, colour),
        _kpis(p),
        _recommendations(p),
        _day_pressure(p),
        _venues(p),
        _groups(p),
        _issues(p),
        _observations(p),
        _appendix(p),
        foot,
        "</body></html>",
    ]
    return "".join(parts)


def _cover(p, verdict, colour) -> str:
    run = p.get("run") or {}
    counts = p.get("counts") or {}
    return (
        '<div class="cover">'
        f"<h1>Post-Allocation Audit Report</h1>"
        f'<p class="sub">{e(run.get("semester"))}'
        + (f' &middot; scope {e(run.get("scope"))}' if run.get("scope") else "")
        + f' &middot; allocation run #{e(run.get("id"))}</p>'
        f'<p class="meta">Generated {e(p.get("generated_at"))} &middot; '
        f'data hash {e((p.get("data_hash") or "")[:16])}</p>'
        f'<p><span class="verdict" style="background:{colour}">'
        f'{e(verdict.get("label")) or "Not recorded"}</span></p>'
        f'<p class="meta">{e(counts.get("groups"))} groups in scope &middot; '
        f'{e(counts.get("students"))} students &middot; '
        f'{e(counts.get("issues"))} findings '
        f'({e(counts.get("validity"))} affecting validity)</p>'
        "</div>"
    )


def _kpis(p) -> str:
    tfi = p.get("tfi") or {}
    counts = p.get("counts") or {}
    cards = [
        ("Mean TFI", _num(tfi.get("mean"))),
        ("Worst TFI", _num(tfi.get("worst"))),
        ("Requirements met", None if tfi.get("completion") is None else f'{tfi["completion"]}%'),
        ("Groups at D/E", e(tfi.get("groups_at_de"))),
        ("Findings", e(counts.get("issues"))),
        ("Validity", e(counts.get("validity"))),
    ]
    body = "".join(
        f'<div class="kpi"><span class="n">{value}</span>'
        f'<span class="l">{e(label)}</span></div>'
        for label, value in cards
    )
    return f'<h2>Headline</h2><div class="grid">{body}</div>'


def _recommendations(p) -> str:
    recs = p.get("recommendations") or []
    if not recs:
        return (
            "<h2>Recommendations</h2><p>No recommendations: the analysis "
            "produced no finding that has an action attached to it.</p>"
        )
    out = ["<h2>Recommendations</h2>"]
    for r in recs:
        if r.get("presentation") == "table":
            out.append(_rec_table(r))
        else:
            out.append(_rec_card(r))
    return "".join(out)


def _rec_card(r) -> str:
    priority = (r.get("priority") or "low").lower()
    ev = "".join(
        f'<p class="ev">&bull; {e(item.get("label"))}: {e(item.get("detail"))}</p>'
        for item in (r.get("evidence") or [])[:6]
    )
    return (
        f'<div class="rec {e(priority)}">'
        f'<p><span class="rank">{e(r.get("rank"))}.</span> {e(r.get("title"))}</p>'
        f'<p>{e(r.get("headline"))}</p>'
        f'<p><strong>Do this:</strong> {e(r.get("action"))}</p>'
        f'<p class="meta">affects {e(r.get("groups"))} group(s) &middot; '
        f'{e(r.get("students"))} student-attendances &middot; '
        f'{_num(r.get("hours"), " student-hours")}</p>'
        f"{ev}</div>"
    )


def _rec_table(r) -> str:
    """A long list of the same kind of thing, as a table.

    Reads the *normalised* alternative rows (see :func:`_alternative_rows`), so
    the column set here and the column set in the on-screen table cannot drift:
    both are handed the same flat record.
    """
    head = ""
    body = ""
    for alt in r.get("alternatives") or []:
        head = (
            "<tr><th>Group</th><th>Why it is listed</th><th>Suggested session</th>"
            "<th>When</th><th>Room</th></tr>"
        )
        body += (
            "<tr>"
            f"<td>{e(alt.get('group'))}</td>"
            f"<td>{e(alt.get('problem'))}</td>"
            f"<td>{e(alt.get('session') or 'no existing session fits')}"
            f"{e(' &mdash; ' + alt['why']) if alt.get('why') else ''}</td>"
            f"<td>{e(alt.get('day'))} {e(alt.get('time'))}</td>"
            f"<td>{e(alt.get('venue'))}</td>"
            "</tr>"
        )
    return (
        f'<div class="rec">'
        f'<p><span class="rank">{e(r.get("rank"))}.</span> {e(r.get("title"))}</p>'
        f'<p>{e(r.get("headline"))}</p>'
        f'<p><strong>Do this:</strong> {e(r.get("action"))}</p>'
        + (f'<table>{head}{body}</table>' if head else "")
        + "</div>"
    )


def _day_pressure(p) -> str:
    days = p.get("day_pressure") or []
    if not days:
        return ""
    peak = max((d.get("student_hours") or 0 for d in days), default=0) or 1
    rows = ""
    for data in days:
        value = data.get("student_hours") or 0
        width = max(2, int(round(value / peak * 100)))
        rows += (
            f"<tr><td>{e(data.get('day'))}</td>"
            f'<td class="num">{_num(value)}</td>'
            f'<td class="num">{e(data.get("sessions"))}</td>'
            f'<td><div class="bar"><span style="width:{width}%"></span></div></td>'
            "</tr>"
        )
    return (
        "<h2>Day pressure</h2><table><thead><tr><th>Day</th>"
        '<th class="num">Student-hours</th><th class="num">Sessions</th>'
        f"<th>Share</th></tr></thead><tbody>{rows}</tbody></table>"
    )


def _venues(p) -> str:
    rows_data = p.get("venues") or []
    if not rows_data:
        return ""
    rows = ""
    for v in rows_data:
        over = int(v.get("over_capacity") or 0)
        if over:
            note = (
                f'<span style="color:#b91c1c;font-weight:600">'
                f'{e(over)} session(s) over capacity</span>'
            )
        elif not v.get("has_venue"):
            note = '<span style="color:#b45309">no room recorded</span>'
        elif v.get("capacity") is None:
            note = '<span style="color:#b45309">capacity not recorded</span>'
        else:
            note = ""
        rows += (
            "<tr>"
            f"<td>{e(v.get('venue'))}</td>"
            f'<td class="num">{e(v.get("capacity"))}</td>'
            f'<td class="num">{e(v.get("sessions"))}</td>'
            f'<td class="num">{e(v.get("students"))}</td>'
            f'<td class="num">{_pct(v.get("seat_util"), 0)}</td>'
            f'<td class="num">{_pct(v.get("time_util"), 0)}</td>'
            f"<td>{note}</td>"
            "</tr>"
        )
    return (
        "<h2>Room estate</h2><table><thead><tr><th>Room</th>"
        '<th class="num">Capacity</th><th class="num">Sessions</th>'
        '<th class="num">Students</th><th class="num">Seat use</th>'
        f'<th class="num">Time use</th><th>Note</th></tr></thead>'
        f"<tbody>{rows}</tbody></table>"
    )


def _groups(p) -> str:
    rows_data = p.get("groups") or []
    if not rows_data:
        return "<h2>Groups</h2><p>Not recorded</p>"
    rows = ""
    for g in rows_data:
        grade = (g.get("grade") or "").upper()
        tone = {"A": "#15803d", "B": "#0b4f9e", "C": "#b45309", "D": "#c2410c",
                "E": "#b91c1c"}.get(grade, "#334155")
        gap = g.get("longest_gap")
        rows += (
            "<tr>"
            f"<td><strong>{e(g.get('code'))}</strong></td>"
            f"<td>{e(g.get('programme'))}</td>"
            f'<td class="num">{e(g.get("size"))}</td>'
            f'<td class="num">{e(g.get("sessions"))}</td>'
            f'<td class="num" style="color:{tone};font-weight:700">'
            f'{e(g.get("score"))} {e(grade)}</td>'
            f'<td class="num">{_num(g.get("dead_ratio"), "", 2)}</td>'
            f'<td class="num">{"&mdash;" if gap is None else f"{gap}m"}</td>'
            f'<td class="num">{e(g.get("pct_campus"))}</td>'
            f"<td>{e(g.get('flags'))}</td>"
            "</tr>"
        )
    unscored = p.get("unscored") or []
    tail = (
        "<p class='meta'>Not scored (no requirement list, or nothing "
        f"timetabled): {e(', '.join(u['code'] for u in unscored))}</p>"
        if unscored
        else ""
    )
    return (
        "<h2>Groups</h2><table><thead><tr><th>Group</th><th>Programme</th>"
        '<th class="num">Size</th><th class="num">Sessions</th>'
        '<th class="num">TFI</th><th class="num">Dead ratio</th>'
        f'<th class="num">Longest gap</th><th class="num">Percentile</th>'
        f"<th>Flags</th></tr></thead>"
        f"<tbody>{rows}</tbody></table>{tail}"
    )


def _issues(p) -> str:
    rows_data = p.get("issues") or []
    if not rows_data:
        return "<h2>Finding register</h2><p>No findings were recorded.</p>"
    rows = ""
    for i in rows_data:
        tone = "#b91c1c" if i.get("affects_validity") else "#475569"
        rows += (
            "<tr>"
            f'<td style="color:{tone};font-weight:600">{e(i.get("severity"))}</td>'
            f"<td>{e(i.get('category'))}</td>"
            f"<td>{e(i.get('description'))}</td>"
            f"<td>{e(i.get('group_count'))}</td>"
            f'<td class="num">{e(i.get("students"))}</td>'
            "</tr>"
        )
    return (
        "<h2>Finding register</h2><table><thead><tr><th>Severity</th>"
        "<th>Category</th><th>Finding</th><th class='num'>Groups</th>"
        f"<th class='num'>Students</th></tr></thead><tbody>{rows}</tbody></table>"
    )


def _observations(p) -> str:
    obs = p.get("observations") or []
    if not obs:
        return ""
    items = "".join(
        f"<li>{e(_observation_text(o))}</li>" for o in obs
    )
    return f"<h2>Observations</h2><ul>{items}</ul>"


def _observation_text(o) -> str:
    """Observations are dicts of the offending record, not sentences."""
    if "text" in o:
        return o["text"]
    bits = [
        f"{k} {o[k]}"
        for k in ("venue", "course", "activity")
        if o.get(k)
    ]
    if not bits:
        return ", ".join(f"{k}={v}" for k, v in list(o.items())[:4])
    if o.get("ratio"):
        return (
            f"{', '.join(bits)}: {o.get('students')} students into "
            f"{o.get('capacity')} seats ({o.get('ratio')}x over)"
        )
    return " — ".join(bits)


def _appendix(p) -> str:
    yaml_text = p.get("config_yaml") or ""
    return (
        "<h2>Appendix C &mdash; configuration</h2>"
        f'<div class="appendix"><pre>{e(yaml_text)}</pre></div>'
    )


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------


class PdfRenderError(RuntimeError):
    """Chrome could not be found or did not produce a file."""


def pdf_dir() -> Path:
    base = Path(getattr(settings, "MEDIA_ROOT", "") or (Path.cwd() / "media"))
    target = base / "audit-reports"
    target.mkdir(parents=True, exist_ok=True)
    return target


def html_to_pdf(html_text: str, out_path: Path, timeout: int = 120) -> Path:
    """Print *html_text* to a PDF with headless Chrome.

    Raises :class:`PdfRenderError` rather than returning a missing file, so the
    report row records a real reason instead of silently 404-ing later.
    """
    chrome = find_chrome()
    if chrome is None:
        raise PdfRenderError(
            "No Chrome or Chromium executable found. Install Google Chrome, or "
            "put chrome on PATH, to generate PDF reports."
        )
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="audit-render-") as tmp:
        source = Path(tmp) / "report.html"
        source.write_text(html_text, encoding="utf-8")
        # ``file:///`` with a native path. Windows paths need the drive letter
        # kept and the separators as slashes.
        uri = source.resolve().as_uri()
        base = [
            chrome,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--no-pdf-header-footer",
            # A throwaway profile, and this is not a micro-optimisation: with
            # the default profile Chrome *attaches to the already-running
            # browser* instead of starting a clean headless process, and the
            # print job then sits in that browser's queue. It was observed to
            # take 96s instead of 14s on a machine with Chrome open. One
            # directory per print keeps this a private, disposable instance.
            f"--user-data-dir={Path(tmp) / 'profile'}",
            # 2s, not 10s. The report is plain HTML and CSS with no scripts and
            # no remote assets, so there is nothing to wait for; a longer
            # virtual-time budget only makes Chrome idle. Both budgets produced
            # a byte-identical 24-page document.
            "--virtual-time-budget=2000",
            f"--print-to-pdf={out_path}",
            uri,
        ]
        # Older Chrome (and Chrome-for-Testing builds) do not know
        # ``--headless=new``; fall back rather than failing the whole report.
        attempts = [base, [chrome, "--headless", *base[2:]]]
        last = ""
        for command in attempts:
            try:
                done = subprocess.run(
                    command, capture_output=True, text=True, timeout=timeout
                )
            except subprocess.TimeoutExpired as exc:
                last = f"Chrome timed out after {timeout}s"
                continue
            if out_path.exists() and out_path.stat().st_size > 0:
                return out_path
            last = (done.stderr or done.stdout or "").strip()[:500]
        raise PdfRenderError(
            f"Chrome produced no PDF. Last output: {last or 'unknown'}"
        )
