"""Venue name quality helpers.

Excel imports frequently arrive with inconsistent casing, irregular spacing
or duplicate records that name the same physical venue differently
(e.g. "a104" vs "A104", "A 104" vs "A104", "DO1 luhanga hall kijitonyama"
vs "DO1 kijitonyama").  The functions here normalise a venue name into a
comparable key and classify the kinds on inconsistency found so the UI can
highlight problem rows and let a user recycle (edit/fix/delete) them, keeping
the database in sync.

It also provides the pairwise detection used by the interactive venue-import
conflict fixer (spacing / casing duplicates) plus the merge primitives used to
fold an accepted venue name back into the database.
"""
import re
import uuid

_BASE_CODE_RE = re.compile(r"^([A-Z]{1,6}\d{1,6})")
_BREAKS_CODE_RE = re.compile(r"^[A-Za-z]+\s+\d+")
_CODELIKE_RE = re.compile(r"^[A-Z]{1,6}\d{1,6}$")


def collapse_ws(value: str) -> str:
    """Trim and collapse any run of whitespace into a single space."""
    return " ".join(str(value).strip().split())


def base_key(name: str) -> str:
    """Return a canonical key for a venue name.

    Casing is ignored, every space is removed and any trailing locator text
    is dropped so "a104", "A104", "A 104" and "A104 - Stage Left" all key to
    "A104"; "DO1 luhanga hall kijitonyama" and "DO1 kijitonyama" both key to
    "DO1".  Compound codes separated by - / or , are kept whole
    ("B4-206" stays distinct from "B4", "A104, A106" stays distinct from
    "A104").
    """
    compact = re.sub(r"\s+", "", str(name).strip().upper())
    if not compact:
        return compact
    m = _BASE_CODE_RE.match(compact)
    if not m:
        return compact
    head = m.group(1)
    rest = compact[len(head):]
    if rest and rest[0] in "-/,":
        return compact
    return head


def issues_for(name: str) -> list:
    """Return the quality issue kinds for a single venue name."""
    text = str(name)
    stripped = text.strip()
    issues = []
    if stripped != stripped.upper():
        issues.append("Casing")
    if text != collapse_ws(text):
        issues.append("Spacing")
    # A space that breaks an alphanumeric code, e.g. "A 104", "P 008".
    if _BREAKS_CODE_RE.match(stripped) and not re.match(r"^[A-Za-z]+\d+", stripped):
        issues.append("Spacing")
    return list(dict.fromkeys(issues))


def suggested_name(name: str, key: str | None = None) -> str:
    """Return the cleaned name that a human should probably adopt.

    Code-like venues collapse to their compact code ("a104" -> "A104",
    "DO1 luhanga hall kijitonyama" -> "DO1"); everything else just gets a
    canonical case + single-space treatment.
    """
    key = key if key is not None else base_key(name)
    if _CODELIKE_RE.match(key):
        return key
    return collapse_ws(name).upper()


def analyse_venues():
    """Analyse all venues and return (issues_map, group_list, has_issues).

    issues_map: {venue pk: {"key", "issues", "summary", "suggested",
    "duplicate_of": [names]}} for every venue that has a problem.
    group_list: duplicate groups [{key, venues}, ...] sorted by key.
    """
    from .models import Venue

    venues = list(Venue.objects.all())
    groups = {}
    for v in venues:
        groups.setdefault(base_key(v.name), []).append(v)

    duplicate_groups = [
        {"key": k, "venues": lst}
        for k, lst in groups.items()
        if len(lst) > 1
    ]
    duplicate_groups.sort(key=lambda g: g["key"])
    dup_pks = {v.pk for g in duplicate_groups for v in g["venues"]}

    issues_map = {}
    for v in venues:
        key = base_key(v.name)
        v_issues = issues_for(v.name)
        if v.pk in dup_pks:
            v_issues.append("Duplicate")
        if not v_issues:
            continue
        duplicates = [u.name for u in groups[key] if u.pk != v.pk]
        issues_map[v.pk] = {
            "key": key,
            "issues": v_issues,
            "summary": ", ".join(v_issues),
            "suggested": suggested_name(v.name, key),
            "duplicate_of": duplicates,
        }
    return issues_map, duplicate_groups, bool(issues_map)


def _format_key(name: str) -> str:
    """Strict formatting key: lowercase, no whitespace at all.

    'PB 06', 'pb06' and 'PB06' all key to 'pb06', while 'DO1 luhanga hall
    kijitonyama' and 'DO1 kijitonyama' stay distinct (they are genuine name
    duplicates, not pure formatting variants).
    """
    return "".join(str(name).lower().split())


def conflict_reasons(a: str, b: str) -> list:
    """Classify why two venue names are formatting duplicates.

    Returns a (possibly empty) list drawn from the labels "Spacing" and
    "Casing" describing the difference between the two names:
    - "Casing": identical once letter case is ignored ("PB 06" vs "pb 06").
    - "Spacing": identical once whitespace is removed ("PB 06" vs "PB06").
    - both: the names differ in both ways ("pb 06" vs "PB06").
    An empty list means the pair is not a pure formatting duplicate.
    """
    a = str(a).strip()
    b = str(b).strip()
    if a == b or _format_key(a) != _format_key(b):
        return []
    reasons = []
    if a.lower() == b.lower():
        reasons.append("Casing")
    if " ".join(a.split()) != " ".join(b.split()) and "".join(a.split()) == "".join(
        b.split()
    ):
        reasons.append("Spacing")
    if not reasons:
        reasons = ["Spacing", "Casing"]
    return reasons


def detect_name_conflicts(names) -> list:
    """Return pairwise spacing/casing conflicts among a set of venue names.

    Every group of names that is identical after removing whitespace and
    ignoring case yields one entry per distinct pair. Each entry has:
    {"id", "key", "names": [a, b], "reasons": ["Spacing"|"Casing", ...]}.
    The id is stable for the lifetime of the process so interactive fixes can
    reference a specific issue.
    """
    groups = {}
    for name in names:
        text = str(name).strip()
        if text:
            groups.setdefault(_format_key(text), set()).add(text)

    conflicts = []
    for key, name_set in groups.items():
        ordered = sorted(name_set, key=lambda n: (n.lower(), n))
        for i in range(len(ordered)):
            for j in range(i + 1, len(ordered)):
                a, b = ordered[i], ordered[j]
                reasons = conflict_reasons(a, b)
                if not reasons:
                    continue
                conflicts.append(
                    {
                        "id": uuid.uuid4().hex[:10],
                        "key": base_key(a),
                        "names": [a, b],
                        "reasons": reasons,
                    }
                )
    return conflicts


def merge_venues(source, target):
    """Fold one venue into another and keep every reference in sync.

    Sessions point at the kept venue, workshop/TD allocation strings are
    overwritten with the kept name, capacity is absorbed and the source row is
    deleted. Returns the kept venue.
    """
    from .models import (
        Session,
        TechnicalDrawingAllocation,
        Venue,
        WorkshopAllocation,
    )

    target.refresh_from_db()
    Session.objects.filter(venue=source).update(venue=target)
    WorkshopAllocation.objects.filter(venue=source.name).update(venue=target.name)
    TechnicalDrawingAllocation.objects.filter(venue=source.name).update(
        venue=target.name
    )
    if (target.capacity or 0) <= 0 and (source.capacity or 0) > 0:
        target.capacity = source.capacity
        target.save(update_fields=["capacity"])
    source.delete()
    return target


def resolve_venue_name_conflict(key: str, official: str):
    """Merge every venue normalising to ``key`` into the venue ``official``.

    The user-chosen official name becomes the surviving venue (created with
    capacity 0 if it does not exist yet) and every other venue whose
    ``base_key`` matches is folded into it. Returns ``(target, merged_names)``.
    """
    from .models import Venue

    target = Venue.objects.filter(name=official).first()
    if target is None:
        target = Venue.objects.create(name=official, capacity=0)
    merged = []
    for venue in Venue.objects.all():
        if venue.pk == target.pk:
            continue
        if base_key(venue.name) == key:
            merged.append(venue.name)
            merge_venues(venue, target)
    return target, merged