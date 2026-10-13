"""Loading and normalising the audit thresholds.

The YAML file is the single source of truth: the same object is used to
compute every number in the report, stored verbatim on the report row
(``thresholds_json``) and printed in Appendix C. Nothing in the analytics
hard-codes a threshold, so re-running an audit with a different file genuinely
produces a different — and reproducible — report.
"""

from __future__ import annotations

import copy
from pathlib import Path

import yaml

DEFAULTS_PATH = Path(__file__).with_name("config.default.yaml")


class Cfg(dict):
    """A dict that also answers to attribute access, recursively.

    ``cfg.venue.low_seat_util`` and ``cfg["venue"]["low_seat_util"]`` are the
    same value, and every nested mapping is itself a ``Cfg``. Missing keys
    return ``None`` rather than raising, because a report must degrade to
    "Not recorded" rather than crash on a hand-edited config.
    """

    def __getattr__(self, name):
        try:
            value = self[name]
        except KeyError:
            return None
        return Cfg(value) if isinstance(value, dict) else value

    def get_path(self, dotted, default=None):
        """``cfg.get_path("venue.low_seat_util")`` with a safe default."""
        node = self
        for part in str(dotted).split("."):
            if not isinstance(node, dict) or part not in node:
                return default
            node = node[part]
        return node


def _deep_cfg(value):
    if isinstance(value, dict):
        return Cfg({k: _deep_cfg(v) for k, v in value.items()})
    if isinstance(value, list):
        return [_deep_cfg(v) for v in value]
    return value


def default_config() -> Cfg:
    with DEFAULTS_PATH.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}
    return _validate(_deep_cfg(raw))


def load_config(path=None) -> Cfg:
    """Load a config file, falling back to the shipped defaults."""
    if path is None:
        return default_config()
    file = Path(path)
    if not file.exists():
        return default_config()
    with file.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}
    return _validate(_deep_cfg(raw))


def merged_config(base: Cfg | None = None, overrides: dict | None = None) -> Cfg:
    """A copy of ``base`` with ``overrides`` deep-merged in.

    Used by the "re-audit with different thresholds" action, so a partial
    override only has to name the values it changes.
    """
    merged = copy.deepcopy(dict(base or default_config()))
    for key, value in (overrides or {}).items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = merged_config(Cfg(merged[key]), value)
        else:
            merged[key] = value
    return _validate(_deep_cfg(merged))


def as_plain_dict(config) -> dict:
    """JSON/YAML-safe plain dict, for ``thresholds_json`` and Appendix C."""
    def plain(value):
        if isinstance(value, dict):
            return {k: plain(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [plain(v) for v in value]
        return value

    return plain(dict(config))


def _validate(config: Cfg) -> Cfg:
    """Fail loudly on a config that cannot produce a meaningful score."""
    weights = config.get("weights") or {}
    total = sum(float(v) for v in weights.values())
    if weights and abs(total - 1.0) > 1e-6:
        raise ValueError(
            f"audit weights must sum to 1.0 (got {total:.3f} from {dict(weights)})"
        )
    day = config.get("working_day") or {}
    if float(day.get("end", 1080)) <= float(day.get("start", 480)):
        raise ValueError("audit working_day.end must be after working_day.start")
    for key in ("day_soft_hours", "day_hard_hours", "run_soft_hours", "run_hard_hours"):
        if float(config.get_path(key, 1)) <= 0:
            raise ValueError(f"audit {key} must be positive")
    if float(config.get("day_hard_hours", 8)) <= float(config.get("day_soft_hours", 5)):
        raise ValueError("audit day_hard_hours must exceed day_soft_hours")
    if float(config.get("run_hard_hours", 6)) <= float(config.get("run_soft_hours", 3)):
        raise ValueError("audit run_hard_hours must exceed run_soft_hours")
    if float(config.get("gap_bad_min", 240)) <= float(config.get("gap_min", 60)):
        raise ValueError("audit gap_bad_min must exceed gap_min")
    return config
