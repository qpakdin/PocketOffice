# -*- coding: utf-8 -*-
"""JSON profile: load, validate, apply to one file.

Profile format (version 2)::

    {
      "name": "Thesis style",
      "scope": {"parts": ["body"], "styles": null, "include_cover": false},
      "settings": {
         "paragraph": {"alignment": "justify", "line_spacing": 1.5, "direction": "rtl"},
         "font": {"font": "B Nazanin", "size": 14},
         "page": {"paper": "A4", "margin_top": "2.5cm"}
      }
    }
"""
import json
from pathlib import Path

from ..core.package import DocxPackage
from ..core.scope import Scope
from ..features import REGISTRY, Context, ProfileError, get_feature


class Profile:
    def __init__(self, name, scope, settings, source=None, legacy_ops=None):
        self.name, self.scope, self.settings, self.source = name, scope, settings, source
        self.legacy_ops = legacy_ops or []

    def to_dict(self):
        return {"version": 2, "name": self.name, "scope": self.scope.to_dict(), "settings": self.settings}


def parse_profile(data, source=None):
    if not isinstance(data, dict):
        raise ProfileError("profile must be a JSON object")
    if "operations" in data and "settings" not in data:      # v1 prototype profile
        if not isinstance(data["operations"], list):
            raise ProfileError("'operations' must be a list")
        return Profile(data.get("name", "legacy profile"), Scope(), {}, source, data["operations"])
    raw = data.get("settings")
    if not isinstance(raw, dict) or not raw:
        raise ProfileError("profile needs a non-empty 'settings' object")
    try:
        scope = Scope.from_dict(data.get("scope"))
    except ValueError as e:
        raise ProfileError("scope: %s" % e)
    settings = {}
    for key, val in raw.items():
        if key.startswith("_"):
            continue
        settings[key] = get_feature(key).normalize(val)
    return Profile(data.get("name", "profile"), scope, settings, source)


def load_profile(path):
    path = Path(path)
    try:
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise ProfileError("%s: invalid JSON (line %d, col %d): %s" % (path.name, e.lineno, e.colno, e.msg))
    except OSError as e:
        raise ProfileError("cannot read profile: %s" % e)
    return parse_profile(data, str(path))


def apply_profile_to_file(src, dst, profile, dry_run=False):
    """Apply ``profile`` to ``src`` and write ``dst``. Returns a result dict."""
    pkg = DocxPackage(src)
    ctx = Context(pkg, profile.scope)
    for key in REGISTRY:                    # deterministic feature order
        if key in profile.settings:
            REGISTRY[key].apply(ctx, profile.settings[key])
            ctx.reset_styles()
    if profile.legacy_ops:
        from .legacy import apply_legacy
        apply_legacy(pkg, profile.legacy_ops, ctx.report)
    if not dry_run:
        pkg.save(dst)
    return {"changes": dict(ctx.report), "warnings": ctx.warnings}
