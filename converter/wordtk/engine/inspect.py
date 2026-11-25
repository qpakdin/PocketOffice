# -*- coding: utf-8 -*-
"""Read every setting group of a DOCX, or turn them into a reusable profile."""
from ..core.package import DocxPackage
from ..features import REGISTRY


def inspect_file(path, include_text=False, groups=None):
    pkg = DocxPackage(path)
    out = {"file": str(path)}
    for key, feat in REGISTRY.items():
        if groups and key not in groups:
            continue
        try:
            out[key] = feat.read(pkg, include_text)
        except Exception as e:           # one broken group must not hide the others
            out[key] = {"error": "%s: %s" % (type(e).__name__, e)}
    return out


def snapshot_profile(path, groups=None, name=None):
    """Profile that reproduces the dominant formatting of ``path`` (template extraction)."""
    pkg = DocxPackage(path)
    settings = {}
    for key, feat in REGISTRY.items():
        if groups and key not in groups:
            continue
        snap = feat.snapshot(pkg)
        if snap:
            settings[key] = snap
    return {"version": 2, "name": name or "from " + str(path).split("/")[-1],
            "scope": {"parts": ["body"], "styles": None, "include_cover": False}, "settings": settings}
