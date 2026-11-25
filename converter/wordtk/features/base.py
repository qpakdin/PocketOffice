# -*- coding: utf-8 -*-
"""Feature base class: field schema, validation/coercion and apply/read contract."""
from collections import Counter

from ..core.scope import Scope
from ..core.styles import StyleResolver
from ..core.units import normalize_color, parse_length, parse_number


class ProfileError(ValueError):
    """Invalid profile content."""


_TRUE = {"true", "yes", "on", "1", "y"}
_FALSE = {"false", "no", "off", "0", "n"}


class Field:
    """One setting. ``kind``: choice|text|int|float|bool|length|color|json|list."""

    def __init__(self, name, kind, label, choices=None, unit="pt", hint="", default=None,
                 minimum=None, maximum=None, group=""):
        self.name, self.kind, self.label = name, kind, label
        self.choices = list(choices) if choices else None
        self.unit, self.hint, self.default = unit, hint, default
        self.minimum, self.maximum, self.group = minimum, maximum, group

    def to_dict(self):
        d = {"name": self.name, "kind": self.kind, "label": self.label}
        if self.choices:
            d["choices"] = self.choices
        if self.kind == "length":
            d["unit"] = self.unit
        if self.hint:
            d["hint"] = self.hint
        if self.default is not None:
            d["default"] = self.default
        return d

    def coerce(self, value):
        """Return canonical value, or None when the field means 'leave unchanged'."""
        if value is None or (isinstance(value, str) and value.strip() == ""):
            return None
        k = self.kind
        try:
            if k == "bool":
                if isinstance(value, bool):
                    return value
                s = str(value).strip().lower()
                if s in _TRUE:
                    return True
                if s in _FALSE:
                    return False
                raise ValueError("expected true/false")
            if k == "choice":
                s = str(value).strip()
                for c in self.choices:
                    if s.lower() == str(c).lower():
                        return c
                raise ValueError("choose one of: %s" % ", ".join(map(str, self.choices)))
            if k == "text":
                return str(value)
            if k in ("int", "float"):
                v = parse_number(value)
                if self.minimum is not None and v < self.minimum or self.maximum is not None and v > self.maximum:
                    raise ValueError("must be between %s and %s" % (self.minimum, self.maximum))
                return int(round(v)) if k == "int" else v
            if k == "length":
                return parse_length(value, self.unit)
            if k == "color":
                return normalize_color(value)
            if k == "list":
                if isinstance(value, str):
                    value = [x.strip() for x in value.split(",") if x.strip()]
                if not isinstance(value, (list, tuple)):
                    raise ValueError("expected a list")
                return list(value)
            if k == "json":
                if isinstance(value, str):
                    import json
                    try:
                        value = json.loads(value)
                    except ValueError:
                        if value.strip()[:1] in "[{":
                            raise ValueError("malformed JSON")
                        # plain text such as 'exact 18pt' is passed on as-is
                return value
        except ValueError as e:
            raise ProfileError("%s: %s" % (self.name, e))
        raise ProfileError("unknown field kind %r" % k)


class Context:
    """Per-file apply state shared between features."""

    def __init__(self, pkg, scope=None):
        self.pkg = pkg
        self.scope = scope or Scope()
        self.report = Counter()
        self.warnings = []
        self._styles = None

    @property
    def styles(self):
        if self._styles is None:
            self._styles = StyleResolver(self.pkg)
        return self._styles

    def reset_styles(self):
        self._styles = None

    def count(self, key, n=1):
        if n:
            self.report[key] += n

    def warn(self, msg):
        if msg not in self.warnings:
            self.warnings.append(msg)


class Feature:
    key = ""
    label = ""
    description = ""
    fields = ()
    uses_scope = True       # accepts a per-feature "scope" override

    def field_map(self):
        return {f.name: f for f in self.fields}

    def schema(self):
        return {"key": self.key, "label": self.label, "description": self.description,
                "uses_scope": self.uses_scope, "fields": [f.to_dict() for f in self.fields]}

    def normalize(self, settings):
        """Validate + coerce a settings dict. Raises ProfileError."""
        if not isinstance(settings, dict):
            raise ProfileError("%s: settings must be an object" % self.key)
        fm, out = self.field_map(), {}
        for k, v in settings.items():
            if k == "scope" and not self.uses_scope:
                continue            # accepted for old profiles, has no effect here
            if k == "scope":
                try:
                    Scope.from_dict(v)
                except ValueError as e:
                    raise ProfileError("%s.scope: %s" % (self.key, e))
                out["scope"] = v
                continue
            if k.startswith("_") or k == "comment":
                continue
            if k not in fm:
                raise ProfileError("%s: unknown setting %r (known: %s)" % (self.key, k, ", ".join(fm)))
            try:
                c = fm[k].coerce(v)
            except ProfileError as e:
                raise ProfileError("%s.%s" % (self.key, e))
            if c is not None:
                out[k] = c
        return self.validate(out)

    def validate(self, s):
        return s

    def scope_for(self, ctx, settings):
        if "scope" in settings:
            return Scope.from_dict(settings["scope"], ctx.scope)
        return ctx.scope

    # contract -------------------------------------------------------------
    def read(self, pkg, include_text=False):
        return {}

    def snapshot(self, pkg):
        """Settings that would reproduce this document's dominant formatting."""
        return {}

    def apply(self, ctx, settings):
        raise NotImplementedError


def F(name, kind, label, **kw):
    return Field(name, kind, label, **kw)


def most_common(counter, minimum_share=0.0):
    """Return the most common non-None key of a Counter (or None)."""
    items = [(k, v) for k, v in counter.most_common() if k is not None]
    if not items:
        return None
    total = sum(v for _k, v in counter.items())
    return items[0][0] if items[0][1] / max(total, 1) >= minimum_share else None
