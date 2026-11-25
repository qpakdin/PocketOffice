# -*- coding: utf-8 -*-
"""Paragraph group: alignment, direction, spacing, indentation, borders, shading, tabs."""
import re
from collections import Counter

from ..core.ooxml import (ensure, ensure_ppr, ensure_rpr, get, insert_ordered, is_on, make, qn,
                          remove, set_attrs, set_onoff)
from ..core.scope import iter_paragraphs
from ..core.units import fmt_length, from_twips, twips
from .base import F, Feature, ProfileError, most_common

ALIGN_CHOICES = ["left", "center", "right", "justify", "distribute", "start", "end",
                 "justify_low", "justify_medium", "justify_high"]
_JC = {"center": "center", "justify": "both", "distribute": "distribute",
       "justify_low": "lowKashida", "justify_medium": "mediumKashida", "justify_high": "highKashida"}
BORDER_STYLES = ["single", "double", "dotted", "dashed", "dotDash", "thick", "wave", "none"]
SIDES = ("top", "left", "bottom", "right")


def jc_value(choice, rtl):
    """Map a user alignment to w:jc. left/right are *visual*; start/end are logical.

    In Word's transitional XML, w:jc="left" means 'start edge', so inside a
    right-to-left paragraph it is rendered on the right. We compensate so that
    "left" always looks left.
    """
    if choice == "left":
        return "right" if rtl else "left"
    if choice == "right":
        return "left" if rtl else "right"
    if choice == "start":
        return "left"
    if choice == "end":
        return "right"
    return _JC[choice]


def visual_alignment(jc, rtl):
    if jc in ("left", None):
        return "right" if rtl else "left"
    if jc == "right":
        return "left" if rtl else "right"
    return {"both": "justify", "lowKashida": "justify_low", "mediumKashida": "justify_medium",
            "highKashida": "justify_high"}.get(jc, jc)


def parse_line_spacing(v):
    """-> (rule, value). multiple: factor; exact/atLeast: points."""
    if isinstance(v, dict):
        rule, val = v.get("rule", "multiple"), v.get("value")
    else:
        s = str(v).strip().lower()
        named = {"single": ("multiple", 1.0), "double": ("multiple", 2.0), "1.5 lines": ("multiple", 1.5)}
        if s in named:
            return named[s]
        m = re.match(r"^(exact|exactly|atleast|at_least|at least|multiple)[ :=]+(.+)$", s)
        if m:
            rule, val = m.group(1), m.group(2)
        else:
            rule, val = "multiple", s
    rule = {"exactly": "exact", "atleast": "atLeast", "at_least": "atLeast", "at least": "atLeast"}.get(rule, rule)
    if rule not in ("multiple", "exact", "atLeast"):
        raise ProfileError("line_spacing: rule must be multiple, exact or atLeast")
    try:
        if rule == "multiple":
            from ..core.units import parse_number
            num = parse_number(val)
        else:
            from ..core.units import parse_length
            num = parse_length(val, "pt")
    except ValueError as e:
        raise ProfileError("line_spacing: %s" % e)
    if num <= 0:
        raise ProfileError("line_spacing must be > 0")
    return rule, num


class ParagraphFeature(Feature):
    key = "paragraph"
    label = "Paragraph"
    description = "Alignment, direction, spacing, indentation, borders, shading, tabs, keep options."
    fields = [
        F("alignment", "choice", "Alignment", choices=ALIGN_CHOICES,
          hint="left/right are visual (auto-flipped in RTL); start/end are logical"),
        F("direction", "choice", "Text direction", choices=["ltr", "rtl"]),
        F("line_spacing", "json", "Line spacing",
          hint="1.5  |  'exact 18pt'  |  'atLeast 14pt'  |  single / double"),
        F("space_before", "length", "Space before", unit="pt"),
        F("space_after", "length", "Space after", unit="pt"),
        F("contextual_spacing", "bool", "Don't add space between same-style paragraphs"),
        F("indent_start", "length", "Indent start (left in LTR, right in RTL)", unit="pt"),
        F("indent_end", "length", "Indent end (right in LTR, left in RTL)", unit="pt"),
        F("first_line_indent", "length", "First-line indent", unit="pt"),
        F("hanging_indent", "length", "Hanging indent", unit="pt"),
        F("keep_with_next", "bool", "Keep with next"),
        F("keep_lines", "bool", "Keep lines together"),
        F("page_break_before", "bool", "Page break before"),
        F("widow_control", "bool", "Widow/orphan control"),
        F("outline_level", "int", "Outline level (1-9, 0 = body text)", minimum=0, maximum=9),
        F("style", "text", "Paragraph style (name or id)"),
        F("shading", "color", "Shading fill colour", hint="hex, or 'auto' to clear"),
        F("border", "json", "Border",
          hint='{"style":"single","width":"0.5pt","color":"000000","space":"1pt","sides":"box|top,bottom"}'),
        F("remove_border", "bool", "Remove paragraph borders"),
        F("tabs", "json", "Tab stops", hint='[{"pos":"8cm","align":"left|center|right|decimal","leader":"none|dot|hyphen|underscore"}]'),
        F("clear_tabs", "bool", "Clear existing tab stops"),
        F("clear_direct", "bool", "Reset direct paragraph formatting first"),
        F("changes", "json", "Change (replace one value with another)",
          hint='[{"property":"alignment","from":"left","to":"justify"},{"property":"line_spacing","from":1,"to":1.5},{"property":"style","from":"Heading 3","to":"Heading 2"}]'),
        F("apply_to", "choice", "Apply to", choices=["paragraphs", "styles", "both"],
          hint="paragraphs = direct formatting; styles = style definitions"),
        F("style_targets", "list", "Style targets when apply_to uses styles",
          hint="names/ids, 'docDefaults' or '*' (default: Normal)"),
    ]

    CHANGE_PROPS = ("alignment", "direction", "line_spacing", "space_before", "space_after",
                    "indent_start", "style")

    @classmethod
    def norm_change(cls, c):
        from ..core.units import parse_length
        if not isinstance(c, dict) or c.get("property") not in cls.CHANGE_PROPS or "to" not in c:
            raise ProfileError("paragraph.changes: each item needs property (%s), from and to" % ", ".join(cls.CHANGE_PROPS))
        prop = c["property"]

        def one(v):
            if isinstance(v, str) and v.strip() in ("", "*"):
                return "*"
            try:
                if prop == "alignment":
                    if v not in ALIGN_CHOICES:
                        raise ValueError("choose one of %s" % ", ".join(ALIGN_CHOICES))
                    return v
                if prop == "direction":
                    if v not in ("ltr", "rtl"):
                        raise ValueError("ltr or rtl")
                    return v
                if prop == "line_spacing":
                    return parse_line_spacing(v)
                if prop == "style":
                    return str(v).strip()
                return parse_length(v, "pt")
            except (ValueError, ProfileError) as e:
                raise ProfileError("paragraph.changes[%s]: %s" % (prop, e))
        frm = c.get("from", "*")
        frm = frm if isinstance(frm, list) else [frm]
        to = one(c["to"])
        if to == "*":
            raise ProfileError("paragraph.changes[%s]: 'to' is required" % prop)
        return {"property": prop, "from": [one(x) for x in frm], "to": to}

    def validate(self, s):
        if "changes" in s:
            if not isinstance(s["changes"], list):
                raise ProfileError("paragraph.changes must be a list")
            s["changes"] = [self.norm_change(c) for c in s["changes"]]
        if "line_spacing" in s:
            s["line_spacing"] = parse_line_spacing(s["line_spacing"])
        if "first_line_indent" in s and "hanging_indent" in s:
            raise ProfileError("paragraph: use either first_line_indent or hanging_indent, not both")
        if "border" in s:
            self._border(s["border"])
        for t in s.get("tabs", []) or []:
            if "pos" not in t:
                raise ProfileError("paragraph.tabs: each tab needs 'pos'")
        return s

    # --------------------------------------------------------------- helpers
    @staticmethod
    def _border(b):
        from ..core.units import normalize_color, parse_length
        if not isinstance(b, dict):
            raise ProfileError("paragraph.border must be an object")
        style = b.get("style", "single")
        if style not in BORDER_STYLES:
            raise ProfileError("border.style must be one of %s" % ", ".join(BORDER_STYLES))
        sides = b.get("sides", "box")
        sides = list(SIDES) if sides == "box" else [x.strip() for x in str(sides).split(",") if x.strip()] if isinstance(sides, str) else list(sides)
        if any(x not in SIDES + ("between",) for x in sides):
            raise ProfileError("border.sides: use top,bottom,left,right,between or box")
        try:
            return {"style": style, "sides": sides,
                    "sz": max(2, min(96, int(round(parse_length(b.get("width", "0.5pt")) * 8)))),
                    "space": max(0, min(31, int(round(parse_length(b.get("space", "1pt")))))),
                    "color": normalize_color(b.get("color", "auto"))}
        except ValueError as e:
            raise ProfileError("border: %s" % e)

    def apply_ppr(self, ppr, s, rtl):
        """Write settings into a pPr element. ``rtl`` = effective direction after this call."""
        if s.get("clear_direct"):
            for c in list(ppr):
                if c.tag.split("}")[1] not in ("pStyle", "numPr", "sectPr", "rPr", "pPrChange"):
                    ppr.remove(c)
        if "direction" in s:
            bidi = ensure(ppr, "w:bidi")
            if s["direction"] == "rtl":
                bidi.attrib.pop(qn("w:val"), None)
            else:
                bidi.set(qn("w:val"), "0")
        for k, tag in (("keep_with_next", "w:keepNext"), ("keep_lines", "w:keepLines"),
                       ("page_break_before", "w:pageBreakBefore"), ("widow_control", "w:widowControl"),
                       ("contextual_spacing", "w:contextualSpacing")):
            if k in s:
                set_onoff(ppr, tag, s[k])
        if "outline_level" in s:
            lv = s["outline_level"]
            ensure(ppr, "w:outlineLvl", {"w:val": 9 if lv == 0 else lv - 1})
        if "shading" in s:
            if s["shading"] == "auto":
                remove(ppr, "w:shd")
            else:
                ensure(ppr, "w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": s["shading"]})
        if s.get("remove_border"):
            remove(ppr, "w:pBdr")
        if "border" in s:
            b = self._border(s["border"])
            pb = ensure(ppr, "w:pBdr")
            for side in b["sides"]:
                e = ensure(pb, "w:" + side)
                set_attrs(e, {"w:val": b["style"], "w:sz": b["sz"], "w:space": b["space"], "w:color": b["color"]})
        if s.get("clear_tabs") or "tabs" in s:
            tabs = ppr.find(qn("w:tabs"))
            if tabs is not None and (s.get("clear_tabs") or "tabs" in s):
                ppr.remove(tabs)
            if s.get("tabs"):
                from ..core.units import parse_length
                tabs = ensure(ppr, "w:tabs")
                rows = []
                for t in s["tabs"]:
                    try:
                        pos = twips(parse_length(t["pos"], "cm"))
                    except ValueError as e:
                        raise ProfileError("paragraph.tabs: %s" % e)
                    rows.append((pos, t.get("align", "left"), t.get("leader", "none")))
                for pos, al, ld in sorted(rows):
                    tabs.append(make("w:tab", {"w:val": al, "w:leader": ld if ld != "none" else None, "w:pos": pos}))
        if "space_before" in s or "space_after" in s or "line_spacing" in s:
            sp = ensure(ppr, "w:spacing")
            if "space_before" in s:
                set_attrs(sp, {"w:before": twips(s["space_before"]), "w:beforeLines": None, "w:beforeAutospacing": None})
            if "space_after" in s:
                set_attrs(sp, {"w:after": twips(s["space_after"]), "w:afterLines": None, "w:afterAutospacing": None})
            if "line_spacing" in s:
                rule, v = s["line_spacing"]
                set_attrs(sp, {"w:line": int(round(v * 240)) if rule == "multiple" else twips(v),
                               "w:lineRule": "auto" if rule == "multiple" else rule})
        if any(k in s for k in ("indent_start", "indent_end", "first_line_indent", "hanging_indent")):
            ind = ensure(ppr, "w:ind")
            if "indent_start" in s:
                set_attrs(ind, {"w:left": twips(s["indent_start"]), "w:start": None, "w:leftChars": None, "w:startChars": None})
            if "indent_end" in s:
                set_attrs(ind, {"w:right": twips(s["indent_end"]), "w:end": None, "w:rightChars": None, "w:endChars": None})
            if "first_line_indent" in s:
                set_attrs(ind, {"w:firstLine": twips(s["first_line_indent"]), "w:hanging": None, "w:firstLineChars": None, "w:hangingChars": None})
            if "hanging_indent" in s:
                set_attrs(ind, {"w:hanging": twips(s["hanging_indent"]), "w:firstLine": None, "w:firstLineChars": None, "w:hangingChars": None})
        if "alignment" in s:
            ensure(ppr, "w:jc", {"w:val": jc_value(s["alignment"], rtl)})

    # ----------------------------------------------------------------- apply
    def apply(self, ctx, s):
        scope = self.scope_for(ctx, s)
        target = s.get("apply_to", "paragraphs")
        res = ctx.styles
        style_id = None
        if "style" in s:
            style_id = res.find_id(s["style"], "paragraph")
            if style_id is None:
                raise ProfileError("paragraph.style: no paragraph style named %r in this document" % s["style"])
        if s.get("changes"):
            self._apply_changes(ctx, scope, s["changes"])
        if self._has_props(s) and target in ("paragraphs", "both"):
            n = 0
            for _name, root, p in iter_paragraphs(ctx, scope, edit=True):
                ppr = ensure_ppr(p)
                if style_id:
                    ensure(ppr, "w:pStyle", {"w:val": style_id})
                rtl = (s["direction"] == "rtl") if "direction" in s else res.is_rtl(p)
                self.apply_ppr(ppr, s, rtl)
                n += 1
            ctx.count("paragraph.paragraphs", n)
        if self._has_props(s) and target in ("styles", "both"):
            self._apply_styles(ctx, s)

    def _has_props(self, s):
        return bool(set(s) - {"changes", "scope", "apply_to", "style_targets"}) or not s.get("changes")

    @staticmethod
    def _effective(res, p, prop):
        if prop == "alignment":
            jc = res.ppr_prop(p, "w:jc")
            return visual_alignment(get(jc, "w:val") if jc is not None else None, res.is_rtl(p))
        if prop == "direction":
            return "rtl" if res.is_rtl(p) else "ltr"
        if prop == "style":
            sid = res.paragraph_style_id(p)
            return res.name_of(sid) if sid else None
        if prop == "line_spacing":
            sp = res.ppr_prop(p, "w:spacing")
            line = get(sp, "w:line") if sp is not None else None
            if line is None:
                return None
            rule = get(sp, "w:lineRule", "auto")
            return ("multiple", int(line) / 240.0) if rule == "auto" else (rule, int(line) / 20.0)
        if prop in ("space_before", "space_after"):
            sp = res.ppr_prop(p, "w:spacing")
            v = get(sp, "w:before" if prop == "space_before" else "w:after") if sp is not None else None
            return int(v) / 20.0 if v is not None else 0.0
        ind = res.ppr_prop(p, "w:ind")
        v = (get(ind, "w:start") or get(ind, "w:left")) if ind is not None else None
        return int(v) / 20.0 if v is not None else 0.0

    @staticmethod
    def _same(cur, want, prop):
        if want == "*":
            return True
        if prop == "line_spacing":
            return cur is not None and cur[0] == want[0] and abs(cur[1] - want[1]) < 0.01
        if prop == "style":
            return cur is not None and str(cur).lower() == str(want).lower()
        if prop in ("space_before", "space_after", "indent_start"):
            return abs(cur - want) < 0.05
        return cur == want

    def _apply_changes(self, ctx, scope, changes):
        res, n = ctx.styles, 0
        for c in changes:
            if c["property"] == "style" and res.find_id(c["to"], "paragraph") is None:
                raise ProfileError("paragraph.changes: no paragraph style named %r in this document" % c["to"])
        for _name, _root, p in iter_paragraphs(ctx, scope, edit=True):
            for c in changes:
                prop = c["property"]
                cur = self._effective(res, p, prop)
                if not any(self._same(cur, w, prop) for w in c["from"]):
                    continue
                ppr = ensure_ppr(p)
                if prop == "style":
                    ensure(ppr, "w:pStyle", {"w:val": res.find_id(c["to"], "paragraph")})
                else:
                    self.apply_ppr(ppr, {prop: c["to"]}, (c["to"] == "rtl") if prop == "direction" else res.is_rtl(p))
                n += 1
        ctx.count("paragraph.changed", n)

    def _apply_styles(self, ctx, s):
        res = ctx.styles
        if res.root is None:
            ctx.warn("paragraph: document has no styles.xml")
            return
        self.pkg_dirty(ctx)
        targets = s.get("style_targets") or ["Normal"]
        n = 0
        want_all = "*" in targets
        for sid, st in res.styles.items():
            if st.get(qn("w:type")) != "paragraph":
                continue
            names = {sid.lower(), res.name_of(sid).lower()}
            if want_all or any(str(t).lower() in names for t in targets):
                ppr = ensure(st, "w:pPr")
                rtl = (s["direction"] == "rtl") if "direction" in s else bool(is_on(ppr.find(qn("w:bidi"))))
                self.apply_ppr(ppr, s, rtl)
                n += 1
        if want_all or any(str(t).lower() == "docdefaults" for t in targets):
            dd = res.root.find(qn("w:docDefaults"))
            if dd is not None:
                pd = dd.find(qn("w:pPrDefault"))
                if pd is None:
                    pd = make("w:pPrDefault")
                    dd.append(pd)
                ppr = pd.find(qn("w:pPr"))
                if ppr is None:
                    ppr = make("w:pPr")
                    pd.append(ppr)
                self.apply_ppr(ppr, s, s.get("direction") == "rtl")
                n += 1
        if n == 0:
            ctx.warn("paragraph: no style matched %s" % targets)
        ctx.count("paragraph.styles", n)
        ctx.reset_styles()

    @staticmethod
    def pkg_dirty(ctx):
        ctx.pkg.xml("word/styles.xml", edit=True)

    # ------------------------------------------------------------------ read
    def _collect(self, ctx_pkg, scope=None):
        from .base import Context
        ctx = Context(ctx_pkg)
        rows = []
        for _n, _r, p in iter_paragraphs(ctx, scope):
            res = ctx.styles
            jc = res.ppr_prop(p, "w:jc")
            sp = res.ppr_prop(p, "w:spacing")
            ind = res.ppr_prop(p, "w:ind")
            rtl = res.is_rtl(p)
            rows.append({
                "rtl": rtl,
                "alignment": visual_alignment(get(jc, "w:val"), rtl),
                "before": get(sp, "w:before"), "after": get(sp, "w:after"),
                "line": (get(sp, "w:line"), get(sp, "w:lineRule")),
                "start": get(ind, "w:left", get(ind, "w:start")), "end": get(ind, "w:right", get(ind, "w:end")),
                "first": get(ind, "w:firstLine"), "hanging": get(ind, "w:hanging"),
                "keepNext": is_on(res.ppr_prop(p, "w:keepNext")),
                "style": res.name_of(res.paragraph_style_id(p)) if res.paragraph_style_id(p) else None,
            })
        return rows

    def read(self, pkg, include_text=False):
        rows = self._collect(pkg)
        if not rows:
            return {"paragraphs": 0}

        def dist(key, conv=lambda x: x):
            c = Counter(conv(r[key]) if r[key] is not None else None for r in rows)
            return {str(k): v for k, v in c.most_common(6)}

        def pt(v):
            return fmt_length(from_twips(v), "pt") if v is not None else None

        def line(v):
            val, rule = v
            if val is None:
                return None
            val = int(val)
            return ("%.2f lines" % (val / 240.0)) if rule in (None, "auto") else "%s %s" % (rule, fmt_length(from_twips(val), "pt"))
        return {
            "paragraphs": len(rows),
            "direction": dist("rtl", lambda v: "rtl" if v else "ltr"),
            "alignment": dist("alignment"),
            "line_spacing": dist("line", line),
            "space_before": dist("before", lambda v: pt(int(v))),
            "space_after": dist("after", lambda v: pt(int(v))),
            "indent_start": dist("start", lambda v: pt(int(v))),
            "indent_end": dist("end", lambda v: pt(int(v))),
            "first_line_indent": dist("first", lambda v: pt(int(v))),
            "hanging_indent": dist("hanging", lambda v: pt(int(v))),
            "styles_used": dist("style"),
        }

    def snapshot(self, pkg):
        rows = self._collect(pkg)
        if not rows:
            return {}
        out = {}
        al = most_common(Counter(r["alignment"] for r in rows), 1.0)
        if al:
            out["alignment"] = al
        out["direction"] = "rtl" if sum(r["rtl"] for r in rows) > len(rows) / 2 else "ltr"
        for key, name in (("before", "space_before"), ("after", "space_after")):
            v = most_common(Counter(r[key] for r in rows), 1.0)
            if v is not None:
                out[name] = round(from_twips(int(v)), 2)
        ln = most_common(Counter(r["line"] for r in rows), 1.0)
        if ln and ln[0] is not None:
            val, rule = int(ln[0]), ln[1]
            out["line_spacing"] = round(val / 240.0, 3) if rule in (None, "auto") else "%s %spt" % (rule, round(from_twips(val), 2))
        for key, name in (("start", "indent_start"), ("end", "indent_end"), ("first", "first_line_indent"), ("hanging", "hanging_indent")):
            v = most_common(Counter(r[key] for r in rows), 1.0)
            if v is not None and int(v) != 0:
                out[name] = round(from_twips(int(v)), 2)
        if "first_line_indent" in out and "hanging_indent" in out:
            del out["hanging_indent"]
        return out
