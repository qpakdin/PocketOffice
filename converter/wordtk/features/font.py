# -*- coding: utf-8 -*-
"""Font group: face, size, colour, highlight, effects, language, symbols."""
import re
from collections import Counter

from ..core.ooxml import ensure, ensure_rpr, get, is_on, make, qn, remove, set_attrs, set_onoff
from ..core.scope import iter_paragraphs
from ..core.textops import Sym, compile_find, paragraph_runs, replace_in_paragraph
from ..core.units import half_points, twips, from_twips
from .base import F, Feature, ProfileError, most_common

HIGHLIGHTS = ["none", "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan",
              "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black", "white"]
_HL = {"green": "green", "brightGreen": "green"}
UNDERLINES = ["none", "single", "words", "double", "thick", "dotted", "dottedHeavy", "dash",
              "dashedHeavy", "dotDash", "dotDotDash", "wave", "wavyHeavy", "wavyDouble"]


class FontFeature(Feature):
    key = "font"
    label = "Font"
    description = "Font face/size/colour/highlight/effects/language and symbol insertion."
    fields = [
        F("font", "text", "Font name (Latin)", hint="e.g. Calibri"),
        F("font_cs", "text", "Font name (complex script: Persian/Arabic)", hint="e.g. B Nazanin; default = font"),
        F("font_east_asia", "text", "Font name (East Asian)"),
        F("sync_scripts", "bool", "Apply font/size/bold/italic to complex script too", default=True),
        F("size", "float", "Size (pt)", minimum=1, maximum=1638),
        F("size_cs", "float", "Size for complex script (pt)", minimum=1, maximum=1638),
        F("bold", "bool", "Bold"),
        F("italic", "bool", "Italic"),
        F("underline", "choice", "Underline", choices=UNDERLINES),
        F("underline_color", "color", "Underline colour"),
        F("strike", "choice", "Strikethrough", choices=["none", "single", "double"]),
        F("vertical_align", "choice", "Superscript / subscript", choices=["baseline", "superscript", "subscript"]),
        F("small_caps", "bool", "Small caps"),
        F("all_caps", "bool", "All caps"),
        F("hidden", "bool", "Hidden"),
        F("color", "color", "Font colour", hint="003399 / #003399 / red / auto"),
        F("highlight", "choice", "Highlight colour", choices=HIGHLIGHTS),
        F("shading", "color", "Character shading fill", hint="'auto' clears"),
        F("char_spacing", "length", "Character spacing (+expanded / -condensed)", unit="pt"),
        F("char_scale", "float", "Character scale (%)", minimum=1, maximum=600),
        F("kerning", "length", "Kerning for fonts above (0 = off)", unit="pt"),
        F("position", "length", "Raise/lower text (+/-)", unit="pt"),
        F("lang_latin", "text", "Proofing language (Latin)", hint="en-US"),
        F("lang_bidi", "text", "Proofing language (complex script)", hint="fa-IR / ar-SA"),
        F("lang_east_asia", "text", "Proofing language (East Asian)"),
        F("run_direction", "choice", "Run direction (w:rtl)", choices=["ltr", "rtl"]),
        F("clear_direct", "bool", "Clear direct character formatting first"),
        F("symbols", "json", "Insert symbols",
          hint='[{"find":"(c)","font":"Symbol","char":"F0A9","regex":false}]'),
        F("changes", "json", "Change (replace one value with another)",
          hint='[{"property":"font","from":"Arial","to":"Calibri"},{"property":"color","from":"red","to":"black"},{"property":"size","from":12,"to":10}]'),
        F("apply_to", "choice", "Apply to", choices=["runs", "styles", "both"],
          hint="runs = direct formatting; styles = style definitions"),
        F("style_targets", "list", "Style targets when apply_to uses styles",
          hint="names/ids, 'docDefaults' or '*' (default: Normal)"),
    ]

    CHANGE_PROPS = ("font", "font_cs", "size", "color", "highlight", "bold", "italic", "underline")

    @classmethod
    def norm_change(cls, c):
        from ..core.units import normalize_color, parse_number
        if not isinstance(c, dict) or c.get("property") not in cls.CHANGE_PROPS or "to" not in c:
            raise ProfileError("font.changes: each item needs property (%s), from and to" % ", ".join(cls.CHANGE_PROPS))
        prop = c["property"]
        frm = c.get("from", "*")
        frm = [x.strip() for x in frm.split(",")] if isinstance(frm, str) and prop != "color" and "," in frm else frm
        frm = frm if isinstance(frm, list) else [frm]

        def one(v):
            if isinstance(v, str) and v.strip() in ("", "*"):
                return "*"
            try:
                if prop in ("font", "font_cs"):
                    return str(v).strip()
                if prop == "size":
                    return parse_number(v)
                if prop == "color":
                    return normalize_color(v)
                if prop in ("bold", "italic"):
                    return v if isinstance(v, bool) else str(v).strip().lower() in ("1", "true", "yes", "on")
                if prop == "highlight":
                    if str(v) not in HIGHLIGHTS:
                        raise ValueError("choose one of %s" % ", ".join(HIGHLIGHTS))
                    return str(v)
                if prop == "underline":
                    if str(v) not in UNDERLINES:
                        raise ValueError("choose one of %s" % ", ".join(UNDERLINES))
                    return str(v)
            except ValueError as e:
                raise ProfileError("font.changes[%s]: %s" % (prop, e))
        to = one(c["to"])
        if to == "*":
            raise ProfileError("font.changes[%s]: 'to' is required" % prop)
        return {"property": prop, "from": [one(x) for x in frm], "to": to}

    def validate(self, s):
        if "changes" in s:
            if not isinstance(s["changes"], list):
                raise ProfileError("font.changes must be a list")
            s["changes"] = [self.norm_change(c) for c in s["changes"]]
        for sym in s.get("symbols", []) or []:
            if not isinstance(sym, dict) or not sym.get("find") or not sym.get("char"):
                raise ProfileError("font.symbols: each item needs 'find' and 'char' (hex code like F0A9)")
            if not re.fullmatch(r"[0-9A-Fa-f]{2,4}", str(sym["char"])):
                raise ProfileError("font.symbols: char must be 2-4 hex digits, got %r" % sym["char"])
        return s

    # ------------------------------------------------------------- rPr writer
    def apply_rpr(self, rpr, s):
        if s.get("clear_direct"):
            for c in list(rpr):
                if c.tag != qn("w:rStyle") and c.tag != qn("w:rPrChange"):
                    rpr.remove(c)
        sync = s.get("sync_scripts", True)
        name = s.get("font")
        if name or s.get("font_cs") or s.get("font_east_asia"):
            rf = ensure(rpr, "w:rFonts")
            if name:
                set_attrs(rf, {"w:ascii": name, "w:hAnsi": name, "w:asciiTheme": None, "w:hAnsiTheme": None})
            cs = s.get("font_cs") or (name if sync else None)
            if cs:
                set_attrs(rf, {"w:cs": cs, "w:cstheme": None})
            ea = s.get("font_east_asia")
            if ea:
                set_attrs(rf, {"w:eastAsia": ea, "w:eastAsiaTheme": None})
            rf.attrib.pop(qn("w:hint"), None)
        if "size" in s or "size_cs" in s:
            if "size" in s:
                ensure(rpr, "w:sz", {"w:val": half_points(s["size"])})
            cs = s.get("size_cs", s.get("size") if sync else None)
            if cs is not None:
                ensure(rpr, "w:szCs", {"w:val": half_points(cs)})
        if "bold" in s:
            set_onoff(rpr, "w:b", s["bold"])
            if sync:
                set_onoff(rpr, "w:bCs", s["bold"])
        if "italic" in s:
            set_onoff(rpr, "w:i", s["italic"])
            if sync:
                set_onoff(rpr, "w:iCs", s["italic"])
        if "underline" in s or "underline_color" in s:
            u = ensure(rpr, "w:u")
            if "underline" in s:
                u.set(qn("w:val"), s["underline"])
            elif get(u, "w:val") is None:
                u.set(qn("w:val"), "single")
            if "underline_color" in s:
                set_attrs(u, {"w:color": s["underline_color"], "w:themeColor": None})
        if "strike" in s:
            remove(rpr, "w:strike")
            remove(rpr, "w:dstrike")
            if s["strike"] == "single":
                ensure(rpr, "w:strike")
            elif s["strike"] == "double":
                ensure(rpr, "w:dstrike")
        if "vertical_align" in s:
            ensure(rpr, "w:vertAlign", {"w:val": s["vertical_align"]})
        for k, tag in (("small_caps", "w:smallCaps"), ("all_caps", "w:caps"), ("hidden", "w:vanish")):
            if k in s:
                set_onoff(rpr, tag, s[k])
        if "color" in s:
            set_attrs(ensure(rpr, "w:color"), {"w:val": s["color"], "w:themeColor": None, "w:themeShade": None, "w:themeTint": None})
        if "highlight" in s:
            ensure(rpr, "w:highlight", {"w:val": s["highlight"]})
        if "shading" in s:
            if s["shading"] == "auto":
                remove(rpr, "w:shd")
            else:
                ensure(rpr, "w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": s["shading"]})
        if "char_spacing" in s:
            ensure(rpr, "w:spacing", {"w:val": twips(s["char_spacing"])})
        if "char_scale" in s:
            ensure(rpr, "w:w", {"w:val": int(round(s["char_scale"]))})
        if "kerning" in s:
            ensure(rpr, "w:kern", {"w:val": half_points(s["kerning"])})
        if "position" in s:
            ensure(rpr, "w:position", {"w:val": half_points(s["position"])})
        if any(k in s for k in ("lang_latin", "lang_bidi", "lang_east_asia")):
            lg = ensure(rpr, "w:lang")
            for k, a in (("lang_latin", "w:val"), ("lang_bidi", "w:bidi"), ("lang_east_asia", "w:eastAsia")):
                if k in s:
                    lg.set(qn(a), s[k])
        if "run_direction" in s:
            if s["run_direction"] == "rtl":
                set_onoff(rpr, "w:rtl", True)
            else:
                remove(rpr, "w:rtl")

    _PROP_KEYS = {"font", "font_cs", "font_east_asia", "size", "size_cs", "bold", "italic", "underline",
                  "underline_color", "strike", "vertical_align", "small_caps", "all_caps", "hidden", "color",
                  "highlight", "shading", "char_spacing", "char_scale", "kerning", "position", "lang_latin",
                  "lang_bidi", "lang_east_asia", "run_direction", "clear_direct"}

    # ------------------------------------------------------------------ apply
    def apply(self, ctx, s):
        scope = self.scope_for(ctx, s)
        target = s.get("apply_to", "runs")
        has_props = bool(self._PROP_KEYS & set(s))
        if s.get("changes"):
            self._apply_changes(ctx, scope, s["changes"])
        if has_props and target in ("runs", "both"):
            nr = 0
            for _n, _root, p in iter_paragraphs(ctx, scope, edit=True):
                any_text = False
                for r in paragraph_runs(p):
                    self.apply_rpr(ensure_rpr(r), s)
                    nr += 1
                    any_text = True
                ppr = p.find(qn("w:pPr"))
                mark = ppr.find(qn("w:rPr")) if ppr is not None else None
                if mark is not None or not any_text:
                    from ..core.ooxml import ensure_ppr
                    if mark is None:
                        mark = ensure(ensure_ppr(p), "w:rPr")
                    self.apply_rpr(mark, s)
            ctx.count("font.runs", nr)
        if has_props and target in ("styles", "both"):
            self._apply_styles(ctx, s)
        if s.get("symbols"):
            n = 0
            for _n, _root, p in iter_paragraphs(ctx, scope, edit=True):
                for sym in s["symbols"]:
                    rx = compile_find(sym["find"], bool(sym.get("regex")), bool(sym.get("match_case", True)))
                    n += replace_in_paragraph(p, rx, Sym(sym.get("font", "Symbol"), str(sym["char"]).upper()))
            ctx.count("font.symbols_inserted", n)

    @staticmethod
    def _effective(res, r, p, prop):
        """Effective value of one run property (comparable with a normalised 'from')."""
        if prop in ("font", "font_cs"):
            rf = res.rpr_prop(r, p, "w:rFonts")
            attr = "w:ascii" if prop == "font" else "w:cs"
            v = get(rf, attr)
            if v is None and rf is not None:
                th = get(rf, "w:asciiTheme" if prop == "font" else "w:cstheme")
                v = ("theme:" + th) if th else None
            return v
        if prop == "size":
            e = res.rpr_prop(r, p, "w:sz")
            return float(get(e, "w:val")) / 2 if e is not None and get(e, "w:val") else None
        if prop == "color":
            e = res.rpr_prop(r, p, "w:color")
            v = get(e, "w:val") if e is not None else None
            return "auto" if v in (None, "auto") else v.upper()
        if prop in ("bold", "italic"):
            return bool(is_on(res.rpr_prop(r, p, "w:b" if prop == "bold" else "w:i")))
        e = res.rpr_prop(r, p, "w:highlight" if prop == "highlight" else "w:u")
        return get(e, "w:val") if e is not None else "none"

    @staticmethod
    def _same(cur, want, prop):
        if want == "*":
            return True
        if cur is None:
            return want in ("auto", "none", False)
        if prop in ("font", "font_cs"):
            return str(cur).lower() == str(want).lower()
        if prop == "size":
            return abs(float(cur) - float(want)) < 0.01
        return cur == want

    def _apply_changes(self, ctx, scope, changes):
        res, n = ctx.styles, 0
        for _n, _root, p in iter_paragraphs(ctx, scope, edit=True):
            for r in paragraph_runs(p):
                if r.find(qn("w:t")) is None and r.find(qn("w:sym")) is None:
                    continue
                for c in changes:
                    prop, to = c["property"], c["to"]
                    cur = self._effective(res, r, p, prop)
                    if not any(self._same(cur, w, prop) for w in c["from"]):
                        continue
                    if cur is not None and self._same(cur, to, prop) and "*" not in c["from"]:
                        continue
                    rpr = ensure_rpr(r)
                    spec = {"sync_scripts": False}
                    if prop == "size":
                        spec.update(size=to, size_cs=to)
                    elif prop in ("bold", "italic"):
                        spec.update({prop: to, "sync_scripts": True})
                    else:
                        spec[prop] = to
                    self.apply_rpr(rpr, spec)
                    n += 1
        ctx.count("font.changed", n)

    def _apply_styles(self, ctx, s):
        res = ctx.styles
        if res.root is None:
            ctx.warn("font: document has no styles.xml")
            return
        ctx.pkg.xml("word/styles.xml", edit=True)
        targets = s.get("style_targets") or ["Normal"]
        want_all = "*" in targets
        n = 0
        for sid, st in res.styles.items():
            if st.get(qn("w:type")) not in ("paragraph", "character"):
                continue
            names = {sid.lower(), res.name_of(sid).lower()}
            if want_all or any(str(t).lower() in names for t in targets):
                self.apply_rpr(ensure(st, "w:rPr"), s)
                n += 1
        if want_all or any(str(t).lower() == "docdefaults" for t in targets):
            dd = res.root.find(qn("w:docDefaults"))
            if dd is not None:
                rd = dd.find(qn("w:rPrDefault"))
                if rd is None:
                    rd = make("w:rPrDefault")
                    dd.insert(0, rd)
                rpr = rd.find(qn("w:rPr"))
                if rpr is None:
                    rpr = make("w:rPr")
                    rd.append(rpr)
                self.apply_rpr(rpr, s)
                n += 1
        if n == 0:
            ctx.warn("font: no style matched %s" % targets)
        ctx.count("font.styles", n)
        ctx.reset_styles()

    # ------------------------------------------------------------------- read
    def _scan(self, pkg):
        from .base import Context
        ctx = Context(pkg)
        res = ctx.styles
        rows, syms = [], Counter()
        for _n, root, p in iter_paragraphs(ctx):
            for r in paragraph_runs(p):
                if r.find(qn("w:t")) is None and r.find(qn("w:sym")) is None:
                    continue
                for sy in r.findall(qn("w:sym")):
                    syms[(get(sy, "w:font"), get(sy, "w:char"))] += 1
                cs = bool(is_on(res.rpr_prop(r, p, "w:rtl")))
                rf = res.rpr_prop(r, p, "w:rFonts")
                sz = res.rpr_prop(r, p, "w:sz")
                szcs = res.rpr_prop(r, p, "w:szCs")
                col = res.rpr_prop(r, p, "w:color")
                hl = res.rpr_prop(r, p, "w:highlight")
                lang = res.rpr_prop(r, p, "w:lang")
                u = res.rpr_prop(r, p, "w:u")
                rows.append({
                    "font": get(rf, "w:ascii") or ("theme:" + get(rf, "w:asciiTheme") if get(rf, "w:asciiTheme") else None),
                    "font_cs": get(rf, "w:cs") or ("theme:" + get(rf, "w:cstheme") if get(rf, "w:cstheme") else None),
                    "size": float(get(sz, "w:val")) / 2 if sz is not None and get(sz, "w:val") else None,
                    "size_cs": float(get(szcs, "w:val")) / 2 if szcs is not None and get(szcs, "w:val") else None,
                    "bold": bool(is_on(res.rpr_prop(r, p, "w:b"))), "italic": bool(is_on(res.rpr_prop(r, p, "w:i"))),
                    "underline": get(u, "w:val") if u is not None else None,
                    "color": get(col, "w:val") if col is not None else None,
                    "highlight": get(hl, "w:val") if hl is not None else None,
                    "lang_latin": get(lang, "w:val"), "lang_bidi": get(lang, "w:bidi"),
                })
        return rows, syms

    def read(self, pkg, include_text=False):
        rows, syms = self._scan(pkg)
        d = {"text_runs": len(rows)}

        def dist(k):
            c = Counter(r[k] for r in rows)
            return {str(a): b for a, b in c.most_common(6)}
        for k in ("font", "font_cs", "size", "size_cs", "color", "highlight", "underline", "lang_latin", "lang_bidi"):
            d[k] = dist(k)
        d["bold_runs"] = sum(r["bold"] for r in rows)
        d["italic_runs"] = sum(r["italic"] for r in rows)
        d["symbols"] = [{"font": f, "char": c, "count": n} for (f, c), n in syms.most_common()]
        return d

    def snapshot(self, pkg):
        rows, _ = self._scan(pkg)
        out = {}
        for k in ("font", "font_cs", "size", "size_cs", "color", "lang_latin", "lang_bidi"):
            v = most_common(Counter(r[k] for r in rows), 1.0)   # only when the whole file agrees
            if v is not None and not (isinstance(v, str) and v.startswith("theme:")):
                out[k] = v
        return out
