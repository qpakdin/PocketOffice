# -*- coding: utf-8 -*-
"""Tables group: borders, eraser, row height, AutoFit, width, alignment, padding, header rows, shading, delete."""
import re

from ..core.ooxml import ensure, get, is_on, make, qn, remove, set_attrs, set_onoff
from ..core.textops import has_visible_content
from ..core.units import normalize_color, parse_length, twips
from .base import F, Feature, ProfileError

BORDER_STYLES = ["single", "double", "dotted", "dashed", "dotDash", "thick", "wave", "triple", "none"]
BORDER_TARGETS = ["all", "outside", "inside", "top", "bottom", "left", "right", "inside_h", "inside_v"]
_SIDE_MAP = {"all": ("top", "left", "bottom", "right", "insideH", "insideV"),
             "outside": ("top", "left", "bottom", "right"), "inside": ("insideH", "insideV"),
             "top": ("top",), "bottom": ("bottom",), "left": ("left",), "right": ("right",),
             "inside_h": ("insideH",), "inside_v": ("insideV",)}


def _parse_width(v):
    """'100%' -> ('pct', 5000); '12cm' -> ('dxa', twips); 'auto' -> ('auto', 0)."""
    s = str(v).strip().lower()
    if s == "auto":
        return "auto", 0
    m = re.fullmatch(r"([0-9.]+)\s*%", s)
    if m:
        return "pct", int(round(float(m.group(1)) * 50))
    return "dxa", twips(parse_length(s, "cm"))


class TablesFeature(Feature):
    key = "tables"
    label = "Tables"
    description = ("Borders, eraser, row height, AutoFit, width, alignment, padding, header row, shading, "
                   "delete tables / empty rows. ('View gridlines' is a Word view option, not stored in the file.)")
    fields = [
        F("tables", "json", "Apply to tables", hint='"all" or [1,3] (1-based, in document order)'),
        F("border_target", "choice", "Borders: which edges", choices=BORDER_TARGETS, hint="default: all"),
        F("border_style", "choice", "Border style", choices=BORDER_STYLES),
        F("border_width", "length", "Border width", unit="pt", hint="0.5pt / 1pt"),
        F("border_color", "color", "Border colour"),
        F("erase_borders", "bool", "Eraser: remove all borders"),
        F("row_height", "length", "Row height", unit="cm"),
        F("row_height_rule", "choice", "Row height rule", choices=["atLeast", "exact", "auto"]),
        F("autofit", "choice", "AutoFit", choices=["contents", "window", "fixed"]),
        F("table_width", "text", "Table width", hint="100%  |  12cm  |  auto"),
        F("alignment", "choice", "Table alignment", choices=["left", "center", "right"]),
        F("direction", "choice", "Table direction", choices=["ltr", "rtl"]),
        F("cell_vertical_align", "choice", "Cell vertical alignment", choices=["top", "center", "bottom"]),
        F("cell_padding", "length", "Cell padding (all sides)", unit="cm"),
        F("cell_spacing", "length", "Spacing between cells", unit="pt"),
        F("header_row", "bool", "Repeat first row as header on each page"),
        F("rows_cant_split", "bool", "Don't let rows break across pages"),
        F("shading", "color", "Shade all cells", hint="'auto' clears"),
        F("header_shading", "color", "Shade first row", hint="'auto' clears"),
        F("table_style", "text", "Table style (name or id)"),
        F("delete_empty_rows", "bool", "Delete rows whose cells are all empty"),
        F("delete_empty_tables", "bool", "Delete tables with no content"),
        F("delete_tables", "bool", "Delete the selected tables"),
    ]

    def validate(self, s):
        if "table_width" in s:
            try:
                _parse_width(s["table_width"])
            except ValueError as e:
                raise ProfileError("tables.table_width: %s" % e)
        t = s.get("tables")
        if t is not None and t != "all" and not (isinstance(t, list) and all(isinstance(i, int) for i in t)):
            raise ProfileError("tables.tables: use \"all\" or a list of numbers")
        if "row_height_rule" in s and "row_height" not in s and s["row_height_rule"] != "auto":
            raise ProfileError("tables: row_height_rule needs row_height")
        return s

    # ------------------------------------------------------------------ helpers
    def _selected(self, ctx, s):
        sel = s.get("tables", "all")
        out, i = [], 0
        for name, _k in ctx.pkg.story_parts(self.scope_for(ctx, s).parts):
            root = ctx.pkg.xml(name, edit=True)
            for tbl in root.iter(qn("w:tbl")):
                i += 1
                if sel == "all" or i in sel:
                    out.append(tbl)
        return out

    @staticmethod
    def _tblpr(tbl):
        pr = tbl.find(qn("w:tblPr"))
        if pr is None:
            pr = make("w:tblPr")
            tbl.insert(0, pr)
        return pr

    @staticmethod
    def _rows(tbl):
        return tbl.findall(qn("w:tr"))

    @staticmethod
    def _cells(tr):
        return tr.findall(qn("w:tc"))

    @staticmethod
    def _cell_empty(tc):
        return not any(has_visible_content(p) for p in tc.iter(qn("w:p"))) and tc.find(".//" + qn("w:tbl")) is None

    def _border(self, pr, s):
        borders = ensure(pr, "w:tblBorders")
        if s.get("erase_borders"):
            for side in _SIDE_MAP["all"]:
                set_attrs(ensure(borders, "w:" + side), {"w:val": "none", "w:sz": 0, "w:space": 0, "w:color": "auto"})
            return
        style = s.get("border_style", "single")
        for side in _SIDE_MAP[s.get("border_target", "all")]:
            e = ensure(borders, "w:" + side)
            attrs = {"w:val": style, "w:space": 0}
            if style == "none":
                attrs.update({"w:sz": 0, "w:color": "auto"})
            else:
                if "border_width" in s or get(e, "w:sz") is None:
                    attrs["w:sz"] = max(2, min(96, int(round(s.get("border_width", 0.5) * 8))))
                if "border_color" in s or get(e, "w:color") is None:
                    attrs["w:color"] = s.get("border_color", "auto")
            set_attrs(e, attrs)

    def _apply_one(self, tbl, s):
        pr = self._tblpr(tbl)
        rtl_tbl = is_on(pr.find(qn("w:bidiVisual")))
        if "table_style" in s:
            pr_style = self._style_id
            if pr_style:
                ensure(pr, "w:tblStyle", {"w:val": pr_style})
        if "direction" in s:
            set_onoff(pr, "w:bidiVisual", True if s["direction"] == "rtl" else None)
            rtl_tbl = s["direction"] == "rtl"
        if any(k in s for k in ("border_style", "border_width", "border_color", "border_target")) or s.get("erase_borders"):
            self._border(pr, s)
        if "alignment" in s:
            a = s["alignment"]
            if rtl_tbl and a in ("left", "right"):      # visual in RTL tables
                a = "right" if a == "left" else "left"
            ensure(pr, "w:jc", {"w:val": a})
        if "autofit" in s:
            mode = s["autofit"]
            ensure(pr, "w:tblLayout", {"w:type": "fixed" if mode == "fixed" else "autofit"})
            if mode == "contents":
                ensure(pr, "w:tblW", {"w:type": "auto", "w:w": 0})
                for tc in tbl.iter(qn("w:tc")):
                    w = tc.find(qn("w:tcPr") + "/" + qn("w:tcW"))
                    if w is not None:
                        set_attrs(w, {"w:type": "auto", "w:w": 0})
            elif mode == "window":
                ensure(pr, "w:tblW", {"w:type": "pct", "w:w": 5000})
        if "table_width" in s:
            t, v = _parse_width(s["table_width"])
            ensure(pr, "w:tblW", {"w:type": t, "w:w": v})
        if "cell_padding" in s:
            mar = ensure(pr, "w:tblCellMar")
            for side in ("top", "left", "bottom", "right"):
                ensure(mar, "w:" + side, {"w:w": twips(s["cell_padding"]), "w:type": "dxa"})
        if "cell_spacing" in s:
            if s["cell_spacing"] <= 0:
                remove(pr, "w:tblCellSpacing")
            else:
                ensure(pr, "w:tblCellSpacing", {"w:w": twips(s["cell_spacing"]), "w:type": "dxa"})
        rows = self._rows(tbl)
        for i, tr in enumerate(rows):
            trpr = tr.find(qn("w:trPr"))
            if trpr is None:
                trpr = make("w:trPr")
                tr.insert(1 if tr.find(qn("w:tblPrEx")) is not None else 0, trpr)
            if "row_height" in s or s.get("row_height_rule") == "auto":
                rule = s.get("row_height_rule", "atLeast")
                if rule == "auto":
                    remove(trpr, "w:trHeight")
                else:
                    ensure(trpr, "w:trHeight", {"w:val": twips(s["row_height"]), "w:hRule": rule})
            if "rows_cant_split" in s:
                set_onoff(trpr, "w:cantSplit", True if s["rows_cant_split"] else None)
            if "header_row" in s and i == 0:
                set_onoff(trpr, "w:tblHeader", True if s["header_row"] else None)
            if not len(trpr):
                tr.remove(trpr)
            for tc in self._cells(tr):
                tcpr = tc.find(qn("w:tcPr"))
                if tcpr is None:
                    tcpr = make("w:tcPr")
                    tc.insert(0, tcpr)
                if "cell_vertical_align" in s:
                    ensure(tcpr, "w:vAlign", {"w:val": s["cell_vertical_align"]})
                fill = s.get("header_shading") if i == 0 and "header_shading" in s else s.get("shading")
                if fill is not None:
                    if fill == "auto":
                        remove(tcpr, "w:shd")
                    else:
                        ensure(tcpr, "w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": fill})
                if not len(tcpr):
                    tc.remove(tcpr)

    # -------------------------------------------------------------------- apply
    def apply(self, ctx, s):
        tables = self._selected(ctx, s)
        ctx.count("tables.selected", len(tables))
        if s.get("delete_tables"):
            for t in tables:
                t.getparent().remove(t)
            ctx.count("tables.deleted", len(tables))
            return
        self._style_id = None
        if "table_style" in s:
            self._style_id = ctx.styles.find_id(s["table_style"], "table")
            if self._style_id is None:
                raise ProfileError("tables.table_style: no table style named %r in this document" % s["table_style"])
        if s.get("delete_empty_rows"):
            n = 0
            for t in tables:
                rows = self._rows(t)
                for tr in rows:
                    cells = self._cells(tr)
                    if cells and all(self._cell_empty(c) for c in cells) and len(self._rows(t)) > 1:
                        t.remove(tr)
                        n += 1
            ctx.count("tables.rows_deleted", n)
        if s.get("delete_empty_tables"):
            n = 0
            for t in list(tables):
                if all(self._cell_empty(c) for c in t.iter(qn("w:tc"))):
                    t.getparent().remove(t)
                    tables.remove(t)
                    n += 1
            ctx.count("tables.empty_deleted", n)
        for t in tables:
            self._apply_one(t, s)

    # --------------------------------------------------------------------- read
    def read(self, pkg, include_text=False):
        out, i = [], 0
        for name, _k in pkg.story_parts():
            for tbl in pkg.xml(name).iter(qn("w:tbl")):
                i += 1
                pr = tbl.find(qn("w:tblPr"))
                rows = tbl.findall(qn("w:tr"))
                b = pr.find(qn("w:tblBorders")) if pr is not None else None
                lay = pr.find(qn("w:tblLayout")) if pr is not None else None
                w = pr.find(qn("w:tblW")) if pr is not None else None
                st = pr.find(qn("w:tblStyle")) if pr is not None else None
                h = [get(r.find(qn("w:trPr") + "/" + qn("w:trHeight")), "w:val") for r in rows
                     if r.find(qn("w:trPr") + "/" + qn("w:trHeight")) is not None]
                out.append({
                    "index": i, "part": name, "rows": len(rows),
                    "columns": len(tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol"))),
                    "style": get(st, "w:val") if st is not None else None,
                    "borders": {c.tag.split("}")[1]: get(c, "w:val") for c in b} if b is not None else None,
                    "layout": get(lay, "w:type", "autofit") if lay is not None else "autofit",
                    "width": ("%s %s" % (get(w, "w:w"), get(w, "w:type"))) if w is not None else None,
                    "row_heights_twips": sorted(set(h)) if h else [],
                    "header_row": any(r.find(qn("w:trPr") + "/" + qn("w:tblHeader")) is not None for r in rows[:1]),
                    "empty_rows": sum(1 for r in rows if all(self._cell_empty(c) for c in r.findall(qn("w:tc")))),
                })
        return {"count": len(out), "tables": out}

    def snapshot(self, pkg):
        return {}
