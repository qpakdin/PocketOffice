# -*- coding: utf-8 -*-
"""Lists group: bullets, numbering, existing ('current') lists and their indents."""
import random
from collections import Counter, defaultdict

from lxml import etree

from ..core.ooxml import NS, RELT, ensure, get, insert_ordered, make, qn, remove, set_attrs
from ..core.scope import iter_paragraphs
from ..core.units import fmt_length, from_twips, twips
from .base import F, Feature, ProfileError

NUMBERING = "word/numbering.xml"
NUM_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
NUMBER_FORMATS = ["decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "decimalZero",
                  "ordinal", "hindiNumbers", "arabicAlpha", "arabicAbjad", "hebrew1"]
_BULLETS = [("", "Symbol"), ("o", "Courier New"), ("", "Wingdings")]
_NUM_CYCLE = ["decimal", "lowerLetter", "lowerRoman"]


def _numbering_root(pkg, create=False):
    root = pkg.try_xml(NUMBERING, edit=create)
    if root is None and create:
        root = etree.Element(qn("w:numbering"), nsmap={"w": NS["w"]})
        pkg.add_xml_part(NUMBERING, root, NUM_CT)
        pkg.add_rel(pkg.main, RELT + "numbering", NUMBERING)
    return root


def _next_id(root, tag, attr):
    ids = [int(get(e, attr, "0")) for e in root.findall(qn(tag)) if (get(e, attr, "0") or "0").isdigit()]
    return max(ids + [0]) + 1


def _lvl(i, kind, fmt, start, text, left, hanging, bullet_font=None):
    lvl = make("w:lvl", {"w:ilvl": i})
    lvl.append(make("w:start", {"w:val": start}))
    lvl.append(make("w:numFmt", {"w:val": "bullet" if kind == "bullet" else fmt}))
    lvl.append(make("w:lvlText", {"w:val": text}))
    lvl.append(make("w:lvlJc", {"w:val": "left"}))
    ppr = make("w:pPr")
    ppr.append(make("w:ind", {"w:left": left, "w:hanging": hanging}))
    lvl.append(ppr)
    if kind == "bullet" and bullet_font:
        rpr = make("w:rPr")
        rpr.append(make("w:rFonts", {"w:ascii": bullet_font, "w:hAnsi": bullet_font, "w:cs": bullet_font, "w:hint": "default"}))
        lvl.append(rpr)
    return lvl


def build_abstract(root, s, kind):
    """Create an abstractNum (9 levels) at the proper position and return its id."""
    aid = _next_id(root, "w:abstractNum", "w:abstractNumId")
    ab = make("w:abstractNum", {"w:abstractNumId": aid})
    ab.append(make("w:nsid", {"w:val": "%08X" % random.randint(0x10000000, 0x7FFFFFFF)}))
    multilevel = bool(s.get("multilevel"))
    ab.append(make("w:multiLevelType", {"w:val": "multilevel" if multilevel else "hybridMultilevel"}))
    start_l = twips(s.get("indent_start", 36.0))
    hang = twips(s.get("indent_hanging", 18.0))
    step = twips(s.get("indent_step", 36.0))
    for i in range(9):
        left = start_l + i * step
        if kind == "bullet":
            ch, font = _BULLETS[i % 3]
            if i == 0 and s.get("bullet_char"):
                ch, font = s["bullet_char"], s.get("bullet_font", "")
            ab.append(_lvl(i, "bullet", None, 1, ch, left, hang, font or None))
        else:
            fmt = s.get("number_format", "decimal") if i == 0 or multilevel else _NUM_CYCLE[i % 3]
            if multilevel:
                fmt = s.get("number_format", "decimal")
                text = ".".join("%%%d" % (j + 1) for j in range(i + 1)) + "."
            elif i == 0 and s.get("number_text"):
                text = s["number_text"]
            else:
                text = "%%%d." % (i + 1)
            ab.append(_lvl(i, "number", fmt, s.get("start_at", 1) if i == 0 else 1, text, left, hang))
    first_num = root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(ab)
    else:
        root.append(ab)
    return aid


def new_num(root, abstract_id, restart_at=None):
    nid = _next_id(root, "w:num", "w:numId")
    n = make("w:num", {"w:numId": nid})
    n.append(make("w:abstractNumId", {"w:val": abstract_id}))
    if restart_at is not None:
        ov = make("w:lvlOverride", {"w:ilvl": 0})
        ov.append(make("w:startOverride", {"w:val": restart_at}))
        n.append(ov)
    root.append(n)
    return nid


class ListsFeature(Feature):
    key = "lists"
    label = "Lists & Numbering"
    description = "Convert to bullets/numbering, remove lists, restyle current lists, change list indents."
    fields = [
        F("mode", "choice", "Convert paragraphs to", choices=["bullet", "numbered", "none"],
          hint="omit to keep list type and only change indents"),
        F("apply_to", "choice", "Which paragraphs", choices=["existing_lists", "all_paragraphs"],
          hint="existing_lists = only paragraphs already in a list; use scope.styles to narrow"),
        F("bullet_char", "text", "Bullet character", hint="default Word bullet; e.g. •  ▪  ◦"),
        F("bullet_font", "text", "Bullet font", hint="e.g. Arial (empty = document font)"),
        F("number_format", "choice", "Number format", choices=NUMBER_FORMATS),
        F("number_text", "text", "Number text pattern (level 1)", hint="%1.  |  (%1)  |  %1-"),
        F("multilevel", "bool", "Multilevel numbering 1. / 1.1. / 1.1.1."),
        F("start_at", "int", "Start at", minimum=0, maximum=9999),
        F("restart", "bool", "Restart numbering for each separate list", default=True),
        F("indent_start", "length", "List indent: level 1 start", unit="pt"),
        F("indent_hanging", "length", "List indent: hanging", unit="pt"),
        F("indent_step", "length", "List indent: extra per level", unit="pt"),
        F("use_list", "int", "Join the existing list with this numId (see 'read')", minimum=1),
    ]

    # ------------------------------------------------------------------ apply
    def apply(self, ctx, s):
        scope = self.scope_for(ctx, s)
        pkg = ctx.pkg
        mode = s.get("mode")
        apply_to = s.get("apply_to", "existing_lists")
        sel = []  # (p, original numId or None)
        for _n, _r, p in iter_paragraphs(ctx, scope, edit=True):
            ppr = p.find(qn("w:pPr"))
            numpr = ppr.find(qn("w:numPr")) if ppr is not None else None
            nid = get(numpr.find(qn("w:numId")), "w:val") if numpr is not None and numpr.find(qn("w:numId")) is not None else None
            if nid is None:
                st_np = ctx.styles.ppr_prop(p, "w:numPr")
                if st_np is not None and st_np.find(qn("w:numId")) is not None:
                    nid = get(st_np.find(qn("w:numId")), "w:val")
            in_list = nid not in (None, "0")
            if apply_to == "existing_lists" and not in_list:
                continue
            if apply_to == "all_paragraphs" and not in_list and not "".join(t.text or "" for t in p.iter(qn("w:t"))).strip():
                continue
            sel.append((p, nid if in_list else None))
        if not sel:
            ctx.warn("lists: no matching paragraphs")
            return

        if mode == "none":
            n = 0
            for p, nid in sel:
                ppr = p.find(qn("w:pPr"))
                if ppr is not None and ppr.find(qn("w:numPr")) is not None:
                    ppr.remove(ppr.find(qn("w:numPr")))
                    n += 1
                if nid is not None and ctx.styles.ppr_prop(p, "w:numPr") is not None:
                    from ..core.ooxml import ensure_ppr
                    np_ = ensure(ensure_ppr(p), "w:numPr")
                    ensure(np_, "w:numId", {"w:val": 0})
                    n += 1
            ctx.count("lists.removed", n)
            return

        root = _numbering_root(pkg, create=bool(mode) or "use_list" in s)
        if root is None:
            ctx.warn("lists: document has no numbering definitions")
            return

        if "use_list" in s:
            if not any(get(x, "w:numId") == str(s["use_list"]) for x in root.findall(qn("w:num"))):
                raise ProfileError("lists.use_list: numId %s does not exist" % s["use_list"])
            for p, _ in sel:
                np_ = ensure(self._ppr(p), "w:numPr")
                ensure(np_, "w:ilvl", {"w:val": get(np_.find(qn("w:ilvl")), "w:val", 0)})
                ensure(np_, "w:numId", {"w:val": s["use_list"]})
            ctx.count("lists.joined", len(sel))

        elif mode in ("bullet", "numbered"):
            kind = "bullet" if mode == "bullet" else "number"
            ab_id = build_abstract(root, s, kind)
            shared = None
            if kind == "bullet" or not s.get("restart", True):
                shared = new_num(root, ab_id)
            for grp in self._groups(sel):
                nid = shared if shared else new_num(root, ab_id, s.get("start_at", 1))
                for p in grp:
                    np_ = ensure(self._ppr(p), "w:numPr")
                    ensure(np_, "w:ilvl", {"w:val": get(np_.find(qn("w:ilvl")), "w:val", 0)})
                    ensure(np_, "w:numId", {"w:val": nid})
                    # numbering indents must not be shadowed by stale direct indents
                    ind = self._ppr(p).find(qn("w:ind"))
                    if ind is not None:
                        self._ppr(p).remove(ind)
            ctx.count("lists.converted", len(sel))

        elif any(k in s for k in ("indent_start", "indent_hanging", "indent_step")):
            n = self._reindent(root, sel, s)
            ctx.count("lists.levels_reindented", n)
            # direct paragraph indents would hide the change
            for p, _ in sel:
                ind = self._ppr(p).find(qn("w:ind"))
                if ind is not None:
                    self._ppr(p).remove(ind)

    @staticmethod
    def _ppr(p):
        from ..core.ooxml import ensure_ppr
        return ensure_ppr(p)

    @staticmethod
    def _groups(sel):
        groups, cur, prev, prev_key = [], [], None, None
        for p, nid in sel:
            contiguous = prev is not None and p.getprevious() is prev and nid == prev_key
            if not contiguous and cur:
                groups.append(cur)
                cur = []
            cur.append(p)
            prev, prev_key = p, nid
        if cur:
            groups.append(cur)
        return groups

    def _reindent(self, root, sel, s):
        nums = {get(n, "w:numId"): n for n in root.findall(qn("w:num"))}
        abstracts = {get(a, "w:abstractNumId"): a for a in root.findall(qn("w:abstractNum"))}
        used = set()
        for _p, nid in sel:
            n = nums.get(nid)
            if n is not None:
                aid = get(n.find(qn("w:abstractNumId")), "w:val")
                if aid in abstracts:
                    used.add(aid)
        count = 0
        for aid in used:
            lvls = abstracts[aid].findall(qn("w:lvl"))

            def left_of(l):
                ind = l.find("w:pPr/w:ind", NS)
                return int(get(ind, "w:left", get(ind, "w:start", 0)) or 0) if ind is not None else 0
            cur0 = left_of(lvls[0]) if lvls else 720
            cur_step = (left_of(lvls[1]) - cur0) if len(lvls) > 1 else 720
            base = twips(s["indent_start"]) if "indent_start" in s else cur0
            step = twips(s["indent_step"]) if "indent_step" in s else cur_step
            relevel = "indent_start" in s or "indent_step" in s
            for l in lvls:
                i = int(get(l, "w:ilvl", "0"))
                ppr = l.find(qn("w:pPr"))
                if ppr is None:
                    ppr = insert_ordered(l, make("w:pPr"))
                ind = ensure(ppr, "w:ind")
                if relevel:
                    set_attrs(ind, {"w:left": base + i * step, "w:start": None})
                if "indent_hanging" in s:
                    set_attrs(ind, {"w:hanging": twips(s["indent_hanging"]), "w:firstLine": None})
                count += 1
        return count

    # ------------------------------------------------------------------- read
    def read(self, pkg, include_text=False):
        root = pkg.try_xml(NUMBERING)
        from .base import Context
        ctx = Context(pkg)
        used = Counter()
        for _n, _r, p in iter_paragraphs(ctx):
            np_ = ctx.styles.ppr_prop(p, "w:numPr")
            if np_ is not None and np_.find(qn("w:numId")) is not None:
                used[get(np_.find(qn("w:numId")), "w:val")] += 1
        out = {"numbering_part": root is not None, "paragraphs_in_lists": sum(v for k, v in used.items() if k != "0"), "lists": []}
        if root is None:
            return out
        abstracts = {get(a, "w:abstractNumId"): a for a in root.findall(qn("w:abstractNum"))}
        for n in root.findall(qn("w:num")):
            nid = get(n, "w:numId")
            aid = get(n.find(qn("w:abstractNumId")), "w:val")
            a = abstracts.get(aid)
            lv0 = a.find(qn("w:lvl")) if a is not None else None
            fmt = get(lv0.find(qn("w:numFmt")), "w:val") if lv0 is not None and lv0.find(qn("w:numFmt")) is not None else None
            txt = get(lv0.find(qn("w:lvlText")), "w:val") if lv0 is not None and lv0.find(qn("w:lvlText")) is not None else None
            ind = lv0.find("w:pPr/w:ind", NS) if lv0 is not None else None
            out["lists"].append({
                "numId": int(nid), "abstractNumId": aid,
                "type": "bullet" if fmt == "bullet" else "numbered", "level1_format": fmt,
                "level1_text": txt if fmt != "bullet" else "(bullet U+%04X)" % ord(txt[0]) if txt else None,
                "levels": len(a.findall(qn("w:lvl"))) if a is not None else 0,
                "indent_start": fmt_length(from_twips(int(get(ind, "w:left", get(ind, "w:start", 0)))), "cm") if ind is not None else None,
                "hanging": fmt_length(from_twips(int(get(ind, "w:hanging", 0))), "cm") if ind is not None else None,
                "paragraphs": used.get(nid, 0)})
        return out
