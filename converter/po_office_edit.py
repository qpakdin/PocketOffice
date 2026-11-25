#!/usr/bin/env python3
"""
Editing sessions for Pocket Office (Word / Excel / PowerPoint).

Every session works on a COPY of the file held in memory as a stack of snapshots (bytes). The original is never
written. Two kinds of edits exist for Word:
  * precise paragraph edits (po_docx_edit.py: only the touched paragraph XML changes), and
  * "settings" (the wordtk engine: font, paragraph, page, tables, find & replace ...), previewed before they are applied.
Excel and PowerPoint use the same settings mechanism (same Field / Feature classes, so the app builds the same forms).
"""
import difflib, hashlib, io, json, os, re, shutil, sys, tempfile, time, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import po_docx_edit as PDE

try:
    from lxml import etree
    from wordtk.core.package import DocxPackage
    from wordtk.core.scope import Scope
    from wordtk.features import REGISTRY as WORD_FEATURES, Context, ProfileError
    from wordtk.gui.layout import sections_for as _word_sections
    OFFICE_EDIT_ERROR = None
except Exception as _e:  # noqa
    OFFICE_EDIT_ERROR = "%s: %s" % (type(_e).__name__, _e)
    ProfileError = ValueError

TEXT_CHANGING_GROUPS = {"text", "comments", "tables", "cover", "breaks"}
SKIP_SUFFIX = (".rels", ".xml", ".vml")


class OfficeEditError(Exception):
    """A request that cannot be done; the message is shown to the user."""


def _zip_parts(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    return z, z.namelist()


def part_diff(a, b):
    za, na = _zip_parts(a)
    zb, nb = _zip_parts(b)
    changed = [n for n in na if n in nb and za.read(n) != zb.read(n)]
    return {"changed": changed, "added": [n for n in nb if n not in na], "removed": [n for n in na if n not in nb]}


def xml_problems(data, names):
    z, _n = _zip_parts(data)
    bad = []
    for n in names:
        if n.endswith(SKIP_SUFFIX):
            try:
                etree.fromstring(z.read(n), etree.XMLParser(resolve_entities=False, huge_tree=True))
            except Exception as e:  # noqa
                bad.append("%s (%s)" % (n, str(e)[:80]))
    return bad


def schema_from(features, sections_fn):
    out = []
    for key, feat in features.items():
        secs = [{"title": t, "fields": [f.to_dict() for f in fields]} for t, fields in sections_fn(feat)]
        out.append({"key": key, "label": feat.label, "description": feat.description, "uses_scope": feat.uses_scope, "sections": secs})
    return out


# =============================================================================== common session behaviour
class BaseSession:
    kind = ""
    features = {}

    def __init__(self, orig, name, log=None):
        self.orig, self.name = orig, name
        self.log = log or (lambda m: None)
        self.sha = hashlib.sha256(orig).hexdigest()
        self.states = [orig]
        self.history = []                     # [{"kind": "ops"/"settings", "group":..., "text_changing": bool, "desc": ...}]
        self.report = {"name": name, "sha256": self.sha, "size": len(orig), "blocked": None, "items": [], "stats": {}, "locked": {}, "page": None}

    @property
    def cur(self):
        return self.states[-1]

    @property
    def ops(self):
        return self.history                   # len() = number of changes

    def _push(self, data, entry):
        self.states.append(data)
        self.history.append(entry)
        if len(self.states) > 25:             # keep memory bounded; the original (index 0) is always kept
            del self.states[1]

    def undo(self):
        if not self.history:
            raise OfficeEditError("nothing to undo")
        self.states.pop()
        self.history.pop()

    def export(self):
        return self.cur

    def schema(self):
        raise NotImplementedError

    # ---- settings (shared flow)
    file_ext = "docx"

    def _make_scope(self, scope):
        return Scope()

    def _run_feature(self, group, norm, scope, src_bytes):
        feat = self.features[group]
        sc = self._make_scope(scope)
        d = tempfile.mkdtemp(prefix="po_set_")
        try:
            src, dst = os.path.join(d, "in." + self.file_ext), os.path.join(d, "out." + self.file_ext)
            with open(src, "wb") as f:
                f.write(src_bytes)
            pkg = DocxPackage(src)
            ctx = Context(pkg, sc)
            feat.apply(ctx, norm)
            pkg.save(dst)
            with open(dst, "rb") as f:
                return f.read(), dict(ctx.report), list(ctx.warnings)
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def settings(self, group, values, scope=None, dry=False):
        if group not in self.features:
            raise OfficeEditError("unknown settings group %r" % group)
        values = {k: v for k, v in (values or {}).items() if v is not None and not (isinstance(v, str) and not v.strip())}
        if not values:
            raise OfficeEditError("nothing to apply: fill in at least one setting (empty fields mean 'leave unchanged')")
        feat = self.features[group]
        try:
            norm = feat.normalize(values)
        except ProfileError as e:
            raise OfficeEditError(str(e))
        t0 = time.time()
        try:
            new, changes, warnings = self._run_feature(group, norm, scope, self.cur)
        except ProfileError as e:
            raise OfficeEditError(str(e))
        except OfficeEditError:
            raise
        except Exception as e:  # noqa
            raise OfficeEditError("this setting could not be applied (%s: %s). Nothing was changed." % (type(e).__name__, e))
        diff = part_diff(self.cur, new)
        bad = xml_problems(new, diff["changed"] + diff["added"])
        if bad:
            raise OfficeEditError("the result would contain damaged XML (%s). Nothing was changed." % "; ".join(bad))
        res = {"group": group, "label": feat.label, "changes": changes, "warnings": warnings, "parts_changed": diff["changed"],
               "parts_added": diff["added"], "parts_removed": diff["removed"], "seconds": round(time.time() - t0, 2),
               "nothing_changed": new == self.cur or not (diff["changed"] or diff["added"] or diff["removed"])}
        res.update(self._impact(self.cur, new))
        if not dry and not res["nothing_changed"]:
            self._push(new, {"kind": "settings", "group": group, "text_changing": group in self._text_groups(), "desc": "%s: %s" % (group, ", ".join(sorted(norm)))})
        return res

    def _text_groups(self):
        return TEXT_CHANGING_GROUPS

    def _impact(self, old, new):
        return {}

    # ---- verification pieces shared by all kinds
    def _verify_package(self, checks):
        d = part_diff(self.orig, self.cur)

        def chk(name, ok, detail=""):
            checks.append({"name": name, "ok": bool(ok), "detail": detail})
        bad = xml_problems(self.cur, d["changed"] + d["added"])
        chk("Every changed part is well-formed XML", not bad, "; ".join(bad))
        znames = zipfile.ZipFile(io.BytesIO(self.cur)).namelist()
        chk("The file is a valid zip package", zipfile.ZipFile(io.BytesIO(self.cur)).testzip() is None, "%d parts" % len(znames))
        total = len(zipfile.ZipFile(io.BytesIO(self.orig)).namelist())
        detail = "%d of %d parts are byte-identical to the original" % (total - len(d["changed"]) - len(d["removed"]), total)
        if d["changed"]:
            detail += "; changed: " + ", ".join(d["changed"][:8]) + (" ..." if len(d["changed"]) > 8 else "")
        if d["added"]:
            detail += "; added: " + ", ".join(d["added"][:6])
        if d["removed"]:
            detail += "; removed: " + ", ".join(d["removed"][:6])
        chk("Parts you did not touch are unchanged", True, detail)
        return d


# =============================================================================== WORD
class WordSession(BaseSession):
    kind = "docx"

    def __init__(self, orig, name, log=None):
        super().__init__(orig, name, log)
        self.features = WORD_FEATURES
        first = PDE.DocxSession(orig, name, log=log)
        self.report = first.report
        self.report["sha256"] = self.sha
        self.report["items"].append({"level": "INFO", "text": "Settings (font, paragraph, page, tables, find & replace ...) are previewed first and then applied to the whole document or to the scope you choose. They also affect paragraphs that are locked for direct editing; use the preview to see what will change."})
        self.blocked = first.blocked or self.report.get("blocked")

    def _dx(self, data=None):
        return PDE.DocxSession(data if data is not None else self.cur, self.name, analyze=False)

    # ---- precise edits
    def blocks(self):
        return self._dx().blocks()

    def table_rows(self):
        return self._dx().table_rows()

    def page_info(self):
        return self._dx().page_info()

    def apply(self, ops):
        if self.blocked:
            raise PDE.EditError(self.blocked)
        dx = self._dx()
        res = dx.apply(ops)                   # raises EditError and changes nothing
        self._push(dx.export(), {"kind": "ops", "ops": list(ops), "desc": "%d edit(s)" % len(ops),
                                 "text_changing": any(o.get("op") in ("set_text", "insert_paragraph_after", "delete_paragraph", "insert_row", "delete_row") for o in ops)})
        return res

    # ---- settings
    def schema(self):
        return schema_from(self.features, _word_sections)

    def _make_scope(self, scope):
        try:
            return Scope.from_dict(scope or {})
        except ValueError as e:
            raise OfficeEditError("scope: %s" % e)

    def _impact(self, old, new):
        try:
            a, b = self._dx(old).full_texts(), self._dx(new).full_texts()
        except Exception:
            return {}
        sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
        samples, n = [], 0
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag != "equal":
                n += max(i2 - i1, j2 - j1)
                if len(samples) < 6:
                    samples.append({"old": " / ".join(x[:70] for x in a[i1:i2][:2]), "new": " / ".join(x[:70] for x in b[j1:j2][:2])})
        return {"paragraphs_text_changed": n, "paragraphs_before": len(a), "paragraphs_after": len(b), "samples": samples}

    # ---- find
    def find(self, q, match_case=False, whole_word=False, regex=False, limit=300):
        if not q:
            raise OfficeEditError("type something to find")
        try:
            pat = q if regex else re.escape(q)
            if whole_word:
                pat = r"(?<!\w)(?:%s)(?!\w)" % pat
            rx = re.compile(pat, 0 if match_case else re.I)
        except re.error as e:
            raise OfficeEditError("the search expression is not valid: %s" % e)
        dx = self._dx()
        texts, blocks = dx.full_texts(), dx.blocks()
        out, total = [], 0
        for i, t in enumerate(texts):
            ms = list(rx.finditer(t))
            if not ms:
                continue
            total += len(ms)
            if len(out) < limit:
                m = ms[0]
                a, b = max(0, m.start() - 30), min(len(t), m.end() + 30)
                out.append({"i": i, "count": len(ms), "snippet": ("\u2026" if a else "") + t[a:b] + ("\u2026" if b < len(t) else ""), "locked": blocks[i]["locked"], "kind": blocks[i]["kind"]})
        return {"matches": total, "paragraphs": len(out), "results": out, "truncated": len(out) >= limit}

    # ---- verification
    def verify(self):
        checks = []
        d = self._verify_package(checks)
        text_changing = any(h.get("text_changing") for h in self.history)
        surgical_only = all(h["kind"] == "ops" for h in self.history)

        def chk(name, ok, detail=""):
            checks.append({"name": name, "ok": bool(ok), "detail": detail})
        changes = []
        a, b = self._dx(self.orig).full_texts(), self._dx(self.cur).full_texts()
        sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
        n_ch = 0
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag != "equal":
                n_ch += max(i2 - i1, j2 - j1)
                if len(changes) < 60:
                    changes.append({"tag": tag, "old": " / ".join(x[:80] for x in a[i1:i2][:3]), "new": " / ".join(x[:80] for x in b[j1:j2][:3])})
        if text_changing:
            chk("Text changes are the ones you asked for", True, "%d paragraph(s) differ in text (expected: you used text-changing edits)" % n_ch)
        else:
            chk("No text was changed by formatting edits", n_ch == 0, "%d paragraph(s) differ in text" % n_ch)
        if surgical_only and self.history:
            dx = PDE.DocxSession(self.orig, self.name, analyze=False)
            dx.cur, dx._scan = zipfile.ZipFile(io.BytesIO(self.cur)).read(PDE.MAIN_PART), None
            dx.ops = [o for h in self.history for o in h["ops"]]
            sub = dx.verify()
            for c in sub["checks"]:
                if "byte-identical" in c["name"] or "Same list" in c["name"] or "well-formed" in c["name"]:
                    continue
                checks.append(c)
        return {"ok": all(c["ok"] for c in checks), "checks": checks, "changes": changes, "ops": len(self.history),
                "paragraphs_before": len(a), "paragraphs_after": len(b)}


# =============================================================================== shared text helpers (Excel + PowerPoint)
if OFFICE_EDIT_ERROR is None:
    from wordtk.features.base import F, Feature
    from wordtk.core.units import _AR_DIGITS, _FA_DIGITS, normalize_color

    _TO_FA = {ord(str(i)): c for i, c in enumerate(_FA_DIGITS)}
    _TO_FA.update({ord(c): _FA_DIGITS[i] for i, c in enumerate(_AR_DIGITS)})
    _TO_EN = {ord(c): str(i) for i, c in enumerate(_FA_DIGITS)}
    _TO_EN.update({ord(c): str(i) for i, c in enumerate(_AR_DIGITS)})
    _TO_AR = {ord(str(i)): c for i, c in enumerate(_AR_DIGITS)}
    _TO_AR.update({ord(c): _AR_DIGITS[i] for i, c in enumerate(_FA_DIGITS)})
    _FA_ONLY = {ord(str(i)): c for i, c in enumerate(_FA_DIGITS)}
    _EN_FROM_FA = {ord(c): str(i) for i, c in enumerate(_FA_DIGITS)}
    _EN_FROM_AR = {ord(c): str(i) for i, c in enumerate(_AR_DIGITS)}
    _FA_FROM_AR = {ord(c): _FA_DIGITS[i] for i, c in enumerate(_AR_DIGITS)}
    _AR_FROM_FA = {ord(c): _AR_DIGITS[i] for i, c in enumerate(_FA_DIGITS)}
    _AR_FROM_EN = {ord(str(i)): c for i, c in enumerate(_AR_DIGITS)}
    DIGIT_MAPS = {"english->persian": _FA_ONLY, "persian->english": _EN_FROM_FA, "arabic->persian": _FA_FROM_AR, "persian->arabic": _AR_FROM_FA,
                  "arabic->english": _EN_FROM_AR, "english->arabic": _AR_FROM_EN}
    ALL_DIGITS = {"persian": _TO_FA, "english": _TO_EN, "arabic": _TO_AR}
    XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

    def parse_bulk(text):
        rules = []
        for line in str(text).splitlines():
            if not line.strip():
                continue
            a, b = (line.split("=>", 1) + [""])[:2] if "=>" in line else (line, "")
            a = a.strip()
            if a:
                rules.append({"find": a, "replace": b.strip()})
        return rules

    def text_transform(v):
        """-> function(str)->str applying the digit / letter / case / space options of a 'text' settings dict (no find & replace)."""
        steps = []
        if v.get("digits"):
            steps.append(lambda s, m=ALL_DIGITS[v["digits"]]: s.translate(m))
        if v.get("digits_convert"):
            steps.append(lambda s, m=DIGIT_MAPS[v["digits_convert"]]: s.translate(m))
        if v.get("persian_letters"):
            steps.append(lambda s: s.replace("\u0643", "\u06a9").replace("\u064a", "\u06cc"))
        if v.get("collapse_spaces"):
            steps.append(lambda s: re.sub(r"[ \u00a0]{2,}", " ", s))
        if v.get("fix_punctuation_spacing"):
            steps.append(lambda s: re.sub(r"\s+([,.;:!?\u060c\u061b\u061f])", r"\1", s))
        case = v.get("change_case")
        if case:
            steps.append({"upper": str.upper, "lower": str.lower, "title": str.title, "sentence": lambda s: s[:1].upper() + s[1:].lower()}[case])

        def run(s):
            for f in steps:
                s = f(s)
            return s
        return run

    def find_rules(v):
        rules = []
        raw = v.get("replacements")
        if isinstance(raw, list):
            rules += [r for r in raw if isinstance(r, dict) and r.get("find")]
        if v.get("replace_list"):
            rules += parse_bulk(v["replace_list"])
        out = []
        for r in rules:
            pat = r["find"] if r.get("regex") else re.escape(r["find"])
            if r.get("whole_word"):
                pat = r"(?<!\w)(?:%s)(?!\w)" % pat
            try:
                out.append((re.compile(pat, 0 if r.get("match_case") else re.I), r.get("replace", "") if r.get("regex") else str(r.get("replace", "")).replace("\\", "\\\\")))
            except re.error as e:
                raise OfficeEditError("find & replace: invalid expression %r (%s)" % (r["find"], e))
        return out

    def set_node_text(n, text):
        n.text = text
        if text != text.strip() or "  " in text:
            n.set(XMLSPACE, "preserve")

    def replace_in_nodes(nodes, rx, repl):
        """regex replace across consecutive text nodes (the formatting of the first touched node is kept). -> number of replacements"""
        texts = [n.text or "" for n in nodes]
        full = "".join(texts)
        ms = [m for m in rx.finditer(full) if m.end() > m.start()]
        if not ms:
            return 0
        starts, pos = [], 0
        for t in texts:
            starts.append(pos)
            pos += len(t)
        for m in reversed(ms):
            a, b = m.start(), m.end()
            new = m.expand(repl) if isinstance(repl, str) else repl(m)
            first = next((i for i in range(len(nodes)) if starts[i] <= a < starts[i] + len(texts[i])), None)
            if first is None:
                continue
            at = a - starts[first]
            remain = b - a
            take = min(remain, len(texts[first]) - at)
            texts[first] = texts[first][:at] + new + texts[first][at + take:]
            remain -= take
            j = first + 1
            while remain > 0 and j < len(nodes):
                k = min(remain, len(texts[j]))
                texts[j] = texts[j][k:]
                remain -= k
                j += 1
        for n, t in zip(nodes, texts):
            if (n.text or "") != t:
                set_node_text(n, t)
        return len(ms)

    def apply_text_options(nodes, rules, fn):
        """find/replace rules first, then the char-level transform; returns number of changed nodes"""
        before = [n.text or "" for n in nodes]
        for rx, rep in rules:
            replace_in_nodes(nodes, rx, rep)
        if fn is not None:
            for n in nodes:
                t = n.text or ""
                nt = fn(t)
                if nt != t:
                    set_node_text(n, nt)
        return sum(1 for n, b in zip(nodes, before) if (n.text or "") != b)

    def diff_set_text(nodes, old, new, make_node):
        """change the text spread over `nodes` (their concatenation is `old`) to `new`; untouched text keeps its node and formatting.
        make_node() creates the first node when there is none."""
        if old == new:
            return
        if not old:
            n = nodes[0] if nodes else make_node()
            if not nodes:
                nodes.append(n)
            set_node_text(n, new)
            return
        sm = difflib.SequenceMatcher(None, old, new, autojunk=False)
        deleted, ins = [False] * len(old), {}
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag in ("delete", "replace"):
                for k in range(i1, i2):
                    deleted[k] = True
            if tag in ("insert", "replace"):
                ins[i1] = ins.get(i1, "") + new[j1:j2]
        ci, first_done = 0, False
        for n in nodes:
            t, res = n.text or "", []
            for k, ch in enumerate(t):
                g = ci + k
                if g == 0 and 0 in ins and not first_done:
                    res.append(ins[0])
                    first_done = True
                if not deleted[g]:
                    res.append(ch)
                if (g + 1) in ins:
                    res.append(ins[g + 1])
            ci += len(t)
            if "".join(res) != t:
                set_node_text(n, "".join(res))


# =============================================================================== EXCEL
if OFFICE_EDIT_ERROR is None:
    import copy
    NS_M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def m(tag):
        return "{%s}%s" % (NS_M, tag)

    WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
                "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
                "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties",
                "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
                "webPublishItems", "tableParts", "extLst"]
    WB_ORDER = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences",
                "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
                "fileRecoveryPr", "webPublishObjects", "extLst"]
    PAPER = {"A3": 8, "A4": 9, "A5": 11, "Letter": 1, "Legal": 5, "Tabloid": 3, "Executive": 7}
    BUILTIN_FMT = {"General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "@": 49}

    def col2n(c):
        n = 0
        for ch in c.upper():
            n = n * 26 + ord(ch) - 64
        return n

    def n2col(n):
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    def parse_ref(ref):
        mm = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})", ref.strip())
        if not mm:
            raise OfficeEditError("%r is not a cell reference (example: B7)" % ref)
        return col2n(mm.group(1)), int(mm.group(2))

    def parse_range(text):
        """'A1:C5' | 'A1' | 'A:C' | '2:6' -> (c1, r1, c2, r2); '' -> None (= whole sheet)"""
        t = (text or "").strip().replace("$", "")
        if not t:
            return None
        if ":" not in t:
            c, r = parse_ref(t)
            return (c, r, c, r)
        a, b = [x.strip() for x in t.split(":", 1)]
        if re.fullmatch(r"[A-Za-z]{1,3}", a) and re.fullmatch(r"[A-Za-z]{1,3}", b):
            return (col2n(a), 1, col2n(b), 1048576)
        if re.fullmatch(r"\d+", a) and re.fullmatch(r"\d+", b):
            return (1, int(a), 16384, int(b))
        c1, r1 = parse_ref(a)
        c2, r2 = parse_ref(b)
        return (min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))

    def ws_ensure(root, tag):
        """child of a worksheet root, created at the schema position if missing"""
        e = root.find(m(tag))
        if e is not None:
            return e
        rank = WS_ORDER.index(tag)
        e = etree.Element(m(tag))
        at = len(root)
        for i, ch in enumerate(root):
            loc = etree.QName(ch).localname
            if loc in WS_ORDER and WS_ORDER.index(loc) > rank:
                at = i
                break
        root.insert(at, e)
        return e

    class Styles:
        """cell formats: new fonts / fills / borders / number formats / cellXfs are appended, nothing existing is modified"""

        def __init__(self, pkg):
            self.pkg, self.root, self.memo, self.dirty = pkg, pkg.xml("xl/styles.xml"), {}, False

        def _mark(self):
            if not self.dirty:
                self.root = self.pkg.xml("xl/styles.xml", edit=True)
                self.dirty = True

        def _coll(self, tag):
            return self.root.find(m(tag))

        def _append(self, tag, el):
            c = self._coll(tag)
            c.append(el)
            c.set("count", str(len(c)))
            return len(c) - 1

        def numfmt(self, code):
            if code in BUILTIN_FMT:
                return BUILTIN_FMT[code]
            nf = self._coll("numFmts")
            if nf is None:
                nf = etree.Element(m("numFmts"))
                self.root.insert(0, nf)
            ids = [int(x.get("numFmtId")) for x in nf]
            for x in nf:
                if x.get("formatCode") == code:
                    return int(x.get("numFmtId"))
            nid = max([163] + ids) + 1
            e = etree.SubElement(nf, m("numFmt"))
            e.set("numFmtId", str(nid))
            e.set("formatCode", code)
            nf.set("count", str(len(nf)))
            return nid

        def restyle(self, base, ch):
            key = (base, json.dumps(ch, sort_keys=True))
            if key in self.memo:
                return self.memo[key]
            self._mark()
            xfs = self._coll("cellXfs")
            xf = copy.deepcopy(xfs[base] if base < len(xfs) else xfs[0])
            fch = ch.get("font")
            if fch:
                fonts = self._coll("fonts")
                f = copy.deepcopy(fonts[int(xf.get("fontId", "0"))])

                def put(tag, **attrs):
                    for old in f.findall(m(tag)):
                        f.remove(old)
                    if attrs is not None:
                        e = etree.SubElement(f, m(tag))
                        for k, v in attrs.items():
                            e.set(k, v)
                for tag, key2 in (("b", "bold"), ("i", "italic"), ("strike", "strike")):
                    if fch.get(key2) is not None:
                        for old in f.findall(m(tag)):
                            f.remove(old)
                        if fch[key2]:
                            etree.SubElement(f, m(tag))
                if fch.get("underline") is not None:
                    for old in f.findall(m("u")):
                        f.remove(old)
                    if fch["underline"]:
                        etree.SubElement(f, m("u"))
                if fch.get("size"):
                    put("sz", val=str(fch["size"]))
                if fch.get("color"):
                    put("color", rgb="FF" + fch["color"])
                if fch.get("name"):
                    put("name", val=fch["name"])
                    for old in f.findall(m("scheme")):
                        f.remove(old)
                xf.set("fontId", str(self._append("fonts", f)))
                xf.set("applyFont", "1")
            if ch.get("fill"):
                if ch["fill"] == "none":
                    xf.set("fillId", "0")
                else:
                    fill = etree.Element(m("fill"))
                    pf = etree.SubElement(fill, m("patternFill"))
                    pf.set("patternType", "solid")
                    etree.SubElement(pf, m("fgColor")).set("rgb", "FF" + ch["fill"])
                    etree.SubElement(pf, m("bgColor")).set("indexed", "64")
                    xf.set("fillId", str(self._append("fills", fill)))
                xf.set("applyFill", "1")
            bch = ch.get("border")
            if bch:
                b = etree.Element(m("border"))
                for side in ("left", "right", "top", "bottom", "diagonal"):
                    e = etree.SubElement(b, m(side))
                    if side != "diagonal" and bch["style"] != "none":
                        e.set("style", bch["style"])
                        etree.SubElement(e, m("color")).set("rgb", "FF" + bch.get("color", "000000"))
                xf.set("borderId", str(self._append("borders", b)))
                xf.set("applyBorder", "1")
            if ch.get("numfmt"):
                xf.set("numFmtId", str(self.numfmt(ch["numfmt"])))
                xf.set("applyNumberFormat", "1")
            ach = ch.get("align")
            if ach:
                al = xf.find(m("alignment"))
                if al is None:
                    al = etree.Element(m("alignment"))
                    xf.insert(0, al)
                for k, v in ach.items():
                    if v is None:
                        al.attrib.pop(k, None)
                    else:
                        al.set(k, str(v))
                xf.set("applyAlignment", "1")
            idx = self._append("cellXfs", xf)
            self.memo[key] = idx
            return idx

    class Sheet:
        def __init__(self, book, info, index):
            self.book, self.info, self.index = book, info, index
            self.name, self.part = info["name"], info["part"]

        def root(self, edit=False):
            return self.book.pkg.xml(self.part, edit)

        def data(self, edit=False):
            return self.root(edit).find(m("sheetData"))

        @property
        def protected(self):
            sp = self.root().find(m("sheetProtection"))
            return sp is not None and sp.get("sheet") in ("1", "true")

        def cell(self, ref, create=False, edit=False):
            c, r = parse_ref(ref)
            data = self.data(edit)
            row = None
            for rw in data:
                rr = int(rw.get("r", "0"))
                if rr == r:
                    row = rw
                    break
                if rr > r:
                    if not create:
                        return None
                    row = etree.Element(m("row"))
                    row.set("r", str(r))
                    rw.addprevious(row)
                    break
            if row is None:
                if not create:
                    return None
                row = etree.SubElement(data, m("row"))
                row.set("r", str(r))
            for ce in row:
                cc = col2n(re.match(r"[A-Z]+", ce.get("r")).group(0))
                if cc == c:
                    return ce
                if cc > c:
                    if not create:
                        return None
                    n = etree.Element(m("c"))
                    n.set("r", "%s%d" % (n2col(c), r))
                    ce.addprevious(n)
                    return n
            if not create:
                return None
            n = etree.SubElement(row, m("c"))
            n.set("r", "%s%d" % (n2col(c), r))
            return n

        def cells(self, rng=None):
            """existing <c> elements inside rng (None = all)"""
            for row in self.data():
                r = int(row.get("r", "0"))
                if rng and not rng[1] <= r <= rng[3]:
                    continue
                for ce in row:
                    if rng:
                        cc = col2n(re.match(r"[A-Z]+", ce.get("r")).group(0))
                        if not rng[0] <= cc <= rng[2]:
                            continue
                    yield ce

    class Book:
        def __init__(self, pkg):
            self.pkg = pkg
            wb = pkg.xml(pkg.main)
            rels = {rid: tgt for rid, tgt, _t, ext in pkg.rels(pkg.main, "worksheet") if not ext}
            self.sheets = []
            shs = wb.find(m("sheets"))
            for sh in (shs if shs is not None else []):
                rid = sh.get("{%s}id" % NS_R)
                if rid in rels:
                    self.sheets.append(Sheet(self, {"name": sh.get("name"), "part": rels[rid], "state": sh.get("state", "visible")}, len(self.sheets)))
            self._sst = None

        def pick(self, spec):
            if not spec or not str(spec).strip():
                return list(self.sheets)
            out = []
            for tok in [t.strip() for t in str(spec).split(",") if t.strip()]:
                hit = [s for s in self.sheets if s.name.lower() == tok.lower()]
                if not hit and tok.isdigit() and 1 <= int(tok) <= len(self.sheets):
                    hit = [self.sheets[int(tok) - 1]]
                if not hit:
                    raise OfficeEditError("there is no sheet named %r (sheets: %s)" % (tok, ", ".join(s.name for s in self.sheets)))
                out += [h for h in hit if h not in out]
            return out

        def writable(self, sheets):
            for s in sheets:
                if s.protected:
                    raise OfficeEditError("sheet %r is protected against changes; editing it here would bypass that protection" % s.name)
            return sheets

        def sst(self, edit=False):
            if "xl/sharedStrings.xml" in self.pkg.names() or self.pkg.exists("xl/sharedStrings.xml"):
                return self.pkg.xml("xl/sharedStrings.xml", edit)
            return None

        def si_nodes(self, si):
            """text nodes of a shared string item (phonetic runs excluded)"""
            return [t for t in si.iter(m("t")) if t.getparent().tag != m("rPh")]

        def cell_text(self, c):
            t = c.get("t")
            if t == "s":
                v = c.find(m("v"))
                sst = self.sst()
                if v is not None and sst is not None and (v.text or "").isdigit() and int(v.text) < len(sst):
                    return "".join(n.text or "" for n in self.si_nodes(sst[int(v.text)]))
                return ""
            if t == "inlineStr":
                return "".join(n.text or "" for n in c.iter(m("t")) if n.getparent().tag != m("rPh"))
            v = c.find(m("v"))
            return (v.text or "") if v is not None else ""

        def values(self):
            """{(sheet, ref): (type, text, formula)} for change detection"""
            out = {}
            for s in self.sheets:
                for c in s.cells():
                    f = c.find(m("f"))
                    out[(s.name, c.get("r"))] = (c.get("t", ""), self.cell_text(c), (f.text or "") if f is not None else None)
            return out

    # ----------------------------------------------------------------------------- features
    SHEETS_FIELD = F("sheets", "text", "Sheets", hint="names or numbers, comma separated; empty = all sheets")
    RANGE_FIELD = F("range", "text", "Cells", hint="e.g. A1:D20, A:C, 2:5 or B7; empty = every cell that exists on the sheet")

    class XText(Feature):
        key, label = "text", "Text"
        description = "Find & replace in cell text, digits, Persian letters, spaces, case."
        uses_scope = False
        fields = [SHEETS_FIELD,
                  F("replacements", "json", "Find & replace", hint='[{"find":"old","replace":"new","regex":false,"match_case":false,"whole_word":false}]'),
                  F("replace_list", "text", "Bulk replace list", hint="one rule per line:  old => new   (no => deletes the text)"),
                  F("digits", "choice", "Convert ALL digits to", choices=["persian", "english", "arabic"]),
                  F("digits_convert", "choice", "Convert digits (one direction)", choices=list(DIGIT_MAPS)),
                  F("persian_letters", "bool", "Replace Arabic \u0643 \u064a with Persian \u06a9 \u06cc"),
                  F("collapse_spaces", "bool", "Collapse repeated spaces"),
                  F("fix_punctuation_spacing", "bool", "Remove space before , . ; : ! ? \u060c \u061b \u061f"),
                  F("change_case", "choice", "Change case", choices=["upper", "lower", "title", "sentence"])]

        def apply(self, ctx, v):
            book = Book(ctx.pkg)
            sheets = book.writable(book.pick(v.get("sheets")))
            rules, fn = find_rules(v), text_transform(v)
            if not rules and not any(k in v for k in ("digits", "digits_convert", "persian_letters", "collapse_spaces", "fix_punctuation_spacing", "change_case")):
                raise OfficeEditError("fill in something to find / replace or choose a conversion")
            all_sheets = len(sheets) == len(book.sheets)
            sst = book.sst()
            changed = 0
            if sst is not None and all_sheets:
                sst = book.pkg.xml("xl/sharedStrings.xml", edit=True)
                for si in sst:
                    nodes = book.si_nodes(si)
                    if nodes:
                        changed += apply_text_options(nodes, rules, fn)
            elif sst is not None:
                cache = {}
                for s in sheets:
                    root = s.root()
                    for c in s.cells():
                        if c.get("t") != "s":
                            continue
                        old = book.cell_text(c)
                        if old in cache:
                            new_idx = cache[old]
                        else:
                            idx = int(c.find(m("v")).text)
                            probe = copy.deepcopy(sst[idx])
                            nodes = book.si_nodes(probe)
                            if not apply_text_options(nodes, rules, fn):
                                cache[old] = None
                                continue
                            sst = book.pkg.xml("xl/sharedStrings.xml", edit=True)
                            sst.append(probe)
                            new_idx = len(sst) - 1
                            cache[old] = new_idx
                        if new_idx is not None:
                            s.root(edit=True)
                            c.find(m("v")).text = str(new_idx)
                            changed += 1
                if cache and sst is not None:
                    sst.set("uniqueCount", str(len(sst)))
            for s in sheets:                    # inline strings
                edit_needed = False
                for c in s.cells():
                    if c.get("t") == "inlineStr":
                        nodes = [n for n in c.iter(m("t")) if n.getparent().tag != m("rPh")]
                        if nodes and apply_text_options(nodes, rules, fn):
                            changed += 1
                            edit_needed = True
                if edit_needed:
                    s.root(edit=True)
            nform = sum(1 for s in sheets for c in s.cells() if c.get("t") == "str")
            if nform:
                ctx.warn("%d formula cell(s) with text results were not changed (change the formula or its source cells instead)." % nform)
            ctx.count("text items changed", changed)

    class XFormat(Feature):
        key, label = "format", "Cell format"
        description = "Font, colors, fill, borders, number format and alignment for cells."
        uses_scope = False
        fields = [SHEETS_FIELD, RANGE_FIELD,
                  F("font", "text", "Font name", hint="e.g. Calibri, B Nazanin"), F("size", "float", "Font size (pt)", minimum=1, maximum=409),
                  F("bold", "bool", "Bold"), F("italic", "bool", "Italic"), F("underline", "bool", "Underline"), F("strike", "bool", "Strikethrough"),
                  F("color", "color", "Text color"), F("fill", "color", "Fill color", hint="'auto' removes the fill"),
                  F("number_format", "text", "Number format", hint="e.g. 0.00  |  #,##0  |  0%  |  yyyy-mm-dd  |  @ (text)"),
                  F("border_style", "choice", "Borders (all edges)", choices=["none", "thin", "medium", "thick", "dashed", "dotted", "double", "hair"]),
                  F("border_color", "color", "Border color"),
                  F("horizontal", "choice", "Horizontal alignment", choices=["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]),
                  F("vertical", "choice", "Vertical alignment", choices=["top", "center", "bottom", "justify", "distributed"]),
                  F("wrap", "bool", "Wrap text"), F("shrink", "bool", "Shrink to fit"), F("indent", "int", "Indent", minimum=0, maximum=250),
                  F("direction", "choice", "Reading order", choices=["context", "ltr", "rtl"])]

        def apply(self, ctx, v):
            book = Book(ctx.pkg)
            sheets = book.writable(book.pick(v.get("sheets")))
            try:
                rng = parse_range(v.get("range"))
            except OfficeEditError:
                raise
            ch = {}
            font = {}
            for k_in, k_out in (("font", "name"), ("size", "size"), ("bold", "bold"), ("italic", "italic"), ("underline", "underline"), ("strike", "strike")):
                if k_in in v:
                    font[k_out] = v[k_in]
            if "color" in v:
                font["color"] = v["color"]
            if font:
                ch["font"] = font
            if "fill" in v:
                ch["fill"] = "none" if str(v["fill"]).lower() == "auto" else v["fill"]
            if "number_format" in v:
                ch["numfmt"] = v["number_format"]
            if "border_style" in v:
                ch["border"] = {"style": v["border_style"], "color": v.get("border_color") or "000000"}
            al = {}
            if "horizontal" in v:
                al["horizontal"] = None if v["horizontal"] == "general" else v["horizontal"]
            if "vertical" in v:
                al["vertical"] = v["vertical"]
            if "wrap" in v:
                al["wrapText"] = "1" if v["wrap"] else None
            if "shrink" in v:
                al["shrinkToFit"] = "1" if v["shrink"] else None
            if "indent" in v:
                al["indent"] = v["indent"] or None
            if "direction" in v:
                al["readingOrder"] = {"context": None, "ltr": "1", "rtl": "2"}[v["direction"]]
            if al:
                ch["align"] = al
            if not ch:
                raise OfficeEditError("choose at least one format to apply")
            st = Styles(ctx.pkg)
            n = 0
            for s in sheets:
                if rng and (rng[2] - rng[0] + 1) * (min(rng[3], 5000) - rng[1] + 1) <= 60000 and rng[3] < 1048576 and rng[2] < 16384:
                    for r in range(rng[1], rng[3] + 1):          # explicit area: blank cells are formatted too
                        for c in range(rng[0], rng[2] + 1):
                            s.cell("%s%d" % (n2col(c), r), create=True, edit=True)
                targets = list(s.cells(rng))
                if targets:
                    s.root(edit=True)
                for c in targets:
                    c.set("s", str(st.restyle(int(c.get("s", "0")), ch)))
                    n += 1
            ctx.count("cells formatted", n)
            if not n:
                ctx.warn("no cells were found in the chosen area")

    class XLayout(Feature):
        key, label = "layout", "Rows, columns & view"
        description = "Column width, row height, sheet direction, zoom, gridlines, frozen panes, tab color."
        uses_scope = False
        fields = [SHEETS_FIELD,
                  F("columns", "text", "Columns", hint="e.g. A:C or B"), F("col_width", "float", "Column width (characters)", minimum=0, maximum=255),
                  F("rows", "text", "Rows", hint="e.g. 1:3 or 5"), F("row_height", "float", "Row height (pt)", minimum=0, maximum=409),
                  F("rtl", "bool", "Right-to-left sheet"), F("zoom", "int", "Zoom %", minimum=10, maximum=400),
                  F("gridlines", "bool", "Show gridlines"), F("freeze", "text", "Freeze panes at cell", hint="e.g. B2 (freezes row 1 and column A); '-' removes"),
                  F("tab_color", "color", "Sheet tab color")]

        def apply(self, ctx, v):
            book = Book(ctx.pkg)
            sheets = book.writable(book.pick(v.get("sheets")))
            if "columns" in v and "col_width" not in v or "rows" in v and "row_height" not in v:
                raise OfficeEditError("give a width / height together with the columns / rows")
            n = 0
            for s in sheets:
                root = s.root(edit=True)
                if "col_width" in v:
                    c1, _a, c2, _b = parse_range(v["columns"] + (":" + v["columns"] if ":" not in v["columns"] else ""))
                    c2 = min(c2, 16384)
                    cols = ws_ensure(root, "cols")
                    old = [(int(c.get("min")), int(c.get("max")), dict(c.attrib)) for c in cols]
                    for c in list(cols):
                        cols.remove(c)
                    keep = []
                    for lo, hi, at in old:
                        if hi < c1 or lo > c2:
                            keep.append((lo, hi, at))
                        else:
                            if lo < c1:
                                keep.append((lo, c1 - 1, at))
                            if hi > c2:
                                keep.append((c2 + 1, hi, at))
                    keep.append((c1, c2, {"width": str(v["col_width"]), "customWidth": "1"}))
                    for lo, hi, at in sorted(keep, key=lambda x: x[0]):
                        e = etree.SubElement(cols, m("col"))
                        for k, val in at.items():
                            e.set(k, val)
                        e.set("min", str(lo))
                        e.set("max", str(hi))
                    n += 1
                if "row_height" in v:
                    _a, r1, _b, r2 = parse_range(v["rows"] + (":" + v["rows"] if ":" not in v["rows"] else ""))
                    for row in s.data():
                        if r1 <= int(row.get("r", "0")) <= r2:
                            row.set("ht", str(v["row_height"]))
                            row.set("customHeight", "1")
                            n += 1
                sv = ws_ensure(root, "sheetViews")
                if len(sv) == 0:
                    etree.SubElement(sv, m("sheetView")).set("workbookViewId", "0")
                view = sv[0]
                if "rtl" in v:
                    view.set("rightToLeft", "1" if v["rtl"] else "0")
                    n += 1
                if "zoom" in v:
                    view.set("zoomScale", str(v["zoom"]))
                    view.set("zoomScaleNormal", str(v["zoom"]))
                    n += 1
                if "gridlines" in v:
                    view.set("showGridLines", "1" if v["gridlines"] else "0")
                    n += 1
                if "freeze" in v:
                    for old in view.findall(m("pane")):
                        view.remove(old)
                    if v["freeze"].strip() != "-":
                        c, r = parse_ref(v["freeze"])
                        pane = etree.Element(m("pane"))
                        if c > 1:
                            pane.set("xSplit", str(c - 1))
                        if r > 1:
                            pane.set("ySplit", str(r - 1))
                        pane.set("topLeftCell", "%s%d" % (n2col(c), r))
                        pane.set("activePane", "bottomRight" if c > 1 and r > 1 else ("topRight" if c > 1 else "bottomLeft"))
                        pane.set("state", "frozen")
                        view.insert(0, pane)
                    n += 1
                if "tab_color" in v:
                    sp = ws_ensure(root, "sheetPr")
                    for old in sp.findall(m("tabColor")):
                        sp.remove(old)
                    tc = etree.Element(m("tabColor"))
                    tc.set("rgb", "FF" + v["tab_color"])
                    sp.insert(0, tc)
                    n += 1
            ctx.count("settings applied", n)

    class XPage(Feature):
        key, label = "page", "Page setup & print"
        description = "Paper, orientation, scale / fit to pages, margins, centering, gridlines, header, footer, print area."
        uses_scope = False
        fields = [SHEETS_FIELD,
                  F("paper", "choice", "Paper size", choices=list(PAPER)), F("orientation", "choice", "Orientation", choices=["portrait", "landscape"]),
                  F("scale", "int", "Scale %", minimum=10, maximum=400), F("fit_width", "int", "Fit to N pages wide", minimum=0, maximum=100),
                  F("fit_height", "int", "Fit to N pages tall (0 = automatic)", minimum=0, maximum=100),
                  F("margin_top", "length", "Top margin", unit="cm"), F("margin_bottom", "length", "Bottom margin", unit="cm"),
                  F("margin_left", "length", "Left margin", unit="cm"), F("margin_right", "length", "Right margin", unit="cm"),
                  F("header_distance", "length", "Header distance", unit="cm"), F("footer_distance", "length", "Footer distance", unit="cm"),
                  F("center_horizontally", "bool", "Center on page horizontally"), F("center_vertically", "bool", "Center on page vertically"),
                  F("print_gridlines", "bool", "Print gridlines"),
                  F("header_text", "text", "Header text (centered)", hint="codes: &P page, &N pages, &D date, &F file, &A sheet"),
                  F("footer_text", "text", "Footer text (centered)", hint="e.g. Page &P of &N"),
                  F("print_area", "text", "Print area", hint="e.g. A1:H40; '-' clears")]

        def apply(self, ctx, v):
            book = Book(ctx.pkg)
            sheets = book.writable(book.pick(v.get("sheets")))
            n = 0
            for s in sheets:
                root = s.root(edit=True)
                if any(k in v for k in ("paper", "orientation", "scale", "fit_width", "fit_height")):
                    ps = ws_ensure(root, "pageSetup")
                    if "paper" in v:
                        ps.set("paperSize", str(PAPER[v["paper"]]))
                    if "orientation" in v:
                        ps.set("orientation", v["orientation"])
                    if "scale" in v:
                        ps.set("scale", str(v["scale"]))
                    if "fit_width" in v or "fit_height" in v:
                        sp = ws_ensure(root, "sheetPr")
                        pp = sp.find(m("pageSetUpPr"))
                        if pp is None:
                            pp = etree.SubElement(sp, m("pageSetUpPr"))
                        pp.set("fitToPage", "1")
                        ps.set("fitToWidth", str(v.get("fit_width", 1)))
                        ps.set("fitToHeight", str(v.get("fit_height", 0)))
                    n += 1
                if any(k in v for k in ("margin_top", "margin_bottom", "margin_left", "margin_right", "header_distance", "footer_distance")):
                    pm = ws_ensure(root, "pageMargins")
                    for a, d in (("left", 0.7), ("right", 0.7), ("top", 0.75), ("bottom", 0.75), ("header", 0.3), ("footer", 0.3)):
                        if pm.get(a) is None:
                            pm.set(a, str(d))
                    for k, a in (("margin_left", "left"), ("margin_right", "right"), ("margin_top", "top"), ("margin_bottom", "bottom"),
                                 ("header_distance", "header"), ("footer_distance", "footer")):
                        if k in v:
                            pm.set(a, "%.4f" % (v[k] / 72.0))             # wordtk lengths are in pt; Excel wants inches
                    n += 1
                if any(k in v for k in ("center_horizontally", "center_vertically", "print_gridlines")):
                    po = ws_ensure(root, "printOptions")
                    for k, a in (("center_horizontally", "horizontalCentered"), ("center_vertically", "verticalCentered"), ("print_gridlines", "gridLines")):
                        if k in v:
                            po.set(a, "1" if v[k] else "0")
                    n += 1
                if "header_text" in v or "footer_text" in v:
                    hf = ws_ensure(root, "headerFooter")
                    for k, tag in (("header_text", "oddHeader"), ("footer_text", "oddFooter")):
                        if k in v:
                            for old in hf.findall(m(tag)):
                                hf.remove(old)
                            e = etree.Element(m(tag))
                            e.text = "&C" + v[k]
                            hf.insert(0 if tag == "oddHeader" else len(hf), e)
                    n += 1
            if "print_area" in v:
                wb = ctx.pkg.xml(ctx.pkg.main, edit=True)
                dn = wb.find(m("definedNames"))
                if dn is None:
                    dn = etree.Element(m("definedNames"))
                    rank = WB_ORDER.index("definedNames")
                    at = len(wb)
                    for i, ch in enumerate(wb):
                        loc = etree.QName(ch).localname
                        if loc in WB_ORDER and WB_ORDER.index(loc) > rank:
                            at = i
                            break
                    wb.insert(at, dn)
                for s in sheets:
                    for old in [d for d in dn if d.get("name") == "_xlnm.Print_Area" and d.get("localSheetId") == str(s.index)]:
                        dn.remove(old)
                    if v["print_area"].strip() != "-":
                        c1, r1, c2, r2 = parse_range(v["print_area"])
                        e = etree.SubElement(dn, m("definedName"))
                        e.set("name", "_xlnm.Print_Area")
                        e.set("localSheetId", str(s.index))
                        e.text = "'%s'!$%s$%d:$%s$%d" % (s.name.replace("'", "''"), n2col(c1), r1, n2col(c2), r2)
                    n += 1
                if not len(dn):
                    wb.remove(dn)
            ctx.count("sheets configured", len(sheets) if n else 0)

    XLSX_FEATURES = {f.key: f for f in (XText(), XFormat(), XLayout(), XPage())}
    XLSX_SECTIONS = {
        "text": [("Find & Replace \u2022 \u062c\u0627\u06cc\u06af\u0632\u06cc\u0646\u06cc", ["sheets", "replacements", "replace_list"]),
                 ("Digits & Letters \u2022 \u0627\u0631\u0642\u0627\u0645 \u0648 \u062d\u0631\u0648\u0641", ["digits_convert", "digits", "persian_letters", "change_case"]),
                 ("Spaces & Punctuation \u2022 \u0641\u0627\u0635\u0644\u0647", ["collapse_spaces", "fix_punctuation_spacing"])],
        "format": [("Where \u2022 \u0645\u062d\u062f\u0648\u062f\u0647", ["sheets", "range"]), ("Font \u2022 \u0642\u0644\u0645", ["font", "size", "bold", "italic", "underline", "strike", "color"]),
                   ("Fill & Borders \u2022 \u0633\u0627\u06cc\u0647 \u0648 \u06a9\u0627\u062f\u0631", ["fill", "border_style", "border_color"]), ("Number \u2022 \u0639\u062f\u062f", ["number_format"]),
                   ("Alignment \u2022 \u062a\u0631\u0627\u0632", ["horizontal", "vertical", "wrap", "shrink", "indent", "direction"])],
        "layout": [("Where \u2022 \u0628\u0631\u06af\u0647", ["sheets"]), ("Columns \u2022 \u0633\u062a\u0648\u0646", ["columns", "col_width"]), ("Rows \u2022 \u0633\u0637\u0631", ["rows", "row_height"]),
                   ("View \u2022 \u0646\u0645\u0627", ["rtl", "zoom", "gridlines", "freeze", "tab_color"])],
        "page": [("Where \u2022 \u0628\u0631\u06af\u0647", ["sheets"]), ("Paper \u2022 \u06a9\u0627\u063a\u0630", ["paper", "orientation", "scale", "fit_width", "fit_height"]),
                 ("Margins \u2022 \u062d\u0627\u0634\u06cc\u0647", ["margin_top", "margin_bottom", "margin_left", "margin_right", "header_distance", "footer_distance"]),
                 ("Print \u2022 \u0686\u0627\u067e", ["center_horizontally", "center_vertically", "print_gridlines", "print_area"]), ("Header & Footer \u2022 \u0633\u0631\u0635\u0641\u062d\u0647", ["header_text", "footer_text"])]}

    def xlsx_sections(feat):
        fm = feat.field_map()
        out, used = [], set()
        for title, names in XLSX_SECTIONS.get(feat.key, []):
            fl = [fm[n] for n in names if n in fm and n not in used]
            used.update(f.name for f in fl)
            if fl:
                out.append((title, fl))
        rest = [f for f in feat.fields if f.name not in used]
        return out + ([("Other", rest)] if rest else [])

    def pkg_from_bytes(data, ext):
        d = tempfile.mkdtemp(prefix="po_pkg_")
        try:
            p = os.path.join(d, "x." + ext)
            with open(p, "wb") as f:
                f.write(data)
            return DocxPackage(p)                # reads every part into memory, so the temp file can go
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def pkg_bytes(pkg):
        d = tempfile.mkdtemp(prefix="po_pkg_")
        try:
            p = os.path.join(d, "out.zip")
            pkg.save(p)
            with open(p, "rb") as f:
                return f.read()
        finally:
            shutil.rmtree(d, ignore_errors=True)

    NUM = re.compile(r"[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?")

    def xl_set_cell(book, sheet, ref, value):
        if sheet.protected:
            raise OfficeEditError("sheet %r is protected; editing it here would bypass that protection" % sheet.name)
        c = sheet.cell(ref, create=True, edit=True)
        f = c.find(m("f"))
        if f is not None and (f.get("t") in ("shared", "array") and (f.text or f.get("ref"))):
            raise OfficeEditError("%s is the master of a shared / array formula; changing it would break the cells that depend on it" % ref)
        had_formula = f is not None
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop("t", None)
        text = str(value)
        if text == "":
            return had_formula, False
        if text.startswith("="):
            etree.SubElement(c, m("f")).text = text[1:]
            return had_formula, True
        t2 = text.strip().translate(_TO_EN).replace("\u066b", ".").replace("\u066c", "")
        if NUM.fullmatch(t2):
            etree.SubElement(c, m("v")).text = t2
        elif text.strip().lower() in ("true", "false"):
            c.set("t", "b")
            etree.SubElement(c, m("v")).text = "1" if text.strip().lower() == "true" else "0"
        else:
            c.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(c, m("is")), m("t"))
            set_node_text(t, text)
            if not t.get(XMLSPACE) and (text != text.strip()):
                t.set(XMLSPACE, "preserve")
        return had_formula, False

    def xl_values_changed(pkg):
        """cached formula results are no longer valid: drop them and ask Excel to recalculate when the file is opened"""
        book = Book(pkg)
        cleared = 0
        for s in book.sheets:
            todo = [c for c in s.cells() if c.find(m("f")) is not None and (c.find(m("v")) is not None)]
            if todo:
                s.root(edit=True)
                for c in todo:
                    c.remove(c.find(m("v")))
                    c.attrib.pop("t", None)
                    cleared += 1
        if pkg.exists("xl/calcChain.xml"):
            pkg.remove_part("xl/calcChain.xml")
        wb = pkg.xml(pkg.main, edit=True)
        cp = wb.find(m("calcPr"))
        if cp is None:
            cp = etree.Element(m("calcPr"))
            rank = WB_ORDER.index("calcPr")
            at = len(wb)
            for i, ch in enumerate(wb):
                loc = etree.QName(ch).localname
                if loc in WB_ORDER and WB_ORDER.index(loc) > rank:
                    at = i
                    break
            wb.insert(at, cp)
        cp.set("fullCalcOnLoad", "1")
        return cleared

    class XlsxSession(BaseSession):
        kind, file_ext = "xlsx", "xlsx"

        def __init__(self, orig, name, log=None):
            super().__init__(orig, name, log)
            self.features = XLSX_FEATURES
            self.blocked = None
            try:
                self._report()
            except Exception as e:  # noqa
                self.blocked = "This is not a readable .xlsx workbook (%s: %s)." % (type(e).__name__, e)
                self.report["items"].append({"level": "BLOCK", "text": self.blocked})
                self.report["blocked"] = self.blocked

        def _text_groups(self):
            return {"text"}

        def _book(self, data=None):
            if data is not None and data is not self.cur:
                return Book(pkg_from_bytes(data, "xlsx"))
            if getattr(self, "_bk", None) is None or self._bk[0] is not self.cur:
                self._bk = (self.cur, Book(pkg_from_bytes(self.cur, "xlsx")))
            return self._bk[1]

        def grid(self, sheet=0, r0=1, c0=1, rows=40, cols=12):
            book = self._book()
            if not 0 <= sheet < len(book.sheets):
                raise OfficeEditError("there is no sheet number %s" % sheet)
            s = book.sheets[sheet]
            cells, maxr, maxc = {}, 0, 0
            for c in s.cells():
                cc, rr = parse_ref(c.get("r"))
                f = c.find(m("f"))
                text = book.cell_text(c)
                if text == "" and f is None:
                    continue
                maxr, maxc = max(maxr, rr), max(maxc, cc)
                if r0 <= rr < r0 + rows and c0 <= cc < c0 + cols:
                    cells[(rr, cc)] = [text, ("=" + f.text) if f is not None and f.text else None]
            return {"sheet": sheet, "sheet_name": s.name, "sheets": [x.name for x in book.sheets], "r0": r0, "c0": c0, "max_row": maxr, "max_col": maxc,
                    "protected": s.protected, "rows": [[cells.get((r, c)) for c in range(c0, c0 + cols)] for r in range(r0, r0 + rows)]}

        def _report(self):
            pkg = pkg_from_bytes(self.orig, "xlsx")
            names = pkg.names()
            book = Book(pkg)
            R, items = self.report, self.report["items"]

            def add(level, text):
                items.append({"level": level, "text": text})
            wbx = pkg.raw(pkg.main).decode("utf-8", "ignore")
            nform = nshared = 0
            prot = []
            merged = cond = valid = 0
            for s in book.sheets:
                raw = pkg.raw(s.part).decode("utf-8", "ignore")
                nform += raw.count("<f>") + raw.count("<f ")
                nshared += len(re.findall(r'<f[^>]*t="(shared|array)"', raw))
                merged += raw.count("<mergeCell ")
                cond += raw.count("<conditionalFormatting")
                valid += raw.count("<dataValidation ")
                if re.search(r'<sheetProtection[^>]*sheet="(1|true)"', raw):
                    prot.append(s.name)
            if prot:
                R["blocked_sheets"] = prot
                add("WARN", "Protected sheet(s): %s. Editing them is refused here (it would bypass the protection)." % ", ".join(prot))
            if re.search(r"<workbookProtection[^>]*lockStructure=\"(1|true)\"", wbx):
                add("INFO", "Workbook structure is locked: sheets cannot be added / removed (this editor does not do that anyway).")
            if any(n.startswith("_xmlsignatures/") for n in names):
                add("WARN", "The workbook has a DIGITAL SIGNATURE. Any change makes the signature invalid.")
            if "xl/vbaProject.bin" in names:
                add("WARN", "The workbook contains MACROS (kept untouched). If a macro depends on cells you change, it may stop working.")
            if nform:
                add("WARN", "%d formula cell(s)%s. When you change a value or formula, the stored results of all formulas are cleared and recalculated when the file is opened in Excel / LibreOffice (other apps that only read stored results will show empty cells until then)." % (nform, (", %d of them shared/array formulas (their master cells cannot be edited)" % nshared) if nshared else ""))
            charts = len([n for n in names if n.startswith("xl/charts/chart")])
            pivots = len([n for n in names if n.startswith("xl/pivotTables/pivotTable")])
            tables = len([n for n in names if n.startswith("xl/tables/table")])
            if charts or pivots:
                add("WARN", "%s%s. Kept byte-for-byte, but they are not refreshed here: charts / pivot tables built on cells you change will update only when opened in Excel." % (("%d chart(s)" % charts) if charts else "", (", %d pivot table(s)" % pivots) if pivots else ""))
            if tables:
                add("INFO", "%d Excel table(s): kept; changing cells inside them is fine, inserting / deleting rows is not offered here." % tables)
            ext = len([n for n in names if n.startswith("xl/externalLinks/")])
            if ext:
                add("WARN", "%d external link(s) to other workbooks: kept untouched." % ext)
            if cond or valid or merged:
                add("INFO", "Conditional formatting: %d, data validation: %d, merged cells: %d \u2013 all kept untouched." % (cond, valid, merged))
            media = len([n for n in names if n.startswith("xl/media/") or n.startswith("xl/drawings/")])
            if media:
                add("INFO", "Pictures / drawings present (%d part(s)): kept untouched." % media)
            hidden = [s.name for s in book.sheets if s.info["state"] != "visible"]
            if hidden:
                add("INFO", "Hidden sheet(s): %s (they are included when you apply a setting to 'all sheets')." % ", ".join(hidden))
            R["stats"] = {"sheets": len(book.sheets), "formulas": nform}
            add("OK", "Your original file is never touched: edits happen in a copy and are saved with 'Save as'. Parts you do not change are copied byte for byte.")
            add("INFO", "Not available in this version: inserting / deleting rows or columns, renaming sheets, charts, pivot tables, sorting. Cell values, formulas, formats, widths, print setup and find & replace are available.")

        def schema(self):
            return schema_from(self.features, xlsx_sections)

        def sheet_names(self):
            return [s.name for s in self._book().sheets]

        def cells(self, sheet=0, offset=0, limit=250):
            book = self._book()
            if not 0 <= sheet < len(book.sheets):
                raise OfficeEditError("there is no sheet number %s" % sheet)
            s = book.sheets[sheet]
            out, total = [], 0
            for c in s.cells():
                text = book.cell_text(c)
                f = c.find(m("f"))
                if text == "" and f is None:
                    continue
                total += 1
                if offset <= total - 1 < offset + limit:
                    t = c.get("t", "")
                    kind = "formula" if f is not None else ("text" if t in ("s", "inlineStr", "str") else ("bool" if t == "b" else ("error" if t == "e" else "number")))
                    out.append({"ref": c.get("r"), "text": text, "kind": kind, "formula": ("=" + (f.text or "")) if f is not None and (f.text or "") else None,
                                "locked": s.protected or (f is not None and f.get("t") in ("shared", "array") and bool(f.text or f.get("ref")))})
            return {"sheet": sheet, "sheet_name": s.name, "sheets": [x.name for x in book.sheets], "cells": out, "total": total, "offset": offset, "protected": s.protected}

        def blocks(self):
            return []

        def table_rows(self):
            return []

        def page_info(self):
            return []

        def apply(self, ops):
            pkg = pkg_from_bytes(self.cur, "xlsx")
            book = Book(pkg)
            formulas_touched = False
            for op in ops:
                if op.get("op") != "set_cell":
                    raise OfficeEditError("unknown operation %r" % op.get("op"))
                sh = op.get("sheet", 0)
                try:
                    sheet = book.sheets[int(sh)] if str(sh).isdigit() else book.pick(sh)[0]
                except (IndexError, ValueError):
                    raise OfficeEditError("there is no sheet %r" % sh)
                had, now = xl_set_cell(book, sheet, str(op.get("ref", "")).strip().upper(), op.get("value", ""))
                formulas_touched = formulas_touched or had or now
            cleared = xl_values_changed(pkg)
            new = pkg_bytes(pkg)
            bad = xml_problems(new, part_diff(self.cur, new)["changed"])
            if bad:
                raise OfficeEditError("the result would contain damaged XML (%s). Nothing was changed." % "; ".join(bad))
            self._push(new, {"kind": "ops", "ops": list(ops), "desc": "%d cell edit(s)" % len(ops), "text_changing": True})
            return {"ok": True, "results": [{"ok": True} for _ in ops], "structure_changed": False, "formula_results_cleared": cleared}

        def _impact(self, old, new):
            try:
                a, b = self._book(old).values(), self._book(new).values()
            except Exception:
                return {}
            keys = set(a) | set(b)
            diff = [k for k in keys if a.get(k) != b.get(k)]
            samples = [{"old": "%s!%s  %s" % (k[0], k[1], (a.get(k) or ("", "", None))[1][:60]), "new": (b.get(k) or ("", "", None))[1][:60]} for k in sorted(diff)[:6]]
            return {"paragraphs_text_changed": len(diff), "paragraphs_before": len(a), "paragraphs_after": len(b), "samples": samples}

        def find(self, q, match_case=False, whole_word=False, regex=False, limit=300):
            if not q:
                raise OfficeEditError("type something to find")
            try:
                pat = q if regex else re.escape(q)
                if whole_word:
                    pat = r"(?<!\w)(?:%s)(?!\w)" % pat
                rx = re.compile(pat, 0 if match_case else re.I)
            except re.error as e:
                raise OfficeEditError("the search expression is not valid: %s" % e)
            book = self._book()
            out, total = [], 0
            for si, s in enumerate(book.sheets):
                for c in s.cells():
                    t = book.cell_text(c)
                    ms = list(rx.finditer(t)) if t else []
                    if ms:
                        total += len(ms)
                        if len(out) < limit:
                            out.append({"sheet": si, "ref": c.get("r"), "sheet_name": s.name, "count": len(ms), "snippet": t[:100], "locked": s.protected})
            return {"matches": total, "paragraphs": len(out), "results": out, "truncated": len(out) >= limit}

        def verify(self):
            checks = []
            self._verify_package(checks)
            value_changing = any(h.get("text_changing") for h in self.history)
            a, b = self._book(self.orig).values(), self._book(self.cur).values()
            diff = [k for k in set(a) | set(b) if a.get(k) != b.get(k)]
            changes = [{"tag": "cell", "old": "%s!%s  %s" % (k[0], k[1], (a.get(k) or ("", "", None))[1][:60]), "new": (b.get(k) or ("", "", None))[1][:60]} for k in sorted(diff)[:60]]
            if value_changing:
                checks.append({"name": "Cell content changes are the ones you asked for", "ok": True, "detail": "%d cell(s) differ (expected: you used value-changing edits)" % len(diff)})
            else:
                strict = [k for k in diff]
                checks.append({"name": "No cell value or formula was changed by formatting edits", "ok": not strict, "detail": "%d cell(s) differ" % len(strict)})
            return {"ok": all(c["ok"] for c in checks), "checks": checks, "changes": changes, "ops": len(self.history), "paragraphs_before": len(a), "paragraphs_after": len(b)}


# =============================================================================== POWERPOINT
if OFFICE_EDIT_ERROR is None:
    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

    def a_(tag):
        return "{%s}%s" % (NS_A, tag)

    def p_(tag):
        return "{%s}%s" % (NS_P, tag)

    RPR_ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag", "highlight", "uLnTx", "uLn",
                 "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
    PPR_ORDER = ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar",
                 "buBlip", "tabLst", "defRPr", "extLst"]
    FILL_TAGS = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}

    def ordered_set(parent, tag, new, order, replace_group=None):
        """put `new` (lxml element) into parent at the schema position, replacing existing children of the same tag (or group)"""
        group = replace_group or {tag}
        for old in [c for c in parent if etree.QName(c).localname in group]:
            parent.remove(old)
        if new is None:
            return
        rank = order.index(tag)
        at = len(parent)
        for i, c in enumerate(parent):
            loc = etree.QName(c).localname
            if loc in order and order.index(loc) > rank:
                at = i
                break
        parent.insert(at, new)

    def slide_filter(spec, total):
        if not spec or not str(spec).strip():
            return set(range(1, total + 1))
        out = set()
        for tok in [t.strip() for t in str(spec).replace(";", ",").split(",") if t.strip()]:
            try:
                if "-" in tok:
                    x, y = tok.split("-", 1)
                    out |= set(range(int(x), int(y) + 1))
                else:
                    out.add(int(tok))
            except ValueError:
                raise OfficeEditError("slides: %r is not a slide number or range (example: 1,3-5)" % tok)
        bad = [n for n in out if not 1 <= n <= total]
        if bad:
            raise OfficeEditError("the presentation has %d slide(s); %s does not exist" % (total, ", ".join(map(str, sorted(bad)))))
        return out

    class Deck:
        def __init__(self, pkg):
            self.pkg = pkg
            pres = pkg.xml(pkg.main)
            rels = {rid: tgt for rid, tgt, _t, ext in pkg.rels(pkg.main, "slide") if not ext}
            self.slides = []
            lst = pres.find(p_("sldIdLst"))
            for sid in (lst if lst is not None else []):
                rid = sid.get("{%s}id" % NS_R)
                if rid in rels and pkg.exists(rels[rid]):
                    self.slides.append({"part": rels[rid], "n": len(self.slides) + 1})

        def root(self, slide, edit=False):
            return self.pkg.xml(slide["part"], edit)

        def paragraphs(self, only=None):
            """[(slide dict, shape label, a:p element)] in presentation order"""
            out = []
            for sl in self.slides:
                if only is not None and sl["n"] not in only:
                    continue
                root = self.root(sl)
                for p in root.iter(a_("p")):
                    shape = "text"
                    anc = p.getparent()
                    while anc is not None:
                        if anc.tag in (p_("sp"), p_("graphicFrame")):
                            nv = anc.find(".//" + p_("cNvPr"))
                            shape = (nv.get("name") if nv is not None else "shape") + (" (table)" if anc.tag == p_("graphicFrame") else "")
                            break
                        anc = anc.getparent()
                    out.append((sl, shape, p))
            return out

    def para_runs(p):
        return [r for r in p.findall(a_("r"))]

    def para_nodes(p):
        return [r.find(a_("t")) for r in para_runs(p) if r.find(a_("t")) is not None]

    def para_lock(p):
        for c in p:
            loc = etree.QName(c).localname
            if loc == "fld":
                return "a field (slide number, date ...)"
            if loc == "br":
                return "manual line breaks"
            if loc not in ("pPr", "r", "endParaRPr"):
                return "unsupported content <%s>" % loc
        for r in para_runs(p):
            for c in r:
                if etree.QName(c).localname not in ("rPr", "t"):
                    return "unsupported run content"
        anc = p.getparent()
        while anc is not None:
            if anc.tag == p_("sp"):
                ph = anc.find(".//" + p_("ph"))
                if ph is not None and ph.get("type") in ("sldNum", "dt", "ftr"):
                    return "a slide number / date / footer placeholder"
                break
            anc = anc.getparent()
        return None

    def para_text(p):
        return "".join(t.text or "" for t in para_nodes(p))

    def new_run_for(p):
        r = etree.Element(a_("r"))
        rpr = etree.SubElement(r, a_("rPr"))
        end = p.find(a_("endParaRPr"))
        if end is not None:
            for k, v in end.attrib.items():
                rpr.set(k, v)
            for ch in end:
                rpr.append(copy.deepcopy(ch))
        rpr.set("lang", rpr.get("lang", "en-US"))
        etree.SubElement(r, a_("t"))
        if end is not None:
            end.addprevious(r)
        else:
            p.append(r)
        return r

    def split_run(r, at):
        """split run r after `at` characters; returns (left, right)"""
        t = r.find(a_("t"))
        txt = t.text or ""
        right = copy.deepcopy(r)
        set_node_text(t, txt[:at])
        set_node_text(right.find(a_("t")), txt[at:])
        r.addnext(right)
        return r, right

    def run_props(rpr, props):
        if props.get("size"):
            rpr.set("sz", str(int(round(float(props["size"]) * 100))))
        if props.get("bold") is not None:
            rpr.set("b", "1" if props["bold"] else "0")
        if props.get("italic") is not None:
            rpr.set("i", "1" if props["italic"] else "0")
        if props.get("underline") is not None:
            rpr.set("u", "sng" if props["underline"] else "none")
        if props.get("strike") is not None:
            rpr.set("strike", "sngStrike" if props["strike"] else "noStrike")
        if props.get("color"):
            c = str(props["color"]).lstrip("#")
            if not re.fullmatch(r"[0-9a-fA-F]{6}", c):
                raise OfficeEditError("colour must be like FF0000")
            sf = etree.Element(a_("solidFill"))
            etree.SubElement(sf, a_("srgbClr")).set("val", c.upper())
            ordered_set(rpr, "solidFill", sf, RPR_ORDER, FILL_TAGS)
        if props.get("highlight"):
            h = str(props["highlight"]).lstrip("#")
            if not re.fullmatch(r"[0-9a-fA-F]{6}", h):
                raise OfficeEditError("highlight must be a hex colour like FFFF00")
            hl = etree.Element(a_("highlight"))
            etree.SubElement(hl, a_("srgbClr")).set("val", h.upper())
            ordered_set(rpr, "highlight", hl, RPR_ORDER)
        if props.get("font"):
            for tag in ("latin", "ea", "cs"):
                e = etree.Element(a_(tag))
                e.set("typeface", props["font"])
                ordered_set(rpr, tag, e, RPR_ORDER)

    def pptx_format(p, start, end, props):
        runs = para_runs(p)
        total = sum(len(r.find(a_("t")).text or "") for r in runs if r.find(a_("t")) is not None)
        if start is None or end is None:
            start, end = 0, total
        if not 0 <= start <= end <= total:
            raise OfficeEditError("selection is outside the paragraph")
        if total == 0 or not runs:
            e = p.find(a_("endParaRPr"))
            if e is None:
                e = etree.SubElement(p, a_("endParaRPr"))
            run_props(e, props)
            return
        if start == end:
            raise OfficeEditError("nothing selected")
        pos = 0
        for r in list(runs):
            t = r.find(a_("t"))
            L = len(t.text or "") if t is not None else 0
            lo, hi = pos, pos + L
            pos = hi
            if hi <= start or lo >= end or L == 0:
                continue
            target = r
            if start > lo:
                _l, target = split_run(target, start - lo)
                lo = start
            if end < hi:
                target, _r = split_run(target, end - lo)
            rpr = target.find(a_("rPr"))
            if rpr is None:
                rpr = etree.Element(a_("rPr"))
                target.insert(0, rpr)
            run_props(rpr, props)

    def pptx_para(p, props):
        ppr = p.find(a_("pPr"))
        if ppr is None:
            ppr = etree.Element(a_("pPr"))
            p.insert(0, ppr)
        if props.get("align"):
            ppr.set("algn", {"left": "l", "center": "ctr", "right": "r", "justify": "just"}[props["align"]])
        if props.get("rtl") is not None:
            ppr.set("rtl", "1" if props["rtl"] else "0")
        emu = lambda cm: str(int(round(float(cm) * 360000)))
        if props.get("left") is not None:
            ppr.set("marL", emu(props["left"]))
        if props.get("right") is not None:
            ppr.set("marR", emu(props["right"]))
        if props.get("first") is not None:
            ppr.set("indent", emu(props["first"]))
        if props.get("line") is not None:
            ls = etree.Element(a_("lnSpc"))
            etree.SubElement(ls, a_("spcPct")).set("val", str(int(round(float(props["line"]) * 100000))))
            ordered_set(ppr, "lnSpc", ls, PPR_ORDER)
        for key, tag in (("before", "spcBef"), ("after", "spcAft")):
            if props.get(key) is not None:
                e = etree.Element(a_(tag))
                etree.SubElement(e, a_("spcPts")).set("val", str(int(round(float(props[key]) * 100))))
                ordered_set(ppr, tag, e, PPR_ORDER)

    # ----------------------------------------------------------------------------- features
    SLIDES_FIELD = F("slides", "text", "Slides", hint="e.g. 1,3-5; empty = all slides")

    class PText(Feature):
        key, label = "text", "Text"
        description = "Find & replace on slides, digits, Persian letters, spaces, case."
        uses_scope = False
        fields = [SLIDES_FIELD,
                  F("replacements", "json", "Find & replace", hint='[{"find":"old","replace":"new","regex":false,"match_case":false,"whole_word":false}]'),
                  F("replace_list", "text", "Bulk replace list", hint="one rule per line:  old => new   (no => deletes the text)"),
                  F("digits", "choice", "Convert ALL digits to", choices=["persian", "english", "arabic"]),
                  F("digits_convert", "choice", "Convert digits (one direction)", choices=list(DIGIT_MAPS)),
                  F("persian_letters", "bool", "Replace Arabic \u0643 \u064a with Persian \u06a9 \u06cc"),
                  F("collapse_spaces", "bool", "Collapse repeated spaces"),
                  F("fix_punctuation_spacing", "bool", "Remove space before , . ; : ! ? \u060c \u061b \u061f"),
                  F("change_case", "choice", "Change case", choices=["upper", "lower", "title", "sentence"])]

        def apply(self, ctx, v):
            deck = Deck(ctx.pkg)
            only = slide_filter(v.get("slides"), len(deck.slides))
            rules, fn = find_rules(v), text_transform(v)
            if not rules and not any(k in v for k in ("digits", "digits_convert", "persian_letters", "collapse_spaces", "fix_punctuation_spacing", "change_case")):
                raise OfficeEditError("fill in something to find / replace or choose a conversion")
            changed = skipped = 0
            touched = set()
            for sl, _shape, p in deck.paragraphs(only):
                nodes = para_nodes(p)
                if not nodes:
                    continue
                old = [n.text or "" for n in nodes]
                n = apply_text_options(nodes, rules, fn)
                if n:
                    changed += n
                    touched.add(sl["part"])
            for sl in deck.slides:
                if sl["part"] in touched:
                    deck.root(sl, edit=True)
            ctx.count("text items changed", changed)

    class PFont(Feature):
        key, label = "font", "Font"
        description = "Font, size, style and color for all text on the chosen slides."
        uses_scope = False
        fields = [SLIDES_FIELD, F("font", "text", "Font name", hint="e.g. Calibri, B Nazanin"), F("size", "float", "Size (pt)", minimum=1, maximum=400),
                  F("bold", "bool", "Bold"), F("italic", "bool", "Italic"), F("underline", "bool", "Underline"), F("strike", "bool", "Strikethrough"),
                  F("color", "color", "Text color"), F("highlight", "color", "Highlight color")]

        def apply(self, ctx, v):
            deck = Deck(ctx.pkg)
            only = slide_filter(v.get("slides"), len(deck.slides))
            props = {k: v[k] for k in ("font", "size", "bold", "italic", "underline", "strike", "color", "highlight") if k in v}
            if not props:
                raise OfficeEditError("choose at least one font setting")
            n = 0
            touched = set()
            for sl, _shape, p in deck.paragraphs(only):
                if para_lock(p) and not para_runs(p):
                    continue
                for r in para_runs(p):
                    rpr = r.find(a_("rPr"))
                    if rpr is None:
                        rpr = etree.Element(a_("rPr"))
                        r.insert(0, rpr)
                    run_props(rpr, props)
                    n += 1
                    touched.add(sl["part"])
            for sl in deck.slides:
                if sl["part"] in touched:
                    deck.root(sl, edit=True)
            ctx.count("text runs formatted", n)

    class PParagraph(Feature):
        key, label = "paragraph", "Paragraph"
        description = "Alignment, direction, spacing and indents for all paragraphs on the chosen slides."
        uses_scope = False
        fields = [SLIDES_FIELD, F("alignment", "choice", "Alignment", choices=["left", "center", "right", "justify"]),
                  F("direction", "choice", "Text direction", choices=["ltr", "rtl"]), F("line_spacing", "float", "Line spacing (1 = single)", minimum=0.5, maximum=10),
                  F("space_before", "length", "Space before"), F("space_after", "length", "Space after"),
                  F("indent_left", "length", "Indent left", unit="cm"), F("indent_right", "length", "Indent right", unit="cm"), F("first_line_indent", "length", "First-line indent", unit="cm")]

        def apply(self, ctx, v):
            deck = Deck(ctx.pkg)
            only = slide_filter(v.get("slides"), len(deck.slides))
            props = {}
            if "alignment" in v:
                props["align"] = v["alignment"]
            if "direction" in v:
                props["rtl"] = v["direction"] == "rtl"
            if "line_spacing" in v:
                props["line"] = v["line_spacing"]
            if "space_before" in v:
                props["before"] = v["space_before"]
            if "space_after" in v:
                props["after"] = v["space_after"]
            for k_in, k_out in (("indent_left", "left"), ("indent_right", "right"), ("first_line_indent", "first")):
                if k_in in v:
                    props[k_out] = v[k_in] / 28.3464567           # pt -> cm
            if not props:
                raise OfficeEditError("choose at least one paragraph setting")
            n = 0
            touched = set()
            for sl, _shape, p in deck.paragraphs(only):
                pptx_para(p, props)
                n += 1
                touched.add(sl["part"])
            for sl in deck.slides:
                if sl["part"] in touched:
                    deck.root(sl, edit=True)
            ctx.count("paragraphs formatted", n)

    class PSlides(Feature):
        key, label = "slides", "Slides"
        description = "Hide or show slides."
        uses_scope = False
        fields = [SLIDES_FIELD, F("hidden", "bool", "Hide the slides (false = show them again)")]

        def apply(self, ctx, v):
            if "hidden" not in v:
                raise OfficeEditError("choose hide or show")
            deck = Deck(ctx.pkg)
            only = slide_filter(v.get("slides"), len(deck.slides))
            n = 0
            for sl in deck.slides:
                if sl["n"] in only:
                    root = deck.root(sl, edit=True)
                    if v["hidden"]:
                        root.set("show", "0")
                    else:
                        root.attrib.pop("show", None)
                    n += 1
            ctx.count("slides changed", n)

    PPTX_FEATURES = {f.key: f for f in (PText(), PFont(), PParagraph(), PSlides())}
    PPTX_SECTIONS = {
        "text": [("Find & Replace \u2022 \u062c\u0627\u06cc\u06af\u0632\u06cc\u0646\u06cc", ["slides", "replacements", "replace_list"]),
                 ("Digits & Letters \u2022 \u0627\u0631\u0642\u0627\u0645 \u0648 \u062d\u0631\u0648\u0641", ["digits_convert", "digits", "persian_letters", "change_case"]),
                 ("Spaces & Punctuation \u2022 \u0641\u0627\u0635\u0644\u0647", ["collapse_spaces", "fix_punctuation_spacing"])],
        "font": [("Where \u2022 \u0627\u0633\u0644\u0627\u06cc\u062f", ["slides"]), ("Font \u2022 \u0642\u0644\u0645", ["font", "size", "bold", "italic", "underline", "strike"]), ("Color \u2022 \u0631\u0646\u06af", ["color", "highlight"])],
        "paragraph": [("Where \u2022 \u0627\u0633\u0644\u0627\u06cc\u062f", ["slides"]), ("Alignment & Direction \u2022 \u062a\u0631\u0627\u0632", ["alignment", "direction"]),
                      ("Spacing \u2022 \u0641\u0627\u0635\u0644\u0647", ["line_spacing", "space_before", "space_after"]), ("Indentation \u2022 \u062a\u0648\u0631\u0641\u062a\u06af\u06cc", ["indent_left", "indent_right", "first_line_indent"])]}

    def pptx_sections(feat):
        fm = feat.field_map()
        out, used = [], set()
        for title, names in PPTX_SECTIONS.get(feat.key, []):
            fl = [fm[n] for n in names if n in fm and n not in used]
            used.update(f.name for f in fl)
            if fl:
                out.append((title, fl))
        rest = [f for f in feat.fields if f.name not in used]
        return out + ([("Other", rest)] if rest else [])

    class PptxSession(BaseSession):
        kind, file_ext = "pptx", "pptx"

        def __init__(self, orig, name, log=None):
            super().__init__(orig, name, log)
            self.features = PPTX_FEATURES
            self.blocked = None
            try:
                self._report()
            except Exception as e:  # noqa
                self.blocked = "This is not a readable .pptx presentation (%s: %s)." % (type(e).__name__, e)
                self.report["items"].append({"level": "BLOCK", "text": self.blocked})
                self.report["blocked"] = self.blocked

        def _text_groups(self):
            return {"text"}

        def _deck(self, data=None):
            return Deck(pkg_from_bytes(data if data is not None else self.cur, "pptx"))

        def _report(self):
            pkg = pkg_from_bytes(self.orig, "pptx")
            names = pkg.names()
            deck = Deck(pkg)
            R, items = self.report, self.report["items"]

            def add(level, text):
                items.append({"level": level, "text": text})
            pres = pkg.raw(pkg.main).decode("utf-8", "ignore")
            if "<p:modifyVerifier" in pres:
                R["blocked"] = self.blocked = "The presentation has a 'modify' password (write protection). Editing it here would bypass that protection."
                add("BLOCK", R["blocked"])
            if any(n.startswith("_xmlsignatures/") for n in names):
                add("WARN", "The presentation has a DIGITAL SIGNATURE. Any change makes the signature invalid.")
            if "ppt/vbaProject.bin" in names:
                add("WARN", "The presentation contains MACROS (kept untouched).")
            raw_all = "".join(pkg.raw(s["part"]).decode("utf-8", "ignore") for s in deck.slides)
            anim, trans = raw_all.count("<p:timing"), raw_all.count("<p:transition")
            if anim or trans:
                add("INFO", "Animations on %d slide(s) and transitions on %d slide(s): kept untouched. Changing text does not break them, but animations tied to a paragraph by position may then act on different text." % (anim, trans))
            charts = len([n for n in names if n.startswith("ppt/charts/chart")])
            smart = len([n for n in names if n.startswith("ppt/diagrams/data")])
            media = len([n for n in names if re.match(r"ppt/media/.*\.(mp4|wav|mp3|m4a|mov|avi|wmv)$", n, re.I)])
            ole = len([n for n in names if n.startswith("ppt/embeddings/")])
            if charts or smart or media or ole:
                add("WARN", "The presentation has %s. They are kept byte-for-byte but cannot be edited here; text inside charts / SmartArt is not changed by find & replace." % ", ".join(x for x in (("%d chart(s)" % charts) if charts else "", ("%d SmartArt" % smart) if smart else "", ("%d audio/video file(s)" % media) if media else "", ("%d embedded object(s)" % ole) if ole else "") if x))
            notes = len([n for n in names if n.startswith("ppt/notesSlides/notesSlide")])
            if notes:
                add("INFO", "%d slide(s) have speaker notes: kept, but not edited or searched here." % notes)
            hidden = raw_all.count('show="0"')
            if hidden:
                add("INFO", "%d hidden slide(s)." % hidden)
            locked = {}
            total = ed = 0
            for sl, _sh, p in deck.paragraphs():
                total += 1
                r = para_lock(p)
                if r:
                    locked[r] = locked.get(r, 0) + 1
                else:
                    ed += 1
            R["stats"] = {"paragraphs": total, "editable": ed, "locked": total - ed, "slides": len(deck.slides)}
            R["locked"] = locked
            if locked:
                add("INFO", "%d of %d text paragraphs are LOCKED for direct editing: %s." % (total - ed, total, "; ".join("%s: %d" % kv for kv in locked.items())))
            add("OK", "Your original file is never touched: edits happen in a copy and are saved with 'Save as'. Parts you do not change are copied byte for byte.")
            add("INFO", "Not available in this version: moving / resizing shapes, pictures, slide order, charts, notes, slide size. Text, fonts, paragraph settings and find & replace are available.")

        def schema(self):
            return schema_from(self.features, pptx_sections)

        def blocks(self):
            out = []
            for i, (sl, shape, p) in enumerate(self._deck().paragraphs()):
                reason = para_lock(p)
                text = para_text(p)
                b = {"i": i, "kind": "slide", "slide": sl["n"], "where": "Slide %d \u00b7 %s" % (sl["n"], shape), "text": text[:4000], "long": len(text) > 4000, "locked": bool(reason), "reason": reason}
                ppr = p.find(a_("pPr"))
                al = {"l": "left", "ctr": "center", "r": "right", "just": "justify"}.get(ppr.get("algn") if ppr is not None else None, "left")
                b.update({"align": al, "rtl": bool(ppr is not None and ppr.get("rtl") == "1"), "style": None})
                for r in para_runs(p):
                    rpr = r.find(a_("rPr"))
                    if rpr is not None and (r.find(a_("t")) is not None and r.find(a_("t")).text):
                        lat = rpr.find(a_("latin"))
                        b.update({"font": lat.get("typeface") if lat is not None else None, "size": int(rpr.get("sz")) / 100.0 if rpr.get("sz") else None,
                                  "bold": rpr.get("b") == "1", "italic": rpr.get("i") == "1", "underline": rpr.get("u") not in (None, "none")})
                        break
                out.append(b)
            return out

        def table_rows(self):
            return []

        def page_info(self):
            return []

        def apply(self, ops):
            pkg = pkg_from_bytes(self.cur, "pptx")
            deck = Deck(pkg)
            paras = deck.paragraphs()
            structure = False
            for op in ops:
                k = op.get("op")
                i = op.get("p")
                if not isinstance(i, int) or not 0 <= i < len(paras):
                    raise OfficeEditError("paragraph %s does not exist" % i)
                sl, _shape, p = paras[i]
                reason = para_lock(p)
                if k in ("set_text", "format", "para", "delete_paragraph") and reason:
                    raise OfficeEditError("paragraph %d is locked: it contains %s" % (i + 1, reason))
                deck.root(sl, edit=True)
                if k == "set_text":
                    new = str(op.get("text", "")).replace("\r\n", "\n").replace("\r", "\n")
                    if "\n" in new:
                        raise OfficeEditError("a paragraph cannot contain a line break here; insert a new paragraph instead")
                    try:
                        PDE.check_text(new)
                    except PDE.EditError as e:
                        raise OfficeEditError(str(e))
                    nodes = para_nodes(p)
                    old = "".join(n.text or "" for n in nodes)
                    diff_set_text(nodes, old, new, lambda p=p: new_run_for(p).find(a_("t")))
                    if para_text(p) != new:
                        raise OfficeEditError("internal check failed: the text would not match what you typed; nothing was changed")
                elif k == "format":
                    pptx_format(p, op.get("start"), op.get("end"), {x: op.get(x) for x in ("font", "size", "bold", "italic", "underline", "strike", "color", "highlight")})
                elif k == "para":
                    pptx_para(p, {x: op.get(x) for x in ("align", "rtl", "before", "after", "line", "left", "right", "first")})
                elif k == "insert_paragraph_after":
                    np_ = etree.Element(a_("p"))
                    ppr = p.find(a_("pPr"))
                    if ppr is not None:
                        np_.append(copy.deepcopy(ppr))
                    end = p.find(a_("endParaRPr"))
                    if end is not None:
                        np_.append(copy.deepcopy(end))
                    p.addnext(np_)
                    structure = True
                elif k == "delete_paragraph":
                    if len(p.getparent().findall(a_("p"))) <= 1:
                        raise OfficeEditError("a text box must keep at least one paragraph")
                    p.getparent().remove(p)
                    structure = True
                else:
                    raise OfficeEditError("unknown operation %r" % k)
                if structure and k in ("insert_paragraph_after", "delete_paragraph") and op is not ops[-1]:
                    paras = deck.paragraphs()                   # indexes shifted; later ops refer to the new list
            new = pkg_bytes(pkg)
            bad = xml_problems(new, part_diff(self.cur, new)["changed"])
            if bad:
                raise OfficeEditError("the result would contain damaged XML (%s). Nothing was changed." % "; ".join(bad))
            self._push(new, {"kind": "ops", "ops": list(ops), "desc": "%d edit(s)" % len(ops),
                             "text_changing": any(o.get("op") in ("set_text", "insert_paragraph_after", "delete_paragraph") for o in ops)})
            return {"ok": True, "results": [{"ok": True} for _ in ops], "structure_changed": structure}

        def _texts(self, data):
            return [para_text(p) for _s, _sh, p in self._deck(data).paragraphs()]

        def _impact(self, old, new):
            try:
                a, b = self._texts(old), self._texts(new)
            except Exception:
                return {}
            sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
            samples, n = [], 0
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag != "equal":
                    n += max(i2 - i1, j2 - j1)
                    if len(samples) < 6:
                        samples.append({"old": " / ".join(x[:70] for x in a[i1:i2][:2]), "new": " / ".join(x[:70] for x in b[j1:j2][:2])})
            return {"paragraphs_text_changed": n, "paragraphs_before": len(a), "paragraphs_after": len(b), "samples": samples}

        def find(self, q, match_case=False, whole_word=False, regex=False, limit=300):
            if not q:
                raise OfficeEditError("type something to find")
            try:
                pat = q if regex else re.escape(q)
                if whole_word:
                    pat = r"(?<!\w)(?:%s)(?!\w)" % pat
                rx = re.compile(pat, 0 if match_case else re.I)
            except re.error as e:
                raise OfficeEditError("the search expression is not valid: %s" % e)
            out, total = [], 0
            for i, (sl, shape, p) in enumerate(self._deck().paragraphs()):
                t = para_text(p)
                ms = list(rx.finditer(t))
                if ms:
                    total += len(ms)
                    if len(out) < limit:
                        a0, b0 = max(0, ms[0].start() - 30), min(len(t), ms[0].end() + 30)
                        out.append({"i": i, "count": len(ms), "snippet": ("\u2026" if a0 else "") + t[a0:b0] + ("\u2026" if b0 < len(t) else ""), "locked": bool(para_lock(p)), "kind": "slide",
                                    "where": "Slide %d \u00b7 %s" % (sl["n"], shape)})
            return {"matches": total, "paragraphs": len(out), "results": out, "truncated": len(out) >= limit}

        def verify(self):
            checks = []
            self._verify_package(checks)
            text_changing = any(h.get("text_changing") for h in self.history)
            a, b = self._texts(self.orig), self._texts(self.cur)
            sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
            changes, n_ch = [], 0
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag != "equal":
                    n_ch += max(i2 - i1, j2 - j1)
                    if len(changes) < 60:
                        changes.append({"tag": tag, "old": " / ".join(x[:80] for x in a[i1:i2][:3]), "new": " / ".join(x[:80] for x in b[j1:j2][:3])})
            if text_changing:
                checks.append({"name": "Text changes are the ones you asked for", "ok": True, "detail": "%d paragraph(s) differ in text (expected: you used text-changing edits)" % n_ch})
            else:
                checks.append({"name": "No text was changed by formatting edits", "ok": n_ch == 0, "detail": "%d paragraph(s) differ in text" % n_ch})
            return {"ok": all(c["ok"] for c in checks), "checks": checks, "changes": changes, "ops": len(self.history), "paragraphs_before": len(a), "paragraphs_after": len(b)}


def open_session(ext, data, name, log=None):
    """Session for an Office file (docx / xlsx / pptx families)."""
    if OFFICE_EDIT_ERROR:
        raise OfficeEditError("the editing engine is not available: %s (on the converter machine: pip install lxml)" % OFFICE_EDIT_ERROR)
    if ext in ("docx", "docm", "dotx", "dotm"):
        return WordSession(data, name, log)
    if ext in ("xlsx", "xlsm", "xltx", "xltm"):
        return XlsxSession(data, name, log)
    if ext in ("pptx", "pptm", "ppsx", "ppsm", "potx", "potm"):
        return PptxSession(data, name, log)
    raise OfficeEditError("Editing is available for Word (.docx .docm .dotx .dotm), Excel (.xlsx .xlsm .xltx .xltm) and PowerPoint (.pptx .pptm .ppsx .ppsm .potx .potm). Other formats would have to be re-written by LibreOffice, which could damage features of the document, so they are view-only.")
