# -*- coding: utf-8 -*-
"""Text group: find/replace, cleanup, digit & Persian normalisation, case, hyperlinks."""
import re

from ..core.ooxml import NS, get, local, qn
from ..core.scope import iter_paragraphs
from ..core.textops import (compile_find, has_visible_content, paragraph_runs, replace_in_paragraph,
                            run_text, set_run_text)
from ..core.units import _AR_DIGITS, _FA_DIGITS
from .base import F, Feature, ProfileError

_TO_FA = {ord(str(i)): c for i, c in enumerate(_FA_DIGITS)}
_TO_FA.update({ord(c): _FA_DIGITS[i] for i, c in enumerate(_AR_DIGITS)})
_TO_EN = {ord(c): str(i) for i, c in enumerate(_FA_DIGITS)}
_TO_EN.update({ord(c): str(i) for i, c in enumerate(_AR_DIGITS)})
_TO_AR = {ord(str(i)): c for i, c in enumerate(_AR_DIGITS)}
_TO_AR.update({ord(c): _AR_DIGITS[i] for i, c in enumerate(_FA_DIGITS)})


def _case(fn):
    return lambda m: fn(m.group(0))


class TextFeature(Feature):
    key = "text"
    label = "Text"
    description = "Find & replace (per-run formatting kept), cleanup, digits, Persian/Arabic letters, case."
    fields = [
        F("replacements", "json", "Find & replace",
          hint='[{"find":"old","replace":"new","regex":false,"match_case":false,"whole_word":false}]'),
        F("replace_list", "text", "Bulk replace list",
          hint="one rule per line:  old => new   (a line without => deletes the text)"),
        F("digits", "choice", "Convert ALL digits to", choices=["persian", "english", "arabic"]),
        F("digits_convert", "choice", "Convert digits (one direction)",
          choices=["english->persian", "persian->english", "arabic->persian", "persian->arabic",
                   "arabic->english", "english->arabic"],
          hint="only the source digit set is touched"),
        F("persian_letters", "bool", "Replace Arabic ك ي with Persian ک ی"),
        F("collapse_spaces", "bool", "Collapse repeated spaces"),
        F("trim_spaces", "bool", "Trim leading/trailing spaces in paragraphs"),
        F("fix_punctuation_spacing", "bool", "Remove space before , . ; : ! ? ، ؛ ؟"),
        F("remove_empty_paragraphs", "bool", "Remove empty lines (empty paragraphs)"),
        F("remove_empty_line_breaks", "bool", "Remove empty lines made of manual line breaks (Shift+Enter)"),
        F("remove_empty_pages", "bool", "Remove empty pages (page breaks with nothing on the page)"),
        F("max_consecutive_empty", "int", "Allow at most N consecutive empty paragraphs", minimum=0, maximum=20),
        F("change_case", "choice", "Change case", choices=["upper", "lower", "title", "sentence"]),
        F("remove_hyperlinks", "bool", "Unlink hyperlinks (keep text)"),
        F("remove_hidden_text", "bool", "Delete hidden text"),
    ]

    @staticmethod
    def parse_bulk(text):
        rules = []
        for line in str(text).splitlines():
            if not line.strip():
                continue
            if "=>" in line:
                a, b = line.split("=>", 1)
                a, b = a.strip(), b.strip()
            else:
                a, b = line.strip(), ""
            if a:
                rules.append({"find": a, "replace": b})
        return rules

    def validate(self, s):
        for r in s.get("replacements", []) or []:
            if not isinstance(r, dict) or "find" not in r or not r["find"]:
                raise ProfileError("text.replacements: each item needs a non-empty 'find'")
            if r.get("regex"):
                try:
                    re.compile(r["find"])
                except re.error as e:
                    raise ProfileError("text.replacements: bad regex %r: %s" % (r["find"], e))
        return s

    def _regex_pass(self, ctx, scope, rx, repl, key):
        n = 0
        for _n, _r, p in iter_paragraphs(ctx, scope, edit=True):
            n += replace_in_paragraph(p, rx, repl)
        ctx.count(key, n)

    def apply(self, ctx, s):
        scope = self.scope_for(ctx, s)
        rules = list(s.get("replacements", []) or []) + self.parse_bulk(s.get("replace_list", ""))
        for r in rules:
            rx = compile_find(r["find"], bool(r.get("regex")), bool(r.get("match_case")), bool(r.get("whole_word")))
            rep = r.get("replace", "")
            self._regex_pass(ctx, scope, rx, rep if r.get("regex") else (lambda m, rep=rep: rep), "text.replacements")
        if "digits_convert" in s:
            src, dst = s["digits_convert"].split("->")
            sets = {"english": "0-9", "persian": "۰-۹", "arabic": "٠-٩"}
            tab = {"persian": _TO_FA, "english": _TO_EN, "arabic": _TO_AR}[dst]
            self._regex_pass(ctx, scope, re.compile("[%s]+" % sets[src]),
                             lambda m, tab=tab: m.group(0).translate(tab), "text.digit_runs")
        if "digits" in s:
            table = {"persian": _TO_FA, "english": _TO_EN, "arabic": _TO_AR}[s["digits"]]
            self._regex_pass(ctx, scope, re.compile(r"[0-9۰-۹٠-٩]+"),
                             lambda m: m.group(0).translate(table), "text.digit_runs")
        if s.get("persian_letters"):
            self._regex_pass(ctx, scope, re.compile("[كيى]"),
                             lambda m: {"ك": "ک", "ي": "ی", "ى": "ی"}[m.group(0)], "text.letters_fixed")
        if s.get("collapse_spaces"):
            self._regex_pass(ctx, scope, re.compile(r" {2,}"), " ", "text.spaces_collapsed")
        if s.get("fix_punctuation_spacing"):
            self._regex_pass(ctx, scope, re.compile(r" +(?=[,.;:!?،؛؟])"), "", "text.punct_fixed")
        if s.get("trim_spaces"):
            self._regex_pass(ctx, scope, re.compile(r"^ +| +$"), "", "text.trimmed")
        if "change_case" in s:
            fn = {"upper": str.upper, "lower": str.lower, "title": str.title,
                  "sentence": lambda t: t[:1].upper() + t[1:].lower()}[s["change_case"]]
            self._regex_pass(ctx, scope, re.compile(r"\S+(?: \S+)*" if s["change_case"] == "sentence" else r"[^\W\d_]+"),
                             _case(fn), "text.case_changed")
        if s.get("remove_hyperlinks"):
            n = 0
            for name, _k in ctx.pkg.story_parts(scope.parts):
                root = ctx.pkg.xml(name, edit=True)
                for h in list(root.iter(qn("w:hyperlink"))):
                    par, i = h.getparent(), h.getparent().index(h)
                    for k in list(h):
                        par.insert(i, k); i += 1
                    par.remove(h); n += 1
            ctx.count("text.hyperlinks_removed", n)
        if s.get("remove_hidden_text"):
            n = 0
            for _n, _r, p in iter_paragraphs(ctx, scope, edit=True):
                for r in paragraph_runs(p):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is not None and rpr.find(qn("w:vanish")) is not None and get(rpr.find(qn("w:vanish")), "w:val", "1") != "0":
                        r.getparent().remove(r); n += 1
            ctx.count("text.hidden_removed", n)
        if s.get("remove_empty_line_breaks"):
            ctx.count("text.line_breaks_removed", self._empty_breaks(ctx, scope))
        if s.get("remove_empty_pages"):
            ctx.count("text.empty_pages_removed", self._empty_pages(ctx, scope))
        if s.get("remove_empty_paragraphs") or "max_consecutive_empty" in s:
            keep = 0 if s.get("remove_empty_paragraphs") else s["max_consecutive_empty"]
            ctx.count("text.empty_removed", self._empty(ctx, scope, keep))

    @staticmethod
    def _classify(p):
        """'content' | 'empty' | 'break' (page-break-only paragraph) | 'section'."""
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            return "section"
        brk = ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None
        for el in p.iter():
            if not isinstance(el.tag, str):
                continue
            l = local(el)
            if l == "t" and (el.text or "").strip():
                return "content"
            if l in ("drawing", "pict", "object", "sym", "fldChar", "fldSimple", "instrText", "tab",
                     "sdt", "hyperlink", "oMath", "commentReference", "footnoteReference", "endnoteReference"):
                return "content"
            if l == "br":
                t = get(el, "w:type", "textWrapping")
                if t == "page":
                    brk = True
                elif t == "column":
                    return "content"
        return "break" if brk else "empty"

    @classmethod
    def _empty_pages(cls, ctx, scope):
        """Drop page breaks that would leave a page with no content (and the empties on it)."""
        if "body" not in scope.parts:
            return 0
        body = ctx.pkg.xml("word/document.xml", edit=True).find(qn("w:body"))
        n, content, pending, last_break = 0, False, [], None
        for el in list(body):
            if el.tag == qn("w:tbl"):
                content, pending, last_break = True, [], None
                continue
            if el.tag != qn("w:p"):
                continue
            kind = cls._classify(el)
            if kind == "content":
                content, pending, last_break = True, [], None
            elif kind == "section":
                content, pending, last_break = False, [], None
            elif kind == "empty":
                pending.append(el)
            else:  # break
                if not content:
                    for e in pending + [el]:
                        body.remove(e)
                    n += 1
                else:
                    last_break = el
                pending = []
                content = False
        if not content and last_break is not None and last_break.getparent() is body:
            for e in pending + [last_break]:
                if e.getparent() is body:
                    body.remove(e)
            n += 1
        elif not content and pending and n == 0 and False:
            pass
        return n

    @staticmethod
    def _empty_breaks(ctx, scope):
        n = 0
        for _n, _r, p in iter_paragraphs(ctx, scope, edit=True):
            runs = paragraph_runs(p)
            plain = [r for r in runs if len(r.findall(qn("w:br"))) and
                     all(local(c) in ("rPr", "br") and (local(c) != "br" or get(c, "w:type", "textWrapping") == "textWrapping")
                         for c in r)]
            if not plain:
                continue
            if all(r in plain for r in runs):          # only line breaks => remove the paragraph's breaks
                if len(plain) > 0:
                    for r in plain:
                        r.getparent().remove(r)
                        n += 1
                continue
            for seq in (runs, list(reversed(runs))):   # trim leading and trailing breaks
                for r in seq:
                    if r in plain:
                        r.getparent().remove(r)
                        n += 1
                    else:
                        break
        return n

    @staticmethod
    def _empty(ctx, scope, keep):
        n = 0
        for name, _k in ctx.pkg.story_parts(scope.parts):
            root = ctx.pkg.xml(name, edit=True)
            for container in {p.getparent() for p in root.iter(qn("w:p"))}:
                if container.tag == qn("w:tc"):
                    continue          # a table cell must keep >= 1 paragraph
                run = 0
                for p in list(container.findall(qn("w:p"))):
                    if has_visible_content(p):
                        run = 0
                        continue
                    run += 1
                    if run > keep:
                        container.remove(p)
                        n += 1
        return n

    def read(self, pkg, include_text=False):
        from .base import Context
        ctx = Context(pkg)
        c = {"paragraphs": 0, "empty_paragraphs": 0, "persian_digits": 0, "english_digits": 0, "arabic_digits": 0,
             "arabic_yeh_kaf": 0, "double_spaces": 0, "hyperlinks": 0, "hidden_runs": 0, "words": 0}
        for _n, root, p in iter_paragraphs(ctx):
            c["paragraphs"] += 1
            if not has_visible_content(p):
                c["empty_paragraphs"] += 1
            t = "".join(x.text or "" for x in p.iter(qn("w:t")))
            c["persian_digits"] += len(re.findall("[۰-۹]", t))
            c["arabic_digits"] += len(re.findall("[٠-٩]", t))
            c["english_digits"] += len(re.findall("[0-9]", t))
            c["arabic_yeh_kaf"] += len(re.findall("[كيى]", t))
            c["double_spaces"] += len(re.findall(" {2,}", t))
            c["words"] += len(t.split())
        for name, _k in pkg.story_parts():
            r = pkg.xml(name)
            c["hyperlinks"] += len(list(r.iter(qn("w:hyperlink"))))
            c["hidden_runs"] += len(list(r.iter(qn("w:vanish"))))
        return c
