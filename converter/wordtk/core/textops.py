# -*- coding: utf-8 -*-
"""Run-level text operations: find/replace across runs, range isolation.

Matching is done on the concatenated text of a paragraph. Runs that contain
anything other than text (tabs, breaks, drawings, fields...) act as barriers,
so a match never spans them.
"""
import copy
import re

from .ooxml import NS, make, qn, local

BARRIER = "\x00"


class Sym:
    """Replacement that inserts a symbol (w:sym) instead of text."""

    def __init__(self, font, char):
        self.font, self.char = font, char


def _is_text_run(r):
    for c in r:
        if local(c) not in ("rPr", "t"):
            return False
    return True


def run_text(r):
    return "".join(t.text or "" for t in r.findall(qn("w:t")))


def set_run_text(r, text):
    ts = r.findall(qn("w:t"))
    for t in ts[1:]:
        r.remove(t)
    if ts:
        t = ts[0]
    else:
        t = make("w:t")
        r.append(t)
    t.text = text
    t.set("{%s}space" % NS["xml"], "preserve")


def paragraph_runs(p):
    """Runs in document order, excluding deleted (tracked) text."""
    return [r for r in p.iter(qn("w:r")) if not any(local(a) == "del" for a in r.iterancestors())]


def _segments(p):
    segs, pos = [], 0
    for r in paragraph_runs(p):
        if _is_text_run(r) and r.find(qn("w:t")) is not None:
            t = run_text(r)
            segs.append({"run": r, "text": t, "start": pos})
            pos += len(t)
        else:
            segs.append({"run": r, "text": None, "start": pos})
            pos += 1  # barrier
    return segs


def paragraph_string(p):
    s = []
    for sg in _segments(p):
        s.append(sg["text"] if sg["text"] is not None else BARRIER)
    return "".join(s)


def _split(run, offset):
    """Split a text run at ``offset``; returns the right-hand new run."""
    text = run_text(run)
    if offset <= 0 or offset >= len(text):
        return None
    right = copy.deepcopy(run)
    set_run_text(run, text[:offset])
    set_run_text(right, text[offset:])
    run.addnext(right)
    return right


def isolate(p, start, end, segs=None):
    """Split runs so that [start, end) is covered by whole runs; return them."""
    segs = segs or _segments(p)
    cover = [s for s in segs if s["text"] is not None and s["start"] < end and s["start"] + len(s["text"]) > start]
    if not cover:
        return []
    last = cover[-1]
    _split(last["run"], end - last["start"])
    first = cover[0]
    first_run = first["run"]
    if start > first["start"]:
        right = _split(first_run, start - first["start"])
        if right is not None:
            if first is last:
                last_run = right
            first_run = right
    runs = [first_run]
    for s in cover[1:]:
        runs.append(s["run"])
    if first is last:
        # after the two splits the covered part is exactly first_run
        runs = [first_run]
    return runs


def replace_in_paragraph(p, regex, repl):
    """Replace all matches of compiled ``regex``.

    ``repl`` is a callable(match)->str|Sym or a str (may use \\1 groups).
    Returns the number of replacements.
    """
    s = paragraph_string(p)
    matches = [m for m in regex.finditer(s) if BARRIER not in m.group(0) and m.end() > m.start()]
    if not matches:
        return 0
    segs = _segments(p)
    n = 0
    for m in reversed(matches):
        runs = isolate(p, m.start(), m.end(), segs)
        if not runs:
            continue
        value = repl if isinstance(repl, Sym) else repl(m) if callable(repl) else m.expand(repl)
        head = runs[0]
        for extra in runs[1:]:
            extra.getparent().remove(extra)
        if isinstance(value, Sym):
            for t in head.findall(qn("w:t")):
                head.remove(t)
            head.append(make("w:sym", {"w:font": value.font, "w:char": value.char}))
        else:
            set_run_text(head, value)
        n += 1
    return n


def compile_find(find, regex=False, match_case=False, whole_word=False):
    pat = find if regex else re.escape(find)
    if whole_word:
        pat = r"(?<!\w)(?:%s)(?!\w)" % pat
    return re.compile(pat, 0 if match_case else re.IGNORECASE | re.UNICODE)


def has_visible_content(p):
    """False for paragraphs that carry no text/objects (safe to call 'empty')."""
    ppr = p.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
        return True
    for el in p.iter():
        l = local(el) if isinstance(el.tag, str) else ""
        if l == "t" and (el.text or "").strip():
            return True
        if l in ("drawing", "pict", "object", "br", "tab", "sym", "fldChar", "fldSimple",
                 "instrText", "bookmarkStart", "commentReference", "footnoteReference",
                 "endnoteReference", "sdt", "hyperlink", "oMath", "ptab", "cr", "noBreakHyphen"):
            return True
    return False
