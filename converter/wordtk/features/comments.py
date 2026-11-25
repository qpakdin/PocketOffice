# -*- coding: utf-8 -*-
"""Comments / revisions / footnotes management."""
from ..core.ooxml import NS, get, qn
from .base import F, Feature

COMMENT_PARTS = ["word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml"]


class CommentsFeature(Feature):
    key = "comments"
    label = "Comments & Revisions"
    description = "Remove comments, accept/reject tracked changes, set author, remove footnotes/endnotes."
    uses_scope = False
    fields = [
        F("remove_comments", "bool", "Remove all comments"),
        F("anonymize_comments", "text", "Rename every comment author to"),
        F("accept_changes", "bool", "Accept all tracked changes"),
        F("reject_changes", "bool", "Reject all tracked changes"),
        F("track_changes", "bool", "Turn track-changes on/off"),
        F("remove_footnotes", "bool", "Remove footnotes & endnotes"),
    ]

    def validate(self, s):
        if s.get("accept_changes") and s.get("reject_changes"):
            from .base import ProfileError
            raise ProfileError("comments: accept_changes and reject_changes are mutually exclusive")
        return s

    def apply(self, ctx, s):
        pkg = ctx.pkg
        if s.get("remove_comments"):
            n = 0
            for name, _k in list(pkg.story_parts()):
                root = pkg.xml(name)
                hits = list(root.iter(qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")))
                if not hits:
                    continue
                root = pkg.xml(name, edit=True)
                for e in list(root.iter(qn("w:commentRangeStart"), qn("w:commentRangeEnd"))):
                    e.getparent().remove(e)
                for e in list(root.iter(qn("w:commentReference"))):
                    r = e.getparent()
                    r.getparent().remove(r)
                    n += 1
            for part in COMMENT_PARTS:
                pkg.remove_part(part)
            ctx.count("comments.removed", n)
        elif "anonymize_comments" in s and pkg.exists("word/comments.xml"):
            root = pkg.xml("word/comments.xml", edit=True)
            for c in root.findall(qn("w:comment")):
                c.set(qn("w:author"), s["anonymize_comments"])
                c.set(qn("w:initials"), s["anonymize_comments"][:2])
                ctx.count("comments.anonymized")
        if s.get("accept_changes") or s.get("reject_changes"):
            ctx.count("revisions.resolved", self._resolve(pkg, accept=bool(s.get("accept_changes"))))
        if "track_changes" in s:
            from ..core.ooxml import set_onoff
            st = pkg.try_xml("word/settings.xml", edit=True)
            if st is not None:
                set_onoff(st, "w:trackRevisions", True if s["track_changes"] else None)
        if s.get("remove_footnotes"):
            n = 0
            for name, _k in list(pkg.story_parts(("body",))):
                root = pkg.xml(name, edit=True)
                for e in list(root.iter(qn("w:footnoteReference"), qn("w:endnoteReference"))):
                    r = e.getparent()
                    r.getparent().remove(r)
                    n += 1
            for part in ("word/footnotes.xml", "word/endnotes.xml"):
                pkg.remove_part(part)
            ctx.count("comments.footnotes_removed", n)

    @staticmethod
    def _resolve(pkg, accept):
        n = 0
        keep, drop = (("ins", "moveTo"), ("del", "moveFrom")) if accept else (("del", "moveFrom"), ("ins", "moveTo"))
        for name, _k in list(pkg.story_parts()):
            root = pkg.xml(name, edit=True)
            for tag in drop[0], drop[1]:
                for e in list(root.iter(qn("w:" + tag))):
                    e.getparent().remove(e); n += 1
            for tag in keep[0], keep[1]:
                for e in list(root.iter(qn("w:" + tag))):
                    parent, idx = e.getparent(), e.getparent().index(e)
                    kids = list(e)
                    for k in kids:
                        if not accept:           # rejected deletions: restore text
                            for dt in k.iter(qn("w:delText")):
                                dt.tag = qn("w:t")
                        parent.insert(idx, k); idx += 1
                    parent.remove(e); n += 1
            for tag in ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "numberingChange"):
                for e in list(root.iter(qn("w:" + tag))):
                    e.getparent().remove(e); n += 1
        return n

    def read(self, pkg, include_text=False):
        out = {"comments": 0, "authors": [], "insertions": 0, "deletions": 0, "format_changes": 0,
               "footnotes": 0, "endnotes": 0}
        c = pkg.try_xml("word/comments.xml")
        if c is not None:
            cs = c.findall(qn("w:comment"))
            out["comments"] = len(cs)
            out["authors"] = sorted({get(x, "w:author", "") for x in cs})
        for name, _k in pkg.story_parts():
            r = pkg.xml(name)
            out["insertions"] += len(r.findall(".//" + qn("w:ins")))
            out["deletions"] += len(r.findall(".//" + qn("w:del")))
            out["format_changes"] += len(list(r.iter(qn("w:rPrChange"), qn("w:pPrChange"))))
        for k, part, tag in (("footnotes", "word/footnotes.xml", "w:footnote"), ("endnotes", "word/endnotes.xml", "w:endnote")):
            r = pkg.try_xml(part)
            if r is not None:
                out[k] = len([e for e in r.findall(qn(tag)) if get(e, "w:type") in (None, "normal")])
        st = pkg.try_xml("word/settings.xml")
        out["track_changes_on"] = st is not None and st.find(qn("w:trackRevisions")) is not None
        return out
