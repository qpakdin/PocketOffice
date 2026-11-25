# -*- coding: utf-8 -*-
"""In-memory OPC (zip) package for a DOCX: parts, relationships, content types."""
import os
import posixpath
import re
import time
import zipfile
from pathlib import Path

from lxml import etree

from .ooxml import NS, RELT, qn, serialize

CT_NS = NS["ct"]
REL_NS = NS["rel"]
MAIN_CT = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
}


class DocxError(Exception):
    """The file is not a readable DOCX package."""


def rels_name(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def source_of_rels(relname):
    d, b = posixpath.split(relname)          # word/_rels , document.xml.rels
    base = posixpath.dirname(d)               # word
    return posixpath.join(base, b[:-5]) if b.endswith(".rels") else ""


def resolve_target(source, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


_PARSER_ARGS = dict(resolve_entities=False, no_network=True, huge_tree=True)


class DocxPackage:
    def __init__(self, path):
        self.path = Path(path)
        self._order, self._raw, self._info = [], {}, {}
        try:
            with zipfile.ZipFile(self.path) as z:
                for i in z.infolist():
                    if i.is_dir():
                        continue
                    self._order.append(i.filename)
                    self._raw[i.filename] = z.read(i.filename)
                    self._info[i.filename] = i
        except zipfile.BadZipFile as e:
            raise DocxError("not a valid DOCX (corrupted, or encrypted/password protected): %s" % self.path.name) from e
        if "[Content_Types].xml" not in self._raw:
            raise DocxError("not an OOXML package: %s" % self.path.name)
        self._xml, self._dirty, self._removed = {}, set(), set()
        self.main = self._find_main()
        if self.main not in self._raw:
            raise DocxError("main document part missing in %s" % self.path.name)

    # ------------------------------------------------------------------ parts
    def _find_main(self):
        try:
            rels = self.xml("_rels/.rels")
        except KeyError:
            return "word/document.xml"
        for r in rels:
            if r.get("Type", "").endswith("/officeDocument"):
                return resolve_target("", r.get("Target"))
        return "word/document.xml"

    def names(self):
        out = [n for n in self._order if n not in self._removed]
        out += [n for n in self._xml if n not in self._raw and n not in self._removed and n not in out]
        return out

    def exists(self, name):
        return name not in self._removed and (name in self._raw or name in self._xml)

    def xml(self, name, edit=False):
        """Parsed root of an XML part (cached). ``edit=True`` marks it for saving."""
        if name in self._removed:
            raise KeyError(name)
        if name not in self._xml:
            if name not in self._raw:
                raise KeyError(name)
            try:
                self._xml[name] = etree.fromstring(self._raw[name], etree.XMLParser(**_PARSER_ARGS))
            except etree.XMLSyntaxError as e:
                raise DocxError("XML error in %s: %s" % (name, e)) from e
        if edit:
            self._dirty.add(name)
        return self._xml[name]

    def try_xml(self, name, edit=False):
        try:
            return self.xml(name, edit)
        except KeyError:
            return None

    def raw(self, name):
        if name in self._xml and name in self._dirty:
            return serialize(self._xml[name])
        return self._raw[name] if name in self._raw else serialize(self._xml[name])

    def set_raw(self, name, data):
        if name not in self._raw and name not in self._xml:
            self._order.append(name)
        self._raw[name] = data
        self._xml.pop(name, None)
        self._dirty.discard(name)
        self._removed.discard(name)

    def add_xml_part(self, name, root, content_type=None):
        if name not in self._raw and name not in self._xml:
            self._order.append(name)
        self._xml[name] = root
        self._dirty.add(name)
        self._removed.discard(name)
        if content_type:
            self.ct_override(name, content_type)
        return root

    def remove_part(self, name):
        if not self.exists(name):
            return
        self._removed.add(name)
        self._xml.pop(name, None)
        self._dirty.discard(name)
        rn = rels_name(name)
        if self.exists(rn):
            self._removed.add(rn)
            self._xml.pop(rn, None)
            self._dirty.discard(rn)
        ct = self.xml("[Content_Types].xml", edit=True)
        for o in ct.findall("{%s}Override" % CT_NS):
            if o.get("PartName") == "/" + name:
                ct.remove(o)
        for relname in [n for n in self.names() if n.endswith(".rels")]:
            src = source_of_rels(relname)
            root = self.xml(relname)
            for r in list(root):
                if r.get("TargetMode") != "External" and resolve_target(src, r.get("Target", "")) == name:
                    self.xml(relname, edit=True).remove(r)

    # ----------------------------------------------------------- content types
    def ct_override(self, name, content_type):
        ct = self.xml("[Content_Types].xml", edit=True)
        for o in ct.findall("{%s}Override" % CT_NS):
            if o.get("PartName") == "/" + name:
                o.set("ContentType", content_type)
                return
        e = etree.SubElement(ct, "{%s}Override" % CT_NS)
        e.set("PartName", "/" + name)
        e.set("ContentType", content_type)

    def ct_default(self, ext, content_type):
        ct = self.xml("[Content_Types].xml", edit=True)
        for d in ct.findall("{%s}Default" % CT_NS):
            if d.get("Extension", "").lower() == ext.lower():
                return
        e = etree.Element("{%s}Default" % CT_NS)
        e.set("Extension", ext)
        e.set("ContentType", content_type)
        ct.insert(0, e)

    # ----------------------------------------------------------- relationships
    def rels_root(self, source, edit=False):
        rn = rels_name(source)
        if not self.exists(rn):
            if not edit:
                return etree.Element("{%s}Relationships" % REL_NS, nsmap={None: REL_NS})
            self.add_xml_part(rn, etree.Element("{%s}Relationships" % REL_NS, nsmap={None: REL_NS}))
        return self.xml(rn, edit)

    def rels(self, source, type_suffix=None):
        """[(rId, absolute part name or external target, type, external)]"""
        out = []
        for r in self.rels_root(source):
            t = r.get("Type", "")
            if type_suffix and not t.endswith("/" + type_suffix):
                continue
            ext = r.get("TargetMode") == "External"
            tgt = r.get("Target", "") if ext else resolve_target(source, r.get("Target", ""))
            out.append((r.get("Id"), tgt, t, ext))
        return out

    def add_rel(self, source, rel_type, part, external=False):
        root = self.rels_root(source, edit=True)
        ids = {r.get("Id") for r in root}
        n = 1
        while "rId%d" % n in ids:
            n += 1
        rid = "rId%d" % n
        e = etree.SubElement(root, "{%s}Relationship" % REL_NS)
        e.set("Id", rid)
        e.set("Type", rel_type if rel_type.startswith("http") else RELT + rel_type)
        e.set("Target", part if external else posixpath.relpath(part, posixpath.dirname(source) or "."))
        if external:
            e.set("TargetMode", "External")
        return rid

    def remove_rel(self, source, rid):
        root = self.rels_root(source, edit=True)
        for r in list(root):
            if r.get("Id") == rid:
                root.remove(r)

    def part_for_rid(self, source, rid):
        for i, tgt, _t, ext in self.rels(source):
            if i == rid:
                return None if ext else tgt
        return None

    # ----------------------------------------------------------------- stories
    def story_parts(self, kinds=("body", "headers", "footers", "footnotes", "endnotes")):
        """Yield (part name, kind) for the requested story kinds."""
        if "body" in kinds:
            yield self.main, "body"
        seen = set()
        for kind, suffix in (("headers", "header"), ("footers", "footer"),
                             ("footnotes", "footnotes"), ("endnotes", "endnotes")):
            if kind in kinds:
                for _rid, part, _t, ext in self.rels(self.main, suffix):
                    if not ext and part not in seen and self.exists(part):
                        seen.add(part)
                        yield part, kind

    def document(self, edit=False):
        return self.xml(self.main, edit)

    def body(self, edit=False):
        return self.document(edit).find(qn("w:body"))

    def sections(self, edit=False):
        """All sectPr elements in document order (last one = body-level)."""
        doc = self.document(edit)
        return doc.xpath("/w:document/w:body//w:pPr/w:sectPr | /w:document/w:body/w:sectPr", namespaces=NS)

    # -------------------------------------------------------------------- save
    def save(self, dst):
        dst = Path(dst)
        if dst.exists() and dst.resolve() == self.path.resolve():
            raise DocxError("output must differ from the input file")
        dst.parent.mkdir(parents=True, exist_ok=True)
        tmp = dst.with_name(dst.name + ".tmp")
        now = time.localtime()[:6]
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for name in self.names():
                    if name in self._dirty or (name not in self._raw):
                        data = serialize(self._xml[name])
                    else:
                        data = self._raw[name]
                    old = self._info.get(name)
                    zi = zipfile.ZipInfo(name, date_time=old.date_time if old else now)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zi.external_attr = old.external_attr if old else 0o600 << 16
                    zout.writestr(zi, data)
            os.replace(tmp, dst)
        finally:
            if tmp.exists():
                tmp.unlink()
        return dst
