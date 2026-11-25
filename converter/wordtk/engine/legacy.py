# -*- coding: utf-8 -*-
"""Legacy v1 profiles ("operations": raw XPath edits) from the original single-file prototype."""
from ..core.ooxml import NS, qn, serialize
from ..features import ProfileError


def _attr(key, ns):
    if ":" in key:
        pre, loc = key.split(":", 1)
        if pre not in ns:
            raise ProfileError("unknown namespace prefix: " + pre)
        return "{%s}%s" % (ns[pre], loc)
    return key


def apply_legacy(pkg, ops, report):
    from lxml import etree
    for op in ops:
        part = op.get("part", "word/document.xml")
        if not pkg.exists(part):
            if op.get("required"):
                raise ProfileError("part missing: " + part)
            continue
        root = pkg.xml(part, edit=True)
        ns = dict(NS); ns.update(op.get("namespaces", {}))
        targets = root.xpath(op.get("xpath", "/w:document"), namespaces=ns)
        if not targets and op.get("required"):
            raise ProfileError("XPath not found: " + op.get("xpath", ""))
        action = op.get("action", "set_attributes")
        for el in targets:
            if action in ("set_property", "set_attributes"):
                node = el
                if action == "set_property":
                    for piece in filter(None, op.get("child_path", "").split("/")):
                        t = qn(piece) if ":" in piece and not piece.startswith("{") else piece
                        c = node.find(t)
                        if c is None:
                            from ..core.ooxml import insert_ordered
                            c = insert_ordered(node, etree.Element(t))
                        node = c
                for k, v in op.get("attributes", {}).items():
                    a = _attr(k, ns)
                    if v is None:
                        node.attrib.pop(a, None)
                    else:
                        node.set(a, str(v).lower() if isinstance(v, bool) else str(v))
            elif action == "remove":
                if el is root:
                    raise ProfileError("cannot remove the root element")
                el.getparent().remove(el)
            elif action == "set_text":
                el.text = str(op.get("text", ""))
            else:
                raise ProfileError("unknown action: " + action)
        report["legacy.operations"] += 1
