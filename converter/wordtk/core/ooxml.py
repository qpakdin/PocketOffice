# -*- coding: utf-8 -*-
"""Namespaces, element creation and schema-ordered child insertion.

Word is strict about the order of children inside pPr / rPr / sectPr / settings,
so every property write goes through :func:`insert_ordered`.
"""
from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xml": "http://www.w3.org/XML/1998/namespace",
}
W = NS["w"]
RELT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def qn(name):
    """'w:p' -> '{ns}p'. Names without a prefix are returned unchanged."""
    if name.startswith("{") or ":" not in name:
        return name
    p, l = name.split(":", 1)
    return "{%s}%s" % (NS[p], l)


def local(el):
    t = el.tag
    return t.split("}", 1)[1] if isinstance(t, str) and "}" in t else t


def make(tag, attrs=None, text=None):
    """Create a standalone element. attrs keys may be 'w:val' style."""
    prefix = tag.split(":", 1)[0] if ":" in tag else "w"
    e = etree.Element(qn(tag), nsmap={prefix: NS[prefix]})
    for k, v in (attrs or {}).items():
        if v is not None:
            e.set(qn(k), str(v))
    if text is not None:
        e.text = text
    return e


def get(el, name, default=None):
    if el is None:
        return default
    return el.get(qn(name), default)


def set_attrs(el, attrs):
    """Set attributes; a value of None removes the attribute."""
    for k, v in attrs.items():
        if v is None:
            el.attrib.pop(qn(k), None)
        else:
            el.set(qn(k), str(v))
    return el


# ----------------------------------------------------------------- ordering
ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
            "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
            "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
            "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
            "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
            "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
            "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
            "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type",
               "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
               "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
               "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
              "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
              "tblLook", "tblCaption", "tblDescription"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
             "tblHeader", "tblCellSpacing", "jc", "hidden"],
    "tblBorders": ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV"],
    "tblCellMar": ["top", "left", "start", "bottom", "right", "end"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark"],
    "numPr": ["ilvl", "numId", "numberingChange", "ins"],
    "pBdr": ["top", "left", "bottom", "right", "between", "bar"],
    "pgBorders": ["top", "left", "bottom", "right"],
    "lvl": ["start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText",
            "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr"],
    "style": ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
              "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal",
              "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr",
              "tcPr", "tblStylePr"],
    "settings": ["writeProtection", "view", "zoom", "removePersonalInformation",
                 "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape",
                 "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData",
                 "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData",
                 "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
                 "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
                 "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign",
                 "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
                 "stylePaneSortMethod", "documentType", "mailMerge", "revisionView",
                 "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting",
                 "documentProtection", "autoFormatOverride", "styleLockTheme",
                 "styleLockQFSet", "defaultTabStop", "autoHyphenation",
                 "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
                 "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
                 "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
                 "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
                 "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
                 "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
                 "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
                 "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
                 "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
                 "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
                 "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
                 "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
                 "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
                 "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
                 "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang",
                 "clrSchemeMapping", "doNotIncludeSubdocsInStats",
                 "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown",
                 "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags",
                 "decimalSymbol", "listSeparator"],
}


def insert_ordered(parent, new):
    """Insert ``new`` into ``parent`` respecting the schema sequence."""
    order = ORDER.get(local(parent))
    ln = local(new)
    if order is None or ln not in order:
        parent.append(new)
        return new
    idx = order.index(ln)
    for ch in parent:
        if not isinstance(ch.tag, str):
            continue
        lc = local(ch)
        if lc in order and order.index(lc) > idx:
            ch.addprevious(new)
            return new
    parent.append(new)
    return new


def child(parent, tag):
    return parent.find(qn(tag)) if parent is not None else None


def ensure(parent, tag, attrs=None):
    """Return child ``tag`` (creating it in schema order) and set attrs."""
    c = parent.find(qn(tag))
    if c is None:
        c = insert_ordered(parent, make(tag))
    if attrs:
        set_attrs(c, attrs)
    return c


def remove(parent, tag):
    if parent is None:
        return
    for c in parent.findall(qn(tag)):
        parent.remove(c)


def set_onoff(parent, tag, value):
    """True -> <w:tag/>, False -> <w:tag w:val="0"/>, None -> removed."""
    if value is None:
        remove(parent, tag)
        return
    c = ensure(parent, tag)
    if value:
        c.attrib.pop(qn("w:val"), None)
    else:
        c.set(qn("w:val"), "0")


def is_on(el):
    """Interpret an on/off element (None -> None)."""
    if el is None:
        return None
    return get(el, "w:val", "1").lower() not in ("0", "false", "off")


def ensure_ppr(p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = make("w:pPr")
        p.insert(0, ppr)
    return ppr


def ensure_rpr(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = make("w:rPr")
        r.insert(0, rpr)
    return rpr


def paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
