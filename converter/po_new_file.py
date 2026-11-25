#!/usr/bin/env python3
"""
Create brand-new, empty, valid Office files from nothing (no template files needed).

The OOXML families are written directly (a Word / Excel / PowerPoint package with only the parts Office requires).
Every other format (OpenDocument, old binary .doc/.xls/.ppt, RTF, StarOffice) is produced by editing the OOXML twin and
converting it with LibreOffice when the file is saved (see EDIT_AS / SAVE_AS).
"""
import io, time, zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

MAIN_CT = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "docm": "application/vnd.ms-word.document.macroEnabled.main+xml",
    "dotm": "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "ppsx": "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
    "potx": "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
    "pptm": "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    "ppsm": "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
    "potm": "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
}
# formats that are edited as an OOXML twin and converted by LibreOffice when saved
EDIT_AS = {"odt": "docx", "ott": "docx", "rtf": "docx", "doc": "docx", "dot": "docx", "fodt": "docx",
           "ods": "xlsx", "ots": "xlsx", "xls": "xlsx", "xlt": "xlsx", "fods": "xlsx",
           "odp": "pptx", "otp": "pptx", "ppt": "pptx", "pps": "pptx", "pot": "pptx", "fodp": "pptx"}
GROUPS = {
    "Word": ["docx", "dotx", "docm", "dotm", "odt", "ott", "rtf", "doc", "dot", "fodt"],
    "Excel": ["xlsx", "xltx", "xlsm", "xltm", "ods", "ots", "xls", "xlt", "fods"],
    "PowerPoint": ["pptx", "ppsx", "potx", "pptm", "ppsm", "potm", "odp", "otp", "ppt", "pps", "pot", "fodp"],
}
OLD_STAR = "The OpenOffice 1.x format (.%s) can be opened but not created: this LibreOffice no longer writes it. Create an OpenDocument file (.odt .ods .odp) instead."
NOT_CREATABLE = {"xlsb": "Excel Binary (.xlsb) can be opened but not created (LibreOffice cannot write it). Create an .xlsx instead.",
                 "sxw": OLD_STAR % "sxw", "stw": OLD_STAR % "stw", "sxc": OLD_STAR % "sxc", "sxi": OLD_STAR % "sxi"}
LABELS = {"docx": "Word document", "dotx": "Word template", "docm": "Word macro-enabled document", "dotm": "Word macro-enabled template",
          "odt": "OpenDocument text", "ott": "OpenDocument text template", "rtf": "Rich Text Format", "doc": "Word 97-2003", "dot": "Word 97-2003 template", "fodt": "Flat OpenDocument text",
          "xlsx": "Excel workbook", "xltx": "Excel template", "xlsm": "Excel macro-enabled workbook", "xltm": "Excel macro-enabled template",
          "ods": "OpenDocument spreadsheet", "ots": "OpenDocument spreadsheet template", "xls": "Excel 97-2003", "xlt": "Excel 97-2003 template", "fods": "Flat OpenDocument spreadsheet",
          "pptx": "PowerPoint presentation", "ppsx": "PowerPoint slide show", "potx": "PowerPoint template", "pptm": "PowerPoint macro-enabled presentation",
          "ppsm": "PowerPoint macro-enabled slide show", "potm": "PowerPoint macro-enabled template", "odp": "OpenDocument presentation",
          "otp": "OpenDocument presentation template", "ppt": "PowerPoint 97-2003", "pps": "PowerPoint 97-2003 slide show", "pot": "PowerPoint 97-2003 template", "fodp": "Flat OpenDocument presentation"}


def catalog():
    return {"groups": GROUPS, "labels": LABELS, "not_creatable": NOT_CREATABLE, "edit_as": EDIT_AS}


def _pack(parts):
    buf = io.BytesIO()
    now = time.localtime()[:6]
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:                              # [Content_Types].xml first, as Office writes it
            zi = zipfile.ZipInfo(name, now)
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, data if isinstance(data, bytes) else data.encode("utf-8"))
    return buf.getvalue()


def _content_types(main_part, main_ct, extra):
    ov = '<Override PartName="/%s" ContentType="%s"/>' % (main_part, main_ct) + "".join('<Override PartName="/%s" ContentType="%s"/>' % kv for kv in extra)
    return DECL + '<Types xmlns="%s"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>%s</Types>' % (CT, ov)


def _root_rels(main_part):
    return DECL + ('<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/officeDocument" Target="%s"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                   '<Relationship Id="rId3" Type="%s/extended-properties" Target="docProps/app.xml"/></Relationships>') % (PR, OD, main_part, OD)


def _core():
    t = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    return DECL + ('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
                   'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Pocket Office</dc:creator>'
                   '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified></cp:coreProperties>') % (t, t)


def _app(name):
    return DECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>%s</Application></Properties>' % name


CORE_APP = [("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"), ("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")]


def new_docx(ext="docx"):
    doc = DECL + ('<w:document xmlns:w="%s" xmlns:r="%s"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>') % (W, R)
    styles = DECL + ('<w:styles xmlns:w="%s"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Arial"/>'
                     '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="fa-IR"/></w:rPr></w:rPrDefault>'
                     '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                     '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
                     '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>'
                     '</w:styles>') % W
    rels = DECL + '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/styles" Target="styles.xml"/></Relationships>' % (PR, OD)
    ct = _content_types("word/document.xml", MAIN_CT[ext], [("word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")] + CORE_APP)
    return _pack([("[Content_Types].xml", ct), ("_rels/.rels", _root_rels("word/document.xml")), ("word/document.xml", doc), ("word/_rels/document.xml.rels", rels),
                  ("word/styles.xml", styles), ("docProps/core.xml", _core()), ("docProps/app.xml", _app("Microsoft Office Word"))])


def new_xlsx(ext="xlsx"):
    S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    wb = DECL + '<workbook xmlns="%s" xmlns:r="%s"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>' % (S, R)
    sheet = DECL + '<worksheet xmlns="%s"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>' % S
    styles = DECL + ('<styleSheet xmlns="%s"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
                     '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
                     '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                     '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                     '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
                     '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>') % S
    rels = DECL + ('<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>'
                   '<Relationship Id="rId2" Type="%s/styles" Target="styles.xml"/></Relationships>') % (PR, OD, OD)
    ct = _content_types("xl/workbook.xml", MAIN_CT[ext], [("xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
                                                        ("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")] + CORE_APP)
    return _pack([("[Content_Types].xml", ct), ("_rels/.rels", _root_rels("xl/workbook.xml")), ("xl/workbook.xml", wb), ("xl/_rels/workbook.xml.rels", rels),
                  ("xl/worksheets/sheet1.xml", sheet), ("xl/styles.xml", styles), ("docProps/core.xml", _core()), ("docProps/app.xml", _app("Microsoft Excel"))])


_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _theme():
    def fills():
        return ('<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
                '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
                '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
                '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
                '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>')
    clr = ''.join('<a:%s><a:srgbClr val="%s"/></a:%s>' % (n, v, n) for n, v in (("dk2", "44546A"), ("lt2", "E7E6E6"), ("accent1", "4472C4"), ("accent2", "ED7D31"), ("accent3", "A5A5A5"),
                                                                              ("accent4", "FFC000"), ("accent5", "5B9BD5"), ("accent6", "70AD47"), ("hlink", "0563C1"), ("folHlink", "954F72")))
    return DECL + ('<a:theme xmlns:a="%s" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
                   '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>%s</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
                   '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office">%s</a:fmtScheme></a:themeElements></a:theme>') % (_A, clr, fills())


def _ph(idx, typ, name, x, y, cx, cy, text=""):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="%s"%s/></p:nvPr></p:nvSpPr><p:spPr>%s</p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></p:txBody></p:sp>') % (
        idx + 1, name, typ, ' idx="1"' if typ == "subTitle" else "", '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (x, y, cx, cy) if cx else "",
        text or '<a:endParaRPr lang="en-US"/>')


def new_pptx(ext="pptx"):
    def hdr(extra=""):
        return DECL + '<%s xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"%s>'
    pres = DECL + ('<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
                   '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>') % (_A, R, _P)
    pres_rels = DECL + ('<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
                        '<Relationship Id="rId2" Type="%s/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="%s/theme" Target="theme/theme1.xml"/></Relationships>') % (PR, OD, OD, OD)
    grp = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    master = DECL + ('<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>%s%s%s</p:spTree></p:cSld>'
                     '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
                     '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>') % (
        _A, R, _P, grp, _ph(1, "title", "Title Placeholder", 838200, 365125, 10515600, 1325563), _ph(2, "body", "Text Placeholder", 838200, 1825625, 10515600, 4351338))
    master_rels = DECL + ('<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                          '<Relationship Id="rId2" Type="%s/theme" Target="../theme/theme1.xml"/></Relationships>') % (PR, OD, OD)
    layout = DECL + ('<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" type="title" preserve="1"><p:cSld name="Title Slide"><p:spTree>%s%s%s</p:spTree></p:cSld>'
                     '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>') % (
        _A, R, _P, grp, _ph(1, "ctrTitle", "Title 1", 1524000, 1122363, 9144000, 2387600), _ph(2, "subTitle", "Subtitle 2", 1524000, 3602038, 9144000, 1655762))
    layout_rels = DECL + '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>' % (PR, OD)
    slide = DECL + ('<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:spTree>%s%s%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>') % (
        _A, R, _P, grp, _ph(1, "ctrTitle", "Title 1", 0, 0, 0, 0), _ph(2, "subTitle", "Subtitle 2", 0, 0, 0, 0))
    slide_rels = DECL + '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>' % (PR, OD)
    pp = DECL + '<p:presentationPr xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"/>' % (_A, R, _P)
    vp = DECL + '<p:viewPr xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"/>' % (_A, R, _P)
    ts = DECL + '<a:tblStyleLst xmlns:a="%s" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>' % _A
    pres_rels = pres_rels.replace("</Relationships>", '<Relationship Id="rId4" Type="%s/presProps" Target="presProps.xml"/><Relationship Id="rId5" Type="%s/viewProps" Target="viewProps.xml"/>'
                                  '<Relationship Id="rId6" Type="%s/tableStyles" Target="tableStyles.xml"/></Relationships>' % (OD, OD, OD))
    base = "application/vnd.openxmlformats-officedocument."
    ct = _content_types("ppt/presentation.xml", MAIN_CT[ext], [
        ("ppt/slideMasters/slideMaster1.xml", base + "presentationml.slideMaster+xml"), ("ppt/slideLayouts/slideLayout1.xml", base + "presentationml.slideLayout+xml"),
        ("ppt/slides/slide1.xml", base + "presentationml.slide+xml"), ("ppt/theme/theme1.xml", base + "theme+xml"), ("ppt/presProps.xml", base + "presentationml.presProps+xml"),
        ("ppt/viewProps.xml", base + "presentationml.viewProps+xml"), ("ppt/tableStyles.xml", base + "presentationml.tableStyles+xml")] + CORE_APP)
    return _pack([("[Content_Types].xml", ct), ("_rels/.rels", _root_rels("ppt/presentation.xml")), ("ppt/presentation.xml", pres), ("ppt/_rels/presentation.xml.rels", pres_rels),
                  ("ppt/slideMasters/slideMaster1.xml", master), ("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels),
                  ("ppt/slideLayouts/slideLayout1.xml", layout), ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels),
                  ("ppt/slides/slide1.xml", slide), ("ppt/slides/_rels/slide1.xml.rels", slide_rels), ("ppt/theme/theme1.xml", _theme()),
                  ("ppt/presProps.xml", pp), ("ppt/viewProps.xml", vp), ("ppt/tableStyles.xml", ts), ("docProps/core.xml", _core()), ("docProps/app.xml", _app("Microsoft Office PowerPoint"))])


def create(ext):
    """-> (bytes of the file to EDIT, edit_ext, save_ext). save_ext != edit_ext means: converted by LibreOffice when saved."""
    ext = (ext or "").lower()
    if ext in NOT_CREATABLE:
        raise ValueError(NOT_CREATABLE[ext])
    base = ext if ext in MAIN_CT else EDIT_AS.get(ext)
    if base is None:
        raise ValueError("cannot create .%s files (supported: %s)" % (ext, ", ".join(sorted(set(MAIN_CT) | set(EDIT_AS)))))
    kind = {"d": new_docx, "x": new_xlsx, "p": new_pptx}[base[0]]
    return kind(base), base, ext
