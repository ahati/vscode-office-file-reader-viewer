#!/usr/bin/env python3
"""Generate minimal but valid OOXML sample files (docx/xlsx/pptx) for manual testing."""

import zipfile
from pathlib import Path

OUT = Path(__file__).parent

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {overrides}
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{main}"/>
</Relationships>"""


def write_zip(path: Path, files: dict[str, str]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            z.writestr(name, content)


def w(
    tag: str,
    attrs: dict[str, str] | str | None = None,
    *children: str,
) -> str:
    if isinstance(attrs, str):
        a = f" {attrs}"
    else:
        a = "".join(f" {k}=\"{v}\"" for k, v in (attrs or {}).items())
    inner = "".join(children)
    return f"<{tag}{a}>{inner}</{tag}>" if children else f"<{tag}{a}/>"


# ---------------------------------------------------------------- DOCX
def make_docx() -> None:
    ns_w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def para(text: str, style: str | None = None, bold: bool = False):
        runs = w(
            "w:r",
            None,
            w("w:rPr", None, w("w:b") if bold else w("w:i")),
            w("w:t", {"xml:space": "preserve"}, text),
        )
        props = w("w:pPr", None, w("w:pStyle", {"w:val": style})) if style else ""
        return w("w:p", ns_w, props, runs) if style else w("w:p", ns_w, runs)

    body = "".join([
        para("Sample Document", style="Title"),
        para("Heading 1", style="Heading1"),
        para("This is a minimal OOXML document used to exercise the viewer."),
        para("Heading 2", style="Heading2"),
        para("A second paragraph with "),
        para("bold text inline.", bold=True),
        para("A bullet item", ),
    ])

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {ns_w} xmlns:r="{R}">'
        f"<w:body>{body}"
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
        "</w:body></w:document>"
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles {ns_w}>'
        + w(
            "w:style",
            {"w:type": "paragraph", "w:styleId": "Title"},
            w("w:name", {"w:val": "Title"}),
            w("w:pPr", None, w("w:jc", {"w:val": "center"})),
            w("w:rPr", None, w("w:sz", {"w:val": "56"}), w("w:b")),
        )
        + w(
            "w:style",
            {"w:type": "paragraph", "w:styleId": "Heading1"},
            w("w:name", {"w:val": "heading 1"}),
            w("w:rPr", None, w("w:sz", {"w:val": "32"}), w("w:b"), w("w:color", {"w:val": "2E74B5"})),
        )
        + w(
            "w:style",
            {"w:type": "paragraph", "w:styleId": "Heading2"},
            w("w:name", {"w:val": "heading 2"}),
            w("w:rPr", None, w("w:sz", {"w:val": "26"}), w("w:b"), w("w:color", {"w:val": "2E74B5"})),
        )
        + "</w:styles>"
    )

    write_zip(
        OUT / "sample.docx",
        {
            "[Content_Types].xml": CONTENT_TYPES.format(
                overrides='<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            ),
            "_rels/.rels": ROOT_RELS.format(main="word/document.xml"),
            "word/_rels/document.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
                "</Relationships>"
            ),
            "word/document.xml": document,
            "word/styles.xml": styles,
        },
    )


# ---------------------------------------------------------------- XLSX
def make_xlsx() -> None:
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def cell(ref: str, value: str, kind: str = "n") -> str:
        t = ' t="inlineStr"' if kind == "s" else ""
        inner = (
            f"<is><t>{value}</t></is>"
            if kind == "s"
            else f"<v>{value}</v>"
        )
        return f'<c r="{ref}"{t}>{inner}</c>'

    sheet1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{ns}">'
        "<sheetData>"
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Q1</t></is></c><c r="C1" t="inlineStr"><is><t>Q2</t></is></c></row>'
        '<row r="2"><c r="A2" t="inlineStr"><is><t>North</t></is></c><c r="B2"><v>1200</v></c><c r="C2"><v>1350</v></c></row>'
        '<row r="3"><c r="A3" t="inlineStr"><is><t>South</t></is></c><c r="B3"><v>980</v></c><c r="C3"><v>1120</v></c></row>'
        '<row r="4"><c r="A4" t="inlineStr"><is><t>East</t></is></c><c r="B4"><v>1440</v></c><c r="C4"><v>1610</v></c></row>'
        '<row r="5"><c r="A5" t="inlineStr"><is><t>West</t></is></c><c r="B5"><v>870</v></c><c r="C5"><v>905</v></c></row>'
        '<row r="7"><c r="A7" t="inlineStr"><is><t>Total (formula cache)</t></is></c><c r="B7"><v>4490</v></c><c r="C7"><v>4985</v></c></row>'
        "</sheetData>"
        '<sheetFormatPr defaultRowHeight="15"/>'
        "</worksheet>"
    )
    sheet2 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{ns}">'
        "<sheetData>"
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Second sheet</t></is></c></row>'
        '<row r="2"><c r="A2" t="inlineStr"><is><t>Switch tabs below.</t></is></c></row>'
        "</sheetData></worksheet>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )

    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{ns}" xmlns:r="{R}">'
        "<sheets>"
        '<sheet name="Sales" sheetId="1" r:id="rId1"/>'
        '<sheet name="Notes" sheetId="2" r:id="rId2"/>'
        "</sheets></workbook>"
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{R}/worksheet" Target="worksheets/sheet2.xml"/>'
        f'<Relationship Id="rId3" Type="{R}/styles" Target="styles.xml"/>'
        "</Relationships>"
    )

    write_zip(
        OUT / "sample.xlsx",
        {
            "[Content_Types].xml": CONTENT_TYPES.format(
                overrides='<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            ),
            "_rels/.rels": ROOT_RELS.format(main="xl/workbook.xml"),
            "xl/workbook.xml": workbook,
            "xl/_rels/workbook.xml.rels": wb_rels,
            "xl/styles.xml": styles_xml,
            "xl/worksheets/sheet1.xml": sheet1,
            "xl/worksheets/sheet2.xml": sheet2,
        },
    )


# ---------------------------------------------------------------- PPTX
def make_pptx() -> None:
    p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def slide(title: str, body: str) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p:sld xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{R}">'
            "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
            '<p:spPr/>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{title}</a:t></a:r></a:p></p:txBody></p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr/>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{body}</a:t></a:r></a:p></p:txBody></p:sp>'
            "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        )

    presentation = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentation xmlns:p="{p}" xmlns:r="{R}" xmlns:a="{a}">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>'
        '<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
    )
    theme = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<a:theme xmlns:a="{a}" name="Sample Theme">'
        '<a:themeElements>'
        '<a:clrScheme name="Sample"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
        '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
        '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
        '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
        '<a:fontScheme name="Sample"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
        '<a:fmtScheme name="Sample"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
        '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
        '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
        '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
        '</a:fmtScheme></a:themeElements></a:theme>'
    )
    master = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sldMaster xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{R}">'
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>'
    )
    layout = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sldLayout xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{R}" type="obj" preserve="1">'
        '<p:cSld name="Title and Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
    )

    write_zip(
        OUT / "sample.pptx",
        {
            "[Content_Types].xml": CONTENT_TYPES.format(
                overrides='<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
                '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
                '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
                '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
                '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
            ),
            "_rels/.rels": ROOT_RELS.format(main="ppt/presentation.xml"),
            "ppt/presentation.xml": presentation,
            "ppt/_rels/presentation.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{R}/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
                f'<Relationship Id="rId2" Type="{R}/slide" Target="slides/slide1.xml"/>'
                f'<Relationship Id="rId3" Type="{R}/slide" Target="slides/slide2.xml"/>'
                f'<Relationship Id="rId4" Type="{R}/slideLayout" Target="slideLayouts/slideLayout1.xml"/>'
                f'<Relationship Id="rId5" Type="{R}/theme" Target="theme/theme1.xml"/>'
                "</Relationships>"
            ),
            "ppt/slideMasters/slideMaster1.xml": master,
            "ppt/slideMasters/_rels/slideMaster1.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                f'<Relationship Id="rId2" Type="{R}/theme" Target="../theme/theme1.xml"/>'
                "</Relationships>"
            ),
            "ppt/theme/theme1.xml": theme,
            "ppt/slideLayouts/slideLayout1.xml": layout,
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
                "</Relationships>"
            ),
            "ppt/slides/slide1.xml": slide("Sample Deck", "Slide 1 — rendered locally with WASM + Canvas."),
            "ppt/slides/slide2.xml": slide("Second slide", "Use the toolbar or arrow keys to navigate."),
        },
    )


if __name__ == "__main__":
    make_docx()
    make_xlsx()
    make_pptx()
    print("generated:", ", ".join(p.name for p in sorted(OUT.glob("sample.*"))))
