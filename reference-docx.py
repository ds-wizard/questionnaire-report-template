"""Restyle src/reference.docx (the pandoc reference document of the MS Word format).

Sets the fonts, colours, page size, header/footer and the paragraph styles used by
src/for-word.html.j2 (`data-custom-style="..."`). The colours match src/default.css.
Edits the file in place; run it again after changing the values below:

    uv run --with lxml python reference-docx.py
"""
import pathlib
import shutil
import tempfile
import zipfile

from lxml import etree

DOCX = pathlib.Path(__file__).parent / 'src' / 'reference.docx'

FONT = 'Calibri'
ACCENT = '1D5F8A'
ACCENT_DARK = '123F5C'
ACCENT_SOFT = 'E8F1F7'
TEXT = '212529'
MUTED = '5F6B76'
LIGHT = '8A949E'
BORDER = 'DDE2E6'
SURFACE = 'F5F7F9'
DANGER = 'B4232F'
DANGER_SOFT = 'FBECED'

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def q(tag):
    return f'{{{W}}}{tag}'


def xml(fragment):
    return etree.fromstring(f'<root xmlns:w="{W}">{fragment}</root>')[0]


def border(side, color, size, space=4):
    return f'<w:{side} w:val="single" w:sz="{size}" w:space="{space}" w:color="{color}"/>'


def rpr(size=None, color=None, bold=False, italic=False, caps=False, spacing=None,
        shade=None, font=True):
    parts = []
    if font:
        parts.append(f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" '
                     f'w:cs="{FONT}"/>')
    if bold:
        parts.append('<w:b/><w:bCs/>')
    else:
        parts.append('<w:b w:val="0"/><w:bCs w:val="0"/>')
    parts.append('<w:i/><w:iCs/>' if italic else '<w:i w:val="0"/><w:iCs w:val="0"/>')
    if caps:
        parts.append('<w:caps/>')
    if color:
        parts.append(f'<w:color w:val="{color}"/>')
    if spacing:
        parts.append(f'<w:spacing w:val="{spacing}"/>')
    if size:
        parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    if shade:
        parts.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>')
    return f'<w:rPr>{"".join(parts)}</w:rPr>'


def ppr(before=0, after=120, jc='left', keep_next=False, borders='', shade=None, ind=None,
        outline=None, line=264):
    parts = []
    if keep_next:
        parts.append('<w:keepNext/><w:keepLines/>')
    if borders:
        parts.append(f'<w:pBdr>{borders}</w:pBdr>')
    if shade:
        parts.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>')
    parts.append(f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" '
                 'w:lineRule="auto"/>')
    if ind:
        parts.append(f'<w:ind w:left="{ind}" w:right="{ind}"/>')
    parts.append(f'<w:jc w:val="{jc}"/>')
    if outline is not None:
        parts.append(f'<w:outlineLvl w:val="{outline}"/>')
    return f'<w:pPr>{"".join(parts)}</w:pPr>'


def set_style(styles, style_id, name, ppr_xml='', rpr_xml='', kind='paragraph',
              based='BodyText', next_style=None, custom=True):
    old = styles.find(f'w:style[@w:styleId="{style_id}"]', NS)
    attrs = f'w:type="{kind}" w:styleId="{style_id}"' + (' w:customStyle="1"' if custom else '')
    inner = f'<w:name w:val="{name}"/>'
    if based:
        inner += f'<w:basedOn w:val="{based}"/>'
    if next_style:
        inner += f'<w:next w:val="{next_style}"/>'
    if old is not None:
        for keep in ('uiPriority', 'unhideWhenUsed', 'link'):
            el = old.find(f'w:{keep}', NS)
            if el is not None:
                inner += etree.tostring(el).decode()
    inner += '<w:qFormat/>' + ppr_xml + rpr_xml
    new = xml(f'<w:style {attrs}>{inner}</w:style>')
    if old is not None:
        old.getparent().replace(old, new)
    else:
        styles.append(new)


def restyle_styles(root):
    defaults = root.find('w:docDefaults', NS)
    defaults.getparent().replace(defaults, xml(
        '<w:docDefaults><w:rPrDefault>'
        + rpr(size=20, color=TEXT).replace('<w:b w:val="0"/><w:bCs w:val="0"/>', '')
                                  .replace('<w:i w:val="0"/><w:iCs w:val="0"/>', '')
                                  .replace('</w:rPr>', '<w:lang w:val="en-US" w:eastAsia="en-US" '
                                                       'w:bidi="ar-SA"/></w:rPr>')
        + '</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" '
          'w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'))

    # pandoc's own styles
    set_style(root, 'BodyText', 'Body Text', ppr(0, 120), based='Normal', custom=False)
    set_style(root, 'FirstParagraph', 'First Paragraph', ppr(0, 120), next_style='BodyText')
    set_style(root, 'Compact', 'Compact', ppr(20, 20))
    set_style(root, 'Title', 'Title',
              ppr(3200, 360, keep_next=True, borders=border('bottom', ACCENT, 24, 10)),
              rpr(56, ACCENT_DARK, bold=True), based='Normal', next_style='BodyText',
              custom=False)
    set_style(root, 'Subtitle', 'Subtitle', ppr(0, 240, keep_next=True),
              rpr(28, MUTED), based='Title', next_style='BodyText', custom=False)
    headings = {  # level: (size, before, after)
        1: (36, 0, 200), 2: (24, 320, 60), 3: (22, 240, 60),
        4: (21, 200, 40), 5: (20, 160, 40), 6: (20, 160, 40),
    }
    for level, (size, before, after) in headings.items():
        borders = border('bottom', ACCENT, 16, 6) if level == 1 else ''
        set_style(root, f'Heading{level}', f'heading {level}',
                  ppr(before, after, keep_next=True, borders=borders, outline=level - 1),
                  rpr(size, ACCENT_DARK, bold=True), based='Normal', next_style='BodyText',
                  custom=False)
    set_style(root, 'Hyperlink', 'Hyperlink', '', rpr(color=ACCENT, font=False)
              .replace('<w:b w:val="0"/><w:bCs w:val="0"/>', '')
              .replace('<w:i w:val="0"/><w:iCs w:val="0"/>', ''),
              kind='character', based=None, custom=False)

    # styles of src/for-word.html.j2
    left_bar = border('left', ACCENT, 18, 6)
    set_style(root, 'Answer', 'Answer',
              ppr(60, 120, borders=left_bar, shade=SURFACE, ind=140),
              rpr(color=TEXT, bold=True))
    set_style(root, 'NotAnswered', 'Not Answered',
              ppr(60, 120, borders=border('left', BORDER, 18, 6), ind=140),
              rpr(18, LIGHT, italic=True))
    set_style(root, 'NotAnsweredRequired', 'Not Answered Required',
              ppr(60, 120, borders=border('left', DANGER, 18, 6), shade=DANGER_SOFT, ind=140),
              rpr(18, DANGER, italic=True))
    set_style(root, 'QuestionText', 'Question Text', ppr(0, 80), rpr(18, MUTED))
    set_style(root, 'QuestionMeta', 'Question Meta', ppr(0, 40), rpr(16, MUTED))
    set_style(root, 'Advice', 'Advice',
              ppr(40, 120, borders=border('left', ACCENT, 12, 6), shade=ACCENT_SOFT, ind=140),
              rpr(18, ACCENT_DARK))
    set_style(root, 'SectionLabel', 'Section Label', ppr(200, 60, keep_next=True),
              rpr(15, MUTED, bold=True, caps=True, spacing=12))
    set_style(root, 'ChoiceSelected', 'Choice Selected', ppr(10, 10, ind=140),
              rpr(color=TEXT, bold=True))
    set_style(root, 'ChoiceUnselected', 'Choice Unselected', ppr(10, 10, ind=140),
              rpr(color=LIGHT))
    set_style(root, 'ItemHeading', 'Item Heading',
              ppr(200, 60, keep_next=True, borders=border('top', BORDER, 8, 4)),
              rpr(16, ACCENT, bold=True))
    set_style(root, 'ToolInfo', 'Tool Info', ppr(2400, 0, jc='center'), rpr(16, LIGHT))
    # only the colour: bold comes from the heading (bold in both would toggle it off)
    set_style(root, 'QuestionNumber', 'Question Number', '',
              f'<w:rPr><w:color w:val="{ACCENT}"/></w:rPr>', kind='character',
              based='DefaultParagraphFont')

    # tables: light horizontal rules, header row on a grey background
    rule = f'w:val="single" w:sz="4" w:space="0" w:color="{BORDER}"'
    set_style(root, 'Table', 'Table', kind='table', based='TableNormal', ppr_xml=(
        '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        f'<w:tblBorders><w:top {rule}/><w:bottom {rule}/><w:insideH {rule}/></w:tblBorders>'
        '<w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
        '<w:bottom w:w="50" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:caps/>'
        f'<w:color w:val="{MUTED}"/><w:sz w:val="15"/><w:szCs w:val="15"/></w:rPr>'
        '<w:tcPr><w:tcBorders>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{ACCENT}"/></w:tcBorders>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{SURFACE}"/><w:vAlign w:val="bottom"/>'
        '</w:tcPr></w:tblStylePr>'
        '<w:tblStylePr w:type="firstCol"><w:rPr>'
        f'<w:color w:val="{MUTED}"/></w:rPr></w:tblStylePr>'))

    for style_id, size in (('Header', 15), ('Footer', 15)):
        style = root.find(f'w:style[@w:styleId="{style_id}"]', NS)
        old = style.find('w:rPr', NS)
        if old is not None:
            style.remove(old)
        style.append(xml(rpr(size, LIGHT)))


def restyle_header(root):
    p = root.find('.//w:p', NS)
    for run in p.findall('w:r', NS):
        if run.find('w:tab', NS) is not None:
            p.remove(run)
    bdr = p.find('w:pPr/w:pBdr', NS)
    bdr.getparent().replace(bdr, xml(f'<w:pBdr>{border("bottom", BORDER, 4, 4)}</w:pBdr>'))


def restyle_footer(root):
    p = root.find('.//w:p', NS)
    for run in p.findall('w:r', NS):
        if run.find('w:tab', NS) is not None:
            p.remove(run)
    ppr_el = p.find('w:pPr', NS)
    bdr = ppr_el.find('w:pBdr', NS)
    if bdr is not None:
        ppr_el.remove(bdr)
    if ppr_el.find('w:jc', NS) is None:
        ppr_el.append(xml('<w:jc w:val="right"/>'))


def restyle_document(root):
    sect = root.find('.//w:body/w:sectPr', NS)
    sect.find('w:pgSz', NS).attrib.update({q('w'): '11906', q('h'): '16838'})  # A4
    sect.find('w:pgMar', NS).attrib.update({
        q('top'): '1418', q('right'): '1134', q('bottom'): '1134', q('left'): '1134',
        q('header'): '567', q('footer'): '567',
    })


def main():
    editors = {
        'word/styles.xml': restyle_styles,
        'word/header1.xml': restyle_header,
        'word/footer1.xml': restyle_footer,
        'word/document.xml': restyle_document,
    }
    with tempfile.TemporaryDirectory() as tmp:
        out = pathlib.Path(tmp) / 'reference.docx'
        with zipfile.ZipFile(DOCX) as src, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename in editors:
                    root = etree.fromstring(data)
                    editors[info.filename](root)
                    data = etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                                          standalone=True)
                dst.writestr(info, data)
        shutil.copyfile(out, DOCX)
    print(f'restyled {DOCX}')


if __name__ == '__main__':
    main()
