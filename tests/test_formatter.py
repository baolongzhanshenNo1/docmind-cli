"""
Tests for generator/formatter.py v5.1 — category+language rule-map.
"""
import zipfile
from pathlib import Path
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _make_docx(path: Path, paragraphs_xml: str, body_sectpr: str = ''):
    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}">
  <w:body>{paragraphs_xml}{body_sectpr}</w:body>
</w:document>'''
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', doc_xml.encode('utf-8'))
        z.writestr('[Content_Types].xml', b'''<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>''')


def _para(text, sz=24, bold=False, ea=''):
    rpr = f'<w:rPr><w:sz w:val="{sz}"/>'
    if bold:
        rpr += '<w:b/>'
    if ea:
        rpr += f'<w:rFonts w:eastAsia="{ea}" w:ascii="{ea}"/>'
    rpr += '</w:rPr>'
    return f'<w:p><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'


class TestExtractRules:
    """extract_rules — parsing annotations into {category: {lang: Rule}}."""

    def test_basic_rules(self, tmp_path):
        from generator.formatter import extract_rules

        # Use "第一章" prefix so it classifies as 'chapter' (separate from body)
        xml = _para('第一章（三号，黑体，居中）', sz=24) + \
              _para('正文（小四号，宋体，1.5倍行距，两端对齐）', sz=24)
        _make_docx(tmp_path / 'tmpl.docx', xml)

        rule_map = extract_rules(tmp_path / 'tmpl.docx')

        # v5.1 keys by category→lang
        assert 'chapter' in rule_map
        ch = rule_map['chapter']['cn']
        assert ch.font_ea == '黑体'
        assert ch.font_size == 32
        assert ch.bold is True
        assert ch.alignment == 'center'

        assert 'body' in rule_map
        bd = rule_map['body']['cn']
        assert bd.font_ea == '宋体'
        assert bd.font_size == 24
        assert bd.line_spacing == 360
        assert bd.alignment == 'both'

    def test_multiple_headings(self, tmp_path):
        from generator.formatter import extract_rules

        # "第一章"→chapter, "1.1"→section, "正文"→body
        xml = _para('第一章（三号，黑体）', sz=24) + \
              _para('1.1（四号，黑体）', sz=24) + \
              _para('正文（小四号，宋体）', sz=24)
        _make_docx(tmp_path / 'tmpl.docx', xml)

        rule_map = extract_rules(tmp_path / 'tmpl.docx')

        assert 'chapter' in rule_map
        assert rule_map['chapter']['cn'].font_size == 32   # 三号
        assert rule_map['chapter']['cn'].bold is True

        assert 'section' in rule_map
        assert rule_map['section']['cn'].font_size == 28   # 四号
        assert rule_map['section']['cn'].bold is True

        assert 'body' in rule_map
        assert rule_map['body']['cn'].font_size == 24      # 小四号
        assert rule_map['body']['cn'].bold is False


class TestApplyFormatting:
    """apply_formatting — applying rule_map to target (v5.1: 3 args)."""

    def test_formats_entire_document(self, tmp_path):
        from generator.formatter import extract_rules, apply_formatting

        tmpl = tmp_path / 'tmpl.docx'
        # Template must include rules for every category the target will use
        _make_docx(tmpl,
            _para('摘要（三号，黑体）', sz=24) +
            _para('第一章（三号，黑体）', sz=24) +
            _para('正文（小四号，宋体）', sz=24))
        rule_map = extract_rules(tmpl)

        # Target: cover (skip) + abstract heading + body paragraph
        target = tmp_path / 'target.docx'
        xml = _para('封面', sz=44, bold=True, ea='黑体') + \
              f'<w:p><w:pPr><w:sectPr/></w:pPr></w:p>' + \
              _para('摘要', sz=32, bold=True) + \
              _para('正文内容', sz=24)
        _make_docx(target, xml)

        out = tmp_path / 'out.docx'
        apply_formatting(target, rule_map, out)

        with zipfile.ZipFile(out, 'r') as z:
            doc = etree.parse(z.open('word/document.xml'))
        ps = doc.findall(f'.//{{{W}}}p')
        # Cover should be preserved
        cover_font = ''
        abstract_font = ''
        body_font = ''
        for p in ps:
            t = ''.join(t.text or '' for t in p.findall(f'.//{{{W}}}t'))
            rPr = p.find(f'.//{{{W}}}rPr')
            if rPr is not None:
                rf = rPr.find(f'{{{W}}}rFonts')
                ea = rf.get(f'{{{W}}}eastAsia') if rf is not None else ''
            else:
                ea = ''
            if '封面' in t:
                cover_font = ea
            elif '摘要' in t and '正文' not in t:
                abstract_font = ea
            elif '正文内容' in t:
                body_font = ea

        assert cover_font == '黑体'  # preserved original cover font
        assert abstract_font == '黑体'  # formatted by abstract rule (三号，黑体)
        assert body_font == '宋体'  # formatted by body rule (小四号，宋体)
