"""Unit tests for template_analyzer.py — docx template analysis pipeline."""
import tempfile
import zipfile
from pathlib import Path
from collections import OrderedDict

import pytest
import yaml
from lxml import etree

from generator.template_analyzer import (
    # Pure helpers
    twips_to_pt,
    twips_to_cm,
    normalize_spaces,
    # Data models
    OoxmlSection,
    FontInfo,
    PageInfo,
    SectionClassification,
    # Core classes
    DocxReader,
    SectionClassifier,
    FontExtractor,
    PageAnalyzer,
    TemplateGenerator,
    # Public API
    analyze_template,
    analyze_and_save,
    _custom_yaml_dump,
)

# ═══════════════════════════════════════════════════════
# Namespace
# ═══════════════════════════════════════════════════════
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


# ── Helper: build a minimal docx in-memory ────────────────────
def _make_minimal_docx(body_elements: list[etree._Element]) -> bytes:
    """Wrap a list of OOXML body children into a valid minimal .docx zip."""
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}">'
        '<w:body>'
        + ''.join(etree.tostring(e, encoding='unicode') for e in body_elements) +
        '</w:body>'
        '</w:document>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '</Relationships>'
    )

    buf = tempfile.SpooledTemporaryFile()
    with zipfile.ZipFile(buf, 'w') as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', document_xml)
    buf.seek(0)
    return buf.read()


def _write_docx(body_elements: list[etree._Element], dir_path: Path, name: str = 'test.docx') -> Path:
    """Write a minimal docx to disk and return its path."""
    docx_bytes = _make_minimal_docx(body_elements)
    p = dir_path / name
    p.write_bytes(docx_bytes)
    return p


def _rPr(font='宋体', sz=None, bold=False, italic=False, eastAsia=None):
    """Build a ``w:rPr`` element."""
    rPr = etree.Element(f'{{{W}}}rPr')
    rFonts = etree.SubElement(rPr, f'{{{W}}}rFonts')
    rFonts.set(f'{{{W}}}ascii', font)
    rFonts.set(f'{{{W}}}hAnsi', font)
    if eastAsia:
        rFonts.set(f'{{{W}}}eastAsia', eastAsia)
    if sz is not None:
        sz_el = etree.SubElement(rPr, f'{{{W}}}sz')
        sz_el.set(f'{{{W}}}val', str(sz))
    if bold:
        etree.SubElement(rPr, f'{{{W}}}b')
    if italic:
        etree.SubElement(rPr, f'{{{W}}}i')
    return rPr


def _para_with_runs(runs=None, text=''):
    """Build a ``w:p`` element with explicit run elements."""
    p = etree.Element(f'{{{W}}}p')
    if runs:
        for rPr_el in runs:
            r = etree.SubElement(p, f'{{{W}}}r')
            r.append(rPr_el)
            t = etree.SubElement(r, f'{{{W}}}t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = text
    elif text:
        r = etree.SubElement(p, f'{{{W}}}r')
        t = etree.SubElement(r, f'{{{W}}}t')
        t.text = text
    return p


def _para(text=''):
    """Build a simple ``w:p`` with plain text."""
    p = etree.Element(f'{{{W}}}p')
    if text:
        r = etree.SubElement(p, f'{{{W}}}r')
        t = etree.SubElement(r, f'{{{W}}}t')
        t.text = text
    return p


def _sectPr_element(margins=None, page_size=None, even_and_odd_headers=False, title_pg=False):
    """Build a ``w:sectPr`` element."""
    sp = etree.Element(f'{{{W}}}sectPr')
    if page_size:
        pgSz = etree.SubElement(sp, f'{{{W}}}pgSz')
        pgSz.set(f'{{{W}}}w', str(page_size[0]))
        pgSz.set(f'{{{W}}}h', str(page_size[1]))
    if margins:
        pgMar = etree.SubElement(sp, f'{{{W}}}pgMar')
        for attr, val in margins.items():
            pgMar.set(f'{{{W}}}{attr}', str(val))
    if even_and_odd_headers:
        etree.SubElement(sp, f'{{{W}}}evenAndOddHeaders')
    if title_pg:
        etree.SubElement(sp, f'{{{W}}}titlePg')
    return sp


def _make_paragraph_xml(text, font='宋体', sz=24, bold=False):
    """Build a paragraph with run properties as real OOXML XML bytes."""
    rPr_el = _rPr(font=font, sz=sz, bold=bold)
    r = etree.Element(f'{{{W}}}r')
    r.append(rPr_el)
    t = etree.SubElement(r, f'{{{W}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    p = etree.Element(f'{{{W}}}p')
    p.append(r)
    return p


# ═══════════════════════════════════════════════════════
# 1. Pure helpers
# ═══════════════════════════════════════════════════════

class TestTwipsToPt:
    """twips_to_pt — font-size conversion (w:sz half-points → pt)."""

    def test_24_half_pt_is_12_pt(self):
        assert twips_to_pt(24) == 12.0

    def test_28_half_pt_is_14_pt(self):
        assert twips_to_pt(28) == 14.0

    def test_32_half_pt_is_16_pt(self):
        assert twips_to_pt(32) == 16.0

    def test_zero(self):
        assert twips_to_pt(0) == 0.0

    def test_large_value(self):
        assert twips_to_pt(72) == 36.0


class TestTwipsToCm:
    """twips_to_cm — page-margin conversion."""

    def test_1440_twips_is_about_2_54_cm(self):
        val = twips_to_cm(1440)
        assert pytest.approx(val, 0.05) == 2.54

    def test_standard_a4_margin(self):
        val = twips_to_cm(1800)
        assert pytest.approx(val, 0.1) == 3.17

    def test_zero(self):
        assert twips_to_cm(0) == 0.0


class TestNormalizeSpaces:
    """normalize_spaces — whitespace collapsing."""

    def test_multiple_spaces(self):
        assert normalize_spaces('  hello   world  ') == 'hello world'

    def test_newlines_and_tabs(self):
        assert normalize_spaces('line1\nline2\t\tline3') == 'line1 line2 line3'

    def test_chinese_text_with_spaces(self):
        assert normalize_spaces('  摘要  ： 测试  ') == '摘要 ： 测试'

    def test_empty_string(self):
        assert normalize_spaces('') == ''

    def test_only_whitespace(self):
        assert normalize_spaces('   \n\t  ') == ''


# ═══════════════════════════════════════════════════════
# 2. Font extraction
# ═══════════════════════════════════════════════════════

class TestFontExtraction:
    """FontExtractor — extracting font name, size, bold from OOXML sections."""

    def test_font_name_and_size(self):
        """Extract font name (ascii) and size (half-points → pt)."""
        p = _make_paragraph_xml('测试', font='宋体', sz=24)
        section = OoxmlSection(index=0, paragraphs=[p])
        extractor = FontExtractor([section])
        fonts = extractor.extract_all()
        assert 'body' in fonts
        assert fonts['body'].font_name == 'Times New Roman'
        assert fonts['body'].font_name_east_asia == '宋体'
        assert fonts['body'].size_pt == 12.0

    def test_bold_detection(self):
        p = _make_paragraph_xml('标题', font='黑体', sz=32, bold=True)
        section = OoxmlSection(index=0, paragraphs=[p])
        extractor = FontExtractor([section])
        fonts = extractor.extract_all()
        # FontExtractor may classify large bold fonts as 'heading' not 'body'
        target = fonts.get('heading', fonts.get('body'))
        assert target is not None
        assert target.bold is True

    def test_not_bold_when_absent(self):
        p = _make_paragraph_xml('正文', font='宋体', sz=24, bold=False)
        section = OoxmlSection(index=0, paragraphs=[p])
        extractor = FontExtractor([section])
        fonts = extractor.extract_all()
        assert fonts['body'].bold is False

    def test_east_asian_font(self):
        rPr = _rPr(font='Times New Roman', eastAsia='黑体', sz=28)
        p = _para_with_runs(runs=[rPr], text='中文标题')
        section = OoxmlSection(index=0, paragraphs=[p])
        extractor = FontExtractor([section])
        fonts = extractor.extract_all()
        info = fonts.get('heading', fonts.get('body'))
        assert info is not None
        # CJK text prefers eastAsia font
        assert info.font_name_east_asia == '黑体'
        assert info.font_name in ('黑体', 'Times New Roman')

    def test_size_14pt_from_half_pts_28(self):
        p = _make_paragraph_xml('标题', font='黑体', sz=28)
        section = OoxmlSection(index=0, paragraphs=[p])
        extractor = FontExtractor([section])
        fonts = extractor.extract_all()
        info = fonts.get('heading', fonts.get('body'))
        assert info is not None
        assert info.size_pt in (14.0, 16.0)

    def test_empty_section_skipped(self):
        p = _para(text='')
        section = OoxmlSection(index=0, paragraphs=[p])
        extractor = FontExtractor([section])
        fonts = extractor.extract_all()
        # Empty section returns defaults (always provides body fallback)
        assert 'body' in fonts


# ═══════════════════════════════════════════════════════
# 3. Page setup extraction
# ═══════════════════════════════════════════════════════

class TestPageSetupExtraction:
    """PageAnalyzer — extracting page size and margins."""

    def test_a4_with_standard_margins(self):
        sp = _sectPr_element(
            page_size=(11906, 16838),
            margins={'top': 1440, 'bottom': 1440, 'left': 1800, 'right': 1800},
        )
        section = OoxmlSection(index=0, paragraphs=[_para('content')], sect_pr=sp)
        analyzer = PageAnalyzer([section])
        info = analyzer.extract()

        assert info.width_twips == 11906
        assert info.height_twips == 16838
        assert info.margin_top_twips == 1440
        assert info.margin_left_twips == 1800

    def test_page_size_name_is_a4(self):
        sp = _sectPr_element(page_size=(11906, 16838))
        section = OoxmlSection(index=0, paragraphs=[_para('content')], sect_pr=sp)
        analyzer = PageAnalyzer([section])
        info = analyzer.extract()
        assert info.page_size_name == 'A4'

    def test_margins_to_dict_cm(self):
        sp = _sectPr_element(margins={'top': 1440, 'bottom': 1440, 'left': 1800, 'right': 1800})
        section = OoxmlSection(index=0, paragraphs=[_para('content')], sect_pr=sp)
        analyzer = PageAnalyzer([section])
        info = analyzer.extract()
        d = info.to_dict()
        assert 'margins' in d
        assert pytest.approx(d['margins']['top'], 0.05) == 2.54
        assert pytest.approx(d['margins']['left'], 0.1) == 3.18

    def test_custom_margins(self):
        sp = _sectPr_element(margins={'top': 1134, 'bottom': 1134, 'left': 1701, 'right': 1701})
        section = OoxmlSection(index=0, paragraphs=[_para('content')], sect_pr=sp)
        analyzer = PageAnalyzer([section])
        info = analyzer.extract()
        d = info.to_dict()
        assert pytest.approx(d['margins']['top'], 0.05) == 2.0
        assert pytest.approx(d['margins']['left'], 0.05) == 3.0

    def test_empty_section_returns_default(self):
        section = OoxmlSection(index=0, paragraphs=[_para('content')], sect_pr=None)
        analyzer = PageAnalyzer([section])
        info = analyzer.extract()
        assert info.width_twips is None
        assert info.margin_top_twips is None

    def test_even_and_odd_headers_detected(self):
        """PageAnalyzer should detect w:evenAndOddHeaders in sectPr."""
        sp = _sectPr_element(even_and_odd_headers=True)
        section = OoxmlSection(index=0, paragraphs=[_para('content')], sect_pr=sp)
        analyzer = PageAnalyzer([section])
        info = analyzer.extract()
        assert info.even_and_odd_headers is True

    def test_even_odd_in_to_dict(self):
        """PageInfo.to_dict() should include even_odd when flag is set."""
        info = PageInfo(even_and_odd_headers=True)
        d = info.to_dict()
        assert 'even_odd' in d
        assert d['even_odd'] is True

    def test_even_odd_not_in_dict_when_false(self):
        """PageInfo.to_dict() should NOT include even_odd when flag is false."""
        info = PageInfo(even_and_odd_headers=False)
        d = info.to_dict()
        assert 'even_odd' not in d


# ═══════════════════════════════════════════════════════
# 4. Section classification
# ═══════════════════════════════════════════════════════

class TestSectionClassification:
    """SectionClassifier — classifying sections by text content."""

    def _make_section(self, text, index=0):
        """Helper to create an OoxmlSection with given text."""
        p = _para(text=text)
        return OoxmlSection(index=index, paragraphs=[p])

    def test_classify_cover(self):
        section = self._make_section('毕业设计（论文） 封面 题目：xxx')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'cover'

    def test_classify_statement(self):
        section = self._make_section('郑重声明：本文为原创...')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'statement'

    def test_classify_abstract(self):
        section = self._make_section('摘要')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'abstract'

    def test_classify_abstract_en(self):
        section = self._make_section('ABSTRACT')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'abstract_en'

    def test_classify_toc(self):
        section = self._make_section('目  录')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'toc'

    def test_classify_body(self):
        section = self._make_section('第1章 绪论')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'body'

    def test_classify_references(self):
        section = self._make_section('参考文献')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'references'

    def test_classify_conclusion(self):
        section = self._make_section('结论与展望')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'conclusion'

    def test_classify_acknowledgments(self):
        section = self._make_section('致  谢')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        # "致  谢" normalizes to "致 谢" which still matches the pattern r'致谢'
        assert results[0].section_type == 'acknowledgments'

    def test_classify_appendix(self):
        section = self._make_section('附录 A 数据表格')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type == 'appendix'

    def test_unknown_falls_back_to_body(self):
        section = self._make_section('一些无法归类的内容')
        classifier = SectionClassifier([section])
        results = classifier.classify_all()
        assert results[0].section_type in ('body', 'cover')  # may default to cover if idx==0

    def test_multiple_sections(self):
        texts = [
            '毕业设计（论文）',
            '郑重声明',
            '摘要',
            'ABSTRACT',
            '目 录',
            '第1章 绪论',
            '第2章 相关工作',
            '参考文献',
            '致谢',
        ]
        sections = [self._make_section(t, i) for i, t in enumerate(texts)]
        classifier = SectionClassifier(sections)
        results = classifier.classify_all()
        types = [r.section_type for r in results]
        # Allow some flexibility in first/last defaults
        assert types[0] in ('cover', 'body')
        assert types[1] == 'statement'
        assert types[2] == 'abstract'
        assert types[3] == 'abstract_en'
        assert types[4] == 'toc'
        assert types[5] == 'body'
        assert types[6] == 'body'
        assert types[7] == 'references'
        assert types[8] == 'acknowledgments'

    def test_section_keywords_cover_all_types(self):
        """Ensure known section types match expectations."""
        expected = {
            'cover', 'cover_info', 'statement',
            'abstract', 'abstract_en', 'toc',
            'body', 'conclusion', 'references',
            'acknowledgments', 'appendix', 'empty',
        }
        assert set(SectionClassifier.SECTION_TYPES) == expected


# ═══════════════════════════════════════════════════════
# 5. YAML generation
# ═══════════════════════════════════════════════════════

class TestYamlGeneration:
    """TemplateGenerator — producing valid thesis-format YAML."""

    SAMPLE_SECTIONS_DATA = [
        ('cover', 'cover', 0.9, '', None),
        ('abstract', 'abstract_cn', 0.9, 'roman_upper', 1),
        ('body', 'body', 0.85, 'arabic', 1),
        ('references', 'references', 0.9, 'arabic', None),
    ]

    def _make_sections_and_classifications(self):
        """Create OoxmlSection list and matching SectionClassification list."""
        texts = ['毕业设计', '摘要', '第1章 绪论', '参考文献']
        sections = []
        for i, t in enumerate(texts):
            p = _para(text=t)
            sp = _sectPr_element(
                page_size=(11906, 16838),
                margins={'top': 1440, 'bottom': 1440, 'left': 1800, 'right': 1800},
            )
            sections.append(OoxmlSection(index=i, paragraphs=[p], sect_pr=sp))

        classifications = [
            SectionClassification(section_type='cover', confidence=0.9, evidence=['pattern: cover=0.90']),
            SectionClassification(section_type='abstract', confidence=0.9, evidence=['pattern: abstract=0.90'],
                                  page_number_format='roman_upper', page_number_start=1),
            SectionClassification(section_type='body', confidence=0.85, evidence=['numbering: chapter: "第1章"'],
                                  page_number_format='arabic', page_number_start=1),
            SectionClassification(section_type='references', confidence=0.9, evidence=['pattern: references=0.90'],
                                  page_number_format='arabic', page_number_start=None),
        ]
        return sections, classifications

    def test_generates_valid_structure(self):
        sections, classifications = self._make_sections_and_classifications()
        fonts = {
            'body': FontInfo(font_name='宋体', size_pt=12.0),
            'heading': FontInfo(font_name='黑体', size_pt=16.0, bold=True),
        }
        page_info = PageInfo(
            width_twips=11906, height_twips=16838,
            margin_top_twips=1440, margin_bottom_twips=1440,
            margin_left_twips=1800, margin_right_twips=1800,
        )
        hf_configs = [
            {'header': '', 'footer': '', 'page_number_format': '', 'page_number_start': None, 'different_first_page': False},
            {'header': '摘要', 'footer': '', 'page_number_format': 'roman_upper', 'page_number_start': 1, 'different_first_page': False},
            {'header': '{chapter_title}', 'footer': 'centered', 'page_number_format': 'arabic', 'page_number_start': 1, 'different_first_page': False},
            {'header': '参考文献', 'footer': 'centered', 'page_number_format': 'arabic', 'page_number_start': None, 'different_first_page': False},
        ]

        gen = TemplateGenerator(sections, classifications, fonts, page_info, hf_configs)
        result = gen.generate()

        assert isinstance(result, OrderedDict)
        assert 'name' in result
        assert 'page' in result
        assert 'fonts' in result
        assert 'sections' in result

    def test_page_contains_a4_size(self):
        sections, classifications = self._make_sections_and_classifications()
        fonts = {'body': FontInfo(font_name='宋体', size_pt=12.0)}
        page_info = PageInfo(width_twips=11906, height_twips=16838)
        hf_configs = [{} for _ in sections]

        gen = TemplateGenerator(sections, classifications, fonts, page_info, hf_configs)
        result = gen.generate()
        assert result['page']['size'] == 'A4'

    def test_yaml_output_is_parseable(self):
        sections, classifications = self._make_sections_and_classifications()
        fonts = {
            'body': FontInfo(font_name='宋体', size_pt=12.0),
            'heading': FontInfo(font_name='黑体', size_pt=16.0, bold=True),
        }
        page_info = PageInfo(
            width_twips=11906, height_twips=16838,
            margin_top_twips=1440, margin_bottom_twips=1440,
            margin_left_twips=1800, margin_right_twips=1800,
        )
        hf_configs = [
            {'header': '', 'footer': '', 'page_number_format': '', 'page_number_start': None, 'different_first_page': False},
            {'header': '摘要', 'footer': '', 'page_number_format': 'roman_upper', 'page_number_start': 1, 'different_first_page': False},
            {'header': '{chapter_title}', 'footer': 'centered', 'page_number_format': 'arabic', 'page_number_start': 1, 'different_first_page': False},
            {'header': '参考文献', 'footer': 'centered', 'page_number_format': 'arabic', 'page_number_start': None, 'different_first_page': False},
        ]

        gen = TemplateGenerator(sections, classifications, fonts, page_info, hf_configs)
        result = gen.generate()
        yaml_str = _custom_yaml_dump(result)

        # Must be parseable
        parsed = yaml.safe_load(yaml_str)
        assert isinstance(parsed, dict)
        assert 'sections' in parsed
        assert 'cover' in parsed['sections']

    def test_body_has_heading_numbering(self):
        sections, classifications = self._make_sections_and_classifications()
        fonts = {'body': FontInfo(font_name='宋体', size_pt=12.0)}
        page_info = PageInfo(width_twips=11906, height_twips=16838)
        hf_configs = [{}, {}, {}, {}]

        gen = TemplateGenerator(sections, classifications, fonts, page_info, hf_configs)
        result = gen.generate()
        # The body section (third in list, second in yaml after cover)
        assert 'body' in result['sections']
        assert result['sections']['body'].get('heading_numbering') is True

    def test_body_has_even_odd_when_flag_set(self):
        """TemplateGenerator should add even_odd for body when hf flag is set."""
        sections, classifications = self._make_sections_and_classifications()
        fonts = {'body': FontInfo(font_name='宋体', size_pt=12.0)}
        page_info = PageInfo(width_twips=11906, height_twips=16838)
        # Set even_and_odd_headers on the body section's hf config
        hf_configs = [
            {},
            {},
            {'even_and_odd_headers': True},
            {},
        ]

        gen = TemplateGenerator(sections, classifications, fonts, page_info, hf_configs)
        result = gen.generate()
        assert 'body' in result['sections']
        assert result['sections']['body'].get('even_odd') is True

    def test_body_no_even_odd_for_non_body_sections(self):
        """even_odd should NOT be added for non-body sections even if flag is set."""
        sections, classifications = self._make_sections_and_classifications()
        fonts = {'body': FontInfo(font_name='宋体', size_pt=12.0)}
        page_info = PageInfo(width_twips=11906, height_twips=16838)
        # Set even_and_odd_headers on abstract (non-body)
        hf_configs = [
            {},
            {'even_and_odd_headers': True},  # abstract, should NOT get even_odd
            {},
            {},
        ]

        gen = TemplateGenerator(sections, classifications, fonts, page_info, hf_configs)
        result = gen.generate()
        assert 'abstract_cn' in result['sections']
        assert 'even_odd' not in result['sections']['abstract_cn']

    def test_cover_has_no_page_number(self):
        sections, classifications = self._make_sections_and_classifications()
        fonts = {'body': FontInfo(font_name='宋体', size_pt=12.0)}
        page_info = PageInfo(width_twips=11906, height_twips=16838)
        hf_configs = [{}, {}, {}, {}]

        gen = TemplateGenerator(sections, classifications, fonts, page_info, hf_configs)
        result = gen.generate()
        assert result['sections']['cover']['page_number'] == 'none'

    def test_analyze_and_save_writes_file(self, tmp_path):
        """End-to-end via analyze_and_save on a synthetic docx."""
        sp = _sectPr_element(
            page_size=(11906, 16838),
            margins={'top': 1440, 'bottom': 1440, 'left': 1800, 'right': 1800},
        )
        p1 = _make_paragraph_xml('毕业设计', font='黑体', sz=32, bold=True)
        p2 = _make_paragraph_xml('摘要', font='宋体', sz=24)
        p3 = _make_paragraph_xml('第1章 绪论', font='宋体', sz=24)
        p4 = _make_paragraph_xml('参考文献', font='宋体', sz=24)

        docx = _write_docx([p1, p2, p3, p4, sp], tmp_path)
        out = tmp_path / 'out.yaml'

        result_path = analyze_and_save(str(docx), str(out))
        assert result_path == out
        assert out.exists()

        content = out.read_text(encoding='utf-8')
        assert 'name:' in content
        assert 'sections:' in content
        # Verify parseable
        parsed = yaml.safe_load(content)
        assert isinstance(parsed, dict)


# ═══════════════════════════════════════════════════════
# 6. End-to-end: analyze_template
# ═══════════════════════════════════════════════════════

class TestAnalyzeTemplateE2E:
    """End-to-end pipeline tests using analyze_template()."""

    def test_analyze_mvp_docx(self):
        """Run the full pipeline against mvp_test.docx if it exists."""
        project_root = Path(__file__).resolve().parent.parent
        mvp = project_root / 'output' / 'mvp_test.docx'
        if not mvp.exists():
            pytest.skip(f'{mvp} not found — skipping end-to-end test')

        result = analyze_template(str(mvp))
        assert isinstance(result, OrderedDict)
        assert 'sections' in result
        assert 'fonts' in result
        assert 'page' in result

        # YAML output must be parseable
        yaml_str = _custom_yaml_dump(result)
        parsed = yaml.safe_load(yaml_str)
        assert isinstance(parsed, dict)

    def test_analyze_minimal_docx(self, tmp_path):
        """Full pipeline on a synthetic multi-section docx."""
        sp = _sectPr_element(
            page_size=(11906, 16838),
            margins={'top': 1440, 'bottom': 1440, 'left': 1800, 'right': 1800},
        )
        p1 = _make_paragraph_xml('毕业设计（论文）', font='黑体', sz=32, bold=True)

        # Section break para
        p2 = _make_paragraph_xml('郑重声明', font='宋体', sz=24)
        pPr2 = etree.SubElement(p2, f'{{{W}}}pPr')
        etree.SubElement(pPr2, f'{{{W}}}sectPr')

        p3 = _make_paragraph_xml('摘要', font='宋体', sz=24)
        pPr3 = etree.SubElement(p3, f'{{{W}}}pPr')
        etree.SubElement(pPr3, f'{{{W}}}sectPr')

        p4 = _make_paragraph_xml('第1章 绪论', font='宋体', sz=24)

        docx = _write_docx([p1, p2, p3, p4, sp], tmp_path)

        result = analyze_template(str(docx))
        assert isinstance(result, OrderedDict)
        assert len(result.get('sections', {})) >= 2

        # YAML must be parseable
        yaml_str = _custom_yaml_dump(result)
        yaml.safe_load(yaml_str)

    def test_missing_file_raises(self):
        with pytest.raises(FileNotFoundError):
            analyze_template('nonexistent.docx')
