"""
DocMind 2.0 — 模板分析器
从 .docx 模板中提取节结构、字体、页面设置、页眉页脚，生成 YAML 模板。

用法：
    python -m generator.template_analyzer <input.docx> [-o output.yaml]

也可编程调用：
    from generator.template_analyzer import analyze_template

    result = analyze_template("path/to/template.docx")
    # result 是 dict，等同于生成的 YAML 结构
"""

import re
import yaml
import zipfile
import argparse
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional
from collections import OrderedDict

from lxml import etree

# ── OOXML namespaces ──
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def _ns(tag: str) -> str:
    """Resolve namespace prefix to full Clark notation, e.g. 'w:t' → '{...}t'."""
    if ':' in tag:
        prefix, local = tag.split(':', 1)
        uri = NSMAP.get(prefix, '')
        return f'{{{uri}}}{local}'
    return tag


# ═══════════════════════════════════════════════════════════════════
# Pure conversion helpers
# ═══════════════════════════════════════════════════════════════════

TWIPS_PER_CM = 567       # 1440 twips/inch ÷ 2.54 cm/inch ≈ 567
HALF_PTS_PER_PT = 2       # w:sz is stored in half-points


def twips_to_pt(half_pts: float) -> float:
    """Convert half-points (w:sz) to points.

    In OOXML, font sizes in ``w:sz`` are stored in **half-points**
    (1/144 inch each).  This function divides by 2.
    """
    return half_pts / HALF_PTS_PER_PT


def twips_to_cm(twips: float) -> float:
    """Convert twips to centimeters.

    1 inch = 1440 twips = 2.54 cm  →  1 cm ≈ 567 twips.
    """
    return twips / TWIPS_PER_CM


def normalize_spaces(text: str) -> str:
    """Collapse all whitespace runs into a single space and strip.

    >>> normalize_spaces('  hello   world\\n\\t  ')
    'hello world'
    """
    return re.sub(r'\s+', ' ', text).strip()


# ═══════════════════════════════════════════════════════════════════
# Data models
# ═══════════════════════════════════════════════════════════════════

@dataclass
class OoxmlSection:
    """Represents one document section extracted from OOXML."""
    index: int                                    # 0-based section index
    paragraphs: list = field(default_factory=list) # list of paragraph elements
    sect_pr: Optional[etree._Element] = None       # the w:sectPr for this section
    start_para_idx: int = 0                        # first paragraph index (in document)

    def get_text(self, normalize: bool = True) -> str:
        """Extract all text from this section's paragraphs."""
        texts = []
        for p in self.paragraphs:
            t_nodes = p.findall(f'.//{_ns("w:t")}')
            for t in t_nodes:
                if t.text:
                    texts.append(t.text)
        result = ''.join(texts)
        if normalize:
            result = re.sub(r'[\s\u3000]+', '', result)
        return result

    def get_first_nonempty_paragraph(self) -> Optional[etree._Element]:
        """Return the first paragraph element that contains text."""
        for p in self.paragraphs:
            t_nodes = p.findall(f'.//{_ns("w:t")}')
            has_text = any(t.text and t.text.strip() for t in t_nodes)
            if has_text:
                return p
        return None

    def __repr__(self):
        return f'<OoxmlSection idx={self.index} paras={len(self.paragraphs)}>'


@dataclass
class FontInfo:
    """Extracted font information for a section."""
    font_name: str = ""
    font_name_east_asia: str = ""
    size_pt: Optional[float] = None     # points
    bold: bool = False
    italic: bool = False
    alignment: str = ""                 # left|center|right|both|distribute
    spacing_before: Optional[float] = None  # twips
    spacing_after: Optional[float] = None   # twips
    line_spacing: Optional[float] = None    # 240=单倍行距, 360=1.5倍
    line_spacing_rule: str = ""             # auto|exact|atLeast

    def to_dict(self) -> dict:
        d = {}
        if self.font_name:
            d['name'] = self.font_name
        if self.font_name_east_asia and self.font_name_east_asia != self.font_name:
            d['name_east_asia'] = self.font_name_east_asia
        if self.size_pt is not None:
            d['size'] = self.size_pt
        if self.bold:
            d['bold'] = True
        if self.italic:
            d['italic'] = True
        return d


@dataclass
class PageInfo:
    """Extracted page settings."""
    width_twips: Optional[int] = None
    height_twips: Optional[int] = None
    margin_top_twips: Optional[int] = None
    margin_bottom_twips: Optional[int] = None
    margin_left_twips: Optional[int] = None
    margin_right_twips: Optional[int] = None
    even_and_odd_headers: bool = False
    title_page: bool = False

    @property
    def page_size_name(self) -> str:
        """Return common paper size name (A4, A3, etc.) or custom."""
        if self.width_twips and self.height_twips:
            w_cm = self.width_twips / 567.0
            h_cm = self.height_twips / 567.0
            if abs(w_cm - 21.0) < 0.5 and abs(h_cm - 29.7) < 0.5:
                return 'A4'
            if abs(w_cm - 29.7) < 0.5 and abs(h_cm - 21.0) < 0.5:
                return 'A4'  # landscape
            if abs(w_cm - 14.8) < 0.5 and abs(h_cm - 21.0) < 0.5:
                return 'A5'
            return f'{w_cm:.1f}x{h_cm:.1f}_cm'
        return 'A4'

    def to_dict(self) -> dict:
        """Convert to YAML-ready dict."""
        d = {'size': self.page_size_name}
        margins = {}
        if self.margin_top_twips is not None:
            margins['top'] = round(self.margin_top_twips / 567.0, 2)
        if self.margin_bottom_twips is not None:
            margins['bottom'] = round(self.margin_bottom_twips / 567.0, 2)
        if self.margin_left_twips is not None:
            margins['left'] = round(self.margin_left_twips / 567.0, 2)
        if self.margin_right_twips is not None:
            margins['right'] = round(self.margin_right_twips / 567.0, 2)
        if margins:
            d['margins'] = margins
        if self.even_and_odd_headers:
            d['even_odd'] = True
        return d


@dataclass
class SectionClassification:
    """Result of section classification."""
    section_type: str          # one of the known types
    confidence: float          # 0.0-1.0
    evidence: list = field(default_factory=list)  # human-readable reasons
    page_number_format: str = ""    # decimal|upperRoman|lowerRoman|none
    page_number_start: Optional[int] = None


# ═══════════════════════════════════════════════════════════════════
# 1. OOXML Reader
# ═══════════════════════════════════════════════════════════════════

class DocxReader:
    """Reads and parses a .docx file into structured OOXML data."""

    def __init__(self, docx_path: str | Path):
        self.path = Path(docx_path)
        if not self.path.exists():
            raise FileNotFoundError(f"File not found: {self.path}")
        self._zf: Optional[zipfile.ZipFile] = None
        self._document_xml: Optional[etree._Element] = None
        self._relationships: dict = {}
        self._sections: list[OoxmlSection] = []

    def __enter__(self):
        self._zf = zipfile.ZipFile(self.path, 'r')
        return self

    def __exit__(self, *args):
        if self._zf:
            self._zf.close()

    def _read_xml(self, path_in_zip: str) -> etree._Element:
        """Read an XML file from the ZIP archive."""
        if self._zf is None:
            raise RuntimeError("DocxReader must be used as context manager")
        try:
            with self._zf.open(path_in_zip) as f:
                return etree.parse(f).getroot()
        except KeyError:
            raise FileNotFoundError(f"'{path_in_zip}' not found in {self.path.name}")

    def _load_relationships(self):
        """Parse word/_rels/document.xml.rels to map rId → file path."""
        try:
            rels_xml = self._read_xml('word/_rels/document.xml.rels')
        except FileNotFoundError:
            return {}
        rels = {}
        for rel in rels_xml:
            r_id = rel.get('Id', '')
            target = rel.get('Target', '')
            rel_type = rel.get('Type', '')
            # Resolve relative paths: target may be "header1.xml" or "../media/image1.png"
            if target.startswith('..'):
                # Not handling external references for now
                continue
            full_target = 'word/' + target
            rels[r_id] = {'target': full_target, 'type': rel_type}
        return rels

    def parse(self):
        """Parse the document and populate sections."""
        if self._zf is None:
            raise RuntimeError("DocxReader must be used as context manager")

        self._document_xml = self._read_xml('word/document.xml')
        self._relationships = self._load_relationships()

        body = self._document_xml.find(_ns('w:body'))
        if body is None:
            raise ValueError("No w:body found in document.xml")

        # Collect children: paragraphs, section properties, tables
        children = list(body)

        # Find all section properties (w:sectPr) — they mark section boundaries
        # sectPr can be:
        #   1. Last child of w:body (final section)
        #   2. Inside w:pPr of a paragraph (section break)
        sections = []
        current_paragraphs = []
        section_index = 0
        start_para_idx = 0

        for child in children:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'p':
                # Check if this paragraph has a sectPr in its pPr
                pPr = child.find(_ns('w:pPr'))
                has_sectpr = False
                sectpr_elem = None
                if pPr is not None:
                    sectpr_elem = pPr.find(_ns('w:sectPr'))
                    if sectpr_elem is not None:
                        has_sectpr = True

                current_paragraphs.append(child)

                if has_sectpr:
                    sections.append(OoxmlSection(
                        index=section_index,
                        paragraphs=list(current_paragraphs),
                        sect_pr=sectpr_elem,
                        start_para_idx=start_para_idx,
                    ))
                    section_index += 1
                    start_para_idx += len(current_paragraphs)
                    current_paragraphs = []

            elif tag == 'tbl':
                # Tables are part of the section content
                current_paragraphs.append(child)

            elif tag == 'sectPr':
                # Final section properties (child of body)
                if current_paragraphs:
                    sections.append(OoxmlSection(
                        index=section_index,
                        paragraphs=list(current_paragraphs),
                        sect_pr=child,
                        start_para_idx=start_para_idx,
                    ))
                    section_index += 1
                elif sections:
                    # No new paragraphs; update last section's sectPr
                    sections[-1].sect_pr = child

        self._sections = sections
        return sections

    @property
    def sections(self) -> list[OoxmlSection]:
        return self._sections

    @property
    def relationships(self) -> dict:
        return self._relationships

    def read_header_footer_xml(self, r_id: str) -> Optional[etree._Element]:
        """Read a header or footer XML file by relationship ID."""
        rel = self._relationships.get(r_id)
        if not rel:
            return None
        try:
            return self._read_xml(rel['target'])
        except FileNotFoundError:
            return None


# ═══════════════════════════════════════════════════════════════════
# 2. Section Classifier
# ═══════════════════════════════════════════════════════════════════

class SectionClassifier:
    """Classify OOXML sections using a three-layer strategy:
       1. Page number signal (w:pgNumType / w:start)
       2. Numbering detection (e.g. "第1章", "1 绪论", "1.1")
       3. Text pattern matching + section index context + header text
    """

    # Known section types in order of likelihood for a thesis
    SECTION_TYPES = [
        'cover', 'cover_info', 'statement',
        'abstract', 'abstract_en', 'toc',
        'body', 'conclusion', 'references',
        'acknowledgments', 'appendix', 'empty',
    ]

    # Text patterns for classification (after whitespace normalization)
    # Patterns are matched against the BEGINNING of text for higher accuracy.
    PATTERNS = {
        'cover': [
            # Strong cover indicators (must be near start of text)
            r'毕业设计\s*[（(]\s*论文\s*[）)]',
            r'毕业论文',
            r'学位论文',
            r'学士学位论文',
            r'硕士学位论文',
            r'博士学位论文',
            r'学生姓名[：:]',
            r'学号[：:]',
            r'指导教师[：:]',
            r'导师[：:]',
            r'专业名称[：:]',
        ],
        'cover_info': [
            r'中图分类号', r'UDC', r'密级',
        ],
        'statement': [
            r'郑重声明',
            r'独创性声明',
            r'原创性声明',
            r'学位论文使用授权',
            r'本人郑重声明',
            r'知识产权声明',
            r'保密',
        ],
        'abstract': [
            r'^摘\s*要',          # "摘 要" or "摘要" at start
            r'中文摘要',
        ],
        'abstract_en': [
            r'^ABSTRACT\b',
            r'^Abstract\b',
            r'英文摘要',
        ],
        'toc': [
            r'^目\s*录',           # "目 录" or "目录" at start
            r'^CONTENTS\b',
            r'^Table of Contents',
        ],
        'body': [
            # "第X章" classical Chinese chapter
            r'第[一二三四五六七八九十\d]+章',
            # "1 绪论", "2 核心技术" — number+Chinese char (no space needed)
            r'^[1-9]\d*\s*[\u4e00-\u9fff]',
            # "1.1", "1、xxx"
            r'^[1-9]\d*[.、]',
        ],
        'conclusion': [
            r'总结与展望',
            r'^[1-9]\d*\s*总结',       # "7总结" or "7 总结"
            r'^总结',
            r'^结论',
            r'^结\s*论',
            r'^总\s*结',
            r'^CONCLUSION\b',
            r'^Conclusion\b',
            r'结束语',
        ],
        'references': [
            r'^参考文献',
            r'^REFERENCE',
            r'^Reference',
            r'引用文献',
        ],
        'acknowledgments': [
            r'^致谢',
            r'^ACKNOWLEDGMENT',
            r'^Acknowledgment',
            r'^鸣谢',
        ],
        'appendix': [
            r'^附录',
            r'^APPENDIX',
            r'^Appendix',
            r'附件',
        ],
    }

    # Section types that typically have no page numbers
    NO_PAGE_NUMBER_TYPES = {'cover', 'cover_info', 'statement'}

    # Section types that may start a new page numbering system
    PAGE_NUMBER_RESTART_TYPES = {'abstract', 'abstract_en', 'body'}

    def __init__(self, sections: list[OoxmlSection],
                 section_headers: Optional[dict[int, str]] = None):
        """
        Args:
            sections: List of parsed OOXML sections.
            section_headers: Optional dict mapping section index to its header text
                             (from HeaderFooterExtractor). Used to improve classification.
        """
        self.sections = sections
        self.section_headers = section_headers or {}
        self._classifications: list[SectionClassification] = []

    def classify_all(self) -> list[SectionClassification]:
        """Classify all sections and return results."""
        self._classifications = []
        for i, section in enumerate(self.sections):
            classification = self._classify_section(section, i)
            self._classifications.append(classification)

        # ── Post-classification refinement ──

        # Fix 1: Subclassify front-matter cover sections (cover2/cover3/cover4)
        # Sections after the first cover that matched cover patterns may actually be
        # cover_info or statement based on keywords.
        first_cover_idx = None
        for i, cls in enumerate(self._classifications):
            if cls.section_type == 'cover':
                first_cover_idx = i
                break

        if first_cover_idx is not None:
            for i, cls in enumerate(self._classifications):
                if i == first_cover_idx:
                    continue
                if cls.section_type == 'cover':
                    text = self.sections[i].get_text(normalize=True)
                    # Check for cover_info keywords
                    if any(kw in text for kw in ['学校代码', '学号', 'UDC']):
                        cls.section_type = 'cover_info'
                        cls.evidence.append('post: cover→cover_info (学校代码/学号/UDC)')
                    # Check for statement keywords
                    elif any(kw in text for kw in ['声明', '授权', '原创']):
                        cls.section_type = 'statement'
                        cls.evidence.append('post: cover→statement (声明/授权/原创)')

        # Fix 2: Detect Chinese abstract from section headers (position-independent)
        # Scan all section_headers for headers containing 摘+要, reclassify unclassified
        # or misclassified sections as abstract.
        for idx, header_text in self.section_headers.items():
            if idx >= len(self._classifications):
                continue
            if '摘' in header_text and '要' in header_text:
                cls = self._classifications[idx]
                if cls.section_type not in ('abstract', 'abstract_en'):
                    # Reclassify as Chinese abstract
                    cls.section_type = 'abstract'
                    cls.evidence.append(f'post: header "{header_text}"→abstract_cn')
                    if cls.confidence < 0.8:
                        cls.confidence = 0.8

        return self._classifications

    def _classify_section(self, section: OoxmlSection, idx: int) -> SectionClassification:
        """Apply three-layer strategy to classify a single section."""
        evidence = []
        text = section.get_text(normalize=True)
        header_text = self.section_headers.get(idx, '')

        # ── Layer 1: Page number signal ──
        pn_signal = self._page_number_signal(section)
        if pn_signal:
            evidence.append(f'pn_signal: {pn_signal}')

        # ── Layer 2: Numbering detection ──
        numbering_hit = self._numbering_detection(text)
        if numbering_hit:
            evidence.append(f'numbering: {numbering_hit}')

        # ── Layer 3: Text pattern matching ──
        pattern_matches = self._pattern_match(text)
        for match_type, score in pattern_matches:
            evidence.append(f'pattern: {match_type}={score:.2f}')

        # ── Layer 4: Header text as signal ──
        header_signal = self._header_signal(header_text)
        if header_signal:
            evidence.append(f'header: {header_signal}')

        # ── Handle empty sections ──
        if not text.strip():
            # Empty section: check header for clues
            if 'toc' in (header_signal or '').lower() or '目录' in header_text:
                section_type, confidence = 'toc', 0.7
                evidence.append('empty+toc_header')
            else:
                section_type, confidence = 'empty', 1.0
                evidence = ['empty: no text content']
        else:
            # ── Resolve type ──
            section_type, confidence = self._resolve_type(
                text, pn_signal, numbering_hit, pattern_matches,
                header_signal, idx
            )

        # Extract page number format
        pn_format, pn_start = self._extract_page_number_info(section)

        return SectionClassification(
            section_type=section_type,
            confidence=confidence,
            evidence=evidence,
            page_number_format=pn_format,
            page_number_start=pn_start,
        )

    def _header_signal(self, header_text: str) -> Optional[str]:
        """Detect section type from header text."""
        if not header_text:
            return None
        h = header_text.strip()
        if re.match(r'^摘\s*要$', h):
            return 'abstract'
        if re.match(r'^目\s*录$', h):
            return 'toc'
        if re.match(r'^ABSTRACT$', h, re.IGNORECASE):
            return 'abstract_en'
        if re.match(r'^参考文献$', h):
            return 'references'
        if re.match(r'^致谢$', h):
            return 'acknowledgments'
        if re.match(r'^结论|总结', h):
            return 'conclusion'
        if re.match(r'^附录', h):
            return 'appendix'
        return None

    def _page_number_signal(self, section: OoxmlSection) -> Optional[str]:
        """Detect page number signals from section properties."""
        if section.sect_pr is None:
            return None

        pgNumType = section.sect_pr.find(_ns('w:pgNumType'))
        if pgNumType is None:
            return None

        fmt = pgNumType.get(_ns('w:fmt'), '')
        start = pgNumType.get(_ns('w:start'), '')

        parts = []
        fmt_map = {
            'decimal': 'arabic',
            'upperRoman': 'roman_upper',
            'lowerRoman': 'roman_lower',
            'none': 'none',
        }
        if fmt:
            parts.append(fmt_map.get(fmt, fmt))
        if start:
            parts.append(f'start={start}')

        return ':'.join(parts) if parts else None

    def _numbering_detection(self, text: str) -> Optional[str]:
        """Detect if text begins with chapter/section numbering."""
        # Check for "第X章" pattern
        m = re.match(r'第[一二三四五六七八九十\d]+章', text)
        if m:
            return f'chapter: "{m.group()}"'
        # Check for numbered heading: "1绪论", "1 绪论", "2核心技术"
        m = re.match(r'^([1-9]\d*)\s*[\u4e00-\u9fff]', text)
        if m:
            return f'numbered_ch: "{m.group()}"'
        # Check for sub-numbered: "1.1", "1.1.1"
        m = re.match(r'^([1-9]\d*)[.、]\d+', text)
        if m:
            return f'numbered_sub: "{m.group()}"'
        return None

    def _pattern_match(self, text: str) -> list[tuple[str, float]]:
        """Match text against known section patterns.

        Returns list of (section_type, score) tuples where score is 0.0-1.0.
        Uses position-weighted scoring: matches near text start score higher.
        """
        results = []
        for sec_type, patterns in self.PATTERNS.items():
            best_score = 0.0
            for pat in patterns:
                m = re.search(pat, text, re.IGNORECASE)
                if m:
                    # Score based on match position: earlier = higher
                    pos_ratio = 1.0 - min(m.start() / max(len(text), 1), 0.95)
                    # Base score 0.3 + up to 0.7 for position
                    score = 0.3 + 0.7 * pos_ratio
                    best_score = max(best_score, score)
            if best_score > 0:
                results.append((sec_type, best_score))

        # Sort by score descending
        results.sort(key=lambda x: x[1], reverse=True)
        return results

    def _resolve_type(self, text: str, pn_signal: Optional[str],
                      numbering_hit: Optional[str],
                      pattern_matches: list[tuple[str, float]],
                      header_signal: Optional[str],
                      idx: int) -> tuple[str, float]:
        """Resolve section type from combined evidence."""

        n_sections = len(self.sections)

        # ── Rule 0: Header signal overrides weak pattern matches ──
        if header_signal:
            top_pattern = pattern_matches[0] if pattern_matches else (None, 0)

            # Header signals for abstract/toc are very reliable
            if header_signal in ('abstract', 'abstract_en', 'toc'):
                if top_pattern[0] in ('abstract', 'abstract_en', 'toc'):
                    return header_signal, 0.9
                # Even without pattern match, these are reliable
                if not numbering_hit:
                    return header_signal, 0.7

            # Header signals for references/conclusion/acknowledgments/appendix:
            # Only trust if NO numbering found (body chapters often have
            # static headers like "参考文献")
            if header_signal in ('references', 'conclusion', 'acknowledgments', 'appendix'):
                if not numbering_hit:
                    if top_pattern[0] == header_signal:
                        return header_signal, 0.9
                    return header_signal, 0.7
                # Numbering found → this is body, not references
                # (fall through to Rule 1)

        # ── Rule 1: numbering detection → body (unless contradicted) ──
        if numbering_hit:
            # Check for contradicting strong abstract/toc/conclusion patterns
            if pattern_matches:
                top_type, top_score = pattern_matches[0]
                if top_score > 0.8 and top_type in ('abstract', 'abstract_en', 'toc',
                                                     'references', 'conclusion',
                                                     'acknowledgments', 'appendix'):
                    return top_type, top_score
            # Check if number is followed by conclusion/references/acknowledgments keyword
            # e.g., "7总结与展望" → conclusion, not body
            if pattern_matches:
                for m_type, m_score in pattern_matches:
                    if m_type in ('conclusion', 'references', 'acknowledgments') and m_score > 0.6:
                        return m_type, m_score
            return 'body', 0.85

        # ── Rule 2: page number reset → likely abstract or body start ──
        if pn_signal and 'start=1' in pn_signal:
            if pattern_matches:
                top_type, top_score = pattern_matches[0]
                if top_score > 0.5:
                    return top_type, top_score
            if len(text) < 200:
                return 'abstract', 0.5
            return 'body', 0.4

        # ── Rule 3: use best pattern match ──
        if pattern_matches:
            top_type, top_score = pattern_matches[0]

            # Special: if best is 'body' pattern but section looks like conclusion
            if top_type == 'body' and top_score < 0.9:
                # Check for conclusion-like content near end
                if idx >= n_sections - 3:
                    for m_type, m_score in pattern_matches:
                        if m_type in ('conclusion', 'references', 'acknowledgments') and m_score > 0.5:
                            return m_type, m_score

            # Abstract + abstract_en in same section
            if len(pattern_matches) > 1:
                second_type, second_score = pattern_matches[1]
                if {top_type, second_type} == {'abstract', 'abstract_en'}:
                    return 'abstract', 0.8

            # Cover + statement co-occurring (first section only)
            if idx == 0 and top_type == 'cover':
                return 'cover', min(top_score + 0.1, 1.0)

            return top_type, top_score

        # ── Rule 4: Section index heuristics ──
        # First section is usually cover
        if idx == 0 and text:
            return 'cover', 0.5

        # Section after cover/abstract/toc is usually body
        if idx >= 1 and idx <= n_sections - 4 and text:
            return 'body', 0.3

        # Last 3 sections are usually conclusion/references/acknowledgments
        if idx >= n_sections - 3 and text:
            remaining = n_sections - idx
            if remaining == 3:
                return 'conclusion', 0.3
            elif remaining == 2:
                return 'references', 0.3
            elif remaining == 1:
                return 'acknowledgments', 0.3

        # Default
        if text:
            return 'body', 0.1
        return 'empty', 0.8

    def _extract_page_number_info(self, section: OoxmlSection) -> tuple[str, Optional[int]]:
        """Extract page number format and start value from sectPr."""
        if section.sect_pr is None:
            return '', None

        pgNumType = section.sect_pr.find(_ns('w:pgNumType'))
        if pgNumType is None:
            return '', None

        fmt = pgNumType.get(_ns('w:fmt'), '')
        start_str = pgNumType.get(_ns('w:start'), '')

        fmt_map = {
            'decimal': 'arabic',
            'upperRoman': 'roman_upper',
            'lowerRoman': 'roman_lower',
            'none': 'none',
        }

        pn_format = fmt_map.get(fmt, fmt) if fmt else ''
        pn_start = int(start_str) if start_str else None

        return pn_format, pn_start

    @property
    def classifications(self) -> list[SectionClassification]:
        return self._classifications


# ═══════════════════════════════════════════════════════════════════
# 3. Font Extractor
# ═══════════════════════════════════════════════════════════════════

class FontExtractor:
    """Extract font information from sections.

    For each section, reads paragraphs and extracts:
    - Font name (w:rFonts → w:ascii, w:hAnsi, w:eastAsia)
    - Font size (w:sz → half-points, convert to points)
    - Bold (w:b), Italic (w:i)
    - Alignment (w:jc)
    - Paragraph spacing (w:spacing)

    When run-level font properties are absent, falls back to paragraph style
    (from word/styles.xml) and document defaults.
    """

    def __init__(self, sections: list[OoxmlSection],
                 styles_xml: Optional[etree._Element] = None):
        self.sections = sections
        self._style_cache: dict[str, dict] = {}
        self._default_fonts: dict = {}
        if styles_xml is not None:
            self._parse_styles(styles_xml)

    def _parse_styles(self, styles_xml: etree._Element):
        """Parse word/styles.xml to cache style→font mappings."""
        # Parse document defaults
        defaults = styles_xml.find(_ns('w:docDefaults'))
        if defaults is not None:
            rPrDefault = defaults.find(_ns('w:rPrDefault'))
            if rPrDefault is not None:
                rPr = rPrDefault.find(_ns('w:rPr'))
                if rPr is not None:
                    self._default_fonts = self._extract_run_props(rPr)

        # Parse individual styles
        for style in styles_xml.findall(_ns('w:style')):
            style_id = style.get(_ns('w:styleId'), '')
            if not style_id:
                continue

            # Collect font info from style's rPr and pPr
            font_data = {}
            rPr = style.find(_ns('w:rPr'))
            if rPr is not None:
                font_data = self._extract_run_props(rPr)
            pPr = style.find(_ns('w:pPr'))
            if pPr is not None:
                jc = pPr.find(_ns('w:jc'))
                if jc is not None:
                    font_data['alignment'] = jc.get(_ns('w:val'), '')

            # Inherit from basedOn style
            basedOn = style.find(_ns('w:basedOn'))
            if basedOn is not None:
                parent_id = basedOn.get(_ns('w:val'), '')
                if parent_id and parent_id in self._style_cache:
                    parent_data = self._style_cache[parent_id]
                    for k, v in parent_data.items():
                        if k not in font_data:
                            font_data[k] = v

            self._style_cache[style_id] = font_data

    def _extract_run_props(self, rPr: etree._Element) -> dict:
        """Extract font properties from a w:rPr element."""
        data = {}
        rFonts = rPr.find(_ns('w:rFonts'))
        if rFonts is not None:
            for attr in ('ascii', 'hAnsi', 'eastAsia'):
                val = rFonts.get(_ns(f'w:{attr}'), '')
                if val:
                    data[f'font_{attr}'] = val
        sz = rPr.find(_ns('w:sz'))
        if sz is not None:
            val = sz.get(_ns('w:val'), '')
            if val:
                data['size_pt'] = int(val) / 2.0
        b = rPr.find(_ns('w:b'))
        if b is not None:
            data['bold'] = True
        i = rPr.find(_ns('w:i'))
        if i is not None:
            data['italic'] = True
        return data

    def extract_all(self) -> dict[str, FontInfo]:
        """Extract font info for different semantic roles.

        Returns dict with keys like 'heading', 'body', 'abstract_title', 'header_footer'.
        Scans all sections and picks the most representative font for each role.
        """
        # Collect font samples from different section types
        heading_samples = []   # large centered bold fonts
        body_samples = []      # body text fonts
        small_samples = []     # header/footer fonts

        for i, section in enumerate(self.sections):
            text = section.get_text(normalize=True)
            if not text.strip():
                continue

            # For heading: use first non-empty paragraph
            first_p = section.get_first_nonempty_paragraph()
            if first_p is not None:
                fi = self._extract_from_paragraph(first_p)
                if fi.size_pt and fi.size_pt >= 14 and fi.bold:
                    heading_samples.append(fi)
                elif fi.size_pt and fi.size_pt <= 10:
                    small_samples.append(fi)

            # For body text: look at paragraph after the heading
            # (second non-empty paragraph, or first paragraph that is not large+centered+bold)
            body_p = self._find_body_paragraph(section)
            if body_p is not None:
                fi = self._extract_from_paragraph(body_p)
                if fi.font_name or fi.size_pt:
                    body_samples.append(fi)

            # Also check last few paragraphs for body font if still missing
            if len(body_samples) == 0 and len(section.paragraphs) >= 3:
                # Try third paragraph
                non_empty = [p for p in section.paragraphs
                           if any(t.text and t.text.strip()
                                  for t in p.findall(f'.//{_ns("w:t")}'))]
                if len(non_empty) >= 3:
                    fi = self._extract_from_paragraph(non_empty[2])
                    if fi.font_name or fi.size_pt:
                        body_samples.append(fi)

        # ── Build result dict ──
        result = OrderedDict()

        # Heading font: prefer the largest one
        if heading_samples:
            best_heading = max(heading_samples,
                              key=lambda f: (f.size_pt or 0, len(f.font_name)))
            result['heading'] = best_heading

        # Abstract title: second heading-like font if different
        if len(heading_samples) >= 2:
            # Check if there's a different heading font (e.g., 黑体 vs 宋体)
            h1 = heading_samples[0]
            for h in heading_samples[1:]:
                if h.font_name != h1.font_name or abs((h.size_pt or 0) - (h1.size_pt or 0)) >= 2:
                    result['abstract_title'] = h
                    break

        # Body font — if identical to heading, fall back to defaults (empty template)
        if body_samples:
            best_body = max(body_samples,
                           key=lambda f: (len(f.font_name), f.size_pt or 0))
            result['body'] = best_body

        # If body font matches heading, it's likely a sparse template — use defaults
        if 'body' in result and 'heading' in result:
            h = result['heading']
            b = result['body']
            if (b.font_name == h.font_name and b.size_pt == h.size_pt and b.bold == h.bold):
                result['body'] = FontInfo(font_name='Times New Roman', size_pt=12,
                                          font_name_east_asia='宋体')

        # Header/footer font
        if small_samples:
            result['header_footer'] = small_samples[0]

        # ── Fill defaults ──
        defaults = {
            'heading': FontInfo(font_name='Times New Roman', size_pt=16, bold=True,
                               font_name_east_asia='黑体'),
            'body': FontInfo(font_name='Times New Roman', size_pt=12,
                            font_name_east_asia='宋体'),
            'header_footer': FontInfo(font_name='Times New Roman', size_pt=9,
                                     font_name_east_asia='宋体'),
        }
        for dk, dv in defaults.items():
            if dk not in result:
                result[dk] = dv

        return result

    def _find_body_paragraph(self, section: OoxmlSection) -> Optional[etree._Element]:
        """Find a body-text paragraph (not a heading) in the section.

        Skips the first paragraph if it's centered+bold+large (heading).
        Tries to find a paragraph with Chinese text for accurate font detection.
        Returns the first suitable paragraph.
        """
        paras = []
        for p in section.paragraphs:
            t_nodes = p.findall(f'.//{_ns("w:t")}')
            if any(t.text and t.text.strip() for t in t_nodes):
                paras.append(p)

        if len(paras) <= 1:
            return None

        first = paras[0]
        fi = self._extract_from_paragraph(first)
        if fi.alignment == 'center' and fi.bold and fi.size_pt and fi.size_pt >= 14:
            # First paragraph is a heading, look for body text
            # Prefer a paragraph with Chinese characters
            for p in paras[1:]:
                t_nodes = p.findall(f'.//{_ns("w:t")}')
                para_text = ''.join(t.text or '' for t in t_nodes)
                if re.search(r'[\u4e00-\u9fff]', para_text):
                    return p
            # Fallback: return second paragraph
            return paras[1] if len(paras) > 1 else None

        return first

    def _infer_font_key(self, section: OoxmlSection, font_info: FontInfo) -> str:
        """Infer the font role key (heading, body, header_footer, etc.)."""
        # Check alignment: centered + large + bold → heading
        if font_info.alignment == 'center' and font_info.size_pt and font_info.size_pt >= 14:
            return 'heading'
        if font_info.alignment == 'center' and font_info.bold:
            return 'heading'
        # Small font → header_footer
        if font_info.size_pt and font_info.size_pt <= 10:
            return 'header_footer'
        return 'body'

    def _extract_from_paragraph(self, p: etree._Element) -> FontInfo:
        """Extract font info from a single paragraph element."""
        info = FontInfo()

        # ── Paragraph properties ──
        pPr = p.find(_ns('w:pPr'))
        pStyle_id = None
        if pPr is not None:
            # Alignment
            jc = pPr.find(_ns('w:jc'))
            if jc is not None:
                jc_val = jc.get(_ns('w:val'), '')
                jc_map = {
                    'left': 'left', 'center': 'center', 'right': 'right',
                    'both': 'both', 'distribute': 'distribute',
                }
                info.alignment = jc_map.get(jc_val, jc_val)

            # Paragraph spacing
            spacing = pPr.find(_ns('w:spacing'))
            if spacing is not None:
                before = spacing.get(_ns('w:before'))
                if before:
                    info.spacing_before = int(before)
                after = spacing.get(_ns('w:after'))
                if after:
                    info.spacing_after = int(after)
                line = spacing.get(_ns('w:line'))
                if line:
                    info.line_spacing = int(line)
                line_rule = spacing.get(_ns('w:lineRule'))
                if line_rule:
                    info.line_spacing_rule = line_rule

            # Paragraph style reference
            pStyle = pPr.find(_ns('w:pStyle'))
            if pStyle is not None:
                pStyle_id = pStyle.get(_ns('w:val'), '')

        # ── Run properties (collect from all runs) ──
        found_font = False
        runs = p.findall(_ns('w:r'))
        for r in runs:
            rPr = r.find(_ns('w:rPr'))
            if rPr is None:
                continue

            # Font name
            rFonts = rPr.find(_ns('w:rFonts'))
            if rFonts is not None:
                ascii_font = rFonts.get(_ns('w:ascii'), '')
                hAnsi_font = rFonts.get(_ns('w:hAnsi'), '')
                ea_font = rFonts.get(_ns('w:eastAsia'), '')
                if ascii_font and not info.font_name:
                    info.font_name = ascii_font
                elif hAnsi_font and not info.font_name:
                    info.font_name = hAnsi_font
                if ea_font and not info.font_name_east_asia:
                    info.font_name_east_asia = ea_font

            # Font size (half-points → points)
            sz = rPr.find(_ns('w:sz'))
            if sz is not None:
                sz_val = sz.get(_ns('w:val'))
                if sz_val and info.size_pt is None:
                    info.size_pt = int(sz_val) / 2.0

            szCs = rPr.find(_ns('w:szCs'))
            if szCs is not None and info.size_pt is None:
                sz_val = szCs.get(_ns('w:val'))
                if sz_val:
                    info.size_pt = int(sz_val) / 2.0

            # Bold
            b = rPr.find(_ns('w:b'))
            if b is not None:
                b_val = b.get(_ns('w:val'), 'true')
                if b_val not in ('false', '0', 'off'):
                    info.bold = True

            # Italic
            i = rPr.find(_ns('w:i'))
            if i is not None:
                i_val = i.get(_ns('w:val'), 'true')
                if i_val not in ('false', '0', 'off'):
                    info.italic = True

            if info.font_name or info.size_pt is not None:
                found_font = True

        # ── Fallback: inherit from paragraph style ──
        if not found_font and pStyle_id and pStyle_id in self._style_cache:
            style_data = self._style_cache[pStyle_id]
            if 'font_ascii' in style_data:
                info.font_name = style_data['font_ascii']
            elif 'font_hAnsi' in style_data:
                info.font_name = style_data['font_hAnsi']
            if 'font_eastAsia' in style_data:
                info.font_name_east_asia = style_data['font_eastAsia']
            if 'size_pt' in style_data and info.size_pt is None:
                info.size_pt = style_data['size_pt']
            if 'bold' in style_data:
                info.bold = style_data['bold']
            if 'alignment' in style_data and not info.alignment:
                info.alignment = style_data['alignment']

        # ── Fallback: use document defaults ──
        if not found_font and self._default_fonts:
            if not info.font_name:
                info.font_name = self._default_fonts.get('font_ascii', '') or \
                                  self._default_fonts.get('font_hAnsi', '')
            if not info.font_name_east_asia:
                info.font_name_east_asia = self._default_fonts.get('font_eastAsia', '')
            if info.size_pt is None:
                info.size_pt = self._default_fonts.get('size_pt')

        return info


# ═══════════════════════════════════════════════════════════════════
# 4. Page Analyzer
# ═══════════════════════════════════════════════════════════════════

class PageAnalyzer:
    """Extract page setup from section properties."""

    def __init__(self, sections: list[OoxmlSection]):
        self.sections = sections

    def extract(self) -> PageInfo:
        """Extract page settings from the first section's sectPr (most representative)."""
        # Use the first section's sectPr, or body-level sectPr
        for section in self.sections:
            if section.sect_pr is not None:
                return self._extract_from_sectpr(section.sect_pr)

        # Fallback: construct default
        return PageInfo()

    def extract_all(self) -> list[PageInfo]:
        """Extract page settings from all sections."""
        result = []
        for section in self.sections:
            if section.sect_pr is not None:
                result.append(self._extract_from_sectpr(section.sect_pr))
            else:
                result.append(PageInfo())
        return result

    def _extract_from_sectpr(self, sectPr: etree._Element) -> PageInfo:
        """Parse page settings from a sectPr element."""
        info = PageInfo()

        # Page size
        pgSz = sectPr.find(_ns('w:pgSz'))
        if pgSz is not None:
            w = pgSz.get(_ns('w:w'))
            h = pgSz.get(_ns('w:h'))
            if w:
                info.width_twips = int(w)
            if h:
                info.height_twips = int(h)

        # Page margins
        pgMar = sectPr.find(_ns('w:pgMar'))
        if pgMar is not None:
            top = pgMar.get(_ns('w:top'))
            bottom = pgMar.get(_ns('w:bottom'))
            left = pgMar.get(_ns('w:left'))
            right = pgMar.get(_ns('w:right'))
            if top:
                info.margin_top_twips = int(top)
            if bottom:
                info.margin_bottom_twips = int(bottom)
            if left:
                info.margin_left_twips = int(left)
            if right:
                info.margin_right_twips = int(right)

        # Even and odd headers
        even_odd = sectPr.find(_ns('w:evenAndOddHeaders'))
        if even_odd is not None:
            info.even_and_odd_headers = True

        # Title page (different first page)
        title_pg = sectPr.find(_ns('w:titlePg'))
        if title_pg is not None:
            info.title_page = True

        return info


# ═══════════════════════════════════════════════════════════════════
# 5. Header/Footer Extractor
# ═══════════════════════════════════════════════════════════════════

class HeaderFooterExtractor:
    """Extract header/footer content from .docx OOXML.

    Reads each section's headerReference/footerReference rIds,
    then parses the actual headerN.xml / footerN.xml files.
    """

    HEADER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
    FOOTER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'

    def __init__(self, reader: DocxReader):
        self.reader = reader

    def extract_section_hf(self, section: OoxmlSection) -> dict:
        """Extract header/footer config for a single section.

        Returns dict with keys: header, footer, page_number_format,
        page_number_start, different_first_page.
        """
        result = {
            'header': '',
            'footer': '',
            'page_number_format': '',
            'page_number_start': None,
            'different_first_page': False,
        }

        if section.sect_pr is None:
            return result

        # ── Header references ──
        header_refs = section.sect_pr.findall(_ns('w:headerReference'))
        for ref in header_refs:
            r_id = ref.get(_ns('r:id'), '')
            hdr_type = ref.get(_ns('w:type'), 'default')  # default, even, first
            header_text = self._read_header_footer_text(r_id)
            if header_text:
                if hdr_type == 'default':
                    result['header'] = header_text
                    # Also check for PAGE field
                    if self._has_page_field(r_id):
                        result['has_page_field'] = True

        # ── Footer references ──
        footer_refs = section.sect_pr.findall(_ns('w:footerReference'))
        for ref in footer_refs:
            r_id = ref.get(_ns('r:id'), '')
            ftr_type = ref.get(_ns('w:type'), 'default')
            footer_text = self._read_header_footer_text(r_id)
            if footer_text:
                if ftr_type == 'default':
                    result['footer'] = footer_text
                    if self._has_page_field(r_id):
                        result['has_page_field'] = True

        # ── Page number info ──
        pgNumType = section.sect_pr.find(_ns('w:pgNumType'))
        if pgNumType is not None:
            fmt = pgNumType.get(_ns('w:fmt'), '')
            start = pgNumType.get(_ns('w:start'), '')
            fmt_map = {
                'decimal': 'arabic',
                'upperRoman': 'roman_upper',
                'lowerRoman': 'roman_lower',
                'none': 'none',
            }
            result['page_number_format'] = fmt_map.get(fmt, fmt)
            if start:
                result['page_number_start'] = int(start)

        # ── Different first page ──
        titlePg = section.sect_pr.find(_ns('w:titlePg'))
        if titlePg is not None:
            result['different_first_page'] = True

        # ── Even and odd headers ──
        even_odd = section.sect_pr.find(_ns('w:evenAndOddHeaders'))
        if even_odd is not None:
            result['even_and_odd_headers'] = True

        return result

    def _read_header_footer_text(self, r_id: str) -> str:
        """Read text content from a header/footer XML file."""
        xml_root = self.reader.read_header_footer_xml(r_id)
        if xml_root is None:
            return ''

        # Collect all w:t text nodes
        texts = []
        for t in xml_root.iter(_ns('w:t')):
            if t.text:
                texts.append(t.text)
        full_text = ''.join(texts)

        # Normalize: remove extra whitespace
        full_text = re.sub(r'\s+', ' ', full_text).strip()
        return full_text

    def _has_page_field(self, r_id: str) -> bool:
        """Check if a header/footer contains a PAGE field code."""
        xml_root = self.reader.read_header_footer_xml(r_id)
        if xml_root is None:
            return False

        # Look for w:fldChar with fldCharType="begin" AND w:instrText with "PAGE"
        has_begin = False
        has_page_instr = False

        for elem in xml_root.iter():
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if tag == 'fldChar':
                if elem.get(_ns('w:fldCharType')) == 'begin':
                    has_begin = True
            if tag == 'instrText':
                if elem.text and 'PAGE' in elem.text.upper():
                    has_page_instr = True

        return has_begin and has_page_instr


# ═══════════════════════════════════════════════════════════════════
# 6. Template YAML Generator
# ═══════════════════════════════════════════════════════════════════

class TemplateGenerator:
    """Generate YAML template from extracted analysis data."""

    # Mapping from section_type to YAML key
    TYPE_TO_YAML_KEY = {
        'cover': 'cover',
        'cover_info': 'cover_info',
        'statement': 'statement',
        'abstract': 'abstract_cn',
        'abstract_en': 'abstract_en',
        'toc': 'toc',
        'body': 'body',
        'conclusion': 'conclusion',
        'references': 'references',
        'acknowledgments': 'acknowledgments',
        'appendix': 'appendix',
        'empty': None,  # skip empty sections
    }

    def __init__(self,
                 sections: list[OoxmlSection],
                 classifications: list[SectionClassification],
                 fonts: dict[str, FontInfo],
                 page_info: PageInfo,
                 hf_configs: list[dict]):
        self.sections = sections
        self.classifications = classifications
        self.fonts = fonts
        self.page_info = page_info
        self.hf_configs = hf_configs

    def generate(self) -> OrderedDict:
        """Generate the complete template YAML structure."""
        template = OrderedDict()

        # ── Name & description ──
        template['name'] = '从模板提取的配置'
        template['description'] = '由 template_analyzer.py 自动分析 .docx 生成'

        # ── Page settings ──
        template['page'] = self.page_info.to_dict()

        # ── Fonts ──
        template['fonts'] = self._build_fonts_section()

        # ── Sections ──
        template['sections'] = self._build_sections_section()

        return template

    def _build_fonts_section(self) -> OrderedDict:
        """Build the fonts section of the YAML."""
        fonts = OrderedDict()
        for key, info in self.fonts.items():
            fonts[key] = info.to_dict()
        # Ensure default keys exist
        defaults = {
            'heading': {'name': '黑体', 'size': 16, 'bold': True},
            'body': {'name': '宋体', 'size': 12},
            'header_footer': {'name': '宋体', 'size': 9},
        }
        for dk, dv in defaults.items():
            if dk not in fonts:
                fonts[dk] = dv
        return fonts

    def _build_sections_section(self) -> OrderedDict:
        """Build the sections section of the YAML.

        Only the FIRST occurrence of each section type is included in the template.
        Body sections are special: we use the first body section but note if the
        header text is static across all body sections.
        """
        sections_yaml = OrderedDict()
        seen_types = {}  # track first occurrence index

        # First pass: collect all classifications per type to detect patterns
        body_headers = []
        for i, (section, cls, hf) in enumerate(
            zip(self.sections, self.classifications, self.hf_configs)
        ):
            if cls.section_type == 'body':
                body_headers.append(hf.get('header', ''))

        # Detect if all body sections share the same static header
        static_body_header = None
        if body_headers:
            unique_headers = set(h for h in body_headers if h)
            if len(unique_headers) == 1:
                static_body_header = list(unique_headers)[0]

        # Second pass: build YAML entries — keep ALL sections, use counter for repeats
        seen_type_counts = {}  # count occurrences per type
        
        for i, (section, cls, hf) in enumerate(
            zip(self.sections, self.classifications, self.hf_configs)
        ):
            yaml_key = self.TYPE_TO_YAML_KEY.get(cls.section_type)
            if yaml_key is None:
                continue
            
            # Allow multiple instances of the same type
            count = seen_type_counts.get(yaml_key, 0)
            seen_type_counts[yaml_key] = count + 1
            if count > 0:
                yaml_key = f"{yaml_key}{count + 1}"  # statement2, statement3, etc.

            section_config = OrderedDict()

            # Header
            header_text = hf.get('header', '')
            if cls.section_type == 'body' and static_body_header:
                # If all body sections share the same header, it's likely a
                # template artifact — suggest using dynamic chapter title
                section_config['header'] = '{chapter_title}'
                # Add comment about detected static header
                sections_yaml[
                    f'# ℹ️ Body header was static: "{static_body_header}" — '
                    f'changed to {{chapter_title}}'
                ] = None
            else:
                section_config['header'] = header_text

            # Footer
            footer_text = hf.get('footer', '')
            has_page_field = hf.get('has_page_field', False)
            if has_page_field:
                # If the footer text looks like a page number (just digits),
                # use 'centered' instead
                if footer_text and re.match(r'^\d+$', footer_text.strip()):
                    section_config['footer'] = 'centered'
                else:
                    section_config['footer'] = 'centered' if not footer_text else footer_text
            elif footer_text:
                section_config['footer'] = footer_text
            else:
                section_config['footer'] = ''

            # Page number
            pn_format = hf.get('page_number_format', '')
            pn_start = hf.get('page_number_start')
            has_page_field = hf.get('has_page_field', False)

            if cls.section_type in SectionClassifier.NO_PAGE_NUMBER_TYPES:
                section_config['page_number'] = 'none'
            elif has_page_field or pn_format:
                pn_config = OrderedDict()
                if pn_format:
                    pn_config['format'] = pn_format
                else:
                    # Smart defaults based on section type
                    if cls.section_type in ('abstract', 'abstract_en', 'toc'):
                        pn_config['format'] = 'roman_upper'
                    else:
                        pn_config['format'] = 'arabic'
                if pn_start is not None:
                    pn_config['start'] = pn_start
                elif cls.section_type in ('abstract', 'body'):
                    # First abstract or body section gets start: 1
                    pn_config['start'] = 1
                if len(pn_config) > 1:
                    section_config['page_number'] = dict(pn_config)
                else:
                    section_config['page_number'] = {'format': pn_config.get('format', 'arabic')}
            else:
                # Default: centered page numbers with smart format
                if cls.section_type in ('abstract', 'abstract_en', 'toc'):
                    section_config['page_number'] = {
                        'format': 'roman_upper',
                        'start': 1,
                    }
                elif cls.section_type == 'body':
                    section_config['page_number'] = {
                        'format': 'arabic',
                        'start': 1,
                    }
                else:
                    section_config['page_number'] = {'format': 'arabic'}

            # Break
            section_config['break'] = 'nextPage'

            # Different first page
            if hf.get('different_first_page'):
                section_config['different_first_page'] = True

            # Even and odd headers for body
            if cls.section_type == 'body' and hf.get('even_and_odd_headers'):
                section_config['even_odd'] = True

            # Heading numbering for body
            if cls.section_type == 'body':
                section_config['heading_numbering'] = True

            # Add review comment for low confidence
            if cls.confidence < 0.7:
                comment_key = (
                    f'# ⚠️ REVIEW: confidence={cls.confidence:.2f} | '
                    f'evidence: {"; ".join(cls.evidence[:3])}'
                )
                sections_yaml[comment_key] = None

            sections_yaml[yaml_key] = section_config

        return sections_yaml

    def to_yaml(self) -> str:
        """Generate YAML string with proper formatting."""
        template = self.generate()
        return self._dump_yaml(template)

    def _dump_yaml(self, data: OrderedDict) -> str:
        """Dump OrderedDict to YAML, handling None values for comments."""
        lines = []

        def _dump_section(d, indent=0):
            prefix = '  ' * indent
            if isinstance(d, OrderedDict):
                for key, value in d.items():
                    # Comment keys (start with #)
                    if isinstance(key, str) and key.startswith('#'):
                        lines.append(f'{prefix}{key}')
                        if value is not None:
                            _dump_section(value, indent)
                        continue

                    if value is None:
                        lines.append(f'{prefix}{key}:')
                    elif isinstance(value, (OrderedDict, dict)):
                        lines.append(f'{prefix}{key}:')
                        _dump_section(value, indent + 1)
                    elif isinstance(value, list):
                        lines.append(f'{prefix}{key}:')
                        for item in value:
                            if isinstance(item, (dict, OrderedDict)):
                                lines.append(f'{prefix}  - ')
                                _dump_section(item, indent + 2)
                            else:
                                lines.append(f'{prefix}  - {self._format_value(item)}')
                    elif isinstance(value, bool):
                        lines.append(f'{prefix}{key}: {str(value).lower()}')
                    elif isinstance(value, str):
                        if '\n' in value or ':' in value or value.startswith('{'):
                            lines.append(f'{prefix}{key}: "{value}"')
                        else:
                            lines.append(f'{prefix}{key}: "{value}"')
                    else:
                        lines.append(f'{prefix}{key}: {value}')
            elif isinstance(d, dict):
                for key, value in d.items():
                    if isinstance(key, str) and key.startswith('#'):
                        lines.append(f'{prefix}{key}')
                        continue
                    if isinstance(value, (OrderedDict, dict)):
                        lines.append(f'{prefix}{key}:')
                        _dump_section(value, indent + 1)
                    elif isinstance(value, str):
                        if '\n' in value or ':' in value:
                            lines.append(f'{prefix}{key}: "{value}"')
                        else:
                            lines.append(f'{prefix}{key}: "{value}"')
                    else:
                        lines.append(f'{prefix}{key}: {value}')
            else:
                lines.append(f'{prefix}{self._format_value(d)}')

        _dump_section(data)
        return '\n'.join(lines) + '\n'

    @staticmethod
    def _format_value(v) -> str:
        if isinstance(v, bool):
            return str(v).lower()
        if isinstance(v, str):
            return f'"{v}"'
        return str(v)


# ═══════════════════════════════════════════════════════════════════
# Main orchestration
# ═══════════════════════════════════════════════════════════════════

def analyze_template(docx_path: str | Path) -> OrderedDict:
    """Full pipeline: analyze a .docx and return the template dict.

    Args:
        docx_path: Path to the .docx file.

    Returns:
        OrderedDict representing the YAML template structure.
    """
    docx_path = Path(docx_path)

    with DocxReader(docx_path) as reader:
        # Parse document structure
        sections = reader.parse()
        if not sections:
            raise ValueError(f"No sections found in {docx_path}")

        # 4. Extract headers/footers FIRST (used by classifier)
        hf_extractor = HeaderFooterExtractor(reader)
        hf_configs = [hf_extractor.extract_section_hf(s) for s in sections]

        # Build section_headers dict for classifier
        section_headers = {
            i: hf.get('header', '') for i, hf in enumerate(hf_configs)
        }

        # 1. Classify sections (with header context)
        classifier = SectionClassifier(sections, section_headers=section_headers)
        classifications = classifier.classify_all()

        # 2. Extract fonts (with style inheritance)
        styles_xml = None
        try:
            styles_xml = reader._read_xml('word/styles.xml')
        except Exception:
            pass
        font_extractor = FontExtractor(sections, styles_xml=styles_xml)
        fonts = font_extractor.extract_all()

        # 3. Extract page settings
        page_analyzer = PageAnalyzer(sections)
        page_info = page_analyzer.extract()

        # 5. Generate YAML template
        generator = TemplateGenerator(
            sections=sections,
            classifications=classifications,
            fonts=fonts,
            page_info=page_info,
            hf_configs=hf_configs,
        )

        return generator.generate()


def analyze_and_save(docx_path: str | Path, output_path: str | Path = None) -> Path:
    """Analyze a .docx and save the resulting YAML template to a file.

    Args:
        docx_path: Path to the .docx file.
        output_path: Optional output path. Default: same dir, .yaml extension.

    Returns:
        Path to the saved YAML file.
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path.with_suffix('.yaml')
    else:
        output_path = Path(output_path)

    template = analyze_template(docx_path)

    # Use a proper YAML dumper for the final output
    yaml_str = _custom_yaml_dump(template)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(yaml_str)

    return output_path


def _custom_yaml_dump(data: OrderedDict) -> str:
    """Dump template data to YAML with proper formatting.

    Uses a custom approach because we need:
    - Ordered sections
    - Comment lines for low-confidence items
    - Proper indentation
    """
    lines = []
    comments = []  # collect comment lines to insert

    def _emit_comment(key, value, indent_level):
        """Check if key is a comment placeholder."""
        if isinstance(key, str) and key.startswith('#'):
            comments.append((indent_level, key))
            return True
        return False

    def _dump(d, indent=0):
        prefix = '  ' * indent

        if isinstance(d, OrderedDict):
            keys = list(d.keys())
            for key in keys:
                value = d[key]

                # Handle comment keys
                if isinstance(key, str) and key.startswith('#'):
                    lines.append(f'{prefix}{key}')
                    if value is not None:
                        _dump(value, indent)
                    continue

                if isinstance(value, dict):
                    lines.append(f'{prefix}{key}:')
                    _dump(value, indent + 1)
                elif isinstance(value, list):
                    lines.append(f'{prefix}{key}:')
                    for item in value:
                        lines.append(f'{prefix}  - {_yaml_value(item)}')
                elif value is None:
                    lines.append(f'{prefix}{key}:')
                elif isinstance(value, bool):
                    lines.append(f'{prefix}{key}: {str(value).lower()}')
                else:
                    lines.append(f'{prefix}{key}: {_yaml_value(value)}')
        elif isinstance(d, dict):
            for key, value in d.items():
                if isinstance(value, dict):
                    lines.append(f'{prefix}{key}:')
                    _dump(value, indent + 1)
                elif isinstance(value, list):
                    lines.append(f'{prefix}{key}:')
                    for item in value:
                        lines.append(f'{prefix}  - {_yaml_value(item)}')
                elif value is None:
                    lines.append(f'{prefix}{key}:')
                elif isinstance(value, bool):
                    lines.append(f'{prefix}{key}: {str(value).lower()}')
                else:
                    lines.append(f'{prefix}{key}: {_yaml_value(value)}')

    def _yaml_value(v):
        if isinstance(v, bool):
            return str(v).lower()
        if isinstance(v, str):
            # Escape if needed
            if '#' in v or ':' in v or v.startswith('{') or v.startswith('[') or '\n' in v:
                return f'"{v}"'
            # Always quote strings that might be misinterpreted
            if v in ('true', 'false', 'yes', 'no', 'on', 'off', 'none', 'null'):
                return f'"{v}"'
            if v == '':
                return '""'
            return v
        if isinstance(v, (int, float)):
            return str(v)
        return str(v)

    _dump(data)
    return '\n'.join(lines) + '\n'


# ═══════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description='DocMind 模板分析器 — 从 .docx 提取节结构/字体/页面/页眉 → 生成 YAML 模板',
    )
    parser.add_argument('input', help='输入的 .docx 文件路径')
    parser.add_argument('-o', '--output', default=None,
                        help='输出 YAML 文件路径 (默认: 与输入同目录, .yaml 后缀)')
    parser.add_argument('--print', action='store_true', default=False,
                        help='打印 YAML 到 stdout 而不保存文件')

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'❌ 文件不存在: {input_path}')
        return 1

    if args.print:
        template = analyze_template(input_path)
        yaml_str = _custom_yaml_dump(template)
        print(yaml_str)
    else:
        output_path = analyze_and_save(input_path, args.output)
        print(f'✓ YAML 模板已生成: {output_path}')

    return 0


if __name__ == '__main__':
    import sys
    sys.exit(main())
