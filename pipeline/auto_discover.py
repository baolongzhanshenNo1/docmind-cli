"""Pipeline 步骤0: 自动发现文档结构

从规范模板提取格式规则 + 从目标文档发现结构。
"""
import sys
from pathlib import Path

# Archived modules (docmind3) path

import docx
from docx.oxml.ns import qn
from lxml import etree
import re

from docmind3 import DocxReader
from .config import PipelineConfig
from .spec_reader import read_spec, merge_with_config

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def auto_discover(config: PipelineConfig) -> PipelineConfig:
    """发现文档结构，填充配置"""
    print('[auto_discover] Reading spec template...')
    
    # ── 1. 从规范模板提取格式规则 ──
    if config.spec_docx and config.spec_docx.exists():
        rules = read_spec(config.spec_docx)
        rules = merge_with_config(rules, config)
        config.spec_rules = rules
        print(f'  Rules: {list(rules.keys())}')
        h1_info = rules.get('h1', {})
        if h1_info:
            print('  H1: {} {}pt'.format(h1_info.get('font_east'), h1_info.get('font_size_pt')))
    else:
        print('  No spec template found, using defaults')
    
    # ── 2. 用 Reader 解析目标文档 ──
    print('[auto_discover] Analyzing target document...')
    reader = DocxReader()
    doc = reader.read(str(config.input_docx))
    
    # ── 3. 样式 ID ──
    style_map = _discover_style_ids(doc, config)
    if style_map:
        config.style_ids = style_map
        print(f'  Style IDs: {style_map}')
    
    # ── 4. 节结构 ──
    section_count = _get_section_count(config)
    print(f'  Sections: {section_count}')
    
    h1_count = sum(1 for h in doc.all_headings if h.level == 1)
    h2_count = sum(1 for h in doc.all_headings if h.level == 2)
    h3_count = sum(1 for h in doc.all_headings if h.level == 3)
    print(f'  Headings: H1={h1_count} H2={h2_count} H3={h3_count}')
    
    # ── 5. 自动推导节页眉映射 ──
    if not config.section_headers:
        config.section_headers = _derive_section_headers(doc, config)
        print(f'  Section headers: {len(config.section_headers)} sections mapped')
    
    # ── 6. 诊断 ──
    from docmind3 import DocxFixer
    fixer = DocxFixer(doc)
    issues = fixer.diagnose()
    print(f'  Issues found: {len(issues)}')
    
    return config


def _discover_style_ids(doc, config):
    """从文档自动发现 H1/H2/Body 样式 ID"""
    result = {}
    
    for h in doc.all_headings:
        style = getattr(h, 'style_id', None)
        if style:
            if h.level == 0 and 'h1' not in result:
                result['h1'] = style
            elif h.level == 1 and 'h2' not in result:
                result['h2'] = style
    
    result.setdefault('h1', '1')
    result.setdefault('h2', '20')
    result.setdefault('body', 'a8')
    
    return result


def _get_section_count(config):
    docx_obj = docx.Document(str(config.input_docx))
    return len(docx_obj.sections)


def _derive_section_headers(doc, config) -> dict:
    """根据文档标题树推导节页眉映射
    
    策略:
    - 摘要/ABSTRACT/目录 → 使用各自标题文字
    - 1-7章 → 奇数=config.odd_header_template, 偶数=config.even_header_template
    - 参考文献/致谢 → 使用各自标题文字
    """
    result = {}
    h1_headings = [h for h in doc.all_headings if h.level == 1]
    
    for i, h in enumerate(h1_headings):
        text = re.sub(r'\s+', '', h.text)  # 去所有空格（含全角）
        
        if text in ['摘要', 'ABSTRACT', '目录']:
            result[i] = (text, '', '')  # 前部 matter，用去空格后的文字
        elif text in ['参考文献', '致谢', '附录']:
            result[i] = (text, '', '')  # 后部 matter，用去空格后的文字
        else:
            # 正文章节
            odd = config.odd_header_template.format(year=config.year)
            even = config.even_header_template
            result[i] = (odd, even, '')
    
    return result


def _derive_page_number_format(doc, config) -> dict:
    """推导每个 section 的页码格式（罗马 vs 阿拉伯）"""
    result = {}
    h1_headings = [h for h in doc.all_headings if h.level == 1]
    
    roman_headings = {'摘要', 'ABSTRACT', '目录'}
    arabic_start = None
    
    for i, h in enumerate(h1_headings):
        text = re.sub(r'\s+', '', h.text)
        if text in roman_headings:
            result[i] = 'roman'
        else:
            if arabic_start is None:
                arabic_start = i
            result[i] = 'arabic'
    
    result['_arabic_start'] = arabic_start
    return result
