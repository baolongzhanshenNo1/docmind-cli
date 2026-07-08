"""
DocDiscover — 自动发现文档结构

读取目标 docx 文件，输出当前状态：
- 章节（H1 标题）分类：front_matter / body / back_matter
- 每个章节的页眉内容与字体（从 header XML 读取）
- 分节符类型（从 sectPr 的 w:type 读取）
- 页脚是否有 PAGE 域
- 是否有奇偶页页眉

用法:
    from pipeline.doc_discover import discover_doc
    doc_state = discover_doc("论文.docx")

章节分类规则:
    首个匹配 r'\d+[\u4e00-\u9fff]' 的 H1 之前 → front_matter
    最后两个 H1 → back_matter
    其余 → body
"""

import zipfile
import re
from pathlib import Path
from lxml import etree

# ── OOXML 命名空间 ──
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WPML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

W_NS = f'{{{W}}}'
R_NS = f'{{{R}}}'


def discover_doc(target_path):
    """从目标 docx 发现文档当前状态。

    Args:
        target_path: docx 文件路径 (str 或 Path)

    Returns:
        dict: {
            "sections": [
                {
                    "index": 0,
                    "heading": "摘要",
                    "category": "front_matter",
                    "current_header_text": "2024年 毕业设计",
                    "current_header_font_east": "宋体",
                    "current_header_font_ascii": "Times New Roman",
                    "current_header_font_size_pt": 9.0,
                    "current_header_underline": True,
                    "current_sectpr_type": "nextPage",
                    "current_footer_has_page": False,
                    "current_footer_page_format": None,
                    "has_even_header": False,
                },
                ...
            ],
            "headings": [
                {"text": "摘要", "level": 1, "section": 0},
                ...
            ]
        }
    """
    target_path = Path(target_path)
    if not target_path.exists():
        raise FileNotFoundError(f'Target file not found: {target_path}')

    with zipfile.ZipFile(target_path) as z:
        namelist = z.namelist()

        # 1. 读取 document.xml
        if 'word/document.xml' not in namelist:
            raise ValueError("Not a valid docx: word/document.xml not found")
        doc_tree = etree.parse(z.open('word/document.xml'))
        body = doc_tree.getroot().find(f'{W_NS}body')
        if body is None:
            raise ValueError("document.xml has no <w:body>")

        # 2. 读取 relationships (rId → header/footer 文件名)
        rels_map = _read_rels(z, namelist)

        # 2.5. 读取 styles.xml → 样式→标题级别映射
        style_level_map = _read_style_heading_levels(z, namelist)

        # 3. 提取所有 H1 标题（含段落元素引用，用于定位所属 section）
        h1_headings = _extract_h1_headings(body, style_level_map)

        # 4. 提取所有 sectPr 元素及其属性
        sectprs = _extract_sectprs(body)

        # 5. 为每个 H1 确定所属的 docx section index（基于 sectPr 位置）
        _assign_section_indices(h1_headings, body, sectprs)

        # 6. 读取每个 sectPr 对应的 header/footer 文件内容
        _read_header_footer_details(sectprs, z, rels_map, namelist)

        # 7. 章节分类（front_matter / body / back_matter）
        _categorize_sections(h1_headings)

        # 7.5. 回退：如果没有 H1 标题但有 sectPr，创建单个 body section
        if not h1_headings and sectprs:
            h1_headings.append({
                'text': '(无标题)',
                'level': 1,
                'element': None,
                'sectpr_index': 0,
                'category': 'body',
            })

    # 8. 构建输出
    sections = []
    for i, h in enumerate(h1_headings):
        sp = sectprs[h['sectpr_index']] if h['sectpr_index'] < len(sectprs) else {}
        sections.append({
            "index": i,
            "sectpr_index": h.get('sectpr_index', i),
            "heading": h['text'],
            "category": h.get('category', 'body'),
            "current_header_text": sp.get('header_text', ''),
            "current_header_font_east": sp.get('header_font_east'),
            "current_header_font_ascii": sp.get('header_font_ascii'),
            "current_header_font_size_pt": sp.get('header_font_size_pt'),
            "current_header_underline": sp.get('header_underline', False),
            "current_sectpr_type": sp.get('sectpr_type'),
            "current_footer_has_page": sp.get('footer_has_page', False),
            "current_footer_page_format": sp.get('footer_page_format'),
            "has_even_header": sp.get('has_even_header', False),
        })

    all_headings = _extract_all_headings(body, h1_headings, style_level_map)

    return {
        "sections": sections,
        "headings": all_headings,
    }


# ══════════════════════════════════════════════════════════════
# 内部函数
# ══════════════════════════════════════════════════════════════

def _read_rels(z, namelist):
    """读取 word/_rels/document.xml.rels，返回 rId → target 映射。"""
    rels_map = {}
    rels_path = 'word/_rels/document.xml.rels'
    if rels_path in namelist:
        try:
            rels_tree = etree.parse(z.open(rels_path))
            nsmap = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            for rel in rels_tree.iter():
                if rel.tag.endswith('}Relationship'):
                    rid = rel.get('Id')
                    target = rel.get('Target', '')
                    if rid:
                        rels_map[rid] = target
        except Exception:
            pass
    return rels_map


def _text_of(element) -> str:
    """提取 OOXML 元素中所有 w:t 文字。"""
    parts = []
    for t in element.iter(f'{W_NS}t'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts)


def _read_style_heading_levels(z, namelist) -> dict:
    """从 styles.xml 读取 styleId → heading level 映射。

    同时使用两种方式检测:
    1. outlineLvl (最可靠)
    2. 样式名匹配 (Heading 1, 标题 1, heading 1 等)

    Returns:
        dict: style_id → heading_level (1-9), 非标题样式不包含在内
    """
    level_map = {}

    styles_path = 'word/styles.xml'
    if styles_path not in namelist:
        return level_map

    try:
        styles_root = etree.parse(z.open(styles_path)).getroot()
    except Exception:
        return level_map

    for style in styles_root.iter(f'{W_NS}style'):
        style_id = style.get(f'{W_NS}styleId', '')
        if not style_id:
            continue

        pPr = style.find(f'{W_NS}pPr')
        level = None

        # 方式1: outlineLvl
        if pPr is not None:
            ol = pPr.find(f'{W_NS}outlineLvl')
            if ol is not None:
                try:
                    level = int(ol.get(f'{W_NS}val', -1)) + 1  # 0-based → 1-based
                except (ValueError, TypeError):
                    pass

        # 方式2: 样式名匹配
        if level is None:
            name_el = style.find(f'{W_NS}name')
            style_name = name_el.get(f'{W_NS}val', '') if name_el is not None else ''
            level = _heading_level_from_name(style_name, style_id)

        if level is not None and 1 <= level <= 9:
            level_map[style_id] = level

    return level_map


def _heading_level_from_name(name: str, style_id: str) -> int | None:
    """从样式名或 ID 推断标题级别。

    匹配模式:
    - "Heading 1", "heading 1" → 1
    - "标题 1", "标题1" → 1
    - "1" (仅当单独的数字样式ID) → 1
    """
    import re

    # "Heading N" / "heading N"
    m = re.match(r'^[Hh]eading\s*(\d+)$', name.strip())
    if m:
        return int(m.group(1))

    # "标题 N" / "标题N"
    m = re.match(r'^标题\s*(\d+)$', name.strip())
    if m:
        return int(m.group(1))

    # 纯数字样式 ID (如 "1", "2", "3")
    m = re.match(r'^(\d+)$', style_id.strip())
    if m and name.strip() in ('heading', 'Heading', '标题', ''):
        return int(m.group(1))

    return None


def _detect_heading_level(p, style_level_map: dict) -> int | None:
    """检测段落的标题级别。

    优先级:
    1. 段落级 outlineLvl
    2. 段落样式映射 (styles.xml)
    """
    pPr = p.find(f'{W_NS}pPr')
    if pPr is None:
        return None

    # 1. outlineLvl
    ol = pPr.find(f'{W_NS}outlineLvl')
    if ol is not None:
        try:
            return int(ol.get(f'{W_NS}val', -1)) + 1
        except (ValueError, TypeError):
            pass

    # 2. 样式映射
    pStyle = pPr.find(f'{W_NS}pStyle')
    if pStyle is not None:
        sid = pStyle.get(f'{W_NS}val', '')
        if sid in style_level_map:
            return style_level_map[sid]

    return None


# 内容模式匹配：常见中文论文章节标题
_NON_STYLED_H1_PATTERN = re.compile(
    r'^(?:摘\s*要|ABSTRACT|目\s*录|目录|'
    r'\d+[\u4e00-\u9fff]|'
    r'参考\s*文献|参考文献|致\s*谢|致谢|'
    r'附\s*录|附录)'
)

def _is_toc_entry(text: str) -> bool:
    """检测是否为 TOC 条目（章标题+页码，如 '1 绪论 1'）。"""
    # TOC 条目通常以数字+页码结尾，如 "1 绪论 1" 中末尾的 "1"
    # 正文标题可能是 "1 绪论"，TOC 中是 "1 绪论............1"
    # 简单检测：文本末尾有数字，且去除末尾数字后仍匹配 H1 模式
    m = re.match(r'^(.+?)\s*(\d+)$', text)
    if m:
        prefix = m.group(1).strip()
        norm = re.sub(r'\s+', '', prefix)
        if _NON_STYLED_H1_PATTERN.match(norm):
            return True
    return False

def _extract_h1_headings(body, style_level_map: dict) -> list[dict]:
    """从 body 提取所有 H1 标题。

    检测方式（优先级递减）:
    1. outlineLvl=0 或 1 → H1
    2. 样式映射 (需 styles.xml)
    3. 内容模式匹配（无样式时）：摘要/ABSTRACT/目录/数字章节/参考文献/致谢

    Returns:
        list of dicts with keys: text, level, element (lxml Element)
    """
    headings = []
    seen_texts = set()  # 去重

    for p in body.iter(f'{W_NS}p'):
        text = _text_of(p).strip()
        if not text or text in seen_texts:
            continue

        # 跳过 TOC 条目（如 "1 绪论 1", "参考文献 25"）
        if _is_toc_entry(text):
            continue

        level = _detect_heading_level(p, style_level_map)
        is_h1 = (level == 1)

        # 方式3: 内容模式匹配（无 heading level 时的回退）
        if level is None:
            norm = re.sub(r'\s+', '', text)
            if _NON_STYLED_H1_PATTERN.match(norm):
                is_h1 = True

        if is_h1:
            seen_texts.add(text)
            headings.append({
                'text': text,
                'level': 1,
                'element': p,
            })
    return headings


def _extract_all_headings(body, h1_headings, style_level_map: dict) -> list[dict]:
    """提取所有级别的标题。

    同时确定每个标题所属的 H1 section index。
    """
    # 建立 H1 元素 → section index 映射
    h1_elements = {}  # id(element) → section_index
    for i, h in enumerate(h1_headings):
        h1_elements[id(h['element'])] = i

    all_h = []
    current_section = 0
    for p in body.iter(f'{W_NS}p'):
        pPr = p.find(f'{W_NS}pPr')
        if pPr is None:
            continue

        # 检测 section break
        if pPr.find(f'{W_NS}sectPr') is not None:
            continue

        level = _detect_heading_level(p, style_level_map)
        if level is None:
            continue

        text = _text_of(p)
        if not text.strip():
            continue

        # 更新当前 section
        eid = id(p)
        if eid in h1_elements:
            current_section = h1_elements[eid]

        all_h.append({
            'text': text.strip(),
            'level': level,
            'section': current_section,
        })

    return all_h


def _extract_sectprs(body) -> list[dict]:
    """从 body 提取所有 sectPr 元素及其属性。

    遍历 body 的直接子元素，同时收集：
    - 作为 body 直接子元素的 sectPr（末尾节属性）
    - 位于 pPr 内部的 sectPr（段落级分节符）

    Returns:
        list of dicts，每个 dict 包含:
        - element: lxml Element (sectPr)
        - sectpr_type: w:type 属性值 (nextPage/oddPage/evenPage/continuous)
        - header_refs: {type: rId}
        - footer_refs: {type: rId}
        - title_pg: bool
        - pg_num_fmt: w:pgNumType 的格式/起始值
    """
    sectprs = []

    for child in body:
        # 情况1: 直接子元素是 sectPr
        if child.tag == f'{W_NS}sectPr':
            sectprs.append(_parse_sectpr(child))
        # 情况2: w:p 内部的 sectPr
        elif child.tag == f'{W_NS}p':
            pPr = child.find(f'{W_NS}pPr')
            if pPr is not None:
                sp = pPr.find(f'{W_NS}sectPr')
                if sp is not None:
                    sectprs.append(_parse_sectpr(sp))

    return sectprs


def _parse_sectpr(sp_element) -> dict:
    """解析单个 sectPr 元素的属性。"""
    info = {
        'element': sp_element,
        'sectpr_type': None,
        'header_refs': {},
        'footer_refs': {},
        'title_pg': False,
        'pg_num_fmt': None,
        'pg_num_start': None,
        # 以下由 _read_header_footer_details 填充
        'header_text': '',
        'header_font_east': None,
        'header_font_ascii': None,
        'header_font_size_pt': None,
        'header_underline': False,
        'footer_has_page': False,
        'footer_page_format': None,
        'has_even_header': False,
    }

    # w:type
    type_el = sp_element.find(f'{W_NS}type')
    if type_el is not None:
        info['sectpr_type'] = type_el.get(f'{W_NS}val')

    # headerReference
    for hr in sp_element.findall(f'{W_NS}headerReference'):
        typ = hr.get(f'{W_NS}type', 'default')
        rid = hr.get(f'{R_NS}id')
        info['header_refs'][typ] = rid
        if typ == 'even':
            info['has_even_header'] = True

    # footerReference
    for fr in sp_element.findall(f'{W_NS}footerReference'):
        typ = fr.get(f'{W_NS}type', 'default')
        rid = fr.get(f'{R_NS}id')
        info['footer_refs'][typ] = rid

    # titlePg
    if sp_element.find(f'{W_NS}titlePg') is not None:
        info['title_pg'] = True

    # pgNumType
    pnt = sp_element.find(f'{W_NS}pgNumType')
    if pnt is not None:
        info['pg_num_fmt'] = pnt.get(f'{W_NS}fmt')
        start = pnt.get(f'{W_NS}start')
        if start is not None:
            info['pg_num_start'] = int(start)

    return info


def _assign_section_indices(h1_headings, body, sectprs):
    """为每个 H1 标题确定所属的 docx section index。

    遍历 body 的直接子元素，跟踪当前 section index。
    每当遇到段落内包含 sectPr（分节符），递增 section index。
    最后一个 sectPr（body 直接子元素）属于最后一节。
    """
    current_sec = 0
    # 建立 H1 元素 id → heading dict 的快速查找
    h1_by_elem = {id(h['element']): h for h in h1_headings}

    for child in body:
        if child.tag == f'{W_NS}p':
            # 检查这个段落是否是 H1
            eid = id(child)
            if eid in h1_by_elem:
                h1_by_elem[eid]['sectpr_index'] = min(current_sec, len(sectprs) - 1) if sectprs else 0

            # 检查段落内是否有分节符
            pPr = child.find(f'{W_NS}pPr')
            if pPr is not None and pPr.find(f'{W_NS}sectPr') is not None:
                current_sec += 1

    # 确保所有 H1 都有 sectpr_index
    default_idx = len(sectprs) - 1 if sectprs else 0
    for h in h1_headings:
        if 'sectpr_index' not in h:
            h['sectpr_index'] = default_idx


def _read_header_footer_details(sectprs, z, rels_map, namelist):
    """读取每个 sectPr 引用的 header/footer XML 文件，提取文字、字体等信息。

    对每个 sectPr:
    - 通过 header_refs 找到 headerN.xml，读取文字、字体、下划线
    - 通过 footer_refs 找到 footerN.xml，检查是否有 PAGE 域

    优先级: default > even > first（默认页眉优先）
    """
    # 缓存已读取的 header/footer 文件
    header_cache = {}
    footer_cache = {}

    for sp in sectprs:
        # ── Header ──
        # 按优先级排序: default 最优先
        header_priority = ['default', 'even', 'first']
        for htype in header_priority:
            rid = sp['header_refs'].get(htype)
            if not rid or rid not in rels_map:
                continue
            target = rels_map[rid]
            filepath = f'word/{target}'

            if filepath in header_cache:
                hinfo = header_cache[filepath]
            elif filepath in namelist:
                hinfo = _read_header_xml(z, filepath)
                header_cache[filepath] = hinfo
            else:
                hinfo = {}

            if hinfo:
                sp['header_text'] = hinfo.get('text', '')
                sp['header_font_east'] = hinfo.get('font_east')
                sp['header_font_ascii'] = hinfo.get('font_ascii')
                sp['header_font_size_pt'] = hinfo.get('font_size_pt')
                sp['header_underline'] = hinfo.get('underline', False)
                break  # 取优先级最高的

        # ── Footer ──
        footer_priority = ['default', 'even', 'first']
        for ftype in footer_priority:
            rid = sp['footer_refs'].get(ftype)
            if not rid or rid not in rels_map:
                continue
            target = rels_map[rid]
            filepath = f'word/{target}'

            if filepath in footer_cache:
                finfo = footer_cache[filepath]
            elif filepath in namelist:
                finfo = _read_footer_xml(z, filepath)
                footer_cache[filepath] = finfo
            else:
                finfo = {}

            if finfo:
                sp['footer_has_page'] = finfo.get('has_page', False)
                sp['footer_page_format'] = finfo.get('page_format')
            break


def _read_header_xml(z, filepath) -> dict:
    """读取 header XML 文件，提取文字内容和字体格式。

    Returns:
        dict with keys: text, font_east, font_ascii, font_size_pt, underline
    """
    try:
        root = etree.parse(z.open(filepath)).getroot()
    except Exception:
        return {}

    info = {'text': '', 'font_east': None, 'font_ascii': None,
             'font_size_pt': None, 'underline': False}

    # 提取所有文字
    info['text'] = _text_of(root)

    # 从第一个 run 提取字体信息
    for r in root.iter(f'{W_NS}r'):
        rPr = r.find(f'{W_NS}rPr')
        if rPr is not None:
            # 字体
            rf = rPr.find(f'{W_NS}rFonts')
            if rf is not None:
                east = rf.get(f'{W_NS}eastAsia')
                ascii_f = rf.get(f'{W_NS}ascii')
                if east:
                    info['font_east'] = east
                if ascii_f:
                    info['font_ascii'] = ascii_f

            # 字号
            sz = rPr.find(f'{W_NS}sz')
            if sz is not None:
                try:
                    info['font_size_pt'] = int(sz.get(f'{W_NS}val', '18')) / 2
                except (ValueError, TypeError):
                    pass

            # 下划线（文字级）
            if rPr.find(f'{W_NS}u') is not None:
                info['underline'] = True

            # 找到第一个有字体信息的 run 就退出
            if info['font_east'] or info['font_ascii']:
                break

    # 检查段落级边框（页眉下划线通常用 pBdr 实现）
    for p in root.iter(f'{W_NS}p'):
        pPr = p.find(f'{W_NS}pPr')
        if pPr is not None:
            pBdr = pPr.find(f'{W_NS}pBdr')
            if pBdr is not None:
                bottom = pBdr.find(f'{W_NS}bottom')
                if bottom is not None:
                    val = bottom.get(f'{W_NS}val', '')
                    if val and val != 'none':
                        info['underline'] = True
            break

    return info


def _read_footer_xml(z, filepath) -> dict:
    """读取 footer XML 文件，检查是否有 PAGE 域。

    PAGE 域结构:
        <w:fldChar w:fldCharType="begin"/>
        <w:instrText> PAGE </w:instrText>
        <w:fldChar w:fldCharType="separate"/>
        ...
        <w:fldChar w:fldCharType="end"/>

    Returns:
        dict with keys: has_page, page_format
    """
    try:
        root = etree.parse(z.open(filepath)).getroot()
    except Exception:
        return {}

    info = {'has_page': False, 'page_format': None}

    # 检查 PAGE 域
    for instr in root.iter(f'{W_NS}instrText'):
        if instr.text and 'PAGE' in instr.text.upper():
            info['has_page'] = True
            # 尝试提取格式：PAGE \* ROMAN 等
            text = instr.text.strip()
            if 'ROMAN' in text.upper():
                info['page_format'] = 'upperRoman'
            elif 'roman' in text:
                info['page_format'] = 'lowerRoman'
            elif 'ALPHABETIC' in text.upper():
                info['page_format'] = 'upperLetter'
            else:
                info['page_format'] = 'decimal'
            break

    # 也检查 NUM PAGES 域
    if not info['has_page']:
        for instr in root.iter(f'{W_NS}instrText'):
            if instr.text and 'NUMPAGES' in instr.text.upper():
                info['has_page'] = True
                break

    return info


# ── 章节分类 ──────────────────────────────────────────────

# 匹配 "1绪论", "2相关技术" 等正文章节标题
BODY_H1_PATTERN = re.compile(r'^\d+[\u4e00-\u9fff]')


def _categorize_sections(h1_headings):
    """根据 H1 标题文字分类章节。

    规则:
    - 首个匹配 r'\d+[\u4e00-\u9fff]' 的 H1 之前 → front_matter
    - 最后两个 H1 → back_matter
    - 其余 → body

    特殊情况处理:
    - 无匹配时：全部归为 body（或根据总数量启发式判断）
    - H1 数量 ≤ 2 时：无 body，前面为 front，后面为 back
    """
    n = len(h1_headings)
    if n == 0:
        return

    # 找到第一个匹配正文模式的 H1 索引
    first_body_idx = None
    for i, h in enumerate(h1_headings):
        if BODY_H1_PATTERN.match(h['text'].replace(' ', '')):
            first_body_idx = i
            break

    if first_body_idx is None:
        # 没有匹配到正文章节模式
        # 如果 ≥ 3 个 H1，尝试启发式：前 1-2 个为 front，后 1-2 个为 back
        if n >= 3:
            # 前 2 个 front_matter，后 2 个 back_matter，中间的 body
            for i in range(min(2, n)):
                h1_headings[i]['category'] = 'front_matter'
            for i in range(max(n - 2, 2), n):
                h1_headings[i]['category'] = 'back_matter'
            for i in range(2, n - 2):
                h1_headings[i]['category'] = 'body'
        elif n == 2:
            h1_headings[0]['category'] = 'front_matter'
            h1_headings[1]['category'] = 'back_matter'
        else:
            h1_headings[0]['category'] = 'body'
        return

    # 正常分类
    back_count = min(2, n - first_body_idx) if n - first_body_idx >= 2 else max(1, n - first_body_idx - 1)

    for i, h in enumerate(h1_headings):
        if i < first_body_idx:
            h['category'] = 'front_matter'
        elif i >= n - back_count:
            h['category'] = 'back_matter'
        else:
            h['category'] = 'body'


# ══════════════════════════════════════════════════════════════
# 便捷函数：直接输出可读摘要
# ══════════════════════════════════════════════════════════════

def print_summary(doc_state):
    """打印 discover_doc 结果的可读摘要。"""
    print('=' * 60)
    print('DocDiscover Summary')
    print('=' * 60)

    sections = doc_state.get('sections', [])
    for sec in sections:
        cat = sec['category']
        emoji = {'front_matter': '📋', 'body': '📄', 'back_matter': '📎'}.get(cat, '❓')
        print(f"\n{emoji} Section {sec['index']} [{cat}] — \"{sec['heading']}\"")
        print(f"   Header text:  \"{sec['current_header_text']}\"")
        print(f"   Header font:  east={sec['current_header_font_east']}, "
              f"ascii={sec['current_header_font_ascii']}, "
              f"size={sec['current_header_font_size_pt']}pt, "
              f"underline={sec['current_header_underline']}")
        print(f"   SectPr type:  {sec['current_sectpr_type']}")
        print(f"   Even header:  {sec['has_even_header']}")
        print(f"   Footer PAGE:  {sec['current_footer_has_page']}"
              f" (fmt={sec['current_footer_page_format']})")

    print(f"\n---\nTotal: {len(sections)} sections, "
          f"{len(doc_state.get('headings', []))} headings")
