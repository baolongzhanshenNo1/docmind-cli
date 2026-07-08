"""
SpecReader v2 — 从规范模板 OOXML 提取纯格式规则

提取（且只提取）以下格式信息：
1. 各级标题格式（从 styles.xml + document.xml 的 outlineLvl）
2. 页眉格式（从 header XML 文件）
3. 页脚格式（从 footer XML 文件）
4. 分节符与页边距（从 sectPr）

不提取任何文字内容。
"""
import zipfile
from pathlib import Path
from typing import Dict, Optional, Any, List
from collections import defaultdict

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ── 外部 interface ──


def read_spec_v2(spec_path: Path) -> dict:
    """从规范模板读取所有格式规则（纯格式，不含文字）。

    Args:
        spec_path: .docx 规范模板路径

    Returns:
        dict with keys:
          - h1, h2, h3: 各级标题格式
          - body: 正文默认格式
          - header_format: 页眉格式（可能为 None）
          - footer_format: 页脚格式（可能为 None）
          - section_break: 分节符类型
          - page: 页边距信息
    """
    if not spec_path or not spec_path.exists():
        return {}

    with zipfile.ZipFile(spec_path) as z:
        namelist = z.namelist()

        styles_xml = z.read('word/styles.xml') if 'word/styles.xml' in namelist else None
        doc_xml = z.read('word/document.xml') if 'word/document.xml' in namelist else None
        rels_xml = z.read('word/_rels/document.xml.rels') if 'word/_rels/document.xml.rels' in namelist else None

        header_xmls = {}
        footer_xmls = {}
        for name in namelist:
            if name.startswith('word/header') and name.endswith('.xml'):
                header_xmls[name] = z.read(name)
            elif name.startswith('word/footer') and name.endswith('.xml'):
                footer_xmls[name] = z.read(name)

    rules: Dict[str, Any] = {}

    # 1. 标题格式 + 正文格式
    if styles_xml is not None:
        styles_root = etree.fromstring(styles_xml)
        _extract_heading_from_styles(styles_root, rules)
        _extract_body_from_styles(styles_root, rules)

    if doc_xml is not None:
        doc_root = etree.fromstring(doc_xml)
        _extract_heading_from_document(doc_root, rules)
        _extract_section_info(doc_root, rules)

    # 2. 页眉格式
    rules['header_format'] = (_extract_header_format(doc_xml, rels_xml, header_xmls)
                              if rels_xml is not None and header_xmls else None)

    # 3. 页脚格式
    rules['footer_format'] = (_extract_footer_format(doc_xml, rels_xml, footer_xmls)
                              if rels_xml is not None and footer_xmls else None)

    return rules


# ═══════════════════════════════════════════════════════════════
# 1. 标题格式提取
# ═══════════════════════════════════════════════════════════════

HEADING_STYLE_NAMES = {
    'heading 1': 'h1', 'heading1': 'h1', '标题 1': 'h1', '标题1': 'h1',
    '1': 'h1',
    'heading 2': 'h2', 'heading2': 'h2', '标题 2': 'h2', '标题2': 'h2',
    '2': 'h2',
    'heading 3': 'h3', 'heading3': 'h3', '标题 3': 'h3', '标题3': 'h3',
    '3': 'h3',
}

OUTLINE_TO_KEY = {0: 'h1', 1: 'h2', 2: 'h3'}


def _extract_heading_from_styles(styles_root, rules: dict):
    """从 styles.xml 提取各级标题格式（通过 outlineLvl 或样式名）。"""
    for style in styles_root.iter(f'{{{W}}}style'):
        style_id = style.get(f'{{{W}}}styleId', '')
        pPr = style.find(f'{{{W}}}pPr')
        rPr = style.find(f'{{{W}}}rPr')

        key = None

        if pPr is not None:
            ol = pPr.find(f'{{{W}}}outlineLvl')
            if ol is not None:
                level = int(ol.get(f'{{{W}}}val'))
                key = OUTLINE_TO_KEY.get(level)

        if key is None:
            name_el = style.find(f'{{{W}}}name')
            name_lower = (name_el.get(f'{{{W}}}val') or '').lower() if name_el is not None else ''
            key = HEADING_STYLE_NAMES.get(name_lower) or HEADING_STYLE_NAMES.get(style_id.lower())

        if key is None:
            continue

        fmt = _read_format(pPr, rPr)
        if key not in rules or not rules[key]:
            rules[key] = fmt
        else:
            for k, v in fmt.items():
                if v is not None and k not in rules[key]:
                    rules[key][k] = v


def _extract_heading_from_document(doc_root, rules: dict):
    """从 document.xml 提取段落直接格式，补充/覆盖 styles 中的标题规则。

    使用投票策略：对每个 outlineLvl，按出现次数最多的格式作为最终规则。
    同时收集 spacing 信息。
    """
    # 收集每个 level 的格式（包括 spacing）
    level_formats: Dict[int, List[dict]] = defaultdict(list)

    for p in doc_root.iter(f'{{{W}}}p'):
        pPr = p.find(f'{{{W}}}pPr')
        if pPr is None:
            continue
        ol = pPr.find(f'{{{W}}}outlineLvl')
        if ol is None:
            continue

        level = int(ol.get(f'{{{W}}}val'))
        key = OUTLINE_TO_KEY.get(level)
        if key is None:
            continue

        # 从第一个 run 提取格式
        rPr = None
        for r in p.iter(f'{{{W}}}r'):
            rPr = r.find(f'{{{W}}}rPr')
            if rPr is not None:
                break

        fmt = _read_format(pPr, rPr)
        level_formats[level].append(fmt)

    for level, fmts in level_formats.items():
        if not fmts:
            continue

        key = OUTLINE_TO_KEY.get(level)
        font_fmt = _pick_best_format(fmts)

        # 合并 spacing（从具有最完整信息的条目中获取）
        spacing_fmt = _pick_best_spacing(fmts)

        merged = {}
        if font_fmt:
            merged.update(font_fmt)
        if spacing_fmt:
            merged.update(spacing_fmt)

        if merged:
            if key not in rules:
                rules[key] = merged
            else:
                for k, v in merged.items():
                    if v is not None:
                        rules[key][k] = v

    # 清理空 dict
    for lvl_key in ['h1', 'h2', 'h3']:
        if lvl_key in rules and not rules[lvl_key]:
            del rules[lvl_key]


def _pick_best_format(fmts: List[dict]) -> Optional[dict]:
    """从多个格式 dict 中选出最优的字符/段落格式。

    策略（按优先级）：
    1. 过滤掉明显是注释/说明的格式（楷体、小字号）
    2. 过滤掉 TOC 格式（font_east 为空但 font_ascii 为宋体 + distribute 对齐）
    3. 优先选择字号最大的（标题字号 > 正文 > TOC）
    4. 若字号相同，优先选带font_east的（更可能是正式标题）
    """
    if not fmts:
        return None

    # 过滤明显非标题格式
    candidates = []
    for fmt in fmts:
        font_east = (fmt.get('font_east') or '').lower()
        font_ascii = (fmt.get('font_ascii') or '').lower()
        sz = fmt.get('font_size_pt', 0) or 0
        align = fmt.get('alignment', '')

        # 排除楷体注释
        if '楷体' in font_east or '楷体' in font_ascii:
            continue

        # 排除 TOC 条目（无 font_east + distribute 对齐 + 小字号）
        if (not font_east and font_ascii and '宋体' in font_ascii
                and align == 'distribute' and sz <= 12):
            continue

        candidates.append(fmt)

    if not candidates:
        candidates = fmts  # fallback

    # 按字号降序排列（标题通常字号较大）
    candidates.sort(key=lambda f: f.get('font_size_pt', 0) or 0, reverse=True)

    # 取字号最大的第一个
    return candidates[0] if candidates else None


def _pick_best_spacing(fmts: List[dict]) -> Optional[dict]:
    """从多个格式中选出间距信息（before/after/lines/lineRule）。"""
    spacing_keys = ['before_twips', 'after_twips', 'before_lines', 'after_lines',
                     'line', 'lineRule', 'first_line_indent']
    result = {}
    for fmt in fmts:
        for k in spacing_keys:
            if k in fmt and fmt[k] is not None and k not in result:
                result[k] = fmt[k]
    return result if result else None


# ═══════════════════════════════════════════════════════════════
# 1b. 正文默认格式
# ═══════════════════════════════════════════════════════════════

def _extract_body_from_styles(styles_root, rules: dict):
    """提取正文默认段落格式。"""
    # 优先：default paragraph style
    for style in styles_root.iter(f'{{{W}}}style'):
        if style.get(f'{{{W}}}type') == 'paragraph' and style.get(f'{{{W}}}default') == '1':
            pPr = style.find(f'{{{W}}}pPr')
            rPr = style.find(f'{{{W}}}rPr')
            rules['body'] = _read_format(pPr, rPr)
            rules['body_style_id'] = style.get(f'{{{W}}}styleId', 'Normal')
            return

    # fallback: Normal style
    for style in styles_root.iter(f'{{{W}}}style'):
        if style.get(f'{{{W}}}styleId') == 'Normal' and style.get(f'{{{W}}}type') == 'paragraph':
            pPr = style.find(f'{{{W}}}pPr')
            rPr = style.find(f'{{{W}}}rPr')
            rules['body'] = _read_format(pPr, rPr)
            rules['body_style_id'] = 'Normal'
            break

    # fallback: docDefaults
    if 'body' not in rules:
        dd = styles_root.find(f'{{{W}}}docDefaults')
        if dd is not None:
            rPrDefault = dd.find(f'{{{W}}}rPrDefault')
            if rPrDefault is not None:
                rPr = rPrDefault.find(f'{{{W}}}rPr')
                if rPr is not None:
                    rules['body'] = _read_format(None, rPr)


# ═══════════════════════════════════════════════════════════════
# 2. 页眉格式提取
# ═══════════════════════════════════════════════════════════════

def _extract_header_format(doc_xml: Optional[bytes],
                           rels_xml: bytes,
                           header_xmls: Dict[str, bytes]) -> Optional[dict]:
    """提取页眉格式。"""
    header_file = _resolve_part_file(doc_xml, rels_xml, 'headerReference', header_xmls)
    if header_file is None:
        return None

    xml_bytes = header_xmls.get(header_file)
    if xml_bytes is None:
        return None

    return _parse_header_footer_xml(xml_bytes)


def _resolve_part_file(doc_xml: Optional[bytes],
                       rels_xml: bytes,
                       ref_tag: str,
                       part_xmls: Dict[str, bytes]) -> Optional[str]:
    """从关系解析 header/footer 文件路径。"""
    rid = None
    if doc_xml is not None:
        doc_root = etree.fromstring(doc_xml)
        for sp in doc_root.iter(f'{{{W}}}sectPr'):
            ref = sp.find(f'{{{W}}}{ref_tag}')
            if ref is not None:
                rid = ref.get(f'{{{R}}}id')
                break

    if rid is not None:
        rels_root = etree.fromstring(rels_xml)
        for rel in rels_root:
            if rel.get('Id') == rid:
                target = rel.get('Target', '')
                return f'word/{target}'

    # fallback: 返回第一个匹配的 part
    for name in sorted(part_xmls.keys()):
        return name

    return None


def _parse_header_footer_xml(xml_bytes: bytes) -> dict:
    """从 header/footer XML 提取格式信息。

    提取：字体、字号、对齐、加粗、边框（下划线）。
    不提取任何文字内容。
    """
    root = etree.fromstring(xml_bytes)
    fmt: Dict[str, Any] = {}

    # 处理第一个段落
    for p in root.iter(f'{{{W}}}p'):
        pPr = p.find(f'{{{W}}}pPr')
        if pPr is not None:
            # 对齐
            jc = pPr.find(f'{{{W}}}jc')
            if jc is not None:
                fmt['alignment'] = jc.get(f'{{{W}}}val')

            # 边框（下划线 / 顶线等）
            pBdr = pPr.find(f'{{{W}}}pBdr')
            if pBdr is not None:
                borders = {}
                for border_tag in ['top', 'bottom', 'left', 'right']:
                    b_el = pBdr.find(f'{{{W}}}{border_tag}')
                    if b_el is not None:
                        val = b_el.get(f'{{{W}}}val', 'nil')
                        if val not in ('nil', 'none', ''):
                            borders[border_tag] = {
                                'val': val,
                                'sz': int(b_el.get(f'{{{W}}}sz', '0')),
                                'color': b_el.get(f'{{{W}}}color', 'auto'),
                            }
                if borders:
                    fmt['borders'] = borders

        # 提取 run 字符格式：逐属性取第一个出现的值。
        # 不能死认第一个 run —— 页脚首个 run 常是无 sz 的空白占位符（页码文字在后续 run）。
        for r in p.iter(f'{{{W}}}r'):
            rPr = r.find(f'{{{W}}}rPr')
            if rPr is None:
                continue
            rf = rPr.find(f'{{{W}}}rFonts')
            if rf is not None:
                east = rf.get(f'{{{W}}}eastAsia')
                ascii_ = rf.get(f'{{{W}}}ascii')
                if east and 'font_east' not in fmt:
                    fmt['font_east'] = east
                if ascii_ and 'font_ascii' not in fmt:
                    fmt['font_ascii'] = ascii_

            sz = rPr.find(f'{{{W}}}sz')
            if sz is not None and 'font_size_pt' not in fmt:
                fmt['font_size_pt'] = int(sz.get(f'{{{W}}}val')) / 2

            b = rPr.find(f'{{{W}}}b')
            if b is not None and 'bold' not in fmt:
                fmt['bold'] = b.get(f'{{{W}}}val', 'true') not in ('false', '0')

        break  # 只看第一个段落

    return fmt


# ═══════════════════════════════════════════════════════════════
# 3. 页脚格式提取（含页码格式）
# ═══════════════════════════════════════════════════════════════

def _extract_footer_format(doc_xml: Optional[bytes],
                           rels_xml: bytes,
                           footer_xmls: Dict[str, bytes]) -> Optional[dict]:
    """提取页脚格式和页码格式。"""
    footer_file = _resolve_part_file(doc_xml, rels_xml, 'footerReference', footer_xmls)
    if footer_file is None:
        return None

    xml_bytes = footer_xmls.get(footer_file)
    if xml_bytes is None:
        return None

    fmt = _parse_header_footer_xml(xml_bytes)

    # 页码格式检测
    pn_format = _detect_page_number_format(xml_bytes)
    if pn_format:
        fmt['page_number_format'] = pn_format

    return fmt


def _detect_page_number_format(footer_xml: bytes) -> Optional[str]:
    """检测页码格式。

    解析整个 footer XML（递归包含嵌套结构），寻找 PAGE 字段及其上下文文本。

    Returns:
        '{PAGE}' 或 '第{PAGE}页' 或 '- {PAGE} -' 等
    """
    root = etree.fromstring(footer_xml)

    # 全局收集所有 items（递归，含嵌套 text box）
    all_items = []
    _collect_footer_items(root, all_items)

    if not all_items:
        return None

    # 构建完整序列
    sequence = []
    i = 0
    while i < len(all_items):
        typ, val = all_items[i]
        if typ == 'field_begin':
            has_page = False
            j = i + 1
            while j < len(all_items):
                t2, v2 = all_items[j]
                if t2 == 'instrText' and 'PAGE' in v2.upper():
                    has_page = True
                elif t2 == 'field_end':
                    break
                j += 1
            if has_page:
                sequence.append('{PAGE}')
            i = j + 1
        elif typ == 'text':
            txt = val.strip()
            if txt:
                sequence.append(txt)
            i += 1
        else:
            i += 1

    if '{PAGE}' not in sequence:
        return None

    # 找所有 {PAGE} 位置
    page_positions = [idx for idx, item in enumerate(sequence) if item == '{PAGE}']
    if not page_positions:
        return None

    first_page = page_positions[0]
    last_page = page_positions[-1]

    # 收集第一个 PAGE 前的所有文本，和最后一个 PAGE 后的所有文本
    before_items = [item for idx, item in enumerate(sequence)
                    if idx < first_page and item != '{PAGE}']
    after_items = [item for idx, item in enumerate(sequence)
                   if idx > last_page and item != '{PAGE}']

    before_text = ''.join(before_items).strip()
    after_text = ''.join(after_items).strip()

    # 如果 before_text 有多个语义片段，最后一个可能是 suffix
    # （当结构为：prefix text, suffix text, {PAGE} 的情况）
    if before_text and not after_text and len(before_items) >= 2:
        # 尝试拆分：最后一个元素作为 after_text
        potential_after = before_items[-1].strip()
        potential_before = ''.join(before_items[:-1]).strip()
        # 判断是否是常见的页脚后缀（如 "页"）
        if potential_after in ('页', '）', ')', '.', '】'):
            before_text = potential_before
            after_text = potential_after

    if before_text and after_text:
        return f'{before_text}{{PAGE}}{after_text}'
    elif after_text:
        return f'{{PAGE}}{after_text}'
    elif before_text:
        return f'{before_text}{{PAGE}}'
    else:
        return '{PAGE}'


def _collect_footer_items(element, items: list):
    """递归收集 footer/header 中的文本片段和 field 标记。

    不关心文字内容本身，只关心结构标记和 field 类型。
    """
    tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag

    if tag == 't':
        text = (element.text or '')
        items.append(('text', text))
        return

    if tag == 'fldChar':
        fld_type = element.get(f'{{{W}}}fldCharType')
        if fld_type == 'begin':
            items.append(('field_begin', ''))
        elif fld_type == 'end':
            items.append(('field_end', ''))
        elif fld_type == 'separate':
            items.append(('field_separate', ''))
        return

    if tag == 'instrText':
        text = (element.text or '')
        items.append(('instrText', text))
        return

    # 递归子元素
    for child in element:
        _collect_footer_items(child, items)


# ═══════════════════════════════════════════════════════════════
# 4. 分节符 + 页边距
# ═══════════════════════════════════════════════════════════════

def _extract_section_info(doc_root, rules: dict):
    """从 document.xml 的 sectPr 提取分节符类型、奇偶页、页边距。"""
    sects = list(doc_root.iter(f'{{{W}}}sectPr'))
    if not sects:
        return

    # 分节符类型
    type_el = sects[0].find(f'{{{W}}}type')
    if type_el is not None:
        rules['section_break'] = type_el.get(f'{{{W}}}val')
    else:
        rules['section_break'] = 'nextPage'

    # 奇偶页
    eo = sects[0].find(f'{{{W}}}evenAndOddHeaders')
    rules['even_and_odd_headers'] = eo is not None

    # 页边距
    pm = sects[0].find(f'{{{W}}}pgMar')
    if pm is not None:
        page = {}
        for attr in ['top', 'bottom', 'left', 'right', 'header', 'footer']:
            val = pm.get(f'{{{W}}}{attr}')
            if val is not None:
                page[attr] = int(val)
        if page:
            rules['page'] = page

    # 纸张大小
    pgSz = sects[0].find(f'{{{W}}}pgSz')
    if pgSz is not None:
        page = rules.get('page', {})
        w = pgSz.get(f'{{{W}}}w')
        h = pgSz.get(f'{{{W}}}h')
        if w:
            page['width'] = int(w)
        if h:
            page['height'] = int(h)
        if page:
            rules['page'] = page


# ═══════════════════════════════════════════════════════════════
# 通用工具函数
# ═══════════════════════════════════════════════════════════════

def _read_format(pPr, rPr) -> dict:
    """读取段落格式(pPr) + 字符格式(rPr)，返回 dict。

    只提取非 None 的属性。不读取任何文字内容。
    """
    fmt: Dict[str, Any] = {}

    if pPr is not None:
        jc = pPr.find(f'{{{W}}}jc')
        if jc is not None:
            fmt['alignment'] = jc.get(f'{{{W}}}val', 'left')

        sp = pPr.find(f'{{{W}}}spacing')
        if sp is not None:
            before = sp.get(f'{{{W}}}before')
            after = sp.get(f'{{{W}}}after')
            bl = sp.get(f'{{{W}}}beforeLines')
            al = sp.get(f'{{{W}}}afterLines')
            line = sp.get(f'{{{W}}}line')
            lineRule = sp.get(f'{{{W}}}lineRule')

            if before:
                fmt['before_twips'] = int(before)
            if after:
                fmt['after_twips'] = int(after)
            if bl:
                fmt['before_lines'] = int(bl) / 100
            if al:
                fmt['after_lines'] = int(al) / 100
            if line:
                fmt['line'] = int(line)
                fmt['lineRule'] = lineRule

        ind = pPr.find(f'{{{W}}}ind')
        if ind is not None:
            first_line = ind.get(f'{{{W}}}firstLine')
            if first_line:
                fmt['first_line_indent'] = int(first_line)
            left = ind.get(f'{{{W}}}left')
            if left:
                fmt['left_indent'] = int(left)
            right = ind.get(f'{{{W}}}right')
            if right:
                fmt['right_indent'] = int(right)

        ol = pPr.find(f'{{{W}}}outlineLvl')
        if ol is not None:
            fmt['outline_level'] = int(ol.get(f'{{{W}}}val'))

    if rPr is not None:
        rf = rPr.find(f'{{{W}}}rFonts')
        if rf is not None:
            east = rf.get(f'{{{W}}}eastAsia')
            ascii_ = rf.get(f'{{{W}}}ascii')
            hAnsi = rf.get(f'{{{W}}}hAnsi')
            if east:
                fmt['font_east'] = east
            if ascii_:
                fmt['font_ascii'] = ascii_
            if hAnsi:
                fmt['font_hAnsi'] = hAnsi

        sz = rPr.find(f'{{{W}}}sz')
        if sz is not None:
            fmt['font_size_pt'] = int(sz.get(f'{{{W}}}val')) / 2

        b = rPr.find(f'{{{W}}}b')
        if b is not None:
            val = b.get(f'{{{W}}}val', 'true')
            fmt['bold'] = val not in ('false', '0', 'off')

        i = rPr.find(f'{{{W}}}i')
        if i is not None:
            val = i.get(f'{{{W}}}val', 'true')
            fmt['italic'] = val not in ('false', '0', 'off')

        color = rPr.find(f'{{{W}}}color')
        if color is not None:
            cval = color.get(f'{{{W}}}val')
            if cval and cval.lower() != 'auto':
                fmt['color'] = cval

    return fmt


# ═══════════════════════════════════════════════════════════════
# 测试入口
# ═══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    import json
    spec = Path('output/附件7：毕业设计（论文）撰写规范.docx')
    if not spec.exists():
        print(f'Spec not found: {spec}')
    else:
        rules = read_spec_v2(spec)
        print(json.dumps(rules, indent=2, ensure_ascii=False, default=str))
