"""
Reconciler v2 — 对比规范规则与当前状态，生成差异修复计划

核心原则:
- 只输出需要修改的差异
- 保留不动的不输出
- 页眉文字内容保持不动（除非规范规则说字体需要改）
- 规范模板缺的属性标 null，Reconciler 跳过不检查

用法:
    from pipeline.spec_reader_v2 import read_spec_v2
    from pipeline.doc_discover import discover_doc
    from pipeline.reconciler_v2 import reconcile

    spec_rules = read_spec_v2(spec_path)
    doc_state = discover_doc(target_path)
    fix_plan = reconcile(spec_rules, doc_state)

fix_plan 格式:
    [
        {"section": 3, "action": "set_sectpr_type", "to": "nextPage"},
        {"section": 5, "action": "set_body_font_ascii", "to": "Times New Roman"},
        {"section": None, "action": "set_style", "style_id": "a8", "font_ascii": "Times New Roman"},
        ...
    ]
"""

import re
from typing import Optional


def reconcile(spec_rules: dict, doc_state: dict) -> list[dict]:
    """对比规范规则与文档当前状态，生成修复计划。

    Args:
        spec_rules: 规范模板的格式规则。期望结构:

            {
                "sections": {
                    "front_matter": {
                        "header": {
                            "font_east": "宋体",       # None 表示不检查
                            "font_ascii": "Times New Roman",
                            "font_size_pt": 9.0,
                            "underline": True,
                        },
                        "footer": {
                            "has_page": True,
                            "page_format": "upperRoman",  # decimal / upperRoman / lowerRoman
                        },
                        "sectpr_type": "nextPage",      # None 表示不检查
                        "even_header": True,             # 是否需要奇偶页页眉
                    },
                    "body": { ... },
                    "back_matter": { ... },
                },
                "body": {
                    "font_east": "宋体",
                    "font_ascii": "Times New Roman",
                    "font_size_pt": 12.0,
                },
                "body_style_id": "a8",
                "h1": { ... },   # 标题格式规则（可选）
                "h2": { ... },
                "h3": { ... },
                "page": { ... },  # 页面设置
            }

        doc_state: DocDiscover 输出的当前状态。结构:

            {
                "sections": [
                    {
                        "index": 0,
                        "heading": "摘要",
                        "category": "front_matter",
                        "current_header_text": "...",
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
                "headings": [...]
            }

    Returns:
        list[dict]: 修复计划，每个元素是一个修复动作。
    """
    fix_plan = []
    sections = doc_state.get('sections', [])

    # ── 0. 从扁平 spec_rules 构建节规范（适配 spec_reader_v2 输出）──
    spec_sections = _build_section_specs(spec_rules)

    for sec in sections:
        idx = sec['index']
        cat = sec.get('category', 'body')

        # 获取该类别规范
        cat_spec = spec_sections.get(cat, {})
        if not cat_spec:
            continue

        # ── 1a. 页眉格式 ──
        _reconcile_header(sec, cat_spec, idx, fix_plan)

        # ── 1b. 分节符类型 ──
        _reconcile_sectpr_type(sec, cat_spec, idx, fix_plan)

        # ── 1c. 页脚页码 ──
        _reconcile_footer(sec, cat_spec, idx, fix_plan)

        # ── 1d. 奇偶页页眉 ──
        _reconcile_even_header(sec, cat_spec, idx, fix_plan)

    # ── 2. 全局样式检查 ────────────────────────────────
    _reconcile_body_styles(spec_rules, doc_state, fix_plan)
    _reconcile_heading_styles(spec_rules, fix_plan)
    _reconcile_page_setup(spec_rules, doc_state, fix_plan)

    return fix_plan


# ══════════════════════════════════════════════════════════════
# 内部：逐项对比函数
# ══════════════════════════════════════════════════════════════

def _reconcile_header(sec: dict, cat_spec: dict, idx: int, fix_plan: list):
    """对比页眉格式：字体、字号、下划线。不改变文字内容。"""
    header_spec = cat_spec.get('header', {})
    if not header_spec:
        return

    # 中文字体
    spec_font_east = header_spec.get('font_east')
    if spec_font_east is not None:
        cur = sec.get('current_header_font_east')
        if cur != spec_font_east:
            fix_plan.append({
                'section': idx,
                'action': 'set_header_font_east',
                'to': spec_font_east,
                'current': cur,
            })

    # 西文字体
    spec_font_ascii = header_spec.get('font_ascii')
    if spec_font_ascii is not None:
        cur = sec.get('current_header_font_ascii')
        if cur != spec_font_ascii:
            fix_plan.append({
                'section': idx,
                'action': 'set_header_font_ascii',
                'to': spec_font_ascii,
                'current': cur,
            })

    # 字号
    spec_size = header_spec.get('font_size_pt')
    if spec_size is not None:
        cur = sec.get('current_header_font_size_pt')
        # 浮点比较
        if cur is None or abs(cur - spec_size) > 0.01:
            fix_plan.append({
                'section': idx,
                'action': 'set_header_font_size',
                'to': spec_size,
                'current': cur,
            })

    # 下划线（pBdr bottom border）
    spec_underline = header_spec.get('underline')
    if spec_underline is not None:
        cur = sec.get('current_header_underline', False)
        if cur != spec_underline:
            fix_plan.append({
                'section': idx,
                'action': 'set_header_underline',
                'to': spec_underline,
                'current': cur,
            })


def _reconcile_sectpr_type(sec: dict, cat_spec: dict, idx: int, fix_plan: list):
    """对比分节符类型（nextPage / oddPage / evenPage / continuous）。"""
    spec_type = cat_spec.get('sectpr_type')
    if spec_type is None:
        return

    cur = sec.get('current_sectpr_type')
    if cur != spec_type:
        fix_plan.append({
            'section': idx,
            'action': 'set_sectpr_type',
            'to': spec_type,
            'current': cur,
        })


def _reconcile_footer(sec: dict, cat_spec: dict, idx: int, fix_plan: list):
    """对比页脚页码：是否有 PAGE 域、页码格式。"""
    footer_spec = cat_spec.get('footer', {})
    if not footer_spec:
        return

    # 是否需要页码
    spec_has_page = footer_spec.get('has_page')
    if spec_has_page is not None:
        cur_has = sec.get('current_footer_has_page', False)
        if cur_has != spec_has_page:
            if spec_has_page:
                page_fmt = footer_spec.get('page_format', 'decimal')
                fix_plan.append({
                    'section': idx,
                    'action': 'add_footer_page',
                    'format': page_fmt,
                    'current': cur_has,
                })
            else:
                fix_plan.append({
                    'section': idx,
                    'action': 'remove_footer_page',
                    'current': cur_has,
                })

    # 页码格式（仅当已有页码时检查）
    spec_format = footer_spec.get('page_format')
    if spec_format is not None and sec.get('current_footer_has_page'):
        cur_fmt = sec.get('current_footer_page_format')
        if cur_fmt and cur_fmt != spec_format:
            fix_plan.append({
                'section': idx,
                'action': 'set_page_format',
                'to': spec_format,
                'current': cur_fmt,
            })


def _reconcile_even_header(sec: dict, cat_spec: dict, idx: int, fix_plan: list):
    """对比是否需要奇偶页页眉。"""
    spec_even = cat_spec.get('even_header')
    if spec_even is None:
        return

    cur = sec.get('has_even_header', False)
    if cur != spec_even:
        fix_plan.append({
            'section': idx,
            'action': 'set_even_header_enabled',
            'to': spec_even,
            'current': cur,
        })


def _reconcile_body_styles(spec_rules: dict, doc_state: dict, fix_plan: list):
    """对比正文样式（body 字体等）。

    如果规范中有 body 字体定义，生成全局样式修改。
    """
    body_spec = spec_rules.get('body', {})
    if not body_spec:
        return

    action = {
        'section': None,
        'action': 'set_style',
        'style_id': spec_rules.get('body_style_id', 'a8'),
    }

    has_change = False

    # 注意：这里只能做全局对比。doc_state 中没有直接的 body 字体信息，
    # 因为 DocDiscover 主要关注节级别的页眉页脚。
    # body 字体通常从 styles.xml 读取，此处生成"应设"指令。
    for attr_key, action_key in [
        ('font_east', 'font_east'),
        ('font_ascii', 'font_ascii'),
        ('font_size_pt', 'font_size_pt'),
    ]:
        val = body_spec.get(attr_key)
        if val is not None:
            action[action_key] = val
            has_change = True

    if has_change:
        fix_plan.append(action)


def _reconcile_heading_styles(spec_rules: dict, fix_plan: list):
    """对比标题样式（H1/H2/H3 字体等）。

    如果规范中有标题字体定义，生成全局样式修改。
    """
    for level_key, outline_lvl in [('h1', 0), ('h2', 1), ('h3', 2)]:
        heading_spec = spec_rules.get(level_key, {})
        if not heading_spec:
            continue

        style_id = spec_rules.get(f'{level_key}_style_id')
        if not style_id:
            continue

        action = {
            'section': None,
            'action': 'set_style',
            'style_id': style_id,
            'outline_lvl': outline_lvl,
        }

        has_change = False
        for attr in ['font_east', 'font_ascii', 'font_size_pt', 'bold', 'alignment']:
            val = heading_spec.get(attr)
            if val is not None:
                action[attr] = val
                has_change = True

        if has_change:
            fix_plan.append(action)


def _reconcile_page_setup(spec_rules: dict, doc_state: dict, fix_plan: list):
    """对比页面设置（页边距等）。

    如果规范中有 page 设置，生成全局页面设置修改。
    """
    page_spec = spec_rules.get('page', {})
    if not page_spec:
        return

    action = {
        'section': None,
        'action': 'set_page_margins',
    }

    has_change = False
    for key in ['top', 'bottom', 'left', 'right', 'header', 'footer']:
        val = page_spec.get(key)
        if val is not None:
            action[key] = val
            has_change = True

    if has_change:
        fix_plan.append(action)


# ══════════════════════════════════════════════════════════════
# 辅助函数
# ══════════════════════════════════════════════════════════════

def _build_section_specs(spec_rules: dict) -> dict:
    """从扁平 spec_rules 构建 Reconciler 期望的嵌套节规范。

    spec_reader_v2 输出扁平结构:
        {section_break, even_and_odd_headers, header_format, footer_format, ...}
    
    转换为 Reconciler 期望的嵌套结构:
        {front_matter: {header:{..}, footer:{..}, sectpr_type, even_header},
         body: {...},
         back_matter: {...}}
    """
    cats = {}
    header_fmt = spec_rules.get('header_format', {})
    footer_fmt = spec_rules.get('footer_format', {})
    section_break = spec_rules.get('section_break')
    even_hdr = spec_rules.get('even_and_odd_headers')

    # Header spec: 只取格式属性，不取文字
    hdr_spec = {}
    if header_fmt:
        for k in ['font_east', 'font_ascii', 'font_size_pt', 'underline']:
            if k in header_fmt:
                hdr_spec[k] = header_fmt[k]

    # Footer spec
    ftr_spec = {}
    if footer_fmt:
        ftr_spec['has_page'] = True  # 有 footer XML 就认为应有页码
        if 'page_number_format' in footer_fmt:
            ftr_spec['page_format'] = 'decimal'

    base = {'sectpr_type': section_break, 'even_header': even_hdr}
    if hdr_spec:
        base['header'] = hdr_spec
    if ftr_spec:
        base['footer'] = ftr_spec

    # 三类章节共用相同格式规则
    for cat in ['front_matter', 'body', 'back_matter']:
        cats[cat] = dict(base)

    return cats


def print_fix_plan(fix_plan: list):
    """打印修复计划的可读摘要。"""
    if not fix_plan:
        print('✅ No fixes needed — document already matches spec.')
        return

    print(f'🔧 Fix Plan ({len(fix_plan)} actions):')
    print('-' * 60)

    for i, fix in enumerate(fix_plan):
        sec = fix.get('section')
        sec_str = f'sec {sec}' if sec is not None else 'GLOBAL'
        action = fix['action']
        to_val = fix.get('to', fix.get('format', ''))
        cur = fix.get('current', '?')

        emoji = _action_emoji(action)
        extra = ''
        if action == 'set_style':
            extra = f' style_id={fix.get("style_id","?")}'
            parts = []
            for k in ['font_east', 'font_ascii', 'font_size_pt', 'bold', 'alignment']:
                if k in fix:
                    parts.append(f'{k}={fix[k]}')
            if parts:
                extra += ' (' + ', '.join(parts) + ')'
        elif action == 'set_page_margins':
            parts = []
            for k in ['top', 'bottom', 'left', 'right', 'header', 'footer']:
                if k in fix:
                    parts.append(f'{k}={fix[k]}')
            extra = ' (' + ', '.join(parts) + ')'

        print(f'  {i+1}. {emoji} [{sec_str}] {action}: {cur} → {to_val}{extra}')

    print('-' * 60)


def _action_emoji(action: str) -> str:
    return {
        'set_header_font_east': '🔤',
        'set_header_font_ascii': '🔤',
        'set_header_font_size': '📏',
        'set_header_underline': '➖',
        'set_sectpr_type': '📄',
        'add_footer_page': '🔢',
        'remove_footer_page': '🚫',
        'set_page_format': '🔢',
        'set_even_header_enabled': '📑',
        'set_style': '🎨',
        'set_page_margins': '📐',
    }.get(action, '❓')


# ══════════════════════════════════════════════════════════════
# 便捷函数：一键发现 + 协调
# ══════════════════════════════════════════════════════════════

def discover_and_reconcile(target_path, spec_rules: dict) -> tuple[dict, list[dict]]:
    """便捷函数：一次调用完成发现 + 协调。

    Args:
        target_path: 目标 docx 文件路径
        spec_rules: 规范格式规则

    Returns:
        (doc_state, fix_plan)
    """
    from .doc_discover import discover_doc
    doc_state = discover_doc(target_path)
    fix_plan = reconcile(spec_rules, doc_state)
    return doc_state, fix_plan
