"""
从规范模板 docx 自动读取格式规则

替代 pipeline/config.py 中的硬编码格式值。
"""
import zipfile
from pathlib import Path
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def read_spec(spec_path: Path) -> dict:
    """从规范模板读取格式规则
    
    Returns:
        dict with keys: h1, h2, h3, body, blank_line, page
        Each value is a dict of format attributes.
    """
    if not spec_path or not spec_path.exists():
        return {}
    
    with zipfile.ZipFile(spec_path) as z:
        styles_xml = z.read('word/styles.xml') if 'word/styles.xml' in z.namelist() else None
        doc_xml = z.read('word/document.xml') if 'word/document.xml' in z.namelist() else None
    
    rules = {}
    
    if styles_xml:
        _extract_heading_rules(etree.fromstring(styles_xml), rules)
    
    if doc_xml:
        _extract_section_rules(etree.fromstring(doc_xml), rules)
        _extract_page_setup(etree.fromstring(doc_xml), rules)
    
    return rules


def _extract_heading_rules(styles_root, rules: dict):
    """从 styles.xml 提取各级标题格式"""
    for style in styles_root.iter(f'{{{W}}}style'):
        style_id = style.get(f'{{{W}}}styleId', '')
        pPr = style.find(f'{{{W}}}pPr')
        rPr = style.find(f'{{{W}}}rPr')
        
        level = None
        if pPr is not None:
            ol = pPr.find(f'{{{W}}}outlineLvl')
            if ol is not None:
                level = int(ol.get(f'{{{W}}}val'))
        
        if level is None:
            continue
        
        fmt = _read_format(pPr, rPr)
        
        key = {0: 'h1', 1: 'h2', 2: 'h3'}.get(level)
        if key:
            rules[key] = fmt
            rules[f'{key}_style_id'] = style_id
    
    # Body format from default paragraph style
    for style in styles_root.iter(f'{{{W}}}style'):
        if style.get(f'{{{W}}}type') == 'paragraph' and style.get(f'{{{W}}}default') == '1':
            rules['body_style_id'] = style.get(f'{{{W}}}styleId', 'a8')
            pPr = style.find(f'{{{W}}}pPr')
            rPr = style.find(f'{{{W}}}rPr')
            rules['body'] = _read_format(pPr, rPr)
            break


def _extract_section_rules(doc_root, rules: dict):
    """从 document.xml 提取节信息"""
    sects = list(doc_root.iter(f'{{{W}}}sectPr'))
    rules['section_count'] = len(sects)
    
    # 检测每个 section 的页眉文字
    header_texts = []
    for sp in sects:
        hdrs = {}
        for hr in sp.findall(f'{{{W}}}headerReference'):
            typ = hr.get(f'{{{W}}}type', 'default')
            # 页眉文字无法直接从 XML 读取（需要读实际 header 文件）
            # 这里只记录引用，由 setup_headers 填充
            hdrs[typ] = hr.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        header_texts.append(hdrs)
    rules['section_headers_raw'] = header_texts


def _extract_page_setup(doc_root, rules: dict):
    """提取页面设置"""
    sects = list(doc_root.iter(f'{{{W}}}sectPr'))
    if not sects:
        return
    
    pm = sects[0].find(f'{{{W}}}pgMar')
    if pm is not None:
        rules['page'] = {
            'top': int(pm.get(f'{{{W}}}top', '1418')),
            'bottom': int(pm.get(f'{{{W}}}bottom', '850')),
            'left': int(pm.get(f'{{{W}}}left', '1587')),
            'right': int(pm.get(f'{{{W}}}right', '1134')),
            'header': int(pm.get(f'{{{W}}}header', '1134')),
            'footer': int(pm.get(f'{{{W}}}footer', '567')),
        }


def _read_format(pPr, rPr) -> dict:
    """读取段落+字符格式"""
    fmt = {}
    
    if pPr is not None:
        jc = pPr.find(f'{{{W}}}jc')
        if jc is not None:
            fmt['alignment'] = jc.get(f'{{{W}}}val', 'left')
        
        sp = pPr.find(f'{{{W}}}spacing')
        if sp is not None:
            bl = sp.get(f'{{{W}}}beforeLines')
            al = sp.get(f'{{{W}}}afterLines')
            line = sp.get(f'{{{W}}}line')
            rule = sp.get(f'{{{W}}}lineRule')
            if bl:
                fmt['before_lines'] = int(bl) / 100
            if al:
                fmt['after_lines'] = int(al) / 100
            if line:
                fmt['line'] = int(line)
                fmt['lineRule'] = rule
    
    if rPr is not None:
        rf = rPr.find(f'{{{W}}}rFonts')
        if rf is not None:
            fmt['font_east'] = rf.get(f'{{{W}}}eastAsia', '宋体')
            fmt['font_ascii'] = rf.get(f'{{{W}}}ascii', 'Times New Roman')
        
        sz = rPr.find(f'{{{W}}}sz')
        if sz is not None:
            fmt['font_size_pt'] = int(sz.get(f'{{{W}}}val')) / 2
        
        fmt['bold'] = rPr.find(f'{{{W}}}b') is not None
    
    return fmt


def merge_with_config(rules: dict, config) -> dict:
    """合并规范模板规则 + 用户配置（用户配置优先）"""
    for key in ['h1', 'h2', 'h3', 'body']:
        if key in rules:
            spec_rules = rules[key]
            config_attr = getattr(config, key, None)
            if config_attr:
                # 用户通过对话设置的优先
                for attr in ['font_east', 'font_ascii', 'font_size_pt', 'bold',
                            'alignment', 'before_lines', 'after_lines']:
                    user_val = getattr(config_attr, attr, None)
                    default_val = getattr(type(config_attr)(), attr, None)
                    if user_val != default_val:  # 用户改过
                        spec_rules[attr] = user_val
    
    return rules
