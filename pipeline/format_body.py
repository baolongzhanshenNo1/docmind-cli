"""
流水线步骤1: 正文格式 — 样式 + 段落 + 页顶空行

合并原来 fix_v6/v7/v8/v9 的功能。
"""
import re, sys
from pathlib import Path
from lxml import etree
import docx
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

# Archived modules (docmind3) path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / '_archive' / 'docmind'))
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .config import PipelineConfig

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _norm(s: str) -> str:
    """去除所有空格（解决 '摘  要' 匹配问题）"""
    return re.sub(r'\s+', '', s)


def _text_of(p) -> str:
    return ''.join(t.text or '' for t in p.iter(qn('w:t')))


def _ensure_rPr(run, **kwargs):
    """确保 run 有 rPr，并按需设置属性"""
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        rPr = etree.SubElement(run._element, qn('w:rPr'))
    
    if 'size' in kwargs:
        for tag in ['w:sz', 'w:szCs']:
            el = rPr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rPr, qn(tag))
            el.set(qn('w:val'), str(int(kwargs['size'] * 2)))  # pt → half-pt
    
    if 'bold' in kwargs:
        b = rPr.find(qn('w:b'))
        if kwargs['bold']:
            if b is None:
                etree.SubElement(rPr, qn('w:b'))
        else:
            if b is not None:
                rPr.remove(b)
    
    if 'font_east' in kwargs:
        rf = rPr.find(qn('w:rFonts'))
        if rf is None:
            rf = etree.SubElement(rPr, qn('w:rFonts'))
        rf.set(qn('w:eastAsia'), kwargs['font_east'])
    
    return rPr


def _spacing_el(before_lines=0, after_lines=0, line_spacing=1.0):
    """创建 <w:spacing> 元素"""
    attrs = {qn('w:snapToGrid'): '0'}
    if before_lines:
        attrs[qn('w:beforeLines')] = str(int(before_lines * 100))
    if after_lines:
        attrs[qn('w:afterLines')] = str(int(after_lines * 100))
    if line_spacing != 1.0:
        attrs[qn('w:line')] = str(int(line_spacing * 240))
        attrs[qn('w:lineRule')] = 'auto'
    return etree.Element(qn('w:spacing'), attrs)


def _jc_val(alignment: str) -> str:
    return {'left': 'left', 'center': 'center', 'right': 'right', 'both': 'both'}.get(alignment, 'left')


def run(config: PipelineConfig):
    """执行正文格式流水线
    
    1. 修改样式定义 (styles.xml)
    2. 修改段落级格式 (document.xml)
    3. 插入页顶空行
    """
    print('[format_body] Loading...')
    doc = docx.Document(str(config.input_docx))
    
    # ── Step 1: 样式定义 ──
    print('[format_body] Modifying styles...')
    
    # 从 spec_rules 或默认值获取格式
    rules = config.spec_rules
    h1_fmt = rules.get('h1', {'font_east':'黑体','font_ascii':'Times New Roman','font_size_pt':16,
                               'bold':True,'alignment':'center','before_lines':0,'after_lines':0.5,
                               'line':240,'lineRule':'auto'})
    h2_fmt = rules.get('h2', {'font_east':'黑体','font_ascii':'Times New Roman','font_size_pt':14,
                               'bold':True,'alignment':'left','before_lines':0.5,'after_lines':0.5,
                               'line':240,'lineRule':'auto'})
    h3_fmt = rules.get('h3', {'font_east':'宋体','font_ascii':'Times New Roman','font_size_pt':14,
                               'bold':False,'alignment':'left','before_lines':0.5,'after_lines':0,
                               'line':240,'lineRule':'auto'})
    body_fmt = rules.get('body', {'font_east':'宋体','font_ascii':'Times New Roman','font_size_pt':12,
                                   'bold':False,'alignment':'both','before_lines':0,'after_lines':0,
                                   'line':360,'lineRule':'auto'})
    
    for style_id, fmt in [
        (config.style_ids.get('h1','1'), h1_fmt),
        (config.style_ids.get('h2','20'), h2_fmt),
        (config.style_ids.get('body','a8'), body_fmt),
    ]:
        try:
            style = doc.styles[style_id]
        except KeyError:
            print(f'  WARNING: style {style_id} not found')
            continue
        
        # 字体
        rPr = style.element.find(qn('w:rPr'))
        if rPr is None:
            rPr = etree.SubElement(style.element, qn('w:rPr'))
        
        rf = rPr.find(qn('w:rFonts'))
        if rf is None:
            rf = etree.SubElement(rPr, qn('w:rFonts'))
        rf.set(qn('w:eastAsia'), fmt.get("font_east","宋体"))
        rf.set(qn('w:ascii'), fmt.get("font_ascii","Times New Roman"))
        
        for tag in ['w:sz', 'w:szCs']:
            el = rPr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rPr, qn(tag))
            el.set(qn('w:val'), str(int(fmt.get("font_size_pt",12) * 2)))
        
        if fmt.get("bold",False):
            b = rPr.find(qn('w:b'))
            if b is None:
                etree.SubElement(rPr, qn('w:b'))
        
        # 段落属性
        pPr = style.element.find(qn('w:pPr'))
        if pPr is None:
            pPr = etree.SubElement(style.element, qn('w:pPr'))
        
        jc = pPr.find(qn('w:jc'))
        if jc is None:
            jc = etree.SubElement(pPr, qn('w:jc'))
        jc.set(qn('w:val'), _jc_val(fmt.get("alignment","left")))
        
        old_sp = pPr.find(qn('w:spacing'))
        if old_sp is not None:
            pPr.remove(old_sp)
        pPr.append(_spacing_el(fmt.get("before_lines",0), fmt.get("after_lines",0), fmt.get("line",240)/240))
    
    # ── Step 2: 段落级格式 ──
    print('[format_body] Formatting paragraphs...')
    body = doc.element.body
    h3_style_id = config.style_ids['h2']  # H3 使用 H2 的样式ID
    
    h2_count = h3_count = body_count = 0
    
    for p in body.iter(qn('w:p')):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            continue
        
        ol = pPr.find(qn('w:outlineLvl'))
        level = int(ol.get(qn('w:val'))) if ol is not None else None
        
        # H2 段落级间距
        if level == 1:
            old_sp = pPr.find(qn('w:spacing'))
            if old_sp is not None:
                pPr.remove(old_sp)
            pPr.append(_spacing_el(h2_fmt.get('before_lines',0), h2_fmt.get('after_lines',0), 
                                   h2_fmt.get('line',240)/240))
            h2_count += 1
        
        # H3 段落级格式
        elif level == 2:
            # 字号 + 不加粗 + 间距
            for r in p.findall(qn('w:r')):
                _ensure_rPr(docx.text.run.Run(r, None), size=h3_fmt.get('font_size_pt',14), 
                           bold=h3_fmt.get('bold',False), font_east=h3_fmt.get('font_east','宋体'))
            old_sp = pPr.find(qn('w:spacing'))
            if old_sp is not None:
                pPr.remove(old_sp)
            pPr.append(_spacing_el(h3_fmt.get('before_lines',0), h3_fmt.get('after_lines',0),
                                   h3_fmt.get('line',240)/240))
            h3_count += 1
        
        # Body 清除段落级间距（让样式生效）
        elif level is None:
            pStyle = pPr.find(qn('w:pStyle'))
            sid = pStyle.get(qn('w:val')) if pStyle is not None else None
            if sid == config.style_ids['body']:
                old_sp = pPr.find(qn('w:spacing'))
                if old_sp is not None:
                    pPr.remove(old_sp)
                    body_count += 1
    
    print(f'[format_body] H2={h2_count} H3={h3_count} Body={body_count}')
    
    # ── Step 3: 页顶空行 ──
    print('[format_body] Adding blank lines...')
    bl_count = insert_blank_lines(doc, config)
    print(f'[format_body] Blank lines: {bl_count}')
    
    doc.save(str(config.input_docx))
    print(f'[format_body] Saved to {config.input_docx}')


def insert_blank_lines(doc, config: PipelineConfig) -> int:
    """在每个 H1 标题前插入页顶空行（使用 Reader 发现标题）"""
    from docmind3 import DocxReader
    
    # 用 Reader 解析当前文档
    reader = DocxReader()
    doc_model = reader.read(str(config.input_docx))
    
    body = doc.element.body
    # 从 spec_rules 或默认值获取空行格式
    bl_line_pt = config.spec_rules.get('page', {}).get('line_height', 20)  # 固定20磅
    bl_font_pt = config.spec_rules.get('body', {}).get('font_size_pt', 12)  # 小四
    
    # 收集所有 H1 标题的已规范化文字
    h1_texts = set()
    for h in doc_model.all_headings:
        if h.level == 1:
            h1_texts.add(_norm(h.text))
    
    count = 0
    for child in list(body):
        if child.tag != qn('w:p'):
            continue
        
        t = _text_of(child)
        norm = _norm(t)
        
        if norm not in h1_texts:
            continue
        
        # 创建空行段落
        bp = etree.Element(qn('w:p'))
        bpP = etree.SubElement(bp, qn('w:pPr'))
        etree.SubElement(bpP, qn('w:spacing'), {
            qn('w:line'): str(int(bl_line_pt * 20)),
            qn('w:lineRule'): 'exact',
            qn('w:snapToGrid'): '0',
        })
        br_r = etree.SubElement(bp, qn('w:r'))
        br_rP = etree.SubElement(br_r, qn('w:rPr'))
        etree.SubElement(br_rP, qn('w:sz'), {qn('w:val'): str(int(bl_font_pt * 2))})
        t_el = etree.SubElement(br_r, qn('w:t'))
        t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t_el.text = ' '
        
        child.addprevious(bp)
        count += 1
    
    return count
