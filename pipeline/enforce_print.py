"""
流水线步骤3: 打印版强制奇数页

使用唯一标记法检测页码 + 简化空白页插入。
合并原来 enforce_v49/v51/v52/v53 的最优方案。
"""
import re, sys, zipfile
from pathlib import Path
from lxml import etree
import fitz

# Archived modules (generator) path

from .config import PipelineConfig

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def run(config: PipelineConfig):
    """为打印版插入空白页，确保每节从奇数页开始"""
    src = Path(config.input_docx)
    print('[enforce_print] Starting from', src)
    
    # ═══ Step 1: 唯一标记 → PDF → 检测页码 ═══
    with zipfile.ZipFile(src) as z:
        raw = z.read('word/document.xml')
    root = etree.fromstring(raw)
    body = root.find(f'{{{W}}}body')
    kids = list(body)
    
    # 找所有 sectPr 段落
    sectpr_idx = []
    for i, c in enumerate(kids):
        if c.tag == f'{{{W}}}p':
            pp = c.find(f'{{{W}}}pPr')
            if pp is not None and pp.find(f'{{{W}}}sectPr') is not None:
                sectpr_idx.append(i)
    
    print(f'[enforce_print] Found {len(sectpr_idx)} sections')
    
    # 插入唯一标记
    for si in range(len(sectpr_idx) - 1, -1, -1):
        sp_i = sectpr_idx[si]
        for j in range(sp_i + 1, len(kids)):
            if kids[j].tag == f'{{{W}}}p':
                r = etree.Element(f'{{{W}}}r')
                rPr = etree.SubElement(r, f'{{{W}}}rPr')
                etree.SubElement(rPr, f'{{{W}}}sz', {f'{{{W}}}val': '2'})
                etree.SubElement(rPr, f'{{{W}}}color', {f'{{{W}}}val': 'FFFFFF'})
                t = etree.SubElement(r, f'{{{W}}}t')
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t.text = f'SEC{si}'
                kids[j].insert(0, r)
                break
    
    marked_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    tmp = src.parent / '_enforce_marked.docx'
    with zipfile.ZipFile(src) as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = marked_xml
                zout.writestr(item, data)
    
    # PDF 转换
    try:
        from generator.pdf_converter import convert_to_pdf
        pdf_path = src.parent / '_enforce_marked.pdf'
        convert_to_pdf(tmp, pdf_path, 60)
    except Exception:
        import subprocess
        subprocess.run([config.libreoffice_path, '--headless',
                       '--convert-to', 'pdf', '--outdir', str(src.parent), str(tmp)],
                      capture_output=True, timeout=60, check=True)
        pdf_path = src.parent / '_enforce_marked.pdf'
    
    # 检测页码
    d = fitz.open(str(pdf_path))
    sp = {}
    for pn in range(d.page_count):
        for block in d[pn].get_text('blocks'):
            for m in re.finditer(r'SEC(\d+)', block[4]):
                i = int(m.group(1))
                if i not in sp:
                    sp[i] = pn + 1
    d.close()
    print(f'[enforce_print] Section start pages: {sp}')
    
    # ═══ Step 2: 计算需要空白页的节 ═══
    needs = []
    shift = 0
    for i in range(len(sectpr_idx)):
        p = sp.get(i)
        if p is None:
            continue
        eff = p + shift
        if eff % 2 == 0:
            needs.append(i)
            shift += 1
    
    print(f'[enforce_print] {len(needs)} sections need blanks')
    
    if not needs:
        print('[enforce_print] All sections already start on odd pages. Done.')
        return
    
    # ═══ Step 3: 提取 sectPr + 插入空白页 ═══
    with zipfile.ZipFile(src) as z:
        raw2 = z.read('word/document.xml')
    root2 = etree.fromstring(raw2)
    body2 = root2.find(f'{{{W}}}body')
    kids2 = list(body2)
    
    # 清除标记
    for p in root2.iter(f'{{{W}}}p'):
        for r in list(p.findall(f'{{{W}}}r')):
            for t in r.findall(f'{{{W}}}t'):
                if t.text and 'SEC' in str(t.text):
                    p.remove(r)
                    break
    
    # 找 sectPr
    sectpr_paras = []
    for i, c in enumerate(kids2):
        if c.tag == f'{{{W}}}p':
            pp = c.find(f'{{{W}}}pPr')
            if pp is not None:
                sp_el = pp.find(f'{{{W}}}sectPr')
                if sp_el is not None:
                    sectpr_paras.append((i, c, pp, sp_el))
    
    # 从后往前插入
    for sec_i in reversed(needs):
        if sec_i == 0:
            continue
        para_idx, para, pPr, sectPr = sectpr_paras[sec_i]
        
        # 检查前一段是否有分页符
        has_br = False
        for chk in kids2[max(0, para_idx - 3):para_idx]:
            if chk.tag == f'{{{W}}}p' and chk.find(f'.//{{{W}}}br') is not None:
                has_br = True
                break
        
        # 提取 sectPr
        pPr.remove(sectPr)
        new_sp = etree.Element(f'{{{W}}}p')
        new_pPr = etree.SubElement(new_sp, f'{{{W}}}pPr')
        new_pPr.append(sectPr)
        
        if has_br:
            # 已有分页符，只补空段
            ep = etree.Element(f'{{{W}}}p')
            r_ep = etree.SubElement(ep, f'{{{W}}}r')
            etree.SubElement(r_ep, f'{{{W}}}t',
                           {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' '
            body2.insert(para_idx + 1, new_sp)
            body2.insert(para_idx + 1, ep)
        else:
            # 分页符段 + 空段 + 新分节符段
            pb = etree.Element(f'{{{W}}}p')
            r_br = etree.SubElement(pb, f'{{{W}}}r')
            etree.SubElement(r_br, f'{{{W}}}br', {f'{{{W}}}type': 'page'})
            r_sp = etree.SubElement(pb, f'{{{W}}}r')
            etree.SubElement(r_sp, f'{{{W}}}t',
                           {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' '
            ep = etree.Element(f'{{{W}}}p')
            r_ep = etree.SubElement(ep, f'{{{W}}}r')
            etree.SubElement(r_ep, f'{{{W}}}t',
                           {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' '
            body2.insert(para_idx + 1, new_sp)
            body2.insert(para_idx + 1, ep)
            body2.insert(para_idx + 1, pb)
    
    new_xml = etree.tostring(root2, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(src) as zin:
        with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = new_xml
                zout.writestr(item, data)
    
    print(f'[enforce_print] Blank pages inserted. Saved to {src}')
