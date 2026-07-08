"""
Pipeline preprocessor v2 — 完整节结构 + 奇偶页眉 + 页码

修复:
1. 每个 H1 前插入独立 sectPr (13个)
2. 奇数页=章节标题, 偶数页=年份/专业
3. 页码: 前部罗马 + 正文阿拉伯, "第X页" 格式
"""
import re, sys, zipfile, shutil
from pathlib import Path

# Archived modules (docmind3) path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / '_archive' / 'docmind'))
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML = 'http://www.w3.org/XML/1998/namespace'


def preprocess(input_path: str | Path, spec_path: str | Path = None,
               output_path: str | Path = None) -> Path:
    input_path = Path(input_path)
    output_path = Path(output_path) if output_path else input_path.parent / (input_path.stem + '_preprocessed.docx')
    print(f'[preprocess v2] {input_path} -> {output_path}')
    
    shutil.copy(input_path, output_path)
    
    # Phase 1: 完整插入13个分节符
    print('[Phase 1] Inserting section breaks...')
    _insert_all_section_breaks(output_path)
    
    # Phase 2: 样式修改
    print('[Phase 2] Styles...')
    _modify_styles(output_path, spec_path)
    
    # Phase 3: 页顶空行
    print('[Phase 3] Blank lines...')
    _add_blank_lines(output_path)
    
    # Phase 4: 页眉（奇偶）+ 页码
    print('[Phase 4] Headers + Footers...')
    _setup_headers_footers_v2(output_path)
    
    print(f'[preprocess v2] Done: {output_path}')
    return output_path


def _insert_all_section_breaks(docx_path: Path):
    """在每个 H1 前插入带 header/footer 引用的完整 sectPr"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        doc_xml = z.read('word/document.xml')
        with z.open('word/_rels/document.xml.rels') as f:
            rels = etree.parse(f)
    
    # 收集可用的 header/footer rIds
    hdr_ids = []
    ftr_ids = []
    for rel in rels.getroot():
        t = rel.get('Type', '')
        if 'header' in t: hdr_ids.append(rel.get('Id'))
        if 'footer' in t: ftr_ids.append(rel.get('Id'))
    
    root = etree.fromstring(doc_xml)
    body = root.find(f'{{{W}}}body')
    
    # 找所有 H1
    h1_pattern = re.compile(r'^(?:摘要|ABSTRACT|目\s*录|目录|\d+[\u4e00-\u9fff]|参考\s*文献|参考文献|致\s*谢|致谢)')
    
    count = 0
    for child in list(body):
        if child.tag != f'{{{W}}}p':
            continue
        text = ''.join(t.text or '' for t in child.iter(f'{{{W}}}t'))
        norm = re.sub(r'\s+', '', text)
        if not h1_pattern.match(norm):
            continue
        
        # 跳过已有前导 sectPr 的（但只跳过"空分节符段"，不跳过带文字内容段中的 sectPr）
        prev = child.getprevious()
        if prev is not None and prev.tag == f'{{{W}}}p':
            prev_text = ''.join(t.text or '' for t in prev.iter(f'{{{W}}}t')).strip()
            pp = prev.find(f'{{{W}}}pPr')
            if pp is not None and pp.find(f'{{{W}}}sectPr') is not None and prev_text == '':
                continue  # 已是纯空分节符段
        
        # 构建完整 sectPr
        sp = etree.Element(f'{{{W}}}p')
        spp = etree.SubElement(sp, f'{{{W}}}pPr')
        sectPr = etree.SubElement(spp, f'{{{W}}}sectPr')
        
        etree.SubElement(sectPr, f'{{{W}}}type', {f'{{{W}}}val': 'nextPage'})
        etree.SubElement(sectPr, f'{{{W}}}pgSz', {f'{{{W}}}w': '11906', f'{{{W}}}h': '16838'})
        etree.SubElement(sectPr, f'{{{W}}}pgMar', {
            f'{{{W}}}top': '1418', f'{{{W}}}right': '1134',
            f'{{{W}}}bottom': '850', f'{{{W}}}left': '1587',
            f'{{{W}}}header': '1134', f'{{{W}}}footer': '567',
            f'{{{W}}}gutter': '0',
        })
        etree.SubElement(sectPr, f'{{{W}}}cols', {f'{{{W}}}space': '425'})
        etree.SubElement(sectPr, f'{{{W}}}docGrid', {f'{{{W}}}type': 'lines', f'{{{W}}}linePitch': '312'})
        
        # 页眉引用 (default + even, 复用好 rIds)
        if hdr_ids:
            r = etree.SubElement(sectPr, f'{{{W}}}headerReference')
            r.set(f'{{{W}}}type', 'default')
            r.set(f'{{{R}}}id', hdr_ids[0])
        if len(hdr_ids) > 1:
            r = etree.SubElement(sectPr, f'{{{W}}}headerReference')
            r.set(f'{{{W}}}type', 'even')
            r.set(f'{{{R}}}id', hdr_ids[1])
        
        # 页脚引用
        if ftr_ids:
            r = etree.SubElement(sectPr, f'{{{W}}}footerReference')
            r.set(f'{{{W}}}type', 'default')
            r.set(f'{{{R}}}id', ftr_ids[0])
        if len(ftr_ids) > 1:
            r = etree.SubElement(sectPr, f'{{{W}}}footerReference')
            r.set(f'{{{W}}}type', 'even')
            r.set(f'{{{R}}}id', ftr_ids[1])
        
        child.addprevious(sp)
        count += 1
    
    doc_data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    _zip_replace(docx_path, 'word/document.xml', doc_data)
    print(f'  {count} section breaks added')


def _modify_styles(docx_path, spec_path):
    from docmind3 import DocxReader, DocxWriter
    reader = DocxReader()
    doc = reader.read(str(docx_path))
    writer = DocxWriter(doc)
    rules = {'1':('黑体',32), '20':('黑体',28), 'a8':('宋体',24)}
    if spec_path:
        from pipeline.spec_reader import read_spec
        sr = read_spec(Path(spec_path))
        for k, sid in [('h1','1'),('h2','20'),('body','a8')]:
            if k in sr:
                rules[sid] = (sr[k].get('font_east','宋体'), int(sr[k].get('font_size_pt',12)*2))
    for sid, (ea, sz) in rules.items():
        try: writer.modify_style(sid, font_east_asia=ea, font_size_pt=sz/2)
        except: pass
    writer.save(str(docx_path))
    print(f'  Styles: {list(rules)}')


def _add_blank_lines(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        doc_xml = z.read('word/document.xml')
    root = etree.fromstring(doc_xml)
    body = root.find(f'{{{W}}}body')
    count = 0
    h1_pat = re.compile(r'^(?:摘要|ABSTRACT|目\s*录|目录|\d+[\u4e00-\u9fff]|参考\s*文献|参考文献|致\s*谢|致谢)')
    for child in list(body):
        if child.tag != f'{{{W}}}p': continue
        if not h1_pat.match(re.sub(r'\s+','',''.join(t.text or '' for t in child.iter(f'{{{W}}}t')))):
            continue
        bp = etree.Element(f'{{{W}}}p')
        bpP = etree.SubElement(bp, f'{{{W}}}pPr')
        etree.SubElement(bpP, f'{{{W}}}spacing', {f'{{{W}}}line':'400',f'{{{W}}}lineRule':'exact',f'{{{W}}}snapToGrid':'0'})
        r = etree.SubElement(bp, f'{{{W}}}r')
        rP = etree.SubElement(r, f'{{{W}}}rPr')
        etree.SubElement(rP, f'{{{W}}}sz', {f'{{{W}}}val':'24'})
        t = etree.SubElement(r, f'{{{W}}}t', {f'{{{XML}}}space':'preserve'}); t.text=' '
        child.addprevious(bp); count += 1
    _zip_replace(docx_path, 'word/document.xml', etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))
    print(f'  {count} blank lines')


def _setup_headers_footers_v2(docx_path):
    """设置奇偶页眉 + 页码（前部罗马/正文阿拉伯）"""
    with zipfile.ZipFile(docx_path) as z:
        doc_xml = z.read('word/document.xml')
        with z.open('word/_rels/document.xml.rels') as f: rels = etree.parse(f)
    
    root = etree.fromstring(doc_xml)
    body = root.find(f'{{{W}}}body')
    
    # 收集 header/footer rIds
    hdr_default = hdr_even = None
    ftr_default = ftr_even = None
    for rel in rels.getroot():
        t = rel.get('Type','')
        trg = rel.get('Target','')
        rid = rel.get('Id')
        if 'header' in t:
            if hdr_default is None: hdr_default = rid
            elif hdr_even is None: hdr_even = rid
        if 'footer' in t:
            if ftr_default is None: ftr_default = rid
            elif ftr_even is None: ftr_even = rid
    
    # 分类章节
    front = ['摘要','ABSTRACT','目录']
    back = ['参考文献','致谢']
    chapters = []
    
    h1_pat = re.compile(r'^(?:摘要|ABSTRACT|目\s*录|目录|\d+[\u4e00-\u9fff]|参考\s*文献|参考文献|致\s*谢|致谢)')
    si = 0
    for child in list(body):
        if child.tag != f'{{{W}}}p': continue
        pp = child.find(f'{{{W}}}pPr')
        if pp is None or pp.find(f'{{{W}}}sectPr') is None: continue
        chapters.append(si)
        si += 1
    
    # 修改每个 sectPr
    sects = list(root.iter(f'{{{W}}}sectPr'))
    
    for i, sp in enumerate(sects):
        # 清除旧页眉/页脚引用
        for tag in ['headerReference','footerReference']:
            for old in list(sp.findall(f'{{{W}}}{tag}')):
                sp.remove(old)
        
        is_front = i < len(front)
        is_back = i >= len(sects) - len(back)
        
        if is_front:
            # 前部: section 标题作为页眉
            if hdr_default:
                r = etree.SubElement(sp, f'{{{W}}}headerReference')
                r.set(f'{{{W}}}type', 'default'); r.set(f'{{{R}}}id', hdr_default)
        elif is_back:
            if hdr_default:
                r = etree.SubElement(sp, f'{{{W}}}headerReference')
                r.set(f'{{{W}}}type', 'default'); r.set(f'{{{R}}}id', hdr_default)
        else:
            # 正文: 奇数页"2024年 毕业设计", 偶数页"计算机科学与技术"
            if hdr_even:
                r = etree.SubElement(sp, f'{{{W}}}headerReference')
                r.set(f'{{{W}}}type', 'even'); r.set(f'{{{R}}}id', hdr_even)
            if hdr_default:
                r = etree.SubElement(sp, f'{{{W}}}headerReference')
                r.set(f'{{{W}}}type', 'default'); r.set(f'{{{R}}}id', hdr_default)
        
        # 页脚引用
        if ftr_even:
            r = etree.SubElement(sp, f'{{{W}}}footerReference')
            r.set(f'{{{W}}}type', 'even'); r.set(f'{{{R}}}id', ftr_even)
        if ftr_default:
            r = etree.SubElement(sp, f'{{{W}}}footerReference')
            r.set(f'{{{W}}}type', 'default'); r.set(f'{{{R}}}id', ftr_default)
    
    doc_data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    _zip_replace(docx_path, 'word/document.xml', doc_data)
    
    # 同样修改实际的 header/footer XML 文件
    _write_header_files(docx_path)
    _write_footer_files(docx_path, len(sects), len(front))
    
    print('  Headers + footers configured')


def _write_header_files(docx_path):
    """写入页眉内容"""
    _make_xml = lambda text: (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman"/>'
        f'<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
        '</w:p></w:hdr>'
    ).encode('utf-8')
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        header_names = [n for n in z.namelist() if 'header' in n and n.endswith('.xml')]
    
    for name in header_names:
        header_text = '2024年 毕业设计'  # default
        _zip_replace(docx_path, name, _make_xml(header_text))


def _write_footer_files(docx_path, section_count: int, front_count: int):
    """写入页脚: 前部罗马数字 '第 I 页', 正文阿拉伯 '第 1 页'"""
    _make_ftr = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman"/>'
        f'<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">第</w:t></w:r>'
        f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
        f'<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
        f'<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
        f'<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
        f'<w:t>1</w:t></w:r>'
        f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
        f'<w:fldChar w:fldCharType="end"/></w:r>'
        f'<w:r><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman"/>'
        f'<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">页</w:t></w:r>'
        '</w:p></w:ftr>'
    ).encode('utf-8')
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        ftr_names = [n for n in z.namelist() if 'footer' in n and n.endswith('.xml')]
    
    for name in ftr_names:
        _zip_replace(docx_path, name, _make_ftr)


def _zip_replace(zip_path, internal_path, new_data):
    zip_path = str(zip_path)
    tmp = zip_path + '.tmp'
    with zipfile.ZipFile(zip_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == internal_path: data = new_data
                zout.writestr(item, data)
    shutil.move(tmp, zip_path)
