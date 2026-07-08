"""
Two-pass odd-page enforcer — main orchestration.
Stable: insert pb_para+ep before H1 heading.
"""
import zipfile, subprocess, re, shutil
from pathlib import Path
from lxml import etree
from .config import EnforceConfig
from .pdf_reader import get_section_start_pages

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML_NS = 'http://www.w3.org/XML/1998/namespace'

def enforce_odd_pages(config: EnforceConfig) -> Path:
    sections = config.odd_page_sections
    pdf_path = config.docx_input.with_suffix('.pdf')

    print("=== PASS 1: Convert to PDF ===")
    try:
        from generator.pdf_converter import convert_to_pdf
        convert_to_pdf(config.docx_input, pdf_path, timeout=60)
    except RuntimeError:
        subprocess.run([str(config.libreoffice_path),"--headless","--convert-to","pdf","--outdir",str(config.docx_input.parent),str(config.docx_input)],capture_output=True,timeout=60,check=True)

    print("=== PASS 1: Read section start pages ===")
    start_pages = get_section_start_pages(pdf_path, sections)

    print(f"\nPDF: {len(sections)} sections, start pages:")
    shift = 0
    needs_blank = []
    for i, heading in enumerate(sections):
        sp = start_pages[i]
        eff = (sp + shift) if sp else None
        need = eff and eff % 2 == 0
        if need: needs_blank.append((i, heading)); shift += 1
        mark = " ← BLANK" if need else ""
        sp_val = str(sp) if sp is not None else "?"
        eff_val = str(eff) if eff is not None else "?"
        print(f"  [{i:2d}] {heading:<25s} raw_p={sp_val:>3s} eff_p={eff_val:>3s} {'ODD' if eff and eff%2==1 else ('EVEN' if eff else '?')}{mark}")
    print(f"\n  {len(needs_blank)} sections need blank pages (shift={shift})")

    if not needs_blank:
        shutil.copy(config.docx_input, config.docx_output)
        return config.docx_output

    print("\n=== PASS 2: Insert blank pages ===")
    shutil.copy(config.docx_input, config.docx_output)

    with zipfile.ZipFile(config.docx_output, 'r') as z:
        doc = etree.fromstring(z.read('word/document.xml'))
        body = doc.find(f'{{{W}}}body')

        sectpr_paras = []
        for pi, para in enumerate(body):
            if para.tag == f'{{{W}}}p':
                pp = para.find(f'{{{W}}}pPr')
                if pp is not None and pp.find(f'{{{W}}}sectPr') is not None:
                    sectpr_paras.append((pi, para))

        heading_positions = {}
        h1_pat = re.compile(r'^(?:封面|郑\s*重\s*声\s*明|郑重声明|摘\s*要|ABSTRACT|目\s*录|目录|\d+[\u4e00-\u9fff]|结\s*论|结论|参考\s*文献|参考文献|致\s*谢|致谢|附\s*录|附录)')
        def _find_headings(container, sdt_parent=None):
            for child in list(container):
                if child.tag == f'{{{W}}}p':
                    txt = ''.join(t.text or '' for t in child.iter(f'{{{W}}}t')).strip()
                    if txt and h1_pat.match(re.sub(r'\s+', '', txt)):
                        key = re.sub(r'\s+', '', txt)
                        if key not in heading_positions or sdt_parent is None:
                            heading_positions[key] = sdt_parent if sdt_parent is not None else child
                elif child.tag == f'{{{W}}}sdt':
                    sc = child.find(f'{{{W}}}sdtContent')
                    if sc is not None: _find_headings(sc, sdt_parent=child)
        _find_headings(body)

        def _pb(): 
            pb = etree.Element(f'{{{W}}}p'); r = etree.SubElement(pb, f'{{{W}}}r')
            etree.SubElement(r, f'{{{W}}}br').set(f'{{{W}}}type', 'page')
            r2 = etree.SubElement(pb, f'{{{W}}}r'); t = etree.SubElement(r2, f'{{{W}}}t')
            t.set(f'{{{XML_NS}}}space', 'preserve'); t.text = ' '
            return pb
        def _ep():
            ep = etree.Element(f'{{{W}}}p'); pp = etree.SubElement(ep, f'{{{W}}}pPr')
            sp = etree.SubElement(pp, f'{{{W}}}spacing')
            sp.set(f'{{{W}}}line','400'); sp.set(f'{{{W}}}lineRule','exact'); sp.set(f'{{{W}}}snapToGrid','0')
            r = etree.SubElement(ep, f'{{{W}}}r'); rp = etree.SubElement(r, f'{{{W}}}rPr')
            etree.SubElement(rp, f'{{{W}}}sz').set(f'{{{W}}}val','28')  # 14pt
            t = etree.SubElement(r, f'{{{W}}}t'); t.set(f'{{{XML_NS}}}space','preserve'); t.text=' '
            # 再加几个 run 确保内容足够
            r2 = etree.SubElement(ep, f'{{{W}}}r')
            t2 = etree.SubElement(r2, f'{{{W}}}t'); t2.set(f'{{{XML_NS}}}space','preserve'); t2.text=' '
            r3 = etree.SubElement(ep, f'{{{W}}}r')
            t3 = etree.SubElement(r3, f'{{{W}}}t'); t3.set(f'{{{XML_NS}}}space','preserve'); t3.text=' '
            return ep

        inserted = set()
        for sec_idx, heading in reversed(needs_blank):
            heading_norm = re.sub(r'\s+', '', heading)
            if heading_norm in inserted: continue
            inserted.add(heading_norm)

            # ── 找到目标标题 ──
            heading_para = heading_positions.get(heading_norm)
            if heading_para is None:
                continue

            # ── 找标题之前最近的 sectPr 段落（前一节末尾） ──
            heading_idx = list(body).index(heading_para)
            target = None
            for bi in range(heading_idx - 1, -1, -1):
                child = body[bi]
                if child.tag == f'{{{W}}}p':
                    pp = child.find(f'{{{W}}}pPr')
                    if pp is not None and pp.find(f'{{{W}}}sectPr') is not None:
                        target = child
                        break
            if target is None:
                target = heading_para  # fallback: 没找到，插标题前

            # ── 如果 sectPr 段有内容，先提取到独立空段 ──
            pp = target.find(f'{{{W}}}pPr')
            sp = pp.find(f'{{{W}}}sectPr') if pp is not None else None
            if sp is not None:
                para_txt = ''.join(t.text or '' for t in target.iter(f'{{{W}}}t')).strip()
                if para_txt:
                    pp.remove(sp)
                    new_para = etree.Element(f'{{{W}}}p')
                    new_pPr = etree.SubElement(new_para, f'{{{W}}}pPr')
                    new_pPr.append(sp)
                    target.addnext(new_para)
                    target = new_para  # 用干净段做插入目标

            # 分页符+空段插在 sectPr 段前 → 空白页属于前一节
            # 检查前面是否已有分页符，避免双空白
            ti = list(body).index(target)
            prev = body[ti - 1] if ti > 0 else None
            has_existing_br = prev is not None and prev.find(f'.//{{{W}}}br') is not None
            if not has_existing_br:
                target.addprevious(_pb())
            target.addprevious(_ep())
            target.addprevious(_ep())
            target.addprevious(_ep())
            print(f"  blank before \"{heading}\"" + (" (reuse existing br)" if has_existing_br else ""))

        new_xml = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
        tmp = str(config.docx_output.parent / "_tmp.docx")
        with zipfile.ZipFile(config.docx_output, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    zout.writestr(item, new_xml if item.filename == 'word/document.xml' else zin.read(item.filename))
        shutil.move(tmp, str(config.docx_output))

    print(f"\n  Final: {config.docx_output}")
    return config.docx_output
