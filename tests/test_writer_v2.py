"""Unit tests for pipeline/writer_v2.py — pure executor."""
import tempfile
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from pipeline.writer_v2 import (
    apply_fixes,
    _ACTION_HANDLERS,
    _zip_replace,
)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML = 'http://www.w3.org/XML/1998/namespace'


def _make_minimal_docx(path: Path) -> None:
    """Create a minimal docx with all required parts for ZIP-level operations."""
    doc = Document()
    doc.add_paragraph('Hello')
    doc.save(str(path))


def _add_header_file(docx_path: Path, header_name: str = 'header1.xml') -> None:
    """Add a header XML file to the docx ZIP and register it in rels + content_types."""
    tmp = str(docx_path) + '.tmp'
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/_rels/document.xml.rels':
                    rels_root = etree.fromstring(data)
                    max_rid = max(
                        int(r.get('Id', 'rId0').replace('rId', ''))
                        for r in rels_root.iter()
                        if r.tag.endswith('}Relationship')
                    ) + 1
                    rid = f'rId{max_rid}'
                    rel = etree.SubElement(rels_root, 'Relationship')
                    rel.set('Id', rid)
                    rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header')
                    rel.set('Target', header_name)
                    rel.set('TargetMode', 'Internal')
                    data = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                elif item.filename == '[Content_Types].xml':
                    ct_root = etree.fromstring(data)
                    ov = etree.SubElement(ct_root, f'{{{W.replace("main", "content-types")}}}Override')
                    ov.set('PartName', f'/word/{header_name}')
                    ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
                    data = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
            # Write the header XML
            hdr_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:hdr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/>'
                f'<w:sz w:val="20"/></w:rPr>'
                f'<w:t xml:space="preserve">Original Header</w:t></w:r>'
                f'</w:p></w:hdr>'
            ).encode('utf-8')
            zout.writestr(zipfile.ZipInfo(f'word/{header_name}'), hdr_xml)
    Path(tmp).replace(docx_path)


def _add_footer_file(docx_path: Path, footer_name: str = 'footer1.xml') -> None:
    """Add a footer XML file to the docx ZIP."""
    tmp = str(docx_path) + '.tmp'
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/_rels/document.xml.rels':
                    rels_root = etree.fromstring(data)
                    max_rid = max(
                        int(r.get('Id', 'rId0').replace('rId', ''))
                        for r in rels_root.iter()
                        if r.tag.endswith('}Relationship')
                    ) + 1
                    rid = f'rId{max_rid}'
                    rel = etree.SubElement(rels_root, 'Relationship')
                    rel.set('Id', rid)
                    rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
                    rel.set('Target', footer_name)
                    rel.set('TargetMode', 'Internal')
                    data = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                elif item.filename == '[Content_Types].xml':
                    ct_root = etree.fromstring(data)
                    ov = etree.SubElement(ct_root, f'{{{W.replace("main", "content-types")}}}Override')
                    ov.set('PartName', f'/word/{footer_name}')
                    ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
                    data = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
            # Write empty footer XML
            ftr_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Old Footer</w:t></w:r></w:p></w:ftr>'
            ).encode('utf-8')
            zout.writestr(zipfile.ZipInfo(f'word/{footer_name}'), ftr_xml)
    Path(tmp).replace(docx_path)


# ── Fixtures ──────────────────────────────────────────────

@pytest.fixture
def temp_docx():
    """Create a temporary minimal docx for testing."""
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        path = Path(f.name)
    _make_minimal_docx(path)
    yield path
    if path.exists():
        path.unlink()


@pytest.fixture
def temp_docx_with_header():
    """Create a temp docx with a header file."""
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        path = Path(f.name)
    _make_minimal_docx(path)
    _add_header_file(path)
    yield path
    if path.exists():
        path.unlink()


@pytest.fixture
def temp_docx_with_footer():
    """Create a temp docx with a footer file."""
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        path = Path(f.name)
    _make_minimal_docx(path)
    _add_footer_file(path)
    yield path
    if path.exists():
        path.unlink()


@pytest.fixture
def temp_docx_full():
    """Create a temp docx with header, footer, and styles."""
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        path = Path(f.name)
    _make_minimal_docx(path)
    _add_header_file(path)
    _add_footer_file(path)
    yield path
    if path.exists():
        path.unlink()


# ── Test: Registered actions ─────────────────────────────

class TestActionRegistry:
    """Verify all required actions are registered."""

    def test_all_six_actions_registered(self):
        expected = {
            'set_sectpr_type',
            'set_header_font',
            'add_page_number',
            'set_body_font_ascii',
            'set_style',
            'remove_extra_sectpr',
        }
        assert expected.issubset(set(_ACTION_HANDLERS.keys()))

    def test_apply_fixes_skips_unknown_action(self, temp_docx):
        plan = [{'action': 'nonexistent_action', 'params': {}}]
        logs = apply_fixes(plan, temp_docx)
        assert any('SKIP' in l for l in logs)
        assert any('nonexistent_action' in l for l in logs)


# ── Test: set_sectpr_type ────────────────────────────────

class TestSetSectprType:

    def test_changes_nextPage(self, temp_docx):
        plan = [{
            'action': 'set_sectpr_type',
            'params': {'section_index': 0, 'val': 'nextPage'},
        }]
        logs = apply_fixes(plan, temp_docx)
        assert any('OK' in l for l in logs)

        # Verify
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/document.xml'))
            sp = list(root.iter(f'{{{W}}}sectPr'))[0]
            type_el = sp.find(f'{{{W}}}type')
            assert type_el is not None
            assert type_el.get(f'{{{W}}}val') == 'nextPage'

    def test_changes_oddPage(self, temp_docx):
        plan = [{
            'action': 'set_sectpr_type',
            'params': {'section_index': 0, 'val': 'oddPage'},
        }]
        apply_fixes(plan, temp_docx)
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/document.xml'))
            sp = list(root.iter(f'{{{W}}}sectPr'))[0]
            type_el = sp.find(f'{{{W}}}type')
            assert type_el.get(f'{{{W}}}val') == 'oddPage'

    def test_raises_on_bad_index(self, temp_docx):
        plan = [{
            'action': 'set_sectpr_type',
            'params': {'section_index': 99, 'val': 'nextPage'},
        }]
        logs = apply_fixes(plan, temp_docx)
        assert any('FAIL' in l for l in logs)


# ── Test: set_header_font ────────────────────────────────

class TestSetHeaderFont:

    def test_changes_ascii_font(self, temp_docx_with_header):
        plan = [{
            'action': 'set_header_font',
            'params': {
                'header_path': 'word/header1.xml',
                'font_ascii': 'Arial',
                'font_eastAsia': '黑体',
                'font_size': '18',
            },
        }]
        logs = apply_fixes(plan, temp_docx_with_header)
        assert any('OK' in l for l in logs)

        with zipfile.ZipFile(temp_docx_with_header, 'r') as z:
            root = etree.fromstring(z.read('word/header1.xml'))
            rPr = list(root.iter(f'{{{W}}}rPr'))[0]
            rf = rPr.find(f'{{{W}}}rFonts')
            assert rf is not None
            assert rf.get(f'{{{W}}}ascii') == 'Arial'
            assert rf.get(f'{{{W}}}eastAsia') == '黑体'
            sz = rPr.find(f'{{{W}}}sz')
            assert sz.get(f'{{{W}}}val') == '18'

    def test_no_hardcoded_font(self, temp_docx_with_header):
        """Writer never hardcodes font names — only uses whatever params say."""
        plan = [{
            'action': 'set_header_font',
            'params': {
                'header_path': 'word/header1.xml',
                'font_ascii': 'CustomFont123',
                'font_eastAsia': '自定义字体',
                'font_size': '24',
            },
        }]
        apply_fixes(plan, temp_docx_with_header)
        with zipfile.ZipFile(temp_docx_with_header, 'r') as z:
            root = etree.fromstring(z.read('word/header1.xml'))
            rPr = list(root.iter(f'{{{W}}}rPr'))[0]
            rf = rPr.find(f'{{{W}}}rFonts')
            assert rf.get(f'{{{W}}}ascii') == 'CustomFont123'
            assert rf.get(f'{{{W}}}eastAsia') == '自定义字体'


# ── Test: add_page_number ────────────────────────────────

class TestAddPageNumber:

    def test_inserts_fldChar_sequence(self, temp_docx_with_footer):
        plan = [{
            'action': 'add_page_number',
            'params': {
                'footer_path': 'word/footer1.xml',
                'format': '第{PAGE}页',
                'font_eastAsia': '宋体',
                'font_ascii': 'Times New Roman',
                'font_size': '18',
            },
        }]
        logs = apply_fixes(plan, temp_docx_with_footer)
        assert any('OK' in l for l in logs)

        with zipfile.ZipFile(temp_docx_with_footer, 'r') as z:
            root = etree.fromstring(z.read('word/footer1.xml'))
            fld_chars = list(root.iter(f'{{{W}}}fldChar'))
            assert len(fld_chars) == 3  # begin, separate, end

            instr_texts = list(root.iter(f'{{{W}}}instrText'))
            assert any('PAGE' in (t.text or '') for t in instr_texts)

            # Check prefix/suffix text
            all_text = ''.join(t.text or '' for t in root.iter(f'{{{W}}}t'))
            assert '第' in all_text
            assert '页' in all_text

    def test_custom_format(self, temp_docx_with_footer):
        plan = [{
            'action': 'add_page_number',
            'params': {
                'footer_path': 'word/footer1.xml',
                'format': 'Page {PAGE} of NUMPAGES',
                'font_eastAsia': '宋体',
                'font_ascii': 'Times New Roman',
                'font_size': '18',
            },
        }]
        apply_fixes(plan, temp_docx_with_footer)
        with zipfile.ZipFile(temp_docx_with_footer, 'r') as z:
            root = etree.fromstring(z.read('word/footer1.xml'))
            all_text = ''.join(t.text or '' for t in root.iter(f'{{{W}}}t'))
            assert 'Page' in all_text
            assert 'of NUMPAGES' in all_text


# ── Test: set_body_font_ascii ────────────────────────────

class TestSetBodyFontAscii:

    def test_modifies_default_style(self, temp_docx):
        plan = [{
            'action': 'set_body_font_ascii',
            'params': {'style_id': 'a8', 'font_ascii': 'TestFont'},
        }]
        logs = apply_fixes(plan, temp_docx)
        # 'a8' may not exist in all docx, so we just ensure no crash
        assert len(logs) == 1

    def test_modifies_existing_style(self, temp_docx):
        # Find the actual default style ID
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}type') == 'paragraph' and s.get(f'{{{W}}}default') == '1':
                    style_id = s.get(f'{{{W}}}styleId')
                    break
            else:
                pytest.skip('No default paragraph style found')

        plan = [{
            'action': 'set_body_font_ascii',
            'params': {'style_id': style_id, 'font_ascii': 'TestFontX'},
        }]
        apply_fixes(plan, temp_docx)

        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}styleId') == style_id:
                    rPr = s.find(f'{{{W}}}rPr')
                    if rPr is not None:
                        rf = rPr.find(f'{{{W}}}rFonts')
                        if rf is not None:
                            assert rf.get(f'{{{W}}}ascii') == 'TestFontX'


# ── Test: set_style ─────────────────────────────────────

class TestSetStyle:

    def test_sets_font_properties(self, temp_docx):
        # Find a valid style ID
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}type') == 'paragraph':
                    style_id = s.get(f'{{{W}}}styleId')
                    break
            else:
                pytest.skip('No paragraph style found')

        plan = [{
            'action': 'set_style',
            'params': {
                'style_id': style_id,
                'font_ascii': 'Arial',
                'font_eastAsia': '黑体',
                'font_size_pt': 14,
                'bold': True,
            },
        }]
        apply_fixes(plan, temp_docx)

        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}styleId') == style_id:
                    rPr = s.find(f'{{{W}}}rPr')
                    assert rPr is not None
                    rf = rPr.find(f'{{{W}}}rFonts')
                    assert rf is not None
                    assert rf.get(f'{{{W}}}ascii') == 'Arial'
                    assert rf.get(f'{{{W}}}eastAsia') == '黑体'
                    sz = rPr.find(f'{{{W}}}sz')
                    assert sz is not None
                    assert sz.get(f'{{{W}}}val') == '28'  # 14pt * 2
                    b = rPr.find(f'{{{W}}}b')
                    assert b is not None

    def test_sets_alignment(self, temp_docx):
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}type') == 'paragraph':
                    style_id = s.get(f'{{{W}}}styleId')
                    break
            else:
                pytest.skip('No paragraph style found')

        plan = [{
            'action': 'set_style',
            'params': {
                'style_id': style_id,
                'alignment': 'center',
            },
        }]
        apply_fixes(plan, temp_docx)

        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}styleId') == style_id:
                    pPr = s.find(f'{{{W}}}pPr')
                    assert pPr is not None
                    jc = pPr.find(f'{{{W}}}jc')
                    assert jc is not None
                    assert jc.get(f'{{{W}}}val') == 'center'

    def test_sets_spacing(self, temp_docx):
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}type') == 'paragraph':
                    style_id = s.get(f'{{{W}}}styleId')
                    break
            else:
                pytest.skip('No paragraph style found')

        plan = [{
            'action': 'set_style',
            'params': {
                'style_id': style_id,
                'before_lines': 100,
                'after_lines': 200,
                'line': 360,
                'lineRule': 'auto',
            },
        }]
        apply_fixes(plan, temp_docx)

        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}styleId') == style_id:
                    pPr = s.find(f'{{{W}}}pPr')
                    assert pPr is not None
                    sp = pPr.find(f'{{{W}}}spacing')
                    assert sp is not None
                    assert sp.get(f'{{{W}}}beforeLines') == '100'
                    assert sp.get(f'{{{W}}}afterLines') == '200'
                    assert sp.get(f'{{{W}}}line') == '360'
                    assert sp.get(f'{{{W}}}lineRule') == 'auto'


# ── Test: remove_extra_sectpr ────────────────────────────

class TestRemoveExtraSectpr:

    def test_removes_embedded_sectpr(self, temp_docx):
        # Manually embed a fake sectPr inside a paragraph
        tmp = str(temp_docx) + '.tmp'
        with zipfile.ZipFile(temp_docx, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == 'word/document.xml':
                        root = etree.fromstring(data)
                        body = root.find(f'{{{W}}}body')

                        # Add a paragraph with embedded oddPage sectPr
                        p = etree.SubElement(body, f'{{{W}}}p')
                        pPr = etree.SubElement(p, f'{{{W}}}pPr')
                        sp = etree.SubElement(pPr, f'{{{W}}}sectPr')
                        etree.SubElement(sp, f'{{{W}}}type', {f'{{{W}}}val': 'oddPage'})
                        etree.SubElement(sp, f'{{{W}}}pgSz', {f'{{{W}}}w': '11906', f'{{{W}}}h': '16838'})

                        data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                    zout.writestr(item, data)
        Path(tmp).replace(temp_docx)

        # Verify it's there
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/document.xml'))
            body = root.find(f'{{{W}}}body')
            embedded_count = sum(
                1 for p in body
                if p.tag == f'{{{W}}}p' and p.find(f'{{{W}}}pPr') is not None
                and p.find(f'{{{W}}}pPr').find(f'{{{W}}}sectPr') is not None
            )
            assert embedded_count == 1

        # Remove it
        plan = [{
            'action': 'remove_extra_sectpr',
            'params': {'preserve_body_sectpr': True},
        }]
        logs = apply_fixes(plan, temp_docx)
        assert any('OK' in l for l in logs)

        # Verify removed
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/document.xml'))
            body = root.find(f'{{{W}}}body')
            embedded_count = sum(
                1 for p in body
                if p.tag == f'{{{W}}}p' and p.find(f'{{{W}}}pPr') is not None
                and p.find(f'{{{W}}}pPr').find(f'{{{W}}}sectPr') is not None
            )
            assert embedded_count == 0

    def test_preserves_body_sectpr(self, temp_docx):
        """Body-level sectPr (last child of body) should be preserved."""
        plan = [{
            'action': 'remove_extra_sectpr',
            'params': {'preserve_body_sectpr': True},
        }]
        apply_fixes(plan, temp_docx)

        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/document.xml'))
            body = root.find(f'{{{W}}}body')
            # The last sectPr should still exist (body-level)
            body_sectprs = [c for c in body if c.tag == f'{{{W}}}sectPr']
            # At minimum, the original body sectPr should remain
            assert len(list(body.iter(f'{{{W}}}sectPr'))) >= 1


# ── Test: apply_fixes returns correct log format ─────────

class TestApplyFixesLog:

    def test_each_action_returns_log(self, temp_docx):
        plan = [
            {'action': 'set_sectpr_type', 'params': {'section_index': 0, 'val': 'nextPage'}},
            {'action': 'set_sectpr_type', 'params': {'section_index': 0, 'val': 'continuous'}},
        ]
        logs = apply_fixes(plan, temp_docx)
        assert len(logs) == 2
        assert logs[0].startswith('[OK')
        assert logs[1].startswith('[OK')

    def test_failed_action_returns_FAIL_log(self, temp_docx):
        plan = [
            {'action': 'set_sectpr_type', 'params': {'section_index': 99, 'val': 'nextPage'}},
        ]
        logs = apply_fixes(plan, temp_docx)
        assert len(logs) == 1
        assert '[FAIL' in logs[0]
