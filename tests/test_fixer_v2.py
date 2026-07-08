"""Unit tests for pipeline/fixer_v2.py — diagnostic verifier."""
import tempfile
import zipfile
from pathlib import Path

import pytest
from docx import Document
from lxml import etree

from pipeline.fixer_v2 import (
    diagnose,
    FixerDiagnostic,
    _VERIFIERS,
    summary,
)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def _make_minimal_docx(path: Path) -> None:
    doc = Document()
    doc.add_paragraph('Hello')
    doc.save(str(path))


def _add_header_file(docx_path: Path, header_name: str = 'header1.xml') -> None:
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
                    rel.set('Type',
                            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header')
                    rel.set('Target', header_name)
                    rel.set('TargetMode', 'Internal')
                    data = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                elif item.filename == '[Content_Types].xml':
                    ns_ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
                    ct_root = etree.fromstring(data)
                    ov = etree.SubElement(ct_root, f'{{{ns_ct}}}Override')
                    ov.set('PartName', f'/word/{header_name}')
                    ov.set('ContentType',
                           'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
                    data = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
            hdr_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:hdr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/>'
                f'<w:sz w:val="20"/></w:rPr>'
                f'<w:t xml:space="preserve">Header</w:t></w:r></w:p></w:hdr>'
            ).encode('utf-8')
            zout.writestr(zipfile.ZipInfo(f'word/{header_name}'), hdr_xml)
    Path(tmp).replace(docx_path)


def _add_footer_file(docx_path: Path, footer_name: str = 'footer1.xml') -> None:
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
                    rel.set('Type',
                            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
                    rel.set('Target', footer_name)
                    rel.set('TargetMode', 'Internal')
                    data = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                elif item.filename == '[Content_Types].xml':
                    ns_ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
                    ct_root = etree.fromstring(data)
                    ov = etree.SubElement(ct_root, f'{{{ns_ct}}}Override')
                    ov.set('PartName', f'/word/{footer_name}')
                    ov.set('ContentType',
                           'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
                    data = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
            ftr_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>'
            ).encode('utf-8')
            zout.writestr(zipfile.ZipInfo(f'word/{footer_name}'), ftr_xml)
    Path(tmp).replace(docx_path)


@pytest.fixture
def temp_docx():
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        path = Path(f.name)
    _make_minimal_docx(path)
    yield path
    if path.exists():
        path.unlink()


@pytest.fixture
def temp_docx_with_header():
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        path = Path(f.name)
    _make_minimal_docx(path)
    _add_header_file(path)
    yield path
    if path.exists():
        path.unlink()


@pytest.fixture
def temp_docx_with_footer():
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        path = Path(f.name)
    _make_minimal_docx(path)
    _add_footer_file(path)
    yield path
    if path.exists():
        path.unlink()


# ── Test: FixerDiagnostic dataclass ─────────────────────

class TestFixerDiagnostic:
    def test_fields(self):
        d = FixerDiagnostic(
            code='HEADER_FONT_MISMATCH',
            action_index=0,
            severity='error',
            detail='test detail',
            expected='A',
            actual='B',
        )
        assert d.code == 'HEADER_FONT_MISMATCH'
        assert d.action_index == 0
        assert d.severity == 'error'
        assert 'HEADER_FONT_MISMATCH' in str(d)
        assert 'test detail' in str(d)

    def test_default_values(self):
        d = FixerDiagnostic(code='TEST', action_index=0, severity='info', detail='x')
        assert d.expected is None
        assert d.actual is None


# ── Test: verifier registry ─────────────────────────────

class TestVerifierRegistry:
    def test_all_actions_have_verifiers(self):
        from pipeline.writer_v2 import _ACTION_HANDLERS
        writer_actions = set(_ACTION_HANDLERS.keys())
        fixer_verifiers = set(_VERIFIERS.keys())
        missing = writer_actions - fixer_verifiers
        assert not missing, f'Missing verifiers for: {missing}'


# ── Test: diagnose with matching fix_plan ───────────────

class TestDiagnoseMatching:

    def test_empty_plan_returns_empty_issues(self, temp_docx):
        issues = diagnose([], temp_docx)
        assert issues == []

    def test_unknown_action_reports_warning(self, temp_docx):
        plan = [{'action': 'nonexistent', 'params': {}}]
        issues = diagnose(plan, temp_docx)
        assert len(issues) == 1
        assert issues[0].code == 'UNKNOWN_ACTION'
        assert issues[0].severity == 'warning'

    def test_correct_sectpr_type_passes(self, temp_docx):
        # First set it
        from pipeline.writer_v2 import apply_fixes
        plan = [{
            'action': 'set_sectpr_type',
            'params': {'section_index': 0, 'val': 'nextPage'},
        }]
        apply_fixes(plan, temp_docx)

        # Then verify
        issues = diagnose(plan, temp_docx)
        assert issues == [], f'Expected no issues, got {issues}'

    def test_mismatched_sectpr_type_detected(self, temp_docx):
        plan = [{
            'action': 'set_sectpr_type',
            'params': {'section_index': 0, 'val': 'neverSetValue'},
        }]
        issues = diagnose(plan, temp_docx)
        assert len(issues) == 1
        assert issues[0].code == 'SECTPR_TYPE_MISMATCH'

    def test_out_of_range_section_index(self, temp_docx):
        plan = [{
            'action': 'set_sectpr_type',
            'params': {'section_index': 99, 'val': 'nextPage'},
        }]
        issues = diagnose(plan, temp_docx)
        assert len(issues) == 1
        assert 'SECTPR_INDEX' in issues[0].code


# ── Test: header font verification ──────────────────────

class TestDiagnoseHeaderFont:

    def test_matching_header_font_passes(self, temp_docx_with_header):
        from pipeline.writer_v2 import apply_fixes
        plan = [{
            'action': 'set_header_font',
            'params': {
                'header_path': 'word/header1.xml',
                'font_ascii': 'Arial',
                'font_eastAsia': '黑体',
                'font_size': '18',
            },
        }]
        apply_fixes(plan, temp_docx_with_header)
        issues = diagnose(plan, temp_docx_with_header)
        assert issues == [], f'Got issues: {issues}'

    def test_mismatched_header_font_detected(self, temp_docx_with_header):
        plan = [{
            'action': 'set_header_font',
            'params': {
                'header_path': 'word/header1.xml',
                'font_ascii': 'WrongFont',
                'font_eastAsia': 'WrongEast',
                'font_size': '99',
            },
        }]
        issues = diagnose(plan, temp_docx_with_header)
        assert len(issues) > 0
        assert any('HEADER_FONT_MISMATCH' in i.code for i in issues)

    def test_missing_header_file_reports_error(self, temp_docx):
        plan = [{
            'action': 'set_header_font',
            'params': {
                'header_path': 'word/nonexistent_header.xml',
                'font_ascii': 'Arial',
                'font_eastAsia': '黑体',
                'font_size': '18',
            },
        }]
        issues = diagnose(plan, temp_docx)
        # Should report error (file missing)
        assert len(issues) > 0


# ── Test: page number verification ──────────────────────

class TestDiagnosePageNumber:

    def test_matching_page_number_passes(self, temp_docx_with_footer):
        from pipeline.writer_v2 import apply_fixes
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
        apply_fixes(plan, temp_docx_with_footer)
        issues = diagnose(plan, temp_docx_with_footer)
        assert issues == [], f'Got issues: {issues}'

    def test_missing_page_number_detected(self, temp_docx_with_footer):
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
        # Don't apply — just diagnose
        issues = diagnose(plan, temp_docx_with_footer)
        assert any('MISSING_PAGE_NUMBER' in i.code for i in issues)

    def test_wrong_prefix_detected(self, temp_docx_with_footer):
        from pipeline.writer_v2 import apply_fixes
        # Apply one format
        plan_apply = [{
            'action': 'add_page_number',
            'params': {
                'footer_path': 'word/footer1.xml',
                'format': '第{PAGE}页',
                'font_eastAsia': '宋体',
                'font_ascii': 'Times New Roman',
                'font_size': '18',
            },
        }]
        apply_fixes(plan_apply, temp_docx_with_footer)

        # Diagnose with different format
        plan_check = [{
            'action': 'add_page_number',
            'params': {
                'footer_path': 'word/footer1.xml',
                'format': 'Page-{PAGE}',
                'font_eastAsia': '宋体',
                'font_ascii': 'Times New Roman',
                'font_size': '18',
            },
        }]
        issues = diagnose(plan_check, temp_docx_with_footer)
        # Will report missing 'Page-' prefix
        # That's correct behavior — the prefix isn't there
        assert any('MISSING_PAGE_NUMBER' in i.code for i in issues)


# ── Test: style verification ────────────────────────────

class TestDiagnoseStyle:

    def test_matching_style_passes(self, temp_docx):
        from pipeline.writer_v2 import apply_fixes
        # Find existing style ID
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
            },
        }]
        apply_fixes(plan, temp_docx)
        issues = diagnose(plan, temp_docx)
        assert issues == [], f'Got issues: {issues}'

    def test_unapplied_style_detected(self, temp_docx):
        plan = [{
            'action': 'set_style',
            'params': {
                'style_id': 'nonexistent_style_xyz',
                'font_ascii': 'Arial',
            },
        }]
        issues = diagnose(plan, temp_docx)
        assert any('STYLE_NOT_APPLIED' in i.code for i in issues)

    def test_mismatched_font_ascii_detected(self, temp_docx):
        from pipeline.writer_v2 import apply_fixes
        with zipfile.ZipFile(temp_docx, 'r') as z:
            root = etree.fromstring(z.read('word/styles.xml'))
            for s in root.iter(f'{{{W}}}style'):
                if s.get(f'{{{W}}}type') == 'paragraph':
                    style_id = s.get(f'{{{W}}}styleId')
                    break
            else:
                pytest.skip('No paragraph style found')

        # Apply Arial
        apply_fixes([{
            'action': 'set_style',
            'params': {'style_id': style_id, 'font_ascii': 'Arial'},
        }], temp_docx)

        # Diagnose expecting Times
        issues = diagnose([{
            'action': 'set_style',
            'params': {'style_id': style_id, 'font_ascii': 'Times New Roman'},
        }], temp_docx)
        assert any('STYLE_NOT_APPLIED' in i.code for i in issues)


# ── Test: remove_extra_sectpr verification ──────────────

class TestDiagnoseExtraSectpr:

    def test_no_extra_sectpr_passes(self, temp_docx):
        plan = [{
            'action': 'remove_extra_sectpr',
            'params': {},
        }]
        issues = diagnose(plan, temp_docx)
        assert issues == [], f'Got issues: {issues}'

    def test_embedded_sectpr_detected(self, temp_docx):
        # Embed a sectPr in a paragraph
        tmp = str(temp_docx) + '.tmp'
        with zipfile.ZipFile(temp_docx, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == 'word/document.xml':
                        root = etree.fromstring(data)
                        body = root.find(f'{{{W}}}body')
                        p = etree.SubElement(body, f'{{{W}}}p')
                        pPr = etree.SubElement(p, f'{{{W}}}pPr')
                        sp = etree.SubElement(pPr, f'{{{W}}}sectPr')
                        etree.SubElement(sp, f'{{{W}}}type', {f'{{{W}}}val': 'oddPage'})
                        data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                    zout.writestr(item, data)
        Path(tmp).replace(temp_docx)

        plan = [{
            'action': 'remove_extra_sectpr',
            'params': {},
        }]
        issues = diagnose(plan, temp_docx)
        # Should detect the extra sectPr
        assert any('EXTRA_ODDPAGE_SECTPR' in i.code for i in issues)


# ── Test: summary function ──────────────────────────────

class TestSummary:
    def test_empty_issues(self):
        s = summary([])
        assert s['total'] == 0
        assert s['passed'] is True
        assert s['by_code'] == {}
        assert s['by_severity'] == {'error': 0, 'warning': 0, 'info': 0}

    def test_with_issues(self):
        issues = [
            FixerDiagnostic(code='A', action_index=0, severity='error', detail='x'),
            FixerDiagnostic(code='A', action_index=1, severity='error', detail='y'),
            FixerDiagnostic(code='B', action_index=2, severity='warning', detail='z'),
        ]
        s = summary(issues)
        assert s['total'] == 3
        assert s['passed'] is False
        assert s['by_code'] == {'A': 2, 'B': 1}
        assert s['by_severity'] == {'error': 2, 'warning': 1, 'info': 0}
