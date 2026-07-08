"""Tests for EnforceConfig dataclass."""
from enforce.config import EnforceConfig
from pathlib import Path


class TestEnforceConfig:
    """Unit tests for EnforceConfig dataclass."""

    def test_construct_with_required_fields(self):
        cfg = EnforceConfig(
            libreoffice_path=Path("/usr/bin/soffice"),
            docx_input=Path("/tmp/input.docx"),
            docx_output=Path("/tmp/output.docx"),
        )
        assert cfg.libreoffice_path == Path("/usr/bin/soffice")
        assert cfg.docx_input == Path("/tmp/input.docx")
        assert cfg.docx_output == Path("/tmp/output.docx")
        assert cfg.odd_page_sections == []

    def test_construct_with_odd_page_sections(self):
        cfg = EnforceConfig(
            libreoffice_path=Path("/usr/bin/soffice"),
            docx_input=Path("/tmp/in.docx"),
            docx_output=Path("/tmp/out.docx"),
            odd_page_sections=["封面", "绪论", "参考文献"],
        )
        assert cfg.odd_page_sections == ["封面", "绪论", "参考文献"]

    def test_default_odd_page_sections_is_empty_list(self):
        cfg = EnforceConfig(
            libreoffice_path=Path("/soffice"),
            docx_input=Path("in.docx"),
            docx_output=Path("out.docx"),
        )
        assert isinstance(cfg.odd_page_sections, list)
        assert len(cfg.odd_page_sections) == 0

    def test_all_fields_are_path_objects(self):
        cfg = EnforceConfig(
            libreoffice_path=Path("/opt/libreoffice/soffice.exe"),
            docx_input=Path("C:/Users/T/doc.docx"),
            docx_output=Path("C:/Users/T/out.docx"),
        )
        assert isinstance(cfg.libreoffice_path, Path)
        assert isinstance(cfg.docx_input, Path)
        assert isinstance(cfg.docx_output, Path)

    def test_odd_page_sections_mutable_list(self):
        cfg = EnforceConfig(
            libreoffice_path=Path("/b"),
            docx_input=Path("/a"),
            docx_output=Path("/c"),
            odd_page_sections=["绪论"],
        )
        cfg.odd_page_sections.append("结论")
        assert cfg.odd_page_sections == ["绪论", "结论"]

    def test_windows_paths(self):
        """Test with Windows-style paths."""
        cfg = EnforceConfig(
            libreoffice_path=Path("C:/Program Files/LibreOffice/program/soffice.exe"),
            docx_input=Path("D:/papers/thesis.docx"),
            docx_output=Path("D:/papers/thesis_fixed.docx"),
            odd_page_sections=["Abstract", "Introduction", "Conclusion"],
        )
        assert "Program Files" in str(cfg.libreoffice_path)
        assert len(cfg.odd_page_sections) == 3
