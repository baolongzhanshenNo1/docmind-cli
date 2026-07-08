"""Tests for toc_real_pages.py — mock fitz testing."""
import subprocess
from pathlib import Path
from unittest.mock import patch, MagicMock, call

import pytest

from generator.toc_real_pages import (
    get_real_page_numbers,
    _match_titles_to_pages,
    _convert_to_pdf,
)


# ── helpers ──

def _make_text_block(x0=0, y0=80, x1=500, y1=100, text="", block_type=0, block_no=0):
    """Return a tuple simulating a fitz text block."""
    return (x0, y0, x1, y1, text, block_type, block_no)


def _mock_pdf_doc(pages_blocks, page_count=None):
    """Build a mock fitz Document with per-page text blocks."""
    if page_count is None:
        page_count = max(pages_blocks.keys()) + 1 if pages_blocks else 0

    page_mocks = {}
    for pn in range(page_count):
        pm = MagicMock()
        pm.get_text.return_value = pages_blocks.get(pn, [])
        page_mocks[pn] = pm

    mock_doc = MagicMock()
    mock_doc.page_count = page_count

    def _getitem(pn):
        if pn < page_count:
            return page_mocks[pn]
        raise IndexError(pn)

    mock_doc.__getitem__.side_effect = _getitem
    return mock_doc


def _make_section_dict(type="body", title="Ch1"):
    return {"type": type, "title": title, "content": ""}


class TestMatchTitlesToPages:
    """Unit tests for _match_titles_to_pages() with mocked fitz."""

    def test_imports_without_fitz(self):
        """Module should still be importable without fitz installed."""
        from generator.toc_real_pages import get_real_page_numbers
        assert callable(get_real_page_numbers)

    @patch("generator.toc_real_pages.fitz")
    def test_single_title_found(self, mock_fitz):
        mock_doc = _mock_pdf_doc({
            1: [_make_text_block(y0=80, text="1 绪论")],
        }, page_count=3)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), ["1 绪论"])
        assert result == {"0": 2}  # 1-based page 2

    @patch("generator.toc_real_pages.fitz")
    def test_title_not_found_returns_minus_one(self, mock_fitz):
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="摘要")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), ["绪论"])
        assert result == {"0": -1}

    @patch("generator.toc_real_pages.fitz")
    def test_multiple_sections_in_order(self, mock_fitz):
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="摘要")],
            1: [_make_text_block(y0=90, text="1 绪论")],
            4: [_make_text_block(y0=75, text="参考文献")],
            5: [_make_text_block(y0=85, text="致谢")],
        }, page_count=6)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(
            Path("fake.pdf"),
            ["摘要", "1 绪论", "参考文献", "致谢"],
        )
        assert result == {"0": 1, "1": 2, "2": 5, "3": 6}

    @patch("generator.toc_real_pages.fitz")
    def test_cover_always_page_1(self, mock_fitz):
        """封面固定为第 1 页，即使 PDF 中未显式标出或在不同页。"""
        mock_doc = _mock_pdf_doc({
            0: [],  # no text on page 0
            1: [_make_text_block(y0=150, text="封面")],
        }, page_count=2)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(
            Path("fake.pdf"),
            ["封面", "摘要"],
        )
        # 封面 index=0 固定为 1，不扫描
        assert result["0"] == 1

    @patch("generator.toc_real_pages.fitz")
    def test_cover_not_first_section_not_forced(self, mock_fitz):
        """只有 index=0 的封面强制为第 1 页。"""
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="摘要")],
            1: [_make_text_block(y0=150, text="封面")],
        }, page_count=2)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(
            Path("fake.pdf"),
            ["摘要", "封面"],  # 封面在索引 1
        )
        # 封面 y₀=150 > 100，应该匹配到第 2 页
        assert result["1"] == 2

    @patch("generator.toc_real_pages.fitz")
    def test_heading_at_wrong_y_ignored(self, mock_fitz):
        """标题在错误 Y 位置应忽略。"""
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=200, text="绪论")],  # too low
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), ["绪论"])
        assert result == {"0": -1}

    @patch("generator.toc_real_pages.fitz")
    def test_y0_boundary_70_excluded(self, mock_fitz):
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=70, text="绪论")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), ["绪论"])
        assert result == {"0": -1}

    @patch("generator.toc_real_pages.fitz")
    def test_y0_boundary_100_excluded(self, mock_fitz):
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=100, text="绪论")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), ["绪论"])
        assert result == {"0": -1}

    @patch("generator.toc_real_pages.fitz")
    def test_first_occurrence_is_returned(self, mock_fitz):
        """标题出现在多页时，返回第一页。"""
        mock_doc = _mock_pdf_doc({
            1: [_make_text_block(y0=80, text="实验")],
            2: [_make_text_block(y0=80, text="实验 - 续")],
        }, page_count=4)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), ["实验"])
        assert result == {"0": 2}

    @patch("generator.toc_real_pages.fitz")
    def test_empty_titles_returns_empty_dict(self, mock_fitz):
        mock_doc = _mock_pdf_doc({}, page_count=5)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), [])
        assert result == {}

    @patch("generator.toc_real_pages.fitz")
    def test_doc_close_is_called(self, mock_fitz):
        mock_doc = _mock_pdf_doc({0: []}, page_count=1)
        mock_fitz.open.return_value = mock_doc

        _match_titles_to_pages(Path("fake.pdf"), ["绪论"])
        mock_doc.close.assert_called_once()

    @patch("generator.toc_real_pages.fitz")
    def test_mixed_found_and_not_found(self, mock_fitz):
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=90, text="摘要")],
            1: [_make_text_block(y0=80, text="1 绪论")],
            2: [],  # "结论" not here
            3: [_make_text_block(y0=85, text="参考文献")],
        }, page_count=4)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(
            Path("fake.pdf"),
            ["摘要", "1 绪论", "结论", "参考文献"],
        )
        assert result == {"0": 1, "1": 2, "2": -1, "3": 4}

    @patch("generator.toc_real_pages.fitz")
    def test_substring_match(self, mock_fitz):
        """标题为子串时也应匹配（如 "绪论" 匹配 "1 绪论"）。"""
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="1 绪论")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = _match_titles_to_pages(Path("fake.pdf"), ["绪论"])
        assert result == {"0": 1}


class TestConvertToPdf:
    """Tests for _convert_to_pdf() — mocked subprocess."""

    @patch("generator.toc_real_pages.subprocess.run")
    def test_successful_conversion(self, mock_run):
        mock_run.return_value = MagicMock(returncode=0, stdout=b"", stderr=b"")

        # We need the PDF to "exist" after conversion
        with patch("generator.toc_real_pages.Path.exists", return_value=True):
            result = _convert_to_pdf(
                Path("/tmp/test.docx"),
                Path("/usr/bin/soffice"),
            )
            assert result == Path("/tmp/test.pdf")

    @patch("generator.toc_real_pages.subprocess.run")
    def test_conversion_failure_raises(self, mock_run):
        mock_run.side_effect = subprocess.CalledProcessError(
            1, "cmd", stderr=b"Error: something went wrong"
        )
        with pytest.raises(RuntimeError, match="LibreOffice 转换 PDF 失败"):
            _convert_to_pdf(Path("/tmp/test.docx"), Path("/usr/bin/soffice"))

    @patch("generator.toc_real_pages.subprocess.run")
    def test_timeout_raises(self, mock_run):
        mock_run.side_effect = subprocess.TimeoutExpired("cmd", 120)
        with pytest.raises(RuntimeError, match="超时"):
            _convert_to_pdf(Path("/tmp/test.docx"), Path("/usr/bin/soffice"))


class TestGetRealPageNumbersIntegration:
    """Integration-style tests for get_real_page_numbers() with mocked components."""

    @patch("generator.toc_real_pages.fitz")
    @patch("generator.toc_real_pages._convert_to_pdf")
    def test_full_flow_with_sections(self, mock_convert, mock_fitz):
        """Full flow: convert → match titles → return page map."""
        # Mock the conversion to return a fake PDF path
        fake_pdf = Path("/tmp/test.pdf")
        mock_convert.return_value = fake_pdf

        # Mock fitz to return a document with sections on pages 1, 3, 5
        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="摘要")],
            2: [_make_text_block(y0=90, text="1 绪论")],
            4: [_make_text_block(y0=75, text="参考文献")],
        }, page_count=5)
        mock_fitz.open.return_value = mock_doc

        # Also need to mock Path.unlink to prevent cleanup error
        with patch.object(Path, 'unlink', return_value=None):
            sections = [
                {"type": "abstract_cn", "title": "摘要", "content": "..."},
                {"type": "body", "title": "1 绪论", "content": "..."},
                {"type": "references", "title": "参考文献", "content": "..."},
            ]
            result = get_real_page_numbers(
                Path("/tmp/test.docx"),
                sections,
                Path("/usr/bin/soffice"),
            )

        assert result == {"0": 1, "1": 3, "2": 5}
        mock_convert.assert_called_once()

    @patch("generator.toc_real_pages.fitz")
    @patch("generator.toc_real_pages._convert_to_pdf")
    def test_cover_handled_as_page_1(self, mock_convert, mock_fitz):
        """封面固定为第 1 页。"""
        fake_pdf = Path("/tmp/test.pdf")
        mock_convert.return_value = fake_pdf

        mock_doc = _mock_pdf_doc({
            0: [],  # cover may have no text
            2: [_make_text_block(y0=80, text="摘要")],
        }, page_count=3)
        mock_fitz.open.return_value = mock_doc

        with patch.object(Path, 'unlink', return_value=None):
            sections = [
                {"type": "cover", "title": "封面", "content": ""},
                {"type": "abstract_cn", "title": "摘要", "content": "..."},
            ]
            result = get_real_page_numbers(
                Path("/tmp/test.docx"),
                sections,
                Path("/usr/bin/soffice"),
            )

        assert result == {"0": 1, "1": 3}

    @patch("generator.toc_real_pages.fitz")
    @patch("generator.toc_real_pages._convert_to_pdf")
    def test_section_not_found_returns_minus_one(self, mock_convert, mock_fitz):
        """未找到的 section 返回 -1（调用方应 fallback）。"""
        fake_pdf = Path("/tmp/test.pdf")
        mock_convert.return_value = fake_pdf

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="摘要")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        with patch.object(Path, 'unlink', return_value=None):
            sections = [
                {"type": "abstract_cn", "title": "摘要", "content": "..."},
                {"type": "body", "title": "绪论", "content": "..."},  # not found
            ]
            result = get_real_page_numbers(
                Path("/tmp/test.docx"),
                sections,
                Path("/usr/bin/soffice"),
            )

        assert result == {"0": 1, "1": -1}
