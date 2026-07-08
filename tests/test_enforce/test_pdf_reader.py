"""Tests for pdf_reader.get_section_start_pages() — mock fitz.open."""
from pathlib import Path
from unittest.mock import patch, MagicMock


def _make_text_block(x0=0, y0=80, x1=500, y1=100, text="", block_type=0, block_no=0):
    """Return a tuple simulating a fitz text block."""
    return (x0, y0, x1, y1, text, block_type, block_no)


def _mock_pdf_doc(pages_blocks, page_count=None):
    """Build a mock fitz Document with per-page text blocks.

    Args:
        pages_blocks: dict mapping 0-based page index → list of text block tuples.
                      e.g. {0: [_make_text_block(text="摘要")], 1: []}
        page_count: total number of pages (defaults to max key + 1).
    """
    if page_count is None:
        page_count = max(pages_blocks.keys()) + 1 if pages_blocks else 0

    # Each page is its own MagicMock
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


class TestGetSectionStartPages:
    """Unit tests for get_section_start_pages() with mocked fitz."""

    def test_imports_without_fitz_available(self):
        """The function should be importable (real fitz may not be installed)."""
        from enforce.pdf_reader import get_section_start_pages
        assert callable(get_section_start_pages)

    @patch("enforce.pdf_reader.fitz")
    def test_returns_page_numbers_when_heading_found(self, mock_fitz):
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="摘要"),
                _make_text_block(y0=80, text="目录")],
            1: [_make_text_block(y0=80, text="绪论")],
            2: [_make_text_block(y0=85, text="绪论：研究背景")],  # "绪论" is a substring
        }, page_count=3)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), ["绪论"])
        assert result == [3]  # LAST match (avoid TOC), page 3 = 1-based page 3

    @patch("enforce.pdf_reader.fitz")
    def test_returns_none_when_heading_not_found(self, mock_fitz):
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=80, text="摘要")],
            1: [_make_text_block(y0=80, text="参考文献")],
        }, page_count=2)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), ["绪论"])
        assert result == [None]

    @patch("enforce.pdf_reader.fitz")
    def test_multiple_sections_in_order(self, mock_fitz):
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            1: [_make_text_block(y0=90, text="第一章 绪论")],
            3: [_make_text_block(y0=75, text="第二章 方法")],
            4: [_make_text_block(y0=88, text="第三章 实验")],
        }, page_count=5)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(
            Path("fake.pdf"),
            ["第一章 绪论", "第二章 方法", "第三章 实验"],
        )
        assert result == [2, 4, 5]

    @patch("enforce.pdf_reader.fitz")
    def test_cover_heading_special_y_threshold(self, mock_fitz):
        """Cover heading uses standard Y threshold. LAST match returns later page."""
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=150, text="封面")],
            1: [_make_text_block(y0=80, text="封面")],
        }, page_count=2)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), ["封面"])
        assert result == [2]  # LAST match: page 2 (1-based)

    @patch("enforce.pdf_reader.fitz")
    def test_heading_matches_any_y_with_page_text(self, mock_fitz):
        """With page-level normalized matching, headings at any y can match."""
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=200, text="绪论")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), ["绪论"])
        assert result == [1]  # page-level match finds it regardless of y

    @patch("enforce.pdf_reader.fitz")
    def test_last_occurrence_for_non_toc_is_returned(self, mock_fitz):
        """非目录标题使用 LAST 匹配（避免 TOC 误匹配）"""
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            1: [_make_text_block(y0=80, text="实验")],
            2: [_make_text_block(y0=80, text="实验 - 续")],
        }, page_count=4)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), ["实验"])
        assert result == [3]  # LAST: page 3 (1-based), avoids TOC on page 2

    @patch("enforce.pdf_reader.fitz")
    def test_empty_sections_returns_empty_list(self, mock_fitz):
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({}, page_count=5)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), [])
        assert result == []

    @patch("enforce.pdf_reader.fitz")
    def test_y0_70_matches_with_page_text(self, mock_fitz):
        """y0=70 matches via page-level normalized text."""
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=70, text="绪论")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), ["绪论"])
        assert result == [1]

    @patch("enforce.pdf_reader.fitz")
    def test_y0_100_matches_with_page_text(self, mock_fitz):
        """y0=100 matches via page-level normalized text."""
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=100, text="绪论")],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(Path("fake.pdf"), ["绪论"])
        assert result == [1]

    @patch("enforce.pdf_reader.fitz")
    def test_doc_close_is_called(self, mock_fitz):
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [],
        }, page_count=1)
        mock_fitz.open.return_value = mock_doc

        get_section_start_pages(Path("fake.pdf"), ["绪论"])
        mock_doc.close.assert_called_once()

    @patch("enforce.pdf_reader.fitz")
    def test_mixed_found_and_not_found(self, mock_fitz):
        from enforce.pdf_reader import get_section_start_pages

        mock_doc = _mock_pdf_doc({
            0: [_make_text_block(y0=90, text="封面")],
            1: [_make_text_block(y0=80, text="绪论")],
            2: [],  # "结论" not found
        }, page_count=3)
        mock_fitz.open.return_value = mock_doc

        result = get_section_start_pages(
            Path("fake.pdf"),
            ["封面", "绪论", "结论"],
        )
        assert result == [1, 2, None]
