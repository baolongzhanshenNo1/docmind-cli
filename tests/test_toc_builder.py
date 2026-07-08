"""Unit tests for toc_builder.py — Dynamic TOC page number builder."""
import pytest
from generator.toc_builder import (
    estimate_page_count, build_toc_entries, format_toc_line, format_toc,
    DEFAULT_SECTION_PAGES, CHARS_PER_PAGE,
)
from generator.models import Section


class TestEstimatePageCount:
    """estimate_page_count() per-section page estimation."""

    def test_default_for_body_without_content(self):
        sec = Section(type='body', title='Chapter 1')
        assert estimate_page_count(sec) == DEFAULT_SECTION_PAGES['body']

    def test_default_for_cover(self):
        sec = Section(type='cover', title='Cover')
        assert estimate_page_count(sec) == DEFAULT_SECTION_PAGES['cover']

    def test_default_for_unknown_type(self):
        sec = Section(type='unknown', title='Something')
        assert estimate_page_count(sec) == 3  # fallback

    def test_content_based_estimation(self):
        # content with exactly 1 page worth of characters
        content = 'x' * CHARS_PER_PAGE
        sec = Section(type='body', title='Ch', content=content)
        assert estimate_page_count(sec) == 1

    def test_content_spans_multiple_pages(self):
        content = 'x' * (CHARS_PER_PAGE * 2 + 1)
        sec = Section(type='body', title='Ch', content=content)
        assert estimate_page_count(sec) == 3

    def test_minimum_one_page(self):
        sec = Section(type='body', title='Ch', content='short')
        assert estimate_page_count(sec) == 1


class TestBuildTocEntries:
    """build_toc_entries() generates TOC entry list from sections."""

    def test_empty_sections(self):
        entries = build_toc_entries([], start_page=1)
        assert entries == []

    def test_simple_linear_pages(self):
        sections = [
            Section(type='body', title='Ch1'),
            Section(type='body', title='Ch2'),
        ]
        entries = build_toc_entries(sections, start_page=1)
        assert len(entries) == 2
        assert entries[0]['title'] == 'Ch1'
        assert entries[0]['page'] == 1
        # Ch2 starts after Ch1's pages (5 default for body)
        assert entries[1]['title'] == 'Ch2'
        assert entries[1]['page'] == 1 + DEFAULT_SECTION_PAGES['body']

    def test_skips_toc_section(self):
        sections = [
            Section(type='toc', title='目录'),
            Section(type='body', title='Ch1'),
        ]
        entries = build_toc_entries(sections, start_page=1)
        assert len(entries) == 1
        assert entries[0]['title'] == 'Ch1'

    def test_supports_custom_toc_label(self):
        sections = [
            Section(type='toc', title='Table of Contents'),
            Section(type='body', title='Ch1'),
        ]
        entries = build_toc_entries(sections, start_page=1, toc_label='Table of Contents')
        assert len(entries) == 1

    def test_custom_start_page(self):
        sections = [
            Section(type='body', title='Ch1'),
        ]
        entries = build_toc_entries(sections, start_page=10)
        assert entries[0]['page'] == 10


class TestFormatTocLine:
    """format_toc_line() formats a single TOC entry."""

    def test_basic_format(self):
        entry = {'title': '1 绪论', 'page': 1}
        line = format_toc_line(entry)
        assert line.startswith('1 绪论')
        assert line.endswith('1')
        assert '..' in line or '.' in line  # dotted leader

    def test_page_number_present(self):
        entry = {'title': '参考文献', 'page': 25}
        line = format_toc_line(entry)
        assert '25' in line

    def test_very_long_title(self):
        entry = {'title': '这是一个非常非常非常非常非常非常非常长的章节标题', 'page': 1}
        line = format_toc_line(entry)
        # Should not crash, uses abbreviated format
        assert '1' in line


class TestFormatToc:
    """format_toc() generates a complete TOC text block."""

    def test_includes_title(self):
        entries = []
        text = format_toc(entries, title='目录')
        assert '目录' in text

    def test_includes_entries(self):
        entries = [
            {'title': 'Ch1', 'page': 1},
            {'title': 'Ch2', 'page': 6},
        ]
        text = format_toc(entries)
        assert 'Ch1' in text
        assert 'Ch2' in text

    def test_newlines_separate_entries(self):
        entries = [
            {'title': 'A', 'page': 1},
            {'title': 'B', 'page': 2},
        ]
        text = format_toc(entries)
        lines = text.split('\n')
        # title + blank + 2 entries = 4 lines
        assert len(lines) >= 3
