"""Unit tests for ai_pipeline.py — AI pipeline skeleton."""
import pytest
from generator.ai_pipeline import (
    MockAIGenerator, AIOutputParser, analyze_markdown, _auto_detect_type,
)
from generator.models import Section, AnalysisDocument


class TestMockAIGenerator:
    """MockAIGenerator produces the expected mock data."""

    def test_generate_returns_list_of_dicts(self):
        data = MockAIGenerator.generate()
        assert isinstance(data, list)
        assert all(isinstance(item, dict) for item in data)

    def test_generate_has_required_sections(self):
        data = MockAIGenerator.generate()
        types = [item['type'] for item in data]
        assert 'cover' in types
        assert 'toc' in types
        assert 'body' in types
        assert 'references' in types
        assert 'conclusion' in types

    def test_generate_document_returns_analysis_document(self):
        doc = MockAIGenerator.generate_document()
        assert isinstance(doc, AnalysisDocument)
        assert doc.title == "毕业论文"
        assert doc.author == "Mock Author"
        assert len(doc.sections) == len(MockAIGenerator.generate())
        for sec in doc.sections:
            assert isinstance(sec, Section)

    def test_generate_and_generate_document_are_consistent(self):
        data = MockAIGenerator.generate()
        doc = MockAIGenerator.generate_document()
        assert len(doc.sections) == len(data)
        for sec, item in zip(doc.sections, data):
            assert sec.type == item['type']
            assert sec.title == item['title']


class TestAnalyzeMarkdown:
    """analyze_markdown() extracts structure from Markdown text."""

    def test_simple_headings(self):
        md = """# My Thesis

## Introduction

## Conclusion
"""
        doc = analyze_markdown(md)
        assert doc.title == "My Thesis"
        assert len(doc.sections) == 2
        assert doc.sections[0].title == "Introduction"
        assert doc.sections[1].title == "Conclusion"

    def test_front_matter_title_and_author(self):
        md = """---
title: 毕业论文
author: 张三
---

# 毕业论文

## 1 绪论
"""
        doc = analyze_markdown(md)
        assert doc.title == "毕业论文"
        assert doc.author == "张三"
        # First H1 is same as title, so not added as section
        # Second heading "## 1 绪论" is a section
        assert len(doc.sections) >= 1

    def test_auto_detect_types(self):
        md = """# Doc

## 参考文献

## 致谢

## 附录
"""
        doc = analyze_markdown(md)
        types = {s.type for s in doc.sections}
        assert 'references' in types
        assert 'acknowledgments' in types
        assert 'appendix' in types

    def test_empty_markdown(self):
        doc = analyze_markdown("")
        assert doc.title == "未命名文档"
        assert doc.sections == []

    def test_content_not_in_headings_ignored(self):
        md = """# Title

Some paragraph here.

## Chapter 1

More text.
"""
        doc = analyze_markdown(md)
        assert doc.title == "Title"
        assert len(doc.sections) == 1
        assert doc.sections[0].title == "Chapter 1"


class TestAutoDetectType:
    """_auto_detect_type keyword matching."""

    def test_cover_keyword(self):
        sec = Section(type='body', title='封面')
        _auto_detect_type(sec)
        assert sec.type == 'cover'

    def test_statement_keyword(self):
        sec = Section(type='body', title='郑重声明')
        _auto_detect_type(sec)
        assert sec.type == 'statement'

    def test_abstract_cn_keyword(self):
        sec = Section(type='body', title='摘要')
        _auto_detect_type(sec)
        assert sec.type == 'abstract_cn'

    def test_abstract_en_keyword(self):
        sec = Section(type='body', title='ABSTRACT')
        _auto_detect_type(sec)
        assert sec.type == 'abstract_en'

    def test_toc_keyword(self):
        sec = Section(type='body', title='目录')
        _auto_detect_type(sec)
        assert sec.type == 'toc'

    def test_references_keyword(self):
        sec = Section(type='body', title='参考文献')
        _auto_detect_type(sec)
        assert sec.type == 'references'

    def test_conclusion_keyword(self):
        sec = Section(type='body', title='结论')
        _auto_detect_type(sec)
        assert sec.type == 'conclusion'

    def test_acknowledgments_keyword(self):
        sec = Section(type='body', title='致谢')
        _auto_detect_type(sec)
        assert sec.type == 'acknowledgments'

    def test_appendix_keyword(self):
        sec = Section(type='body', title='附录')
        _auto_detect_type(sec)
        assert sec.type == 'appendix'

    def test_body_unchanged_for_unknown(self):
        sec = Section(type='body', title='未知章节')
        _auto_detect_type(sec)
        assert sec.type == 'body'


class TestAIOutputParser:
    """AIOutputParser static methods delegate correctly."""

    def test_from_json_dict(self):
        data = {"title": "Test", "sections": []}
        doc = AIOutputParser.from_json(data)
        assert isinstance(doc, AnalysisDocument)
        assert doc.title == "Test"

    def test_from_markdown(self):
        doc = AIOutputParser.from_markdown("# Hello\n\n## World\n")
        assert doc.title == "Hello"
        assert len(doc.sections) == 1
