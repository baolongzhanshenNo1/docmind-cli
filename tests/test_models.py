"""Unit tests for models.py — Section / AnalysisDocument dataclasses."""
import json
import os
import tempfile
from generator.models import Section, AnalysisDocument


class TestSection:
    """Section dataclass construction and default values."""

    def test_construct_all_fields(self):
        s = Section(
            type="body",
            title="Introduction",
            content="Hello world.",
            level=2,
            metadata={"pages": [1, 2]},
        )
        assert s.type == "body"
        assert s.title == "Introduction"
        assert s.content == "Hello world."
        assert s.level == 2
        assert s.metadata == {"pages": [1, 2]}

    def test_default_content(self):
        s = Section(type="cover", title="Cover Page")
        assert s.content == ""

    def test_default_level(self):
        s = Section(type="toc", title="Contents")
        assert s.level == 1

    def test_default_metadata(self):
        s = Section(type="appendix", title="Appendix A")
        assert s.metadata == {}

    def test_equality(self):
        s1 = Section(type="body", title="Intro", content="x")
        s2 = Section(type="body", title="Intro", content="x")
        assert s1 == s2

    def test_inequality(self):
        s1 = Section(type="body", title="Intro")
        s2 = Section(type="body", title="Summary")
        assert s1 != s2


class TestAnalysisDocumentFromDict:
    """AnalysisDocument.from_dict() construction."""

    def test_minimal_dict(self):
        data = {"title": "My Paper"}
        doc = AnalysisDocument.from_dict(data)
        assert doc.title == "My Paper"
        assert doc.author == ""
        assert doc.sections == []
        assert doc.metadata == {}

    def test_full_dict(self):
        data = {
            "title": "Full Paper",
            "author": "Jane Doe",
            "metadata": {"lang": "zh"},
            "sections": [
                {
                    "type": "cover",
                    "title": "Cover",
                    "content": "C",
                    "level": 1,
                    "metadata": {"page": 1},
                },
                {
                    "type": "body",
                    "title": "Body",
                    "content": "B",
                    "level": 2,
                },
            ],
        }
        doc = AnalysisDocument.from_dict(data)
        assert doc.title == "Full Paper"
        assert doc.author == "Jane Doe"
        assert doc.metadata == {"lang": "zh"}
        assert len(doc.sections) == 2

        s0 = doc.sections[0]
        assert isinstance(s0, Section)
        assert s0.type == "cover"
        assert s0.title == "Cover"
        assert s0.content == "C"
        assert s0.level == 1
        assert s0.metadata == {"page": 1}

        s1 = doc.sections[1]
        assert s1.type == "body"
        assert s1.level == 2
        assert s1.metadata == {}  # default

    def test_extra_keys_ignored(self):
        data = {"title": "T", "unknown": 42, "sections": []}
        doc = AnalysisDocument.from_dict(data)
        assert doc.title == "T"
        assert len(doc.sections) == 0

    def test_section_defaults_on_missing_keys(self):
        data = {
            "title": "T",
            "sections": [{"type": "body"}],  # no title, content, etc.
        }
        doc = AnalysisDocument.from_dict(data)
        s = doc.sections[0]
        assert s.title == ""
        assert s.content == ""
        assert s.level == 1
        assert s.metadata == {}


class TestAnalysisDocumentFromJson:
    """AnalysisDocument.from_json() loading from file."""

    def test_from_json_roundtrip(self):
        data = {
            "title": "Roundtrip Doc",
            "author": "Author",
            "sections": [
                {
                    "type": "abstract_cn",
                    "title": "摘要",
                    "content": "中文摘要内容",
                }
            ],
        }
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8"
        ) as f:
            json.dump(data, f)
            tmp_path = f.name

        try:
            doc = AnalysisDocument.from_json(tmp_path)
            assert doc.title == "Roundtrip Doc"
            assert doc.author == "Author"
            assert len(doc.sections) == 1
            assert doc.sections[0].type == "abstract_cn"
            assert doc.sections[0].title == "摘要"
        finally:
            os.unlink(tmp_path)

    def test_from_json_empty_file(self):
        data = {"title": "Empty"}
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8"
        ) as f:
            json.dump(data, f)
            tmp_path = f.name
        try:
            doc = AnalysisDocument.from_json(tmp_path)
            assert doc.title == "Empty"
            assert doc.sections == []
        finally:
            os.unlink(tmp_path)
