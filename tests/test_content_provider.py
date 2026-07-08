"""Unit tests for content_provider.py — Provider registry and base class."""
import pytest
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from generator.content_provider import (
    ContentProvider,
    PROVIDER_REGISTRY,
    add_section_heading,
)


# ── Test helper: minimal concrete provider ──────────────────────────
class _FakeProvider(ContentProvider):
    section_type = "fake_test"

    def render(self, doc, section, config):
        doc.add_paragraph(f"Rendered: {section.title}")


# ══════════════════════════════════════════════════════════════════════
# ContentProvider registration / lookup
# ══════════════════════════════════════════════════════════════════════

class TestContentProviderRegistry:
    """ContentProvider.register() and get_provider()."""

    def test_register_adds_to_registry(self):
        ContentProvider.register(_FakeProvider)
        assert "fake_test" in PROVIDER_REGISTRY
        assert isinstance(PROVIDER_REGISTRY["fake_test"], _FakeProvider)

    def test_get_provider_returns_registered(self):
        ContentProvider.register(_FakeProvider)
        provider = ContentProvider.get_provider("fake_test")
        assert isinstance(provider, _FakeProvider)

    def test_get_provider_raises_KeyError_for_unknown(self):
        with pytest.raises(KeyError):
            ContentProvider.get_provider("nonexistent_type")

    def test_register_overwrites_previous(self):
        class _FakeOverride(ContentProvider):
            section_type = "fake_test"

            def render(self, doc, section, config):
                pass

        ContentProvider.register(_FakeProvider)
        assert isinstance(PROVIDER_REGISTRY["fake_test"], _FakeProvider)

        ContentProvider.register(_FakeOverride)
        assert isinstance(PROVIDER_REGISTRY["fake_test"], _FakeOverride)

    def test_register_multiple_providers(self):
        class _FakeA(ContentProvider):
            section_type = "type_a"

            def render(self, doc, section, config):
                pass

        class _FakeB(ContentProvider):
            section_type = "type_b"

            def render(self, doc, section, config):
                pass

        ContentProvider.register(_FakeA)
        ContentProvider.register(_FakeB)

        assert isinstance(
            ContentProvider.get_provider("type_a"), _FakeA
        )
        assert isinstance(
            ContentProvider.get_provider("type_b"), _FakeB
        )


# ══════════════════════════════════════════════════════════════════════
# Registered providers are singletons (created once at register time)
# ══════════════════════════════════════════════════════════════════════

class TestProviderSingleton:
    """Registered provider instances are created at registration."""

    def test_same_instance_returned(self):
        class _Fake(ContentProvider):
            section_type = "singleton_test"

            def render(self, doc, section, config):
                pass

        ContentProvider.register(_Fake)
        p1 = ContentProvider.get_provider("singleton_test")
        p2 = ContentProvider.get_provider("singleton_test")
        assert p1 is p2


# ══════════════════════════════════════════════════════════════════════
# add_section_heading
# ══════════════════════════════════════════════════════════════════════

class TestAddSectionHeading:
    """add_section_heading() utility."""

    def test_adds_paragraph_with_title(self, blank_doc):
        p = add_section_heading(blank_doc, "Chapter 1")
        assert p.text == "Chapter 1"
        assert len(blank_doc.paragraphs) == 1

    def test_default_alignment_center(self, blank_doc):
        p = add_section_heading(blank_doc, "Title")
        assert p.alignment == WD_ALIGN_PARAGRAPH.CENTER

    def test_custom_alignment(self, blank_doc):
        p = add_section_heading(
            blank_doc, "Left Title", alignment=WD_ALIGN_PARAGRAPH.LEFT
        )
        assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT

    def test_default_font_size(self, blank_doc):
        p = add_section_heading(blank_doc, "Default Size")
        run = p.runs[0]
        assert run.font.size == Pt(16)

    def test_custom_font_size(self, blank_doc):
        p = add_section_heading(blank_doc, "Big Title", font_size=24)
        run = p.runs[0]
        assert run.font.size == Pt(24)

    def test_bold(self, blank_doc):
        p = add_section_heading(blank_doc, "Bold Title")
        run = p.runs[0]
        assert run.bold is True

    def test_font_size_from_config(self, blank_doc):
        config = {"fonts": {"heading": {"size": 20}}}
        p = add_section_heading(blank_doc, "From Config", config=config)
        run = p.runs[0]
        assert run.font.size == Pt(20)

    def test_font_size_from_config_partial(self, blank_doc):
        """Config missing fonts/heading/size should fall back to default 16."""
        config = {"fonts": {}}  # no heading key
        p = add_section_heading(blank_doc, "Partial Config", config=config)
        run = p.runs[0]
        assert run.font.size == Pt(16)

    def test_config_none_uses_default_size(self, blank_doc):
        p = add_section_heading(blank_doc, "No Config", config=None)
        run = p.runs[0]
        assert run.font.size == Pt(16)

    def test_explicit_font_size_overrides_config(self, blank_doc):
        config = {"fonts": {"heading": {"size": 20}}}
        p = add_section_heading(
            blank_doc, "Explicit", config=config, font_size=30
        )
        run = p.runs[0]
        assert run.font.size == Pt(30)


# ══════════════════════════════════════════════════════════════════════
# KeyError → ValueError conversion (if implemented in future)
# For now, verify KeyError behavior directly
# ══════════════════════════════════════════════════════════════════════

class TestKeyErrorToValueError:
    """ContentProvider.get_provider raises KeyError on unknown type.
    (KeyError → ValueError wrapping can be added if desired.)"""

    def test_keyerror_on_unknown_type(self):
        with pytest.raises(KeyError) as exc_info:
            ContentProvider.get_provider("definitely_not_a_type")
        # The key name is in the message
        assert "definitely_not_a_type" in str(exc_info.value)
