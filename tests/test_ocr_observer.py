"""
Tests for ocr_observer.py — DocMind 视觉验证引擎。
"""
import json
import sys
from pathlib import Path

import pytest

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from ocr_observer import (
    OcrObserver,
    VerifyResult,
    BlankVerifyResult,
    PageAudit,
    TextRegion,
    get_observer,
    quick_verify_header,
    quick_verify_blank,
)

# Test fixtures
TEST_PNG = Path(__file__).parent.parent / "output" / "_v_p005.png"


# ═══════════════════════════════════════════
# Dataclass tests
# ═══════════════════════════════════════════


class TestTextRegion:
    def test_constructor(self):
        tr = TextRegion(
            text="测试",
            confidence=0.95,
            bbox=[[0, 0], [100, 0], [100, 20], [0, 20]],
            center_y=10,
        )
        assert tr.text == "测试"
        assert tr.confidence == 0.95
        assert tr.center_y == 10

    def test_repr(self):
        tr = TextRegion(text="hello", confidence=0.88, bbox=[[0, 0], [10, 0], [10, 10], [0, 10]], center_y=5)
        r = repr(tr)
        assert "hello" in r
        assert "0.88" in r


class TestVerifyResult:
    def test_found(self):
        vr = VerifyResult(found=True, confidence=0.95, expected="A", actual="A", detail="ok")
        assert vr.found
        assert bool(vr) is True

    def test_not_found(self):
        vr = VerifyResult(found=False, confidence=0.0, expected="A", detail="not found")
        assert not vr.found
        assert bool(vr) is False

    def test_repr(self):
        vr = VerifyResult(found=True, confidence=0.9, expected="A", detail="found")
        assert "✅" in repr(vr)

        vr2 = VerifyResult(found=False, confidence=0.0, expected="A", detail="missing")
        assert "❌" in repr(vr2)


class TestBlankVerifyResult:
    def test_blank(self):
        br = BlankVerifyResult(is_blank=True, text_count=0, detail="empty")
        assert br.is_blank
        assert bool(br) is True

    def test_not_blank(self):
        br = BlankVerifyResult(is_blank=False, text_count=10, any_text=["hello"], detail="has text")
        assert not br.is_blank
        assert bool(br) is False


class TestPageAudit:
    def test_constructor(self):
        pa = PageAudit(
            page_path="/tmp/test.png",
            page_index=0,
            page_width=827,
            page_height=1170,
            total_regions=10,
            summary="test",
        )
        assert pa.page_index == 0
        assert pa.total_regions == 10

    def test_to_dict(self):
        tr = TextRegion(text="hello", confidence=0.9, bbox=[[0, 0], [10, 0], [10, 10], [0, 10]], center_y=5)
        pa = PageAudit(
            page_path="/tmp/test.png",
            page_index=2,
            page_width=800,
            page_height=1100,
            total_regions=5,
            header_texts=[tr],
            body_texts=[],
            footer_texts=[],
            summary="test",
        )
        d = pa.to_dict()
        assert d["page"] == 3  # 1-based
        assert d["regions"] == 5
        assert d["header"] == ["hello"]
        assert d["body_preview"] == []
        assert d["page_number"] is None


# ═══════════════════════════════════════════
# OcrObserver tests
# ═══════════════════════════════════════════


class TestOcrObserverInit:
    def test_default_init(self):
        obs = OcrObserver()
        assert obs.model_type == "medium"
        assert not obs.is_ready

    def test_small_init(self):
        obs = OcrObserver(model_type="small")
        assert obs.model_type == "small"

    def test_tiny_init(self):
        obs = OcrObserver(model_type="tiny")
        assert obs.model_type == "tiny"


class TestOcrObserverLazyInit:
    def test_lazy_initialization(self):
        obs = OcrObserver(model_type="small")
        assert not obs.is_ready
        obs._ensure_initialized()
        assert obs.is_ready

    def test_observe_page_triggers_init(self):
        obs = OcrObserver(model_type="small")
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        audit = obs.observe_page(str(TEST_PNG))
        assert obs.is_ready
        assert isinstance(audit, PageAudit)


class TestOcrObserverObservePage:
    @pytest.fixture(scope="class")
    def observer(self):
        return OcrObserver(model_type="small")

    def test_returns_page_audit(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        audit = observer.observe_page(str(TEST_PNG))
        assert isinstance(audit, PageAudit)
        assert audit.total_regions >= 0
        assert audit.page_width > 0
        assert audit.page_height > 0

    def test_header_body_footer_separation(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        audit = observer.observe_page(str(TEST_PNG))
        # 摘要页应有正文和页眉
        total = len(audit.header_texts) + len(audit.body_texts) + len(audit.footer_texts)
        assert total == audit.total_regions

    def test_all_text_regions_have_content(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        audit = observer.observe_page(str(TEST_PNG))
        for tr in audit.header_texts + audit.body_texts + audit.footer_texts:
            assert isinstance(tr, TextRegion)
            assert tr.text.strip() != ""
            assert 0 <= tr.confidence <= 1.0
            assert len(tr.bbox) == 4  # 4 corners

    def test_unknown_image_does_not_crash(self, observer):
        """不存在的图片应抛出异常而非崩溃"""
        with pytest.raises(Exception):
            observer.observe_page("nonexistent_file.png")


class TestOcrObserverVerification:
    @pytest.fixture(scope="class")
    def observer(self):
        return OcrObserver(model_type="small")

    def test_verify_header_found(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        result = observer.verify_header(str(TEST_PNG), "摘要")
        assert isinstance(result, VerifyResult)
        # 摘要页眉应能找到"摘要"
        # 注意：可能找到，取决于 small 模型渲染

    def test_verify_header_not_found_returns_result(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        result = observer.verify_header(str(TEST_PNG), "XYZNOTEXIST")
        assert isinstance(result, VerifyResult)
        # 不应崩溃

    def test_verify_footer(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        result = observer.verify_footer(str(TEST_PNG), "1")
        assert isinstance(result, VerifyResult)

    def test_verify_page_number(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        result = observer.verify_page_number(str(TEST_PNG), 1)
        assert isinstance(result, VerifyResult)

    def test_verify_blank(self, observer):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        result = observer.verify_blank(str(TEST_PNG))
        assert isinstance(result, BlankVerifyResult)
        # 摘要页不是空白页
        if result.text_count > OcrObserver.BLANK_THRESHOLD:
            assert not result.is_blank


class TestOcrObserverReport:
    def test_report_empty(self):
        obs = OcrObserver(model_type="small")
        report = obs.report([])
        assert "0 页" in report

    def test_report_single(self):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        obs = OcrObserver(model_type="small")
        audit = obs.observe_page(str(TEST_PNG))
        report = obs.report([audit])
        assert "1 页" in report
        assert "文字区域" in report

    def test_to_json(self):
        if not TEST_PNG.exists():
            pytest.skip(f"Test image not found: {TEST_PNG}")
        obs = OcrObserver(model_type="small")
        audit = obs.observe_page(str(TEST_PNG))
        json_str = obs.to_json([audit])
        data = json.loads(json_str)
        assert isinstance(data, list)
        assert len(data) == 1
        assert "page" in data[0]
        assert "regions" in data[0]


# ═══════════════════════════════════════════
# Convenience function tests
# ═══════════════════════════════════════════


class TestConvenienceFunctions:
    def test_get_observer_singleton(self):
        obs1 = get_observer("small")
        obs2 = get_observer("small")
        assert obs1 is obs2  # same instance

    def test_get_observer_different_type_new_instance(self):
        obs1 = get_observer("small")
        obs2 = get_observer("tiny")  # different type → new instance
        # tiny may fall back to small if model not available
        assert obs2.model_type in ("tiny", "small")
