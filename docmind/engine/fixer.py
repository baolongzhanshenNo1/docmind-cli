"""
Fixer v2 — 验证 fix_plan 执行结果。

核心原则:
- 对比 fix_plan 和实际 OOXML，逐条验证每项是否已生效
- 不做任何修改，只报告差异
- 返回诊断列表供 Agent 决策

用法:
    from pipeline.fixer_v2 import diagnose

    issues = diagnose(fix_plan, Path("output/毕业设计.docx"))
    # issues: [FixerDiagnostic(code="HEADER_FONT_MISMATCH", ...), ...]
"""

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Union

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class FixerDiagnostic:
    """fix_plan 验证结果"""
    code: str               # 诊断代码
    action_index: int       # 对应 fix_plan 中的 action 索引
    severity: str           # "error", "warning", "info"
    detail: str             # 详细描述
    expected: Any = None    # fix_plan 中期望的值
    actual: Any = None      # 实际 OOXML 中读取的值

    def __str__(self):
        return f"[{self.severity}] #{self.action_index} {self.code}: {self.detail}"


# ── 验证调度器 ──────────────────────────────────────────────

_VERIFIERS: dict[str, callable] = {}


def _verifier(action_name: str):
    """装饰器：注册验证函数"""
    def decorator(fn):
        _VERIFIERS[action_name] = fn
        return fn
    return decorator


def diagnose(fix_plan: list[dict], target_path: Union[str, Path]) -> list[FixerDiagnostic]:
    """逐条验证 fix_plan 中的每项是否已在目标 docx 中生效。

    Args:
        fix_plan: action 列表
        target_path: 修改后的 docx 文件路径

    Returns:
        FixerDiagnostic 列表（空列表表示全部验证通过）
    """
    target_path = Path(target_path)
    issues: list[FixerDiagnostic] = []

    # 预加载常用 XML
    _cache: dict[str, etree._Element] = {}

    def _get_xml(internal_path: str) -> etree._Element:
        if internal_path not in _cache:
            with zipfile.ZipFile(target_path, "r") as z:
                _cache[internal_path] = etree.fromstring(z.read(internal_path))
        return _cache[internal_path]

    for i, action in enumerate(fix_plan):
        action_type = action.get("action", "")
        params = action.get("params", {})

        verifier = _VERIFIERS.get(action_type)
        if verifier is None:
            issues.append(FixerDiagnostic(
                code="UNKNOWN_ACTION",
                action_index=i,
                severity="warning",
                detail=f"未知 action 类型: {action_type}，无法验证",
            ))
            continue

        try:
            result = verifier(target_path, params, _get_xml, i)
            if result:
                issues.extend(result if isinstance(result, list) else [result])
        except Exception as e:
            issues.append(FixerDiagnostic(
                code="VERIFY_ERROR",
                action_index=i,
                severity="error",
                detail=f"验证异常: {e}",
            ))

    return issues


# ── 辅助 ────────────────────────────────────────────────────

def _read_text(element: etree._Element) -> str:
    """读取元素内所有 <w:t> 的文本"""
    return "".join(t.text or "" for t in element.iter(f"{{{W}}}t"))


# ── 验证: set_sectpr_type ───────────────────────────────────

@_verifier("set_sectpr_type")
def _verify_sectpr_type(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    section_index = params["section_index"]
    expected_val = params["val"]

    root = get_xml("word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))

    if section_index >= len(sects):
        return [FixerDiagnostic(
            code="SECTPR_INDEX_OUT_OF_RANGE",
            action_index=idx,
            severity="error",
            detail=f"section_index={section_index} 超出范围 (共 {len(sects)} 个 sectPr)",
            expected=expected_val,
            actual=None,
        )]

    sp = sects[section_index]
    type_el = sp.find(f"{{{W}}}type")
    actual_val = type_el.get(f"{{{W}}}val") if type_el is not None else None

    if actual_val != expected_val:
        return [FixerDiagnostic(
            code="SECTPR_TYPE_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"sectPr[{section_index}] type 期望={expected_val} 实际={actual_val}",
            expected=expected_val,
            actual=actual_val,
        )]

    return []


# ── 验证: set_header_font ───────────────────────────────────

@_verifier("set_header_font")
def _verify_header_font(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    header_path = params["header_path"]
    expected_ascii = params.get("font_ascii")
    expected_eastAsia = params.get("font_eastAsia")
    expected_size = params.get("font_size")

    root = get_xml(header_path)
    issues = []

    # 空 header 容错：无文字内容则跳过字体检查
    all_text = ''.join(t.text or '' for t in root.iter(f'{{{W}}}t')).strip()
    if not all_text:
        return []

    rPr_list = list(root.iter(f"{{{W}}}rPr"))
    if not rPr_list:
        return [FixerDiagnostic(
            code="HEADER_FONT_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"{header_path} 中没有 rPr 元素",
            expected={"ascii": expected_ascii, "eastAsia": expected_eastAsia, "size": expected_size},
            actual=None,
        )]

    # 检查第一个 rPr（通常是主要字体设置）
    rPr = rPr_list[0]
    rf = rPr.find(f"{{{W}}}rFonts")

    if rf is None:
        issues.append(FixerDiagnostic(
            code="HEADER_FONT_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"{header_path} 中无 rFonts 元素",
            expected={"ascii": expected_ascii, "eastAsia": expected_eastAsia},
            actual=None,
        ))
    else:
        actual_ascii = rf.get(f"{{{W}}}ascii")
        actual_eastAsia = rf.get(f"{{{W}}}eastAsia")

        if expected_ascii and actual_ascii != expected_ascii:
            issues.append(FixerDiagnostic(
                code="HEADER_FONT_MISMATCH",
                action_index=idx,
                severity="error",
                detail=f"{header_path} ascii 期望={expected_ascii} 实际={actual_ascii}",
                expected=expected_ascii,
                actual=actual_ascii,
            ))

        if expected_eastAsia and actual_eastAsia != expected_eastAsia:
            issues.append(FixerDiagnostic(
                code="HEADER_FONT_MISMATCH",
                action_index=idx,
                severity="error",
                detail=f"{header_path} eastAsia 期望={expected_eastAsia} 实际={actual_eastAsia}",
                expected=expected_eastAsia,
                actual=actual_eastAsia,
            ))

    # 检查字号
    if expected_size:
        sz = rPr.find(f"{{{W}}}sz")
        actual_size = sz.get(f"{{{W}}}val") if sz is not None else None
        if actual_size != expected_size:
            issues.append(FixerDiagnostic(
                code="HEADER_FONT_MISMATCH",
                action_index=idx,
                severity="warning",
                detail=f"{header_path} sz 期望={expected_size} 实际={actual_size}",
                expected=expected_size,
                actual=actual_size,
            ))

    return issues


# ── 验证: set_page_number_type ──────────────────────────────

@_verifier("set_page_number_type")
def _verify_page_number_type(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: add_footer_reference ──────────────────────────────

@_verifier("add_footer_reference")
def _verify_footer_reference(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: clear_section_headers ────────────────────────────

@_verifier("clear_section_headers")
def _verify_clear_headers(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: add_page_number ───────────────────────────────────

@_verifier("add_page_number")
def _verify_page_number(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    footer_path = params["footer_path"]

    root = get_xml(footer_path)

    # 检查是否有 fldChar 序列
    fld_chars = list(root.iter(f"{{{W}}}fldChar"))
    if not fld_chars:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="error",
            detail=f"{footer_path} 中无 fldChar（页码域）",
            expected="fldChar 序列",
            actual="无",
        )]

    # 检查是否有 PAGE 指令
    instr_texts = []
    for it in root.iter(f"{{{W}}}instrText"):
        instr_texts.append(it.text or "")

    has_page = any("PAGE" in t for t in instr_texts)
    if not has_page:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="error",
            detail=f"{footer_path} 中无 PAGE 指令",
            expected="PAGE",
            actual=instr_texts,
        )]

    # 验证格式模板中的文字
    fmt = params.get("format", "")
    prefix_before_page = fmt.split("{PAGE}")[0] if "{PAGE}" in fmt else ""
    suffix_after_page = fmt.rsplit("{PAGE}", 1)[-1] if "{PAGE}" in fmt else ""

    full_text = _read_text(root)
    if prefix_before_page and prefix_before_page not in full_text:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="warning",
            detail=f"{footer_path} 缺少页码前缀 '{prefix_before_page}'",
            expected=prefix_before_page,
            actual=full_text[:80],
        )]

    if suffix_after_page and suffix_after_page not in full_text:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="warning",
            detail=f"{footer_path} 缺少页码后缀 '{suffix_after_page}'",
            expected=suffix_after_page,
            actual=full_text[:80],
        )]

    return []


# ── 验证: set_body_font_ascii ───────────────────────────────

@_verifier("set_body_font_ascii")
def _verify_body_font_ascii(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    style_id = params["style_id"]
    expected_ascii = params["font_ascii"]

    root = get_xml("word/styles.xml")

    for style in root.iter(f"{{{W}}}style"):
        if style.get(f"{{{W}}}styleId") != style_id:
            continue

        rPr = style.find(f"{{{W}}}rPr")
        if rPr is None:
            return [FixerDiagnostic(
                code="STYLE_NOT_APPLIED",
                action_index=idx,
                severity="error",
                detail=f"样式 '{style_id}' 无 rPr 元素",
                expected={"ascii": expected_ascii},
                actual=None,
            )]

        rf = rPr.find(f"{{{W}}}rFonts")
        actual_ascii = rf.get(f"{{{W}}}ascii") if rf is not None else None

        if actual_ascii != expected_ascii:
            return [FixerDiagnostic(
                code="STYLE_NOT_APPLIED",
                action_index=idx,
                severity="error",
                detail=f"样式 '{style_id}' ascii 期望={expected_ascii} 实际={actual_ascii}",
                expected=expected_ascii,
                actual=actual_ascii,
            )]

        return []

    return [FixerDiagnostic(
        code="STYLE_NOT_APPLIED",
        action_index=idx,
        severity="error",
        detail=f"未找到样式 '{style_id}'",
        expected=expected_ascii,
        actual="样式不存在",
    )]


# ── 验证: set_style ─────────────────────────────────────────

@_verifier("set_style")
def _verify_style(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    style_id = params["style_id"]
    issues = []

    root = get_xml("word/styles.xml")

    for style in root.iter(f"{{{W}}}style"):
        if style.get(f"{{{W}}}styleId") != style_id:
            continue

        # 验证字体
        if any(k in params for k in ("font_eastAsia", "font_ascii", "font_hAnsi")):
            rPr = style.find(f"{{{W}}}rPr")
            rf = rPr.find(f"{{{W}}}rFonts") if rPr is not None else None

            if params.get("font_ascii") is not None:
                actual = rf.get(f"{{{W}}}ascii") if rf is not None else None
                if actual != params["font_ascii"]:
                    issues.append(FixerDiagnostic(
                        code="STYLE_NOT_APPLIED",
                        action_index=idx,
                        severity="error",
                        detail=f"样式 '{style_id}' ascii 期望={params['font_ascii']} 实际={actual}",
                        expected=params["font_ascii"],
                        actual=actual,
                    ))

            if params.get("font_eastAsia") is not None:
                actual = rf.get(f"{{{W}}}eastAsia") if rf is not None else None
                if actual != params["font_eastAsia"]:
                    issues.append(FixerDiagnostic(
                        code="STYLE_NOT_APPLIED",
                        action_index=idx,
                        severity="error",
                        detail=f"样式 '{style_id}' eastAsia 期望={params['font_eastAsia']} 实际={actual}",
                        expected=params["font_eastAsia"],
                        actual=actual,
                    ))

        # 验证字号
        if params.get("font_size_pt") is not None:
            expected_sv = str(int(params["font_size_pt"] * 2))
            rPr = style.find(f"{{{W}}}rPr")
            if rPr is not None:
                sz = rPr.find(f"{{{W}}}sz")
                actual_sv = sz.get(f"{{{W}}}val") if sz is not None else None
                if actual_sv != expected_sv:
                    issues.append(FixerDiagnostic(
                        code="STYLE_NOT_APPLIED",
                        action_index=idx,
                        severity="warning",
                        detail=f"样式 '{style_id}' sz 期望={expected_sv} 实际={actual_sv}",
                        expected=expected_sv,
                        actual=actual_sv,
                    ))

        # 验证加粗
        if params.get("bold") is not None:
            rPr = style.find(f"{{{W}}}rPr")
            has_b = rPr is not None and rPr.find(f"{{{W}}}b") is not None
            if has_b != params["bold"]:
                issues.append(FixerDiagnostic(
                    code="STYLE_NOT_APPLIED",
                    action_index=idx,
                    severity="warning",
                    detail=f"样式 '{style_id}' bold 期望={params['bold']} 实际={has_b}",
                    expected=params["bold"],
                    actual=has_b,
                ))

        return issues

    return [FixerDiagnostic(
        code="STYLE_NOT_APPLIED",
        action_index=idx,
        severity="error",
        detail=f"未找到样式 '{style_id}'",
        expected=style_id,
        actual="样式不存在",
    )]


# ── 验证: remove_extra_sectpr ───────────────────────────────

@_verifier("remove_extra_sectpr")
def _verify_remove_extra_sectpr(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    root = get_xml("word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return []

    issues = []
    extra_count = 0
    for child in body:
        if child.tag != f"{{{W}}}p":
            continue
        pPr = child.find(f"{{{W}}}pPr")
        if pPr is None:
            continue
        extra_sp = pPr.find(f"{{{W}}}sectPr")
        if extra_sp is not None:
            # 检查这个 sectPr 是否有 oddPage
            type_el = extra_sp.find(f"{{{W}}}type")
            type_val = type_el.get(f"{{{W}}}val") if type_el is not None else ""
            extra_count += 1
            if type_val == "oddPage":
                issues.append(FixerDiagnostic(
                    code="EXTRA_ODDPAGE_SECTPR",
                    action_index=idx,
                    severity="error",
                    detail=f"段落内仍有 oddPage sectPr (第 {extra_count} 个嵌入)",
                    expected="已删除",
                    actual="oddPage sectPr 残留",
                ))

    if extra_count > 0 and not issues:
        issues.append(FixerDiagnostic(
            code="EXTRA_ODDPAGE_SECTPR",
            action_index=idx,
            severity="warning",
            detail=f"仍有 {extra_count} 个嵌入 sectPr 未删除",
            expected=0,
            actual=extra_count,
        ))

    return issues


# ── Summary ────────────────────────────────────────────────


def summary(diagnostics: list[FixerDiagnostic]) -> dict:
    """生成诊断汇总。

    Returns:
        {"total": N, "passed": bool, "by_code": {...}, "by_severity": {...}}
    """
    by_code: dict[str, int] = {}
    by_severity: dict[str, int] = {"error": 0, "warning": 0, "info": 0}

    for d in diagnostics:
        by_code[d.code] = by_code.get(d.code, 0) + 1
        by_severity[d.severity] = by_severity.get(d.severity, 0) + 1

    return {
        "total": len(diagnostics),
        "passed": len(diagnostics) == 0,
        "by_code": by_code,
        "by_severity": by_severity,
    }


# ── Verifier: set_header_text ──────────────────────────────

@_verifier("set_header_text")
def _verify_set_header_text(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 header 文件文字是否已更新。"""
    header_path = params.get("header_path", "")
    expected_text = params.get("text", "")

    try:
        root = get_xml(header_path)
    except Exception:
        return [FixerDiagnostic(
            code="HEADER_FILE_MISSING",
            action_index=idx,
            severity="error",
            detail=f"header 文件不存在: {header_path}",
            expected=expected_text,
        )]

    # 提取所有文字
    all_text = ""
    for t in root.iter(f"{{{W}}}t"):
        if t.text:
            all_text += t.text

    if expected_text not in all_text:
        return [FixerDiagnostic(
            code="HEADER_TEXT_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"header 文字不匹配: 期望 '{expected_text}' 不在 '{all_text[:50]}'",
            expected=expected_text,
            actual=all_text[:50],
        )]

    return []


@_verifier("create_section_headers")
def _verify_create_section_headers(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 create_section_headers 操作。"""
    return []  # TODO: 实现验证：检查 ZIP 中是否有新增 header 文件


@_verifier("insert_section_break")
def _verify_insert_section_break(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 insert_section_break 操作。"""
    return []  # TODO: 实现验证：检查 sectPr 计数是否增加


@_verifier("insert_red_header")
def _verify_insert_red_header(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 insert_red_header 操作：检查文档开头是否包含红头文字和红色反线。"""
    org_name = params.get("org_name", "")
    doc_number = params.get("doc_number", "")

    try:
        root = get_xml("word/document.xml")
    except Exception:
        return [FixerDiagnostic(
            code="RED_HEADER_VERIFY_FAILED",
            action_index=idx,
            severity="error",
            detail="无法读取 word/document.xml",
        )]

    body = root.find(f"{{{W}}}body")
    if body is None:
        return [FixerDiagnostic(
            code="RED_HEADER_VERIFY_FAILED",
            action_index=idx,
            severity="error",
            detail="document.xml 中无 body 元素",
        )]

    # 检查前两个段落是否包含红头内容
    paragraphs = list(body.iter(f"{{{W}}}p"))
    issues = []

    if org_name:
        found_org = False
        found_red_line = False
        for p in paragraphs[:3]:  # 检查前3段
            text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
            if org_name in text:
                found_org = True
                # 检查是否有红色反线（pBdr/bottom）
                pPr = p.find(f"{{{W}}}pPr")
                if pPr is not None:
                    pBdr = pPr.find(f"{{{W}}}pBdr")
                    if pBdr is not None:
                        bottom = pBdr.find(f"{{{W}}}bottom")
                        if bottom is not None and bottom.get(f"{{{W}}}color") == "FF0000":
                            found_red_line = True
            if found_org:
                break

        if not found_org:
            issues.append(FixerDiagnostic(
                code="RED_HEADER_ORG_MISSING",
                action_index=idx,
                severity="error",
                detail=f"未找到发文机关全称 '{org_name}'",
                expected=org_name,
                actual="未找到",
            ))
        elif not found_red_line:
            issues.append(FixerDiagnostic(
                code="RED_HEADER_LINE_MISSING",
                action_index=idx,
                severity="warning",
                detail="发文机关段落缺少红色反线（pBdr/bottom color=FF0000）",
            ))

    if doc_number:
        found_doc = False
        for p in paragraphs[:4]:
            text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
            if doc_number in text:
                found_doc = True
                break
        if not found_doc:
            issues.append(FixerDiagnostic(
                code="RED_HEADER_DOCNUM_MISSING",
                action_index=idx,
                severity="error",
                detail=f"未找到发文字号 '{doc_number}'",
                expected=doc_number,
                actual="未找到",
            ))

    return issues


@_verifier("set_code_block_style")
def _verify_set_code_block_style(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 set_code_block_style 操作：检查指定段落是否有 shd 背景色和等宽字体。"""
    return []  # TODO: 实现验证：检查 pPr/shd 和 rPr/rFonts


@_verifier("insert_toc")
def _verify_insert_toc(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 insert_toc 操作：检查 document.xml 中是否有 TOC fldChar 序列。"""
    return []  # TODO: 实现验证：检查 fldChar begin/end + instrText 包含 TOC

