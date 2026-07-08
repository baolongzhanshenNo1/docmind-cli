"""Pipeline Steps — 可组合的排版步骤。

每个 Step 是高内聚的纯函数：读文档 + spec → 生成 actions。
Writer 只执行 fix_plan，不知道 Step 来源。
"""

from __future__ import annotations

import re
import zipfile
from abc import ABC, abstractmethod
from pathlib import Path
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ── H1 正则（各领域可覆盖） ──
THESIS_H1 = (
    r'^(?:封面|郑\s*重\s*声\s*明|郑重声明|摘\s*要|ABSTRACT|'
    r'目\s*录|目录|\d+[\u4e00-\u9fff]|'
    r'结\s*论|结论|参考\s*文献|参考文献|致\s*谢|致谢|'
    r'附\s*录|附录)'
)

GOV_H1 = r'^(?:[一二三四五六七八九十]+、)'

TECH_H1 = r'^(?:#{1,3}\s|## |### |\d+\.\s|[A-Z][a-z]+)'


# ═══════════════════════════════════════════
# 接口
# ═══════════════════════════════════════════

class PipelineStep(ABC):
    """一个排版步骤。"""

    @abstractmethod
    def run(self, target: Path, spec: dict) -> list[dict]:
        """返回 action 列表，追加到 fix_plan。"""
        ...


class Reconciler:
    """按顺序执行 Steps，合并成 fix_plan。"""

    def __init__(self, steps: list[PipelineStep]):
        self.steps = steps

    def reconcile(self, target: Path, spec: dict) -> list[dict]:
        plan = []
        for step in self.steps:
            plan.extend(step.run(target, spec))
        return plan


# ═══════════════════════════════════════════
# Base Steps — 所有领域共享
# ═══════════════════════════════════════════

class StyleStep(PipelineStep):
    """从 spec 读取 H1/H2/Body 样式，生成 set_style actions。"""

    def run(self, target: Path, spec: dict) -> list[dict]:
        plan = []

        for level, key in [("h1", "1"), ("h2", "20"), ("h3", "30")]:
            rules = spec.get(level, {})
            if not rules:
                continue
            params = {"style_id": key}
            for field, param in [
                ("font_east", "font_eastAsia"),
                ("font_ascii", "font_ascii"),
                ("font_size_pt", "font_size_pt"),
                ("bold", "bold"),
                ("alignment", "alignment"),
            ]:
                if field in rules:
                    params[param] = rules[field]
            if len(params) > 1:
                plan.append({"action": "set_style", "params": params})

        body = spec.get("body", {})
        if body:
            params = {"style_id": spec.get("body_style_id", "a8")}
            for field, param in [
                ("font_east", "font_eastAsia"),
                ("font_ascii", "font_ascii"),
                ("font_size_pt", "font_size_pt"),
                ("alignment", "alignment"),
            ]:
                if field in body:
                    params[param] = body[field]
            if len(params) > 1:
                plan.append({"action": "set_style", "params": params})

        return plan


class SectionBreakStep(PipelineStep):
    """确保所有 sectPr type 为 nextPage。"""

    def run(self, target: Path, spec: dict) -> list[dict]:
        plan = []
        with zipfile.ZipFile(target, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        for i, sp in enumerate(root.iter(f"{{{W}}}sectPr")):
            type_el = sp.find(f"{{{W}}}type")
            current = type_el.get(f"{{{W}}}val") if type_el is not None else None
            if current != "nextPage":
                plan.append({
                    "action": "set_sectpr_type",
                    "params": {"section_index": i, "val": "nextPage"},
                })
        return plan


class PageNumberStep(PipelineStep):
    """添加页码 + footer reference + 页码格式。"""

    def run(self, target: Path, spec: dict) -> list[dict]:
        plan = []
        footer_fmt = spec.get("footer_format", {}) or {}
        fmt = footer_fmt.get("page_number_format", "{PAGE}")
        alignment = footer_fmt.get("alignment", "right")
        body = spec.get("body", {})
        font = body.get("font_ascii", "Times New Roman")

        with zipfile.ZipFile(target, "r") as z:
            ftrs = [n for n in z.namelist() if "footer" in n and n.endswith(".xml")]

        for ftr in ftrs:
            plan.append({
                "action": "add_page_number",
                "params": {
                    "footer_path": ftr,
                    "format": fmt,
                    "alignment": alignment,
                    "font_eastAsia": font,
                    "font_ascii": font,
                    "font_size": "18",
                },
            })

        # footer references
        with zipfile.ZipFile(target, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        for i, sp in enumerate(root.iter(f"{{{W}}}sectPr")):
            fref = sp.find(f"{{{W}}}footerReference")
            if fref is None:
                plan.append({
                    "action": "add_footer_reference",
                    "params": {"section_index": i},
                })

        return plan


class HeaderFontStep(PipelineStep):
    """设置所有 header XML 的字体。"""

    def run(self, target: Path, spec: dict) -> list[dict]:
        plan = []
        body = spec.get("body", {})
        fe = body.get("font_east", "宋体")
        fa = body.get("font_ascii", "Times New Roman")

        with zipfile.ZipFile(target, "r") as z:
            hdrs = [n for n in z.namelist() if "header" in n and n.endswith(".xml")]

        for hdr in hdrs:
            plan.append({
                "action": "set_header_font",
                "params": {
                    "header_path": hdr,
                    "font_ascii": fa,
                    "font_eastAsia": fe,
                    "font_size": "18",
                },
            })
        return plan


class HeaderTextStep(PipelineStep):
    """为每节创建独立 header + 设置页眉文字。

    参数:
        h1_pattern: H1 标题正则
        front_skip_headings: 不设页眉的标题集合
        normalizer: 页眉文字标准化函数 (默认: 去空格 + 去编号前缀)
    """

    def __init__(
        self,
        h1_pattern: str = THESIS_H1,
        front_skip_headings: set = None,
        normalizer=None,
    ):
        self.h1_pat = re.compile(h1_pattern)
        self.front_skip = front_skip_headings or set()
        self.normalizer = normalizer or _default_normalizer

    def run(self, target: Path, spec: dict) -> list[dict]:
        from pipeline.agent import DocMindAgent

        body_rules = spec.get("body", {})
        font_east = body_rules.get("font_east", "宋体")
        font_ascii = body_rules.get("font_ascii", "Times New Roman")

        with zipfile.ZipFile(target, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        body = root.find(f"{{{W}}}body")

        # 收集 H1
        h1_positions = []
        for i, child in enumerate(body):
            para = None
            if child.tag == f"{{{W}}}p":
                para = child
            elif child.tag == f"{{{W}}}sdt":
                sc = child.find(f"{{{W}}}sdtContent")
                if sc is not None:
                    para = sc.find(f"{{{W}}}p")
            if para is not None:
                txt = ''.join(t.text or '' for t in para.iter(f"{{{W}}}t")).strip()
                if txt and self.h1_pat.match(re.sub(r'\s+', '', txt)):
                    h1_positions.append((i, txt))

        # 收集 sectPr
        sects_in_order = []
        for child_idx, child in enumerate(body):
            sp = None
            if child.tag == f"{{{W}}}sectPr":
                sp = child
            elif child.tag == f"{{{W}}}p":
                pp = child.find(f"{{{W}}}pPr")
                if pp is not None:
                    sp = pp.find(f"{{{W}}}sectPr")
            if sp is not None:
                sects_in_order.append((child_idx, sp))

        # 为每个 sectPr 匹配 governing H1
        section_headers = []
        for si, (child_idx, _sp) in enumerate(sects_in_order):
            heading_text = None
            for h1_idx, h1_txt in reversed(h1_positions):
                if h1_idx < child_idx:
                    heading_text = h1_txt
                    break
            if heading_text:
                norm = re.sub(r'\s+', '', heading_text)
                if norm in self.front_skip:
                    continue
                header_text = self.normalizer(heading_text)
                section_headers.append({
                    'sectpr_index': si,
                    'default_text': header_text,
                    'even_text': None,
                })

        if not section_headers:
            return []

        return [{
            "action": "create_section_headers",
            "params": {
                "section_headers": section_headers,
                "font_eastAsia": font_east,
                "font_ascii": font_ascii,
                "font_size": "18",
            },
        }]


class FooterLinkStep(PipelineStep):
    """清除封面(节0) + 郑重声明(节1)的页眉——前置两节不应有页眉。

    结构假设：thesis 文档首节=封面、次节=郑重声明（与既有约定一致）。
    机制：这两节是文档最前的节，清掉 headerReference 后：
      - 节0(封面) 是第一节，无前节可继承 → 空白页眉
      - 节1(郑重声明) 继承已清空的节0 → 空白页眉
    （历史 bug：只清了节1，节0 保留陈旧共享页眉[摘要]，节1 反而继承了它。）
    """

    def run(self, target: Path, spec: dict) -> list[dict]:
        return [
            {"action": "clear_section_headers", "params": {"section_index": 0}},  # 封面
            {"action": "clear_section_headers", "params": {"section_index": 1}},  # 郑重声明
        ]


# ═══════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════

def _default_normalizer(text: str) -> str:
    """论文默认标准化：去空格 + 去编号前缀。"""
    text = re.sub(r'\s+', '', text)
    m = re.match(r'^\d+([\u4e00-\u9fff].*)', text)
    if m:
        text = m.group(1)
    return text
