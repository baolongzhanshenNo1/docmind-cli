"""技术文档专属排版步骤。"""

import re
import zipfile
from pathlib import Path
from lxml import etree

from docmind.steps.base import PipelineStep, W


class CodeBlockStep(PipelineStep):
    """检测代码块段落，添加背景色和等宽字体。

    检测规则：
    1. 段落文本以 4 个空格或 Tab 开头
    2. 段落 run 中的 rFonts 使用了等宽字体（Consolas, Courier, Monaco 等）
    """

    MONOSPACE_FONTS = {
        "consolas", "courier", "courier new", "monaco",
        "source code pro", "fira code", "jetbrains mono",
        "menlo", "dejavu sans mono", "liberation mono",
        "monospace",
    }

    def run(self, target: Path, spec: dict) -> list[dict]:
        with zipfile.ZipFile(target, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))

        body = root.find(f"{{{W}}}body")
        if body is None:
            return []

        code_block_indices = []
        p_index = 0
        for child in body:
            if child.tag != f"{{{W}}}p":
                continue

            if self._is_code_block(child):
                code_block_indices.append(p_index)
            p_index += 1

        if not code_block_indices:
            return []

        return [{
            "action": "set_code_block_style",
            "params": {"paragraph_indices": code_block_indices},
        }]

    def _is_code_block(self, para) -> bool:
        """判断段落是否为代码块。"""
        # 提取段落纯文本
        text = "".join(t.text or "" for t in para.iter(f"{{{W}}}t"))

        # 规则1：以 4 个空格或 Tab 开头
        if text.startswith("    ") or text.startswith("\t"):
            return True

        # 规则2：任意 run 的 rFonts 使用了等宽字体
        for run in para.iter(f"{{{W}}}r"):
            rPr = run.find(f"{{{W}}}rPr")
            if rPr is not None:
                rf = rPr.find(f"{{{W}}}rFonts")
                if rf is not None:
                    ascii_font = (rf.get(f"{{{W}}}ascii") or "").lower()
                    han_font = (rf.get(f"{{{W}}}hAnsi") or "").lower()
                    if ascii_font in self.MONOSPACE_FONTS or han_font in self.MONOSPACE_FONTS:
                        return True

        return False


class TocStep(PipelineStep):
    """自动生成目录 TOC 域。

    策略：
    1. 查找"目录"或"目  录"标题段落
    2. 在"目录"标题后的第一个空段落处插入 TOC 域
    3. 如果找不到"目录"标题，在文档开头插入 TOC
    """

    def run(self, target: Path, spec: dict) -> list[dict]:
        with zipfile.ZipFile(target, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))

        body = root.find(f"{{{W}}}body")
        if body is None:
            return []

        # 收集 p 段落及其索引
        p_elements = []
        for child in body:
            if child.tag == f"{{{W}}}p":
                p_elements.append(child)

        # 查找"目录"标题
        toc_heading_idx = -1
        for idx, p in enumerate(p_elements):
            text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t")).strip()
            normalized = re.sub(r"\s+", "", text)
            if normalized in ("目录", "目錄"):
                toc_heading_idx = idx
                break

        # 在"目录"标题后查找第一个空段落
        if toc_heading_idx >= 0:
            for offset in range(1, min(len(p_elements) - toc_heading_idx, 5)):
                candidate_idx = toc_heading_idx + offset
                candidate = p_elements[candidate_idx]
                text = "".join(t.text or "" for t in candidate.iter(f"{{{W}}}t")).strip()
                if not text:
                    # 找到空段，插入 TOC 到它的位置（替代它）
                    return [{
                        "action": "insert_toc",
                        "params": {"after_paragraph_index": candidate_idx - 1},
                    }]

            # 没有空段落，插在"目录"标题后
            return [{
                "action": "insert_toc",
                "params": {"after_paragraph_index": toc_heading_idx},
            }]

        # 找不到"目录"标题，插入到文档开头
        return [{
            "action": "insert_toc",
            "params": {"after_paragraph_index": None},
        }]
