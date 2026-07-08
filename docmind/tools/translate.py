"""文档翻译工具 — 中英文互译 + 格式保持 + 字体分离。

核心流程：
1. OOXML 提取每段文本（document.xml body，含表格单元格段落）
2. DeepSeek 批量翻译（术语库约束）
3. 回填到 OOXML（target_only：首个 w:t 写译文、其余清空；中英字体分离）

限制（v1）：仅翻译 document.xml 主体；页眉/页脚(headerN.xml)、图片/公式内文字暂不处理。
"""

import os
import re
import json
import zipfile
from pathlib import Path
from typing import Optional, Literal

from lxml import etree

from docmind.tools.base import DocMindTool, ToolResult

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "http://www.w3.org/XML/1998/namespace"

# ── 翻译模式 ──
TRANSLATION_MODES = {
    "bilingual": "双语对照（原文 + 译文并列）",
    "target_only": "仅译文（替换原文）",
    "paragraph_pair": "逐段对照（原文一段 + 译文一段）",
}

_ENV_PATH = r"D:\Microsoft VS Code Projects\office\server\.env"


# ============================================================
# LLM 翻译
# ============================================================

def _read_llm_config(llm_api_key: str = "") -> tuple[str, str, str]:
    """读取 DeepSeek 配置：优先入参 key，否则从 server/.env 读取。"""
    key = llm_api_key or os.getenv("DEEPSEEK_API_KEY", "")
    base = os.getenv("DEEPSEEK_BASE_URL", "")
    model = os.getenv("DEEPSEEK_MODEL", "")
    if not key or not base or not model:
        try:
            for line in Path(_ENV_PATH).read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if line.startswith("DEEPSEEK_API_KEY=") and not key:
                    key = line.split("=", 1)[1].strip().strip('"')
                elif line.startswith("DEEPSEEK_BASE_URL=") and not base:
                    base = line.split("=", 1)[1].strip().strip('"')
                elif line.startswith("DEEPSEEK_MODEL=") and not model:
                    model = line.split("=", 1)[1].strip().strip('"')
        except Exception:
            pass
    return key, (base or "https://api.deepseek.com").rstrip("/"), (model or "deepseek-v4-flash")


def _llm_call(messages: list, key: str, base: str, model: str,
              max_tokens: int = 8192, timeout: int = 120) -> str:
    import urllib.request
    body = {"model": model, "messages": messages,
            "temperature": 0, "max_tokens": max_tokens}
    req = urllib.request.Request(
        base + "/chat/completions",
        data=json.dumps(body).encode("utf-8"),
        headers={"Authorization": "Bearer " + key,
                 "Content-Type": "application/json"},
    )
    r = json.load(urllib.request.urlopen(req, timeout=timeout))
    return (r["choices"][0]["message"].get("content") or "").strip()


def _parse_numbered(text: str, n: int) -> dict:
    """解析 '[i] 译文' 分段（贪婪：一个 [n] 到下一个 [m] 之间都是该段）。"""
    result: dict[int, str] = {}
    parts = re.split(r"\[(\d+)\]", text)
    # parts = [pre, num, seg, num, seg, ...]
    seq = parts[1:]
    for i in range(0, len(seq) - 1, 2):
        try:
            idx = int(seq[i])
            result[idx] = seq[i + 1].strip()
        except (ValueError, IndexError):
            continue
    return result


def _translate_all(texts: list[str], target_lang: str, key: str, base: str,
                   model: str, glossary: dict, logs: list) -> list[str]:
    """批量翻译；失败/缺失的段落回退原文。"""
    lang = "英文" if target_lang == "en" else "中文"
    result = [""] * len(texts)
    BATCH = 20
    gloss = ""
    if glossary:
        pairs = "；".join(f"{k}→{v}" for k, v in list(glossary.items())[:60])
        gloss = f"\n【术语对照，必须严格遵守】{pairs}"
    sys = (f"你是专业文档翻译。把每个带 [编号] 前缀的段落翻译成{lang}，"
           f"严格保持相同的 [编号] 前缀，每个编号对应一段，"
           f"绝不合并/拆分/新增/删除/重排段落，只输出译文本身，不要任何解释。{gloss}")
    for start in range(0, len(texts), BATCH):
        batch = texts[start:start + BATCH]
        numbered = "\n".join(f"[{i}] {t}" for i, t in enumerate(batch))
        try:
            out = _llm_call(
                [{"role": "system", "content": sys},
                 {"role": "user", "content": numbered}],
                key, base, model,
            )
            parsed = _parse_numbered(out, len(batch))
            for i in range(len(batch)):
                result[start + i] = parsed.get(i) or texts[start + i]
            logs.append(f"  批次 {start}-{start + len(batch) - 1}: 译回 {len(parsed)}/{len(batch)}")
        except Exception as e:
            for i in range(len(batch)):
                result[start + i] = texts[start + i]
            logs.append(f"  [WARN] 批次 {start} 翻译失败，保留原文: {e}")
    return result


# ============================================================
# OOXML 回填
# ============================================================

def _set_run_fonts(t_elem) -> None:
    """在 w:t 所属 w:r 上设置中英字体分离（英=Times New Roman，中=宋体）。"""
    r = t_elem.getparent()
    if r is None or r.tag != f"{{{W}}}r":
        return
    rpr = r.find(f"{{{W}}}rPr")
    if rpr is None:
        rpr = etree.Element(f"{{{W}}}rPr")
        r.insert(0, rpr)  # rPr 必须是 w:r 的第一个子元素
    rfonts = rpr.find(f"{{{W}}}rFonts")
    if rfonts is None:
        rfonts = etree.SubElement(rpr, f"{{{W}}}rFonts")
    rfonts.set(f"{{{W}}}ascii", "Times New Roman")
    rfonts.set(f"{{{W}}}hAnsi", "Times New Roman")
    rfonts.set(f"{{{W}}}eastAsia", "宋体")


def _write_docx(source: Path, out: Path, doc_xml_bytes: bytes) -> None:
    """复制原 docx 到 out，仅替换 word/document.xml。"""
    with zipfile.ZipFile(source, "r") as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = doc_xml_bytes
            zout.writestr(item, data)


class TranslateTool(DocMindTool):
    """文档翻译工具。"""

    @property
    def tool_name(self) -> str:
        return "translate"

    @property
    def description(self) -> str:
        return "文档翻译：中英互译、格式保持、术语库约束、字体分离"

    def run(
        self,
        source_docx: str | Path,
        output_docx: Optional[str | Path] = None,
        target_lang: Literal["en", "zh"] = "en",
        mode: Literal["bilingual", "target_only", "paragraph_pair"] = "target_only",
        glossary: Optional[str | Path] = None,
        llm_provider: str = "",
        llm_api_key: str = "",
    ) -> ToolResult:
        source = Path(source_docx)
        if not source.exists():
            return ToolResult(success=False, tool_name=self.tool_name,
                              logs=[f"文件不存在: {source}"])

        suffix = "_en" if target_lang == "en" else "_zh"
        out = Path(output_docx) if output_docx else source.parent / f"{source.stem}{suffix}.docx"

        logs: list[str] = []
        try:
            logs.append(f"[translate] {source.name} → "
                        f"{'英文' if target_lang == 'en' else '中文'}, 模式={mode}")

            key, base, model = _read_llm_config(llm_api_key)
            if not key:
                return ToolResult(success=False, tool_name=self.tool_name,
                                  logs=logs + ["未配置 DEEPSEEK_API_KEY"])

            # 术语库
            glossary_terms: dict = {}
            if glossary and Path(glossary).exists():
                try:
                    glossary_terms = json.loads(Path(glossary).read_text(encoding="utf-8"))
                    logs.append(f"  术语库: {len(glossary_terms)} 条")
                except Exception:
                    logs.append("  [WARN] 术语库加载失败，跳过")

            # ── Step 1: 提取段落（含其 w:t 元素引用，便于回填） ──
            with zipfile.ZipFile(source, "r") as z:
                root = etree.fromstring(z.read("word/document.xml"))

            para_items = []  # [(w:p, full_text, [w:t...])]
            for p in root.iter(f"{{{W}}}p"):
                ts = p.findall(f".//{{{W}}}t")
                text = "".join(t.text or "" for t in ts)
                if text.strip():
                    para_items.append((p, text, ts))
            logs.append(f"  提取 {len(para_items)} 个文本段落")
            if not para_items:
                return ToolResult(success=False, tool_name=self.tool_name,
                                  logs=logs + ["文档无可翻译文本"])

            # ── Step 2: 批量翻译 ──
            texts = [it[1] for it in para_items]
            translations = _translate_all(texts, target_lang, key, base, model,
                                           glossary_terms, logs)

            # ── Step 3: 回填（target_only：首 w:t 写译文，其余清空）+ 字体分离 ──
            filled = 0
            for (p, orig, ts), trans in zip(para_items, translations):
                if not trans or not ts:
                    continue
                ts[0].text = trans
                ts[0].set(f"{{{XML_SPACE}}}space", "preserve")
                for t in ts[1:]:
                    t.text = ""
                for t in ts:
                    _set_run_fonts(t)
                filled += 1

            doc_bytes = etree.tostring(root, xml_declaration=True,
                                       encoding="UTF-8", standalone=True)
            _write_docx(source, out, doc_bytes)
            logs.append(f"  回填 {filled} 段，字体分离(中=宋体/英=Times New Roman) → {out.name}")

            return ToolResult(
                success=True,
                tool_name=self.tool_name,
                output_path=out,
                logs=logs[-6:],
                metadata={
                    "target_lang": target_lang,
                    "mode": mode,
                    "paragraphs": len(para_items),
                    "translated": filled,
                    "glossary_terms": len(glossary_terms),
                    "status": "ok",
                },
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            return ToolResult(success=False, tool_name=self.tool_name,
                              logs=logs + [f"异常: {e}"])
