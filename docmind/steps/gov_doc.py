"""公文专属排版步骤。"""

from pathlib import Path
from docmind.steps.base import PipelineStep


class CenteredPageNumStep(PipelineStep):
    """公文页码：居中 "— X —" 格式。"""

    def run(self, target: Path, spec: dict) -> list[dict]:
        footer_fmt = spec.get("footer_format", {}) or {}
        fmt = footer_fmt.get("page_number_format", "— {PAGE} —")
        body = spec.get("body", {})
        font = body.get("font_ascii", "仿宋")

        import zipfile
        with zipfile.ZipFile(target, "r") as z:
            ftrs = [n for n in z.namelist() if "footer" in n and n.endswith(".xml")]

        plan = []
        for ftr in ftrs:
            plan.append({
                "action": "add_page_number",
                "params": {
                    "footer_path": ftr,
                    "format": fmt,
                    "alignment": "center",
                    "font_eastAsia": font,
                    "font_ascii": font,
                    "font_size": "16",
                },
            })
        return plan


class RedHeaderStep(PipelineStep):
    """公文红头渲染：发文机关全称 + 红色反线 + 发文字号。

    依据 GB/T 9704《党政机关公文格式》：
    - 发文机关全称：居中、红色、初号/小初（42pt）
    - 红色反线：位于机关名称下方
    - 发文字号：居中、红色、三号（16pt）
    - 红头区域位于文档正文起始处（第一个 section 顶部）

    从 spec.red_header 读取：
        org_name: str   — 发文机关全称（如 "XX市人民政府"）
        doc_number: str — 发文字号（如 "X政发〔2024〕1号"）
    """

    def run(self, target: Path, spec: dict) -> list[dict]:
        red_header = spec.get("red_header", {}) or {}
        org_name = red_header.get("org_name", "")
        doc_number = red_header.get("doc_number", "")

        if not org_name and not doc_number:
            return []

        body_rules = spec.get("body", {})
        font_east = body_rules.get("font_east", "方正小标宋简体")

        return [{
            "action": "insert_red_header",
            "params": {
                "org_name": org_name,
                "doc_number": doc_number,
                "font_eastAsia": font_east,
                "org_font_size": "84",   # 42pt → 84 half-points（初号）
                "doc_font_size": "32",   # 16pt → 32 half-points（三号）
            },
        }]
