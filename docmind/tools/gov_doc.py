"""公文排版工具 — 使用可组合 Steps + Writer。"""

import shutil
from pathlib import Path
from typing import Optional, Literal

from docmind.tools.base import DocMindTool, ToolResult
from docmind.steps.base import (
    Reconciler,
    StyleStep,
    SectionBreakStep,
    HeaderFontStep,
    HeaderTextStep,
    FooterLinkStep,
    GOV_H1,
)
from docmind.steps.gov_doc import CenteredPageNumStep, RedHeaderStep


DOC_TYPE_DEFAULTS: dict[str, dict] = {
    "notice": {"h1_font_size": 22, "body_font_size": 16},
    "report": {"h1_font_size": 22, "body_font_size": 16},
    "approval": {"h1_font_size": 22, "body_font_size": 16},
    "letter": {"h1_font_size": 16, "body_font_size": 16},
    "minutes": {"h1_font_size": 22, "body_font_size": 16},
}


class GovDocTool(DocMindTool):
    """政府公文排版工具。"""

    @property
    def tool_name(self) -> str:
        return "gov_doc"

    @property
    def description(self) -> str:
        return "政府公文排版：红头文件、层级编号（一、/1.）、密级标注、印章落款"

    def run(
        self,
        target_docx: str | Path,
        output_docx: Optional[str | Path] = None,
        doc_type: Literal["notice", "report", "approval", "letter", "minutes"] = "notice",
        org_name: Optional[str] = None,
        doc_number: Optional[str] = None,
        libreoffice_path: Optional[str] = None,
    ) -> ToolResult:
        target = Path(target_docx)
        if not target.exists():
            return ToolResult(success=False, tool_name=self.tool_name,
                            logs=[f"文件不存在: {target}"])

        if output_docx is None:
            out = target.parent / f"{target.stem}_formatted.docx"
        else:
            out = Path(output_docx)

        defaults = DOC_TYPE_DEFAULTS.get(doc_type, DOC_TYPE_DEFAULTS["notice"])

        spec = {
            "h1": {
                "font_east": "方正小标宋简体",
                "font_size_pt": defaults["h1_font_size"],
                "alignment": "center",
            },
            "body": {
                "font_east": "仿宋",
                "font_ascii": "Times New Roman",
                "font_size_pt": defaults["body_font_size"],
            },
            "footer_format": {
                "page_number_format": "— {PAGE} —",
                "alignment": "center",
            },
            "red_header": {
                "org_name": org_name or "",
                "doc_number": doc_number or "",
            },
        }

        try:
            shutil.copy2(target, out)

            # 构建 Reconciler
            steps = [
                StyleStep(),
                SectionBreakStep(),
                HeaderFontStep(),
                HeaderTextStep(
                    h1_pattern=GOV_H1,
                    front_skip_headings=set(),
                    normalizer=lambda t: re.sub(r'\s+', '', t),
                ),
                FooterLinkStep(),
                CenteredPageNumStep(),
                RedHeaderStep(),
            ]
            reconciler = Reconciler(steps)
            fix_plan = reconciler.reconcile(out, spec)

            # 执行
            from pipeline.writer_v2 import apply_fixes
            writer_logs = apply_fixes(fix_plan, out)

            from pipeline.fixer_v2 import diagnose, summary
            issues = diagnose(fix_plan, out)
            stats = summary(issues)

            return ToolResult(
                success=stats["total"] == 0,
                tool_name=self.tool_name,
                output_path=out,
                issues=[{"code": i.code, "detail": i.detail} for i in issues],
                logs=writer_logs[-5:],
                metadata={
                    "doc_type": doc_type,
                    "fixed_count": sum(1 for l in writer_logs if "[OK" in l),
                    "remaining": stats["total"],
                },
            )
        except Exception as e:
            return ToolResult(success=False, tool_name=self.tool_name,
                            logs=[f"异常: {e}"])


import re  # noqa: E402
