"""技术文档工具 — 使用可组合 Steps + Writer。"""

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
    PageNumberStep,
    TECH_H1,
)
from docmind.steps.tech_doc import CodeBlockStep, TocStep


DOC_TYPE_DEFAULTS = {
    "api": {"h1_size": 16, "body_size": 10.5, "code_size": 9},
    "spec": {"h1_size": 16, "body_size": 10.5, "code_size": 9},
    "readme": {"h1_size": 18, "body_size": 10.5, "code_size": 9},
    "changelog": {"h1_size": 14, "body_size": 10.5, "code_size": 9},
}


class TechDocTool(DocMindTool):
    """技术文档排版工具。"""

    @property
    def tool_name(self) -> str:
        return "tech_doc"

    @property
    def description(self) -> str:
        return "技术文档排版：代码块高亮、API 文档模板、Markdown 转换、自动目录"

    def run(
        self,
        target_docx: str | Path,
        output_docx: Optional[str | Path] = None,
        doc_type: Literal["api", "spec", "readme", "changelog"] = "readme",
    ) -> ToolResult:
        target = Path(target_docx)
        if not target.exists():
            return ToolResult(success=False, tool_name=self.tool_name,
                            logs=[f"文件不存在: {target}"])

        if output_docx is None:
            out = target.parent / f"{target.stem}_formatted.docx"
        else:
            out = Path(output_docx)

        defaults = DOC_TYPE_DEFAULTS.get(doc_type, DOC_TYPE_DEFAULTS["readme"])

        spec = {
            "h1": {
                "font_east": "黑体",
                "font_ascii": "Arial",
                "font_size_pt": defaults["h1_size"],
                "bold": True,
            },
            "body": {
                "font_east": "宋体",
                "font_ascii": "Consolas",
                "font_size_pt": defaults["body_size"],
            },
            "footer_format": {
                "page_number_format": "{PAGE}",
                "alignment": "center",
            },
        }

        try:
            shutil.copy2(target, out)

            steps = [
                StyleStep(),
                SectionBreakStep(),
                HeaderFontStep(),
                HeaderTextStep(
                    h1_pattern=TECH_H1,
                    front_skip_headings=set(),
                    normalizer=lambda t: t.strip(),
                ),
                FooterLinkStep(),
                PageNumberStep(),
                CodeBlockStep(),
                TocStep(),
            ]
            reconciler = Reconciler(steps)
            fix_plan = reconciler.reconcile(out, spec)

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
