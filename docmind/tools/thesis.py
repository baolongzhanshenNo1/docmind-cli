"""论文排版工具 — 使用可组合 Steps。"""

from pathlib import Path
from typing import Optional

from docmind.tools.base import DocMindTool, ToolResult
from docmind.steps.base import (
    Reconciler,
    StyleStep,
    SectionBreakStep,
    PageNumberStep,
    HeaderFontStep,
    HeaderTextStep,
    FooterLinkStep,
    THESIS_H1,
)
from docmind.steps.thesis import RomanArabicStep


THESIS_STEPS = [
    StyleStep(),
    SectionBreakStep(),
    HeaderFontStep(),
    HeaderTextStep(
        h1_pattern=THESIS_H1,
        front_skip_headings={"封面", "郑重声明", "毕业设计"},
    ),
    FooterLinkStep(),
    PageNumberStep(),
    RomanArabicStep(),
]

_thesis_reconciler = Reconciler(THESIS_STEPS)


class ThesisTool(DocMindTool):
    """高校论文排版工具。"""

    @property
    def tool_name(self) -> str:
        return "thesis"

    @property
    def description(self) -> str:
        return "高校论文排版：页眉页码、奇偶页、空白页插入、样式规范化"

    def run(
        self,
        spec_docx: str | Path,
        target_docx: str | Path,
        output_docx: Optional[str | Path] = None,
        print_mode: bool = True,
        libreoffice_path: Optional[str] = None,
    ) -> ToolResult:
        spec = Path(spec_docx)
        target = Path(target_docx)

        if not spec.exists():
            return ToolResult(success=False, tool_name=self.tool_name,
                            logs=[f"规范模板不存在: {spec}"])
        if not target.exists():
            return ToolResult(success=False, tool_name=self.tool_name,
                            logs=[f"目标文档不存在: {target}"])

        if output_docx is None:
            out = target.parent / f"{target.stem}_formatted.docx"
        else:
            out = Path(output_docx)

        out_print = out.parent / f"{out.stem}_print.docx" if print_mode else None

        # ── 使用 Pipeline ──
        from pipeline.agent import DocMindAgent

        agent = DocMindAgent()
        try:
            result = agent.run(
                spec_path=spec,
                target_path=target,
                output_path=out,
                print_mode=print_mode,
                output_print=out_print,
                libreoffice_path=libreoffice_path,
            )
        except Exception as e:
            return ToolResult(
                success=False, tool_name=self.tool_name,
                logs=[f"Pipeline 异常: {e}"],
            )

        return ToolResult(
            success=result.remaining == 0,
            tool_name=self.tool_name,
            output_path=result.output_path,
            report_path=result.output_print,
            issues=[{"code": i.code, "detail": i.detail} for i in result.issues],
            logs=result.logs[-10:],
            metadata={
                "fixed_count": result.fixed_count,
                "remaining": result.remaining,
            },
        )
