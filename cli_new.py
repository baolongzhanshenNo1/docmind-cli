"""DocMind CLI — 专治 Word 排版疑难杂症。

用法:
  docmind format <论文.docx> --spec <规范.docx>  对论文按规范模板排版
  docmind enforce <文档.docx>                     奇偶页强制（空白页、页眉重整）
  docmind analyze <规范.docx>                     提取规范模板的格式规则 → YAML
"""

from __future__ import annotations

import sys
import re
import shutil
from pathlib import Path

import click
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text
from rich import box

# ---------------------------------------------------------------------------
# 确保项目模块可导入
# ---------------------------------------------------------------------------
_project = Path(__file__).resolve().parent
if str(_project) not in sys.path:
    sys.path.insert(0, str(_project))

console = Console()

BRAND = Text("DocMind — 专治 Word 排版疑难杂症", style="bold cyan")

# ---------------------------------------------------------------------------
# 自动检测 LibreOffice
# ---------------------------------------------------------------------------
_LO_CANDIDATES = [
    Path(r"D:\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    Path(r"C:\LibreOffice\program\soffice.exe"),
]


def _find_libreoffice() -> Path | None:
    """自动检测 LibreOffice。也试 PATH 里的 soffice。"""
    lo = shutil.which("soffice")
    if lo:
        return Path(lo)
    for p in _LO_CANDIDATES:
        if p.exists():
            return p
    return None


# ═══════════════════════════════════════════════
# CLI 群组
# ═══════════════════════════════════════════════

@click.group()
@click.version_option(version="2.0.0", prog_name="docmind", message="DocMind %(version)s")
def cli():
    """DocMind — 专治 Word 排版疑难杂症。

    本地离线、深度 OOXML、AI Agent 实时协作。
    """


# ═══════════════════════════════════════════════
# format
# ═══════════════════════════════════════════════

@cli.command("format")
@click.argument("target", type=click.Path(exists=True, dir_okay=False))
@click.option("--spec", "-s", required=True, type=click.Path(exists=True, dir_okay=False),
              help="规范模板 .docx（如学校论文撰写规范）")
@click.option("--output", "-o", default=None,
              help="输出 .docx 路径（默认: <target>_formatted.docx）")
@click.option("--print/--no-print", default=False,
              help="是否同时生成打印版（奇偶页强制）")
def format_cmd(target, spec, output, print):
    """对论文按规范模板排版（格式修正，不改文字内容）。

    TARGET: 待排版的 .docx 文件
    """
    target_path = Path(target)
    spec_path = Path(spec)

    if output is None:
        output = str(target_path.parent / f"{target_path.stem}_formatted.docx")
    output_path = Path(output)

    # 品牌头
    console.print()
    console.print(Panel(BRAND, border_style="cyan", padding=(1, 2)))
    console.print()

    with console.status("[cyan]加载排版引擎…", spinner="dots"):
        from docmind.tools import get_tool
        tool = get_tool("thesis")

    console.print(f"  [dim]目标文档[/dim] → {target_path.name}")
    console.print(f"  [dim]规范模板[/dim] → {spec_path.name}")

    # 执行排版
    with console.status(f"[cyan]正在排版: {target_path.name}…", spinner="dots"):
        result = tool.run(
            spec_docx=str(spec_path),
            target_docx=str(target_path),
            output_docx=str(output_path),
            print_mode=print,
        )

    # Step 3: 结果
    console.print()
    if result.success:
        table = Table(show_header=False, box=box.SIMPLE, padding=(0, 2))
        table.add_column(style="dim")
        table.add_column(style="bold green")
        table.add_row("📄 修正项", str(result.metadata.get("fixed_count", 0)))
        table.add_row("🔄 残留  ", str(result.metadata.get("remaining", 0)))
        table.add_row("⚠  诊断  ", str(len(result.issues or [])))
        table.add_row("📦 输出  ", str(output_path))
        console.print(table)

        if result.issues:
            console.print("\n  [yellow]⚠  残留问题:[/yellow]")
            for issue in result.issues:
                code = issue.get("code", "?")
                detail = issue.get("detail", "")
                console.print(f"    [yellow]{code}[/yellow] {detail}")

        if result.logs:
            console.print("\n  [dim]操作日志 (最后 6 条):[/dim]")
            for log in result.logs[-6:]:
                c = "green" if log.startswith("[OK") else ("red" if log.startswith("[FAIL") else "dim")
                console.print(f"    [{c}]{log}[/{c}]")
    else:
        console.print("  [red]✗ 排版失败[/red]")
        if result.logs:
            for log in result.logs[-10:]:
                if "[FAIL" in log or "[ERR" in log:
                    console.print(f"    [red]{log}[/red]")
        raise SystemExit(1)

    console.print()
    console.print("[green]✅ 排版完成[/green]")


# ═══════════════════════════════════════════════
# enforce
# ═══════════════════════════════════════════════

@cli.command("enforce")
@click.argument("docx", type=click.Path(exists=True, dir_okay=False))
@click.option("--output", "-o", default=None, help="输出 .docx（默认覆盖原文件）")
@click.option("--libreoffice", default=None, help="soffice.exe 路径（自动检测）")
@click.option("--sections", default=None, help="需强制奇页的章节, 逗号分隔 (默认:论文标准章节)")
def enforce_cmd(docx, output, libreoffice, sections):
    """对已有 docx 执行奇偶页强制（空白页插入、页眉重整）。"""
    input_path = Path(docx)
    output_path = Path(output) if output else input_path

    lo_path = _find_libreoffice() if not libreoffice else Path(libreoffice)
    if not lo_path or not lo_path.exists():
        console.print("[red]✗ 未找到 LibreOffice[/red]")
        console.print("  [dim]请通过 --libreoffice 指定 soffice.exe 路径[/dim]")
        raise SystemExit(1)

    console.print()
    console.print(Panel(BRAND, border_style="cyan", padding=(1, 2)))
    console.print()

    # 默认章节（论文标准结构）
    if sections:
        odd_sections = [s.strip() for s in sections.split(",") if s.strip()]
    else:
        odd_sections = [
            "封面", "郑重声明", "摘要", "ABSTRACT", "目录",
            "1 绪论", "参考文献", "结论", "致谢", "附录",
        ]

    with console.status(f"[cyan]奇偶页强制: {input_path.name}…", spinner="dots"):
        from enforce.config import EnforceConfig
        from enforce.odd_pages import enforce_odd_pages

        config = EnforceConfig(
            libreoffice_path=lo_path,
            docx_input=input_path,
            docx_output=output_path,
            odd_page_sections=odd_sections,
        )
        result = enforce_odd_pages(config)

    console.print(f"  [green]✅ 奇偶页强制完成[/green] → {output_path}")


# ═══════════════════════════════════════════════
# analyze
# ═══════════════════════════════════════════════

@cli.command("analyze")
@click.argument("spec_docx", type=click.Path(exists=True, dir_okay=False))
@click.option("--output", "-o", default=None, help="输出 YAML 路径（默认: <spec>_template.yaml）")
def analyze_cmd(spec_docx, output):
    """提取规范模板的格式规则（字体/边距/页眉/页码）→ YAML。

    生成的结构化 YAML 可直接作为 `format` 命令的 --spec 参数。
    """
    spec_path = Path(spec_docx)
    if output is None:
        output = str(spec_path.parent / f"{spec_path.stem}_template.yaml")
    output_path = Path(output)

    console.print()
    console.print(Panel(BRAND, border_style="cyan", padding=(1, 2)))
    console.print()

    with console.status(f"[cyan]分析规范模板: {spec_path.name}…", spinner="dots"):
        # Legacy template_analyzer（来自 _archive，功能完好）
        sys.path.insert(0, str(_project.parent / "_archive" / "docmind"))
        from generator.template_analyzer import analyze_template

        import yaml

        data = analyze_template(str(spec_path))

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(yaml.dump(data, allow_unicode=True, default_flow_style=False),
                               encoding="utf-8")

    # 摘要
    font_count = len(data.get("fonts", {})) if isinstance(data, dict) else 0
    page = data.get("page", {}) if isinstance(data, dict) else {}

    console.print(f"  [green]✅ 分析完成[/green] → {output_path}")
    console.print(f"  [dim]  字体规则: {font_count} 条[/dim]")
    if page:
        margins = page.get("margins", {})
        if margins:
            console.print(f"  [dim]  页面设置: A4, 上{margins.get('top','?')} 下{margins.get('bottom','?')} 左{margins.get('left','?')} 右{margins.get('right','?')}[/dim]")


def _spec_to_dict(spec) -> dict:
    """SpecReader 结果 → 纯 dict（保留以防旧版调用，已不用）。"""
    return {"fonts": {}, "page": {}}


# ═══════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════

if __name__ == "__main__":
    cli()
