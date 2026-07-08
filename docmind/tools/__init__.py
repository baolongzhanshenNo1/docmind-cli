"""DocMind 工具注册。"""

from __future__ import annotations
from docmind.tools.base import DocMindTool

_registry: dict[str, DocMindTool] = {}

def register(tool: DocMindTool) -> None:
    _registry[tool.tool_name] = tool

def get_tool(name: str) -> DocMindTool | None:
    return _registry.get(name)

def list_tools() -> list[DocMindTool]:
    return list(_registry.values())

# ── 注册 5 个内置工具 ──
from docmind.tools.thesis import ThesisTool         # noqa: E402
from docmind.tools.contract import ContractTool     # noqa: E402
from docmind.tools.gov_doc import GovDocTool        # noqa: E402
from docmind.tools.translate import TranslateTool   # noqa: E402
from docmind.tools.tech_doc import TechDocTool      # noqa: E402

register(ThesisTool())
register(ContractTool())
register(GovDocTool())
register(TranslateTool())
register(TechDocTool())
