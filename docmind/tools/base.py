"""DocMind 工具基类。

每个垂直领域能力是一个 DocMindTool 子类。
Agent 根据用户意图路由到对应工具，工具调用共享 Format Engine。
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional


@dataclass
class ToolResult:
    """工具执行结果"""

    success: bool
    tool_name: str = ""
    output_path: Optional[Path] = None   # 输出文件（docx）
    report_path: Optional[Path] = None   # 审查报告（md/json）
    issues: list[dict] = field(default_factory=list)
    logs: list[str] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


class DocMindTool(ABC):
    """所有工具的基类。

    子类只需实现三个属性 + run() 方法：
      - tool_name: 工具标识
      - description: 工具描述（给 Router 看）
      - supported_extensions: 支持的输入文件后缀
    """

    @property
    @abstractmethod
    def tool_name(self) -> str:
        """工具唯一标识，如 'thesis', 'contract', 'gov_doc'"""
        ...

    @property
    @abstractmethod
    def description(self) -> str:
        """一句话描述，用于意图路由"""
        ...

    @property
    def supported_extensions(self) -> list[str]:
        """支持的输入文件后缀"""
        return [".docx"]

    @abstractmethod
    def run(self, **kwargs) -> ToolResult:
        """执行工具。

        每个工具定义自己的参数，基类只约束返回类型。
        """
        ...
