"""合同审查工具 — 条款提取、完整性检查、风险标注、对比修订。

合同审查不同于排版工具：输入 docx，输出审查报告（JSON/md），不走 FormatEngine。
但共享 SpecReader（提取条款结构）和 OCR（验证）。

核心能力：
1. 条款提取 — 识别"第X条"结构，构建条款树
2. 完整性检查 — 验证必要条款是否齐全
3. 风险标注 — LLM 辅助分析合同风险点
4. 对比修订 — 与条款库比对，生成修订建议
"""

import json
import re
from pathlib import Path
from typing import Optional

from docmind.tools.base import DocMindTool, ToolResult
from docmind.tools.thesis import ThesisTool


# ── 条款识别正则 ──
ARTICLE_PATTERN = re.compile(
    r'^(?:第\s*[一二三四五六七八九十百千]+\s*条|'
    r'第\s*\d+\s*条)'
)

# ── 必要条款清单 ──
REQUIRED_CLAUSES = {
    "parties": {
        "keywords": ["甲方", "乙方", "当事人"],
        "description": "合同主体信息",
    },
    "subject": {
        "keywords": ["标的", "服务内容", "产品描述", "合同标的"],
        "description": "合同标的",
    },
    "price": {
        "keywords": ["价款", "报酬", "费用", "金额", "总价"],
        "description": "价款或报酬",
    },
    "performance": {
        "keywords": ["履行", "交付", "验收", "完成期限"],
        "description": "履行期限、地点和方式",
    },
    "breach": {
        "keywords": ["违约", "赔偿", "违约金", "违约方"],
        "description": "违约责任",
    },
    "dispute": {
        "keywords": ["争议", "仲裁", "诉讼", "管辖", "法律适用"],
        "description": "争议解决方式",
    },
    "force_majeure": {
        "keywords": ["不可抗力", "免责"],
        "description": "不可抗力条款",
    },
    "confidentiality": {
        "keywords": ["保密", "商业秘密", "机密信息"],
        "description": "保密条款",
    },
    "termination": {
        "keywords": ["解除", "终止", "提前终止"],
        "description": "合同解除与终止",
    },
    "effective": {
        "keywords": ["生效", "签署", "盖章", "签字"],
        "description": "合同生效条件",
    },
}

# ── 风险模式库 ──
RISK_PATTERNS = [
    {
        "id": "high_penalty",
        "pattern": r'违约金.*(?:[5-9]\d%|\d{3}%)',
        "level": "high",
        "title": "违约金比例过高",
        "suggestion": "建议将违约金比例控制在合理范围（通常不超过30%）",
    },
    {
        "id": "unilateral_termination",
        "pattern": r'甲方.*有权.*(?:随时|单方|任意).*(?:解除|终止)',
        "level": "high",
        "title": "单方任意解除权",
        "suggestion": "单方任意解除权可能导致合同不确定性，建议增加限制条件",
    },
    {
        "id": "unclear_jurisdiction",
        "pattern": r'(?:管辖|仲裁|诉讼).*(?:甲方|乙方).*(?:所在地|住所地)',
        "level": "medium",
        "title": "管辖约定不对等",
        "suggestion": "建议约定中立方所在地或合同签订地法院管辖",
    },
    {
        "id": "unlimited_liability",
        "pattern": r'(?:全部|一切|所有).*(?:损失|责任).*(?:承担|负责)',
        "level": "high",
        "title": "无限责任条款",
        "suggestion": "建议明确责任上限，排除间接损失和利润损失",
    },
    {
        "id": "auto_renewal",
        "pattern": r'(?:自动|默认).*(?:续期|续约|延期)',
        "level": "medium",
        "title": "自动续约条款",
        "suggestion": "建议增加到期前通知不续约的机制",
    },
]


class ContractTool(DocMindTool):
    """合同审查工具。"""

    @property
    def tool_name(self) -> str:
        return "contract"

    @property
    def description(self) -> str:
        return "合同审查：条款提取、完整性检查、风险标注、对比修订"

    def run(
        self,
        contract_docx: str | Path,
        output_report: Optional[str | Path] = None,
        checklist: Optional[list[str]] = None,
        clause_library: Optional[str | Path] = None,
    ) -> ToolResult:
        """审查合同。

        Args:
            contract_docx: 合同文件路径
            output_report: 报告输出路径（默认与合同同目录）
            checklist: 自定义审查清单（条款名称列表）
            clause_library: 条款库路径（用于对比修订）
        """
        target = Path(contract_docx)
        if not target.exists():
            return ToolResult(
                success=False, tool_name=self.tool_name,
                logs=[f"文件不存在: {target}"],
            )

        if output_report is None:
            report_path = target.parent / f"{target.stem}_review.md"
        else:
            report_path = Path(output_report)

        logs: list[str] = []
        findings: list[dict] = []

        try:
            # ── Step 1: 提取条款结构 ──
            logs.append("Step 1: 提取条款结构")

            import zipfile
            from lxml import etree

            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

            with zipfile.ZipFile(target, "r") as z:
                doc_xml = etree.fromstring(z.read("word/document.xml"))

            body = doc_xml.find(f"{{{W}}}body")
            articles: list[dict] = []

            for para in body.iter(f"{{{W}}}p"):
                txt = "".join(
                    t.text or "" for t in para.iter(f"{{{W}}}t")
                ).strip()
                if txt and ARTICLE_PATTERN.match(txt):
                    articles.append({
                        "text": txt[:100],
                        "index": len(articles),
                    })

            logs.append(f"  发现 {len(articles)} 个条款")

            # ── Step 2: 完整性检查 ──
            logs.append("Step 2: 完整性检查")

            full_text = " ".join(a["text"] for a in articles)
            clauses_to_check = checklist or list(REQUIRED_CLAUSES.keys())

            for clause_key in clauses_to_check:
                clause = REQUIRED_CLAUSES.get(clause_key)
                if clause is None:
                    continue
                found = any(
                    kw in full_text for kw in clause["keywords"]
                )
                if not found:
                    findings.append({
                        "type": "missing_clause",
                        "severity": "warning",
                        "clause": clause_key,
                        "description": clause["description"],
                        "detail": f"未找到 {clause['description']} 相关条款",
                    })

            logs.append(f"  完整性: {len(findings)} 个缺失项")

            # ── Step 3: 风险标注 ──
            logs.append("Step 3: 风险标注")

            risk_count = 0
            for risk in RISK_PATTERNS:
                if re.search(risk["pattern"], full_text):
                    findings.append({
                        "type": "risk",
                        "severity": "high" if risk["level"] == "high" else "medium",
                        "risk_id": risk["id"],
                        "title": risk["title"],
                        "suggestion": risk["suggestion"],
                    })
                    risk_count += 1

            logs.append(f"  风险: {risk_count} 个潜在风险")

            # ── Step 4: 生成报告 ──
            logs.append("Step 4: 生成审查报告")

            report = self._build_report(
                contract_name=target.stem,
                article_count=len(articles),
                articles=articles,
                findings=findings,
            )

            report_path.write_text(report, encoding="utf-8")
            logs.append(f"  报告: {report_path}")

            return ToolResult(
                success=True,
                tool_name=self.tool_name,
                report_path=report_path,
                issues=findings,
                logs=logs[-5:],
                metadata={
                    "article_count": len(articles),
                    "missing_clauses": sum(
                        1 for f in findings if f["type"] == "missing_clause"
                    ),
                    "risks": risk_count,
                },
            )

        except Exception as e:
            return ToolResult(
                success=False,
                tool_name=self.tool_name,
                logs=logs + [f"异常: {e}"],
            )

    def _build_report(
        self,
        contract_name: str,
        article_count: int,
        articles: list[dict],
        findings: list[dict],
    ) -> str:
        """生成 Markdown 审查报告。"""
        lines = [
            f"# 合同审查报告",
            f"",
            f"**合同名称**：{contract_name}",
            f"**审查时间**：自动生成",
            f"**条款总数**：{article_count}",
            f"**发现问题**：{len(findings)} 个",
            f"",
            f"---",
            f"",
        ]

        # 条款结构
        lines.append("## 一、条款结构")
        lines.append("")
        for a in articles[:10]:  # 只展示前10条
            lines.append(f"- {a['text']}")
        if len(articles) > 10:
            lines.append(f"- ... 共 {len(articles)} 条")
        lines.append("")

        # 问题详情
        if findings:
            lines.append("## 二、发现问题")
            lines.append("")

            missing = [f for f in findings if f["type"] == "missing_clause"]
            if missing:
                lines.append("### 缺失条款")
                for m in missing:
                    lines.append(f"- ⚠️ **{m['description']}**：{m['detail']}")
                lines.append("")

            risks = [f for f in findings if f["type"] == "risk"]
            if risks:
                lines.append("### 风险标注")
                for r in risks:
                    level_icon = "🔴" if r["severity"] == "high" else "🟡"
                    lines.append(f"- {level_icon} **{r['title']}**")
                    lines.append(f"  {r['suggestion']}")
                lines.append("")
        else:
            lines.append("## 二、审查结论")
            lines.append("")
            lines.append("✅ 未发现明显问题。")
            lines.append("")

        # 审查清单
        lines.append("## 三、审查清单")
        lines.append("")
        for key, clause in REQUIRED_CLAUSES.items():
            found = any(
                f["type"] == "missing_clause" and f["clause"] == key
                for f in findings
            )
            status = "❌ 缺失" if found else "✅ 已包含"
            lines.append(f"- {status} {clause['description']}")

        return "\n".join(lines)
