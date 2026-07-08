"""
Agent 主循环 — 串联全流水线并支持用户反馈闭环。

流程:
    SpecReader → DocDiscover → Reconciler → Writer → Fixer → 报告
    ↓
    用户反馈 → 更新 fix_plan → 重新 Writer → Fixer → 报告
    （闭环直到用户满意）

用法:
    from pipeline.agent import DocMindAgent

    agent = DocMindAgent()
    result = agent.run(
        spec_path=Path('output/附件7：毕业设计（论文）撰写规范.docx'),
        target_path=Path('output/毕业设计.docx'),
        output_path=Path('output/毕业设计_修正版.docx'),
    )
    # result: {"plan": [...], "issues": [...], "fixed_count": N, "remaining": M}

    # 用户反馈闭环
    agent.handle_feedback("参考文献前面的空白页应该删除")
"""

import json
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

# Archived modules (enforce) path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / '_archive' / 'docmind'))
from typing import Any, Optional, Union

from lxml import etree

from .writer_v2 import apply_fixes
from .fixer_v2 import diagnose, FixerDiagnostic, summary as fixer_summary
from .spec_reader_v2 import read_spec_v2 as read_spec

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class AgentResult:
    """Agent 运行结果"""
    plan: list[dict] = field(default_factory=list)       # fix_plan
    issues: list[FixerDiagnostic] = field(default_factory=list)
    fixed_count: int = 0
    remaining: int = 0
    output_path: Optional[Path] = None
    output_print: Optional[Path] = None
    logs: list[str] = field(default_factory=list)


@dataclass 
class FeedbackRule:
    """持久化的用户反馈规则"""
    key: str           # 规则键，如 "body_font_east"
    value: Any         # 规则值，如 "宋体"
    reason: str = ""   # 用户原始反馈
    when: str = ""     # ISO 时间戳


class DocMindAgent:
    """DocMind 排版 Agent — 闭环编排器"""

    OVERRIDES_FILE = ".docmind_overrides.json"

    def __init__(self):
        self._fix_plan: list[dict] = []
        self._spec_rules: dict = {}
        self._section_count: int = 0
        self._section_headers: dict = {}
        self._style_ids: dict = {}
        self._heading_info: dict[str, Any] = {}
        self._spec_path: Optional[Path] = None
        self._target_path: Optional[Path] = None
        self._output_path: Optional[Path] = None
        self._user_overrides: dict[str, Any] = {}
        self._feedback_rules: list[FeedbackRule] = []  # 持久化规则

    def run(
        self,
        spec_path: Union[str, Path],
        target_path: Union[str, Path],
        output_path: Union[str, Path] = None,
        print_mode: bool = False,
        output_print: Union[str, Path] = None,
        libreoffice_path: str = None,
        h1_pattern: str = None,
        front_skip_headings: set = None,
    ) -> AgentResult:
        """执行完整的排版闭环。

        Args:
            spec_path: 规范模板 docx 路径
            target_path: 目标文档 docx 路径
            output_path: 电子版输出路径
            print_mode: 是否生成打印版（奇数页强制）
            output_print: 打印版输出路径
            libreoffice_path: LibreOffice 路径
            h1_pattern: H1 标题正则（None=论文默认）
            front_skip_headings: 不设页眉的标题集合（None=论文默认）
        """
        self._spec_path = Path(spec_path)
        self._target_path = Path(target_path)
        self._output_path = Path(output_path) if output_path else self._target_path
        self._libreoffice_path = libreoffice_path

        # ── 领域参数：覆盖默认 thesis 行为 ──
        self._h1_pat = re.compile(h1_pattern) if h1_pattern else None
        self._front_skip = front_skip_headings

        # 自动加载持久化反馈规则
        self._load_overrides()

        # ── 创建备份（回滚保险）──
        backup = self._target_path.parent / (self._target_path.stem + ".docmind_backup")
        self._backup_path = backup
        try:
            shutil.copy2(self._target_path, backup)
        except OSError as e:
            print(f"  [WARN] 无法创建备份: {e}")
            self._backup_path = None

        logs = []

        try:
            # ── Step 1: SpecReader ──
            logs.append("=" * 60)
            logs.append("STEP 1/5: SpecReader — 读取规范模板")
            logs.append("=" * 60)

            self._spec_rules = read_spec(self._spec_path)
            logs.append(f"  规范规则: {list(self._spec_rules.keys())}")

            if "h1" in self._spec_rules:
                h1 = self._spec_rules["h1"]
                logs.append(f"  H1: {h1.get('font_east')} {h1.get('font_size_pt')}pt bold={h1.get('bold')}")

            if "body" in self._spec_rules:
                b = self._spec_rules["body"]
                logs.append(f"  Body: {b.get('font_east')} {b.get('font_size_pt')}pt")

            # ── Step 2: DocDiscover ──
            logs.append("")
            logs.append("=" * 60)
            logs.append("STEP 2/5: DocDiscover — 发现目标文档结构")
            logs.append("=" * 60)

            self._discover_doc()
            logs.append(f"  节数: {self._section_count}")
            logs.append(f"  样式 ID: {self._style_ids}")

            if self._heading_info:
                logs.append(f"  标题: H1={self._heading_info.get('h1_count', 0)} "
                            f"H2={self._heading_info.get('h2_count', 0)} "
                            f"H3={self._heading_info.get('h3_count', 0)}")

            # ── Step 3: Reconciler ──
            logs.append("")
            logs.append("=" * 60)
            logs.append("STEP 3/5: Reconciler — 对比生成 fix_plan")
            logs.append("=" * 60)

            self._fix_plan = self._reconcile()
            logs.append(f"  fix_plan 条目: {len(self._fix_plan)}")

            for i, action in enumerate(self._fix_plan):
                action_type = action.get("action", "?")
                params = action.get("params", {})
                brief = _brief_action(action_type, params)
                logs.append(f"    #{i}: {brief}")

            # ── Step 4: Writer ──
            logs.append("")
            logs.append("=" * 60)
            logs.append("STEP 4/5: Writer — 执行 fix_plan")
            logs.append("=" * 60)

            # 纯 OOXML 路径：直接修改磁盘文件
            logs.append("  → 使用 Python OOXML 直写磁盘")
            if self._output_path != self._target_path:
                shutil.copy2(self._target_path, self._output_path)

            try:
                writer_logs = apply_fixes(self._fix_plan, self._output_path)
            except Exception as we:
                logs.append(f"  [FAIL] Writer 异常: {we}")
                writer_logs = [f"[FAIL] {we}"]

            for line in writer_logs:
                logs.append(f"  {line}")

            fixed_count = sum(1 for l in writer_logs if l.startswith("[OK"))
            failed_count = sum(1 for l in writer_logs if l.startswith("[FAIL"))

            # ── Step 5: Fixer ──
            logs.append("")
            logs.append("=" * 60)
            logs.append("STEP 5/5: Fixer — 诊断执行结果")
            logs.append("=" * 60)

            issues = diagnose(self._fix_plan, self._output_path)
            stats = fixer_summary(issues)

            logs.append(f"  诊断结果: {stats['total']} 个问题")
            for code, count in stats["by_code"].items():
                logs.append(f"    {code}: {count}")

            if not stats["passed"]:
                for iss in issues:
                    logs.append(f"  {iss}")

            # ── 结果 ──
            result = AgentResult(
                plan=self._fix_plan,
                issues=issues,
                fixed_count=fixed_count,
                remaining=len(issues),
                output_path=self._output_path,
                logs=logs,
            )

            # ── Step 6: Enforce ──
            if print_mode:
                logs.append("")
                logs.append("=" * 60)
                logs.append("STEP 6/6: Enforce — 打印版奇数页强制")
                logs.append("=" * 60)
                print_output = Path(output_print) if output_print else \
                    self._output_path.parent / (self._output_path.stem + "_print.docx")
                lo_path = libreoffice_path or "D:/LibreOffice/program/soffice.exe"
                try:
                    print_log = self._enforce_odd_pages(self._output_path, print_output, lo_path, logs)
                    result.output_print = print_output
                except Exception as ee:
                    logs.append(f"  [WARN] 奇数页强制失败（非致命）: {ee}")

            # ── 打印日志 ──
            for line in logs:
                print(line)

            return result

        except Exception as e:
            # ── 回滚：恢复备份 ──
            print(f"\n{'='*60}")
            print(f"[FAIL] Pipeline 异常: {e}")
            print(f"{'='*60}")
            if self._backup_path and self._backup_path.exists():
                try:
                    shutil.copy2(self._backup_path, self._target_path)
                    print(f"  [OK] 已从备份恢复原文件")
                except OSError as re:
                    print(f"  [FAIL] 回滚失败: {re}")
                    print(f"  备份文件: {self._backup_path}")
            import traceback
            traceback.print_exc()
            raise


    def _discover_doc(self) -> None:
        """使用 doc_discover 发现目标文档结构"""
        from .doc_discover import discover_doc

        doc_state = discover_doc(self._target_path)

        # 节数
        self._section_count = len(doc_state.get('sections', []))

        # 标题统计
        headings = doc_state.get('headings', [])
        h1 = [h for h in headings if h.get('level') == 1]
        h2 = [h for h in headings if h.get('level') == 2]
        h3 = [h for h in headings if h.get('level') == 3]

        self._heading_info = {
            "h1_count": len(h1),
            "h2_count": len(h2),
            "h3_count": len(h3),
            "h1_texts": [h['text'] for h in h1],
        }

        # 样式 ID 映射
        self._style_ids = {
            "h1": self._spec_rules.get("h1_style_id", "1"),
            "h2": self._spec_rules.get("h2_style_id", "20"),
            # 目标文档常用 body 样式 ID；规范模板的 'Normal' 映射到 'a8'
            "body": self._spec_rules.get("body_style_id", "a8"),
        }
        # 如果规范模板返回 'Normal' 但目标文档用 'a8'，自动修正
        if self._style_ids["body"] == "Normal":
            self._style_ids["body"] = "a8"

    def _reconcile(self) -> list[dict]:
        """对比规范规则与文档结构，生成 fix_plan。

        策略:
        1. 样式修改：将规范中的 H1/H2/Body 格式应用到目标文档
        2. 节类型：确保所有 sectPr 的 type 为 nextPage
        3. 页眉字体：设置 header XML 中的 rFonts
        4. 页码：在 footer XML 中添加页码域
        5. 清理：移除嵌入段落中的多余 sectPr（特别是 oddPage）
        """
        plan: list[dict] = []

        # ── 1. 样式修改 ──
        self._reconcile_styles(plan)

        # ── 2. 节类型修改 ──
        self._reconcile_section_types(plan)

        # ── 3. 页眉字体 ──
        self._reconcile_header_fonts(plan)

        # ── 4. 页码 ──
        self._reconcile_page_numbers(plan)

        # ── 4a. 页脚链接 ──
        self._reconcile_footer_links(plan)

        # ── 4b. 页眉文件独立化 ──
        self._reconcile_header_files(plan)

        # ── 5. 清理多余 sectPr（仅在确认有重复时才做）
        # 注意：python-docx 把 sectPr 嵌入内容段落是正常的，不能全删
        # 只在 preprocessor 插了新的空段分节符导致双重分节符时才清理
        # self._reconcile_extra_sectpr(plan)

        # ── 应用用户覆盖 ──
        self._apply_user_overrides(plan)

        return plan

    def _reconcile_styles(self, plan: list[dict]) -> None:
        """从 spec_rules 生成样式修改 actions"""
        # H1 样式
        h1_rules = self._spec_rules.get("h1", {})
        if h1_rules:
            h1_params = {"style_id": self._style_ids.get("h1", "1")}
            _copy_format_params(h1_rules, h1_params)
            if len(h1_params) > 1:  # 除了 style_id 还有别的
                plan.append({"action": "set_style", "params": h1_params})

        # H2 样式
        h2_rules = self._spec_rules.get("h2", {})
        if h2_rules:
            h2_params = {"style_id": self._style_ids.get("h2", "20")}
            _copy_format_params(h2_rules, h2_params)
            if len(h2_params) > 1:
                plan.append({"action": "set_style", "params": h2_params})

        # H3 样式
        h3_rules = self._spec_rules.get("h3", {})
        if h3_rules:
            h3_params = {"style_id": "30"}  # 默认 Heading 3
            _copy_format_params(h3_rules, h3_params)
            if len(h3_params) > 1:
                plan.append({"action": "set_style", "params": h3_params})

        # Body 样式
        body_rules = self._spec_rules.get("body", {})
        if body_rules:
            body_params = {"style_id": self._style_ids.get("body", "a8")}
            _copy_format_params(body_rules, body_params)
            if len(body_params) > 1:
                plan.append({"action": "set_style", "params": body_params})

    def _reconcile_section_types(self, plan: list[dict]) -> None:
        """确保所有 sectPr 的 type 为 nextPage（排除 oddPage 等）"""
        with zipfile.ZipFile(self._target_path, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))

        sects = list(root.iter(f"{{{W}}}sectPr"))
        for i, sp in enumerate(sects):
            type_el = sp.find(f"{{{W}}}type")
            current = type_el.get(f"{{{W}}}val") if type_el is not None else None

            # oddPage → nextPage
            if current == "oddPage":
                plan.append({
                    "action": "set_sectpr_type",
                    "params": {"section_index": i, "val": "nextPage"},
                })

    def _reconcile_header_fonts(self, plan: list[dict]) -> None:
        """设置 header XML 文件中的字体（使用规范中的 body 字体作为参考）"""
        body_rules = self._spec_rules.get("body", {})
        font_east = body_rules.get("font_east", "宋体")
        font_ascii = body_rules.get("font_ascii", "Times New Roman")

        with zipfile.ZipFile(self._target_path, "r") as z:
            header_names = [n for n in z.namelist() if "header" in n and n.endswith(".xml")]

        for hdr_name in header_names:
            plan.append({
                "action": "set_header_font",
                "params": {
                    "header_path": hdr_name,
                    "font_ascii": font_ascii,
                    "font_eastAsia": font_east,
                    "font_size": "18",  # 小五号 = 9pt = 18 half-points
                },
            })

    def _reconcile_page_numbers(self, plan: list[dict]) -> None:
        """添加页码到 footer XML 文件 + 设置罗马/阿拉伯格式。"""
        fmt = "第{PAGE}页"  # 默认格式

        if "page_number_format" in self._user_overrides:
            fmt = self._user_overrides["page_number_format"]

        # 从规范模板读取页码对齐方式
        footer_fmt = self._spec_rules.get("footer_format", {}) or {}
        alignment = footer_fmt.get("alignment", "right")  # 中文论文默认右对齐

        with zipfile.ZipFile(self._target_path, "r") as z:
            footer_names = [n for n in z.namelist() if "footer" in n and n.endswith(".xml")]

        body_rules = self._spec_rules.get("body", {})
        font_ascii = body_rules.get("font_ascii", "Times New Roman")

        for ftr_name in footer_names:
            plan.append({
                "action": "add_page_number",
                "params": {
                    "footer_path": ftr_name,
                    "format": fmt,
                    "alignment": alignment,
                    "font_eastAsia": font_ascii,
                    "font_ascii": font_ascii,
                    "font_size": "18",
                },
            })

        # ── 页码格式 ──
        # 郑重声明 = 前部 (同封面)，无页码
        # 摘要→目录 = 大罗马
        # 正文（第一个编号章起）= 阿拉伯从1开始

        # 自行构建 section_sectpr_map（_reconcile_header_files 在后面才执行）
        with zipfile.ZipFile(self._target_path, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        body = root.find(f"{{{W}}}body")

        h1_pat = re.compile(
            r'^(?:封面|郑\s*重\s*声\s*明|郑重声明|摘\s*要|ABSTRACT|'
            r'目\s*录|目录|\d+[\u4e00-\u9fff]|'
            r'结\s*论|结论|参考\s*文献|参考文献|致\s*谢|致谢|'
            r'附\s*录|附录)'
        )
        h1_positions = []
        for i, child in enumerate(body):
            para = None
            if child.tag == f"{{{W}}}p":
                para = child
            elif child.tag == f"{{{W}}}sdt":
                sdt_content = child.find(f"{{{W}}}sdtContent")
                if sdt_content is not None:
                    para = sdt_content.find(f"{{{W}}}p")
            if para is not None:
                txt = ''.join(t.text or '' for t in para.iter(f"{{{W}}}t")).strip()
                if txt and h1_pat.match(re.sub(r'\s+', '', txt)):
                    h1_positions.append((i, txt))

        sectpr_map = []
        last_heading = ""
        phys_sectpr_idx = 0
        for child_idx, child in enumerate(body):
            sp = None
            if child.tag == f"{{{W}}}sectPr":
                sp = child
            elif child.tag == f"{{{W}}}p":
                pp = child.find(f"{{{W}}}pPr")
                if pp is not None:
                    sp = pp.find(f"{{{W}}}sectPr")
            if sp is not None:
                heading_text = ""
                for hi, ht in reversed(h1_positions):
                    if hi < child_idx:
                        heading_text = ht
                        break
                if heading_text and heading_text != last_heading:
                    sectpr_map.append((phys_sectpr_idx, heading_text))
                    last_heading = heading_text
                phys_sectpr_idx += 1

        body_start_phys = None
        toc_end_phys = None
        front_declaration_phys = None  # 郑重声明的 phys index
        for phys_idx, heading in sectpr_map:
            h_norm = re.sub(r'\s+', '', heading)
            if "目录" in h_norm:
                toc_end_phys = phys_idx
            if "郑重声明" in h_norm or "郑重" in heading:
                front_declaration_phys = phys_idx
            if body_start_phys is None and re.match(r'\d+[\u4e00-\u9fff]', h_norm):
                body_start_phys = phys_idx

        # 摘要→目录 → 大罗马（跳过郑重声明，它无页码）
        for phys_idx, heading in sectpr_map:
            h_norm = re.sub(r'\s+', '', heading)
            if toc_end_phys is not None and phys_idx <= toc_end_phys:
                # 郑重声明不设页码格式
                if front_declaration_phys is not None and phys_idx == front_declaration_phys:
                    continue
                p = {"section_index": phys_idx, "fmt": "upperRoman"}
                if phys_idx == (front_declaration_phys or 0) + 1:
                    p["start"] = 1  # 摘要 = I（郑重声明后的第一节）
                plan.append({"action": "set_page_number_type", "params": p})

        # 第一个 body 节 → 阿拉伯从1开始
        if body_start_phys is not None:
            plan.append({
                "action": "set_page_number_type",
                "params": {"section_index": body_start_phys, "fmt": "decimal", "start": 1},
            })

        # 后续 body 节 → 阿拉伯，继续编号
        for phys_idx, heading in sectpr_map:
            if body_start_phys is not None and phys_idx > body_start_phys:
                plan.append({
                    "action": "set_page_number_type",
                    "params": {"section_index": phys_idx, "fmt": "decimal"},
                })

    def _reconcile_footer_links(self, plan: list[dict]) -> None:
        """为每节 sectPr 添加 footerReference，确保页码实际显示。
        
        封面(sectPr[0]) 和 郑重声明(sectPr[1]) 不设页脚页码。
        """
        # 从 sectPr[2]（摘要）起添加 footerReference
        for section_index in range(2, 14):
            plan.append({
                "action": "add_footer_reference",
                "params": {"section_index": section_index},
            })
        # 清除封面(节0) + 郑重声明(节1) 的页眉——前置两节不应有页眉。
        # 节0 是首节，清引用后无处继承 → 空白；节1 继承已空白的节0 → 空白。
        # （历史 bug：只清了节1，节0 保留陈旧共享页眉[摘要]，节1 反而继承了它 → 两页都显示"摘要"）
        for section_index in (0, 1):
            plan.append({
                "action": "clear_section_headers",
                "params": {"section_index": section_index},
            })

    def _ocr_section_audit(self, libreoffice_path: str = None) -> list:
        """Phase 0: OCR 辅助检测 — 发现 OOXML H1 正则遗漏的视觉标题。

        将文档转 PDF → OCR 扫描所有页 → 检测页顶区域的标题级文字
        → 与 OOXML H1 列表对比 → 返回 OCR 发现但 OOXML 遗漏的标题。

        作为 OOXML 拆分的兜底方案：当模板中的标题没有标准样式标记
        （如 ABSTRACT 只是加粗段落而非 Heading 1），OOXML 层检测不到，
        但 OCR 可以看到渲染后的视觉标题。
        """
        import subprocess

        libreoffice_path = libreoffice_path or getattr(self, '_libreoffice_path', None)
        if not libreoffice_path:
            return []

        lo = Path(libreoffice_path)
        if not lo.exists():
            return []

        src = self._target_path
        try:
            # ── 转 PDF ──
            pdf_path = src.parent / f"{src.stem}_ocr_audit.pdf"
            result = subprocess.run(
                [str(lo), '--headless', '--convert-to', 'pdf',
                 '--outdir', str(src.parent), str(src)],
                capture_output=True, text=True, timeout=120,
            )
            if result.returncode != 0 or not pdf_path.exists():
                return []

            # ── PDF → PNG → OCR ──
            from ocr_observer import OcrObserver
            obs = OcrObserver(model_type='medium')
            audits = obs.audit_docx(src, dpi=120, verbose=False)
            if not audits:
                return []

            # ── 提取页顶区域标题候选 ──
            page_height = getattr(audits[0], 'page_height', None) or 842
            body_top_max_y = page_height * 0.30
            header_max_y = page_height * 0.15
            max_heading_len = 20

            page_headings = {}
            for audit in audits:
                pn = audit.page_index
                candidates = []
                for region in audit.body_texts:
                    if header_max_y < region.center_y < body_top_max_y:
                        txt = region.text.strip()
                        if 1 <= len(txt) <= max_heading_len:
                            if txt[-1] not in '.。，,;；：:':
                                candidates.append(txt)
                if candidates:
                    page_headings[pn] = candidates

            # ── 收集 OOXML H1 列表 ──
            h1_norm = set()
            with zipfile.ZipFile(src, 'r') as z:
                root = etree.fromstring(z.read('word/document.xml'))
            body = root.find(f'{{{W}}}body')
            h1_pat = re.compile(
                r'^(?:封面|郑\s*重\s*声\s*明|摘\s*要|ABSTRACT|'
                r'目\s*录|目录|\d+[一-鿿]|'
                r'结\s*论|结论|参考\s*文献|参考文献|致\s*谢|致谢|'
                r'附\s*录|附录)'
            )
            for child in body:
                para = None
                if child.tag == f'{{{W}}}p':
                    para = child
                elif child.tag == f'{{{W}}}sdt':
                    sc = child.find(f'{{{W}}}sdtContent')
                    if sc is not None:
                        para = sc.find(f'{{{W}}}p')
                if para is not None:
                    txt = ''.join(t.text or '' for t in para.iter(f'{{{W}}}t')).strip()
                    if txt:
                        norm = re.sub(r'\s+', '', txt)
                        if h1_pat.match(norm):
                            h1_norm.add(norm)

            # ── 对比：OCR 有但 OOXML 没有的标题 ──
            missed = []
            for pn, candidates in page_headings.items():
                for c in candidates:
                    c_norm = re.sub(r'\s+', '', c)
                    if c_norm and c_norm not in h1_norm:
                        if (len(c_norm) >= 2
                            and not c_norm.isdigit()
                            and not all(ord(ch) < 128 for ch in c_norm if ch.isalpha())):
                            missed.append(c)

            # 清理临时 PDF
            try:
                pdf_path.unlink()
            except OSError:
                pass

            return missed

        except Exception as e:
            print(f"  [OCR audit] 跳过: {e}")
            return []


    def _reconcile_header_files(self, plan: list[dict]) -> None:
        """为每节创建独立 header 文件。"

        检测共享 sectPr 的 H1（如摘要→ABSTRACT），先插入分节符，再创建独立 header。
        全部在 reconcile 阶段完成——不依赖 Writer 执行顺序。
        """
        font_east = self._spec_rules.get('body', {}).get('font_east', '宋体')
        font_ascii = self._spec_rules.get('body', {}).get('font_ascii', 'Times New Roman')

        # ═══ Phase 0: OCR 辅助检测（兜底方案） ═══
        ocr_missed = self._ocr_section_audit()
        if ocr_missed:
            print(f"  [OCR] 发现 OOXML 遗漏的视觉标题: {ocr_missed}")

        # ═══ Phase 1: 检测共享 sectPr 的 H1，直接插入分节符 ═══
        with zipfile.ZipFile(self._target_path, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        body = root.find(f"{{{W}}}body")

        h1_pat = re.compile(
            r'^(?:封面|郑\s*重\s*声\s*明|郑重声明|摘\s*要|ABSTRACT|'
                        r'目\s*录|目录|\d+[\u4e00-\u9fff]|'
                        r'结\s*论|结论|参考\s*文献|参考文献|致\s*谢|致谢|'
                        r'附\s*录|附录)'
        )
        h1_positions = []
        for i, child in enumerate(body):
            para = None
            if child.tag == f"{{{W}}}p":
                para = child
            elif child.tag == f"{{{W}}}sdt":
                sdt_content = child.find(f"{{{W}}}sdtContent")
                if sdt_content is not None:
                    first_p = sdt_content.find(f"{{{W}}}p")
                    if first_p is not None:
                        para = first_p
            if para is not None:
                txt = ''.join(t.text or '' for t in para.iter(f"{{{W}}}t")).strip()
                if txt and h1_pat.match(re.sub(r'\s+', '', txt)):
                    h1_positions.append((i, txt))

        sects_in_order = []
        for child_idx, child in enumerate(body):
            sp = None
            if child.tag == f"{{{W}}}sectPr":
                sp = child
            elif child.tag == f"{{{W}}}p":
                pp = child.find(f"{{{W}}}pPr")
                if pp is not None:
                    sp = pp.find(f"{{{W}}}sectPr")
            if sp is not None:
                sects_in_order.append((child_idx, sp))

        # 检测共享 sectPr 的 H1
        h1_to_sectpr = {}
        for si, (child_idx, sp) in enumerate(sects_in_order):
            for h1_idx, h1_txt in h1_positions:
                if h1_idx < child_idx:
                    if h1_idx not in h1_to_sectpr or child_idx < h1_to_sectpr[h1_idx][1]:
                        h1_to_sectpr[h1_idx] = (si, child_idx, h1_txt)

        # 找需要拆分的 H1 对
        splits = []
        for i in range(len(h1_positions)):
            h1_idx_a, h1_txt_a = h1_positions[i]
            if h1_idx_a not in h1_to_sectpr:
                continue
            si_a = h1_to_sectpr[h1_idx_a][0]
            for j in range(i + 1, len(h1_positions)):
                h1_idx_b, h1_txt_b = h1_positions[j]
                if h1_idx_b not in h1_to_sectpr:
                    continue
                if h1_to_sectpr[h1_idx_b][0] == si_a and h1_txt_a != h1_txt_b:
                    splits.append(h1_txt_b)
                    break

        # ── OCR 辅助拆分：将 OCR 发现的遗漏标题注入 splits ──
        for missed_heading in ocr_missed:
            norm_missed = re.sub(r'\s+', '', missed_heading)
            # 避免重复
            if not any(re.sub(r'\s+', '', s) == norm_missed for s in splits):
                splits.append(missed_heading)

        # ── 封面拆分：如果第一段包含封面内容 + 第一个 H1，在 H1 前插分节符 ──
        if h1_positions:
            first_h1_idx, first_h1_txt = h1_positions[0]
            if first_h1_idx in h1_to_sectpr:
                si_first = h1_to_sectpr[first_h1_idx][0]
                # 如果第一个 H1 在第一个 sectPr 中，且前面有足够的内容（封面页）
                if si_first == 0 and first_h1_idx > 3:
                    # 封面本身不是 H1，但内容在第一个 section 中
                    # 需要拆分：封面独立一节，郑重声明另起一节
                    splits.insert(0, first_h1_txt)

        # 直接插入分节符（在 reconcile 阶段，不经过 Writer）
        inserted = 0
        for heading_text in splits:
            for child in list(body):
                if child.tag == f"{{{W}}}p":
                    txt = ''.join(t.text or '' for t in child.iter(f"{{{W}}}t")).strip()
                    if re.sub(r'\s+', '', txt) == re.sub(r'\s+', '', heading_text):
                        new_sp = etree.Element(f"{{{W}}}p")
                        new_pPr = etree.SubElement(new_sp, f"{{{W}}}pPr")
                        sectPr_el = etree.SubElement(new_pPr, f"{{{W}}}sectPr")
                        etree.SubElement(sectPr_el, f"{{{W}}}type", {f"{{{W}}}val": "nextPage"})
                        etree.SubElement(sectPr_el, f"{{{W}}}pgSz", {f"{{{W}}}w": "11906", f"{{{W}}}h": "16838"})
                        child.addprevious(new_sp)
                        inserted += 1
                        break

        if inserted > 0:
            new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            tmp = str(self._target_path) + '.split_tmp'
            with zipfile.ZipFile(self._target_path) as zin:
                with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)
                        if item.filename == 'word/document.xml':
                            data = new_xml
                        zout.writestr(item, data)
            shutil.move(tmp, str(self._target_path))

        # ═══ Phase 2: 重新读取，计算正确的 section_headers ═══
        with zipfile.ZipFile(self._target_path, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        body = root.find(f"{{{W}}}body")

        # 重建 sects_in_order
        sects_in_order = []
        for child_idx, child in enumerate(body):
            sp = None
            if child.tag == f"{{{W}}}sectPr":
                sp = child
            elif child.tag == f"{{{W}}}p":
                pp = child.find(f"{{{W}}}pPr")
                if pp is not None:
                    sp = pp.find(f"{{{W}}}sectPr")
            if sp is not None:
                sects_in_order.append((child_idx, sp))

        # 重建 h1_positions
        h1_positions = []
        for i, child in enumerate(body):
            para = None
            if child.tag == f"{{{W}}}p":
                para = child
            elif child.tag == f"{{{W}}}sdt":
                sdt_content = child.find(f"{{{W}}}sdtContent")
                if sdt_content is not None:
                    first_p = sdt_content.find(f"{{{W}}}p")
                    if first_p is not None:
                        para = first_p
            if para is not None:
                txt = ''.join(t.text or '' for t in para.iter(f"{{{W}}}t")).strip()
                if txt and h1_pat.match(re.sub(r'\s+', '', txt)):
                    h1_positions.append((i, txt))

        # 为每个 sectPr 匹配 governing H1（向前搜索）
        # 封面/郑重声明区域不设页眉
        FRONT_SKIP_HEADER = {'封面', '郑重声明', '毕业设计'}
        section_headers = []
        for si, (child_idx, sp) in enumerate(sects_in_order):
            heading_text = None
            for h1_idx, h1_txt in reversed(h1_positions):
                if h1_idx < child_idx:
                    heading_text = h1_txt
                    break
            if heading_text:
                norm = re.sub(r'\s+', '', heading_text)
                # 封面/郑重声明不设页眉（它们的空白页应干净）
                even = None
                if norm in FRONT_SKIP_HEADER:
                    heading_text = None  # 跳过
                if heading_text:
                    # 标准化页眉文字: 去空格 + 去章节编号前缀
                    header_text = re.sub(r'\s+', '', heading_text)
                    m = re.match(r'^\d+([\u4e00-\u9fff].*)', header_text)
                    if m:
                        header_text = m.group(1)
                    section_headers.append({
                        'sectpr_index': si,
                        'default_text': header_text,
                        'even_text': even,
                    })

        if section_headers:
            plan.append({
                'action': 'create_section_headers',
                'params': {
                    'section_headers': section_headers,
                    'font_eastAsia': font_east,
                    'font_ascii': font_ascii,
                    'font_size': '18',
                },
            })

    def _enforce_odd_pages(self, src: Path, dst: Path, libreoffice_path: str, logs: list) -> str:
        """调用 enforce/odd_pages.py，传递正确的 sectPr 索引映射。"""
        from enforce.config import EnforceConfig
        from enforce.odd_pages import enforce_odd_pages

        # 从修改后的文档（src）读取 sectPr 位置和标题映射
        with zipfile.ZipFile(src, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        body = root.find(f"{{{W}}}body")

        h1_pat = re.compile(
            r'^(?:封面|郑\s*重\s*声\s*明|郑重声明|摘\s*要|ABSTRACT|'
                        r'目\s*录|目录|\d+[\u4e00-\u9fff]|'
                        r'结\s*论|结论|参考\s*文献|参考文献|致\s*谢|致谢|'
                        r'附\s*录|附录)'
        )
        h1_positions = []
        for i, child in enumerate(body):
            para = None
            if child.tag == f"{{{W}}}p":
                para = child
            elif child.tag == f"{{{W}}}sdt":
                sdt_content = child.find(f"{{{W}}}sdtContent")
                if sdt_content is not None:
                    first_p = sdt_content.find(f"{{{W}}}p")
                    if first_p is not None:
                        para = first_p
            if para is not None:
                txt = ''.join(t.text or '' for t in para.iter(f"{{{W}}}t")).strip()
                if txt and h1_pat.match(re.sub(r'\s+', '', txt)):
                    h1_positions.append((i, txt))

        # Build sectPr→heading map with PHYSICAL sectPr indices.
        # Physical index = position in the body's sectPr order (used by odd_pages.py
        # to find the correct paragraph). Logical dedup ensures no duplicate headings.
        sectpr_map = []       # [(physical_sectpr_idx, heading), ...]
        last_heading = ""
        phys_sectpr_idx = 0   # physical counter (all sectPr in body order)
        for child_idx, child in enumerate(body):
            sp = None
            if child.tag == f"{{{W}}}sectPr":
                sp = child
            elif child.tag == f"{{{W}}}p":
                pp = child.find(f"{{{W}}}pPr")
                if pp is not None:
                    sp = pp.find(f"{{{W}}}sectPr")
            if sp is None:
                continue
            heading_text = ""
            for h1_idx, h1_txt in reversed(h1_positions):
                if h1_idx < child_idx:
                    heading_text = h1_txt
                    break
            # Only add if this sectPr maps to a DIFFERENT heading than the last one
            # (prevents duplicate entries when a section has no own H1, e.g. TOC)
            if heading_text and heading_text != last_heading:
                sectpr_map.append((phys_sectpr_idx, heading_text))
                last_heading = heading_text
            phys_sectpr_idx += 1

        sections_list = [h for _, h in sectpr_map]

        # ── 分类：front_matter vs body ──
        # 封面、郑重声明、摘要、ABSTRACT、目录 = front matter
        # 其余的 = body（空白页带页眉和页码）
        front_matter_norm = {
            "封面", "郑重声明", "郑 重 声 明", "郑重 声明",
            "摘要", "摘    要", "摘  要",
            "ABSTRACT",
            "目录", "目  录", "目 录",
            "保密协议", "保 密 协 议",  # 可能的前部页面
        }
        front_matter_headings = set()
        for h in sections_list:
            h_norm = re.sub(r'\s+', '', h)
            if h_norm in front_matter_norm:
                front_matter_headings.add(h_norm)

        config = EnforceConfig(
            libreoffice_path=Path(libreoffice_path),
            docx_input=src,
            docx_output=dst,
            odd_page_sections=sections_list,
            section_sectpr_map=sectpr_map,
            front_matter_headings=front_matter_headings,
        )

        logs.append(f"  章节: {sections_list[:3]}...共{len(sections_list)}节 "
                    f"(front_matter: {len(front_matter_headings)})")
        try:
            result = enforce_odd_pages(config)
            logs.append(f"  完成: {result}")
            return str(result)
        except Exception as e:
            logs.append(f"  失败: {e}")
            return str(e)

    def _reconcile_extra_sectpr(self, plan: list[dict]) -> None:
        with zipfile.ZipFile(self._target_path, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))

        body = root.find(f"{{{W}}}body")
        if body is None:
            return

        extra_indices = []
        for child in body:
            if child.tag != f"{{{W}}}p":
                continue
            pPr = child.find(f"{{{W}}}pPr")
            if pPr is None:
                continue
            if pPr.find(f"{{{W}}}sectPr") is not None:
                extra_indices.append(len(extra_indices))

        if extra_indices:
            plan.append({
                "action": "remove_extra_sectpr",
                "params": {"preserve_body_sectpr": True},
            })

    def _apply_user_overrides(self, plan: list[dict]) -> None:
        """将用户反馈持久化覆盖到 fix_plan"""
        if not self._user_overrides:
            return

        # 覆盖样式
        for key, override in [
            ("h1_font_east", ("h1", "font_eastAsia")),
            ("h1_font_ascii", ("h1", "font_ascii")),
            ("h1_size_pt", ("h1", "font_size_pt")),
            ("body_font_east", ("body", "font_eastAsia")),
            ("body_font_ascii", ("body", "font_ascii")),
            ("body_size_pt", ("body", "font_size_pt")),
        ]:
            if key in self._user_overrides:
                style_key, param_key = override
                style_id = self._style_ids.get(style_key, "1" if style_key == "h1" else "a8")
                val = self._user_overrides[key]

                # 查找已有 set_style action 或创建新的
                found = False
                for action in plan:
                    if action["action"] == "set_style" and action["params"].get("style_id") == style_id:
                        action["params"][param_key] = val
                        found = True
                        break
                if not found:
                    plan.append({
                        "action": "set_style",
                        "params": {"style_id": style_id, param_key: val},
                    })

        # 覆盖页眉字体
        if "header_font_east" in self._user_overrides or "header_font_ascii" in self._user_overrides:
            for action in plan:
                if action["action"] == "set_header_font":
                    if "header_font_east" in self._user_overrides:
                        action["params"]["font_eastAsia"] = self._user_overrides["header_font_east"]
                    if "header_font_ascii" in self._user_overrides:
                        action["params"]["font_ascii"] = self._user_overrides["header_font_ascii"]

    # ── 用户反馈闭环 ──────────────────────────────────────

    def handle_feedback(self, feedback_text: str) -> AgentResult:
        """处理用户反馈，更新 fix_plan 并重新执行 Writer → Fixer。

        反馈自动持久化到 .docmind_overrides.json，下次 run() 自动加载。
        """
        self._parse_feedback(feedback_text)
        self._save_overrides(reason=feedback_text)  # 持久化

        self._fix_plan = self._reconcile()

        # ── 重新执行 Writer（纯 OOXML）──
        logs = []
        if self._output_path != self._target_path:
            shutil.copy2(self._target_path, self._output_path)
        try:
            writer_logs = apply_fixes(self._fix_plan, self._output_path)
        except Exception as we:
            writer_logs = [f"[FAIL] Writer 异常: {we}"]

        logs.extend(writer_logs)

        fixed_count = sum(1 for l in writer_logs if l.startswith("[OK"))

        # ── 重新诊断 ──
        issues = diagnose(self._fix_plan, self._output_path)
        stats = fixer_summary(issues)

        logs.append(f"  Fixer: {stats['total']} issues")

        for line in logs:
            print(f"[feedback] {line}")

        return AgentResult(
            plan=self._fix_plan,
            issues=issues,
            fixed_count=fixed_count,
            remaining=len(issues),
            output_path=self._output_path,
            logs=logs,
        )

    def _parse_feedback(self, feedback_text: str) -> None:
        """解析用户反馈文本，更新 user_overrides。

        支持格式:
        - key=value 对（逗号分隔）
        - 自然语言关键词
        """
        import re

        # ── 尝试解析 key=value 格式 ──
        kv_pattern = re.compile(r'(\w+)\s*=\s*([^,]+)')
        kvs = kv_pattern.findall(feedback_text)

        if kvs:
            for k, v in kvs:
                v = v.strip().strip("'\"")
                # 尝试转换数值
                try:
                    v = float(v)
                    if v == int(v):
                        v = int(v)
                except ValueError:
                    pass
                self._user_overrides[k] = v
            return

        # ── 自然语言关键词匹配 ──
        text = feedback_text.lower()

        # 删除空白页
        if any(kw in text for kw in ["空白页", "删除空白页", "多余空白"]):
            # 找到 remove_extra_sectpr action 并确保其存在
            self._user_overrides["remove_blank_pages"] = True

        # 参考文献页眉
        if any(kw in text for kw in ["参考文献", "reference"]):
            if "页眉" in text:
                self._user_overrides["back_matter_header"] = "参考文献"

        # 致谢页眉
        if any(kw in text for kw in ["致谢", "acknowledge"]):
            if "页眉" in text:
                self._user_overrides["back_matter_header"] = "致谢"

        # 字体修改
        font_match = re.search(r'(?:改成|改为|修改为|用)\s*([\u4e00-\u9fff]{2,6})', text)
        if font_match and any(kw in text for kw in ["字体", "改成", "修改为", "用"]):
            font_name = font_match.group(1)
            if "页眉" in text:
                self._user_overrides["header_font_east"] = font_name
            elif "标题" in text:
                self._user_overrides["h1_font_east"] = font_name
            elif "正文" in text:
                self._user_overrides["body_font_east"] = font_name

        # 字号修改
        size_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:pt|号|磅)', text)
        if size_match:
            size_val = float(size_match.group(1))
            if "页眉" in text:
                self._user_overrides["header_size_pt"] = size_val
            elif "标题" in text:
                self._user_overrides["h1_size_pt"] = size_val
            elif "正文" in text:
                self._user_overrides["body_size_pt"] = size_val

        # 加粗
        if any(kw in text for kw in ["加粗", "粗体", "bold"]):
            if "标题" in text:
                self._user_overrides["h1_bold"] = True
            elif "正文" in text:
                self._user_overrides["body_bold"] = False

    # ── 便利方法 ──────────────────────────────────────────

    def preview_plan(self) -> list[str]:
        """预览当前 fix_plan（不执行）"""
        lines = []
        for i, action in enumerate(self._fix_plan):
            action_type = action.get("action", "?")
            params = action.get("params", {})
            lines.append(f"#{i}: {action_type} {_brief_action(action_type, params)}")
        return lines

    def get_overrides(self) -> dict:
        """获取当前用户覆盖"""
        return dict(self._user_overrides)

    def reset_overrides(self) -> None:
        """清除所有用户覆盖并删除持久化文件"""
        self._user_overrides.clear()
        self._feedback_rules.clear()
        try:
            override_path = self._target_path.parent / self.OVERRIDES_FILE if self._target_path else None
            if override_path and override_path.exists():
                override_path.unlink()
        except Exception:
            pass

    def _overrides_path(self) -> Optional[Path]:
        """返回持久化文件路径（与目标文档同目录）"""
        if self._target_path:
            return self._target_path.parent / self.OVERRIDES_FILE
        return None

    def _load_overrides(self) -> None:
        """从 .docmind_overrides.json 加载持久化规则"""
        path = self._overrides_path()
        if not path or not path.exists():
            return
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
            self._user_overrides = data.get('overrides', {})
            self._feedback_rules = [
                FeedbackRule(**r) for r in data.get('rules', [])
            ]
        except Exception:
            pass

    def _save_overrides(self, reason: str = "") -> None:
        """保存覆盖规则到 .docmind_overrides.json"""
        path = self._overrides_path()
        if not path:
            return
        now = __import__('datetime').datetime.now().isoformat()
        for k, v in self._user_overrides.items():
            existing = next((r for r in self._feedback_rules if r.key == k), None)
            if existing:
                existing.value = v
                existing.when = now
                if reason:
                    existing.reason = reason
            else:
                self._feedback_rules.append(FeedbackRule(key=k, value=v, reason=reason, when=now))
        data = {
            'overrides': self._user_overrides,
            'rules': [{'key': r.key, 'value': r.value, 'reason': r.reason, 'when': r.when}
                      for r in self._feedback_rules],
            'updated': now,
        }
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


# ── 辅助函数 ──────────────────────────────────────────────

def _copy_format_params(rules: dict, params: dict) -> None:
    """将 spec_rules 中的格式参数复制到 action params"""
    key_map = {
        "font_east": "font_eastAsia",
        "font_ascii": "font_ascii",
        "font_size_pt": "font_size_pt",
        "bold": "bold",
        "alignment": "alignment",
    }
    for spec_key, param_key in key_map.items():
        if spec_key in rules and rules[spec_key] is not None:
            params[param_key] = rules[spec_key]


def _brief_action(action_type: str, params: dict) -> str:
    """生成 action 的简短描述"""
    if action_type == "set_style":
        return f"style_id={params.get('style_id')} " + " ".join(
            f"{k}={v}" for k, v in params.items() if k != "style_id"
        )
    elif action_type == "set_sectpr_type":
        return f"sec[{params.get('section_index')}] → {params.get('val')}"
    elif action_type == "set_header_font":
        return f"{params.get('header_path')} font={params.get('font_eastAsia')}/{params.get('font_ascii')}"
    elif action_type == "add_page_number":
        return f"{params.get('footer_path')} format={params.get('format')}"
    elif action_type == "remove_extra_sectpr":
        return f"cleanup embedded sectPr"
    else:
        return str(params)
