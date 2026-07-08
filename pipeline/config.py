"""
DocMind Pipeline — 统一流水线配置 v3

格式规则从规范模板自动提取（spec_reader）。
用户通过对话覆盖的优先。
"""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class PipelineConfig:
    """DocMind 格式流水线配置"""
    
    # ── 文件路径 ──
    input_docx: Path = Path('output/毕业设计.docx')
    spec_docx: Optional[Path] = Path('output/附件7：毕业设计（论文）撰写规范.docx')
    output_digital: Path = Path('output/毕业设计_电子版.docx')
    output_print: Path = Path('output/毕业设计_打印版.docx')
    
    # ── 模式 ──
    print_mode: bool = True
    libreoffice_path: str = 'C:/Program Files/LibreOffice/program/soffice.exe'
    
    # ── 格式规则（从规范模板自动填充，用户对话可覆盖） ──
    # 这些值会在 auto_discover() 中从 spec_docx 读取
    # 如果用户说"页顶空行用单倍行距"，设置 override_blank_line_spacing='auto'
    spec_rules: dict = field(default_factory=dict)
    
    # ── 用户覆盖（对话式修改） ──
    # 示例: override_blank_line_type = 'auto'  # 从固定值改为单倍行距
    override_h1_font: Optional[str] = None
    override_h1_size: Optional[float] = None
    override_blank_line_type: Optional[str] = None  # 'auto' | 'exact'
    
    # ── 节页眉映射 ──
    # 自动从文档结构推导：Front matter (摘要/ABSTRACT/目录) = 各自页眉
    # Body (1-7章) = 奇"XX年 毕业设计" 偶"专业名"
    # Back matter (参考文献/致谢) = 各自页眉
    # 用户可覆盖
    section_headers: dict = field(default_factory=dict)
    
    # 页眉文字模板
    odd_header_template: str = '{year}年 毕业设计'
    even_header_template: str = '计算机科学与技术'
    year: int = 2024
    
    # ── 页脚 ──
    footer_font_size_pt: float = 9  # 小五
    footer_format: str = '第{PAGE}页'
    
    # ── 样式 ID（从 auto_discover 自动填充，用户可覆盖） ──
    style_ids: dict = field(default_factory=dict)
    
    # ── 强制奇数页的节索引（None=全部） ──
    odd_page_section_indices: Optional[list] = None
