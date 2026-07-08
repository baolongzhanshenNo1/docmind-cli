"""
DocMind Pipeline v2 — 自动发现 + 统一排版流水线

流水线步骤:
    0. auto_discover  — Reader 自动发现样式/节/标题
    1. format_body    — 样式/段落/页顶空行
    2. setup_headers  — 页眉/页码/页边距
    3. enforce_print  — 打印版奇数页强制

用法:
    from pipeline import PipelineConfig, run_pipeline
    config = PipelineConfig(input_docx=Path('论文.docx'))
    run_pipeline(config)
"""
import shutil
from pathlib import Path

from .config import PipelineConfig

# 延迟导入，避免缺失可选依赖时阻塞其他模块（如 agent.py）
_defers = ['auto_discover', 'format_body', 'setup_headers', 'enforce_print']


def _get_deferred(name):
    import importlib
    return importlib.import_module(f'.{name}', __package__)


def run_pipeline(config: PipelineConfig):
    """执行完整排版流水线"""
    from . import auto_discover, format_body, setup_headers
    import shutil
    if not config.input_docx.exists():
        raise FileNotFoundError(f'Input file not found: {config.input_docx}')
    
    # ── Step 0: 自动发现文档结构 ──
    print('=' * 50)
    print('PIPELINE: Auto-Discovery')
    print('=' * 50)
    config = auto_discover.auto_discover(config)
    
    # ── 电子版 ──
    print()
    print('=' * 50)
    print('PIPELINE: Digital Version')
    print('=' * 50)
    
    if config.output_digital != config.input_docx:
        shutil.copy(config.input_docx, config.output_digital)
    
    digital_config = PipelineConfig(
        input_docx=config.output_digital,
        print_mode=False,
        section_headers=config.section_headers,
        style_ids=config.style_ids,
        **{k: v for k, v in config.__dict__.items() 
           if k not in ('input_docx','print_mode','output_digital','output_print',
                       'section_headers','style_ids')}
    )
    
    format_body.run(digital_config)
    setup_headers.run(digital_config)
    print(f'Digital: {digital_config.input_docx}')
    
    if not config.print_mode:
        print('\nPipeline complete!')
        return
    
    # ── 打印版 ──
    print()
    print('=' * 50)
    print('PIPELINE: Print Version')
    print('=' * 50)
    
    shutil.copy(digital_config.input_docx, config.output_print)
    
    print_config = PipelineConfig(
        input_docx=config.output_print,
        print_mode=True,
        section_headers=config.section_headers,
        style_ids=config.style_ids,
        **{k: v for k, v in config.__dict__.items() 
           if k not in ('input_docx','print_mode','output_digital','output_print',
                       'section_headers','style_ids')}
    )
    
    from . import enforce_print
    enforce_print.run(print_config)
    print(f'Print: {print_config.input_docx}')
    print('\nPipeline complete!')
