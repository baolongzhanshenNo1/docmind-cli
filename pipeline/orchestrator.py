"""
AgentOrchestrator — 五层闭环编排

Pipeline → Reader → Reconciler → Writer → Fixer

用法:
    orchestrator = AgentOrchestrator(PipelineConfig(...))
    result = orchestrator.run()
"""
import shutil, sys
from pathlib import Path

# Archived modules (docmind3) path

from .config import PipelineConfig
from .preprocessor import preprocess
from docmind3 import DocxReader, Reconciler
from docmind3.writer import DocxWriter


class AgentOrchestrator:
    """编排完整的排版流水线"""
    
    def __init__(self, config: PipelineConfig):
        self.config = config
        self.doc = None
        self.issues_fixed = 0
    
    def run(self) -> dict:
        """执行完整流水线，返回结果字典"""
        result = {'phases': {}, 'success': False}
        
        # ── Phase 1: Preprocess (ZIP 级别初始化) ──
        print('=' * 60)
        print('ORCHESTRATOR: Phase 1 — Preprocess')
        print('=' * 60)
        
        preprocessed = preprocess(
            self.config.input_docx,
            self.config.spec_docx,
            output_path=self.config.output_digital,
        )
        result['phases']['preprocess'] = str(preprocessed)
        
        # ── Phase 2: Reader (理解结构) ──
        print()
        print('=' * 60)
        print('ORCHESTRATOR: Phase 2 — Reader')
        print('=' * 60)
        
        reader = DocxReader()
        self.doc = reader.read(str(preprocessed))
        
        result['phases']['reader'] = {
            'sections': len(self.doc.sections),
            'headings': len(self.doc.all_headings),
            'h1_count': sum(1 for h in self.doc.all_headings if h.level == 1),
            'h2_count': sum(1 for h in self.doc.all_headings if h.level == 2),
            'h3_count': sum(1 for h in self.doc.all_headings if h.level == 3),
        }
        print(f"  {result['phases']['reader']}")
        
        # ── Phase 3: Reconciler (NL → OOXML 映射) ──
        print()
        print('=' * 60)
        print('ORCHESTRATOR: Phase 3 — Reconciler')
        print('=' * 60)
        
        reconciler = Reconciler(self.doc)
        reconciler.set_header_rules(
            front_matter=lambda h: h.text,
            body_odd=self.config.odd_header_template.format(year=self.config.year),
            body_even=self.config.even_header_template,
            back_matter=lambda h: h.text,
        )
        reconciler.set_footer_rules(
            front_format='upperRoman',
            body_format='decimal',
            front_start=1,
            body_start=1,
        )
        
        result['phases']['reconciler'] = {
            'header_rules': len(reconciler.header_rules),
            'footer_rules': len(reconciler.footer_rules),
        }
        print(f"  {result['phases']['reconciler']}")
        
        # ── Phase 4: Writer (执行修改) ──
        print()
        print('=' * 60)
        print('ORCHESTRATOR: Phase 4 — Writer')
        print('=' * 60)
        
        reconciler.apply(preprocessed)
        result['phases']['writer'] = str(preprocessed)
        print(f"  Applied to: {preprocessed}")
        
        # ── Phase 5: Fixer (诊断) ──
        print()
        print('=' * 60)
        print('ORCHESTRATOR: Phase 5 — Fixer')
        print('=' * 60)
        
        from docmind3 import DocxFixer
        fixer = DocxFixer(self.doc)
        self.doc = reader.read(str(preprocessed))
        fixer.doc = self.doc
        try:
            issues = fixer.diagnose()
        except:
            issues = []
        
        result['phases']['fixer_before'] = len(issues)
        print(f"  Issues found: {len(issues)}")
        
        # ── Phase 5b: Auto-fix ──
        if issues:
            writer = DocxWriter(self.doc)
            logs = fixer.fix(issues, writer)
            writer.save(str(preprocessed))
            # Re-diagnose
            self.doc = reader.read(str(preprocessed))
            fixer.doc = self.doc
            remaining = fixer.diagnose()
            result['phases']['fixer'] = {
                'issues_before': len(issues),
                'issues_after': len(remaining),
                'fix_logs': logs[:5],
            }
            print(f"  Fixed: {len(issues) - len(remaining)} issues remaining: {len(remaining)}")
        else:
            result['phases']['fixer'] = {'issues': 0}
        
        result['success'] = True
        result['output'] = str(preprocessed)
        
        return result


def run_pipeline(config: PipelineConfig) -> dict:
    """便捷入口"""
    orch = AgentOrchestrator(config)
    return orch.run()
