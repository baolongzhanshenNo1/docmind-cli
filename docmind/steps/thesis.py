"""论文专属排版步骤。"""

from pathlib import Path
from docmind.steps.base import PipelineStep


class RomanArabicStep(PipelineStep):
    """前部→罗马数字，正文→阿拉伯数字。"""

    def run(self, target: Path, spec: dict) -> list[dict]:
        # 前部 = 摘要/ABSTRACT/目录（section 2-4），正文 = section 5+
        # 简化：按索引分配
        roman_sections = [2, 3, 4]  # 摘要, ABSTRACT, 目录
        arabic_start = 5            # 第1章

        plan = [
            {"action": "set_page_number_type", "params": {
                "section_index": 2, "fmt": "upperRoman", "start": 1}},
        ]
        for si in [3, 4]:
            plan.append({"action": "set_page_number_type", "params": {
                "section_index": si, "fmt": "upperRoman"}})
        plan.append({"action": "set_page_number_type", "params": {
            "section_index": arabic_start, "fmt": "decimal", "start": 1}})

        # 后续 body 节
        import zipfile
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with zipfile.ZipFile(target, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))
        sect_count = len(list(root.iter(f"{{{W}}}sectPr")))

        for si in range(arabic_start + 1, sect_count + 1):
            plan.append({"action": "set_page_number_type", "params": {
                "section_index": si, "fmt": "decimal"}})

        return plan
