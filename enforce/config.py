"""
Config dataclass for odd-page enforcement.
"""
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class EnforceConfig:
    """Configuration for enforce_odd_pages().

    Attributes:
        libreoffice_path: Path to soffice.exe (LibreOffice).
        docx_input: Input .docx file to analyze and modify.
        docx_output: Output .docx file (may differ from input).
        odd_page_sections: List of section heading texts (in document order)
                           that must start on odd pages.
        section_sectpr_map: Optional list of (sectpr_index, heading) tuples.
                           When provided, bypasses PDF heading matching and
                           uses the given sectPr indices directly.
        front_matter_headings: Set of section headings that are front matter.
                              Blank pages before these sections get NO headers/page numbers.
                              All other sections' blank pages get body headers + page numbers.
        body_header_rId_odd: rId for the body odd-page header XML file.
        body_header_rId_even: rId for the body even-page header XML file.
        body_footer_rId: rId for the body footer XML file (with PAGE field).
        body_footer_rId_even: rId for the body even-page footer XML file.
    """
    libreoffice_path: Path
    docx_input: Path
    docx_output: Path
    odd_page_sections: list[str] = field(default_factory=list)
    section_sectpr_map: list[tuple[int, str]] = field(default_factory=list)

    # Blank page behavior
    front_matter_headings: set[str] = field(default_factory=set)

    # Header/footer rIds for body blank pages (discovered from existing docx sections)
    body_header_rId_odd: str = ""
    body_header_rId_even: str = ""
    body_footer_rId: str = ""
    body_footer_rId_even: str = ""

    # Page margin for blank page sections (discovered from existing body sections)
    pg_mar_top: str = "1418"
    pg_mar_right: str = "1134"
    pg_mar_bottom: str = "850"
    pg_mar_left: str = "1587"
    pg_mar_header: str = "1134"
    pg_mar_footer: str = "567"

    # Page number format for body (discovered from existing body sections)
    body_pg_num_fmt: str = "decimal"

    @property
    def has_body_header_config(self) -> bool:
        """Whether body blank page header/footer config is complete."""
        return bool(
            self.body_header_rId_odd
            and self.body_footer_rId
        )
