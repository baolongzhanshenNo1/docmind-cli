"""
PDF reading utilities for section-page analysis.
"""
import re
import fitz
from pathlib import Path


def _normalize(text: str) -> str:
    """Remove all whitespace for fuzzy matching."""
    return re.sub(r'\s+', '', text)


def get_section_start_pages(pdf_path: Path, sections: list[str]) -> list[int | None]:
    """Return the starting page number (1-based) for each section heading.

    Uses LAST match (reverse search) with y=50-200 range only.
    - y=50-200: content headings on chapter start pages
    - y<50: headers (excluded)
    - y>200: TOC entries (excluded)
    No full-page fallback — it picks up header text.
    """
    doc = fitz.open(str(pdf_path))
    
    norm_headings = [_normalize(h) for h in sections]
    start_pages: list[int | None] = []
    
    for heading_norm in norm_headings:
        found: int | None = None

        # LAST match (reverse search) avoids TOC entries before chapter content
        for pn in range(doc.page_count - 1, -1, -1):
            page = doc[pn]
            blocks = page.get_text("blocks")
            
            for block in blocks:
                x0, y0, x1, y1, text, block_type, block_no = block
                text_norm = _normalize(text)
                if heading_norm in text_norm and (50 < y0 <= 200):
                    found = pn + 1
                    break
            if found:
                break
        
        start_pages.append(found)

    doc.close()
    return start_pages
