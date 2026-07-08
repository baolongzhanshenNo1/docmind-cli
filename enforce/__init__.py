"""
DocMind 2.0 — enforce/ package
Odd-page enforcement for thesis-style documents.
"""
from .config import EnforceConfig
from .pdf_reader import get_section_start_pages
from .odd_pages import enforce_odd_pages

__all__ = [
    "EnforceConfig",
    "get_section_start_pages",
    "enforce_odd_pages",
]
