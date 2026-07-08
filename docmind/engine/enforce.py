# Engine wrapper for enforce module — kept at original location
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent.parent / '_archive' / 'docmind'))

from enforce.odd_pages import enforce_odd_pages
from enforce.config import EnforceConfig
from enforce.pdf_reader import get_section_start_pages
