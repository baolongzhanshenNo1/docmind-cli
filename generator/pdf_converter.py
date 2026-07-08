"""
PDF conversion with Word COM (primary) + LibreOffice fallback.
Word COM renders with 100% fidelity to what the user sees in Word.
"""
import subprocess
import shutil
from pathlib import Path
from typing import Optional

try:
    import pythoncom
    import win32com.client
    HAS_WORD_COM = True
except ImportError:
    HAS_WORD_COM = False


def _convert_via_word_com(docx_path: Path, pdf_path: Path, timeout: int = 60) -> bool:
    """Convert docx to PDF using Word COM. Returns True on success."""
    if not HAS_WORD_COM:
        return False
    
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        try:
            doc.ExportAsFixedFormat(
                str(pdf_path.resolve()),
                ExportFormat=17,  # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=0,    # wdExportOptimizeForPrint
                CreateBookmarks=1  # wdExportCreateHeadingBookmarks
            )
            return True
        finally:
            doc.Close(SaveChanges=0)
    except Exception:
        return False
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _convert_via_libreoffice(docx_path: Path, pdf_dir: Path, timeout: int = 60) -> bool:
    """Convert docx to PDF using LibreOffice headless. Returns True on success."""
    # Try common LibreOffice paths
    candidates = [
        shutil.which("soffice"),
        Path("D:/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        Path("/usr/bin/soffice"),
    ]
    
    soffice = None
    for c in candidates:
        if c and Path(c).exists():
            soffice = c
            break
    
    if soffice is None:
        return False
    
    try:
        subprocess.run(
            [str(soffice), "--headless", "--convert-to", "pdf",
             "--outdir", str(pdf_dir), str(docx_path)],
            capture_output=True, timeout=timeout, check=True
        )
        # LibreOffice names the output <basename>.pdf beside the docx
        # but --outdir should place it in pdf_dir. Handle both cases.
        expected = pdf_dir / f"{docx_path.stem}.pdf"
        if expected.exists():
            return True
        # Maybe it ended up next to the docx
        alt = docx_path.parent / f"{docx_path.stem}.pdf"
        if alt.exists():
            shutil.move(str(alt), str(expected))
            return True
        return False
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return False


def convert_to_pdf(docx_path: str | Path,
                   pdf_path: Optional[str | Path] = None,
                   timeout: int = 60) -> Path:
    """
    Convert a .docx file to PDF.
    
    Primary: Word COM (100% rendering fidelity)
    Fallback: LibreOffice headless
    
    Args:
        docx_path: Path to the .docx file.
        pdf_path: Optional output path. Default: same dir, same stem, .pdf.
        timeout: Conversion timeout in seconds.
    
    Returns:
        Path to the generated PDF.
    
    Raises:
        RuntimeError: If both Word COM and LibreOffice fail.
    """
    docx_path = Path(docx_path)
    if pdf_path is None:
        pdf_path = docx_path.parent / f"{docx_path.stem}.pdf"
    pdf_path = Path(pdf_path)
    
    # Try Word COM first (best fidelity)
    if _convert_via_word_com(docx_path, pdf_path, timeout):
        return pdf_path
    
    # Fallback to LibreOffice
    if _convert_via_libreoffice(docx_path, pdf_path.parent, timeout):
        # LibreOffice may produce a different filename
        actual = pdf_path.parent / f"{docx_path.stem}.pdf"
        if actual.exists() and actual != pdf_path:
            shutil.move(str(actual), str(pdf_path))
        if pdf_path.exists():
            return pdf_path
    
    raise RuntimeError(
        f"Failed to convert {docx_path} to PDF. "
        "Neither Word COM nor LibreOffice is available. "
        "Install Microsoft Word or LibreOffice."
    )
