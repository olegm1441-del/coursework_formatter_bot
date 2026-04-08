"""
Phase 3 infra — layout_render.py
Converts a .docx file to PDF via LibreOffice headless.

Usage:
    pdf_path = render_docx_to_pdf(docx_path)
    # pdf_path is a Path in the same temp dir, caller must clean up

Raises:
    LibreOfficeNotFoundError  — LO not installed, Phase 3 should be skipped
    RuntimeError              — conversion failed for other reasons
"""

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)


class LibreOfficeNotFoundError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# Locate soffice binary
# ---------------------------------------------------------------------------

_CANDIDATE_PATHS = [
    # Linux (apt / Railway)
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    # macOS (brew cask)
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
]


def _find_soffice() -> str:
    """Return path to soffice binary or raise LibreOfficeNotFoundError."""
    # 1. Check PATH first
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found

    # 2. Try known install locations
    for candidate in _CANDIDATE_PATHS:
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate

    raise LibreOfficeNotFoundError(
        "LibreOffice (soffice) not found. "
        "Install it: macOS → brew install --cask libreoffice; "
        "Linux → apt install libreoffice"
    )


# ---------------------------------------------------------------------------
# Conversion
# ---------------------------------------------------------------------------

def render_docx_to_pdf(docx_path: Path, timeout: int = 120) -> Path:
    """
    Convert docx_path to PDF using LibreOffice headless.

    Returns the path to the generated PDF file inside a fresh temp directory.
    The caller is responsible for deleting the temp directory when done.

    Raises LibreOfficeNotFoundError if LibreOffice is not installed.
    Raises RuntimeError if conversion fails.
    """
    docx_path = Path(docx_path)
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    soffice = _find_soffice()
    outdir = Path(tempfile.mkdtemp(prefix="lo_render_"))

    try:
        cmd = [
            soffice,
            "--headless",
            "--norestore",
            "--nofirststartwizard",
            "--convert-to", "pdf",
            "--outdir", str(outdir),
            str(docx_path),
        ]

        logger.info("layout_render: running %s", " ".join(cmd))

        result = subprocess.run(
            cmd,
            timeout=timeout,
            capture_output=True,
            text=True,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (rc={result.returncode}): "
                f"{result.stderr[:500]}"
            )

        # LibreOffice writes <stem>.pdf into outdir
        pdf_path = outdir / (docx_path.stem + ".pdf")
        if not pdf_path.exists():
            # Try any .pdf in outdir
            candidates = list(outdir.glob("*.pdf"))
            if not candidates:
                raise RuntimeError(
                    f"PDF not found in {outdir} after conversion. "
                    f"stdout={result.stdout[:300]}"
                )
            pdf_path = candidates[0]

        logger.info("layout_render: PDF created at %s (%d bytes)", pdf_path, pdf_path.stat().st_size)
        return pdf_path

    except subprocess.TimeoutExpired:
        shutil.rmtree(outdir, ignore_errors=True)
        raise RuntimeError(f"LibreOffice conversion timed out after {timeout}s")

    except LibreOfficeNotFoundError:
        shutil.rmtree(outdir, ignore_errors=True)
        raise

    except Exception:
        shutil.rmtree(outdir, ignore_errors=True)
        raise
