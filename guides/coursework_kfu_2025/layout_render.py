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

import hashlib
import logging
import os
import shutil
import signal
import subprocess
import tempfile
from collections import OrderedDict
from pathlib import Path

logger = logging.getLogger(__name__)


class LibreOfficeNotFoundError(RuntimeError):
    pass


_TRUTHY = {"1", "true", "yes", "on"}


def _flag_on(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in _TRUTHY


def _render_timeout_default() -> int:
    """Per-render wall-clock cap. A legitimate convert is <5s idle; this only
    fences a hung soffice. Override with KPFU_RENDER_TIMEOUT_S."""
    try:
        return max(15, int(os.environ.get("KPFU_RENDER_TIMEOUT_S", "60")))
    except (TypeError, ValueError):
        return 60


# ---------------------------------------------------------------------------
# #2 content-hash render cache — the cleanup probes render/verify the SAME docx
# bytes repeatedly (probe -> verify -> re-probe -> final gate). Memoize the PDF
# bytes keyed by the docx content hash and re-materialize into a fresh temp dir
# on a hit, preserving the "caller deletes its own temp dir" contract.
# ---------------------------------------------------------------------------
_RENDER_CACHE: "OrderedDict[str, bytes]" = OrderedDict()
_RENDER_CACHE_MAX = 12
_render_cache_stats = {"hits": 0, "misses": 0}


def clear_render_cache() -> None:
    _RENDER_CACHE.clear()


def render_cache_stats() -> dict:
    return dict(_render_cache_stats)


def _cache_get(key: str) -> bytes | None:
    data = _RENDER_CACHE.get(key)
    if data is not None:
        _RENDER_CACHE.move_to_end(key)
    return data


def _cache_put(key: str, data: bytes) -> None:
    _RENDER_CACHE[key] = data
    _RENDER_CACHE.move_to_end(key)
    while len(_RENDER_CACHE) > _RENDER_CACHE_MAX:
        _RENDER_CACHE.popitem(last=False)


# ---------------------------------------------------------------------------
# #1 warm LibreOffice user profile — first-run profile creation is a large slice
# of soffice cold-start. A fixed, pre-warmed UserInstallation dir is reused by
# every (sequential) render so that cost is paid once per process, not per call.
# ---------------------------------------------------------------------------
_WARM_PROFILE_DIR: str | None = None


def _warm_profile_arg() -> str | None:
    global _WARM_PROFILE_DIR
    if not _flag_on("KPFU_LO_WARM_PROFILE", True):
        return None
    if _WARM_PROFILE_DIR is None:
        _WARM_PROFILE_DIR = tempfile.mkdtemp(prefix="lo_warm_profile_")
    # LibreOffice wants a file:// URI for UserInstallation
    return f"-env:UserInstallation=file://{_WARM_PROFILE_DIR}"


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

def _materialize(pdf_bytes: bytes, stem: str) -> Path:
    outdir = Path(tempfile.mkdtemp(prefix="lo_render_"))
    pdf_path = outdir / f"{stem}.pdf"
    pdf_path.write_bytes(pdf_bytes)
    return pdf_path


def _run_soffice(cmd: list[str], timeout: int) -> subprocess.CompletedProcess:
    """Run soffice in its own session and, on timeout, kill the WHOLE process
    group — soffice forks a persistent soffice.bin that a plain child-kill
    leaves orphaned (and holding the profile lock, wedging later renders)."""
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, start_new_session=True,
    )
    try:
        out, err = proc.communicate(timeout=timeout)
        return subprocess.CompletedProcess(cmd, proc.returncode, out, err)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
        except (ProcessLookupError, PermissionError):
            proc.kill()
        try:
            proc.communicate(timeout=10)
        except Exception:
            pass
        raise


def render_docx_to_pdf(docx_path: Path, timeout: int | None = None) -> Path:
    """
    Convert docx_path to PDF using LibreOffice headless.

    Returns the path to the generated PDF file inside a fresh temp directory.
    The caller is responsible for deleting the temp directory when done.

    Renders are memoized by docx content hash (see _RENDER_CACHE) and use a warm,
    reusable LibreOffice profile; a hung soffice is fenced by a per-render timeout
    and killed by process group. Raises LibreOfficeNotFoundError if LibreOffice is
    not installed. Raises RuntimeError if conversion fails.
    """
    docx_path = Path(docx_path)
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    if timeout is None:
        timeout = _render_timeout_default()

    cache_on = _flag_on("KPFU_RENDER_CACHE", True)
    key = None
    if cache_on:
        key = hashlib.sha256(docx_path.read_bytes()).hexdigest()
        cached = _cache_get(key)
        if cached is not None:
            _render_cache_stats["hits"] += 1
            logger.info("layout_render: cache HIT %s", docx_path.name)
            return _materialize(cached, docx_path.stem)
        _render_cache_stats["misses"] += 1

    soffice = _find_soffice()
    outdir = Path(tempfile.mkdtemp(prefix="lo_render_"))

    try:
        cmd = [soffice, "--headless", "--norestore", "--nofirststartwizard"]
        warm = _warm_profile_arg()
        if warm:
            cmd.append(warm)
        cmd += ["--convert-to", "pdf", "--outdir", str(outdir), str(docx_path)]

        logger.info("layout_render: running %s", " ".join(cmd))

        result = _run_soffice(cmd, timeout)

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (rc={result.returncode}): "
                f"{(result.stderr or '')[:500]}"
            )

        # LibreOffice writes <stem>.pdf into outdir
        pdf_path = outdir / (docx_path.stem + ".pdf")
        if not pdf_path.exists():
            # Try any .pdf in outdir
            candidates = list(outdir.glob("*.pdf"))
            if not candidates:
                raise RuntimeError(
                    f"PDF not found in {outdir} after conversion. "
                    f"stdout={(result.stdout or '')[:300]}"
                )
            pdf_path = candidates[0]

        logger.info("layout_render: PDF created at %s (%d bytes)", pdf_path, pdf_path.stat().st_size)
        if cache_on and key is not None:
            try:
                _cache_put(key, pdf_path.read_bytes())
            except OSError:
                pass
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
