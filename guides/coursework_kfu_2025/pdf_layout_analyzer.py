"""
Phase 3 infra — pdf_layout_analyzer.py
Extracts per-page text layout from a PDF produced by LibreOffice.

Returns a list of PageInfo objects — one per PDF page — each containing
the text blocks found on that page with their approximate vertical positions.

This is used by table_continuation.py and other Phase 3 rules to find:
  - where tables cross page boundaries
  - empty first lines of pages
  - orphaned captions / headings
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class TextBlock:
    """A contiguous block of text on a page."""
    text: str          # cleaned text content
    top: float         # distance from top of page in points
    bottom: float      # bottom edge in points
    page_num: int      # 1-based page number


@dataclass
class PdfWord:
    """One rendered PDF word with page-local coordinates."""
    text: str
    page_num: int
    x0: float
    x1: float
    top: float
    bottom: float


@dataclass
class PdfLine:
    """One rendered PDF text line with page-local coordinates."""
    text: str
    page_num: int
    top: float
    bottom: float


@dataclass
class PageInfo:
    """All text blocks found on one page."""
    page_num: int      # 1-based
    height: float      # page height in points
    width: float       # page width in points
    blocks: list[TextBlock] = field(default_factory=list)

    @property
    def first_block(self) -> TextBlock | None:
        return self.blocks[0] if self.blocks else None

    @property
    def last_block(self) -> TextBlock | None:
        return self.blocks[-1] if self.blocks else None

    def blocks_in_bottom_fraction(self, fraction: float = 0.15) -> list[TextBlock]:
        """Return blocks whose top starts in the last `fraction` of the page."""
        threshold = self.height * (1 - fraction)
        return [b for b in self.blocks if b.top >= threshold]

    def blocks_in_top_fraction(self, fraction: float = 0.10) -> list[TextBlock]:
        """Return blocks whose bottom ends in the first `fraction` of the page."""
        threshold = self.height * fraction
        return [b for b in self.blocks if b.bottom <= threshold]


# ---------------------------------------------------------------------------
# Regex helpers (reuse from classifier where possible)
# ---------------------------------------------------------------------------

_TABLE_CAPTION_RE = re.compile(
    r"^\s*(таблица|table)\s+\d+(?:\.\d+){0,2}",
    re.IGNORECASE,
)
_TABLE_CONTINUATION_RE = re.compile(
    r"^\s*(продолжение\s+таблицы|continuation\s+of\s+table)\b",
    re.IGNORECASE,
)
_FIGURE_CAPTION_RE = re.compile(
    r"^\s*(рис\.|рисунок|figure|fig\.)\s*\d+(?:\.\d+){0,2}",
    re.IGNORECASE,
)
_HEADING2_RE = re.compile(r"^\s*\d+\.\d+\.?\s+\S")


def _clean(text: str) -> str:
    text = text.replace("\xa0", " ").replace("\t", " ")
    return re.sub(r" {2,}", " ", text).strip()


def is_table_caption(text: str) -> bool:
    return bool(_TABLE_CAPTION_RE.match(_clean(text)))


def is_table_continuation(text: str) -> bool:
    return bool(_TABLE_CONTINUATION_RE.match(_clean(text)))


def is_figure_caption(text: str) -> bool:
    return bool(_FIGURE_CAPTION_RE.match(_clean(text)))


def is_heading2(text: str) -> bool:
    return bool(_HEADING2_RE.match(_clean(text)))


# ---------------------------------------------------------------------------
# Core extraction
# ---------------------------------------------------------------------------

def analyze_pdf(pdf_path: Path) -> list[PageInfo]:
    """
    Parse the PDF and return one PageInfo per page.

    Uses pdfplumber to extract words grouped into lines, then lines grouped
    into blocks by vertical proximity.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "pdfplumber is required for Phase 3. "
            "Install it: pip install pdfplumber"
        )

    pdf_path = Path(pdf_path)
    pages: list[PageInfo] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_info = PageInfo(
                page_num=page_num,
                height=float(page.height),
                width=float(page.width),
            )

            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
            )

            if not words:
                pages.append(page_info)
                continue

            # Group words into lines by shared top coordinate (y_tolerance=2pt)
            lines: list[list[dict]] = []
            current_line: list[dict] = []
            last_top: float | None = None

            for word in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
                top = word["top"]
                if last_top is None or abs(top - last_top) <= 2:
                    current_line.append(word)
                    last_top = top
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = [word]
                    last_top = top

            if current_line:
                lines.append(current_line)

            # Group lines into blocks by vertical gap (> 8pt = new block)
            LINE_GAP_THRESHOLD = 8.0
            blocks: list[list[list[dict]]] = []
            current_block: list[list[dict]] = []
            last_bottom: float | None = None

            for line in lines:
                line_top = min(w["top"] for w in line)
                line_bottom = max(w["bottom"] for w in line)

                if last_bottom is None or (line_top - last_bottom) <= LINE_GAP_THRESHOLD:
                    current_block.append(line)
                else:
                    if current_block:
                        blocks.append(current_block)
                    current_block = [line]

                last_bottom = line_bottom

            if current_block:
                blocks.append(current_block)

            # Convert blocks to TextBlock objects
            for block_lines in blocks:
                all_words = [w for line in block_lines for w in line]
                text = " ".join(
                    " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))
                    for line in block_lines
                )
                text = _clean(text)
                if not text:
                    continue

                top = min(w["top"] for w in all_words)
                bottom = max(w["bottom"] for w in all_words)

                page_info.blocks.append(TextBlock(
                    text=text,
                    top=float(top),
                    bottom=float(bottom),
                    page_num=page_num,
                ))

            pages.append(page_info)
            logger.debug(
                "pdf_analyzer: page %d → %d blocks", page_num, len(page_info.blocks)
            )

    logger.info("pdf_layout_analyzer: analyzed %d pages", len(pages))
    return pages


def analyze_pdf_lines(pdf_path: Path) -> list[PdfLine]:
    """
    Parse the PDF and return rendered text lines with page numbers.

    This is additive to analyze_pdf(); existing block extraction remains
    unchanged. Table splitting uses line-level data to map DOCX rows to the
    page where LibreOffice actually rendered them.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "pdfplumber is required for Phase 3. "
            "Install it: pip install pdfplumber"
        )

    pdf_path = Path(pdf_path)
    out: list[PdfLine] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
            )
            if not words:
                continue

            lines: list[list[dict]] = []
            current_line: list[dict] = []
            last_top: float | None = None

            for word in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
                top = float(word["top"])
                if last_top is None or abs(top - last_top) <= 2:
                    current_line.append(word)
                    last_top = top
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = [word]
                    last_top = top

            if current_line:
                lines.append(current_line)

            for line in lines:
                text = _clean(" ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"])))
                if not text:
                    continue
                out.append(PdfLine(
                    text=text,
                    page_num=page_num,
                    top=float(min(w["top"] for w in line)),
                    bottom=float(max(w["bottom"] for w in line)),
                ))

    logger.info("pdf_layout_analyzer: extracted %d lines", len(out))
    return out
