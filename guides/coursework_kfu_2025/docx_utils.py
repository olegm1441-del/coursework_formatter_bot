"""
Shared DOCX XML utilities used by Phase 2 and Phase 3 modules.

Centralises logic that would otherwise be duplicated between
safe_formatter.py and table_continuation.py / pagination_rules.py.
"""
from __future__ import annotations

import re

from docx.oxml.ns import qn


# ── Image detection ───────────────────────────────────────────────────────────

def xml_has_image(xml_elem) -> bool:
    """
    True if *xml_elem* (a raw lxml Element) contains an inline image:
    w:drawing (modern EMF/PNG) or w:pict (legacy VML).
    """
    return bool(
        xml_elem.findall(".//" + qn("w:drawing"))
        or xml_elem.findall(".//" + qn("w:pict"))
    )


# ── Source / note line detection ──────────────────────────────────────────────

_SOURCE_NOTE_RE = re.compile(
    r"^\s*(источник|примечание|составлено по|рассчитано по)\s*:",
    re.IGNORECASE,
)


def is_source_or_note_line(text: str) -> bool:
    """True if *text* starts with 'Источник:' or 'Примечание:' (case-insensitive)."""
    return bool(_SOURCE_NOTE_RE.match(text))


# ── Formatting report ─────────────────────────────────────────────────────────

class FormattingReport:
    """
    Collects human-readable warnings produced during Phase 3.

    Each warning is a short Russian string (≤80 chars) suitable for
    appending to the Telegram bot caption.

    Usage::
        report = FormattingReport()
        report.warn("При переносе таблицы 1.2 осталось мало строк")
        if not report.is_empty():
            print(report.format_caption())
    """

    def __init__(self) -> None:
        self._warnings: list[str] = []

    def warn(self, message: str) -> None:
        """Append a warning (idempotent — duplicate messages are allowed)."""
        self._warnings.append(message)

    @property
    def warnings(self) -> list[str]:
        """Return a snapshot of the current warning list."""
        return list(self._warnings)

    def is_empty(self) -> bool:
        return not self._warnings

    def format_caption(self) -> str:
        """
        Return all warnings as a newline-separated block, prefixed with ⚠️.
        Example::
            ⚠️ При переносе таблицы 1.1 осталось мало строк
            ⚠️ В конце таблицы 2.3 пришлось сделать перенос источника
        """
        return "\n".join(f"⚠️ {w}" for w in self._warnings)
