#!/usr/bin/env python3
"""Deterministic openpyxl Precedent Matrix Workbook compiler and standalone CLI.

Requires ONLY the Python standard library and ``openpyxl`` so this script runs
identically inside the Slinky repository, an installed Claude plugin directory,
Claude Desktop's sandbox, or ChatGPT Code Interpreter without ``src.*`` on
``PYTHONPATH``.
"""

from __future__ import annotations

import argparse
import io
import json
import math
import re
import sys
import zipfile
from abc import ABC, abstractmethod
from collections import Counter
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

# ── Institutional Styling Constants ──────────────────────────────────────────

FONT_FAMILY = "Calibri"

COLOR_TITLE = "1A1A1A"
COLOR_SUBTITLE = "5A5A5A"
COLOR_NAVY_HEADER_BG = "1F3864"
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_SUBSECTION_BG = "F2F4F8"
COLOR_ZEBRA_BG = "FAFBFC"
COLOR_WHITE_BG = "FFFFFF"

COLOR_INPUT_BLUE = "0000FF"
COLOR_FORMULA_BLACK = "1A1A1A"
COLOR_CROSS_SHEET_GREEN = "008000"
COLOR_HYPERLINK_BLUE = "0B5394"

COLOR_BORDER_SUBTLE = "E0E0E0"
COLOR_BORDER_SUMMARY = "9A9A9A"

FMT_CURRENCY_M = '$#,##0.0"M"'
FMT_PERCENT = "0.0%"
FMT_BPS_DELTA = '+0" bps";-0" bps";0" bps"'
FMT_INT = "#,##0"

DETERMINISTIC_CREATED_AT = datetime(2026, 1, 1, 0, 0, 0)
_DETERMINISTIC_ZIP_TIMESTAMP = (2026, 1, 1, 0, 0, 0)
_DETERMINISTIC_ISO_UTC = "2026-01-01T00:00:00Z"

_SNAKE_TOKEN_MAP: dict[str, str] = {
    "m_and_a": "M&A",
    "r_and_d": "R&D",
    "sg_and_a": "SG&A",
    "fp_and_a": "FP&A",
}
_ACRONYM_WORDS: dict[str, str] = {
    "ai": "AI",
    "api": "API",
    "arr": "ARR",
    "b2b": "B2B",
    "b2c": "B2C",
    "bps": "bps",
    "capex": "CapEx",
    "ceo": "CEO",
    "cfo": "CFO",
    "ebitda": "EBITDA",
    "erp": "ERP",
    "eps": "EPS",
    "eu": "EU",
    "ev": "EV",
    "fda": "FDA",
    "fcf": "FCF",
    "ftc": "FTC",
    "fy": "FY",
    "gaap": "GAAP",
    "gpu": "GPU",
    "ip": "IP",
    "ipo": "IPO",
    "ir": "IR",
    "it": "IT",
    "jv": "JV",
    "kpi": "KPI",
    "llm": "LLM",
    "ma": "M&A",
    "opex": "OpEx",
    "pe": "PE",
    "plg": "PLG",
    "q1": "Q1",
    "q2": "Q2",
    "q3": "Q3",
    "q4": "Q4",
    "rd": "R&D",
    "roi": "ROI",
    "roic": "ROIC",
    "saas": "SaaS",
    "sec": "SEC",
    "sic": "SIC",
    "sku": "SKU",
    "smb": "SMB",
    "uk": "UK",
    "us": "US",
    "usa": "USA",
}

_EMBEDDED_SNAKE_RE = re.compile(r"(?<![/\w.-])[a-z][a-z0-9]*(?:_[a-z0-9]+)+(?![/\w.-])")


def _humanize_snake_match(match: re.Match[str]) -> str:
    token = match.group(0).lower()
    if token in _SNAKE_TOKEN_MAP:
        return _SNAKE_TOKEN_MAP[token]
    normalized = token
    for compound, replacement in _SNAKE_TOKEN_MAP.items():
        normalized = normalized.replace(compound, replacement)
    parts = [_ACRONYM_WORDS.get(part, part.capitalize()) for part in normalized.split("_") if part]
    return " ".join(parts) or match.group(0)


def humanize_label(value: Any) -> str | None:
    """Convert standalone or embedded snake_case tokens into clean Title Case English."""
    cleaned = clean_optional_str(value)
    if cleaned is None:
        return None
    lower = cleaned.lower()
    if lower in _SNAKE_TOKEN_MAP:
        return _SNAKE_TOKEN_MAP[lower]
    return _EMBEDDED_SNAKE_RE.sub(_humanize_snake_match, cleaned)


@dataclass(frozen=True)
class BadgePalette:
    """Foreground and fill hex pair for semantic institutional pills."""

    fill_hex: str
    font_hex: str
    bold: bool = True


BADGE_EXPANDED = BadgePalette(fill_hex="EDF7EF", font_hex="1E6B31")
BADGE_ADVERSE = BadgePalette(fill_hex="FDF2F4", font_hex="B00020")
BADGE_PIVOT = BadgePalette(fill_hex="FEF9EC", font_hex="8A6100")
BADGE_OUTSIDE_DOMAIN = BadgePalette(fill_hex="EEF3FA", font_hex="1F3864")
BADGE_NEUTRAL = BadgePalette(fill_hex="F2F4F8", font_hex="334155", bold=False)
BADGE_DELISTED = BadgePalette(fill_hex="F2F4F8", font_hex="5A5A5A", bold=False)

DELISTED_TICKER_LABEL = "DELISTED"

OUTCOME_BADGES: dict[str, BadgePalette] = {
    "Expanded / Succeeded": BADGE_EXPANDED,
    "Success": BADGE_EXPANDED,
    "Succeeded": BADGE_EXPANDED,
    "Expanded": BADGE_EXPANDED,
    "Failed": BADGE_ADVERSE,
    "Failure": BADGE_ADVERSE,
    "Abandoned": BADGE_ADVERSE,
    "Failed / Abandoned": BADGE_ADVERSE,
    "Pivoted": BADGE_PIVOT,
    "Pivot": BADGE_PIVOT,
    "Merged": BADGE_NEUTRAL,
    "Ongoing": BADGE_NEUTRAL,
}

DOMAIN_BADGES: dict[str, BadgePalette] = {
    "Outside-Industry": BADGE_OUTSIDE_DOMAIN,
    "Direct Peer": BADGE_NEUTRAL,
}

STITCH_BADGES: dict[str, BadgePalette] = {
    "Mapped": BADGE_EXPANDED,
    "Awaiting Stitch": BADGE_PIVOT,
}

VERDICT_BADGES: dict[str, BadgePalette] = {
    "Supportive": BADGE_EXPANDED,
    "Aligned": BADGE_EXPANDED,
    "Favorable": BADGE_EXPANDED,
    "Confirmed": BADGE_EXPANDED,
    "Pass": BADGE_EXPANDED,
    "Holds": BADGE_EXPANDED,
    "Adverse": BADGE_ADVERSE,
    "Broken": BADGE_ADVERSE,
    "Breaks": BADGE_ADVERSE,
    "Unfavorable": BADGE_ADVERSE,
    "At Risk": BADGE_ADVERSE,
    "Violated": BADGE_ADVERSE,
    "Fail": BADGE_ADVERSE,
    "Mixed": BADGE_PIVOT,
    "Watch": BADGE_PIVOT,
    "Conditional": BADGE_PIVOT,
    "Unverified": BADGE_PIVOT,
    "Partial": BADGE_PIVOT,
}

OUTCOME_CANONICAL_ROWS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Expanded / Succeeded", ("Expanded / Succeeded", "Success", "Succeeded", "Expanded")),
    ("Pivoted", ("Pivoted", "Pivot")),
    ("Failed", ("Failed", "Failure")),
    ("Abandoned", ("Abandoned",)),
    ("Merged", ("Merged", "Acquired")),
    ("Ongoing", ("Ongoing",)),
)


class CellRole(Enum):
    """Wall Street font color role for cell values and formulas."""

    BODY_TEXT = COLOR_TITLE
    INPUT_BLUE = COLOR_INPUT_BLUE
    SAME_SHEET_FORMULA = COLOR_FORMULA_BLACK
    CROSS_SHEET_FORMULA = COLOR_CROSS_SHEET_GREEN
    HYPERLINK = COLOR_HYPERLINK_BLUE


# ── Pure Boundary Normalizers & Value Helpers ────────────────────────────────


def clean_optional_str(value: Any) -> str | None:
    """Return stripped non-empty string or ``None`` (never empty string)."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def optional_number(value: Any) -> int | float | None:
    """Preserve numeric ``0``/``0.0`` while keeping ``None``/NaN/blank as ``None``."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    text = str(value).strip()
    if not text:
        return None
    try:
        parsed = float(text)
    except ValueError:
        return None
    if not math.isfinite(parsed):
        return None
    return int(parsed) if parsed.is_integer() and "." not in text else parsed


def optional_int(value: Any) -> int | None:
    """Return integer when present and numeric, otherwise ``None``."""
    num = optional_number(value)
    return int(round(num)) if num is not None else None


def format_sic_industry(row: Mapping[str, Any]) -> str | None:
    """Canonical ``Industry (SIC)`` cell string shared across Matrix & Sector tabs."""
    industry = humanize_label(row.get("sic_industry"))
    if industry is not None:
        return industry
    sic_code = clean_optional_str(row.get("sic_code"))
    if sic_code is not None:
        return f"SIC {sic_code}"
    return "Unclassified Industry" if bool(row) else None


def format_ticker_display(row: Mapping[str, Any]) -> str | None:
    """Return active ticker symbol or ``DELISTED`` for legacy/delisted CIK records."""
    ticker = clean_optional_str(row.get("ticker"))
    if ticker is not None:
        return DELISTED_TICKER_LABEL if ticker.upper() == DELISTED_TICKER_LABEL else ticker
    company = clean_optional_str(row.get("company"))
    if company is not None and company.lower() != "unknown":
        return DELISTED_TICKER_LABEL
    return None


def _escape_excel_string(text: str, *, max_len: int = 240) -> str:
    compact = " ".join(text.splitlines()).strip()
    clipped = compact[:max_len]
    return clipped.replace('"', '""')


def build_hyperlink_formula(
    url: Any, label: Any, *, fallback_label: str = "SEC Filing"
) -> str | None:
    """Build Excel ``=HYPERLINK("url","label")`` when a URL is present."""
    clean_url = clean_optional_str(url)
    if clean_url is None:
        return None
    display = clean_optional_str(label) or fallback_label
    safe_url = _escape_excel_string(clean_url, max_len=2000)
    safe_label = _escape_excel_string(display, max_len=240)
    return f'=HYPERLINK("{safe_url}","{safe_label}")'


def primary_filing_display_label(row: Mapping[str, Any]) -> str | None:
    """Human-readable label for Precedent Matrix Column U (Primary SEC Filing)."""
    form = clean_optional_str(row.get("primary_filing_form"))
    fdate = clean_optional_str(row.get("primary_filing_date"))
    accession = clean_optional_str(row.get("primary_accession"))
    if form and fdate:
        return f"{form} ({fdate})"
    return form or accession or fdate


def wire_precedent_count(row: Mapping[str, Any]) -> int | None:
    """Extract cohort precedent count across wire schema casing aliases."""
    for key in ("precedent_count", "precedent_Count", "cohort_arc_count"):
        val = optional_int(row.get(key))
        if val is not None:
            return val
    return None


def lookup_case_insensitive_count(
    counts: Mapping[str, Any] | None,
    aliases: Sequence[str],
    *,
    has_cohort_pool: bool,
) -> int | None:
    """Resolve full-cohort outcome count across backend label aliases."""
    if not counts:
        return 0 if has_cohort_pool else None
    lowered = {str(k).strip().lower(): optional_int(v) for k, v in counts.items()}
    total = 0
    matched = False
    for alias in aliases:
        candidate = lowered.get(alias.lower())
        if candidate is not None:
            total += candidate
            matched = True
    if matched:
        return total
    return 0 if has_cohort_pool else None


# ── Normalized Immutable Payload Model ───────────────────────────────────────


@dataclass(frozen=True)
class WorkbookPayload:
    """Validated, normalized Precedent Matrix Workbook input contract."""

    cohort_header: Mapping[str, Any]
    anchor_context: Mapping[str, Any] | None
    precedent_rows: tuple[Mapping[str, Any], ...]
    filing_trail_rows: tuple[Mapping[str, Any], ...]
    condition_rows: tuple[Mapping[str, Any], ...]
    wire_rows: tuple[Mapping[str, Any], ...]

    @property
    def matrix_last_row(self) -> int:
        """Last Precedent Matrix data row index (Row 5 minimum for formula ranges)."""
        return max(5, 4 + len(self.precedent_rows))


def _coerce_dict_list(raw: Any) -> tuple[Mapping[str, Any], ...]:
    if not isinstance(raw, Sequence) or isinstance(raw, (str, bytes)):
        return ()
    return tuple(item for item in raw if isinstance(item, Mapping))


def _extract_condition_rows(
    raw_payload: Mapping[str, Any],
    matrix_block: Mapping[str, Any],
    override_conditions: Sequence[Mapping[str, Any]] | Mapping[str, Any] | None,
) -> tuple[Mapping[str, Any], ...]:
    if isinstance(override_conditions, Mapping):
        return _coerce_dict_list(override_conditions.get("condition_rows"))
    if override_conditions is not None:
        return _coerce_dict_list(override_conditions)
    nested = _coerce_dict_list(matrix_block.get("condition_rows"))
    return nested or _coerce_dict_list(raw_payload.get("condition_rows"))


def normalize_workbook_payload(
    raw_payload: Mapping[str, Any],
    *,
    condition_rows: Sequence[Mapping[str, Any]] | Mapping[str, Any] | None = None,
) -> WorkbookPayload:
    """Validate input JSON contract and enforce the 'Absence is never zero' guardrail."""
    if not isinstance(raw_payload, Mapping):
        raise ValueError("Workbook payload must be a JSON object (dict).")

    nested_export = raw_payload.get("matrix_export")
    matrix_block: Mapping[str, Any] = (
        nested_export if isinstance(nested_export, Mapping) else raw_payload
    )

    precedent_rows = _coerce_dict_list(matrix_block.get("precedent_rows"))
    wire_rows = _coerce_dict_list(matrix_block.get("wire_rows")) or _coerce_dict_list(
        raw_payload.get("wire_rows")
    )

    if not precedent_rows and not wire_rows:
        raise ValueError(
            "Refusing to build an empty Precedent Workbook: both precedent_rows and "
            "wire_rows are empty (absence is never zero)."
        )

    raw_header = matrix_block.get("cohort_header")
    fallback_summary = raw_payload.get("cohort_summary")
    cohort_header: Mapping[str, Any] = (
        raw_header
        if isinstance(raw_header, Mapping)
        else (fallback_summary if isinstance(fallback_summary, Mapping) else {})
    )

    raw_anchor = matrix_block.get("anchor_context")
    if raw_anchor is None:
        raw_anchor = raw_payload.get("anchor_context")
    anchor_context = raw_anchor if isinstance(raw_anchor, Mapping) and bool(raw_anchor) else None

    filing_trail_rows = _coerce_dict_list(
        matrix_block.get("filing_trail_rows")
    ) or _coerce_dict_list(raw_payload.get("filing_trail_rows"))

    resolved_conditions = _extract_condition_rows(raw_payload, matrix_block, condition_rows)

    return WorkbookPayload(
        cohort_header=cohort_header,
        anchor_context=anchor_context,
        precedent_rows=precedent_rows,
        filing_trail_rows=filing_trail_rows,
        condition_rows=resolved_conditions,
        wire_rows=wire_rows,
    )


# ── Institutional Worksheet Theme & Declarative Column Strategies ────────────


class InstitutionalTheme:
    """Centralized openpyxl style factory enforcing the Wall Street styling contract."""

    def __init__(self) -> None:
        thin_subtle = Side(style="thin", color=COLOR_BORDER_SUBTLE)
        summary_top = Side(style="thin", color=COLOR_BORDER_SUMMARY)
        summary_bottom = Side(style="double", color=COLOR_BORDER_SUMMARY)
        header_side = Side(style="thin", color=COLOR_NAVY_HEADER_BG)

        self.cell_border = Border(
            left=thin_subtle,
            right=thin_subtle,
            top=thin_subtle,
            bottom=thin_subtle,
        )
        self.summary_border = Border(
            left=thin_subtle,
            right=thin_subtle,
            top=summary_top,
            bottom=summary_bottom,
        )
        self.header_border = Border(
            left=header_side,
            right=header_side,
            top=header_side,
            bottom=header_side,
        )

        self.title_font = Font(name=FONT_FAMILY, size=15, bold=True, color=COLOR_TITLE)
        self.subtitle_font = Font(name=FONT_FAMILY, size=10, italic=True, color=COLOR_SUBTITLE)
        self.header_font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_HEADER_TEXT)
        self.banner_font = Font(name=FONT_FAMILY, size=10, bold=True, color=COLOR_NAVY_HEADER_BG)

        self.header_fill = PatternFill(
            fill_type="solid",
            start_color=COLOR_NAVY_HEADER_BG,
            end_color=COLOR_NAVY_HEADER_BG,
        )
        self.banner_fill = PatternFill(
            fill_type="solid",
            start_color=COLOR_SUBSECTION_BG,
            end_color=COLOR_SUBSECTION_BG,
        )
        self._zebra_fills: dict[bool, PatternFill] = {
            False: PatternFill(
                fill_type="solid",
                start_color=COLOR_WHITE_BG,
                end_color=COLOR_WHITE_BG,
            ),
            True: PatternFill(
                fill_type="solid",
                start_color=COLOR_ZEBRA_BG,
                end_color=COLOR_ZEBRA_BG,
            ),
        }
        self._font_cache: dict[tuple[int, str, bool, bool], Font] = {}
        self._fill_cache: dict[str, PatternFill] = {}
        self._align_cache: dict[tuple[str, str, bool], Alignment] = {}

    def font(
        self,
        *,
        size: int = 10,
        color: str = COLOR_TITLE,
        bold: bool = False,
        underline: bool = False,
    ) -> Font:
        key = (size, color, bold, underline)
        underline_styles: dict[bool, str | None] = {False: None, True: "single"}
        return self._font_cache.setdefault(
            key,
            Font(
                name=FONT_FAMILY,
                size=size,
                bold=bold,
                color=color,
                underline=underline_styles[underline],
            ),
        )

    def solid_fill(self, hex_color: str) -> PatternFill:
        return self._fill_cache.setdefault(
            hex_color,
            PatternFill(fill_type="solid", start_color=hex_color, end_color=hex_color),
        )

    def row_fill(self, is_zebra: bool) -> PatternFill:
        return self._zebra_fills[is_zebra]

    def alignment(
        self,
        horizontal: str = "left",
        *,
        vertical: str = "center",
        wrap_text: bool = True,
    ) -> Alignment:
        key = (horizontal, vertical, wrap_text)
        return self._align_cache.setdefault(
            key,
            Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text),
        )


@dataclass(frozen=True)
class RenderedCell:
    """Value and presentation metadata produced by a column strategy."""

    value: Any
    role: CellRole = CellRole.BODY_TEXT
    number_format: str | None = None
    badge: BadgePalette | None = None
    bold: bool = False
    underline: bool = False
    display_hint: str | None = None


class ColumnStrategy(ABC):
    """Polymorphic table column renderer eliminating per-column branching."""

    def __init__(
        self,
        header: str,
        *,
        min_width: int = 11,
        max_width: int = 36,
        horizontal: str = "left",
        wrap_text: bool = True,
    ) -> None:
        self.header = header
        self.min_width = min_width
        self.max_width = max_width
        self.horizontal = horizontal
        self.wrap_text = wrap_text

    @abstractmethod
    def evaluate(self, row: Mapping[str, Any], excel_row: int) -> RenderedCell:
        """Produce value and styling metadata for a single worksheet cell."""


class TextColumn(ColumnStrategy):
    """Plain-text or humanized categorical column."""

    def __init__(
        self,
        header: str,
        extractor: Callable[[Mapping[str, Any]], str | None],
        *,
        min_width: int = 11,
        max_width: int = 38,
        horizontal: str = "left",
        wrap_text: bool = True,
        bold: bool = False,
    ) -> None:
        super().__init__(
            header,
            min_width=min_width,
            max_width=max_width,
            horizontal=horizontal,
            wrap_text=wrap_text,
        )
        self._extractor = extractor
        self._bold = bold

    def evaluate(self, row: Mapping[str, Any], excel_row: int) -> RenderedCell:
        val = self._extractor(row)
        return RenderedCell(value=val, role=CellRole.BODY_TEXT, bold=self._bold)


class TickerColumn(ColumnStrategy):
    """Ticker column rendering active symbols bold and delisted rows as subtle pills."""

    _CELL_BUILDERS: dict[bool, Callable[[str | None], RenderedCell]] = {
        True: lambda val: RenderedCell(
            value=val,
            role=CellRole.BODY_TEXT,
            badge=BADGE_DELISTED,
            bold=False,
        ),
        False: lambda val: RenderedCell(
            value=val,
            role=CellRole.BODY_TEXT,
            bold=True,
        ),
    }

    def __init__(
        self,
        header: str = "Ticker",
        *,
        min_width: int = 11,
        max_width: int = 13,
    ) -> None:
        super().__init__(
            header,
            min_width=min_width,
            max_width=max_width,
            horizontal="center",
            wrap_text=False,
        )

    def evaluate(self, row: Mapping[str, Any], excel_row: int) -> RenderedCell:
        val = format_ticker_display(row)
        builder = self._CELL_BUILDERS[val == DELISTED_TICKER_LABEL]
        return builder(val)


class BadgeColumn(ColumnStrategy):
    """Semantic status pill column backed by a palette lookup table."""

    def __init__(
        self,
        header: str,
        extractor: Callable[[Mapping[str, Any]], str | None],
        badges: Mapping[str, BadgePalette],
        *,
        min_width: int = 15,
        max_width: int = 24,
        horizontal: str = "center",
    ) -> None:
        super().__init__(
            header,
            min_width=min_width,
            max_width=max_width,
            horizontal=horizontal,
            wrap_text=True,
        )
        self._extractor = extractor
        self._badges = {k.lower(): v for k, v in badges.items()}

    def evaluate(self, row: Mapping[str, Any], excel_row: int) -> RenderedCell:
        label = self._extractor(row)
        badge = self._badges.get((label or "").lower())
        return RenderedCell(value=label, role=CellRole.BODY_TEXT, badge=badge)


class NumericInputColumn(ColumnStrategy):
    """Hardcoded numeric input column styled in Wall Street Blue (#0000FF)."""

    def __init__(
        self,
        header: str,
        extractor: Callable[[Mapping[str, Any]], int | float | None],
        number_format: str,
        *,
        min_width: int = 13,
        max_width: int = 18,
    ) -> None:
        super().__init__(
            header,
            min_width=min_width,
            max_width=max_width,
            horizontal="right",
            wrap_text=False,
        )
        self._extractor = extractor
        self._number_format = number_format

    def evaluate(self, row: Mapping[str, Any], excel_row: int) -> RenderedCell:
        val = self._extractor(row)
        return RenderedCell(
            value=val,
            role=CellRole.INPUT_BLUE,
            number_format=self._number_format,
            display_hint="123,456.0M",
        )


class FormulaColumn(ColumnStrategy):
    """Live Excel formula column styled in Black (same-sheet) or Green (cross-sheet)."""

    def __init__(
        self,
        header: str,
        formula_fn: Callable[[Mapping[str, Any], int], str | None],
        *,
        role: CellRole,
        number_format: str,
        min_width: int = 15,
        max_width: int = 22,
    ) -> None:
        super().__init__(
            header,
            min_width=min_width,
            max_width=max_width,
            horizontal="right",
            wrap_text=False,
        )
        self._formula_fn = formula_fn
        self._role = role
        self._number_format = number_format

    def evaluate(self, row: Mapping[str, Any], excel_row: int) -> RenderedCell:
        formula = self._formula_fn(row, excel_row)
        return RenderedCell(
            value=formula,
            role=self._role,
            number_format=self._number_format,
            display_hint="+1,250 bps",
        )


class HyperlinkColumn(ColumnStrategy):
    """Clickable ``=HYPERLINK(url, label)`` column styled in underlined #0B5394."""

    _BUILDERS: dict[bool, Callable[[str | None, str | None], RenderedCell]] = {
        True: lambda formula, display: RenderedCell(
            value=formula,
            role=CellRole.HYPERLINK,
            underline=True,
            display_hint=display or "SEC Filing",
        ),
        False: lambda _formula, display: RenderedCell(
            value=display,
            role=CellRole.BODY_TEXT,
            underline=False,
            display_hint=display,
        ),
    }

    def __init__(
        self,
        header: str,
        url_extractor: Callable[[Mapping[str, Any]], str | None],
        label_extractor: Callable[[Mapping[str, Any]], str | None],
        *,
        fallback_label: str = "SEC Filing",
        min_width: int = 18,
        max_width: int = 28,
    ) -> None:
        super().__init__(
            header,
            min_width=min_width,
            max_width=max_width,
            horizontal="center",
            wrap_text=False,
        )
        self._url_extractor = url_extractor
        self._label_extractor = label_extractor
        self._fallback_label = fallback_label

    def evaluate(self, row: Mapping[str, Any], excel_row: int) -> RenderedCell:
        url = self._url_extractor(row)
        label = self._label_extractor(row)
        formula = build_hyperlink_formula(url, label, fallback_label=self._fallback_label)
        builder = self._BUILDERS[formula is not None]
        return builder(formula, label)


# ── Worksheet Layout Primitives ──────────────────────────────────────────────


def configure_institutional_sheet(ws: Worksheet, *, freeze_cell: str | None = None) -> None:
    """Enforce hidden gridlines, landscape print fit, and freeze panes."""
    ws.views.sheetView[0].showGridLines = False
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.freeze_panes = freeze_cell


def write_sheet_banner(
    ws: Worksheet,
    theme: InstitutionalTheme,
    *,
    title: str,
    subtitle: str,
    span_cols: int,
) -> None:
    """Render Row 1 institutional title and Row 2 italic subtitle."""
    end_col = get_column_letter(max(1, span_cols))
    ws.merge_cells(f"A1:{end_col}1")
    ws.merge_cells(f"A2:{end_col}2")

    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = theme.title_font
    title_cell.alignment = theme.alignment("left", vertical="center", wrap_text=False)
    ws.row_dimensions[1].height = 26

    sub_cell = ws.cell(row=2, column=1, value=subtitle)
    sub_cell.font = theme.subtitle_font
    sub_cell.alignment = theme.alignment("left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 19
    ws.row_dimensions[3].height = 8


def write_subsection_banner(
    ws: Worksheet,
    theme: InstitutionalTheme,
    *,
    row: int,
    text: str,
    span_cols: int,
) -> None:
    """Render a subtle #F2F4F8 institutional sub-section banner across ``span_cols``."""
    end_col = get_column_letter(max(1, span_cols))
    ws.merge_cells(f"A{row}:{end_col}{row}")
    for col_idx in range(1, span_cols + 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.fill = theme.banner_fill
        cell.border = theme.cell_border
    anchor = ws.cell(row=row, column=1, value=text)
    anchor.font = theme.banner_font
    anchor.alignment = theme.alignment("left", vertical="center", wrap_text=False)
    ws.row_dimensions[row].height = 22


def write_navy_headers(
    ws: Worksheet,
    theme: InstitutionalTheme,
    *,
    row: int,
    headers: Sequence[tuple[str, str]],
) -> None:
    """Render #1F3864 Navy table column headers."""
    for col_idx, (header_text, horizontal) in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header_text)
        cell.font = theme.header_font
        cell.fill = theme.header_fill
        cell.border = theme.header_border
        cell.alignment = theme.alignment(horizontal, vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 26


def apply_rendered_cell(
    ws: Worksheet,
    theme: InstitutionalTheme,
    *,
    row: int,
    col: int,
    rendered: RenderedCell,
    font_size: int = 10,
    horizontal: str = "left",
    wrap_text: bool = True,
    is_zebra: bool = False,
    is_summary: bool = False,
) -> None:
    """Write ``rendered.value`` (keeping ``None`` as ``None``) with full formatting."""
    cell = ws.cell(row=row, column=col, value=rendered.value)
    badge = rendered.badge
    font_color = badge.font_hex if badge is not None else rendered.role.value
    bold = badge.bold if badge is not None else rendered.bold
    fill = theme.solid_fill(badge.fill_hex) if badge is not None else theme.row_fill(is_zebra)

    cell.font = theme.font(
        size=font_size,
        color=font_color,
        bold=bold,
        underline=rendered.underline,
    )
    cell.fill = fill
    cell.border = theme.summary_border if is_summary else theme.cell_border
    cell.alignment = theme.alignment(
        horizontal,
        vertical="top" if wrap_text else "center",
        wrap_text=wrap_text,
    )
    if rendered.number_format is not None:
        cell.number_format = rendered.number_format


def render_strategy_table(
    ws: Worksheet,
    theme: InstitutionalTheme,
    *,
    header_row: int,
    columns: Sequence[ColumnStrategy],
    rows: Sequence[Mapping[str, Any]],
    font_size: int = 10,
    enable_autofilter: bool = True,
) -> int:
    """Render a full table using declarative ``ColumnStrategy`` objects."""
    write_navy_headers(
        ws,
        theme,
        row=header_row,
        headers=[(col.header, col.horizontal) for col in columns],
    )
    width_tracker = [max(col.min_width, min(col.max_width, len(col.header) + 4)) for col in columns]

    effective_rows: Sequence[Mapping[str, Any]] = rows or ({},)
    has_real_rows = bool(rows)
    start_data_row = header_row + 1

    for offset, record in enumerate(effective_rows):
        excel_row = start_data_row + offset
        is_zebra = (offset % 2) == 1
        max_wrapped_lines = 1
        for col_zero, col_spec in enumerate(columns):
            rendered = (
                col_spec.evaluate(record, excel_row)
                if has_real_rows
                else _blank_placeholder_cell(col_spec)
            )
            apply_rendered_cell(
                ws,
                theme,
                row=excel_row,
                col=col_zero + 1,
                rendered=rendered,
                font_size=font_size,
                horizontal=col_spec.horizontal,
                wrap_text=col_spec.wrap_text,
                is_zebra=is_zebra,
            )
            sample = _display_sample_text(rendered)
            desired_width = min(col_spec.max_width, max(col_spec.min_width, len(sample) + 3))
            width_tracker[col_zero] = max(width_tracker[col_zero], desired_width)
            if col_spec.wrap_text and sample:
                col_cap = max(12, width_tracker[col_zero] - 2)
                est_lines = max(1, math.ceil(len(sample) / col_cap))
                max_wrapped_lines = max(max_wrapped_lines, est_lines)

        base_line_pt = 13 if font_size <= 9 else 14
        ws.row_dimensions[excel_row].height = max(20, min(96, max_wrapped_lines * base_line_pt + 6))

    last_row = start_data_row + len(effective_rows) - 1
    for col_zero, width in enumerate(width_tracker):
        ws.column_dimensions[get_column_letter(col_zero + 1)].width = width

    if enable_autofilter:
        end_letter = get_column_letter(len(columns))
        ws.auto_filter.ref = f"A{header_row}:{end_letter}{last_row}"

    return last_row


def _blank_placeholder_cell(col_spec: ColumnStrategy) -> RenderedCell:
    """Blank ('Absence is never zero') cell preserving column number mask."""
    number_format = getattr(col_spec, "_number_format", None)
    return RenderedCell(value=None, role=CellRole.BODY_TEXT, number_format=number_format)


def _display_sample_text(rendered: RenderedCell) -> str:
    if rendered.display_hint is not None:
        return rendered.display_hint
    if rendered.value is None or isinstance(rendered.value, ArrayFormula):
        return ""
    raw = str(rendered.value)
    return "" if raw.startswith("=") else raw


def _build_mean_trajectory_formula(
    _cell_ref: str,
    matrix_col: str,
    last_matrix_row: int,
    basis_token: str,
) -> str:
    """Build an unpooled cross-sheet ``AVERAGEIFS`` formula for a fiscal basis."""
    return (
        f"=IFERROR(AVERAGEIFS('Precedent Matrix'!{matrix_col}$5:{matrix_col}${last_matrix_row}, "
        f'\'Precedent Matrix\'!$K$5:$K${last_matrix_row}, "{basis_token}"), "")'
    )


def _build_median_trajectory_formula(
    cell_ref: str,
    matrix_col: str,
    last_matrix_row: int,
    basis_token: str,
) -> ArrayFormula:
    """Build an Excel CSE-safe ``ArrayFormula`` for conditional unpooled ``MEDIAN``.

    Emits OOXML ``t="array"`` with ``ref=cell_ref`` so Excel evaluates the nested
    ``IF(basis, IF(ISNUMBER(metric), metric))`` array cleanly without implicit
    intersection ``#VALUE!`` errors, ignores blank/missing cells (preventing blank
    coercion to ``0``), and preserves genuine measured zeros.
    """
    basis_range = f"'Precedent Matrix'!$K$5:$K${last_matrix_row}"
    metric_range = f"'Precedent Matrix'!{matrix_col}$5:{matrix_col}${last_matrix_row}"
    formula_text = (
        f'=IFERROR(MEDIAN(IF({basis_range}="{basis_token}", '
        f'IF(ISNUMBER({metric_range}), {metric_range}))), "")'
    )
    return ArrayFormula(ref=cell_ref, text=formula_text)


# ── Polymorphic Sheet Hierarchy ──────────────────────────────────────────────


class WorkbookSheet(ABC):
    """Abstract base worksheet renderer in the Precedent Workbook."""

    sheet_title: str

    @abstractmethod
    def render(
        self, wb: Workbook, payload: WorkbookPayload, theme: InstitutionalTheme
    ) -> Worksheet:
        """Create and populate this worksheet on ``wb``."""


class ContextSheetStrategy(WorkbookSheet, ABC):
    """Polymorphic Strategy for the 4th mode-dependent context worksheet."""

    mode: str
    sheet_position: int

    @abstractmethod
    def matches(self, payload: WorkbookPayload) -> bool:
        """Return ``True`` when this context strategy governs ``payload``."""


# ── Universal Spine Sheet 1: Cohort Base Rates ───────────────────────────────


class CohortSummarySheet(WorkbookSheet):
    """Server-wide 25k-pool denominator, Live Matrix Outcomes, and Unpooled FY vs Q block."""

    sheet_title = "Cohort Base Rates"

    _TRAJECTORY_METRICS: tuple[tuple[str, str, str], ...] = (
        ("Pre Rev ($M)", "L", FMT_CURRENCY_M),
        ("Post Rev ($M)", "M", FMT_CURRENCY_M),
        ("Pre Op Margin", "N", FMT_PERCENT),
        ("Post Op Margin", "O", FMT_PERCENT),
        ("Op Margin Delta", "P", FMT_BPS_DELTA),
        ("Pre FCF Margin", "Q", FMT_PERCENT),
        ("Post FCF Margin", "R", FMT_PERCENT),
        ("FCF Margin Delta", "S", FMT_BPS_DELTA),
    )

    _TRAJECTORY_ROWS: tuple[
        tuple[str, str, Callable[[str, str, int, str], Any]],
        ...,
    ] = (
        ("Annual Basis (FY) — Mean", "FY", _build_mean_trajectory_formula),
        ("Annual Basis (FY) — Median", "FY", _build_median_trajectory_formula),
        ("Standalone Quarterly Basis (Q) — Mean", "Standalone Q", _build_mean_trajectory_formula),
        (
            "Standalone Quarterly Basis (Q) — Median",
            "Standalone Q",
            _build_median_trajectory_formula,
        ),
    )

    def render(
        self, wb: Workbook, payload: WorkbookPayload, theme: InstitutionalTheme
    ) -> Worksheet:
        ws = wb.create_sheet(title=self.sheet_title)
        configure_institutional_sheet(ws, freeze_cell=None)

        header = payload.cohort_header
        title = (
            humanize_label(header.get("title"))
            or "Historical SEC Precedent Cohort — Base Rates & Fiscal Trajectories"
        )
        subtitle = (
            humanize_label(header.get("subtitle"))
            or "Server-wide SEC filing pool denominators paired with live loaded-matrix formulas "
            "(Annual FY and Standalone Quarterly trajectories strictly isolated)"
        )
        write_sheet_banner(ws, theme, title=title, subtitle=subtitle, span_cols=10)

        self._render_pool_denominator_block(ws, payload, theme)
        self._render_loaded_outcome_block(ws, payload, theme)
        self._render_unpooled_trajectory_block(ws, payload, theme)
        self._apply_column_widths(ws)
        return ws

    def _render_pool_denominator_block(
        self,
        ws: Worksheet,
        payload: WorkbookPayload,
        theme: InstitutionalTheme,
    ) -> None:
        header = payload.cohort_header
        ongoing = header.get("ongoing_breakdown")
        ongoing_map: Mapping[str, Any] = ongoing if isinstance(ongoing, Mapping) else {}

        total_cases = optional_int(header.get("total_matching_arcs"))
        resolved_cases = optional_int(header.get("terminal_arcs"))
        active_ongoing = optional_int(
            ongoing_map.get("active_recent_lt_365d", ongoing_map.get("active_ongoing"))
        )
        dormant_ongoing = optional_int(
            ongoing_map.get("stale_dormant_gt_365d", ongoing_map.get("dormant_ongoing"))
        )

        write_subsection_banner(
            ws,
            theme,
            row=4,
            text="Server-Wide SEC Knowledge-Base Cohort Pool (Full 25k+ Universe Denominator)",
            span_cols=4,
        )
        write_navy_headers(
            ws,
            theme,
            row=5,
            headers=(
                ("Cohort Pool Metric", "left"),
                ("Full KB Cases", "right"),
                ("Share of Pool", "right"),
                ("Scope & Methodology", "left"),
            ),
        )

        pool_rows: tuple[tuple[str, int | None, str, str], ...] = (
            (
                "Total Historical SEC Cases",
                total_cases,
                '=IF(ISNUMBER(B6),1,"")',
                "Complete non-provisional SEC precedent cohort matching query predicates",
            ),
            (
                "Resolved Cases",
                resolved_cases,
                '=IF(AND(ISNUMBER(B7),$B$6>0),B7/$B$6,"")',
                "Terminal SEC outcomes (Expanded, Pivoted, Failed, Abandoned, Merged)",
            ),
            (
                "Active Ongoing (<365d)",
                active_ongoing,
                '=IF(AND(ISNUMBER(B8),$B$6>0),B8/$B$6,"")',
                "Ongoing cases with material SEC disclosure activity within 365 days",
            ),
            (
                "Dormant Ongoing (>365d)",
                dormant_ongoing,
                '=IF(AND(ISNUMBER(B9),$B$6>0),B9/$B$6,"")',
                "Ongoing cases without subsequent SEC update for more than 365 days",
            ),
        )

        for idx, (label, count_val, share_formula, note) in enumerate(pool_rows):
            r = 6 + idx
            zebra = (idx % 2) == 1
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=1,
                rendered=RenderedCell(value=label, bold=(idx == 0)),
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=2,
                rendered=RenderedCell(
                    value=count_val,
                    role=CellRole.INPUT_BLUE,
                    number_format=FMT_INT,
                    bold=(idx == 0),
                ),
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=3,
                rendered=RenderedCell(
                    value=share_formula,
                    role=CellRole.SAME_SHEET_FORMULA,
                    number_format=FMT_PERCENT,
                ),
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=4,
                rendered=RenderedCell(value=note),
                is_zebra=zebra,
            )
            ws.row_dimensions[r].height = 20

        ws.row_dimensions[10].height = 10

    def _render_loaded_outcome_block(
        self,
        ws: Worksheet,
        payload: WorkbookPayload,
        theme: InstitutionalTheme,
    ) -> None:
        last = payload.matrix_last_row
        header = payload.cohort_header
        raw_counts = header.get("resolution_counts")
        counts_map = raw_counts if isinstance(raw_counts, Mapping) else None
        has_pool = optional_int(header.get("total_matching_arcs")) is not None or bool(counts_map)

        write_navy_headers(
            ws,
            theme,
            row=11,
            headers=(
                ("Loaded Matrix Outcome", "left"),
                ("Loaded Cases", "right"),
                ("Loaded Matrix Share", "right"),
                ("Full Cohort Cases", "right"),
            ),
        )

        total_row = 18
        for idx, (outcome_label, aliases) in enumerate(OUTCOME_CANONICAL_ROWS):
            r = 12 + idx
            zebra = (idx % 2) == 1
            count_formula = f"=COUNTIF('Precedent Matrix'!$J$5:$J${last}, \"{outcome_label}\")"
            share_formula = f'=IFERROR(B{r}/$B${total_row},"")'
            cohort_count = lookup_case_insensitive_count(
                counts_map,
                aliases,
                has_cohort_pool=has_pool,
            )

            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=1,
                rendered=RenderedCell(
                    value=outcome_label,
                    badge=OUTCOME_BADGES.get(outcome_label),
                ),
                horizontal="left",
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=2,
                rendered=RenderedCell(
                    value=count_formula,
                    role=CellRole.CROSS_SHEET_FORMULA,
                    number_format=FMT_INT,
                ),
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=3,
                rendered=RenderedCell(
                    value=share_formula,
                    role=CellRole.SAME_SHEET_FORMULA,
                    number_format=FMT_PERCENT,
                ),
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=4,
                rendered=RenderedCell(
                    value=cohort_count,
                    role=CellRole.INPUT_BLUE,
                    number_format=FMT_INT,
                ),
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            ws.row_dimensions[r].height = 20

        # Row 18: Total Loaded Precedents summary row (#9A9A9A top/bottom summary border)
        apply_rendered_cell(
            ws,
            theme,
            row=total_row,
            col=1,
            rendered=RenderedCell(value="Total Loaded Precedents", bold=True),
            is_summary=True,
        )
        apply_rendered_cell(
            ws,
            theme,
            row=total_row,
            col=2,
            rendered=RenderedCell(
                value="=SUM(B12:B17)",
                role=CellRole.SAME_SHEET_FORMULA,
                number_format=FMT_INT,
                bold=True,
            ),
            horizontal="right",
            wrap_text=False,
            is_summary=True,
        )
        apply_rendered_cell(
            ws,
            theme,
            row=total_row,
            col=3,
            rendered=RenderedCell(
                value='=IF($B$18>0,SUM(C12:C17),"")',
                role=CellRole.SAME_SHEET_FORMULA,
                number_format=FMT_PERCENT,
                bold=True,
            ),
            horizontal="right",
            wrap_text=False,
            is_summary=True,
        )
        apply_rendered_cell(
            ws,
            theme,
            row=total_row,
            col=4,
            rendered=RenderedCell(
                value='=IF(COUNT(D12:D17)>0,SUM(D12:D17),"")',
                role=CellRole.SAME_SHEET_FORMULA,
                number_format=FMT_INT,
                bold=True,
            ),
            horizontal="right",
            wrap_text=False,
            is_summary=True,
        )
        ws.row_dimensions[total_row].height = 22
        ws.row_dimensions[19].height = 10

    def _render_unpooled_trajectory_block(
        self,
        ws: Worksheet,
        payload: WorkbookPayload,
        theme: InstitutionalTheme,
    ) -> None:
        last = payload.matrix_last_row
        write_subsection_banner(
            ws,
            theme,
            row=20,
            text=(
                "Unpooled Fiscal Trajectory Summary "
                "(Annual FY and Standalone Quarterly Never Pooled)"
            ),
            span_cols=10,
        )
        headers: list[tuple[str, str]] = [
            ("Fiscal Period Basis", "left"),
            ("Cases", "right"),
        ] + [(hdr, "right") for hdr, _col, _fmt in self._TRAJECTORY_METRICS]
        write_navy_headers(ws, theme, row=21, headers=headers)

        for idx, (row_label, basis_token, formula_builder) in enumerate(self._TRAJECTORY_ROWS):
            r = 22 + idx
            zebra = (idx % 2) == 1
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=1,
                rendered=RenderedCell(value=row_label, bold=True),
                is_zebra=zebra,
            )
            count_formula = f"=COUNTIF('Precedent Matrix'!$K$5:$K${last}, \"{basis_token}\")"
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=2,
                rendered=RenderedCell(
                    value=count_formula,
                    role=CellRole.CROSS_SHEET_FORMULA,
                    number_format=FMT_INT,
                ),
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            for metric_offset, (_hdr, matrix_col, num_fmt) in enumerate(self._TRAJECTORY_METRICS):
                col_idx = 3 + metric_offset
                cell_ref = f"{get_column_letter(col_idx)}{r}"
                stat_formula = formula_builder(cell_ref, matrix_col, last, basis_token)
                apply_rendered_cell(
                    ws,
                    theme,
                    row=r,
                    col=col_idx,
                    rendered=RenderedCell(
                        value=stat_formula,
                        role=CellRole.CROSS_SHEET_FORMULA,
                        number_format=num_fmt,
                    ),
                    horizontal="right",
                    wrap_text=False,
                    is_zebra=zebra,
                )
            ws.row_dimensions[r].height = 21

    @staticmethod
    def _apply_column_widths(ws: Worksheet) -> None:
        widths = (42, 16, 18, 42, 18, 18, 19, 18, 18, 19)
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = width


# ── Universal Spine Sheet 2: Precedent Matrix ────────────────────────────────


class PrecedentMatrixSheet(WorkbookSheet):
    """Columns A..U Precedent Comps Matrix with unpooled financials and live bps deltas."""

    sheet_title = "Precedent Matrix"

    def _columns(self) -> tuple[ColumnStrategy, ...]:
        return (
            TickerColumn(),
            TextColumn(
                "Company",
                lambda r: clean_optional_str(r.get("company")),
                min_width=18,
                max_width=28,
            ),
            TextColumn(
                "Industry (SIC)",
                format_sic_industry,
                min_width=18,
                max_width=28,
            ),
            BadgeColumn(
                "Domain",
                lambda r: humanize_label(r.get("domain_label")),
                DOMAIN_BADGES,
                min_width=16,
                max_width=20,
            ),
            TextColumn(
                "Start Date",
                lambda r: clean_optional_str(r.get("start_date")),
                min_width=12,
                max_width=14,
                horizontal="center",
                wrap_text=False,
            ),
            TextColumn(
                "End Date",
                lambda r: clean_optional_str(r.get("end_date")),
                min_width=12,
                max_width=14,
                horizontal="center",
                wrap_text=False,
            ),
            TextColumn(
                "Strategic Play",
                lambda r: humanize_label(r.get("play_label")),
                min_width=18,
                max_width=26,
            ),
            TextColumn(
                "Execution Vehicle",
                lambda r: humanize_label(r.get("vehicle_label")),
                min_width=16,
                max_width=24,
            ),
            TextColumn(
                "Catalyst",
                lambda r: humanize_label(r.get("catalyst_label")),
                min_width=16,
                max_width=24,
            ),
            BadgeColumn(
                "Outcome",
                lambda r: humanize_label(r.get("outcome_label")),
                OUTCOME_BADGES,
                min_width=18,
                max_width=22,
            ),
            TextColumn(
                "Period Basis",
                lambda r: clean_optional_str(r.get("period_basis")),
                min_width=13,
                max_width=16,
                horizontal="center",
                wrap_text=False,
            ),
            NumericInputColumn(
                "Pre Rev ($M)",
                lambda r: optional_number(r.get("pre_revenue_m")),
                FMT_CURRENCY_M,
            ),
            NumericInputColumn(
                "Post Rev ($M)",
                lambda r: optional_number(r.get("post_revenue_m")),
                FMT_CURRENCY_M,
            ),
            NumericInputColumn(
                "Pre Op Margin",
                lambda r: optional_number(r.get("pre_op_margin_ratio")),
                FMT_PERCENT,
            ),
            NumericInputColumn(
                "Post Op Margin",
                lambda r: optional_number(r.get("post_op_margin_ratio")),
                FMT_PERCENT,
            ),
            FormulaColumn(
                "Op Margin Delta (bps)",
                lambda _r, row_idx: (
                    f"=IF(AND(ISNUMBER(N{row_idx}),ISNUMBER(O{row_idx})),"
                    f'ROUND((O{row_idx}-N{row_idx})*10000,0),"")'
                ),
                role=CellRole.SAME_SHEET_FORMULA,
                number_format=FMT_BPS_DELTA,
            ),
            NumericInputColumn(
                "Pre FCF Margin",
                lambda r: optional_number(r.get("pre_fcf_margin_ratio")),
                FMT_PERCENT,
            ),
            NumericInputColumn(
                "Post FCF Margin",
                lambda r: optional_number(r.get("post_fcf_margin_ratio")),
                FMT_PERCENT,
            ),
            FormulaColumn(
                "FCF Margin Delta (bps)",
                lambda _r, row_idx: (
                    f"=IF(AND(ISNUMBER(Q{row_idx}),ISNUMBER(R{row_idx})),"
                    f'ROUND((R{row_idx}-Q{row_idx})*10000,0),"")'
                ),
                role=CellRole.SAME_SHEET_FORMULA,
                number_format=FMT_BPS_DELTA,
            ),
            TextColumn(
                "Outcome Summary",
                lambda r: clean_optional_str(r.get("outcome_summary")),
                min_width=28,
                max_width=46,
            ),
            HyperlinkColumn(
                "Primary SEC Filing",
                lambda r: clean_optional_str(r.get("primary_sec_url")),
                primary_filing_display_label,
                fallback_label="SEC Filing",
                min_width=18,
                max_width=24,
            ),
        )

    def render(
        self, wb: Workbook, payload: WorkbookPayload, theme: InstitutionalTheme
    ) -> Worksheet:
        ws = wb.create_sheet(title=self.sheet_title)
        configure_institutional_sheet(ws, freeze_cell="E5")

        columns = self._columns()
        title = (
            humanize_label(payload.cohort_header.get("title")) or "SEC Precedent Comparison Matrix"
        )
        subtitle = (
            "Period-locked pre/post SEC financials (Blue = SEC filing inputs, "
            "Black = live bps margin delta formulas; blank cells = unmeasured, never zero)"
        )
        write_sheet_banner(ws, theme, title=title, subtitle=subtitle, span_cols=len(columns))
        render_strategy_table(
            ws,
            theme,
            header_row=4,
            columns=columns,
            rows=payload.precedent_rows,
            font_size=10,
            enable_autofilter=True,
        )
        return ws


# ── Universal Spine Sheet 3: SEC Filing Trail ────────────────────────────────


class FilingTrailSheet(WorkbookSheet):
    """Columns A..J chronological SEC milestone filing trail with clickable EDGAR links."""

    sheet_title = "SEC Filing Trail"

    def _columns(self) -> tuple[ColumnStrategy, ...]:
        return (
            TickerColumn(),
            TextColumn(
                "Company",
                lambda r: clean_optional_str(r.get("company")),
                min_width=16,
                max_width=26,
            ),
            TextColumn(
                "Precedent Case",
                lambda r: humanize_label(r.get("case_label")),
                min_width=20,
                max_width=32,
            ),
            TextColumn(
                "Filing Date",
                lambda r: clean_optional_str(r.get("filing_date")),
                min_width=12,
                max_width=14,
                horizontal="center",
                wrap_text=False,
            ),
            TextColumn(
                "Form & Section",
                lambda r: clean_optional_str(r.get("form_and_section")),
                min_width=16,
                max_width=24,
            ),
            TextColumn(
                "Stage",
                lambda r: humanize_label(r.get("stage_label")),
                min_width=13,
                max_width=18,
                horizontal="center",
                wrap_text=False,
            ),
            TextColumn(
                "Direction",
                lambda r: humanize_label(r.get("direction_label")),
                min_width=13,
                max_width=18,
                horizontal="center",
                wrap_text=False,
            ),
            TextColumn(
                "Disclosed Action",
                lambda r: clean_optional_str(r.get("action")),
                min_width=28,
                max_width=46,
            ),
            TextColumn(
                "Stated Motivation",
                lambda r: clean_optional_str(r.get("motivation")),
                min_width=26,
                max_width=44,
            ),
            HyperlinkColumn(
                "SEC Filing Link",
                lambda r: clean_optional_str(r.get("sec_url")),
                lambda r: (
                    clean_optional_str(r.get("accession_number"))
                    or clean_optional_str(r.get("form_and_section"))
                ),
                fallback_label="EDGAR Filing",
                min_width=20,
                max_width=26,
            ),
        )

    def render(
        self, wb: Workbook, payload: WorkbookPayload, theme: InstitutionalTheme
    ) -> Worksheet:
        ws = wb.create_sheet(title=self.sheet_title)
        configure_institutional_sheet(ws, freeze_cell="C5")

        columns = self._columns()
        write_sheet_banner(
            ws,
            theme,
            title="SEC Milestone Filing Provenance Trail",
            subtitle=(
                "Chronological non-procedural SEC filing milestones backing the loaded "
                "precedent cases with direct EDGAR accession links"
            ),
            span_cols=len(columns),
        )
        render_strategy_table(
            ws,
            theme,
            header_row=4,
            columns=columns,
            rows=payload.filing_trail_rows,
            font_size=9,
            enable_autofilter=True,
        )
        return ws


# ── Context Strategy 1: Disclosure Wire (Mode A Watchlist Triage, Sheet 1) ──


class DisclosureWireSheet(ContextSheetStrategy):
    """Placed as Sheet 1 when ``wire_rows`` is non-empty (Mode A watchlist triage)."""

    mode = "wire"
    sheet_title = "Disclosure Wire"
    sheet_position = 0

    def matches(self, payload: WorkbookPayload) -> bool:
        return bool(payload.wire_rows)

    def _columns(self) -> tuple[ColumnStrategy, ...]:
        return (
            TextColumn(
                "Ticker",
                lambda r: clean_optional_str(r.get("ticker")),
                min_width=10,
                max_width=12,
                horizontal="center",
                wrap_text=False,
                bold=True,
            ),
            TextColumn(
                "Company",
                lambda r: clean_optional_str(r.get("company")),
                min_width=18,
                max_width=28,
            ),
            TextColumn(
                "Industry (SIC)",
                format_sic_industry,
                min_width=18,
                max_width=28,
            ),
            TextColumn(
                "Filing Date",
                lambda r: clean_optional_str(r.get("filing_date")),
                min_width=12,
                max_width=14,
                horizontal="center",
                wrap_text=False,
            ),
            TextColumn(
                "Form & Section",
                lambda r: clean_optional_str(r.get("form_and_section")),
                min_width=16,
                max_width=24,
            ),
            BadgeColumn(
                "Stitch Status",
                lambda r: humanize_label(r.get("stitch_state")),
                STITCH_BADGES,
                min_width=16,
                max_width=20,
            ),
            TextColumn(
                "Strategic Play",
                lambda r: humanize_label(r.get("play_label")),
                min_width=18,
                max_width=26,
            ),
            TextColumn(
                "Execution Vehicle",
                lambda r: humanize_label(r.get("vehicle_label")),
                min_width=16,
                max_width=24,
            ),
            TextColumn(
                "Catalyst",
                lambda r: humanize_label(r.get("catalyst_label")),
                min_width=16,
                max_width=24,
            ),
            NumericInputColumn(
                "Cohort Precedents",
                wire_precedent_count,
                FMT_INT,
                min_width=16,
                max_width=19,
            ),
            TextColumn(
                "Disclosed Action",
                lambda r: clean_optional_str(r.get("action")),
                min_width=28,
                max_width=46,
            ),
            TextColumn(
                "Stated Motivation",
                lambda r: clean_optional_str(r.get("motivation")),
                min_width=26,
                max_width=44,
            ),
            HyperlinkColumn(
                "SEC Filing Link",
                lambda r: clean_optional_str(r.get("sec_url")),
                lambda r: (
                    clean_optional_str(r.get("accession_number"))
                    or clean_optional_str(r.get("form_and_section"))
                ),
                fallback_label="EDGAR Filing",
                min_width=20,
                max_width=26,
            ),
        )

    def render(
        self, wb: Workbook, payload: WorkbookPayload, theme: InstitutionalTheme
    ) -> Worksheet:
        ws = wb.create_sheet(title=self.sheet_title)
        configure_institutional_sheet(ws, freeze_cell="C5")

        columns = self._columns()
        write_sheet_banner(
            ws,
            theme,
            title="SEC Disclosure Wire — Recent Filing Situations & Cohort Triage",
            subtitle=(
                "Material recent SEC disclosures classified by cohort stitch state, "
                "matched precedent depth, and direct EDGAR accession links"
            ),
            span_cols=len(columns),
        )
        render_strategy_table(
            ws,
            theme,
            header_row=4,
            columns=columns,
            rows=payload.wire_rows,
            font_size=10,
            enable_autofilter=True,
        )
        return ws


# ── Context Strategy 2: Anchor Scorecard (Mode B Subject Stress-Test, Sheet 3)


@dataclass(frozen=True)
class _AnchorMetricRowSpec:
    """Declarative row spec for the Subject Company Unpooled Financial Card."""

    label: str
    key: str
    number_format: str
    note: str

    def benchmark_rendered_cell(self, last_matrix_row: int, excel_row: int) -> RenderedCell:
        return RenderedCell(
            value=None,
            role=CellRole.BODY_TEXT,
            number_format=self.number_format,
        )


@dataclass(frozen=True)
class _BenchmarkedAnchorMetricRowSpec(_AnchorMetricRowSpec):
    """Anchor metric row paired with a live same-basis Precedent Matrix benchmark formula."""

    matrix_col: str

    def benchmark_rendered_cell(self, last_matrix_row: int, excel_row: int) -> RenderedCell:
        formula = (
            f"=IFERROR(AVERAGEIFS('Precedent Matrix'!{self.matrix_col}$5:"
            f"{self.matrix_col}${last_matrix_row}, "
            f"'Precedent Matrix'!$K$5:$K${last_matrix_row}, C{excel_row}), \"\")"
        )
        return RenderedCell(
            value=formula,
            role=CellRole.CROSS_SHEET_FORMULA,
            number_format=self.number_format,
        )


class AnchorScorecardSheet(ContextSheetStrategy):
    """Placed as Sheet 3 when ``anchor_context`` or ``condition_rows`` is non-empty."""

    mode = "anchor"
    sheet_title = "Anchor Scorecard"
    sheet_position = 2

    _METRIC_SPECS: tuple[_AnchorMetricRowSpec, ...] = (
        _BenchmarkedAnchorMetricRowSpec(
            "Revenue ($M)",
            "revenue_m",
            FMT_CURRENCY_M,
            "Latest reported period revenue (USD millions, unpooled)",
            matrix_col="L",
        ),
        _BenchmarkedAnchorMetricRowSpec(
            "Operating Margin",
            "operating_margin_ratio",
            FMT_PERCENT,
            "Operating income / revenue on locked period basis",
            matrix_col="N",
        ),
        _AnchorMetricRowSpec(
            "Free Cash Flow ($M)",
            "fcf_m",
            FMT_CURRENCY_M,
            "Operating cash flow less capital expenditures (USD millions)",
        ),
        _BenchmarkedAnchorMetricRowSpec(
            "FCF Margin",
            "fcf_margin_ratio",
            FMT_PERCENT,
            "Free cash flow / revenue on locked period basis",
            matrix_col="Q",
        ),
        _AnchorMetricRowSpec(
            "Cash & Equivalents ($M)",
            "cash_m",
            FMT_CURRENCY_M,
            "Balance sheet cash and short-term investments (USD millions)",
        ),
        _AnchorMetricRowSpec(
            "Total Debt ($M)",
            "total_debt_m",
            FMT_CURRENCY_M,
            "Short-term plus long-term debt obligations (USD millions)",
        ),
        _AnchorMetricRowSpec(
            "Debt / Total Assets",
            "debt_to_assets_ratio",
            FMT_PERCENT,
            "Total debt / total assets balance-sheet leverage ratio",
        ),
    )

    def matches(self, payload: WorkbookPayload) -> bool:
        return payload.anchor_context is not None or bool(payload.condition_rows)

    def _condition_columns(self) -> tuple[ColumnStrategy, ...]:
        return (
            TextColumn(
                "Transfer Condition",
                lambda r: clean_optional_str(r.get("condition")),
                min_width=22,
                max_width=32,
                bold=True,
            ),
            TextColumn(
                "Precedent Benchmark",
                lambda r: clean_optional_str(r.get("precedent_benchmark")),
                min_width=24,
                max_width=36,
            ),
            TextColumn(
                "Subject Anchor Status",
                lambda r: clean_optional_str(r.get("anchor_status")),
                min_width=24,
                max_width=36,
            ),
            BadgeColumn(
                "Transfer Verdict",
                lambda r: humanize_label(r.get("verdict")),
                VERDICT_BADGES,
                min_width=16,
                max_width=22,
            ),
            TextColumn(
                "Filing & Cohort Evidence",
                lambda r: clean_optional_str(r.get("evidence")),
                min_width=30,
                max_width=48,
            ),
        )

    def render(
        self, wb: Workbook, payload: WorkbookPayload, theme: InstitutionalTheme
    ) -> Worksheet:
        ws = wb.create_sheet(title=self.sheet_title)
        configure_institutional_sheet(ws, freeze_cell=None)

        anchor: Mapping[str, Any] = payload.anchor_context or {}
        ticker = clean_optional_str(anchor.get("ticker"))
        company = clean_optional_str(anchor.get("company"))
        period_basis = clean_optional_str(anchor.get("period_basis"))
        latest_period = clean_optional_str(anchor.get("latest_period"))
        sic_industry = format_sic_industry(anchor)
        baseline_ind = humanize_label(anchor.get("baseline_industry"))

        entity_parts = [p for p in (ticker, company) if p]
        title_prefix = " — ".join(entity_parts) or "Subject Anchor"
        write_sheet_banner(
            ws,
            theme,
            title=f"{title_prefix} | Unpooled Financial Baseline & Condition Scorecard",
            subtitle=(
                f"Industry: {sic_industry or 'Unclassified'} | "
                f"Baseline Domain: {baseline_ind or 'Unclassified'} | "
                f"Locked Period Basis: {period_basis or 'Unreported'} "
                f"({latest_period or 'N/A'})"
            ),
            span_cols=5,
        )

        next_row = self._render_anchor_card(ws, payload, anchor, theme)
        write_subsection_banner(
            ws,
            theme,
            row=next_row,
            text="Precedent Condition Transfer Scorecard",
            span_cols=5,
        )
        render_strategy_table(
            ws,
            theme,
            header_row=next_row + 1,
            columns=self._condition_columns(),
            rows=payload.condition_rows,
            font_size=10,
            enable_autofilter=False,
        )
        return ws

    def _render_anchor_card(
        self,
        ws: Worksheet,
        payload: WorkbookPayload,
        anchor: Mapping[str, Any],
        theme: InstitutionalTheme,
    ) -> int:
        last = payload.matrix_last_row
        period_basis = clean_optional_str(anchor.get("period_basis"))
        latest_period = clean_optional_str(anchor.get("latest_period"))

        write_subsection_banner(
            ws,
            theme,
            row=4,
            text="Subject Company Financial Snapshot (Strictly Unpooled Period Basis)",
            span_cols=5,
        )
        write_navy_headers(
            ws,
            theme,
            row=5,
            headers=(
                ("Subject Baseline Metric", "left"),
                ("Subject Value", "right"),
                ("Period Basis", "center"),
                ("Same-Basis Precedent Avg", "right"),
                ("Reporting Window & Definition", "left"),
            ),
        )

        window_prefix = f"{latest_period}: " if latest_period else ""
        for idx, spec in enumerate(self._METRIC_SPECS):
            r = 6 + idx
            zebra = (idx % 2) == 1
            metric_val = optional_number(anchor.get(spec.key))
            bench_cell = spec.benchmark_rendered_cell(last, r)
            window_note = f"{window_prefix}{spec.note}"

            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=1,
                rendered=RenderedCell(value=spec.label, bold=True),
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=2,
                rendered=RenderedCell(
                    value=metric_val,
                    role=CellRole.INPUT_BLUE,
                    number_format=spec.number_format,
                ),
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=3,
                rendered=RenderedCell(value=period_basis),
                horizontal="center",
                wrap_text=False,
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=4,
                rendered=bench_cell,
                horizontal="right",
                wrap_text=False,
                is_zebra=zebra,
            )
            apply_rendered_cell(
                ws,
                theme,
                row=r,
                col=5,
                rendered=RenderedCell(value=window_note),
                is_zebra=zebra,
            )
            ws.row_dimensions[r].height = 20

        spacer_row = 6 + len(self._METRIC_SPECS)
        ws.row_dimensions[spacer_row].height = 10
        return spacer_row + 1


# ── Context Strategy 3: Sector Breakdown (Mode C Top-Down Sweep, Sheet 3) ───


class SectorBreakdownSheet(ContextSheetStrategy):
    """Placed as Sheet 3 on anchor-less Mode C sweeps with live ``=COUNTIFS`` formulas."""

    mode = "sweep"
    sheet_title = "Sector Breakdown"
    sheet_position = 2

    def matches(self, payload: WorkbookPayload) -> bool:
        return True

    @staticmethod
    def _distinct_sectors(precedent_rows: Sequence[Mapping[str, Any]]) -> list[str]:
        counts: Counter[str] = Counter(filter(None, map(format_sic_industry, precedent_rows)))
        ordered = sorted(counts.items(), key=lambda item: (-item[1], item[0]))
        return [sector for sector, _cnt in ordered]

    def _columns(self, last_matrix_row: int, total_row: int) -> tuple[ColumnStrategy, ...]:
        c_range = f"'Precedent Matrix'!$C$5:$C${last_matrix_row}"
        j_range = f"'Precedent Matrix'!$J$5:$J${last_matrix_row}"
        return (
            TextColumn(
                "Industry (SIC)",
                lambda r: clean_optional_str(r.get("industry")),
                min_width=24,
                max_width=38,
                bold=True,
            ),
            FormulaColumn(
                "Cases",
                lambda _r, row_idx: f"=COUNTIFS({c_range}, $A{row_idx})",
                role=CellRole.CROSS_SHEET_FORMULA,
                number_format=FMT_INT,
                min_width=12,
                max_width=15,
            ),
            FormulaColumn(
                "Matrix Share",
                lambda _r, row_idx: f'=IFERROR(B{row_idx}/$B${total_row},"")',
                role=CellRole.SAME_SHEET_FORMULA,
                number_format=FMT_PERCENT,
                min_width=14,
                max_width=16,
            ),
            FormulaColumn(
                "Expanded",
                lambda _r, row_idx: (
                    f'=COUNTIFS({c_range}, $A{row_idx}, {j_range}, "Expanded / Succeeded")'
                ),
                role=CellRole.CROSS_SHEET_FORMULA,
                number_format=FMT_INT,
                min_width=13,
                max_width=16,
            ),
            FormulaColumn(
                "Pivoted",
                lambda _r, row_idx: f'=COUNTIFS({c_range}, $A{row_idx}, {j_range}, "Pivoted")',
                role=CellRole.CROSS_SHEET_FORMULA,
                number_format=FMT_INT,
                min_width=13,
                max_width=16,
            ),
            FormulaColumn(
                "Failed / Abandoned",
                lambda _r, row_idx: (
                    f'=COUNTIFS({c_range}, $A{row_idx}, {j_range}, "Failed")+'
                    f'COUNTIFS({c_range}, $A{row_idx}, {j_range}, "Abandoned")'
                ),
                role=CellRole.CROSS_SHEET_FORMULA,
                number_format=FMT_INT,
                min_width=18,
                max_width=22,
            ),
            FormulaColumn(
                "Merged / Ongoing",
                lambda _r, row_idx: (
                    f'=COUNTIFS({c_range}, $A{row_idx}, {j_range}, "Merged")+'
                    f'COUNTIFS({c_range}, $A{row_idx}, {j_range}, "Ongoing")'
                ),
                role=CellRole.CROSS_SHEET_FORMULA,
                number_format=FMT_INT,
                min_width=17,
                max_width=21,
            ),
        )

    def render(
        self, wb: Workbook, payload: WorkbookPayload, theme: InstitutionalTheme
    ) -> Worksheet:
        ws = wb.create_sheet(title=self.sheet_title)
        configure_institutional_sheet(ws, freeze_cell="B5")

        sectors = self._distinct_sectors(payload.precedent_rows)
        sector_records = [{"industry": name} for name in sectors]
        row_count = max(1, len(sector_records))
        total_row = 5 + row_count
        columns = self._columns(payload.matrix_last_row, total_row)

        write_sheet_banner(
            ws,
            theme,
            title="Cross-Sector Precedent Distribution by SIC Industry",
            subtitle=(
                "Live cross-sheet outcome counts across distinct SIC industries "
                "represented in the Precedent Matrix"
            ),
            span_cols=len(columns),
        )
        last_data_row = render_strategy_table(
            ws,
            theme,
            header_row=4,
            columns=columns,
            rows=sector_records,
            font_size=10,
            enable_autofilter=True,
        )

        # Render Total Summary Row (#9A9A9A top/bottom summary border)
        apply_rendered_cell(
            ws,
            theme,
            row=total_row,
            col=1,
            rendered=RenderedCell(value="Total Classified Sector Cases", bold=True),
            is_summary=True,
        )
        summary_specs: tuple[tuple[int, str, str], ...] = (
            (2, f"=SUM(B5:B{last_data_row})", FMT_INT),
            (3, f'=IF($B${total_row}>0,SUM(C5:C{last_data_row}),"")', FMT_PERCENT),
            (4, f"=SUM(D5:D{last_data_row})", FMT_INT),
            (5, f"=SUM(E5:E{last_data_row})", FMT_INT),
            (6, f"=SUM(F5:F{last_data_row})", FMT_INT),
            (7, f"=SUM(G5:G{last_data_row})", FMT_INT),
        )
        for col_idx, sum_formula, num_fmt in summary_specs:
            apply_rendered_cell(
                ws,
                theme,
                row=total_row,
                col=col_idx,
                rendered=RenderedCell(
                    value=sum_formula,
                    role=CellRole.SAME_SHEET_FORMULA,
                    number_format=num_fmt,
                    bold=True,
                ),
                horizontal="right",
                wrap_text=False,
                is_summary=True,
            )
        ws.row_dimensions[total_row].height = 22
        return ws


# ── Strategy Registry & Workbook Compiler ───────────────────────────────────

CONTEXT_STRATEGY_REGISTRY: tuple[ContextSheetStrategy, ...] = (
    DisclosureWireSheet(),
    AnchorScorecardSheet(),
    SectorBreakdownSheet(),
)


def select_context_strategy(payload: WorkbookPayload) -> ContextSheetStrategy:
    """Select the governing polymorphic ``ContextSheetStrategy`` from the registry."""
    return next(strategy for strategy in CONTEXT_STRATEGY_REGISTRY if strategy.matches(payload))


class PrecedentWorkbookCompiler:
    """Deterministic 4-sheet institutional Precedent Matrix ``.xlsx`` compiler."""

    def __init__(
        self,
        *,
        strategies: Sequence[ContextSheetStrategy] = CONTEXT_STRATEGY_REGISTRY,
    ) -> None:
        self._strategies = tuple(strategies)

    def select_strategy(self, payload: WorkbookPayload) -> ContextSheetStrategy:
        return next(strategy for strategy in self._strategies if strategy.matches(payload))

    def compile(self, payload: WorkbookPayload) -> Workbook:
        wb = Workbook()
        placeholder_sheet = wb.worksheets[0]

        theme = InstitutionalTheme()
        context_sheet = self.select_strategy(payload)
        ordered_sheets: list[WorkbookSheet] = [
            CohortSummarySheet(),
            PrecedentMatrixSheet(),
            FilingTrailSheet(),
        ]
        ordered_sheets.insert(context_sheet.sheet_position, context_sheet)

        for sheet_spec in ordered_sheets:
            sheet_spec.render(wb, payload, theme)

        wb.remove(placeholder_sheet)

        wb.properties.creator = "Slinky Precedent Compiler"
        wb.properties.lastModifiedBy = "Slinky Precedent Compiler"
        wb.properties.created = DETERMINISTIC_CREATED_AT
        wb.properties.modified = DETERMINISTIC_CREATED_AT
        return wb


# ── Deterministic Byte-Identical XLSX Serializer ────────────────────────────

_CREATED_TAG_RE = re.compile(r"(<(?:[\w]+:)?created\b[^>]*>).*?(</(?:[\w]+:)?created>)")
_MODIFIED_TAG_RE = re.compile(r"(<(?:[\w]+:)?modified\b[^>]*>).*?(</(?:[\w]+:)?modified>)")


def _canonicalize_core_xml(xml_bytes: bytes) -> bytes:
    text = xml_bytes.decode("utf-8")
    text = _CREATED_TAG_RE.sub(rf"\g<1>{_DETERMINISTIC_ISO_UTC}\g<2>", text)
    text = _MODIFIED_TAG_RE.sub(rf"\g<1>{_DETERMINISTIC_ISO_UTC}\g<2>", text)
    return text.encode("utf-8")


def serialize_deterministic_workbook(wb: Workbook) -> bytes:
    """Serialize ``wb`` into byte-identical ``.xlsx`` archive bytes across runs."""
    wb.properties.created = DETERMINISTIC_CREATED_AT
    wb.properties.modified = DETERMINISTIC_CREATED_AT

    raw_buffer = io.BytesIO()
    wb.save(raw_buffer)

    canonical_buffer = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(raw_buffer.getvalue()), mode="r") as src_zip,
        zipfile.ZipFile(
            canonical_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=6,
        ) as dst_zip,
    ):
        for filename in sorted(src_zip.namelist()):
            entry_bytes = src_zip.read(filename)
            if filename == "docProps/core.xml":
                entry_bytes = _canonicalize_core_xml(entry_bytes)
            info = zipfile.ZipInfo(filename=filename, date_time=_DETERMINISTIC_ZIP_TIMESTAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            dst_zip.writestr(info, entry_bytes)

    return canonical_buffer.getvalue()


# ── Public Library Functions ─────────────────────────────────────────────────


def compile_precedent_workbook(
    payload: Mapping[str, Any],
    *,
    condition_rows: Sequence[Mapping[str, Any]] | Mapping[str, Any] | None = None,
) -> Workbook:
    """Compile an in-memory ``openpyxl.Workbook`` from a ``matrix_export`` payload."""
    normalized = normalize_workbook_payload(payload, condition_rows=condition_rows)
    compiler = PrecedentWorkbookCompiler()
    return compiler.compile(normalized)


def compile_precedent_workbook_bytes(
    payload: Mapping[str, Any],
    *,
    condition_rows: Sequence[Mapping[str, Any]] | Mapping[str, Any] | None = None,
) -> bytes:
    """Compile deterministic byte-identical ``.xlsx`` bytes from ``payload``."""
    wb = compile_precedent_workbook(payload, condition_rows=condition_rows)
    return serialize_deterministic_workbook(wb)


def build_precedent_workbook(
    payload: Mapping[str, Any],
    output_path: str | Path,
    *,
    condition_rows: Sequence[Mapping[str, Any]] | Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Compile and write deterministic ``.xlsx`` workbook to ``output_path``."""
    normalized = normalize_workbook_payload(payload, condition_rows=condition_rows)
    compiler = PrecedentWorkbookCompiler()
    strategy = compiler.select_strategy(normalized)
    wb = compiler.compile(normalized)
    xlsx_bytes = serialize_deterministic_workbook(wb)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(xlsx_bytes)

    return {
        "output": str(out),
        "mode": strategy.mode,
        "context_sheet": strategy.sheet_title,
        "sheets": list(wb.sheetnames),
        "precedent_count": len(normalized.precedent_rows),
        "filing_trail_count": len(normalized.filing_trail_rows),
        "condition_count": len(normalized.condition_rows),
        "wire_count": len(normalized.wire_rows),
        "bytes_written": len(xlsx_bytes),
    }


# ── Standalone CLI Entrypoint ────────────────────────────────────────────────


def _read_json_input(path_or_dash: str) -> Any:
    if path_or_dash == "-":
        raw_stdin = sys.stdin.buffer.read().decode("utf-8-sig")
        return json.loads(raw_stdin)
    return json.loads(Path(path_or_dash).read_text(encoding="utf-8-sig"))


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="build_precedent_workbook",
        description=(
            "Compile a deterministic 4-sheet institutional Precedent Matrix .xlsx "
            "workbook from Slinky KB matrix_export or wire_rows JSON."
        ),
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Path to input JSON file (or '-' to read JSON from stdin).",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Destination .xlsx path to write.",
    )
    parser.add_argument(
        "--conditions-json",
        default=None,
        help="Optional path to JSON file containing condition_rows for Anchor Scorecard.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    try:
        raw_payload = _read_json_input(args.input)
        extra_conditions = _read_json_input(args.conditions_json) if args.conditions_json else None
        summary = build_precedent_workbook(
            raw_payload,
            args.output,
            condition_rows=extra_conditions,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(json.dumps(summary))
    return 0


__all__ = [
    "BADGE_ADVERSE",
    "BADGE_DELISTED",
    "BADGE_EXPANDED",
    "BADGE_NEUTRAL",
    "BADGE_OUTSIDE_DOMAIN",
    "BADGE_PIVOT",
    "COLOR_BORDER_SUBTLE",
    "COLOR_BORDER_SUMMARY",
    "COLOR_CROSS_SHEET_GREEN",
    "COLOR_FORMULA_BLACK",
    "COLOR_HEADER_TEXT",
    "COLOR_HYPERLINK_BLUE",
    "COLOR_INPUT_BLUE",
    "COLOR_NAVY_HEADER_BG",
    "COLOR_SUBSECTION_BG",
    "COLOR_SUBTITLE",
    "COLOR_TITLE",
    "COLOR_WHITE_BG",
    "COLOR_ZEBRA_BG",
    "CONTEXT_STRATEGY_REGISTRY",
    "DELISTED_TICKER_LABEL",
    "DETERMINISTIC_CREATED_AT",
    "DOMAIN_BADGES",
    "FMT_BPS_DELTA",
    "FMT_CURRENCY_M",
    "FMT_INT",
    "FMT_PERCENT",
    "FONT_FAMILY",
    "OUTCOME_BADGES",
    "OUTCOME_CANONICAL_ROWS",
    "STITCH_BADGES",
    "VERDICT_BADGES",
    "AnchorScorecardSheet",
    "BadgeColumn",
    "BadgePalette",
    "CellRole",
    "CohortSummarySheet",
    "ColumnStrategy",
    "ContextSheetStrategy",
    "DisclosureWireSheet",
    "FilingTrailSheet",
    "FormulaColumn",
    "HyperlinkColumn",
    "InstitutionalTheme",
    "NumericInputColumn",
    "PrecedentMatrixSheet",
    "PrecedentWorkbookCompiler",
    "RenderedCell",
    "SectorBreakdownSheet",
    "TextColumn",
    "TickerColumn",
    "WorkbookPayload",
    "WorkbookSheet",
    "build_arg_parser",
    "build_hyperlink_formula",
    "build_precedent_workbook",
    "clean_optional_str",
    "compile_precedent_workbook",
    "compile_precedent_workbook_bytes",
    "format_sic_industry",
    "format_ticker_display",
    "humanize_label",
    "main",
    "normalize_workbook_payload",
    "optional_int",
    "optional_number",
    "select_context_strategy",
    "serialize_deterministic_workbook",
]


if __name__ == "__main__":
    raise SystemExit(main())
