from __future__ import annotations

from typing import Sequence, Tuple, Union

# --- Public constants (expected by imports) ---
_PT_TO_EMU = 12700
PT_PER_INCH: float = 72.0  # typographic points per inch
EMU_PER_INCH: int = 914_400  # Office EMUs per inch
EMU_PER_PT: float = EMU_PER_INCH / PT_PER_INCH  # 12,700.0 EMU per point


# --- Conversions ---


def pt_to_emu(value_pt: float | int) -> int:
    """Convert points (1/72 inch) to EMUs."""
    return int(round(float(value_pt) * EMU_PER_PT))


def emu_to_pt(value_emu: int | float) -> float:
    """Convert EMUs to points."""
    return float(value_emu) / EMU_PER_PT


def inches_to_emu(value_in: float | int) -> int:
    """Convert inches to EMUs."""
    return int(round(float(value_in) * EMU_PER_INCH))


def emu_to_inches(value_emu: int | float) -> float:
    """Convert EMUs to inches."""
    return float(value_emu) / EMU_PER_INCH


def _pt_to_emu(pt: float) -> int:
    return int(round(float(pt) * _PT_TO_EMU))


def _emu_to_pt(emu: int) -> float:
    return float(emu) / _PT_TO_EMU


# --- Lightweight layout helpers (deterministic for tests/CI) ---


def line_height_pt(
    font_size_pt: float,
    line_spacing: float = 1.2,
    padding_pt: float = 0.0,  # <-- default 0 so 10 * 1.2 == 12.0
) -> float:
    """
    Single-line height estimate in *points*:
    (font_size * line_spacing) + top/bottom padding
    """
    return (font_size_pt * line_spacing) + (2 * padding_pt)


def calc_row_height_emu(
    font_size_pt: float,
    line_spacing: float = 1.2,
    padding_pt: float = 0.0,
) -> int:
    """
    Estimate a single text-row height in EMUs:
      (font_size * line_spacing) + top/bottom padding
    """
    height_pt = line_height_pt(font_size_pt, line_spacing, padding_pt)
    return pt_to_emu(height_pt)


def para_spacing_before_after_pt(
    font_size_pt: float | None,
    before_factor: float | None = 0.0,
    after_factor: float | None = 0.0,
    *,
    min_before_pt: float = 0.0,
    min_after_pt: float = 0.0,
) -> Tuple[float, float]:
    """
    Returns (before_pt, after_pt).

    Rules expected by tests:
    - If font_size_pt is None:
        * Treat the 2nd positional argument as an absolute "after" spacing in points.
        * before is 0.0.
    - If font_size_pt is provided:
        * before_factor=None  => before = 1.0 * font_size_pt
        * after_factor default (0.0) => after = 0.0 * font_size_pt
    - Apply min_before_pt / min_after_pt floors.
    """
    # Case 1: font size unknown -> interpret 2nd arg as absolute "after" (pt)
    if font_size_pt is None:
        before_pt = 0.0
        after_pt = 0.0 if before_factor is None else float(before_factor)
        # floors
        before_pt = max(before_pt, float(min_before_pt))
        after_pt = max(after_pt, float(min_after_pt))
        return (before_pt, after_pt)

    # Case 2: font size provided -> factors
    fs = float(font_size_pt)
    # before_factor=None means "use 1.0 × font-size" for before
    bf = 1.0 if before_factor is None else float(before_factor)
    # after_factor default is 0.0 (means no extra spacing)
    af = 0.0 if after_factor is None else float(after_factor)

    before_pt = max(fs * bf, float(min_before_pt))
    after_pt = max(fs * af, float(min_after_pt))
    return (before_pt, after_pt)


def _wrap_by_char_capacity(s: str, max_chars_per_line: int) -> list[str]:
    if max_chars_per_line <= 0:
        return [s]
    chunks = [s[i : i + max_chars_per_line] for i in range(0, len(s), max_chars_per_line)]
    return chunks or [""]


def estimate_text_bounds_emu(
    text: Union[str, Sequence[str]],
    font_size_pt: float,
    max_width_emu: int | None = None,
    *,
    avg_char_width_factor: float = 0.5,
    line_spacing: float = 1.2,
    padding_pt: float = 2.0,
) -> Tuple:  # quotes to avoid importing at module import time
    if font_size_pt <= 0:
        from pptx.util import Emu as _Emu  # local import avoids circulars

        return _Emu(0), _Emu(0)

    avg_char_width_pt = max(0.1, float(font_size_pt) * float(avg_char_width_factor))

    # Build lines
    if isinstance(text, (list, tuple)):
        lines: list[str] = [("" if t is None else str(t)) for t in text] or [""]
    else:
        s = "" if text is None else str(text)
        if max_width_emu is not None and max_width_emu > 0:
            max_chars = max(1, int(max_width_emu / _pt_to_emu(avg_char_width_pt)))
            lines = _wrap_by_char_capacity(s, max_chars)
        else:
            lines = [s]

    # Height
    line_height_pt = float(font_size_pt) * float(line_spacing)
    height_pt = (line_height_pt * len(lines)) + (2.0 * float(padding_pt))
    height_emu_int = _pt_to_emu(height_pt)

    # Width (longest line)
    longest_len = max((len(line) for line in lines), default=0)
    if max_width_emu is not None and max_width_emu > 0:
        width_pt = min(longest_len * avg_char_width_pt, _emu_to_pt(max_width_emu))
    else:
        width_pt = longest_len * avg_char_width_pt
    width_emu_int = _pt_to_emu(width_pt)

    # Return real Emu objects (tests check isinstance(..., Emu))
    from pptx.util import Emu as _Emu  # import right before returning

    return _Emu(int(width_emu_int)), _Emu(int(height_emu_int))


# --- Public API (keep at end to avoid F822) ---
__all__ = [
    "PT_PER_INCH",
    "EMU_PER_INCH",
    "EMU_PER_PT",
    "pt_to_emu",
    "emu_to_pt",
    "inches_to_emu",
    "emu_to_inches",
    "line_height_pt",
    "calc_row_height_emu",
    "para_spacing_before_after_pt",
    "estimate_text_bounds_emu",
]
