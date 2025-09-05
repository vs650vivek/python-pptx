import math

from pptx.layout._measure import (
    emu_to_pt,
    estimate_text_bounds_emu,
    line_height_pt,
    para_spacing_before_after_pt,
    pt_to_emu,
)
from pptx.util import Emu


def test_pt_emu_roundtrip_is_stable():
    for pt in (0, 0.5, 1, 10, 12, 24, 72, 100.25):
        emu = pt_to_emu(pt)
        back = emu_to_pt(emu)
        assert math.isclose(back, float(pt), rel_tol=0, abs_tol=1 / 12700)


def test_line_height_defaults_and_multiplier():
    assert math.isclose(line_height_pt(10), 12.0)
    assert math.isclose(line_height_pt(10, 1.5), 15.0)


def test_para_spacing_normalization():
    b, a = para_spacing_before_after_pt(None, 6)
    assert b == 0.0 and a == 6.0
    b, a = para_spacing_before_after_pt(3, None)
    assert b == 3.0 and a == 0.0


def test_estimate_text_bounds_scales_with_lines_and_chars():
    w1, h1 = estimate_text_bounds_emu(["abc"], 10)
    w2, h2 = estimate_text_bounds_emu(["abcdefgh"], 10)
    w3, h3 = estimate_text_bounds_emu(["a", "b"], 10)
    assert int(w2) > int(w1)  # more chars → wider
    assert int(h3) > int(h1)  # more lines → taller
    w4, h4 = estimate_text_bounds_emu(["abc"], 10, padding_pt=10)
    assert int(h4) > int(h1)
    assert isinstance(w4, Emu) and isinstance(h4, Emu)
