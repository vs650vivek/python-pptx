import io
from pptx import Presentation

def _shape(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shape = slide.shapes.add_shape(1, 0, 0, 1_000_000, 1_000_000, 1_000_000)  # MSO_AUTO_SHAPE_TYPE = 1 (rectangle)
    return shape, slide

def _roundtrip(prs):
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return Presentation(buf)

def test_hyperlink_external_url_with_anchor_roundtrip():
    prs = Presentation()
    shape, _ = _shape(prs)
    hl = shape.click_action.hyperlink
    hl.address = "https://example.com/page#section"
    prs2 = _roundtrip(prs)
    hl2 = prs2.slides[0].shapes[0].click_action.hyperlink
    assert "example.com/page#section" in (hl2.address or "")

def test_hyperlink_file_uri_roundtrip(tmp_path):
    prs = Presentation()
    shape, _ = _shape(prs)
    target = tmp_path / "other.pptx"
    target.write_bytes(b"")  # just create file path; no open needed
    shape.click_action.hyperlink.address = f"file://{target}"
    prs2 = _roundtrip(prs)
    addr = prs2.slides[0].shapes[0].click_action.hyperlink.address or ""
    assert addr.startswith("file://")

def test_hyperlink_mailto_with_subject_roundtrip():
    prs = Presentation()
    shape, _ = _shape(prs)
    shape.click_action.hyperlink.address = "mailto:abc@example.com?subject=Hello"
    prs2 = _roundtrip(prs)
    addr = prs2.slides[0].shapes[0].click_action.hyperlink.address or ""
    assert addr.startswith("mailto:") and "subject=Hello" in addr

def test_hyperlink_slide_anchor_via_sub_address_roundtrip():
    prs = Presentation()
    # make two slides so there's something to anchor to
    s1 = prs.slides.add_slide(prs.slide_layouts[5])
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    # anchor from slide-0 shape to slide-1
    shape = s1.shapes.add_shape(1, 0, 0, 1_000_000, 1_000_000, 1_000_000)
    hl = shape.click_action.hyperlink
    # Many builds use sub_address for internal anchors (PowerPoint resolves to slide)
    hl.sub_address = s2.slide_id  # sub_address is commonly used for internal targets
    prs2 = _roundtrip(prs)
    sub = prs2.slides[0].shapes[-1].click_action.hyperlink.sub_address
    assert sub is not None

def test_hyperlink_invalid_target_graceful():
    prs = Presentation()
    shape, _ = _shape(prs)
    shape.click_action.hyperlink.address = "nota://protocol"
    prs2 = _roundtrip(prs)
    assert prs2.slides[0].shapes[0].click_action.hyperlink.address.startswith("nota://")
