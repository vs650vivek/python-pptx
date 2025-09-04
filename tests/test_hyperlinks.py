import io
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # Blank

def _shape(prs):
    slide = _blank_slide(prs)
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 1_000_000, 1_000_000
    )
    return shape, slide

def _roundtrip(prs):
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return Presentation(buf)

def _last_shape(slide):
    return slide.shapes[-1]

def test_hyperlink_external_url_with_anchor_roundtrip():
    prs = Presentation()
    shape, _ = _shape(prs)
    hl = shape.click_action.hyperlink
    hl.address = "https://example.com/page#section"
    prs2 = _roundtrip(prs)
    shp2 = _last_shape(prs2.slides[0])
    hl2 = shp2.click_action.hyperlink
    assert (hl2.address or "").endswith("example.com/page#section") or "example.com/page#section" in (hl2.address or "")

def test_hyperlink_file_uri_roundtrip(tmp_path: Path):
    prs = Presentation()
    shape, _ = _shape(prs)
    target = tmp_path / "other.pptx"
    target.write_bytes(b"")  # create an empty file
    shape.click_action.hyperlink.address = target.as_uri()  # file://... URI
    prs2 = _roundtrip(prs)
    shp2 = _last_shape(prs2.slides[0])
    addr = shp2.click_action.hyperlink.address or ""
    assert addr.startswith("file://")

def test_hyperlink_mailto_with_subject_roundtrip():
    prs = Presentation()
    shape, _ = _shape(prs)
    shape.click_action.hyperlink.address = "mailto:abc@example.com?subject=Hello"
    prs2 = _roundtrip(prs)
    shp2 = _last_shape(prs2.slides[0])
    addr = shp2.click_action.hyperlink.address or ""
    assert addr.startswith("mailto:") and "subject=Hello" in addr

def test_hyperlink_slide_anchor_via_target_slide_roundtrip():
    prs = Presentation()
    s1 = _blank_slide(prs)
    s2 = _blank_slide(prs)
    shape = s1.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 1_000_000, 1_000_000
    )
    shape.click_action.target_slide = s2
    prs2 = _roundtrip(prs)
    s1r = prs2.slides[0]
    shp2 = s1r.shapes[-1]
    ca2 = shp2.click_action
    assert (ca2.target_slide is not None), "target_slide should roundtrip"

def test_hyperlink_invalid_target_graceful():
    prs = Presentation()
    shape, _ = _shape(prs)
    shape.click_action.hyperlink.address = "nota://protocol"
    prs2 = _roundtrip(prs)
    shp2 = _last_shape(prs2.slides[0])
    addr = shp2.click_action.hyperlink.address or ""
    assert addr.startswith("nota://")
