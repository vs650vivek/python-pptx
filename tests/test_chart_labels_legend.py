import io
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

def _add_line_chart(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("S1", (1, 2, 3))
    x, y, cx, cy = 1_000_000, 1_000_000, 6_000_000, 4_000_000
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, chart_data).chart
    return chart

def test_datalabels_toggle_and_number_format_roundtrip(tmp_path):
    prs = Presentation()
    chart = _add_line_chart(prs)
    plot = chart.plots[0]
    dlabels = plot.data_labels
    # toggles (example: show value only)
    dlabels.show_value = True
    dlabels.show_category_name = False
    dlabels.show_series_name = False

    # number format
    dlabels.number_format = "#,##0.00"

    # round-trip save+load
    buf = io.BytesIO()
    prs.save(buf); buf.seek(0)
    prs2 = Presentation(buf)
    chart2 = prs2.slides[0].shapes[0].chart
    d2 = chart2.plots[0].data_labels

    assert d2.show_value is True
    assert (d2.show_category_name or False) is False
    assert (d2.show_series_name or False) is False
    assert d2.number_format in ("#,##0.00",)  # PP may normalize but should keep code

def test_line_chart_defaults_legend_and_vary_by_categories():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    data = CategoryChartData()
    data.categories = ["Q1","Q2"]
    data.add_series("S1", (1,2))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS, 0, 0, 6_000_000, 4_000_000, data
    ).chart

    assert chart.has_legend is True
    assert chart.plots[0].vary_by_categories is False