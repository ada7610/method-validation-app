import io
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import numpy as np
import streamlit as st

# ==========================================
# 🎨 Page & Sidebar Configuration
# ==========================================
st.set_page_config(
    page_title="Method Validation System | Abdulrahman Alamri",
    page_icon="🧪",
    layout="wide",
)

with st.sidebar:
    st.header("👑 Developer Info")
    st.markdown("""
    **Developed & Designed By:**  
    👨‍🔬 **Abdulrahman Alamri**  
    
    **Specialization:**  
    🔬 Quality Control & Method Validation Specialist  
    
    ---
    *All Rights Reserved © 2026*
    """)


# ==========================================
# 📊 Excel Generation Function
# ==========================================
def generate_validation_excel(
    calib_df, level1_df, test_title, unit_str, target_conc, t_val, std_purity,
    lod_method, extra_params
):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validation Report"
    ws.views.sheetView[0].showGridLines = True

    COLOR_GREEN_HEADER = "C6EFCE"
    COLOR_BLUE_HEADER = "8EA9DB"
    COLOR_ORANGE_HEADER = "F4B084"
    COLOR_GRAY_NOTE = "D9D9D9"
    COLOR_PURPLE_HEADER = "7030A0"
    COLOR_PURPLE_SUB = "D9E1F2"

    font_main_title = Font(name="Calibri", size=14, bold=True)
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=10)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=thin_border_side
    )

    # 1. Main File Title
    title_text = f"Calculation of Validation for {test_title}" if test_title else "Calculation of Validation"
    ws.merge_cells("A1:N1")
    ws["A1"] = title_text
    ws["A1"].font = font_main_title
    ws["A1"].fill = PatternFill("solid", fgColor=COLOR_GREEN_HEADER)
    ws["A1"].alignment = align_center

    # 2. Calibration Table
    ws.merge_cells("A4:C4")
    ws["A4"] = "Calibration STD"
    ws["A4"].font = font_bold
    ws["A4"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws["A4"].alignment = align_center

    unit_header = f"concentration ({unit_str})" if unit_str else "concentration"
    ws["A5"] = "Level"
    ws["B5"] = unit_header
    ws["C5"] = "area"

    for col in ["A", "B", "C"]:
        ws[f"{col}5"].font = font_bold
        ws[f"{col}5"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"{col}5"].alignment = align_center

    start_cal_row = 6
    for idx, row in calib_df.iterrows():
        r = idx + start_cal_row
        ws[f"A{r}"] = str(row.get("Level", ""))
        ws[f"B{r}"] = float(row.get("Concentration", 0)) if pd.notnull(row.get("Concentration")) else 0.0
        ws[f"C{r}"] = float(row.get("Area", 0)) if pd.notnull(row.get("Area")) else 0.0
        ws[f"B{r}"].number_format = "0.0000"
        ws[f"C{r}"].number_format = "0.0000"

    end_cal_row = max(len(calib_df) + start_cal_row - 1, start_cal_row)

    # 3. Grubbs Critical Values Table (Columns M & N)
    ws.merge_cells("M4:N4")
    ws["M4"] = "Critical values of G (P=0.05)"
    ws["M4"].font = font_white_bold
    ws["M4"].fill = PatternFill("solid", fgColor=COLOR_PURPLE_HEADER)
    ws["M4"].alignment = align_center

    ws["M5"] = "Sample size"
    ws["N5"] = "Critical value"
    for col_ref in ["M5", "N5"]:
        ws[col_ref].font = font_bold
        ws[col_ref].fill = PatternFill("solid", fgColor=COLOR_PURPLE_SUB)
        ws[col_ref].alignment = align_center
        ws[col_ref].border = thin_border

    grubbs_table = [
        (3, 1.155), (4, 1.481), (5, 1.715), (6, 1.887),
        (7, 2.020), (8, 2.126), (9, 2.215), (10, 2.290)
    ]

    last_grubbs_row = 5
    for row_idx, (n_val, g_crit) in enumerate(grubbs_table, 6):
        ws[f"M{row_idx}"] = n_val
        ws[f"N{row_idx}"] = g_crit
        ws[f"M{row_idx}"].alignment = align_center
        ws[f"N{row_idx}"].number_format = "0.0000"
        for c in ["M", "N"]:
            ws[f"{c}{row_idx}"].font = font_regular
            ws[f"{c}{row_idx}"].border = thin_border
        last_grubbs_row = row_idx

    ref_row = last_grubbs_row + 1
    ws.merge_cells(f"M{ref_row}:N{ref_row+1}")
    ws[f"M{ref_row}"] = "Ref: Miller& Miller, Statistics and chemometrics for analytical chemistry, 6th edition."
    ws[f"M{ref_row}"].font = Font(name="Calibri", size=9, italic=True, bold=True)
    ws[f"M{ref_row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 3.1 Student's t-Critical Values Table
    t_table_start_row = ref_row + 3
    ws.merge_cells(f"M{t_table_start_row}:N{t_table_start_row}")
    ws[f"M{t_table_start_row}"] = "Critical values of t (95% Confidence)"
    ws[f"M{t_table_start_row}"].font = font_white_bold
    ws[f"M{t_table_start_row}"].fill = PatternFill("solid", fgColor=COLOR_PURPLE_HEADER)
    ws[f"M{t_table_start_row}"].alignment = align_center

    t_header_row = t_table_start_row + 1
    ws[f"M{t_header_row}"] = "Sample size (n)"
    ws[f"N{t_header_row}"] = "t-Critical"
    for col_ref in [f"M{t_header_row}", f"N{t_header_row}"]:
        ws[col_ref].font = font_bold
        ws[col_ref].fill = PatternFill("solid", fgColor=COLOR_PURPLE_SUB)
        ws[col_ref].alignment = align_center
        ws[col_ref].border = thin_border

    t_crit_table = [
        (1, 12.706), (2, 4.303), (3, 3.182), (4, 2.776), (5, 2.571),
        (6, 2.447), (7, 2.365), (8, 2.306), (9, 2.262), (10, 2.228)
    ]

    last_t_row = t_header_row
    for idx, (n_val, t_crit) in enumerate(t_crit_table):
        row_idx = t_header_row + 1 + idx
        ws[f"M{row_idx}"] = n_val
        ws[f"N{row_idx}"] = t_crit
        ws[f"M{row_idx}"].alignment = align_center
        ws[f"N{row_idx}"].number_format = "0.0000"
        for c in ["M", "N"]:
            ws[f"{c}{row_idx}"].font = font_regular
            ws[f"{c}{row_idx}"].border = thin_border
        last_t_row = row_idx

    t_ref_row = last_t_row + 1
    ws.merge_cells(f"M{t_ref_row}:N{t_ref_row+1}")
    ws[f"M{t_ref_row}"] = "Ref: Student's t-distribution critical values table (95% Confidence)."
    ws[f"M{t_ref_row}"].font = Font(name="Calibri", size=9, italic=True, bold=True)
    ws[f"M{t_ref_row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 4. Scatter Chart
    try:
        chart = ScatterChart()
        chart.title = None
        chart.x_axis.title = None
        chart.y_axis.title = None
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.crosses = "autoZero"
        chart.y_axis.crosses = "autoZero"
        chart.x_axis.tickLblPos = "nextTo"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.number_format = "0.00"
        chart.y_axis.number_format = "0.00"
        chart.graphicalProperties = GraphicalProperties()
        chart.graphicalProperties.noFill = True
        chart.plot_area.graphicalProperties = GraphicalProperties()
        chart.plot_area.graphicalProperties.noFill = True

        chart.x_axis.majorGridlines = ChartLines()
        chart.x_axis.majorGridlines.graphicalProperties = GraphicalProperties()
        chart.x_axis.majorGridlines.graphicalProperties.line = LineProperties(solidFill="D9D9D9")

        chart.y_axis.majorGridlines = ChartLines()
        chart.y_axis.majorGridlines.graphicalProperties = GraphicalProperties()
        chart.y_axis.majorGridlines.graphicalProperties.line = LineProperties(solidFill="D9D9D9")

        xvalues = Reference(ws, min_col=2, min_row=start_cal_row, max_row=end_cal_row)
        yvalues = Reference(ws, min_col=3, min_row=start_cal_row, max_row=end_cal_row)

        series = Series(yvalues, xvalues, title_from_data=False)
        series.marker.symbol = "circle"
        series.marker.size = 6
        series.graphicalProperties.line.noFill = True
        series.dataLabels = DataLabelList()

        trendline = openpyxl.chart.trendline.Trendline(
            trendlineType="linear", dispEq=True, dispRSqr=True
        )
        series.trendline = trendline
        chart.series.append(series)
        chart.legend = None
        chart.width = 13
        chart.height = 7
        ws.add_chart(chart, "F3")
    except Exception:
        pass

    # 5. Regression Parameters
    rsq_row = end_cal_row + 2
    slope_row = rsq_row + 1
    tval_row = slope_row + 1
    spiked_row = tval_row + 1

    ws[f"A{rsq_row}"] = "RSQ (R²)"
    ws[f"B{rsq_row}"] = f"=RSQ(C{start_cal_row}:C{end_cal_row}, B{start_cal_row}:B{end_cal_row})"

    ws[f"A{slope_row}"] = "Slope (S)"
    ws[f"B{slope_row}"] = f"=SLOPE(C{start_cal_row}:C{end_cal_row}, B{start_cal_row}:B{end_cal_row})"

    ws[f"A{tval_row}"] = "t-value (t test)"
    ws[f"B{tval_row}"] = float(t_val)

    ws[f"A{spiked_row}"] = "Spiked Level"
    ws[f"B{spiked_row}"] = float(target_conc)

    for r in [rsq_row, slope_row, tval_row, spiked_row]:
        ws[f"A{r}"].font = font_bold
        ws[f"A{r}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"A{r}"].alignment = align_left
        ws[f"B{r}"].font = font_bold
        ws[f"B{r}"].alignment = align_center
        ws[f"B{r}"].number_format = "0.0000"

    # 6. Level 1 Table
    l1_title_row = spiked_row + 3
    l1_header_row = l1_title_row + 1
    start_sample_row = l1_header_row + 1

    ws.merge_cells(f"A{l1_title_row}:E{l1_title_row}")
    ws[f"A{l1_title_row}"] = f"Level 1 ({unit_str})" if unit_str else "Level 1"
    ws[f"A{l1_title_row}"].font = font_bold
    ws[f"A{l1_title_row}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws[f"A{l1_title_row}"].alignment = align_center

    headers_l1 = ["Samples name", unit_header, "Recovery %", "Outlier (Z-Score)", "Outlier Status"]
    for c, h in zip(["A", "B", "C", "D", "E"], headers_l1):
        ws[f"{c}{l1_header_row}"] = h
        ws[f"{c}{l1_header_row}"].font = font_bold
        ws[f"{c}{l1_header_row}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"{c}{l1_header_row}"].alignment = align_center

    for idx, row in level1_df.iterrows():
        r = idx + start_sample_row
        ws[f"A{r}"] = str(row.get("Sample Name", ""))
        ws[f"B{r}"] = float(row.get("Concentration", 0)) if pd.notnull(row.get("Concentration")) else 0.0
        ws[f"B{r}"].number_format = "0.0000"

    end_sample_row = max(len(level1_df) + start_sample_row - 1, start_sample_row)

    # 7. Statistics & Calculations
