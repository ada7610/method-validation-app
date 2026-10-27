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
    calib_df, level1_df, test_title, unit_str, target_conc, std_purity,
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

    # 3.1 Student's t-Critical Values Table (يظهر فقط إذا كان الخيار t-Statistic & SD)
    last_t_row = ref_row
    if lod_method == "t-Statistic & SD (Default)":
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
        ws[f"M{t_ref_row}"] = "Ref
