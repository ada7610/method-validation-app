التحديث الأخير للكود بإضافة 2 ليفل 
import io
import streamlit as st

# ==========================================
# 🎨 Page & Sidebar Configuration (MUST BE FIRST)
# ==========================================
st.set_page_config(
    page_title="Method Validation System | Abdulrahman Alamri",
    page_icon="🧪",
    layout="wide",
)

# Safe Environment & Libraries Import Check
try:
    import openpyxl
    from openpyxl.chart import Reference, ScatterChart, Series
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    import pandas as pd
    import numpy as np
except ImportError as imp_err:
    st.error(f"تأكد من تثبيت المكتبات المطلوبة عبر التيرمنال: pip install streamlit pandas numpy openpyxl الخطأ: {imp_err}")
    st.stop()

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
    calib_df, level1_df, level2_df, has_level2, test_title, unit_str, 
    target_conc_l1, target_conc_l2, std_purity, lod_method, extra_params
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
    if calib_df is not None and not calib_df.empty:
        for idx, row in calib_df.iterrows():
            r = idx + start_cal_row
            ws[f"A{r}"] = str(row.get("Level", ""))
            ws[f"B{r}"] = float(row.get("Concentration", 0)) if pd.notnull(row.get("Concentration")) else 0.0
            ws[f"C{r}"] = float(row.get("Area", 0)) if pd.notnull(row.get("Area")) else 0.0
            ws[f"B{r}"].number_format = "0.0000"
            ws[f"C{r}"].number_format = "0.0000"
        end_cal_row = max(len(calib_df) + start_cal_row - 1, start_cal_row)
    else:
        end_cal_row = start_cal_row

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
    spiked_l1_row = slope_row + 1

    ws[f"A{rsq_row}"] = "RSQ (R²)"
    ws[f"B{rsq_row}"] = f"=RSQ(C{start_cal_row}:C{end_cal_row}, B{start_cal_row}:B{end_cal_row})"

    ws[f"A{slope_row}"] = "Slope (S)"
    ws[f"B{slope_row}"] = f"=SLOPE(C{start_cal_row}:C{end_cal_row}, B{start_cal_row}:B{end_cal_row})"

    ws[f"A{spiked_l1_row}"] = "Spiked Level 1"
    ws[f"B{spiked_l1_row}"] = float(target_conc_l1)

    spiked_l2_row = spiked_l1_row + 1
    if has_level2:
        ws[f"A{spiked_l2_row}"] = "Spiked Level 2"
        ws[f"B{spiked_l2_row}"] = float(target_conc_l2)

    for r in [rsq_row, slope_row, spiked_l1_row] + ([spiked_l2_row] if has_level2 else []):
        ws[f"A{r}"].font = font_bold
        ws[f"A{r}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"A{r}"].alignment = align_left
        ws[f"B{r}"].font = font_bold
        ws[f"B{r}"].alignment = align_center
        ws[f"B{r}"].number_format = "0.0000"

    # 6. Level 1 Table
    l1_title_row = (spiked_l2_row if has_level2 else spiked_l1_row) + 2
    l1_header_row = l1_title_row + 1
    start_l1_sample_row = l1_header_row + 1

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

    if level1_df is not None and not level1_df.empty:
        for idx, row in level1_df.iterrows():
            r = idx + start_l1_sample_row
            ws[f"A{r}"] = str(row.get("Sample Name", ""))
            ws[f"B{r}"] = float(row.get("Concentration", 0)) if pd.notnull(row.get("Concentration")) else 0.0
            ws[f"B{r}"].number_format = "0.0000"
        end_l1_sample_row = max(len(level1_df) + start_l1_sample_row - 1, start_l1_sample_row)
    else:
        end_l1_sample_row = start_l1_sample_row

    # Level 1 Statistics & Calculations
    l1_stats_start = end_l1_sample_row + 2
    l1_mean_row = l1_stats_start
    l1_rec_row = l1_stats_start + 1
    l1_sd_row = l1_stats_start + 2
    l1_rsd_row = l1_stats_start + 3
    l1_lod_row = l1_stats_start + 4
    l1_loq_row = l1_stats_start + 5

    # Dynamic formulas for LOD/LOQ
    if lod_method == "t-Statistic & SD (Default)":
        t_val_default = extra_params.get("t_val", 2.571) if isinstance(extra_params, dict) else 2.571
        lod_formula = f"={t_val_default}*B{l1_sd_row}"
        loq_formula = f"=10*B{l1_sd_row}"
    elif lod_method == "Slope & SD of Intercept (Sa)":
        sa_val = extra_params.get("sa", 0.0) if isinstance(extra_params, dict) else 0.0
        lod_formula = f"=(3.3*{sa_val})/B{slope_row}"
        loq_formula = f"=(10*{sa_val})/B{slope_row}"
    elif lod_method == "Slope & Residual SD (Sy/x)":
        syx_val = extra_params.get("syx", 0.0) if isinstance(extra_params, dict) else 0.0
        lod_formula = f"=(3.3*{syx_val})/B{slope_row}"
        loq_formula = f"=(10*{syx_val})/B{slope_row}"
    elif lod_method == "Signal-to-Noise Ratio (S/N)":
        noise_val = extra_params.get("noise", 1.0) if isinstance(extra_params, dict) else 1.0
        lod_formula = f"=3*{noise_val}"
        loq_formula = f"=10*{noise_val}"
    elif lod_method == "Standard Deviation of Blanks":
        blank_sd_val = extra_params.get("blank_sd", 0.0) if isinstance(extra_params, dict) else 0.0
        blank_mean_val = extra_params.get("blank_mean", 0.0) if isinstance(extra_params, dict) else 0.0
        lod_formula = f"={blank_mean_val}+(3*{blank_sd_val})"
        loq_formula = f"={blank_mean_val}+(10*{blank_sd_val})"
    else:
        manual_lod = extra_params.get("manual_lod", 0.0) if isinstance(extra_params, dict) else 0.0
        manual_loq = extra_params.get("manual_loq", 0.0) if isinstance(extra_params, dict) else 0.0
        lod_formula = float(manual_lod)
        loq_formula = float(manual_loq)

    l1_stats_labels = [
        ("Mean", f"=AVERAGE(B{start_l1_sample_row}:B{end_l1_sample_row})"),
        ("Recovery %", f"=AVERAGE(C{start_l1_sample_row}:C{end_l1_sample_row})"),
        ("Standerd Deviation", f"=STDEV(B{start_l1_sample_row}:B{end_l1_sample_row})"),
        ("RSD %", f"=IF(B{l1_mean_row}=0, 0, (B{l1_sd_row}/B{l1_mean_row})*100)"),
        ("LOD", lod_formula),
        ("LOQ", loq_formula),
    ]

    for i, (label, formula) in enumerate(l1_stats_labels, start=l1_stats_start):
        ws[f"A{i}"] = label
        ws[f"A{i}"].font = font_bold
        ws[f"A{i}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
        ws[f"B{i}"] = formula
        ws[f"B{i}"].number_format = "0.0000"

    for r in range(start_l1_sample_row, end_l1_sample_row + 1):
        ws[f"C{r}"] = f"=IF(B{spiked_l1_row}=0, 0, (B{r}/B{spiked_l1_row})*100)"
        ws[f"D{r}"] = f"=ABS(B{r}-B${l1_mean_row})/B${l1_sd_row}"
        ws[f"E{r}"] = f'=IF(D{r}>VLOOKUP(COUNT(B${start_l1_sample_row}:B${end_l1_sample_row}), M$6:N$13, 2, FALSE), "Outlier", "Normal")'
        ws[f"C{r}"].number_format = "0.0000"
        ws[f"D{r}"].number_format = "0.0000"
        ws[f"E{r}"].alignment = align_center

    ws.merge_cells(f"F{start_l1_sample_row}:F{end_l1_sample_row}")
    ws[f"F{start_l1_sample_row}"] = "Any value higher than the critical value in the table is consider outlier"
    ws[f"F{start_l1_sample_row}"].font = Font(name="Calibri", size=9, bold=True)
    ws[f"F{start_l1_sample_row}"].fill = PatternFill("solid", fgColor=COLOR_GRAY_NOTE)
    ws[f"F{start_l1_sample_row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Measurement Uncertainty for Level 1
    unc_header_row = l1_header_row
    unc_start_row = start_l1_sample_row

    ws.merge_cells(f"H{unc_header_row}:I{unc_header_row}")
    ws[f"H{unc_header_row}"] = "Measurment uncertainty"
    ws[f"H{unc_header_row}"].font = font_bold
    ws[f"H{unc_header_row}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
    ws[f"H{unc_header_row}"].alignment = align_center

    ws.merge_cells(f"H{unc_start_row}:I{unc_start_row}")
    ws[f"H{unc_start_row}"] = "Level 1"
    ws[f"H{unc_start_row}"].font = font_bold
    ws[f"H{unc_start_row}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws[f"H{unc_start_row}"].alignment = align_center

    purity_row = unc_header_row - 1
    ws[f"H{purity_row}"] = "Standard Purity"
    ws[f"H{purity_row}"].font = font_bold
    purity_val = std_purity / 100.0 if std_purity > 1.0 else std_purity
    ws[f"I{purity_row}"] = purity_val
    ws[f"I{purity_row}"].number_format = "0.0000"

    uA_row = unc_start_row + 1
    uB_row = uA_row + 1
    uC_row = uB_row + 1
    uD_row = uC_row + 1
    uComb_row = uD_row + 1
    uExp_row = uComb_row + 1
    uPct_row = uExp_row + 1

    l1_unc_labels = [
        (uA_row, "uA", f"=B{l1_rsd_row}/100"),
        (uB_row, "uB", f"=ABS(0.5*(1 - (B{l1_rec_row}/100)))/SQRT(3)"),
        (uC_row, "uC", f"=0.5*(1 - I{purity_row})/SQRT(3)"),
        (uD_row, "uD", f"=1 - SQRT(B{rsq_row})"),
        (uComb_row, "u combiend", f"=SQRT(I{uA_row}^2 + I{uB_row}^2 + I{uC_row}^2 + I{uD_row}^2)"),
        (uExp_row, "U expanded", f"=2*I{uComb_row}"),
        (uPct_row, "U%", f"=I{uExp_row}*100"),
    ]

    for r_num, u_name, u_formula in l1_unc_labels:
        ws[f"H{r_num}"] = u_name
        ws[f"I{r_num}"] = u_formula
        ws[f"I{r_num}"].number_format = "0.0000"
        if u_name in ["u combiend", "U expanded", "U%"]:
            ws[f"H{r_num}"].font = font_bold
            ws[f"I{r_num}"].font = font_bold

    current_max_row = max(l1_stats_start + 6, uPct_row + 2)

    # 7. Level 2 Table (Optional - Without LOD/LOQ)
    if has_level2 and level2_df is not None and not level2_df.empty:
        l2_title_row = current_max_row + 2
        l2_header_row = l2_title_row + 1
        start_l2_sample_row = l2_header_row + 1

        ws.merge_cells(f"A{l2_title_row}:E{l2_title_row}")
        ws[f"A{l2_title_row}"] = f"Level 2 ({unit_str})" if unit_str else "Level 2"
        ws[f"A{l2_title_row}"].font = font_bold
        ws[f"A{l2_title_row}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
        ws[f"A{l2_title_row}"].alignment = align_center

        for c, h in zip(["A", "B", "C", "D", "E"], headers_l1):
            ws[f"{c}{l2_header_row}"] = h
            ws[f"{c}{l2_header_row}"].font = font_bold
            ws[f"{c}{l2_header_row}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
            ws[f"{c}{l2_header_row}"].alignment = align_center

        for idx, row in level2_df.iterrows():
            r = idx + start_l2_sample_row
            ws[f"A{r}"] = str(row.get("Sample Name", ""))
            ws[f"B{r}"] = float(row.get("Concentration", 0)) if pd.notnull(row.get("Concentration")) else 0.0
            ws[f"B{r}"].number_format = "0.0000"

        end_l2_sample_row = max(len(level2_df) + start_l2_sample_row - 1, start_l2_sample_row)

        l2_stats_start = end_l2_sample_row + 2
        l2_mean_row = l2_stats_start
        l2_rec_row = l2_stats_start + 1
        l2_sd_row = l2_stats_start + 2
        l2_rsd_row = l2_stats_start + 3

        l2_stats_labels = [
            ("Mean", f"=AVERAGE(B{start_l2_sample_row}:B{end_l2_sample_row})"),
            ("Recovery %", f"=AVERAGE(C{start_l2_sample_row}:C{end_l2_sample_row})"),
            ("Standerd Deviation", f"=STDEV(B{start_l2_sample_row}:B{end_l2_sample_row})"),
            ("RSD %", f"=IF(B{l2_mean_row}=0, 0, (B{l2_sd_row}/B{l2_mean_row})*100)"),
        ]

        for i, (label, formula) in enumerate(l2_stats_labels, start=l2_stats_start):
            ws[f"A{i}"] = label
            ws[f"A{i}"].font = font_bold
            ws[f"A{i}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
            ws[f"B{i}"] = formula
            ws[f"B{i}"].number_format = "0.0000"

        for r in range(start_l2_sample_row, end_l2_sample_row + 1):
            ws[f"C{r}"] = f"=IF(B{spiked_l2_row}=0, 0, (B{r}/B{spiked_l2_row})*100)"
            ws[f"D{r}"] = f"=ABS(B{r}-B${l2_mean_row})/B${l2_sd_row}"
            ws[f"E{r}"] = f'=IF(D{r}>VLOOKUP(COUNT(B${start_l2_sample_row}:B${end_l2_sample_row}), M$6:N$13, 2, FALSE), "Outlier", "Normal")'
            ws[f"C{r}"].number_format = "0.0000"
            ws[f"D{r}"].number_format = "0.0000"
            ws[f"E{r}"].alignment = align_center

        ws.merge_cells(f"F{start_l2_sample_row}:F{end_l2_sample_row}")
        ws[f"F{start_l2_sample_row}"] = "Any value higher than the critical value in the table is consider outlier"
        ws[f"F{start_l2_sample_row}"].font = Font(name="Calibri", size=9, bold=True)
        ws[f"F{start_l2_sample_row}"].fill = PatternFill("solid", fgColor=COLOR_GRAY_NOTE)
        ws[f"F{start_l2_sample_row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Measurement Uncertainty for Level 2
        l2_unc_start = start_l2_sample_row
        ws.merge_cells(f"H{l2_unc_start-1}:I{l2_unc_start-1}")
        ws[f"H{l2_unc_start-1}"] = "Level 2"
        ws[f"H{l2_unc_start-1}"].font = font_bold
        ws[f"H{l2_unc_start-1}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
        ws[f"H{l2_unc_start-1}"].alignment = align_center

        luA_row = l2_unc_start + 1
        luB_row = luA_row + 1
        luC_row = luB_row + 1
        luD_row = luC_row + 1
        luComb_row = luD_row + 1
        luExp_row = luComb_row + 1
        luPct_row = luExp_row + 1

        l2_unc_labels = [
            (luA_row, "uA", f"=B{l2_rsd_row}/100"),
            (luB_row, "uB", f"=ABS(0.5*(1 - (B{l2_rec_row}/100)))/SQRT(3)"),
            (luC_row, "uC", f"=0.5*(1 - I{purity_row})/SQRT(3)"),
            (luD_row, "uD", f"=1 - SQRT(B{rsq_row})"),
            (luComb_row, "u combiend", f"=SQRT(I{luA_row}^2 + I{luB_row}^2 + I{luC_row}^2 + I{luD_row}^2)"),
            (luExp_row, "U expanded", f"=2*I{luComb_row}"),
            (luPct_row, "U%", f"=I{luExp_row}*100"),
        ]

        for r_num, u_name, u_formula in l2_unc_labels:
            ws[f"H{r_num}"] = u_name
            ws[f"I{r_num}"] = u_formula
            ws[f"I{r_num}"].number_format = "0.0000"
            if u_name in ["u combiend", "U expanded", "U%"]:
                ws[f"H{r_num}"].font = font_bold
                ws[f"I{r_num}"].font = font_bold

    column_widths = {
        "A": 26, "B": 24, "C": 18, "D": 18, "E": 18,
        "F": 25, "H": 22, "I": 18, "M": 25, "N": 25
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==========================================
# ⚙️ Streamlit User Interface (Fully Protected)
# ==========================================
try:
    st.title("🧪 Analytical Method Validation System")

    col_header1, col_header2 = st.columns(2)
    with col_header1:
        test_name = st.text_input("Test / Analysis Name (e.g. Benzo a pyrene)", "", key="test_name_input")
    with col_header2:
        conc_unit = st.text_input("Concentration Unit", "", key="conc_unit_input")

    st.divider()

    st.subheader("📌 Calibration Standard Table (Calibration STD)")
    num_calib_levels = st.number_input(
        "Number of Calibration Levels", min_value=1, max_value=30, value=6, step=1, key="calib_num_input"
    )

    calib_data = [
        {
            "Level": f"STD {i+1}",
            "Concentration": 0.0000,
            "Area": 0.0000
        }
        for i in range(int(num_calib_levels))
    ]

    valid_std_raw = st.data_editor(
        pd.DataFrame(calib_data), num_rows="dynamic", key="calib_table_editor", use_container_width=True
    )

    valid_std = valid_std_raw.dropna(how="all").copy() if valid_std_raw is not None else pd.DataFrame(calib_data)
    if "Level" in valid_std.columns:
        valid_std["Level"] = valid_std["Level"].fillna("")
    if "Concentration" in valid_std.columns:
        valid_std["Concentration"] = pd.to_numeric(valid_std["Concentration"], errors="coerce").fillna(0.0)
    if "Area" in valid_std.columns:
        valid_std["Area"] = pd.to_numeric(valid_std["Area"], errors="coerce").fillna(0.0)

    st.divider()

    # ==========================================
    # 🎯 Selection of LOD & LOQ Calculation Method
    # ==========================================
    st.subheader("🎯 LOD & LOQ Calculation Method")
    st.info("Choose the calculation method compliant with your ICH Q2 / USP / Validation Protocol requirements:")

    lod_method = st.selectbox(
        "Select Calculation Method",
        options=[
            "t-Statistic & SD (Default)",
            "Slope & SD of Intercept (Sa)",
            "Slope & Residual SD (Sy/x)",
            "Signal-to-Noise Ratio (S/N)",
            "Standard Deviation of Blanks",
            "Manual / Visual / Low Recovery",
        ],
        index=0,
        key="lod_method_select"
    )

    extra_params = {}

    if lod_method == "t-Statistic & SD (Default)":
        extra_params["t_val"] = st.number_input("Enter t-statistic value", value=2.5710, min_value=0.0, format="%.4f", key="t_val_input")

    st.divider()

    # ==========================================
    # 📋 Level 1 Inputs & Table
    # ==========================================
    st.subheader("📋 Level 1 Samples & Inputs")

    col_input1, col_input2 = st.columns(2)

    with col_input1:
        target_conc_l1 = st.number_input("Spiked Level 1", value=0.0000, min_value=0.0, format="%.4f", key="target_conc_l1_input")

    with col_input2:
        std_purity = st.number_input("Standard Purity (%)", value=0.0000, min_value=0.0, max_value=100.0, format="%.4f", key="std_purity_input")

    num_samples_l1 = st.number_input(
        "Number of Replicates / Samples (Level 1)", min_value=1, max_value=30, value=6, step=1, key="samples_num_l1_input"
    )

    sample_data_l1 = [
        {"Sample Name": f"Sample {i+1}", "Concentration": 0.0000} for i in range(int(num_samples_l1))
    ]

    edited_samples_l1_raw = st.data_editor(
        pd.DataFrame(sample_data_l1), num_rows="dynamic", key="samples_table_l1_editor", use_container_width=True
    )

    edited_samples_l1 = edited_samples_l1_raw.dropna(how="all").copy() if edited_samples_l1_raw is not None else pd.DataFrame(sample_data_l1)
    if "Sample Name" in edited_samples_l1.columns:
        edited_samples_l1["Sample Name"] = edited_samples_l1["Sample Name"].fillna("")
    if "Concentration" in edited_samples_l1.columns:
        edited_samples_l1["Concentration"] = pd.to_numeric(edited_samples_l1["Concentration"], errors="coerce").fillna(0.0)

    st.divider()

    # ==========================================
    # ➕ Level 2 Optional Addition
    # ==========================================
    has_level2 = st.checkbox("➕ Add Level 2 Samples Table", value=False, key="has_level2_checkbox")

    target_conc_l2 = 0.0000
    edited_samples_l2 = None

    if has_level2:
        st.subheader("📋 Level 2 Samples & Inputs")
        target_conc_l2 = st.number_input("Spiked Level 2", value=0.0000, min_value=0.0, format="%.4f", key="target_conc_l2_input")

        num_samples_l2 = st.number_input(
            "Number of Replicates / Samples (Level 2)", min_value=1, max_value=30, value=6, step=1, key="samples_num_l2_input"
        )

        sample_data_l2 = [
            {"Sample Name": f"Sample {i+1}", "Concentration": 0.0000} for i in range(int(num_samples_l2))
        ]

        edited_samples_l2_raw = st.data_editor(
            pd.DataFrame(sample_data_l2), num_rows="dynamic", key="samples_table_l2_editor", use_container_width=True
        )

        edited_samples_l2 = edited_samples_l2_raw.dropna(how="all").copy() if edited_samples_l2_raw is not None else pd.DataFrame(sample_data_l2)
        if "Sample Name" in edited_samples_l2.columns:
            edited_samples_l2["Sample Name"] = edited_samples_l2["Sample Name"].fillna("")
        if "Concentration" in edited_samples_l2.columns:
            edited_samples_l2["Concentration"] = pd.to_numeric(edited_samples_l2["Concentration"], errors="coerce").fillna(0.0)

        st.divider()

    # ==========================================
    # 🧮 Calculations & Metrics (Level 1)
    # ==========================================
    calib_x = valid_std["Concentration"].values if not valid_std.empty and "Concentration" in valid_std.columns else np.array([0.0])
    calib_y = valid_std["Area"].values if not valid_std.empty and "Area" in valid_std.columns else np.array([0.0])

    n_calib = len(calib_x)
    slope = 1.0
    syx_auto = 0.0
    sa_auto = 0.0

    if n_calib > 2:
        try:
            valid_mask = (calib_x > 0) | (calib_y > 0)
            cx = calib_x[valid_mask]
            cy = calib_y[valid_mask]
            if len(cx) > 2 and np.sum((cx - np.mean(cx))**2) != 0:
                slope, intercept = np.polyfit(cx, cy, 1)
                y_pred = slope * cx + intercept
                residuals = cy - y_pred
                sum_sq_residuals = np.sum(residuals**2)
                df = len(cx) - 2
                syx_auto = np.sqrt(sum_sq_residuals / df)
                
                sum_sq_x = np.sum((cx - np.mean(cx))**2)
                sum_x2 = np.sum(cx**2)
                sa_auto = syx_auto * np.sqrt(sum_x2 / (len(cx) * sum_sq_x))
        except Exception:
            pass

    if lod_method == "Slope & SD of Intercept (Sa)":
        st.success(f"✅ SD of Intercept (Sa) calculated automatically: **{sa_auto:.6f}** (Sy/x: {syx_auto:.6f}, Slope: {slope:.6f})")
        extra_params["sa"] = sa_auto

    elif lod_method == "Slope & Residual SD (Sy/x)":
        st.success(f"✅ Residual SD (Sy/x) calculated automatically: **{syx_auto:.6f}**")
        extra_params["syx"] = syx_auto

    elif lod_method == "Signal-to-Noise Ratio (S/N)":
        extra_params["noise"] = st.number_input(
            "Baseline Noise Level", value=1.0000, min_value=0.0001, format="%.6f", key="noise_input",
            help="LOD will be 3 × Noise, LOQ will be 10 × Noise"
        )
    elif lod_method == "Standard Deviation of Blanks":
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            extra_params["blank_mean"] = st.number_input("Mean of Blank Samples", value=0.0000, format="%.6f", key="blank_mean_input")
        with col_b2:
            extra_params["blank_sd"] = st.number_input("Standard Deviation of Blanks (s_blank)", value=0.0000, min_value=0.0, format="%.6f", key="blank_sd_input")
    elif lod_method == "Manual / Visual / Low Recovery":
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            extra_params["manual_lod"] = st.number_input("Manual / Visual LOD Value", value=0.0000, min_value=0.0, format="%.6f", key="manual_lod_input")
        with col_m2:
            extra_params["manual_loq"] = st.number_input("Manual / Visual LOQ Value", value=0.0000, min_value=0.0, format="%.6f", key="manual_loq_input")

    sample_concs_l1 = edited_samples_l1["Concentration"].values if not edited_samples_l1.empty and "Concentration" in edited_samples_l1.columns else np.array([0.0])
    sample_sd_l1 = np.std(sample_concs_l1, ddof=1) if len(sample_concs_l1) > 1 else 0.0

    if lod_method == "t-Statistic & SD (Default)":
        t_val_val = extra_params.get("t_val", 2.571) if isinstance(extra_params, dict) else 2.571
        calc_lod = t_val_val * sample_sd_l1
        calc_loq = 10 * sample_sd_l1
    elif lod_method == "Slope & SD of Intercept (Sa)":
        calc_lod = (3.3 * sa_auto) / slope if slope != 0 else 0
        calc_loq = (10 * sa_auto) / slope if slope != 0 else 0
    elif lod_method == "Slope & Residual SD (Sy/x)":
        calc_lod = (3.3 * syx_auto) / slope if slope != 0 else 0
        calc_loq = (10 * syx_auto) / slope if slope != 0 else 0
    elif lod_method == "Signal-to-Noise Ratio (S/N)":
        noise = extra_params.get("noise", 1.0) if isinstance(extra_params, dict) else 1.0
        calc_lod = 3 * noise
        calc_loq = 10 * noise
    elif lod_method == "Standard Deviation of Blanks":
        calc_lod = extra_params.get("blank_mean", 0.0) + (3 * extra_params.get("blank_sd", 0.0)) if isinstance(extra_params, dict) else 0.0
        calc_loq = extra_params.get("blank_mean", 0.0) + (10 * extra_params.get("blank_sd", 0.0)) if isinstance(extra_params, dict) else 0.0
    else:
        calc_lod = extra_params.get("manual_lod", 0.0) if isinstance(extra_params, dict) else 0.0
        calc_loq = extra_params.get("manual_loq", 0.0) if isinstance(extra_params, dict) else 0.0

    col_lod, col_loq = st.columns(2)
    with col_lod:
        st.metric(label=f"Calculated LOD ({conc_unit})", value=f"{calc_lod:.6f}")
    with col_loq:
        st.metric(label=f"Calculated LOQ ({conc_unit})", value=f"{calc_loq:.6f}")

    st.divider()

    try:
        calib_export = (
            valid_std[["Level", "Concentration", "Area"]]
            if valid_std is not None and not valid_std.empty and "Level" in valid_std.columns
            else pd.DataFrame(columns=["Level", "Concentration", "Area"])
        )

        excel_file = generate_validation_excel(
            calib_df=calib_export,
            level1_df=edited_samples_l1,
            level2_df=edited_samples_l2,
            has_level2=has_level2,
            test_title=test_name,
            unit_str=conc_unit,
            target_conc_l1=target_conc_l1,
            target_conc_l2=target_conc_l2,
            std_purity=std_purity,
            lod_method=lod_method,
            extra_params=extra_params
        )

        st.download_button(
            label="📥 Download Final Validation Excel Report",
            data=excel_file,
            file_name="Method_Validation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"An error occurred while generating the Excel file: {e}")

except Exception as app_err:
    st.error(f"⚠️ Application Crash Prevented: {app_err}")

st.caption("---")
st.caption("Developed with ❤️ by **Abdulrahman Alamri** | All Rights Reserved © 2026")
