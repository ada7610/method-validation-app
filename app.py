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

        for idx, row in level1_df.iterrows():
            r = idx + start_l1_sample_row
            ws[f"A{r}"] = str(row.get("Sample Name", ""))
            ws[f"B{r}"] = float(row.get("Concentration", 0)) if pd.notnull(row.get("Concentration")) else 0.0
            ws[f"B{r}"].number_format = "0.0000"

        end_l1_sample_row = max(len(level1_df) + start_l1_sample_row - 1, start_l1_sample_row)

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
            ws[f"A{l2_title_row}"].fill
