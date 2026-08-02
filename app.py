import io
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# ==========================================
# 🎨 إعدادات الصفحة والشريط الجانبي (الحقوق)
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
# 📊 دالة إنشاء ملف Excel المنسق
# ==========================================
def generate_validation_excel(
    calib_df, level1_df, test_title, unit_str, target_conc, t_val, std_purity
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
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=False
    )
    align_left = Alignment(horizontal="left", vertical="center")

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )

    # 1. العنوان الرئيسي
    title_text = (
        f"Calculation of Validation for {test_title}"
        if test_title
        else "Calculation of Validation"
    )
    ws.merge_cells("A1:K1")
    ws["A1"] = title_text
    ws["A1"].font = font_main_title
    ws["A1"].fill = PatternFill("solid", fgColor=COLOR_GREEN_HEADER)
    ws["A1"].alignment = align_center

    # 2. جدول المعايرة
    ws.merge_cells("A4:C4")
    ws["A4"] = "Calibration STD"
    ws["A4"].font = font_bold
    ws["A4"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws["A4"].alignment = align_center

    unit_header = f"Concentration ({unit_str})" if unit_str else "Concentration"
    ws["A5"] = "Level"
    ws["B5"] = unit_header
    ws["C5"] = "Area"

    for col in ["A", "B", "C"]:
        ws[f"{col}5"].font = font_bold
        ws[f"{col}5"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"{col}5"].alignment = align_center

    start_cal_row = 6
    for idx, row in calib_df.iterrows():
        r = idx + start_cal_row
        ws[f"A{r}"] = str(row.get("Level", ""))
        ws[f"B{r}"] = (
            float(row.get("Concentration", 0))
            if pd.notnull(row.get("Concentration"))
            else 0.0
        )
        ws[f"C{r}"] = (
            float(row.get("Area", 0)) if pd.notnull(row.get("Area")) else 0.0
        )

    end_cal_row = max(len(calib_df) + start_cal_row - 1, start_cal_row)

    # 3. جدول القيم الحرجة لـ Grubbs
    ws.merge_cells("J4:K4")
    ws["J4"] = "Critical values of G (P=0.05)"
    ws["J4"].font = font_white_bold
    ws["J4"].fill = PatternFill("solid", fgColor=COLOR_PURPLE_HEADER)
    ws["J4"].alignment = align_center

    ws["J5"] = "Sample size"
    ws["K5"] = "Critical value"
    for col_ref in ["J5", "K5"]:
        ws[col_ref].font = font_bold
        ws[col_ref].fill = PatternFill("solid", fgColor=COLOR_PURPLE_SUB)
        ws[col_ref].alignment = align_center
        ws[col_ref].border = thin_border

    grubbs_table = [
        (3, 1.155),
        (4, 1.481),
        (5, 1.715),
        (6, 1.887),
        (7, 2.020),
        (8, 2.126),
        (9, 2.215),
        (10, 2.290),
    ]

    for row_idx, (n_val, g_crit) in enumerate(grubbs_table, 6):
        ws[f"J{row_idx}"] = n_val
        ws[f"K{row_idx}"] = g_crit
        ws[f"J{row_idx}"].alignment = align_center

        for c in ["J", "K"]:
            ws[f"{c}{row_idx}"].font = font_regular
            ws[f"{c}{row_idx}"].border = thin_border

    # 4. المنحنى القياسي
    chart = ScatterChart()
    chart.title = str(test_title) if test_title else "B1"
    chart.title.overlay = False

    chart.graphicalProperties = GraphicalProperties()
    chart.graphicalProperties.noFill = True

    chart.plot_area.graphicalProperties = GraphicalProperties()
    chart.plot_area.graphicalProperties.noFill = True

    chart.x_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines = ChartLines()

    chart.x_axis.graphicalProperties = GraphicalProperties()
    chart.x_axis.graphicalProperties.line = LineProperties(solidFill="BFBFBF")

    chart.y_axis.graphicalProperties = GraphicalProperties()
    chart.y_axis.graphicalProperties.line = LineProperties(solidFill="BFBFBF")

    chart.x_axis.number_format = "0.0000"
    chart.y_axis.number_format = "0.0000"

    xvalues = Reference(
        ws, min_col=2, min_row=start_cal_row, max_row=end_cal_row
    )
    yvalues = Reference(
        ws, min_col=3, min_row=start_cal_row, max_row=end_cal_row
    )

    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker.symbol = "circle"
    series.marker.size = 6
    series.graphicalProperties.line.noFill = True

    trendline = openpyxl.chart.trendline.Trendline(
        trendlineType="linear", dispEq=True, dispRSqr=True
    )
    try:
        trendline.graphicalProperties = GraphicalProperties()
        trendline.graphicalProperties.line = LineProperties(
            prstDash="sysDot", cmpd="sng"
        )
    except Exception:
        pass

    series.trendline = trendline
    chart.series.append(series)
    chart.legend = None
    chart.width = 13
    chart.height = 7.5
    ws.add_chart(chart, "F3")

    # 5. RSQ و t-value و Spiked Level
    rsq_row = end_cal_row + 2
    tval_row = rsq_row + 1
    spiked_row = tval_row + 1

    ws[f"A{rsq_row}"] = "RSQ"
    ws[f"B{rsq_row}"] = (
        f"=RSQ(C{start_cal_row}:C{end_cal_row}, B{start_cal_row}:B{end_cal_row})"
    )

    ws[f"A{tval_row}"] = "t-value (t test)"
    ws[f"B{tval_row}"] = float(t_val)

    ws[f"A{spiked_row}"] = "Spiked Level"
    ws[f"B{spiked_row}"] = float(target_conc)

    for r in [rsq_row, tval_row, spiked_row]:
        ws[f"A{r}"].font = font_bold
        ws[f"A{r}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"A{r}"].alignment = align_left

        ws[f"B{r}"].font = font_bold
        ws[f"B{r}"].fill = PatternFill(fill_type=None)
        ws[f"B{r}"].alignment = align_center

    # 6. جدول Level 1
    l1_title_row = spiked_row + 2
    l1_header_row = l1_title_row + 1
    start_sample_row = l1_header_row + 1

    ws.merge_cells(f"A{l1_title_row}:E{l1_title_row}")
    ws[f"A{l1_title_row}"] = f"Level 1 ({unit_str})" if unit_str else "Level 1"
    ws[f"A{l1_title_row}"].font = font_bold
    ws[f"A{l1_title_row}"].fill = PatternFill(
        "solid", fgColor=COLOR_BLUE_HEADER
    )
    ws[f"A{l1_title_row}"].alignment = align_center

    headers_l1 = [
        "Samples name",
        unit_header,
        "Recovery %",
        "Outlier (Z-Score)",
        "Outlier Status",
    ]
    cols_l1 = ["A", "B", "C", "D", "E"]
    for c, h in zip(cols_l1, headers_l1):
        ws[f"{c}{l1_header_row}"] = h
        ws[f"{c}{l1_header_row}"].font = font_bold
        ws[f"{c}{l1_header_row}"].fill = PatternFill(
            "solid", fgColor=COLOR_ORANGE_HEADER
        )
        ws[f"{c}{l1_header_row}"].alignment = align_center

    for idx, row in level1_df.iterrows():
        r = idx + start_sample_row
        ws[f"A{r}"] = str(row.get("Sample Name", ""))
        sample_conc = (
            float(row.get("Concentration", 0))
            if pd.notnull(row.get("Concentration"))
            else 0.0
        )
        ws[f"B{r}"] = sample_conc

    end_sample_row = max(len(level1_df) + start_sample_row - 1, start_sample_row)

    # 7. الإحصائيات (تم استخدام =STDEV القياسية لتجنب خطأ #NAME?)
    stats_start_row = end_sample_row + 2
    mean_row = stats_start_row
    rec_row = stats_start_row + 1
    sd_row = stats_start_row + 2
    rsd_row = stats_start_row + 3

    for r in range(start_sample_row, end_sample_row + 1):
        ws[f"C{r}"] = f"=IF(B{spiked_row}=0, 0, (B{r}/B{spiked_row})*100)"
        ws[f"D{r}"] = f"=IF(B${sd_row}=0, 0, ABS(B{r}-B${mean_row})/B${sd_row})"

        ws[f"E{r}"] = (
            f'=IF(D{r}>VLOOKUP(COUNT(B${start_sample_row}:B${end_sample_row}), J$6:K$13, 2, FALSE), "Outlier", "Normal")'
        )
        ws[f"E{r}"].alignment = align_center

    ws.merge_cells(f"F{start_sample_row}:F{end_sample_row}")
    ws[f"F{start_sample_row}"] = (
        "Any value higher than the critical value in the table is consider outlier"
    )
    ws[f"F{start_sample_row}"].font = Font(name="Calibri", size=9, bold=True)
    ws[f"F{start_sample_row}"].fill = PatternFill(
        "solid", fgColor=COLOR_GRAY_NOTE
    )
    ws[f"F{start_sample_row}"].alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    stats_labels = [
        ("Mean", f"=AVERAGE(B{start_sample_row}:B{end_sample_row})"),
        ("Recovery %", f"=AVERAGE(C{start_sample_row}:C{end_sample_row})"),
        (
            "Standerd Deviation",
            f"=STDEV(B{start_sample_row}:B{end_sample_row})",
        ),  # استخدام STDEV المباشرة
        ("RSD %", f"=IF(B{mean_row}=0, 0, (B{sd_row}/B{mean_row})*100)"),
        ("LOD", f"=B{tval_row}*B{sd_row}"),
        ("LOQ", f"=10*B{sd_row}"),
    ]

    for i, (label, formula) in enumerate(
        stats_labels, start=stats_start_row
    ):
        ws[f"A{i}"] = label
        ws[f"A{i}"].font = font_bold
        ws[f"A{i}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
        ws[f"B{i}"] = formula

    # 8. جدول Measurement Uncertainty
    unc_header_row = l1_header_row
    unc_start_row = start_sample_row

    ws.merge_cells(f"H{unc_header_row}:I{unc_header_row}")
    ws[f"H{unc_header_row}"] = "Measurment uncertainty"
    ws[f"H{unc_header_row}"].font = font_bold
    ws[f"H{unc_header_row}"].fill = PatternFill(
        "solid", fgColor=COLOR_ORANGE_HEADER
    )
    ws[f"H{unc_header_row}"].alignment = align_center

    ws.merge_cells(f"H{unc_start_row}:I{unc_start_row}")
    ws[f"H{unc_start_row}"] = "Level 1"
    ws[f"H{unc_start_row}"].font = font_bold
    ws[f"H{unc_start_row}"].fill = PatternFill(
        "solid", fgColor=COLOR_BLUE_HEADER
    )
    ws[f"H{unc_start_row}"].alignment = align_center

    purity_row = unc_header_row - 1
    ws[f"H{purity_row}"] = "Standard Purity"
    ws[f"H{purity_row}"].font = font_bold
    purity_val = std_purity / 100.0 if std_purity > 1.0 else std_purity
    ws[f"I{purity_row}"] = purity_val

    uA_row = unc_start_row + 1
    uB_row = uA_row + 1
    uC_row = uB_row + 1
    uD_row = uC_row + 1
    uComb_row = uD_row + 1
    uExp_row = uComb_row + 1

    # تم تحديث صيغة uB لإضافة القيمة المطلقة ABS هنا:
    unc_labels = [
        (uA_row, "uA", f"=B{rsd_row}/100"),
        (uB_row, "uB", f"=ABS(0.5*(1 - (B{rec_row}/100)))/SQRT(3)"),
        (uC_row, "uC", f"=0.5*(1 - I{purity_row})/SQRT(3)"),
        (uD_row, "uD", f"=1 - SQRT(B{rsq_row})"),
        (
            uComb_row,
            "u combiend",
            f"=SQRT(I{uA_row}^2 + I{uB_row}^2 + I{uC_row}^2 + I{uD_row}^2)",
        ),
        (uExp_row, "U expanded", f"=2*I{uComb_row}"),
    ]

    for r_num, u_name, u_formula in unc_labels:
        ws[f"H{r_num}"] = u_name
        ws[f"I{r_num}"] = u_formula
        if u_name in ["u combiend", "U expanded"]:
            ws[f"H{r_num}"].font = font_bold
            ws[f"I{r_num}"].font = font_bold

    column_widths = {
        "A": 22,
        "B": 24,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 25,
        "H": 22,
        "I": 18,
        "J": 22,
        "K": 18,
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==========================================
# ⚙️ واجهة المستخدم (Streamlit UI)
# ==========================================
st.title("🧪 نظام التحقق من كفاءة الطرق التحليلية")

col_header1, col_header2 = st.columns(2)
with col_header1:
    test_name = st.text_input("اسم الاختبار / التحليل (مثلاً B1)", "B1")
with col_header2:
    conc_unit = st.text_input("وحدة التركيز", "ppm")

st.divider()

# 1. جدول المعايرة
st.subheader("📌 جدول المعايرة القياسي (Calibration STD)")

num_calib_levels = st.number_input(
    "عدد مستويات المعايرة",
    min_value=1,
    max_value=30,
    value=5,
    step=1,
    key="calib_num_input",
)

calib_data = [
    {"Level": f"STD {i+1}", "Concentration": 0.0, "Area": 0.0}
    for i in range(int(num_calib_levels))
]

valid_std_raw = st.data_editor(
    pd.DataFrame(calib_data),
    num_rows="dynamic",
    key=f"calib_table_editor_{num_calib_levels}",
    use_container_width=True,
)

valid_std = valid_std_raw.dropna(how="all").copy()
if "Level" in valid_std.columns:
    valid_std["Level"] = valid_std["Level"].fillna("")
if "Concentration" in valid_std.columns:
    valid_std["Concentration"] = (
        pd.to_numeric(valid_std["Concentration"], errors="coerce").fillna(0.0)
    )
if "Area" in valid_std.columns:
    valid_std["Area"] = (
        pd.to_numeric(valid_std["Area"], errors="coerce").fillna(0.0)
    )

st.divider()

# 2. جدول العينات
st.subheader("📋 جدول العينات والمدخلات (Level 1)")

col_input1, col_input2, col_input3 = st.columns(3)
with col_input1:
    target_conc = st.number_input(
        "مستوى الإضافة (Spiked Level)", value=10.0, min_value=0.0
    )
with col_input2:
    t_val = st.number_input(
        "قيمة t-statistic",
        value=2.571,
        help="مثال: القيمة 2.571 تتوافق مع n=6 و 95% confidence level",
    )
with col_input3:
    std_purity = st.number_input(
        "نقاوة المحلول القياسي (Standard Purity)",
        value=0.99,
        min_value=0.0,
        max_value=100.0,
    )

num_samples = st.number_input(
    "عدد التكراريات / العينات",
    min_value=1,
    max_value=30,
    value=6,
    step=1,
    key="samples_num_input",
)

sample_data = [
    {"Sample Name": f"Sample {i+1}", "Concentration": 0.0}
    for i in range(int(num_samples))
]

edited_samples_raw = st.data_editor(
    pd.DataFrame(sample_data),
    num_rows="dynamic",
    key=f"samples_table_editor_{num_samples}",
    use_container_width=True,
)

edited_samples = edited_samples_raw.dropna(how="all").copy()
if "Sample Name" in edited_samples.columns:
    edited_samples["Sample Name"] = edited_samples["Sample Name"].fillna("")
if "Concentration" in edited_samples.columns:
    edited_samples["Concentration"] = (
        pd.to_numeric(edited_samples["Concentration"], errors="coerce")
        .fillna(0.0)
    )

# تصدير الملف
st.divider()

try:
    calib_export = (
        valid_std[["Level", "Concentration", "Area"]]
        if not valid_std.empty
        else pd.DataFrame(columns=["Level", "Concentration", "Area"])
    )

    excel_file = generate_validation_excel(
        calib_df=calib_export,
        level1_df=edited_samples,
        test_title=test_name,
        unit_str=conc_unit,
        target_conc=target_conc,
        t_val=t_val,
        std_purity=std_purity,
    )

    st.download_button(
        label="📥 تحميل تقرير Validation Excel المحدث",
        data=excel_file,
        file_name="Method_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except Exception as e:
    st.error(f"حدث خطأ أثناء إعداد ملف Excel: {e}")

st.caption("---")
st.caption(
    "Developed with ❤️ by **Abdulrahman Alamri** | All Rights Reserved © 2026"
)
