import io
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import streamlit as st


# ==========================================
# 📊 دالة إنشاء ملف Excel المنسق بدوال ومعادلات تفاعلية
# ==========================================
def generate_validation_excel(
    calib_df, level1_df, test_title, unit_str, target_conc, t_val, std_purity
):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validation Report"
    ws.views.sheetView[0].showGridLines = True

    # الألوان والتنسيقات
    COLOR_GREEN_HEADER = "C6EFCE"  # أخضر فاتح للعنوان الرئيسي
    COLOR_BLUE_HEADER = "8EA9DB"  # أزرق العناوين الفرعية
    COLOR_ORANGE_HEADER = "F4B084"  # برتقالي لهيدر الأعمدة
    COLOR_GRAY_NOTE = "D9D9D9"  # رمادي للملاحظات المدمجة

    font_main_title = Font(name="Calibri", size=14, bold=True)
    font_bold = Font(name="Calibri", size=11, bold=True)
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=False
    )
    align_left = Alignment(horizontal="left", vertical="center")

    # 1. العنوان الرئيسي العلوي
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

    # 2. جدول Calibration STD
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
        ws[f"{col}5"].alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=False
        )

    for idx, row in calib_df.iterrows():
        r = idx + 6
        ws[f"A{r}"] = str(row.get("Level", ""))
        ws[f"B{r}"] = (
            float(row.get("Concentration", 0))
            if pd.notnull(row.get("Concentration"))
            else 0.0
        )
        ws[f"C{r}"] = (
            float(row.get("Area", 0)) if pd.notnull(row.get("Area")) else 0.0
        )

    max_cal_row = max(len(calib_df) + 5, 6)

    # RSQ معادلة (في الخلية B14)
    ws["A14"] = "RSQ"
    ws["A14"].font = font_bold
    ws["B14"] = f"=RSQ(C6:C{max_cal_row}, B6:B{max_cal_row})"

    # 3. الرسم البياني (Scatter Chart)
    chart = ScatterChart()
    chart.title = str(test_title) if test_title else "Calibration Curve"
    chart.style = 13

    xvalues = Reference(ws, min_col=2, min_row=6, max_row=max_cal_row)
    yvalues = Reference(ws, min_col=3, min_row=6, max_row=max_cal_row)

    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker.symbol = "circle"
    series.graphicalProperties.line.noFill = True
    series.trendline = openpyxl.chart.trendline.Trendline(
        trendlineType="linear", dispEq=True, dispRSqr=True
    )

    chart.series.append(series)
    chart.width = 12
    chart.height = 7
    ws.add_chart(chart, "F3")

    # 4. خانات t-value و Spiked Level
    ws["A15"] = "t-value (t test)"
    ws["B15"] = float(t_val)

    ws["A16"] = "Spiked Level"
    ws["B16"] = float(target_conc)

    for r in [15, 16]:
        ws[f"A{r}"].font = font_bold
        ws[f"A{r}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"A{r}"].alignment = align_left

        ws[f"B{r}"].font = font_bold
        ws[f"B{r}"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"B{r}"].alignment = align_center

    # 5. جدول Level 1 (العينات)
    ws.merge_cells("A17:E17")
    ws["A17"] = f"Level 1 ({unit_str})" if unit_str else "Level 1"
    ws["A17"].font = font_bold
    ws["A17"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws["A17"].alignment = align_center

    headers_l1 = [
        "Samples name",
        unit_header,
        "Recovery %",
        "Outlier (Z-Score)",
        "Outlier Status",
    ]
    cols_l1 = ["A", "B", "C", "D", "E"]
    for c, h in zip(cols_l1, headers_l1):
        ws[f"{c}18"] = h
        ws[f"{c}18"].font = font_bold
        ws[f"{c}18"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"{c}18"].alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=False
        )

    start_sample_row = 19
    for idx, row in level1_df.iterrows():
        r = idx + start_sample_row
        ws[f"A{r}"] = str(row.get("Sample Name", ""))

        sample_conc = (
            float(row.get("Concentration", 0))
            if pd.notnull(row.get("Concentration"))
            else 0.0
        )
        ws[f"B{r}"] = sample_conc

        # Recovery = (Concentration / Spiked Level) * 100
        ws[f"C{r}"] = f"=IF(B16=0, 0, (B{r}/B16)*100)"

        # Outlier = ABS(Concentration - Mean) / SD
        ws[f"D{r}"] = f"=IF(B$28=0, 0, ABS(B{r}-B$26)/B$28)"

        # Outlier Status
        ws[f"E{r}"] = f'=IF(D{r}<=2.57, "Normal", "Outlier")'
        ws[f"E{r}"].alignment = align_center

    end_sample_row = max(len(level1_df) + start_sample_row - 1, 19)

    # الملاحظة الرمادية
    ws.merge_cells("F19:F24")
    ws["F19"] = (
        "Any value higher than the critical value in the table is consider outlier"
    )
    ws["F19"].font = Font(name="Calibri", size=9, bold=True)
    ws["F19"].fill = PatternFill("solid", fgColor=COLOR_GRAY_NOTE)
    ws["F19"].alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    # 6. ملخص الإحصائيات
    stats_labels = [
        ("Mean", f"=AVERAGE(B{start_sample_row}:B{end_sample_row})"),
        ("Recovery %", f"=AVERAGE(C{start_sample_row}:C{end_sample_row})"),
        (
            "Standerd Deviation",
            f"=STDEV.S(B{start_sample_row}:B{end_sample_row})",
        ),
        ("RSD %", f"=IF(B26=0, 0, (B28/B26)*100)"),
        ("LOD", f"=B15*B28"),
        ("LOQ", f"=10*B28"),
    ]

    for i, (label, formula) in enumerate(stats_labels, start=26):
        ws[f"A{i}"] = label
        ws[f"A{i}"].font = font_bold
        ws[f"A{i}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
        ws[f"B{i}"] = formula

    # 7. جدول Measurement Uncertainty
    ws.merge_cells("H18:I18")
    ws["H18"] = "Measurment uncertainty"
    ws["H18"].font = font_bold
    ws["H18"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
    ws["H18"].alignment = align_center

    ws.merge_cells("H19:I19")
    ws["H19"] = "Level 1"
    ws["H19"].font = font_bold
    ws["H19"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws["H19"].alignment = align_center

    ws["H17"] = "Standard Purity"
    ws["H17"].font = font_bold
    purity_val = std_purity / 100.0 if std_purity > 1.0 else std_purity
    ws["I17"] = purity_val

    unc_labels = [
        ("uA", "=B29/100"),
        ("uB", "=0.5*(1 - (B27/100))/SQRT(3)"),
        ("uC", "=0.5*(1 - I17)/SQRT(3)"),
        ("uD", "=1 - SQRT(B14)"),
        ("u combiend", "=SQRT(I20^2 + I21^2 + I22^2 + I23^2)"),
        ("U expanded", "=2*I24"),
    ]

    for idx, (u_name, u_formula) in enumerate(unc_labels, start=20):
        ws[f"H{idx}"] = u_name
        ws[f"I{idx}"] = u_formula
        if u_name in ["u combiend", "U expanded"]:
            ws[f"H{idx}"].font = font_bold
            ws[f"I{idx}"].font = font_bold

    column_widths = {
        "A": 22,
        "B": 24,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 25,
        "H": 22,
        "I": 18,
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
    test_name = st.text_input("اسم الاختبار / التحليل", "")
with col_header2:
    conc_unit = st.text_input("وحدة التركيز", "")

st.divider()

# ------------------------------------------
# 1. جدول المعايرة القياسي (Calibration STD)
# ------------------------------------------
st.subheader("📌 جدول المعايرة القياسي (Calibration STD)")

default_calib = [
    {"Level": "STD 1", "Concentration": 0.0, "Area": 0.0},
    {"Level": "STD 2", "Concentration": 0.0, "Area": 0.0},
    {"Level": "STD 3", "Concentration": 0.0, "Area": 0.0},
    {"Level": "STD 4", "Concentration": 0.0, "Area": 0.0},
    {"Level": "STD 5", "Concentration": 0.0, "Area": 0.0},
    {"Level": "STD 6", "Concentration": 0.0, "Area": 0.0},
]

valid_std = st.data_editor(
    pd.DataFrame(default_calib),
    num_rows="dynamic",
    key="calib_table",
    use_container_width=True,
)

# 🔄 تحديث وتوليد الترقيم التلقائي لخانة Level
if not valid_std.empty:
    valid_std["Level"] = [f"STD {i+1}" for i in range(len(valid_std))]

st.divider()

# ------------------------------------------
# 2. جدول العينات والمدخلات الإضافية
# ------------------------------------------
st.subheader("📋 جدول العينات والمدخلات (Level 1)")

col_input1, col_input2, col_input3 = st.columns(3)
with col_input1:
    target_conc = st.number_input(
        "مستوى الإضافة (Spiked Level)", value=0.0, min_value=0.0
    )
with col_input2:
    t_val = st.number_input(
        "قيمة t-statistic (لحسبة LOD)",
        value=0.0,
        help="مثال: القيمة 2.571 تتوافق مع n=6 و 95% confidence level",
    )
with col_input3:
    std_purity = st.number_input(
        "نقاوة المحلول القياسي (Standard Purity)",
        value=0.99,
        min_value=0.0,
        max_value=100.0,
        help="أدخل النسبة ككسر عشري (مثلاً 0.99) أو نسبة مئوية (مثلاً 99)",
    )

default_samples = [
    {"Sample Name": "Sample 1", "Concentration": 0.0},
    {"Sample Name": "Sample 2", "Concentration": 0.0},
    {"Sample Name": "Sample 3", "Concentration": 0.0},
    {"Sample Name": "Sample 4", "Concentration": 0.0},
    {"Sample Name": "Sample 5", "Concentration": 0.0},
    {"Sample Name": "Sample 6", "Concentration": 0.0},
]

edited_samples = st.data_editor(
    pd.DataFrame(default_samples),
    num_rows="dynamic",
    key="samples_table",
    use_container_width=True,
)

# 🔄 تحديث وتوليد الترقيم التلقائي لخانة Sample Name
if not edited_samples.empty:
    edited_samples["Sample Name"] = [
        f"Sample {i+1}" for i in range(len(edited_samples))
    ]

# ==========================================
# 📥 قسم تصدير التقرير النهائي إلى Excel
# ==========================================
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
        label="📥 تحميل تقرير Validation Excel المنسق (مع دوال تفاعلية)",
        data=excel_file,
        file_name="Method_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except Exception as e:
    st.error(f"حدث خطأ أثناء إعداد ملف Excel: {e}")
