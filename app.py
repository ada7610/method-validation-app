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
    calib_df, level1_df, test_title, unit_str, target_conc, t_val
):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validation Report"
    ws.views.sheetView[0].showGridLines = True

    # الألوان والتنسيقات المطابقة للصورة
    COLOR_GREEN_HEADER = "C6EFCE"  # أخضر فاتح للعنوان الرئيسي
    COLOR_BLUE_HEADER = "8EA9DB"  # أزرق العناوين الفرعية
    COLOR_ORANGE_HEADER = "F4B084"  # برتقالي لهيدر الأعمدة
    COLOR_GRAY_NOTE = "D9D9D9"  # رمادي للملاحظات المدمجة

    font_main_title = Font(name="Calibri", size=14, bold=True)
    font_bold = Font(name="Calibri", size=11, bold=True)
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    # 1. العنوان الرئيسي العلوي
    ws.merge_cells("A1:K1")
    ws["A1"] = f"Calculation of Validation for {test_title}"
    ws["A1"].font = font_main_title
    ws["A1"].fill = PatternFill("solid", fgColor=COLOR_GREEN_HEADER)
    ws["A1"].alignment = align_center

    # 2. جدول Calibration STD
    ws.merge_cells("A4:B4")
    ws["A4"] = "Calibration STD"
    ws["A4"].font = font_bold
    ws["A4"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws["A4"].alignment = align_center

    ws["A5"] = f"Concentration ({unit_str})"
    ws["B5"] = "Area"
    for col in ["A", "B"]:
        ws[f"{col}5"].font = font_bold
        ws[f"{col}5"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"{col}5"].alignment = align_center

    for idx, row in calib_df.iterrows():
        r = idx + 6
        ws[f"A{r}"] = float(row.get("Concentration", 0))
        ws[f"B{r}"] = float(row.get("Area", 0))

    max_cal_row = max(len(calib_df) + 5, 6)

    # RSQ معادلة
    ws["A14"] = "RSQ"
    ws["A14"].font = font_bold
    ws["B14"] = f"=RSQ(B6:B{max_cal_row}, A6:A{max_cal_row})"

    # 3. الرسم البياني (Scatter Chart)
    chart = ScatterChart()
    chart.title = str(test_title)
    chart.style = 13

    xvalues = Reference(ws, min_col=1, min_row=6, max_row=max_cal_row)
    yvalues = Reference(ws, min_col=2, min_row=6, max_row=max_cal_row)
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker.symbol = "circle"
    series.trendline = openpyxl.chart.trendline.Trendline(
        trendlineType="linear", dispEq=True, dispRSqr=True
    )
    chart.series.append(series)
    chart.width = 12
    chart.height = 7
    ws.add_chart(chart, "E3")

    # Spiked Level و t-value قبل جدول العينات
    ws["A15"] = "t-value (t-test)"
    ws["A15"].font = font_bold
    ws["B15"] = float(t_val)

    ws["A16"] = "Spiked Level"
    ws["A16"].font = font_bold
    ws["B16"] = float(target_conc)

    # 4. جدول Level 1 (العينات)
    ws.merge_cells("A17:D17")
    ws["A17"] = f"Level 1 ({unit_str})"
    ws["A17"].font = font_bold
    ws["A17"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws["A17"].alignment = align_center

    headers_l1 = [
        "Samples name",
        f"Concentration ({unit_str})",
        "Recovery %",
        "Outlier",
    ]
    cols_l1 = ["A", "B", "C", "D"]
    for c, h in zip(cols_l1, headers_l1):
        ws[f"{c}18"] = h
        ws[f"{c}18"].font = font_bold
        ws[f"{c}18"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
        ws[f"{c}18"].alignment = align_center

    start_sample_row = 19
    for idx, row in level1_df.iterrows():
        r = idx + start_sample_row
        ws[f"A{r}"] = str(row.get("Sample Name", ""))

        # التركيز المُدخل مباشرة
        sample_conc = float(row.get("Concentration", 0))
        ws[f"B{r}"] = sample_conc

        # معادلة Recovery = (Concentration / Spiked Level) * 100
        ws[f"C{r}"] = f"=(B{r}/B16)*100"

        # معادلة Outlier (Z-Score)
        ws[f"D{r}"] = (
            f"=IF(STDEV.S(B$19:B$24)=0, 0, ABS(B{r}-AVERAGE(B$19:B$24))/STDEV.S(B$19:B$24))"
        )

    end_sample_row = max(len(level1_df) + start_sample_row - 1, 19)

    # الملاحظة الرمادية الجانبية لـ Outlier
    ws.merge_cells("E19:E24")
    ws["E19"] = (
        "Any value higher than the critical value in the table is consider outlier"
    )
    ws["E19"].font = Font(name="Calibri", size=9, bold=True)
    ws["E19"].fill = PatternFill("solid", fgColor=COLOR_GRAY_NOTE)
    ws["E19"].alignment = align_center

    # 5. ملخص الإحصائيات مع المعادلات
    stats_labels = [
        ("Mean", f"=AVERAGE(B{start_sample_row}:B{end_sample_row})"),
        ("Recovery %", f"=AVERAGE(C{start_sample_row}:C{end_sample_row})"),
        (
            "Standerd Deviation",
            f"=STDEV.S(B{start_sample_row}:B{end_sample_row})",
        ),
        ("RSD %", f"=(B28/B26)*100"),
        ("LOD", f"=B15*B28"),  # t-value * Standard Deviation
        ("LOQ", f"=10*B28"),  # 10 * Standard Deviation
    ]

    for i, (label, formula) in enumerate(stats_labels, start=26):
        ws[f"A{i}"] = label
        ws[f"A{i}"].font = font_bold
        ws[f"A{i}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
        ws[f"B{i}"] = formula

    # 6. جدول Measurement Uncertainty بمعادلات تفاعلية
    ws.merge_cells("G18:H18")
    ws["G18"] = "Measurment uncertainty"
    ws["G18"].font = font_bold
    ws["G18"].fill = PatternFill("solid", fgColor=COLOR_ORANGE_HEADER)
    ws["G18"].alignment = align_center

    ws.merge_cells("G19:H19")
    ws["G19"] = "Level 1"
    ws["G19"].font = font_bold
    ws["G19"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)
    ws["G19"].alignment = align_center

    unc_labels = [
        (
            "uA",
            f"=B28/SQRT(COUNT(B{start_sample_row}:B{end_sample_row}))",
        ),  # SD / sqrt(n)
        ("uB", 0.0017),
        ("uC", 0.0058),
        ("uD", 0.0015),
        ("u combiend", "=SQRT(H20^2 + H21^2 + H22^2 + H23^2)"),
        ("U expanded", "=H24*2"),
    ]

    for idx, (u_name, u_formula) in enumerate(unc_labels, start=20):
        ws[f"G{idx}"] = u_name
        ws[f"H{idx}"] = u_formula
        if u_name in ["u combiend", "U expanded"]:
            ws[f"G{idx}"].font = font_bold
            ws[f"H{idx}"].font = font_bold

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# ⚙️ واجهة المستخدم المبسطة
# ==========================================
st.title("🧪 نظام التحقق من كفاءة الطرق التحليلية")

col_header1, col_header2 = st.columns(2)
with col_header1:
    test_name = st.text_input("اسم الاختبار / التحليل", "Aflatoxins B1")
with col_header2:
    conc_unit = st.text_input("وحدة التركيز", "ng/mL")

st.divider()

# ------------------------------------------
# 1. جدول المعايرة القياسي (Calibration STD)
# ------------------------------------------
st.subheader("📌 جدول المعايرة القياسي (Calibration STD)")

default_calib = [
    {"Concentration": 1.00, "Area": 5.35},
    {"Concentration": 2.50, "Area": 11.84},
    {"Concentration": 5.00, "Area": 21.17},
    {"Concentration": 10.00, "Area": 38.70},
    {"Concentration": 15.00, "Area": 62.22},
    {"Concentration": 20.00, "Area": 84.82},
]

valid_std = st.data_editor(
    pd.DataFrame(default_calib),
    num_rows="dynamic",
    key="calib_table",
    use_container_width=True,
)

st.divider()

# ------------------------------------------
# 2. جدول العينات (Level 1)
# ------------------------------------------
st.subheader("📋 جدول العينات (Level 1)")

col_input1, col_input2 = st.columns(2)
with col_input1:
    target_conc = st.number_input(
        "مستوى الإضافة (Spiked Level / Conc)", value=4.00, min_value=0.01
    )
with col_input2:
    t_val = st.number_input(
        "قيمة t-statistic (لحسبة LOD)",
        value=2.571,
        help="القيمة الافتراضية 2.571 تتوافق مع n=6 و 95% confidence level",
    )

default_samples = [
    {"Sample Name": "LO level A-1", "Concentration": 3.85},
    {"Sample Name": "LO level A-2", "Concentration": 3.72},
    {"Sample Name": "LO level A-3", "Concentration": 3.51},
    {"Sample Name": "LO level A-4", "Concentration": 3.51},
    {"Sample Name": "LO level A-5", "Concentration": 3.81},
    {"Sample Name": "LO level A-6", "Concentration": 3.68},
]

edited_samples = st.data_editor(
    pd.DataFrame(default_samples),
    num_rows="dynamic",
    key="samples_table",
    use_container_width=True,
)

# ==========================================
# 📥 قسم تصدير التقرير النهائي إلى Excel
# ==========================================
st.divider()

try:
    calib_export = (
        valid_std[["Concentration", "Area"]]
        if not valid_std.empty
        else pd.DataFrame(columns=["Concentration", "Area"])
    )

    excel_file = generate_validation_excel(
        calib_df=calib_export,
        level1_df=edited_samples,
        test_title=test_name if test_name else "Aflatoxins B1",
        unit_str=conc_unit if conc_unit else "ng/mL",
        target_conc=target_conc,
        t_val=t_val,
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
