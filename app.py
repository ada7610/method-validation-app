import io
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import streamlit as st

# ==========================================
# 🎨 إعدادات الصفحة الرئيسية
# ==========================================
st.set_page_config(
    page_title="تطبيق التحقق من الطرق التحليلية",
    page_layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================
# 📊 دالة إنشاء ملف Excel المنسق مطابق للصورة
# ==========================================
def generate_validation_excel(
    calib_df, level1_df, lod_val, loq_val, test_title, unit_str, unc_dict
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
        ws[f"A{r}"] = row.get("Concentration", 0)
        ws[f"B{r}"] = row.get("Area", 0)

    # RSQ في السطر 14
    ws["A14"] = "RSQ"
    ws["A14"].font = font_bold
    max_cal_row = len(calib_df) + 5
    ws["B14"] = f"=RSQ(B6:B{max_cal_row}, A6:A{max_cal_row})"

    # 3. الرسم البياني (Scatter Chart)
    chart = ScatterChart()
    chart.title = test_title
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

    for idx, row in level1_df.iterrows():
        r = idx + 19
        ws[f"A{r}"] = row.get("Sample", "")
        ws[f"B{r}"] = row.get("Concentration", 0)
        ws[f"C{r}"] = row.get("Recovery", 0)
        ws[f"D{r}"] = row.get("Outlier", 0)

    # الملاحظة الرمادية الجانبية لـ Outlier
    ws.merge_cells("E19:E24")
    ws["E19"] = (
        "Any value higher than the critical value in the table is consider outlier"
    )
    ws["E19"].font = Font(name="Calibri", size=9, bold=True)
    ws["E19"].fill = PatternFill("solid", fgColor=COLOR_GRAY_NOTE)
    ws["E19"].alignment = align_center

    # 5. ملخص الإحصائيات (Mean, SD, RSD, LOD, LOQ)
    stats_labels = [
        "Mean",
        "Recovery %",
        "Standerd Deviation",
        "RSD %",
        "LOD",
        "LOQ",
    ]
    for i, label in enumerate(stats_labels, start=26):
        ws[f"A{i}"] = label
        ws[f"A{i}"].font = font_bold
        ws[f"A{i}"].fill = PatternFill("solid", fgColor=COLOR_BLUE_HEADER)

    # معادلات الإكسل
    ws["B26"] = "=AVERAGE(B19:B24)"
    ws["B27"] = "=AVERAGE(C19:C24)"
    ws["B28"] = "=STDEV.S(B19:B24)"
    ws["B29"] = "=(B28/B26)*100"
    ws["B30"] = lod_val
    ws["B31"] = loq_val

    # 6. جدول Measurement Uncertainty
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
        ("uA", unc_dict.get("uA", 0)),
        ("uB", unc_dict.get("uB", 0)),
        ("uC", unc_dict.get("uC", 0)),
        ("uD", unc_dict.get("uD", 0)),
        ("u combiend", unc_dict.get("u_comb", 0)),
        ("U expanded", unc_dict.get("U_exp", 0)),
    ]

    for idx, (u_name, u_val) in enumerate(unc_labels, start=20):
        ws[f"G{idx}"] = u_name
        ws[f"H{idx}"] = u_val
        if u_name in ["u combiend", "U expanded"]:
            ws[f"G{idx}"].font = font_bold
            ws[f"H{idx}"].font = font_bold

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# ⚙️ واجهة المستخدم وإدخال البيانات
# ==========================================
st.title("🧪 نظام التحقق من كفاءة الطرق التحليلية (Method Validation)")

with st.sidebar:
    st.header("⚙️ إعدادات الاختبار")
    test_name = st.text_input("اسم الاختبار / التحليل", "Aflatoxins B1")
    conc_unit = st.text_input("وحدة التركيز", "ng/mL")

    st.subheader("📌 جدول المعايرة القياسي (Calibration STD)")
    num_stds = st.number_input("عدد المحاليل القياسية", 3, 10, 6)

    default_conc = [1.00, 2.50, 5.00, 10.00, 15.00, 20.00]
    default_area = [5.35, 11.84, 21.17, 38.70, 62.22, 84.82]

    calib_data = []
    for i in range(int(num_stds)):
        c_val = default_conc[i] if i < len(default_conc) else 1.0 * (i + 1)
        a_val = default_area[i] if i < len(default_area) else 10.0 * (i + 1)
        col1, col2 = st.columns(2)
        with col1:
            c = st.number_input(
                f"Conc {i+1}", value=float(c_val), key=f"c_{i}"
            )
        with col2:
            a = st.number_input(
                f"Area {i+1}", value=float(a_val), key=f"a_{i}"
            )
        calib_data.append({"Concentration": c, "Area": a})

    valid_std = pd.DataFrame(calib_data)

# حساب المعايرة
if len(valid_std) >= 2:
    slope, intercept = np.polyfit(
        valid_std["Concentration"], valid_std["Area"], 1
    )
    r_matrix = np.corrcoef(valid_std["Concentration"], valid_std["Area"])
    r_squared = r_matrix[0, 1] ** 2
else:
    slope, intercept, r_squared = 1, 0, 0

st.header("📋 عينات المستوى الأول (Level 1)")
target_conc = st.number_input("التركيز المستهدف (Spiked Conc)", value=4.00)

default_samples = [
    {"Sample Name": "LO level A-1", "Area": 17.58},
    {"Sample Name": "LO level A-2", "Area": 17.03},
    {"Sample Name": "LO level A-3", "Area": 16.07},
    {"Sample Name": "LO level A-4", "Area": 16.07},
    {"Sample Name": "LO level A-5", "Area": 17.41},
    {"Sample Name": "LO level A-6", "Area": 16.86},
]

edited_samples = st.data_editor(pd.DataFrame(default_samples), num_rows="dynamic")

# الحسابات للعينات
display_samples = []
calculated_concs = []

for _, row in edited_samples.iterrows():
    area_val = row["Area"]
    calc_c = (area_val - intercept) / slope if slope != 0 else 0
    rec = (calc_c / target_conc * 100) if target_conc != 0 else 0
    calculated_concs.append(calc_c)
    display_samples.append({
        "Sample Name": row["Sample Name"],
        "Area": area_val,
        f"Calculated Conc ({conc_unit})": round(calc_c, 2),
        "Recovery %": f"{rec:.2f}%",
    })

display_samples_df = pd.DataFrame(display_samples)

# Outliers حساب الـ
if len(calculated_concs) > 1:
    mean_c = np.mean(calculated_concs)
    std_c = np.std(calculated_concs, ddof=1)
    z_scores = [
        abs(c - mean_c) / std_c if std_c != 0 else 0 for c in calculated_concs
    ]
else:
    mean_c, std_c = 0, 0
    z_scores = [0] * len(calculated_concs)

display_samples_df["Z-Score"] = [round(z, 2) for z in z_scores]

st.subheader("📊 نتائج العينات والحيود")
st.dataframe(display_samples_df, use_container_width=True)

# الإحصائيات الحسابية
lod_result = round(3.3 * (std_c / slope), 2) if slope != 0 else 0
loq_result = round(10 * (std_c / slope), 2) if slope != 0 else 0

st.subheader("📐 حسابات عدم اليقين (Measurement Uncertainty)")
u_A = round(std_c / np.sqrt(len(calculated_concs)) if len(calculated_concs) > 0 else 0, 4)
u_B = 0.0017
u_C = 0.0058
u_D = 0.0015
u_combined = round(np.sqrt(u_A**2 + u_B**2 + u_C**2 + u_D**2), 4)
u_expanded = round(u_combined * 2, 4)

col_u1, col_u2 = st.columns(2)
with col_u1:
    st.write(f"**uA:** {u_A}")
    st.write(f"**uB:** {u_B}")
    st.write(f"**uC:** {u_C}")
with col_u2:
    st.write(f"**uD:** {u_D}")
    st.write(f"**u combined:** {u_combined}")
    st.write(f"**U expanded (k=2):** {u_expanded}")

# ==========================================
# 📥 قسم تصدير التقرير النهائي إلى Excel
# ==========================================
st.divider()
st.subheader("📥 تصدير التقرير النهائي")

try:
    calib_export = (
        valid_std[["Concentration", "Area"]]
        if not valid_std.empty
        else pd.DataFrame(columns=["Concentration", "Area"])
    )

    if not display_samples_df.empty:
        level1_export = pd.DataFrame({
            "Sample": display_samples_df["Sample Name"],
            "Concentration": display_samples_df[
                f"Calculated Conc ({conc_unit})"
            ],
            "Recovery": [
                float(str(r).replace("%", ""))
                for r in display_samples_df["Recovery %"]
            ],
            "Outlier": display_samples_df["Z-Score"],
        })
    else:
        level1_export = pd.DataFrame(
            columns=["Sample", "Concentration", "Recovery", "Outlier"]
        )

    unc_data = {
        "uA": u_A,
        "uB": u_B,
        "uC": u_C,
        "uD": u_D,
        "u_comb": u_combined,
        "U_exp": u_expanded,
    }

    excel_file = generate_validation_excel(
        calib_df=calib_export,
        level1_df=level1_export,
        lod_val=lod_result,
        loq_val=loq_result,
        test_title=test_name if test_name else "Aflatoxins B1",
        unit_str=conc_unit if conc_unit else "ng/mL",
        unc_dict=unc_data,
    )

    st.download_button(
        label="📥 تحميل تقرير Validation Excel المنسق",
        data=excel_file,
        file_name="Method_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except Exception as e:
    st.error(f"حدث خطأ أثناء إعداد ملف Excel: {e}")
