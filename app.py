import io
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import pandas as pd
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Font, PatternFill
from scipy import stats
import streamlit as st
def generate_validation_excel(calib_df, level1_df, lod_val, loq_val, t_val):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Flouride"
    ws.views.sheetView[0].showGridLines = True

    # الألوان والتنسيقات
    COLOR_HEADER_CALIB = "D9E1F2"
    COLOR_HEADER_LEVEL = "8EA9DB"
    COLOR_HEADER_ORANGE = "F4B084"
    COLOR_BLUE_SUMMARY = "D9E1F2"
    COLOR_RED_LOD = "FF0000"

    font_bold = Font(name="Calibri", size=11, bold=True)
    font_white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center")

    # 1. جدول المعايرة
    ws.merge_cells("A1:B1")
    ws["A1"] = "Calibration STD"
    ws["A1"].font = font_bold
    ws["A1"].fill = PatternFill("solid", fgColor=COLOR_HEADER_CALIB)
    ws["A1"].alignment = align_center

    ws["A2"] = "Concentration (mg/L)"
    ws["B2"] = "Area (µs*min)"
    for col in ["A", "B"]:
        ws[f"{col}2"].font = font_bold
        ws[f"{col}2"].fill = PatternFill("solid", fgColor=COLOR_HEADER_CALIB)
        ws[f"{col}2"].alignment = align_center

    for idx, row in calib_df.iterrows():
        r = idx + 3
        ws[f"A{r}"] = row["Concentration"]
        ws[f"B{r}"] = row["Area"]

    # 2. رسم المنحنى البياني
    chart = ScatterChart()
    chart.title = "Flouride Calibration Curve"
    chart.x_axis.title = "Concentration (mg/L)"
    chart.y_axis.title = "Area (µs*min)"

    max_row_calib = len(calib_df) + 2
    xvalues = Reference(ws, min_col=1, min_row=3, max_row=max_row_calib)
    yvalues = Reference(ws, min_col=2, min_row=3, max_row=max_row_calib)
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker.symbol = "circle"
    series.trendline = openpyxl.chart.trendline.Trendline(
        trendlineType="linear", dispEq=True, dispRSqr=True
    )
    chart.series.append(series)
    chart.width = 14
    chart.height = 7
    ws.add_chart(chart, "D1")

    # 3. جدول Level 1
    ws.merge_cells("A18:I18")
    ws["A18"] = "Level1 (0.6 mg/L)"
    ws["A18"].font = font_bold
    ws["A18"].fill = PatternFill("solid", fgColor=COLOR_HEADER_LEVEL)
    ws["A18"].alignment = align_center

    headers = [
        "Samples name",
        "Concentration (mg/L)",
        "Area(µs*min)",
        "Recovery %",
        f"Outlier ({t_val})",
        "",
        "",
        "Measurment uncertainty",
        "",
    ]
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
    for c, h in zip(cols, headers):
        ws[f"{c}19"] = h
        ws[f"{c}19"].font = font_bold
        ws[f"{c}19"].fill = PatternFill("solid", fgColor=COLOR_HEADER_ORANGE)
        ws[f"{c}19"].alignment = align_center

    for idx, row in level1_df.iterrows():
        r = idx + 20
        ws[f"A{r}"] = row["Sample"]
        ws[f"B{r}"] = row["Concentration"]
        ws[f"C{r}"] = row["Area"]
        ws[f"D{r}"] = row["Recovery"]
        ws[f"E{r}"] = row["Outlier"]
        ws[f"F{r}"] = row["U_Type"]
        ws[f"H{r}"] = row["Uncertainty"]

    # 4. ملخص الحسابات
    ws["A25"] = "Mean"
    ws["B25"] = "=AVERAGE(B20:B24)"
    ws["A26"] = "Mean Recovery %"
    ws["B26"] = "=AVERAGE(D20:D24)"
    ws["A27"] = "Standard Deviation"
    ws["B27"] = "=STDEV.S(B20:B24)"
    ws["A28"] = "RSD %"
    ws["B28"] = "=(B27/B25)*100"

    ws["A29"] = "LOD"
    ws["B29"] = lod_val
    ws["A30"] = "LOQ"
    ws["B30"] = "=3*B29"

    for r in [29, 30]:
        ws[f"A{r}"].fill = PatternFill("solid", fgColor=COLOR_RED_LOD)
        ws[f"B{r}"].fill = PatternFill("solid", fgColor=COLOR_RED_LOD)
        ws[f"A{r}"].font = font_white_bold
        ws[f"B{r}"].font = font_white_bold

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# 1. إعدادات الصفحة الأساسية
st.set_page_config(page_title="Method Validation Calculator", page_icon="🔬", layout="wide")

# 2. الهيدر والعنوان المتغير
st.title("🔬 Method Validation Calculator")

col_header1, col_header2 = st.columns([2, 1])

with col_header1:
    test_name = st.text_input("اسم الاختبار (Test Name / Parameter):", value="", placeholder="أدخل اسم الاختبار هنا")

with col_header2:
    conc_unit = st.text_input("وحدة التركيز (Concentration Unit):", value="ppm", placeholder="مثلاً: ppm, mg/L, µg/mL")

if test_name.strip():
    st.header(f"Validation for {test_name}")
else:
    st.header("Validation for ....................")

st.markdown("---")

# قيم افتراضية منعاً لأي أخطاء
slope, intercept, r_squared, steyx = 0.0, 0.0, 0.0, 0.0
mean_conc, sd_conc, mean_recovery, rsd_percent = 0.0, 0.0, 0.0, 0.0
lod, loq = 0.0, 0.0

# 3. قسم منحنى المعايرة والخطية
st.subheader("1. Calibration Curve & Linearity (منحنى المعايرة)")

col_cal1, col_cal2 = st.columns([1, 1.2])

with col_cal1:
    st.markdown("**جدول نقاط الكالبريشن (قابل للزيادة والتعديل):**")
    default_std_data = pd.DataFrame({
        "Standard Level": [f"Std {i+1}" for i in range(6)],
        "Concentration": [1.0, 2.0, 5.0, 10.0, 15.0, 20.0],
        "Area": [1200.0, 2450.0, 6100.0, 12300.0, 18500.0, 24800.0]
    })
    
    std_df = st.data_editor(default_std_data, num_rows="dynamic", key="std_editor", use_container_width=True)

if std_df is not None:
    valid_std = std_df.dropna(subset=["Concentration", "Area"])
    if len(valid_std) >= 2:
        try:
            x_cal = pd.to_numeric(valid_std["Concentration"], errors='coerce').values
            y_cal = pd.to_numeric(valid_std["Area"], errors='coerce').values
            
            mask = ~np.isnan(x_cal) & ~np.isnan(y_cal)
            x_cal, y_cal = x_cal[mask], y_cal[mask]

            if len(x_cal) >= 2:
                slope, intercept, r_value, p_value, std_err = stats.linregress(x_cal, y_cal)
                r_squared = float(r_value**2)
                
                y_pred = slope * x_cal + intercept
                residuals = y_cal - y_pred
                steyx = float(np.sqrt(np.sum(residuals**2) / (len(x_cal) - 2))) if len(x_cal) > 2 else 0.0
                
                with col_cal1:
                    st.success(f"**RSQ ($R^2$):** `{r_squared:.5f}`")
                    
                with col_cal2:
                    fig, ax = plt.subplots(figsize=(6, 4))
                    ax.scatter(x_cal, y_cal, color="#1f77b4", label="Standards", s=50)
                    
                    x_line = np.linspace(min(x_cal), max(x_cal), 100)
                    y_line = slope * x_line + intercept
                    ax.plot(x_line, y_line, color="#d62728", linestyle="--", label="Linear Regression")
                    
                    sign = "+" if intercept >= 0 else "-"
                    eq_text = f"y = {slope:.2f}x {sign} {abs(intercept):.2f}\n$R^2$ = {r_squared:.5f}"
                    
                    ax.set_title("Calibration Curve", fontsize=12, fontweight="bold")
                    ax.set_xlabel(f"Concentration ({conc_unit})", fontsize=10)
                    ax.set_ylabel("Area", fontsize=10)
                    ax.grid(True, linestyle=":", alpha=0.6)
                    ax.legend()
                    
                    ax.text(0.05, 0.82, eq_text, transform=ax.transAxes, fontsize=10,
                            bbox=dict(boxstyle="round,pad=0.5", facecolor="white", alpha=0.8, edgecolor="gray"))
                    
                    st.pyplot(fig)
        except Exception as e:
            st.error(f"خطأ في رسم المعايرة: {e}")

st.markdown("---")

# 4. قسم جدول العينات والـ Recovery % وتقييم Outliers
st.subheader("2. Samples & Recovery Section (جدول العينات وفحص Outliers)")

col_sec1, col_sec2 = st.columns(2)
with col_sec1:
    spike_level_name = st.text_input("Spike Level (المستوى):", value="Level 100%")
with col_sec2:
    spike_nominal_conc = st.number_input(f"Spike Nominal Concentration ({conc_unit}):", value=10.0, step=0.1, format="%.4f")

st.markdown("**جدول نتائج العينات الفردية:** *(جرب إدخال قيمة مرتفعة جداً أو منخفضة جداً لتختبر الـ Outlier)*")

default_samples_data = pd.DataFrame({
    "Sample Name": [f"Sample {i+1}" for i in range(6)],
    f"Calculated Conc ({conc_unit})": [9.80, 9.90, 10.10, 9.70, 10.00, 10.20]
})

samples_df = st.data_editor(default_samples_data, num_rows="dynamic", key="sample_editor", use_container_width=True)

if samples_df is not None and not samples_df.empty and spike_nominal_conc > 0:
    try:
        conc_col_name = samples_df.columns[1]
        calc_concs = pd.to_numeric(samples_df[conc_col_name], errors='coerce').dropna().values
        
        if len(calc_concs) > 0:
            recoveries = (calc_concs / spike_nominal_conc) * 100.0
            
            mean_temp = float(np.mean(calc_concs))
            sd_temp = float(np.std(calc_concs, ddof=1)) if len(calc_concs) > 1 else 0.0
            
            outlier_status = []
            z_score_list = []
            
            # حساب Z-Score لاكتشاف Outliers (معيار 95% Confidence Level)
            if sd_temp > 0 and len(calc_concs) >= 3:
                z_scores = np.abs((calc_concs - mean_temp) / sd_temp)
                for z in z_scores:
                    z_score_list.append(f"{z:.2f}")
                    if z > 1.96:
                        outlier_status.append("⚠️ Outlier")
                    else:
                        outlier_status.append("✅ Normal")
            else:
                z_score_list = ["N/A"] * len(calc_concs)
                outlier_status = ["✅ Normal"] * len(calc_concs)
                
            display_samples_df = pd.DataFrame({
                "Sample Name": samples_df["Sample Name"][:len(calc_concs)],
                f"Calculated Conc ({conc_unit})": calc_concs,
                "Recovery %": [f"{r:.2f}%" for r in recoveries],
                "Z-Score": z_score_list,
                "Outlier Status": outlier_status
            })
            
            st.dataframe(display_samples_df, use_container_width=True)
            
            # 5. الجدول الملحق الإحصائي
            st.subheader("3. Summary Statistics Table (الجدول الإحصائي الملحق)")
            
            mean_conc = float(np.mean(calc_concs))
            mean_recovery = float(np.mean(recoveries))
            sd_conc = float(np.std(calc_concs, ddof=1)) if len(calc_concs) > 1 else 0.0
            rsd_percent = float((sd_conc / mean_conc * 100.0)) if mean_conc > 0 else 0.0
            
            if slope > 0:
                lod = float((3.3 * steyx) / slope)
                loq = float((10.0 * steyx) / slope)
                
            summary_data = {
                "Parameter": ["Mean Concentration", "Mean Recovery", "SD (Standard Deviation)", "%RSD (Relative Standard Deviation)", "LOD (Limit of Detection)", "LOQ (Limit of Quantitation)"],
                "Value": [f"{mean_conc:.4f}", f"{mean_recovery:.2f}%", f"{sd_conc:.4f}", f"{rsd_percent:.2f}%", st.subheader("حساب حد الكشف (LOD) وحد التقدير الكمي (LOQ)")

# 1. إدخال يدوي لقيمة T-test
t_value = st.number_input(
    "أدخل قيمة T-test:", value=3.143, step=0.001, format="%.3f"
)

# 2. المعادلة الجديدة: LOD = t * SD
lod_result = t_value * sd_value  # (تأكد أن sd_value هو اسم متغير الانحراف المعياري لديك)
loq_result = lod_result * 3

# 3. عرض النتائج الجديدة
st.write(f"**LOD:** {lod_result:.4f}")
st.write(f"**LOQ:** {loq_result:.4f}"), f"{loq:.4f}"],
                "Unit": [conc_unit, "%", conc_unit, "%", conc_unit, conc_unit]
            }
            st.table(pd.DataFrame(summary_data))
    except Exception as e:
        st.error(f"خطأ في حساب نتائج العينات: {e}")

st.markdown("---")

# 6. جدول عدم اليقين (Measurement Uncertainty Budget)
st.subheader("4. Measurement Uncertainty Budget (جدول عدم اليقين)")

col_std1, col_std2 = st.columns(2)
with col_std1:
    purity_percent = st.number_input("Standard Reference Purity (%) (نقاوة الأستاندرد):", value=99.50, min_value=0.0, max_value=100.0, step=0.1, format="%.2f")

u_A = float(rsd_percent / 100.0)
u_B = float(abs((0.5 * (1.0 - (mean_recovery / 100.0))) / np.sqrt(3)))
u_C = float((0.5 * (1.0 - (purity_percent / 100.0))) / np.sqrt(3))
u_D = float(1.0 - np.sqrt(r_squared)) if r_squared >= 0 else 0.0

u_combined = float(np.sqrt(u_A**2 + u_B**2 + u_C**2 + u_D**2))
k_factor = 2
u_expanded = float(u_combined * k_factor)

unc_summary_data = {
    "Uncertainty Component": ["u_A", "u_B", "u_C", "u_D", "u_Combined (u_c)", "u_Expanded (U, k=2)"],
    "Formula / Calculation Method": [
        "RSD / 100", 
        "| 0.5 * (1 - Recovery/100) / √3 |", 
        "0.5 * (1 - STD Purity/100) / √3", 
        "1 - √(RSQ)", 
        "√(uA² + uB² + uC² + uD²)", 
        "u_Combined * 2 (95% Confidence Level)"
    ],
    "Calculated Value": [f"{u_A:.6f}", f"{u_B:.6f}", f"{u_C:.6f}", f"{u_D:.6f}", f"{u_combined:.6f}", f"{u_expanded:.6f}"]
}

st.table(pd.DataFrame(unc_summary_data))

final_result_text = f"{mean_conc:.4f} ± {u_expanded:.4f} {conc_unit}"
st.markdown("### 📌 Result for Report:")
st.success(f"**{test_name if test_name else 'Result'} = {final_result_text}** (at 95% Confidence Level, k=2)")

st.markdown("---")

# ==========================================
# 📥 قسم تصدير التقرير النهائي إلى Excel
# ==========================================
st.divider()  # خط فاصل لتنظيم الواجهة
st.subheader("📥 تصدير التقرير النهائي")

try:
    # 1. تجهيز ملف الإكسل المنسق في الذاكرة عبر استدعاء الدالة الجديدة
    excel_file = generate_validation_excel(
        calib_df=calib_df,  # جدول المعايرة
        level1_df=level1_df,  # جدول بيانات المستوى الأول
        lod_val=lod_result,  # قيمة LOD المعمدة بناءً على T-test
        loq_val=loq_result,  # قيمة LOQ
        t_val=t_value,  # قيمة t-score المدخلة
    )

    # 2. عرض زر التحميل للمستخدم
    st.download_button(
        label="📥 تحميل تقرير Validation Excel المنسق",
        data=excel_file,
        file_name="Method_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,  # ليكون الزر بعرض الصفحة وشكله مرتب
    )
except Exception as e:
    st.error(f"حدث خطأ أثناء إعداد ملف Excel: {e}")
)
