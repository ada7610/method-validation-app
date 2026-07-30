import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats
import matplotlib.pyplot as plt

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

# قيم افتراضية منعاً لأي Script Execution Error
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

# 4. قسم جدول العينات والـ Recovery %
st.subheader("2. Samples & Recovery Section (جدول العينات والنتائج)")

col_sec1, col_sec2 = st.columns(2)
with col_sec1:
    spike_level_name = st.text_input("Spike Level (المستوى):", value="Level 100%")
with col_sec2:
    spike_nominal_conc = st.number_input(f"Spike Nominal Concentration ({conc_unit}):", value=10.0, step=0.1, format="%.4f")

st.markdown("**جدول نتائج العينات الفردية:**")

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
            if sd_temp > 0:
                z_scores = np.abs((calc_concs - mean_temp) / sd_temp)
                for z in z_scores:
                    outlier_status.append("Outlier" if z > 2.5 else "Normal")
            else:
                outlier_status = ["Normal"] * len(calc_concs)
                
            display_samples_df = pd.DataFrame({
                "Sample Name": samples_df["Sample Name"][:len(calc_concs)],
                f"Calculated Conc ({conc_unit})": calc_concs,
                "Recovery %": [f"{r:.2f}%" for r in recoveries],
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
                "Value": [f"{mean_conc:.4f}", f"{mean_recovery:.2f}%", f"{sd_conc:.4f}", f"{rsd_percent:.2f}%", f"{lod:.4f}", f"{loq:.4f}"],
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

# 7. تصدير التقرير
st.subheader("5. Export Report (تصدير التقرير)")

export_df = pd.DataFrame({
    "Metric": ["Mean Conc", "Mean Recovery", "RSD%", "u_A", "u_B", "u_C", "u_D", "u_Combined", "u_Expanded (U)"],
    "Value": [f"{mean_conc:.4f}", f"{mean_recovery:.2f}%", f"{rsd_percent:.2f}%", f"{u_A:.6f}", f"{u_B:.6f}", f"{u_C:.6f}", f"{u_D:.6f}", f"{u_combined:.6f}", f"{u_expanded:.6f}"],
    "Unit": [conc_unit, "%", "%", "-", "-", "-", "-", "-", conc_unit]
})

csv_data = export_df.to_csv(index=False).encode('utf-8-sig')

st.download_button(
    label="📥 تحميل التقرير (يفتح مباشرة في Excel)",
    data=csv_data,
    file_name=f"Validation_Report_{test_name if test_name else 'Method'}.csv",
    mime="text/csv"
)
