# ------------------------------------------
# 1. جدول المعايرة القياسي (Calibration STD) - مفتوح العدد
# ------------------------------------------
st.subheader("📌 جدول المعايرة القياسي (Calibration STD)")
st.caption(
    "💡 يمكنك إضافة أسطر جديدة أو حذفها مباشرة من الأسفل بضغط زر (+)"
)

# بيانات أولية افتراضية (تقدر تضيف أو تحذف منها بحرية)
default_calib = pd.DataFrame(
    [
        {"Level": "STD 1", "Concentration": 0.0, "Area": 0.0},
        {"Level": "STD 2", "Concentration": 0.0, "Area": 0.0},
        {"Level": "STD 3", "Concentration": 0.0, "Area": 0.0},
        {"Level": "STD 4", "Concentration": 0.0, "Area": 0.0},
        {"Level": "STD 5", "Concentration": 0.0, "Area": 0.0},
        {"Level": "STD 6", "Concentration": 0.0, "Area": 0.0},
    ]
)

valid_std = st.data_editor(
    default_calib,
    num_rows="dynamic",  # 👈 يجعل إضافة وحذف الصفوف ديناميكي ومفتوح
    key="calib_table_dynamic",
    use_container_width=True,
)

st.divider()

# ------------------------------------------
# 2. جدول العينات (Level 1) - مفتوح العدد
# ------------------------------------------
st.subheader("📋 جدول العينات والمدخلات (Level 1)")
st.caption(
    "💡 يمكنك إضافة أو حذف أي عدد من العينات مباشرة من أسفل الجدول"
)

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

# عينات افتراضية مبدئية
default_samples = pd.DataFrame(
    [{"Sample Name": f"Sample {i+1}", "Concentration": 0.0} for i in range(6)]
)

edited_samples = st.data_editor(
    default_samples,
    num_rows="dynamic",  # 👈 يجعل جدول العينات مفتوح وديناميكي بالكامل
    key="samples_table_dynamic",
    use_container_width=True,
)
