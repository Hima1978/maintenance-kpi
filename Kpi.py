import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="شركة العلمين فليكس للطباعة - نظام KPIs", page_icon="⚙️", layout="wide")

EXCEL_FILE = "maintenance_kpi_database.xlsx"

def load_data():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    return pd.DataFrame(columns=[
        "التاريخ", "المصنع/القسم", "رقم الوردية", "رقم/اسم الماكينة", 
        "اسم مشغل الماكينة", "اسم القائم بالصيانة", "وقت البداية", "وقت النهاية", 
        "مدة العطل (ساعة)", "السبب الرئيسي", "عطل صيانة", 
        "تأخير مشتريات", "صيانة مخططة", "خطأ مشغل", 
        "الساعات المجانية", "توافرية الصيانة (%)", "ملاحظات"
    ])

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False)

# العنوان الرئيسي للتطبيق
st.title("🏭 شركة العلمين فليكس للطباعة")
st.subheader("⚙️ منظومة متابعة أعطال الصيانة والـ KPIs")

tab1, tab2 = st.tabs(["📝 إدخال إخطار عطل", "📊 قاعدة البيانات والتحليلات"])

with tab1:
    # العنوان الجانبي المكتوب بناءً على طلبك
    st.header("🚨 إخطار عطل")
    
    with st.form("kpi_form"):
        # البيانات الأساسية للوردية والموقع
        st.subheader("📌 البيانات الأساسية")
        col1, col2, col3 = st.columns(3)
        with col1:
            date_input = st.date_input("التاريخ", datetime.now())
            factory_site = st.selectbox("المصنع / القسم", [
                "مصنع الطباعة", 
                "مصنع السلندرات", 
                "محطة السولفنت", 
                "الخدمات (Chiller/Dryer/Compressors)"
            ])
        with col2:
            shift_num = st.selectbox("رقم الوردية", ["الوردية الأولى (1)", "الوردية الثانية (2)", "الوردية الثالثة (3)"])
            machine_id = st.text_input("رقم / اسم الماكينة", "Bobst / Comexi / Daetwyler")
        with col3:
            operator_name = st.text_input("اسم مشغل الماكينة", "")
            technician_name = st.text_input("اسم القائم بالصيانة", "")

        st.divider()

        # تفاصيل العطل والأوقات
        st.subheader("⏱️ تفاصيل العطل والوقت")
        col4, col5, col6 = st.columns(3)
        with col4:
            t1_input = st.time_input("وقت بداية العطل", datetime.strptime("08:00", "%H:%M").time())
            t2_input = st.time_input("وقت نهاية العطل", datetime.strptime("09:30", "%H:%M").time())
        with col5:
            cause_cat = st.selectbox("السبب الرئيسي للعطل", ["عطل صيانة", "تأخير مشتريات", "صيانة مخططة", "خطأ مشغل"])
            free_hrs = st.number_input("الساعات التشغيلية المتاحة للوردية", min_value=0.0, value=8.0, step=0.5)
        with col6:
            notes = st.text_area("ملاحظات / أرقام قطع الغيار المستخدمة", "")

        submit = st.form_submit_button("حفظ الإخطار في Excel 💾", use_container_width=True)

        if submit:
            try:
                # حساب مدة العطل تلقائياً بالساعات والدقائق
                t1_dt = datetime.combine(datetime.today(), t1_input)
                t2_dt = datetime.combine(datetime.today(), t2_input)
                
                if t2_dt < t1_dt:  # التعامل مع الأعطال الممتدة عبر منتصف الليل
                    t2_dt += timedelta(days=1)
                
                duration = round((t2_dt - t1_dt).total_seconds() / 3600.0, 2)

                # توزيع الساعات حسب السبب
                maint_h = duration if cause_cat == "عطل صيانة" else 0.0
                proc_h = duration if cause_cat == "تأخير مشتريات" else 0.0
                pm_h = duration if cause_cat == "صيانة مخططة" else 0.0
                op_h = duration if cause_cat == "خطأ مشغل" else 0.0

                # حساب التوافرية
                total_planned = free_hrs + duration
                eff_planned = total_planned - (proc_h + op_h)
                availability = (free_hrs / eff_planned * 100) if eff_planned > 0 else 100.0

                df_curr = load_data()
                new_row = {
                    "التاريخ": str(date_input),
                    "المصنع/القسم": factory_site,
                    "رقم الوردية": shift_num,
                    "رقم/اسم الماكينة": machine_id,
                    "اسم مشغل الماكينة": operator_name,
                    "اسم القائم بالصيانة": technician_name,
                    "وقت البداية": t1_input.strftime("%H:%M"),
                    "وقت النهاية": t2_input.strftime("%H:%M"),
                    "مدة العطل (ساعة)": duration,
                    "السبب الرئيسي": cause_cat,
                    "عطل صيانة": maint_h,
                    "تأخير مشتريات": proc_h,
                    "صيانة مخططة": pm_h,
                    "خطأ مشغل": op_h,
                    "الساعات المجانية": free_hrs,
                    "توافرية الصيانة (%)": round(availability, 2),
                    "ملاحظات": notes
                }
                save_data(pd.concat([df_curr, pd.DataFrame([new_row])], ignore_index=True))
                st.success(f"تم الحفظ بنجاح! مدة العطل: {duration} ساعة | نسبة التوافرية للوردية: {availability:.2f}%")
            except Exception as e:
                st.error(f"حدث خطأ أثناء إدخال البيانات: {e}")

with tab2:
    df_data = load_data()
    st.dataframe(df_data, use_container_width=True)
    if not df_data.empty:
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.subheader("إجمالي مدة الأعطال حسب المصنع/القسم")
            st.bar_chart(df_data.groupby("المصنع/القسم")["مدة العطل (ساعة)"].sum())
        with col_chart2:
            st.subheader("إجمالي مدة الأعطال حسب الوردية")
            st.bar_chart(df_data.groupby("رقم الوردية")["مدة العطل (ساعة)"].sum())
