import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="نظام KPIs الصيانة", page_icon="⚙️", layout="wide")

EXCEL_FILE = "maintenance_kpi_database.xlsx"

def load_data():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    return pd.DataFrame(columns=[
        "التاريخ", "رقم الماكينة", "وقت البداية", "وقت النهاية", 
        "إجمالي التوقف (ساعة)", "السبب الرئيسي", "عطل صيانة", 
        "تأخير مشتريات", "صيانة مخططة", "خطأ مشغل", 
        "الساعات المجانية", "توافرية الصيانة (%)", "ملاحظات"
    ])

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False)

st.title("⚙️ منظومة متابعة أعطال الصيانة (Termux Mobile)")

tab1, tab2 = st.tabs(["📝 إدخال إخطار عطل", "📊 قاعدة البيانات والتحليلات"])

with tab1:
    with st.form("kpi_form"):
        col1, col2 = st.columns(2)
        with col1:
            date_input = st.date_input("التاريخ", datetime.now())
            machine_id = st.selectbox("رقم/اسم الماكينة", ["Bobst-01", "Bobst-02", "Comexi-Laminator", "Daetwyler-Gravostar", "Nordmeccanica-01"])
            cause_cat = st.selectbox("السبب الرئيسي", ["عطل صيانة", "تأخير مشتريات", "صيانة مخططة", "خطأ مشغل"])
            notes = st.text_input("ملاحظات / رقم طلب قطع الغيار", "")
        with col2:
            t1_str = st.text_input("وقت البداية (HH:MM)", "08:00")
            t2_str = st.text_input("وقت النهاية (HH:MM)", "09:30")
            free_hrs = st.number_input("الساعات المجانية المتاحة", min_value=0.0, value=24.0, step=1.0)

        submit = st.form_submit_button("حفظ البيانات في Excel 💾")

        if submit:
            try:
                fmt = "%H:%M"
                t1 = datetime.strptime(t1_str.strip(), fmt)
                t2 = datetime.strptime(t2_str.strip(), fmt)
                duration = (t2 - t1).total_seconds() / 3600.0
                if duration < 0:
                    duration += 24.0
                duration = round(duration, 2)

                maint_h = duration if cause_cat == "عطل صيانة" else 0.0
                proc_h = duration if cause_cat == "تأخير مشتريات" else 0.0
                pm_h = duration if cause_cat == "صيانة مخططة" else 0.0
                op_h = duration if cause_cat == "خطأ مشغل" else 0.0

                total_planned = free_hrs + duration
                eff_planned = total_planned - (proc_h + op_h)
                availability = (free_hrs / eff_planned * 100) if eff_planned > 0 else 100.0

                df_curr = load_data()
                new_row = {
                    "التاريخ": str(date_input), "رقم الماكينة": machine_id,
                    "وقت البداية": t1_str, "وقت النهاية": t2_str,
                    "إجمالي التوقف (ساعة)": duration, "السبب الرئيسي": cause_cat,
                    "عطل صيانة": maint_h, "تأخير مشتريات": proc_h,
                    "صيانة مخططة": pm_h, "خطأ مشغل": op_h,
                    "الساعات المجانية": free_hrs, "توافرية الصيانة (%)": round(availability, 2),
                    "ملاحظات": notes
                }
                save_data(pd.concat([df_curr, pd.DataFrame([new_row])], ignore_index=True))
                st.success(f"تم حفظ البيانات وحساب التوافرية ({availability:.2f}%) بنجاح!")
            except Exception as e:
                st.error(f"تأكد من كتابة الوقت بشكل صحيح (HH:MM): {e}")

with tab2:
    df_data = load_data()
    st.dataframe(df_data, use_container_width=True)
    if not df_data.empty:
        st.bar_chart(df_data.groupby("السبب الرئيسي")["إجمالي التوقف (ساعة)"].sum())
