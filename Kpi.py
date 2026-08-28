import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="شركة العلمين فليكس للطباعة - نظام KPIs المتقدم", page_icon="⚙️", layout="wide")

EXCEL_FILE = "maintenance_kpi_database.xlsx"

def load_data():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        if not df.empty and "التاريخ" in df.columns:
            df["التاريخ_dt"] = pd.to_datetime(df["التاريخ"], errors='coerce')
        return df
    return pd.DataFrame(columns=[
        "رقم الإخطار", "التاريخ", "المصنع/القسم", "رقم الوردية", "رقم/اسم الماكينة", 
        "اسم مشغل الماكينة", "اسم القائم بالصيانة", "تخصص العطل", "طبيعة الصيانة",
        "حالة الماكينة النهائية", "كود/اسم قطعة الغيار", "تكلفة قطعة الغيار (جنيه)",
        "وقت البداية", "وقت النهاية", "مدة العطل (ساعة)", "السبب الرئيسي", 
        "عطل صيانة", "تأخير مشتريات", "صيانة مخططة", "خطأ مشغل", 
        "الساعات التشغيلية المتاحة", "توافرية الصيانة (%)", "ملاحظات"
    ])

def save_data(df):
    if "التاريخ_dt" in df.columns:
        df = df.drop(columns=["التاريخ_dt"])
    df.to_excel(EXCEL_FILE, index=False)

st.title("🏭 شركة العلمين فليكس للطباعة")
st.subheader("⚙️ منظومة إدارة ومتابعة الصيانة الشاملة (CMMS & KPIs)")

tab1, tab2, tab3 = st.tabs(["📝 إدخال إخطار عطل", "📊 قاعدة البيانات والتحليلات", "📈 المؤشرات الهندسية (MTBF / MTTR)"])

with tab1:
    st.header("🚨 إخطار عطل جديد")
    
    st.subheader("⏱️ تفاصيل أوقات العطل (حساب فوري)")
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        t1_input = st.time_input("وقت بداية العطل", datetime.strptime("08:00", "%H:%M").time())
    with col_t2:
        t2_input = st.time_input("وقت نهاية العطل", datetime.strptime("09:30", "%H:%M").time())
        
    t1_dt_live = datetime.combine(datetime.today(), t1_input)
    t2_dt_live = datetime.combine(datetime.today(), t2_input)
    if t2_dt_live < t1_dt_live:
        t2_dt_live += timedelta(days=1)
    live_duration = round((t2_dt_live - t1_dt_live).total_seconds() / 3600.0, 2)
    
    with col_t3:
        st.metric("مدة العطل المحسوبة حالياً", f"{live_duration} ساعة")

    st.divider()

    with st.form("kpi_form"):
        st.subheader("📌 البيانات الأساسية للوردية والأفراد")
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
            machine_id = st.text_input("رقم / اسم الماكينة", "Bobst / Comexi / Daetwyler / Nordmeccanica")
        with col3:
            operator_name = st.text_input("اسم مشغل الماكينة", "")
            technician_name = st.text_input("اسم القائم بالصيانة", "")

        st.divider()

        st.subheader("🛠️ التصنيف الهندسي وحالة الماكينة")
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fault_type = st.selectbox("تخصص العطل", ["كهرباء", "ميكانيكا", "تحكم وآليات PLC", "هيدروليك ونيوماتيك", "كروت إلكترونية"])
            maint_nature = st.selectbox("طبيعة الصيانة", ["عطل طارئ (Emergency)", "صيانة وقائية (PM)", "تحسين وتطوير (Modification)"])
        with col_f2:
            final_status = st.selectbox("حالة الماكينة عند المغادرة", ["تشغيل كلي", "تشغيل جزئي مؤقت", "متوقفة بانتظار قطع غيار"])
            spare_part_code = st.text_input("كود / اسم قطعة الغيار المستهلكة", "بدون / Spare Part Code")
        with col_f3:
            spare_part_cost = st.number_input("تكلفة قطع الغيار التقديرية (جنيه)", min_value=0.0, value=0.0, step=50.0)
            cause_cat = st.selectbox("السبب الرئيسي للعطل", ["عطل صيانة", "تأخير مشتريات", "صيانة مخططة", "خطأ مشغل"])

        st.divider()

        st.subheader("📋 الساعات التشغيلية والملاحظات")
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            # افتراضي 8 ساعات لكل وردية (نظام 3 ورديات = 24 ساعة يومياً)
            free_hrs = st.number_input("الساعات التشغيلية المتاحة للوردية (ساعة)", min_value=0.0, value=8.0, step=0.5)
        with col_n2:
            notes = st.text_area("ملاحظات / أرقام طلبات الشراء والتفاصيل", "")

        submit = st.form_submit_button("حفظ الإخطار في Excel 💾", use_container_width=True)

        if submit:
            try:
                df_curr = load_data()
                str_date = str(date_input)
                str_t1 = t1_input.strftime("%H:%M")

                # التحقق من التكرار لنفس الماكينة والتاريخ والوردية ووقت البداية
                if not df_curr.empty:
                    duplicate_check = df_curr[
                        (df_curr["التاريخ"] == str_date) & 
                        (df_curr["المصنع/القسم"] == factory_site) & 
                        (df_curr["رقم الوردية"] == shift_num) & 
                        (df_curr["رقم/اسم الماكينة"] == machine_id) & 
                        (df_curr["وقت البداية"] == str_t1)
                    ]
                    
                    if not duplicate_check.empty:
                        st.error(f"⚠️ **تم رفض الحفظ!** تم تسجيل إخطار عطل سابق لنفس الماكينة ({machine_id}) في نفس التاريخ والوردية ووقت البداية ({str_t1}).")
                        st.stop()

                duration = live_duration

                maint_h = duration if cause_cat == "عطل صيانة" else 0.0
                proc_h = duration if cause_cat == "تأخير مشتريات" else 0.0
                pm_h = duration if cause_cat == "صيانة مخططة" else 0.0
                op_h = duration if cause_cat == "خطأ مشغل" else 0.0

                total_planned = free_hrs + duration
                eff_planned = total_planned - (proc_h + op_h)
                availability = (free_hrs / eff_planned * 100) if eff_planned > 0 else 100.0

                # ---- توليد رقم إخطار العطل: شهر-سنة + مسلسل شهري يبدأ من 1 وينتهي بآخر رقم بنفس الشهر ----
                month_year_str = date_input.strftime("%m-%Y")
                if not df_curr.empty and "التاريخ_dt" in df_curr.columns:
                    monthly_count = df_curr[
                        (df_curr["التاريخ_dt"].dt.month == date_input.month) &
                        (df_curr["التاريخ_dt"].dt.year == date_input.year)
                    ].shape[0]
                else:
                    monthly_count = 0
                notification_number = f"{month_year_str}-{monthly_count + 1:03d}"

                new_row = {
                    "رقم الإخطار": notification_number,
                    "التاريخ": str_date,
                    "المصنع/القسم": factory_site,
                    "رقم الوردية": shift_num,
                    "رقم/اسم الماكينة": machine_id,
                    "اسم مشغل الماكينة": operator_name,
                    "اسم القائم بالصيانة": technician_name,
                    "تخصص العطل": fault_type,
                    "طبيعة الصيانة": maint_nature,
                    "حالة الماكينة النهائية": final_status,
                    "كود/اسم قطعة الغيار": spare_part_code,
                    "تكلفة قطعة الغيار (جنيه)": spare_part_cost,
                    "وقت البداية": str_t1,
                    "وقت النهاية": t2_input.strftime("%H:%M"),
                    "مدة العطل (ساعة)": duration,
                    "السبب الرئيسي": cause_cat,
                    "عطل صيانة": maint_h,
                    "تأخير مشتريات": proc_h,
                    "صيانة مخططة": pm_h,
                    "خطأ مشغل": op_h,
                    "الساعات التشغيلية المتاحة": free_hrs,
                    "توافرية الصيانة (%)": round(availability, 2),
                    "ملاحظات": notes
                }
                save_data(pd.concat([df_curr, pd.DataFrame([new_row])], ignore_index=True))
                st.success(f"تم الحفظ بنجاح! رقم الإخطار: {notification_number} | مدة العطل: {duration} ساعة | نسبة التوافرية للوردية: {availability:.2f}%")
                st.rerun()
            except Exception as e:
                st.error(f"حدث خطأ أثناء إدخال البيانات: {e}")

def filter_by_date_range(df, key_prefix=""):
    if df.empty or "التاريخ_dt" not in df.columns:
        return df
    
    st.sidebar.header("🗓️ تحديد الفترة الزمنية للتقارير")
    filter_type = st.sidebar.radio(
        "طريقة الفلترة:",
        ["شهري", "فترة مخصصة (من - إلى)", "الكل"],
        key=f"{key_prefix}_filter_type"
    )
    
    if filter_type == "شهري":
        df['سنة_شهر'] = df['التاريخ_dt'].dt.to_period('M')
        available_months = sorted(df['سنة_شهر'].dropna().unique().astype(str), reverse=True)
        if available_months:
            selected_month = st.sidebar.selectbox(
                "اختر الشهر/السنة:",
                available_months,
                key=f"{key_prefix}_selected_month"
            )
            return df[df['سنة_شهر'].astype(str) == selected_month]
    elif filter_type == "فترة مخصصة (من - إلى)":
        min_date = df['التاريخ_dt'].min().date() if not df['التاريخ_dt'].dropna().empty else datetime.now().date()
        max_date = df['التاريخ_dt'].max().date() if not df['التاريخ_dt'].dropna().empty else datetime.now().date()
        start_date = st.sidebar.date_input("من تاريخ:", min_date, key=f"{key_prefix}_start_date")
        end_date = st.sidebar.date_input("إلى تاريخ:", max_date, key=f"{key_prefix}_end_date")
        return df[(df['التاريخ_dt'].dt.date >= start_date) & (df['التاريخ_dt'].dt.date <= end_date)]
    
    return df

df_filtered_shared = filter_by_date_range(load_data(), key_prefix="shared")

with tab2:
    st.subheader("📊 البيانات المفلترة حسب الفترة المحددة")
    st.dataframe(df_filtered_shared.drop(columns=["التاريخ_dt", "سنة_شهر"], errors="ignore"), use_container_width=True)
    
    if not df_filtered_shared.empty:
        col_chart1, col_chart2, col_chart3 = st.columns(3)
        with col_chart1:
            st.subheader("ساعات التوقف حسب المصنع")
            st.bar_chart(df_filtered_shared.groupby("المصنع/القسم")["مدة العطل (ساعة)"].sum())
        with col_chart2:
            st.subheader("ساعات التوقف حسب تخصص العطل")
            st.bar_chart(df_filtered_shared.groupby("تخصص العطل")["مدة العطل (ساعة)"].sum())
        with col_chart3:
            st.subheader("تكلفة قطع الغيار حسب القسم (جنيه)")
            st.bar_chart(df_filtered_shared.groupby("المصنع/القسم")["تكلفة قطعة الغيار (جنيه)"].sum())

with tab3:
    st.header("📈 مؤشرات الأداء الهندسية والاعتمادية ونظام البونص والجزاءات")
    df_kpi = df_filtered_shared
    
    if not df_kpi.empty:
        breakdown_df = df_kpi[df_kpi["السبب الرئيسي"] == "عطل صيانة"]
        
        total_downtime = breakdown_df["مدة العطل (ساعة)"].sum()
        total_failures = len(breakdown_df)
        
        mech_downtime = breakdown_df[breakdown_df["تخصص العطل"] == "ميكانيكا"]["مدة العطل (ساعة)"].sum()
        elec_downtime = breakdown_df[breakdown_df["تخصص العطل"].isin(["كهرباء", "تحكم وآليات PLC", "كروت إلكترونية"])]["مدة العطل (ساعة)"].sum()
        
        # حساب الساعات التشغيلية المتاحة بناءً على نظام 3 ورديات x 8 ساعات بدون تكرار للوردية بنفس اليوم والقسم
        unique_shifts_df = df_kpi.drop_duplicates(subset=["التاريخ", "المصنع/القسم", "رقم الوردية"])
        total_operating_hrs = unique_shifts_df["الساعات التشغيلية المتاحة"].sum()
        
        # حساب صافي ساعات التشغيل الفعلي (Uptime)
        actual_uptime = max(total_operating_hrs - total_downtime, 0.0)
        total_cost = df_kpi["تكلفة قطعة الغيار (جنيه)"].sum()
        
        mttr = round(total_downtime / total_failures, 2) if total_failures > 0 else 0.0
        mtbf = round(actual_uptime / total_failures, 2) if total_failures > 0 else 0.0
        overall_avail = round((actual_uptime / total_operating_hrs) * 100, 2) if total_operating_hrs > 0 else 100.0

        col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
        col_m1.metric("معدل وقت الإصلاح (MTTR)", f"{mttr} ساعة")
        col_m2.metric("معدل التشغيل بين الأعطال (MTBF)", f"{mtbf} ساعة")
        col_m3.metric("عدد أعطال الصيانة", f"{total_failures} عطل")
        col_m4.metric("التوافرية الإجمالية", f"{overall_avail}%")
        col_m5.metric("إجمالي تكلفة الصيانة", f"{total_cost:,.0f} ج.م")
        st.caption(f"عدد الورديات الفعلية المحسوبة: {len(unique_shifts_df)} وردية (إجمالي {total_operating_hrs:.0f} ساعة تشغيل متاح بكتلة 3 ورديات/يوم)")

        st.divider()

        st.subheader("🎯 عداد رصيد البونص والجزاءات (حد 60 ساعة مسموحة)")
        
        ALLOWED_BONUS_HOURS = 60.0
        remaining_bonus = ALLOWED_BONUS_HOURS - total_downtime
        
        b_col1, b_col2, b_col3 = st.columns(3)
        with b_col1:
            st.metric("إجمالي أعطال الصيانة التراكمية", f"{total_downtime:.2f} ساعة")
        with b_col2:
            st.metric("الحد المسموح (البونص)", f"{ALLOWED_BONUS_HOURS} ساعة")
        with b_col3:
            if remaining_bonus >= 0:
                st.metric("رصيد البونص المتبقي", f"{remaining_bonus:.2f} ساعة", delta=f"{remaining_bonus:.2f} ساعة متبقية", delta_color="normal")
            else:
                st.metric("ساعات التجاوز (الجزاءات)", f"{abs(remaining_bonus):.2f} ساعة", delta=f"-{abs(remaining_bonus):.2f} ساعة تجاوز", delta_color="inverse")

        st.markdown("##### 🔔 التنبيه التراكمي العام:")
        if total_downtime <= 45.0:
            st.success(f"🟢 **حالة ممتازة:** إجمالي توقفات الصيانة ({total_downtime:.2f} ساعة) ضمن النطاق الآمن. رصيد البونص كامل وسيتم صرف الحافز المخطط.")
        elif total_downtime <= ALLOWED_BONUS_HOURS:
            st.warning(f"🟡 **تنبيه اقتراب الحد:** إجمالي التوقفات ({total_downtime:.2f} ساعة). متبقي {remaining_bonus:.2f} ساعة فقط قبل استهلاك كامل رصيد 60 ساعة البونص وبدء الخصومات.")
        else:
            excess = abs(remaining_bonus)
            st.error(f"🔴 **تنبيه تجاوز خطير - جزاءات:** تم تجاوز حد البونص المستثنى (60 ساعة) بمقدار **{excess:.2f} ساعة**. تم وقف البونص وتطبيق لائحة الخصومات والجزاءات على القسم.")

        st.divider()
        
        st.markdown("##### ⚡ 🛠️ التنبيهات المنفصلة حسب التخصص:")
        col_warn_m, col_warn_e = st.columns(2)
        
        with col_warn_m:
            st.markdown("**قسم الميكانيكا**")
            st.metric("توقفات الميكانيكا", f"{mech_downtime:.2f} ساعة")
            if mech_downtime > 35.0:
                st.error(f"⚠️ **تنبيه ميكانيكا:** ارتفاع ملحوظ في الأعطال الميكانيكية ({mech_downtime:.2f} ساعة). يلزم مراجعة خطط الصيانة الوقائية للمكونات الميكانيكية.")
            elif mech_downtime > 25.0:
                st.warning(f"⚠️ **تحذير ميكانيكا:** الأعطال الميكانيكية بلغت {mech_downtime:.2f} ساعة.")
            else:
                st.success(f"✅ **الميكانيكا مستقرة:** {mech_downtime:.2f} ساعة.")

        with col_warn_e:
            st.markdown("**قسم الكهرباء والتحكم**")
            st.metric("توقفات الكهرباء والـ PLC", f"{elec_downtime:.2f} ساعة")
            if elec_downtime > 25.0:
                st.error(f"⚠️ **تنبيه كهرباء وتحكم:** الأعطال الكهربائية وتوقفات الإنفرترات/PLC بلغت ({elec_downtime:.2f} ساعة). تقتضي مراجعة دوائر التحكم ونظافة اللوحات.")
            elif elec_downtime > 15.0:
                st.warning(f"⚠️ **تحذير كهرباء:** الأعطال الكهربائية بلغت {elec_downtime:.2f} ساعة.")
            else:
                st.success(f"✅ **الكهرباء والتحكم مستقر:** {elec_downtime:.2f} ساعة.")

        st.divider()
        st.subheader("🎯 تحليل باريتو للأعطال للفترة المحددة (Pareto 80/20)")
        
        pareto_df = df_kpi.groupby("رقم/اسم الماكينة")["مدة العطل (ساعة)"].sum().reset_index()
        pareto_df = pareto_df.sort_values(by="مدة العطل (ساعة)", ascending=False)
        st.bar_chart(pareto_df.set_index("رقم/اسم الماكينة"))

        st.divider()
        st.subheader("🔍 استعلام: الفني الأكثر وقوعاً لأعطال في ورديته والماكينة الأكثر تسبباً بالتوقف")

        col_q1, col_q2 = st.columns(2)

        with col_q1:
            st.markdown("**👨‍🔧 ترتيب الفنيين حسب إجمالي ساعات الأعطال في ورديتهم**")
            tech_df = df_kpi.copy()
            tech_df["اسم القائم بالصيانة"] = tech_df["اسم القائم بالصيانة"].astype(str).str.strip()
            tech_df = tech_df[tech_df["اسم القائم بالصيانة"] != ""]
            tech_downtime = tech_df.groupby("اسم القائم بالصيانة").agg(
                عدد_الاعطال=("مدة العطل (ساعة)", "count"),
                اجمالي_ساعات_الاعطال=("مدة العطل (ساعة)", "sum")
            ).reset_index().sort_values(by="اجمالي_ساعات_الاعطال", ascending=False)

            if not tech_downtime.empty:
                top_tech = tech_downtime.iloc[0]
                st.metric(
                    f"🥇 الأكثر: {top_tech['اسم القائم بالصيانة']}",
                    f"{top_tech['اجمالي_ساعات_الاعطال']:.2f} ساعة",
                    delta=f"{int(top_tech['عدد_الاعطال'])} عطل"
                )
                st.dataframe(tech_downtime, use_container_width=True, hide_index=True)
                st.bar_chart(tech_downtime.set_index("اسم القائم بالصيانة")["اجمالي_ساعات_الاعطال"])
            else:
                st.info("لا توجد بيانات كافية عن الفنيين في هذه الفترة.")

        with col_q2:
            st.markdown("**⚙️ ترتيب الماكينات حسب إجمالي ساعات الأعطال المتسببة بها**")
            machine_downtime = df_kpi.groupby("رقم/اسم الماكينة").agg(
                عدد_الاعطال=("مدة العطل (ساعة)", "count"),
                اجمالي_ساعات_الاعطال=("مدة العطل (ساعة)", "sum")
            ).reset_index().sort_values(by="اجمالي_ساعات_الاعطال", ascending=False)

            if not machine_downtime.empty:
                top_machine = machine_downtime.iloc[0]
                st.metric(
                    f"🥇 الأكثر: {top_machine['رقم/اسم الماكينة']}",
                    f"{top_machine['اجمالي_ساعات_الاعطال']:.2f} ساعة",
                    delta=f"{int(top_machine['عدد_الاعطال'])} عطل"
                )
                st.dataframe(machine_downtime, use_container_width=True, hide_index=True)
            else:
                st.info("لا توجد بيانات كافية عن الماكينات في هذه الفترة.")
    else:
        st.info("لا توجد بيانات مسجلة في الفترة الزمنية المحددة.")
