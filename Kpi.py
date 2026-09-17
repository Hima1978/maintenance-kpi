import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import json
import streamlit.components.v1 as components

st.set_page_config(page_title="شركة العلمين فليكس للطباعة - نظام KPIs المتقدم", page_icon="⚙️", layout="wide")

EXCEL_FILE = "maintenance_kpi_database.xlsx"
MACHINES_CONFIG_FILE = "machines_config.json"

DEFAULT_MACHINES_BY_FACTORY = {
    "مصنع الطباعة": [
        "9 colors press", "8 colors press", "Comexi lam", "Nord1 lam", "Nord2 lam", "Nord3 lam",
        "Bemic slit", "Kesheng slit", "Rewinder mc", "Cheeter mc", "Slave machin",
        "Rewinder mc sm", "Wax mc", "Core cutter old", "Core cutter new"
    ],
    "مصنع السلندرات": [
        "Old engrave mc", "New engrave mc", "Chinese engrave mc", "Old cfm", "New cfm",
        "Finish mc", "Prova mc", "German Chrome tank", "Chinese chrome tank",
        "Copper Chinese tank", "German copper tank", "German nakil tank"
    ],
    "محطة السولفنت": [],
    "الخدمات (Chiller/Dryer/Compressors)": [
        "Big chiller", "Cylinders chiller", "Ink cooling conditioning chiller",
        "Keasir big air compressor", "Keasir small air compressor", "Air dryer"
    ]
}

def load_machines():
    if os.path.exists(MACHINES_CONFIG_FILE):
        try:
            with open(MACHINES_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in DEFAULT_MACHINES_BY_FACTORY.items():
                if k not in data:
                    data[k] = v
            return data
        except Exception:
            return {k: list(v) for k, v in DEFAULT_MACHINES_BY_FACTORY.items()}
    else:
        fresh = {k: list(v) for k, v in DEFAULT_MACHINES_BY_FACTORY.items()}
        save_machines(fresh)
        return fresh

def save_machines(data):
    with open(MACHINES_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "last_notification_number" not in st.session_state:
    st.session_state["last_notification_number"] = None

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

def _record_label(row):
    num = row.get("رقم الإخطار", "")
    num_part = f"{num} | " if pd.notna(num) and str(num).strip() != "" else ""
    return (
        f"{num_part}{row.get('التاريخ','')} | {row.get('المصنع/القسم','')} | "
        f"{row.get('رقم/اسم الماكينة','')} | فني: {row.get('اسم القائم بالصيانة','')} | "
        f"من {row.get('وقت البداية','')} إلى {row.get('وقت النهاية','')}"
    )

# ==================== أدوات الطباعة ====================

def _print_page_wrapper(title, body_html):
    return f"""
    <!DOCTYPE html>
    <html lang="ar" dir="rtl">
    <head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Tahoma', 'Arial', sans-serif;
            direction: rtl;
            padding: 24px;
            color: #1a1a1a;
        }}
        .report-header {{
            text-align: center;
            border-bottom: 3px solid #0d6efd;
            padding-bottom: 12px;
            margin-bottom: 20px;
        }}
        .report-header h1 {{ margin: 0; font-size: 22px; }}
        .report-header h2 {{ margin: 4px 0 0; font-size: 16px; color: #444; font-weight: normal; }}
        .meta-line {{ font-size: 12px; color: #666; margin-top: 6px; }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 16px;
            font-size: 12px;
        }}
        th, td {{
            border: 1px solid #999;
            padding: 6px 8px;
            text-align: center;
        }}
        th {{ background-color: #0d6efd; color: white; }}
        tr:nth-child(even) {{ background-color: #f2f6fc; }}
        .form-table td.label {{
            background-color: #eef3fb;
            font-weight: bold;
            width: 32%;
            text-align: right;
        }}
        .summary-box {{
            margin-top: 20px;
            padding: 12px;
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 13px;
        }}
        .signatures {{
            display: flex;
            justify-content: space-between;
            margin-top: 60px;
        }}
        .signature-box {{
            width: 30%;
            text-align: center;
            border-top: 1px solid #333;
            padding-top: 6px;
            font-size: 13px;
        }}
        .print-toolbar {{ text-align: center; margin-bottom: 16px; }}
        .print-toolbar button {{
            background-color: #0d6efd; color: white; border: none;
            padding: 8px 20px; border-radius: 6px; cursor: pointer; font-size: 14px;
        }}
        @media print {{
            body {{ padding: 0; }}
            .print-toolbar {{ display: none; }}
            @page {{ size: A4; margin: 12mm; }}
        }}
    </style>
    </head>
    <body>
    <div class="print-toolbar"><button onclick="window.print()">🖨️ طباعة</button></div>
    {body_html}
    </body>
    </html>
    """

def print_button(html_content, label="🖨️ طباعة / حفظ PDF", height=55):
    html_json = json.dumps(html_content)
    component_html = f"""
    <button id="printBtnUnique" style="background-color:#0d6efd;color:white;border:none;
    padding:10px 24px;border-radius:8px;cursor:pointer;font-size:15px;font-weight:bold;
    width:100%;">{label}</button>
    <script>
    (function() {{
        var btn = document.getElementById('printBtnUnique');
        btn.addEventListener('click', function() {{
            var content = {html_json};
            var w = window.open('', '_blank');
            w.document.open();
            w.document.write(content);
            w.document.close();
            setTimeout(function() {{ w.focus(); w.print(); }}, 400);
        }});
    }})();
    </script>
    """
    components.html(component_html, height=height)

def build_incident_form_html(row):
    fields = [
        ("رقم الإخطار", row.get("رقم الإخطار", "")),
        ("التاريخ", row.get("التاريخ", "")),
        ("المصنع/القسم", row.get("المصنع/القسم", "")),
        ("رقم الوردية", row.get("رقم الوردية", "")),
        ("رقم/اسم الماكينة", row.get("رقم/اسم الماكينة", "")),
        ("اسم مشغل الماكينة", row.get("اسم مشغل الماكينة", "")),
        ("اسم القائم بالصيانة", row.get("اسم القائم بالصيانة", "")),
        ("تخصص العطل", row.get("تخصص العطل", "")),
        ("طبيعة الصيانة", row.get("طبيعة الصيانة", "")),
        ("حالة الماكينة النهائية", row.get("حالة الماكينة النهائية", "")),
        ("كود/اسم قطعة الغيار", row.get("كود/اسم قطعة الغيار", "")),
        ("تكلفة قطعة الغيار (جنيه)", row.get("تكلفة قطعة الغيار (جنيه)", "")),
        ("وقت البداية", row.get("وقت البداية", "")),
        ("وقت النهاية", row.get("وقت النهاية", "")),
        ("مدة العطل (ساعة)", row.get("مدة العطل (ساعة)", "")),
        ("السبب الرئيسي", row.get("السبب الرئيسي", "")),
        ("الساعات التشغيلية المتاحة", row.get("الساعات التشغيلية المتاحة", "")),
        ("توافرية الصيانة (%)", row.get("توافرية الصيانة (%)", "")),
        ("ملاحظات", row.get("ملاحظات", "")),
    ]
    rows_html = "".join(
        f"<tr><td class='label'>{label}</td><td>{value}</td></tr>"
        for label, value in fields
    )
    body = f"""
    <div class="report-header">
        <h1>🏭 شركة العلمين فليكس للطباعة</h1>
        <h2>نموذج إخطار عطل - Maintenance Notification Form</h2>
        <div class="meta-line">تاريخ الطباعة: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
    </div>
    <table class="form-table">
        {rows_html}
    </table>
    <div class="signatures">
        <div class="signature-box">توقيع مشغل الماكينة</div>
        <div class="signature-box">توقيع فني الصيانة</div>
        <div class="signature-box">توقيع مشرف الصيانة</div>
    </div>
    """
    return _print_page_wrapper("نموذج إخطار عطل", body)

def build_table_report_html(df, title, subtitle="", extra_summary_html=""):
    display_cols = [c for c in df.columns if c not in ("التاريخ_dt", "سنة_شهر")]
    df_show = df[display_cols].copy()
    header_html = "".join(f"<th>{c}</th>" for c in df_show.columns)
    rows_html = ""
    for _, r in df_show.iterrows():
        cells = "".join(f"<td>{r[c]}</td>" for c in df_show.columns)
        rows_html += f"<tr>{cells}</tr>"
    body = f"""
    <div class="report-header">
        <h1>🏭 شركة العلمين فليكس للطباعة</h1>
        <h2>{title}</h2>
        <div class="meta-line">{subtitle} | تاريخ الطباعة: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
    </div>
    <table>
        <thead><tr>{header_html}</tr></thead>
        <tbody>{rows_html}</tbody>
    </table>
    {extra_summary_html}
    """
    return _print_page_wrapper(title, body)

def build_kpi_report_html(period_label, mttr, mtbf, total_failures, overall_avail, total_cost,
                           total_downtime, remaining_bonus, mech_downtime, elec_downtime,
                           top_tech_name, top_tech_hours, top_machine_name, top_machine_hours):
    if total_downtime <= 45.0:
        bonus_status = "🟢 حالة ممتازة - رصيد البونص كامل"
    elif total_downtime <= 60.0:
        bonus_status = "🟡 اقتراب من حد البونص المسموح"
    else:
        bonus_status = "🔴 تجاوز حد البونص - تطبيق الجزاءات"

    body = f"""
    <div class="report-header">
        <h1>🏭 شركة العلمين فليكس للطباعة</h1>
        <h2>تقرير المؤشرات الهندسية (KPI Report)</h2>
        <div class="meta-line">الفترة: {period_label} | تاريخ الطباعة: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
    </div>
    <table class="form-table">
        <tr><td class="label">معدل وقت الإصلاح (MTTR)</td><td>{mttr} ساعة</td></tr>
        <tr><td class="label">معدل التشغيل بين الأعطال (MTBF)</td><td>{mtbf} ساعة</td></tr>
        <tr><td class="label">عدد أعطال الصيانة</td><td>{total_failures} عطل</td></tr>
        <tr><td class="label">التوافرية الإجمالية</td><td>{overall_avail}%</td></tr>
        <tr><td class="label">إجمالي تكلفة الصيانة</td><td>{total_cost:,.0f} ج.م</td></tr>
        <tr><td class="label">إجمالي ساعات توقف الصيانة</td><td>{total_downtime:.2f} ساعة</td></tr>
        <tr><td class="label">رصيد/تجاوز البونص (حد 60 ساعة)</td><td>{remaining_bonus:.2f} ساعة</td></tr>
        <tr><td class="label">توقفات قسم الميكانيكا</td><td>{mech_downtime:.2f} ساعة</td></tr>
        <tr><td class="label">توقفات قسم الكهرباء والتحكم</td><td>{elec_downtime:.2f} ساعة</td></tr>
        <tr><td class="label">الفني الأكثر ارتباطاً بالأعطال</td><td>{top_tech_name} ({top_tech_hours:.2f} ساعة)</td></tr>
        <tr><td class="label">الماكينة الأكثر تسبباً بالتوقف</td><td>{top_machine_name} ({top_machine_hours:.2f} ساعة)</td></tr>
    </table>
    <div class="summary-box">
        <b>الحالة العامة للبونص/الجزاءات:</b> {bonus_status}
    </div>
    """
    return _print_page_wrapper("تقرير KPI", body)

st.title("🏭 شركة العلمين فليكس للطباعة")
st.subheader("⚙️ منظومة إدارة ومتابعة الصيانة الشاملة (CMMS & KPIs)")

tab1, tab2, tab3 = st.tabs(["📝 إدخال إخطار عطل", "📊 قاعدة البيانات والتحليلات", "📈 المؤشرات الهندسية (MTBF / MTTR)"])

with tab1:
    st.header("🚨 إخطار عطل جديد")

    if st.session_state.get("last_notification_number"):
        st.success(f"✅ آخر إخطار تم حفظه في هذه الجلسة: رقم **{st.session_state['last_notification_number']}**")

    st.subheader("⏱️ تفاصيل التاريخ والأوقات (حساب فوري)")
    col_t0, col_t1, col_t2, col_t3 = st.columns(4)
    with col_t0:
        date_input = st.date_input("التاريخ", datetime.now())
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

    # ---- معاينة حية لرقم الإخطار المتوقع بناءً على التاريخ المختار ----
    _preview_df = load_data()
    if not _preview_df.empty and "التاريخ_dt" in _preview_df.columns:
        _monthly_count_preview = _preview_df[
            (_preview_df["التاريخ_dt"].dt.month == date_input.month) &
            (_preview_df["التاريخ_dt"].dt.year == date_input.year)
        ].shape[0]
    else:
        _monthly_count_preview = 0
    preview_notification_number = f"{date_input.strftime('%m-%Y')}-{_monthly_count_preview + 1:03d}"
    st.info(f"🔢 **رقم الإخطار الذي سيتم تسجيله:** {preview_notification_number}")

    st.subheader("🏭 المصنع / القسم والماكينة")
    col_f1, col_f2 = st.columns(2)

    machines_data = load_machines()

    with col_f1:
        factory_site = st.selectbox("المصنع / القسم", [
            "مصنع الطباعة", 
            "مصنع السلندرات", 
            "محطة السولفنت", 
            "الخدمات (Chiller/Dryer/Compressors)"
        ])
    with col_f2:
        _machine_options = machines_data.get(factory_site, [])
        if _machine_options:
            machine_id = st.selectbox("رقم / اسم الماكينة", _machine_options, key="machine_select")
        else:
            machine_id = None
            st.selectbox("رقم / اسم الماكينة", ["— لا توجد ماكينات مضافة بعد —"], disabled=True)

    with st.expander(f"➕ إضافة ماكينة جديدة إلى قائمة: {factory_site}"):
        col_add1, col_add2 = st.columns([3, 1])
        with col_add1:
            new_machine_name = st.text_input("اسم / رقم الماكينة الجديدة", key="new_machine_input", label_visibility="collapsed", placeholder="اكتب اسم الماكينة الجديدة هنا")
        with col_add2:
            add_machine_clicked = st.button("➕ إضافة", key="add_machine_btn", use_container_width=True)
        if add_machine_clicked:
            cleaned_name = new_machine_name.strip()
            if not cleaned_name:
                st.warning("يرجى إدخال اسم الماكينة أولاً.")
            elif cleaned_name in machines_data.get(factory_site, []):
                st.warning("هذه الماكينة موجودة بالفعل في القائمة.")
            else:
                machines_data.setdefault(factory_site, []).append(cleaned_name)
                save_machines(machines_data)
                st.success(f"✅ تمت إضافة '{cleaned_name}' إلى قائمة {factory_site}.")
                st.rerun()

    st.divider()

    with st.form("kpi_form"):
        st.subheader("📌 البيانات الأساسية للوردية والأفراد")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.caption(f"📅 التاريخ المحدد: **{date_input.strftime('%Y-%m-%d')}**")
            st.caption(f"🏭 المصنع/القسم: **{factory_site}**")
            st.caption(f"⚙️ الماكينة: **{machine_id if machine_id else 'لم يتم الاختيار'}**")
            st.caption("(لتغيير أي من الحقول أعلاه، عدّلها في الأعلى قبل الحفظ)")
        with col2:
            shift_num = st.selectbox("رقم الوردية", ["الوردية الأولى (1)", "الوردية الثانية (2)", "الوردية الثالثة (3)"])
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
            if not machine_id:
                st.error("⚠️ لا يمكن الحفظ: لا توجد ماكينة مختارة لهذا القسم. أضف ماكينة أولاً من قسم '➕ إضافة ماكينة جديدة' بالأعلى.")
                st.stop()
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

                # ---- رقم إخطار العطل: شهر-سنة + مسلسل شهري يبدأ من 1 وينتهي بآخر رقم بنفس الشهر ----
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
                st.session_state["last_notification_number"] = notification_number
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

    st.divider()
    st.subheader("💾 نسخ احتياطي واستعادة قاعدة البيانات")

    col_bk1, col_bk2 = st.columns(2)

    with col_bk1:
        st.markdown("**⬇️ تنزيل نسخة احتياطية**")
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                backup_bytes = f.read()
            backup_filename = f"نسخة_احتياطية_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
            st.download_button(
                label="⬇️ تنزيل نسخة احتياطية من قاعدة البيانات",
                data=backup_bytes,
                file_name=backup_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption(f"آخر تعديل على الملف: {datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime('%Y-%m-%d %H:%M')} | عدد السجلات الحالية: {len(load_data())}")
        else:
            st.info("لا توجد قاعدة بيانات محفوظة بعد لعمل نسخة احتياطية منها.")

    with col_bk2:
        st.markdown("**⬆️ استعادة نسخة احتياطية**")
        restore_file = st.file_uploader("اختر ملف Excel (.xlsx) للاستعادة", type=["xlsx"], key="restore_uploader")
        if restore_file is not None:
            try:
                preview_restore_df = pd.read_excel(restore_file)
                st.write(f"عدد السجلات في الملف المرفوع: **{len(preview_restore_df)}** سجل")
                st.dataframe(preview_restore_df.head(5), use_container_width=True)

                expected_cols = {"التاريخ", "المصنع/القسم", "رقم/اسم الماكينة", "مدة العطل (ساعة)"}
                if not expected_cols.issubset(set(preview_restore_df.columns)):
                    st.warning("⚠️ الملف المرفوع لا يحتوي على كل الأعمدة المتوقعة لقاعدة البيانات. تأكد أنه نسخة احتياطية صحيحة قبل المتابعة.")

                confirm_restore = st.checkbox(
                    "⚠️ أؤكد استبدال قاعدة البيانات الحالية بالكامل بهذا الملف (سيتم أخذ نسخة أمان تلقائية من البيانات الحالية قبل الاستبدال)",
                    key="confirm_restore_checkbox"
                )
                if st.button("♻️ استعادة النسخة الاحتياطية الآن", type="primary", disabled=not confirm_restore, use_container_width=True):
                    if os.path.exists(EXCEL_FILE):
                        import shutil
                        safety_copy_name = f"safety_backup_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        shutil.copy(EXCEL_FILE, safety_copy_name)

                    restore_df_to_save = preview_restore_df.drop(columns=["التاريخ_dt"], errors="ignore")
                    restore_df_to_save.to_excel(EXCEL_FILE, index=False)
                    st.success(f"✅ تم استعادة قاعدة البيانات بنجاح من النسخة الاحتياطية ({len(restore_df_to_save)} سجل).")
                    st.rerun()
            except Exception as e:
                st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

    st.divider()
    st.subheader("🗑️ حذف إخطار عطل معين")

    df_all_for_delete = load_data()
    if not df_all_for_delete.empty:
        display_df = df_all_for_delete.copy()
        if "التاريخ_dt" in display_df.columns:
            display_df = display_df.sort_values(by="التاريخ_dt", ascending=False, na_position="last")

        delete_labels = {idx: _record_label(row) for idx, row in display_df.iterrows()}
        selected_idx = st.selectbox(
            "اختر إخطار العطل المراد حذفه:",
            options=list(delete_labels.keys()),
            format_func=lambda i: delete_labels[i],
            key="delete_select"
        )

        st.warning(f"سيتم حذف الإخطار التالي نهائياً:\n\n**{delete_labels[selected_idx]}**")
        confirm_delete = st.checkbox("⚠️ أؤكد رغبتي في حذف هذا الإخطار نهائياً (لا يمكن التراجع عن هذا الإجراء)", key="confirm_delete_checkbox")

        if st.button("🗑️ حذف الإخطار نهائياً", type="primary", disabled=not confirm_delete):
            df_after_delete = df_all_for_delete.drop(index=selected_idx)
            save_data(df_after_delete)
            st.success("✅ تم حذف الإخطار بنجاح.")
            st.rerun()
    else:
        st.info("لا توجد بيانات مسجلة لحذفها حالياً.")

    st.divider()
    st.subheader("🖨️ الطباعة")

    print_col1, print_col2 = st.columns(2)

    with print_col1:
        st.markdown("**📄 طباعة نموذج إخطار عطل واحد**")
        df_all_for_print = load_data()
        if not df_all_for_print.empty:
            print_display_df = df_all_for_print.copy()
            if "التاريخ_dt" in print_display_df.columns:
                print_display_df = print_display_df.sort_values(by="التاريخ_dt", ascending=False, na_position="last")
            print_labels = {idx: _record_label(row) for idx, row in print_display_df.iterrows()}
            selected_print_idx = st.selectbox(
                "اختر الإخطار المراد طباعته:",
                options=list(print_labels.keys()),
                format_func=lambda i: print_labels[i],
                key="print_select"
            )
            selected_row_for_print = print_display_df.loc[selected_print_idx]
            incident_html = build_incident_form_html(selected_row_for_print)
            print_button(incident_html, label="🖨️ طباعة نموذج الإخطار")
        else:
            st.info("لا توجد بيانات مسجلة للطباعة حالياً.")

    with print_col2:
        st.markdown("**📅 طباعة تقرير جميع الأعطال (حسب الفترة المحددة بالشريط الجانبي)**")
        if not df_filtered_shared.empty:
            monthly_report_html = build_table_report_html(
                df_filtered_shared,
                title="تقرير جميع الأعطال - Monthly Faults Report",
                subtitle="الفترة المحددة حالياً في الشريط الجانبي",
                extra_summary_html=f"""
                <div class="summary-box">
                    <b>إجمالي عدد الأعطال:</b> {len(df_filtered_shared)} |
                    <b>إجمالي ساعات التوقف:</b> {df_filtered_shared['مدة العطل (ساعة)'].sum():.2f} ساعة |
                    <b>إجمالي تكلفة قطع الغيار:</b> {df_filtered_shared['تكلفة قطعة الغيار (جنيه)'].sum():,.0f} ج.م
                </div>
                """
            )
            print_button(monthly_report_html, label="🖨️ طباعة تقرير جميع الأعطال")
        else:
            st.info("لا توجد بيانات ضمن الفترة المحددة للطباعة.")

    st.divider()

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

        st.divider()
        st.subheader("🖨️ طباعة تقرير KPI")

        top_tech_name = top_tech["اسم القائم بالصيانة"] if not tech_downtime.empty else "-"
        top_tech_hours = top_tech["اجمالي_ساعات_الاعطال"] if not tech_downtime.empty else 0.0
        top_machine_name_kpi = top_machine["رقم/اسم الماكينة"] if not machine_downtime.empty else "-"
        top_machine_hours_kpi = top_machine["اجمالي_ساعات_الاعطال"] if not machine_downtime.empty else 0.0

        kpi_report_html = build_kpi_report_html(
            period_label="الفترة المحددة حالياً في الشريط الجانبي",
            mttr=mttr, mtbf=mtbf, total_failures=total_failures,
            overall_avail=overall_avail, total_cost=total_cost,
            total_downtime=total_downtime, remaining_bonus=remaining_bonus,
            mech_downtime=mech_downtime, elec_downtime=elec_downtime,
            top_tech_name=top_tech_name, top_tech_hours=top_tech_hours,
            top_machine_name=top_machine_name_kpi, top_machine_hours=top_machine_hours_kpi
        )
        print_button(kpi_report_html, label="🖨️ طباعة تقرير المؤشرات الهندسية (KPI)")
    else:
        st.info("لا توجد بيانات مسجلة في الفترة الزمنية المحددة.")
