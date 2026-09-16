import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
import folium
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
import io

# Page Config
st.set_page_config(
    page_title="بوابة حوكمة وتعديات مشاريع NWC",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Corporate Colors & RTL
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&display=swap');
    
    :root {
        --primary: #153E4B;
        --secondary: #168A89;
        --bg-page: #FBF8F0;
        --bg-phrase: #E2EFEB;
        --bg-card: #EEF3EF;
        --border-line: #D8E2DF;
    }

    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif !important;
        direction: rtl;
        text-align: right;
        background-color: #FBF8F0;
        color: #153E4B;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    /* Headers */
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Cairo', sans-serif !important;
        color: #153E4B !important;
        font-weight: 800 !important;
    }

    /* Metric Cards */
    div[data-testid="stMetric"] {
        background-color: #EEF3EF;
        border: 1px solid #D8E2DF;
        border-radius: 8px;
        padding: 12px 16px;
        border-top: 4px solid #168A89;
    }
    div[data-testid="stMetricValue"] > div {
        color: #153E4B !important;
        font-weight: 900 !important;
        font-size: 26px !important;
    }
    div[data-testid="stMetricLabel"] > div {
        color: #168A89 !important;
        font-weight: 700 !important;
        font-size: 13px !important;
    }

    /* Buttons */
    .stButton > button {
        background-color: #168A89 !important;
        color: #FFFFFF !important;
        font-family: 'Cairo', sans-serif !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
        border: none !important;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        background-color: #116867 !important;
        color: #FFFFFF !important;
    }

    /* Badges */
    .badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 700;
        background-color: #E2EFEB;
        color: #153E4B;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #EEF3EF !important;
        border-left: 1px solid #D8E2DF !important;
    }
</style>
""", unsafe_allow_html=True)

# Load Default Verified Data
DATA_FILE = "data/encroachments_verified.json"
CATALOG_FILE = "data/projects_catalog.json"

@st.cache_data
def load_initial_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

@st.cache_data
def load_catalog():
    if os.path.exists(CATALOG_FILE):
        with open(CATALOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

# Session State Initialization
if "encroachments" not in st.session_state:
    st.session_state.encroachments = load_initial_data()

projects_catalog = load_catalog()
base_date = datetime(2026, 9, 16)

# Sidebar
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c2/National_Water_Company_%28Saudi_Arabia%29_Logo.svg/1200px-National_Water_Company_%28Saudi_Arabia%29_Logo.svg.png", width=160)
    st.markdown("### 🏢 شركة المياه الوطنية")
    st.markdown("**وحدة حوكمة التعديات وحماية الأصول**")
    st.divider()

    st.markdown("#### 📂 استيراد ملف إكسيل جديد:")
    uploaded_file = st.file_uploader("اختر ملف XLSX", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df_up = pd.read_excel(uploaded_file)
            parsed_list = []
            for _, row in df_up.iterrows():
                rep_id = row.get('رقم بلاغ التعدي') or row.get('رقم_بلاغ_التعدي') or row.get('رقم البلاغ') or row.get('ID')
                status = row.get('حالة البلاغ') or row.get('حالة_البلاغ') or row.get('الحالة') or 'تحت معالجة المقاول'
                contractor = row.get('اسم المقاول') or row.get('المقاول المنفذ') or ''
                city = row.get('المدينة') or row.get('المحافظة') or 'مدينة الرياض'
                dist = row.get('الحي') or row.get('الموقع') or ''
                date_val = str(row.get('تاريخ البلاغ') or row.get('تاريخ_البلاغ') or '2026-06-01').split()[0]
                lat = row.get('خط العرض') or row.get('خط_العرض')
                lng = row.get('خط الطول') or row.get('خط_الطول')
                desc = row.get('وصف التعدي') or row.get('تعليق المركز') or ''

                matched_proj = None
                for p in projects_catalog:
                    if str(contractor) in p['contractor'] or p['contractor'] in str(contractor):
                        matched_proj = p
                        break

                if rep_id and pd.notnull(rep_id):
                    parsed_list.append({
                        'id': matched_proj['id'] if matched_proj else 0,
                        'name': matched_proj['name'] if matched_proj else (row.get('اسم المشروع') or 'مشروع رأسمالي NWC'),
                        'contractor': contractor if contractor else (matched_proj['contractor'] if matched_proj else 'مقاول معتمد'),
                        'program_manager_nwc': matched_proj['pm'] if matched_proj else (row.get('مدير البرنامج') or 'م. عبدالله الأسود'),
                        'المدينة': str(city),
                        'المحافظة': str(city),
                        'الحي': str(dist),
                        'خط_العرض': float(lat) if pd.notnull(lat) else (24.7136 if city == 'مدينة الرياض' else 25.2388),
                        'خط_الطول': float(lng) if pd.notnull(lng) else (46.6753 if city == 'مدينة الرياض' else 45.2775),
                        'رقم_بلاغ_التعدي': int(rep_id),
                        'تاريخ_البلاغ': date_val,
                        'حالة_البلاغ': str(status),
                        'وصف_التعدي': str(desc),
                        'الإجراء_المطلوب': 'إلزام المقاول بالمعالجة الميدانية الفورية وإغلاق البلاغ بنظام المركز'
                    })
            if parsed_list:
                st.session_state.encroachments = parsed_list
                st.success(f"تم استيراد {len(parsed_list)} بلاغاً بنجاح!")
        except Exception as ex:
            st.error(f"خطأ أثناء قراءة الملف: {ex}")

    if st.button("🔄 استعادة البيانات الافتراضية (17 بلاغاً)"):
        st.session_state.encroachments = load_initial_data()
        st.rerun()

    st.divider()
    st.markdown("ℹ️ **إعداد:** عبدالله الزغيبي - التعديات والحوادث")
    st.markdown("🌐 [مستودع GitHub](https://github.com/abuahad2000/A.NWC)")

# Top Header
st.title("🏢 منصة حوكمة ومتابعة تعديات المشاريع الرأسمالية (NWC)")
st.caption("القطاع الأوسط - مدينة الرياض والمحافظات | ربط مكاني وتعاقدي مدقق 100%")

# Data Calculations
raw_data = st.session_state.encroachments
for item in raw_data:
    try:
        bdate = datetime.strptime(str(item.get('تاريخ_البلاغ', '2026-06-01'))[:10], '%Y-%m-%d')
        item['age_days'] = max(0, (base_date - bdate).days)
    except:
        item['age_days'] = 0

active_items = [i for i in raw_data if i.get('حالة_البلاغ') != 'تمت المعالجة']
under_contr = [i for i in raw_data if i.get('حالة_البلاغ') == 'تحت معالجة المقاول']
under_rev = [i for i in raw_data if i.get('حالة_البلاغ') == 'بانتظار اعتماد الجهة المتعدية']
resolved = [i for i in raw_data if i.get('حالة_البلاغ') == 'تمت المعالجة']
avg_delay = round(sum(i['age_days'] for i in active_items) / len(active_items)) if active_items else 0

# KPIs Row
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
kpi1.metric("📊 إجمالي البلاغات النشطة", len(active_items), f"{len([i for i in active_items if i.get('المدينة') == 'مدينة الرياض'])} بالرياض + {len([i for i in active_items if i.get('المدينة') != 'مدينة الرياض'])} بالمحافظات")
kpi2.metric("🔧 تحت معالجة المقاول", len(under_contr), "معالجة ميدانية فورية")
kpi3.metric("📝 بانتظار اعتماد الجهة", len(under_rev), "متابعة إخلاء الطرف")
kpi4.metric("✅ تمت المعالجة", len(resolved), "منجز")
kpi5.metric("⏱️ متوسط أيام التأخير", f"{avg_delay} يوم", "من تاريخ الاستخراج")

st.markdown("<br>", unsafe_allow_html=True)

# Tabs
tab_map, tab_table, tab_manage, tab_charts = st.tabs(["🗺️ خريطة OpenStreetMap", "📑 جدول البلاغات المعتمدة", "⚡ إدارة البلاغات (CRUD)", "📈 التحليلات والشارتات"])

# TAB 1: OpenStreetMap
with tab_map:
    st.subheader("🗺️ الخريطة التفاعلية لمواقع التعديات (OpenStreetMap)")
    
    col_filter, col_spacer = st.columns([2, 4])
    with col_filter:
        map_filter = st.radio("تصفية المواقع على الخريطة:", ["الكل", "مدينة الرياض فقط", "المحافظات فقط", "تمت المعالجة"], horizontal=True)
    
    if map_filter == "مدينة الرياض فقط":
        map_data = [i for i in active_items if i.get('المدينة') == 'مدينة الرياض']
    elif map_filter == "المحافظات فقط":
        map_data = [i for i in active_items if i.get('المدينة') != 'مدينة الرياض']
    elif map_filter == "تمت المعالجة":
        map_data = resolved
    else:
        map_data = active_items

    # Render Folium Map with OpenStreetMap tiles
    m = folium.Map(location=[24.7136, 46.6753], zoom_start=8, tiles="OpenStreetMap")
    
    for item in map_data:
        lat = item.get('خط_العرض')
        lng = item.get('خط_الطول')
        if lat and lng:
            is_done = (item.get('حالة_البلاغ') == 'تمت المعالجة')
            color = 'green' if is_done else ('red' if item['age_days'] > 60 else 'cadetblue')
            
            popup_html = f"""
            <div style="font-family:'Cairo', sans-serif; direction:rtl; text-align:right; font-size:12px; min-width:200px;">
                <strong style="color:#153E4B; font-size:14px;">بلاغ #{item.get('رقم_بلاغ_التعدي')}</strong><br>
                <span style="color:#168A89; font-weight:700;">{item.get('name')}</span><br>
                <b>المقاول:</b> {item.get('contractor')}<br>
                <b>الموقع:</b> {item.get('المدينة')} - {item.get('الحي')}<br>
                <b>التأخير:</b> {item['age_days']} يوم<br>
                <b>الحالة:</b> {item.get('حالة_البلاغ')}
            </div>
            """
            folium.Marker(
                [lat, lng],
                popup=folium.Popup(popup_html, max_width=300),
                icon=folium.Icon(color=color, icon='info-sign')
            ).addTo(m)

    st_folium(m, width="100%", height=480)

# TAB 2: Data Table & Export
with tab_table:
    st.subheader("📑 سجل البلاغات المعتمدة للمشاريع الجارية")
    
    df_table = pd.DataFrame(raw_data)
    if not df_table.empty:
        display_cols = ['رقم_بلاغ_التعدي', 'name', 'contractor', 'المدينة', 'الحي', 'تاريخ_البلاغ', 'age_days', 'حالة_البلاغ', 'program_manager_nwc']
        rename_dict = {
            'رقم_بلاغ_التعدي': 'رقم البلاغ',
            'name': 'اسم المشروع',
            'contractor': 'المقاول المنفذ',
            'المدينة': 'المدينة / المحافظة',
            'الحي': 'الحي / الموقع',
            'تاريخ_البلاغ': 'تاريخ البلاغ',
            'age_days': 'أيام التأخير',
            'حالة_البلاغ': 'حالة البلاغ',
            'program_manager_nwc': 'مدير البرنامج'
        }
        df_display = df_table[[c for c in display_cols if c in df_table.columns]].rename(columns=rename_dict)
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        # Download Excel
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_display.to_excel(writer, index=False, sheet_name="البلاغات المعتمدة")
        
        st.download_button(
            label="📥 تنزيل البيانات كملف Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="بلاغات_المشاريع_الجارية_المعتمدة_NWC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# TAB 3: CRUD Management
with tab_manage:
    st.subheader("⚡ إدارة وتحديث البلاغات (إضافة / تعديل / تغيير الحالة / حذف)")
    
    col_add, col_edit = st.columns(2)
    
    with col_add:
        st.markdown("#### ➕ إضافة بلاغ تعدي جديد:")
        with st.form("add_form", clear_on_submit=True):
            add_rep_id = st.number_input("رقم بلاغ التعدي *", min_value=1, step=1, value=18000)
            add_proj_name = st.selectbox("اسم المشروع *", [p['name'] for p in projects_catalog])
            
            # Auto fill contractor and PM from catalog
            sel_p = next((p for p in projects_catalog if p['name'] == add_proj_name), None)
            add_contr = st.text_input("المقاول المنفذ *", value=sel_p['contractor'] if sel_p else "")
            add_pm = st.text_input("مدير البرنامج NWC *", value=sel_p['pm'] if sel_p else "م. عبدالله الأسود")
            
            c1, c2 = st.columns(2)
            add_city = c1.text_input("المدينة / المحافظة *", value="مدينة الرياض")
            add_dist = c2.text_input("الحي / الموقع *", value="الياسمين")
            
            c3, c4 = st.columns(2)
            add_lat = c3.number_input("خط العرض (Latitude)", value=24.7136, format="%.4f")
            add_lng = c4.number_input("خط الطول (Longitude)", value=46.6753, format="%.4f")
            
            add_date = st.date_input("تاريخ البلاغ *", value=datetime.now())
            add_status = st.selectbox("حالة البلاغ *", ["تحت معالجة المقاول", "بانتظار اعتماد الجهة المتعدية", "تمت المعالجة"])
            
            submitted = st.form_submit_button("➕ حفظ البلاغ الجديد")
            if submitted:
                new_item = {
                    'id': sel_p['id'] if sel_p else 0,
                    'name': add_proj_name,
                    'contractor': add_contr,
                    'program_manager_nwc': add_pm,
                    'المدينة': add_city,
                    'المحافظة': add_city,
                    'الحي': add_dist,
                    'خط_العرض': float(add_lat),
                    'خط_الطول': float(add_lng),
                    'رقم_بلاغ_التعدي': int(add_rep_id),
                    'تاريخ_البلاغ': str(add_date),
                    'حالة_البلاغ': add_status,
                    'الإجراء_المطلوب': 'إلزام المقاول بالمعالجة الميدانية الفورية وإغلاق البلاغ بنظام المركز'
                }
                st.session_state.encroachments.insert(0, new_item)
                st.success(f"تمت إضافة البلاغ #{add_rep_id} بنجاح!")
                st.rerun()

    with col_edit:
        st.markdown("#### ✏️ تعديل / تغيير حالة / حذف بلاغ قائم:")
        if raw_data:
            rep_options = [f"بلاغ #{i.get('رقم_بلاغ_التعدي')} - {i.get('name')[:35]}... ({i.get('حالة_البلاغ')})" for i in raw_data]
            selected_rep_idx = st.selectbox("اختر البلاغ المطلوب:", range(len(raw_data)), format_func=lambda x: rep_options[x])
            
            target_item = raw_data[selected_rep_idx]
            st.info(f"**المقاول:** {target_item.get('contractor')} | **الموقع:** {target_item.get('المدينة')} - {target_item.get('الحي')}")
            
            c_action1, c_action2 = st.columns(2)
            with c_action1:
                new_st = st.selectbox("تحديث الحالة إلى:", ["تحت معالجة المقاول", "بانتظار اعتماد الجهة المتعدية", "تمت المعالجة"], index=["تحت معالجة المقاول", "بانتظار اعتماد الجهة المتعدية", "تمت المعالجة"].index(target_item.get('حالة_البلاغ', 'تحت معالجة المقاول')))
                if st.button("💾 حفظ الحالة الجديدة"):
                    st.session_state.encroachments[selected_rep_idx]['حالة_البلاغ'] = new_st
                    st.success("تم تحديث حالة البلاغ بنجاح!")
                    st.rerun()
            
            with c_action2:
                st.write("")
                st.write("")
                if st.button("🗑️ حذف البلاغ نهائياً"):
                    del st.session_state.encroachments[selected_rep_idx]
                    st.warning("تم حذف البلاغ من السجل!")
                    st.rerun()

# TAB 4: Analytics & Charts
with tab_charts:
    st.subheader("📈 التحليلات البيانية ومؤشرات الأداء")
    
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.markdown("##### ⏳ التوزيع الزمني للبلاغات المعلقة:")
        c1_count = len([i for i in active_items if i['age_days'] <= 30])
        c2_count = len([i for i in active_items if 30 < i['age_days'] <= 60])
        c3_count = len([i for i in active_items if 60 < i['age_days'] <= 120])
        c4_count = len([i for i in active_items if i['age_days'] > 120])
        
        fig_age = px.pie(
            values=[c1_count, c2_count, c3_count, c4_count],
            names=['أقل من 30 يوماً', '31 - 60 يوماً', '61 - 120 يوماً', 'أكثر من 120 يوماً'],
            color_discrete_sequence=['#168A89', '#E2EFEB', '#153E4B', '#991B1B'],
            hole=0.45
        )
        fig_age.update_layout(font_family="Cairo", margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_age, use_container_width=True)

    with col_c2:
        st.markdown("##### 👔 البلاغات المعلقة حسب مدير البرنامج:")
        pm_counts = {}
        for i in active_items:
            pm = i.get('program_manager_nwc', 'غير محدد')
            pm_counts[pm] = pm_counts.get(pm, 0) + 1
        
        df_pm = pd.DataFrame(list(pm_counts.items()), columns=['مدير البرنامج', 'عدد البلاغات']).sort_values('عدد البلاغات', ascending=False)
        fig_pm = px.bar(df_pm, x='مدير البرنامج', y='عدد البلاغات', color_discrete_sequence=['#168A89'])
        fig_pm.update_layout(font_family="Cairo", margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_pm, use_container_width=True)
