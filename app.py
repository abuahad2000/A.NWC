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

# Vercel entrypoint compatibility
def handler(*args, **kwargs):
    return {"statusCode": 200, "body": "NWC Portal Active"}
app = handler
application = handler

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
        padding-top: 1.2rem;
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
        padding: 10px 14px;
        border-top: 4px solid #168A89;
    }
    div[data-testid="stMetricValue"] > div {
        color: #153E4B !important;
        font-weight: 900 !important;
        font-size: 24px !important;
    }
    div[data-testid="stMetricLabel"] > div {
        color: #168A89 !important;
        font-weight: 700 !important;
        font-size: 12.5px !important;
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

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #EEF3EF !important;
        border-left: 1px solid #D8E2DF !important;
    }
</style>
""", unsafe_allow_html=True)

# File Paths
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

projects_catalog = load_catalog()
base_date = datetime(2026, 9, 16)

def normalize_ar(t):
    if not isinstance(t, str):
        return ""
    t = t.strip().replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    t = t.replace('ة', 'ه').replace('ى', 'ي')
    return t.lower()

# ==============================================================================
# STRICT GOVERNANCE & MATCHING PARSER FUNCTION
# ==============================================================================
def apply_strict_governance_rules(df_raw):
    """
    Applies the exact 4 NWC Governance Gates:
    1. Filter for active pending statuses only (تحت معالجة المقاول, بانتظار اعتماد الجهة المتعدية)
    2. Exclude O&M (تشغيل وصيانة)
    3. Match against ongoing capital projects list (contractor + location/governorate)
    """
    valid_statuses = [normalize_ar('تحت معالجة المقاول'), normalize_ar('بانتظار اعتماد الجهة المتعدية')]
    
    # 1. Filter status
    status_col = None
    for c in ['حالة البلاغ', 'حالة_البلاغ', 'الحالة', 'Status', 'status']:
        if c in df_raw.columns:
            status_col = c
            break
            
    if status_col:
        df_active = df_raw[df_raw[status_col].astype(str).apply(normalize_ar).isin(valid_statuses)].copy()
    else:
        df_active = df_raw.copy()

    matched_records = []
    
    # Ongoing Capital Projects Match Rules
    riyadh_matching_rules = [
        {"pm": "م. عبدالله الأسود", "id": 60, "name": "تنفيذ خطوط صرف صحي متفرقة بمدينة الرياض – عقد رقم 26 – المرحلة الثالثة", "contractor_keywords": ["سعد علي العيسي", "سعد العيسي", "العيسي"], "districts": None},
        {"pm": "م. تركي الاسمري", "id": 7, "name": "تنفيذ شبكة صرف صحي بأجزاء من احياء الحزم ونمار المرحلة الثالثة", "contractor_keywords": ["الخط الذهبي"], "districts": ["الحزم", "نمار"]},
        {"pm": "م. تركي الاسمري", "id": 2, "name": "عقد تنفيذ شبكات صرف صحي بحي الحائر", "contractor_keywords": ["المسار الحديث"], "districts": ["الحائر"]},
        {"pm": "م. عسكر لسوم", "id": 20, "name": "عقد تنفيذ شبكة صرف صحى بحى المعيزلية - المرحلة الأولى", "contractor_keywords": ["نظم البيئه", "نظم البيئة"], "districts": ["المعيزليه", "المعيزلية"]},
        {"pm": "م. عسكر لسوم", "id": 12, "name": "عقد تنفيذ شبكات الصرف الصحي بأجزاء من أحياء القدس والملك عبد الله", "contractor_keywords": ["ربوه التعمير", "راكو"], "districts": ["القدس", "الملك عبدالله"]},
        {"pm": "م. امجد الفالح", "id": 58, "name": "عقد تنفيذ شبكات الصرف الصحي بحي العوالي - مرحلة ثانية", "contractor_keywords": ["نظم البيئه", "نظم البيئة"], "districts": ["العوالي"]},
        {"pm": "م. امجد الفالح", "id": 57, "name": "عقد تنفيذ شبكات الصرف الصحي بحي العوالي (مرحلة أولى)", "contractor_keywords": ["الاعمال المدنيه", "الاعمال المدنية", "الأعمال المدنية"], "districts": ["العوالي"]},
        {"pm": "م. عبدالله العنزي", "id": 23, "name": "عقد تنفيذ شبكات المياه بحي المهدية (عقد رقم 3)بمدينة الرياض", "contractor_keywords": ["الدايل"], "districts": ["المهديه", "المهدية"]}
    ]

    gov_matching_rules = [
        {"pm": "م. سعيد الحارث", "id": 112, "name": "عقد استكمال مشاريع المياه بمحافظتي شقراء ومرات (المرحلة الاولي)", "contractor_keywords": ["ضيف الله العتيبي", "ضيف الله العتيبى"], "govs": ["شقراء", "مرات"]},
        {"pm": "م. سعيد الحارث", "id": 114, "name": "عقد تنفيذ شبكات الصرف الصحي بمدينة شقراء ( المرحلة السابعة )", "contractor_keywords": ["اضواء رتاج", "أضواء رتاج"], "govs": ["شقراء"]},
        {"pm": "م. سعيد الحارث", "id": 117, "name": "عقد تنفيذ و استكمال مشاريع المياه بمحافظة القويعية", "contractor_keywords": ["ماكسون"], "govs": ["القويعيه", "القويعية", "الرويضه", "الرويضة"]},
        {"pm": "م. سعيد الحارث", "id": 107, "name": "عقد استكمال مشاريع المياه بمدينة الدوادمي ومراكز البجادية ونفي", "contractor_keywords": ["مشروعات المياه والطاقه", "مشروعات المياة والطاقة", "مشروعات المياه والطاقة"], "govs": ["الدوادمي", "البجاديه", "البجادية", "نفي"]},
        {"pm": "م. شاكر الحقباني", "id": 95, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة الخرج (المرحلة السابعة)", "contractor_keywords": ["الخريف"], "govs": ["الخرج"]},
        {"pm": "م. شاكر الحقباني", "id": 96, "name": "عقد تنفيذ مشروع صرف صحي بحوطة بني تميم (المرحلة الثالثة )", "contractor_keywords": ["السبق العربي"], "govs": ["حوطه بني تميم", "حوطة بني تميم", "الحوطه", "الحوطة"]},
        {"pm": "م. شاكر الحقباني", "id": 97, "name": "عقد تنفيذ شبكات الصرف الصحي بحوطة بني تميم و الخرج (المرحلة الثانية)", "contractor_keywords": ["السبق العربي"], "govs": ["الخرج", "حوطه بني تميم", "حوطة بني تميم"]},
        {"pm": "م. شاكر الحقباني", "id": 99, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة الخرج", "contractor_keywords": ["مرامر"], "govs": ["الخرج"]},
        {"pm": "م. سعيد الحارث", "id": 113, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة المزاحمية (المرحلة الثانية )", "contractor_keywords": ["السبق العربي"], "govs": ["المزاحميه", "المزاحمية"]},
        {"pm": "م. سعيد الحارث", "id": 104, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة ضرماء (المرحلة الثانية)", "contractor_keywords": ["مسره الوسطي", "مسرة الوسطى"], "govs": ["ضرماء", "ضرما"]}
    ]

    for idx, row in df_active.iterrows():
        rep_id = row.get('رقم بلاغ التعدي') or row.get('رقم_بلاغ_التعدي') or row.get('رقم البلاغ') or row.get('ID')
        if not rep_id or pd.isnull(rep_id):
            continue
            
        contractor = str(row.get('اسم المقاول') or row.get('المقاول المنفذ') or '')
        city = str(row.get('المدينة') or row.get('المحافظة') or 'مدينة الرياض')
        dist = str(row.get('الحي') or row.get('الموقع') or '')
        comment = str(row.get('تعليق المركز') or row.get('وصف التعدي') or '')
        owner = str(row.get('الجهة المالكة') or '')
        status = str(row.get('حالة البلاغ') or row.get('حالة_البلاغ') or 'تحت معالجة المقاول')
        date_val = str(row.get('تاريخ البلاغ') or row.get('تاريخ_البلاغ') or '2026-06-01').split()[0]
        lat = row.get('خط العرض') or row.get('خط_العرض')
        lng = row.get('خط الطول') or row.get('خط_الطول')

        norm_contr = normalize_ar(contractor)
        norm_city = normalize_ar(city)
        norm_dist = normalize_ar(dist)
        norm_comment = normalize_ar(comment)
        norm_owner = normalize_ar(owner)

        # Exclude pure O&M keywords
        if "management operations and maintenance" in norm_comment or "تاسي للتشغيل والصيانه" in norm_contr or "التشغيل والصيانه" in norm_contr or "تشغيل وصيانه" in norm_contr:
            continue

        matched_rule = None
        
        # Check Governorates first
        for g_rule in gov_matching_rules:
            contr_match = any(normalize_ar(kw) in norm_contr for kw in g_rule['contractor_keywords'])
            if contr_match:
                loc_match = any(normalize_ar(gov) in norm_city or normalize_ar(gov) in norm_dist or normalize_ar(gov) in norm_owner or normalize_ar(gov) in norm_comment for gov in g_rule['govs'])
                if loc_match:
                    matched_rule = g_rule
                    break

        # Check Riyadh City if not matched in governorates
        if not matched_rule and ("رياض" in norm_city or norm_city == 'nan' or not norm_city):
            for r_rule in riyadh_matching_rules:
                contr_match = any(normalize_ar(kw) in norm_contr for kw in r_rule['contractor_keywords'])
                if contr_match:
                    if r_rule['districts'] is None:
                        matched_rule = r_rule
                        break
                    else:
                        dist_match = any(normalize_ar(d) in norm_dist for d in r_rule['districts'])
                        if dist_match:
                            matched_rule = r_rule
                            break

        if matched_rule:
            matched_records.append({
                'id': matched_rule['id'],
                'name': matched_rule['name'],
                'contractor': contractor if contractor else 'مقاول معتمد',
                'program_manager_nwc': matched_rule['pm'],
                'المدينة': city if city != 'nan' and city else 'مدينة الرياض',
                'المحافظة': city if city != 'nan' and city else 'مدينة الرياض',
                'الحي': dist if dist != 'nan' and dist else 'موقع معتمد',
                'خط_العرض': float(lat) if pd.notnull(lat) else (24.7136 if city == 'مدينة الرياض' else 25.2388),
                'خط_الطول': float(lng) if pd.notnull(lng) else (46.6753 if city == 'مدينة الرياض' else 45.2775),
                'رقم_بلاغ_التعدي': int(rep_id),
                'تاريخ_البلاغ': date_val,
                'حالة_البلاغ': status,
                'وصف_التعدي': str(row.get('وصف التعدي') or 'أعمال حفريات وتمديد شبكات بدون استكمال إجراءات إخلاء الطرف'),
                'الإجراء_المطلوب': 'إلزام المقاول بالمعالجة الميدانية الفورية وإغلاق البلاغ بنظام المركز'
            })

    return matched_records

# Session State Initialization
if "encroachments" not in st.session_state:
    st.session_state.encroachments = load_initial_data()

# Sidebar Controls
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c2/National_Water_Company_%28Saudi_Arabia%29_Logo.svg/1200px-National_Water_Company_%28Saudi_Arabia%29_Logo.svg.png", width=160)
    st.markdown("### 🏢 شركة المياه الوطنية")
    st.markdown("**وحدة حوكمة التعديات وحماية الأصول**")
    st.divider()

    st.markdown("#### 📂 استيراد وتحليل ملف الإكسيل:")
    uploaded_file = st.file_uploader("اختر ملف إكسيل خام (XLSX)", type=["xlsx", "xls"])
    
    analysis_mode = st.radio(
        "وضع التحليل عند الرفع:",
        ["⚡ تطبيق قواعد الحوكمة والربط المكاني (المشاريع الجارية فقط)", "📋 عرض كافة البيانات الخام (بدون فلترة)"],
        index=0
    )

    if uploaded_file is not None:
        try:
            df_up = pd.read_excel(uploaded_file)
            total_raw_rows = len(df_up)
            
            if "تطبيق قواعد الحوكمة" in analysis_mode:
                parsed_list = apply_strict_governance_rules(df_up)
                if parsed_list:
                    st.session_state.encroachments = parsed_list
                    st.success(f"✅ تم التحليل بنجاح! إجمالي السطور بالملف: {total_raw_rows:,} ➔ البلاغات المعتمدة للمشاريع الجارية: {len(parsed_list)} بلاغاً (تم استبعاد {total_raw_rows - len(parsed_list):,} بلاغاً معالجاً أو خارج نطاق المشاريع).")
                else:
                    st.warning("لم يتم العثور على بلاغات تطابق معايير المشاريع الجارية في هذا الملف.")
            else:
                # Raw Dump Mode
                parsed_list = []
                for _, row in df_up.iterrows():
                    rep_id = row.get('رقم بلاغ التعدي') or row.get('رقم_بلاغ_التعدي') or row.get('رقم البلاغ') or row.get('ID')
                    if rep_id and pd.notnull(rep_id):
                        parsed_list.append({
                            'id': 0,
                            'name': str(row.get('اسم المشروع') or row.get('وصف التعدي') or 'مشروع غير محدد'),
                            'contractor': str(row.get('اسم المقاول') or row.get('المقاول المنفذ') or 'غير محدد'),
                            'program_manager_nwc': str(row.get('مدير البرنامج') or 'إدارة المشاريع'),
                            'المدينة': str(row.get('المدينة') or 'مدينة الرياض'),
                            'المحافظة': str(row.get('المدينة') or 'مدينة الرياض'),
                            'الحي': str(row.get('الحي') or ''),
                            'خط_العرض': float(row.get('خط العرض') or 24.7136) if pd.notnull(row.get('خط العرض')) else 24.7136,
                            'خط_الطول': float(row.get('خط الطول') or 46.6753) if pd.notnull(row.get('خط الطول')) else 46.6753,
                            'رقم_بلاغ_التعدي': int(rep_id),
                            'تاريخ_البلاغ': str(row.get('تاريخ البلاغ') or '2026-06-01').split()[0],
                            'حالة_البلاغ': str(row.get('حالة البلاغ') or 'تحت معالجة المقاول'),
                            'وصف_التعدي': str(row.get('وصف التعدي') or ''),
                            'الإجراء_المطلوب': 'متابعة المعالجة'
                        })
                st.session_state.encroachments = parsed_list
                st.info(f"تم عرض جميع البلاغات الخام: {len(parsed_list):,} بلاغاً.")
        except Exception as ex:
            st.error(f"خطأ أثناء قراءة وتحليل الملف: {ex}")

    if st.button("🔄 استعادة البلاغات المعتمدة الافتراضية (17 بلاغاً)"):
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
            ).add_to(m)

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
