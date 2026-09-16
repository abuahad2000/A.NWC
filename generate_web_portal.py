import os, zipfile, xml.etree.ElementTree as ET, json
import pandas as pd
from shapely.geometry import Polygon, LineString, mapping

print("Generating clean index.html...")

# 1. Load Projects Master
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])

# 2. Load Encroachments
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx"
try:
    df_enc = pd.read_excel(f_enc, header=3)
    if 'id' not in df_enc.columns:
        df_enc = pd.read_excel(f_enc)
except:
    df_enc = pd.read_excel(f_enc)

CONTRACTOR_CLEANUP = {
    'مؤسسة اضواء رتاج': 'مؤسسة أضواء رتاج للمقاولات',
    'أضواء رتاج': 'مؤسسة أضواء رتاج للمقاولات',
    'اضواء رتاج': 'مؤسسة أضواء رتاج للمقاولات',
    'شركة ماءك': 'شركة ماءك للمقاولات',
    'ماءك': 'شركة ماءك للمقاولات',
    'شركة الأعمال المدنية': 'شركة الأعمال المدنية المحدودة',
    'الأعمال المدنية': 'شركة الأعمال المدنية المحدودة',
    'شركة النمال': 'شركة النمال للمقاولات مساهمة مقفلة',
    'النمال': 'شركة النمال للمقاولات مساهمة مقفلة',
    'شركة الخريف': 'شركة الخريف لتقنية المياه والطاقة',
    'الخريف': 'شركة الخريف لتقنية المياه والطاقة',
    'شركة ضيف الله العتيبى': 'شركة ضيف الله العتيبي للمقاولات',
    'ضيف الله العتيبي': 'شركة ضيف الله العتيبي للمقاولات',
    'شركه احمد محي الدين الحرفي': 'شركة محمد أحمد الحرفي للمقاولات',
    'احمد محي الدين': 'شركة محمد أحمد الحرفي للمقاولات',
    'الحرفي': 'شركة محمد أحمد الحرفي للمقاولات',
    'الشركه الدولية لتوزيع المياه': 'الشركة الدولية لتوزيع المياه',
    'الدولية لتوزيع المياه': 'الشركة الدولية لتوزيع المياه',
    'شركة نظم تقنية المياه': 'شركة نظم تقنية المياه للمقاولات',
    'نظم تقنية المياه': 'شركة نظم تقنية المياه للمقاولات',
    'شركة برق المستقبل': 'شركة برق المستقبل للمقاولات',
    'برق المستقبل': 'شركة برق المستقبل للمقاولات',
    'شركة صلت': 'شركة صلت للمقاولات',
    'صلت': 'شركة صلت للمقاولات',
    'مجموعة سعد علي العيسى': 'مجموعة سعد علي العيسى للمقاولات',
    'سعد العيسى': 'مجموعة سعد علي العيسى للمقاولات',
    'مؤسسة العرين': 'مؤسسة العرين للمقاولات',
    'شركة الاومير': 'شركة الاومير للتجارة والمقاولات',
    'الأومير': 'شركة الاومير للتجارة والمقاولات',
    'مؤسسة ثليل': 'مؤسسة ثليل للمقاولات',
    'شركة اليمامة': 'شركة اليمامة للأعمال التجارية والمقاولات',
    'شركة مستورة': 'شركة مستورة للمقاولات',
    'شركة الفهد': 'شركة الفهد للتجارة والصناعة والمقاولات',
    'شركة البنية الاساسية': 'شركة البنية الأساسية للمقاولات',
    'شركة نظم البيئة': 'شركة نظم البيئة للمقاولات',
    'شركة داثن': 'شركة داثن للمقاولات',
    'شركة ربوة التعمير': 'شركة ربوة التعمير(راكو) للمقاولات',
    'ربوة التعمير': 'شركة ربوة التعمير(راكو) للمقاولات',
    'الشركة الوطنية لأعمال المياه': 'الشركة الوطنية لأعمال المياه',
    'الوطنية لأعمال المياه': 'الشركة الوطنية لأعمال المياه',
    'شركة الخط الذهبي': 'شركة الخط الذهبي للمقاولات',
    'شركة المسار الحديث': 'شركة المسار الحديث المحدودة',
    'شركة فداك العالمية': 'شركة فداك العالمية للمقاولات',
    'شركة العامرية': 'شركة العامرية المتحده للمقاولات',
    'مؤسسة رواكز': 'مؤسسة رواكز العالمية للمقاولات',
    'شركة ابداع الحياة': 'شركة ابداع الحياة للإستثمار'
}

def clean_contractor(val):
    if pd.isna(val) or val is None: return ''
    s = str(val).strip()
    for k, v in CONTRACTOR_CLEANUP.items():
        if k in s: return v
    return s

def clean_str(val):
    if pd.isna(val) or val is None: return '-'
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return '-' if s.lower() == 'nan' or not s else s

def norm_pm(pm_raw, sub_raw):
    pm = str(pm_raw or '')
    sub = str(sub_raw or '')
    if 'تركي' in pm or 'الاسمري' in pm: return 'م. تركي الاسمري'
    if 'عسكر' in pm or 'لسلوم' in pm: return 'م. عسكر لسلوم'
    if 'عبدالله الاسود' in pm or 'عبدالله الأسود' in pm: return 'م. عبدالله الأسود'
    if 'سفر' in pm or 'العتيبي' in pm: return 'م. سفر العتيبي'
    if 'علي الشهري' in pm or ('الشهري' in pm and 'علي' in pm): return 'م. علي الشهري'
    if 'أمجد' in pm or 'امجد' in pm or 'الفالح' in pm: return 'م. أمجد الفالح'
    if 'فهد العنزي' in pm or 'عبدالله العنزي' in pm or 'عبدالله علي' in pm: return 'م. عبدالله العنزي'
    if 'الشمالية' in sub or 'القحطاني' in pm: return 'م. علي القحطاني (الشمالية)'
    if 'الجنوبية' in sub or 'الحقباني' in pm: return 'م. شاكر الحقباني (الجنوبية)'
    if 'الغربية' in sub or 'الحارث' in pm: return 'م. سعيد الحارث (الغربية)'
    return 'مدير برنامج عام'

projects_list = []
for _, r in df_proj.iterrows():
    p_id = int(r['م'])
    c_name = clean_contractor(r.get('إسم المقاول'))
    p_name = clean_str(r.get('إسم المشروع (Ar)'))
    op_num = clean_str(r.get('الرقم التشغيلي'))
    scope = clean_str(r.get('نطاق المشروع'))
    po = clean_str(r.get('PO'))
    consultant = clean_str(r.get('إسم استشاري التنفيذ'))
    status = clean_str(r.get('مرحلة المشروع'))
    sub_prog = clean_str(r.get('البرنامج الفرعي'))
    pm_norm = norm_pm(r.get('مدير برنامج NWC'), sub_prog)
    
    proj_encs = df_enc[df_enc['id'].astype(str) == str(p_id)]
    if proj_encs.empty and c_name:
        proj_encs = df_enc[df_enc['contractor'] == c_name]
        
    enc_items = []
    for _, er in proj_encs.iterrows():
        enc_items.append({
            'id': int(er.get('رقم_بلاغ_التعدي', 0)) if pd.notnull(er.get('رقم_بلاغ_التعدي')) else 0,
            'date': clean_str(er.get('تاريخ_البلاغ')),
            'status': clean_str(er.get('حالة_البلاغ')),
            'target': clean_str(er.get('جهة_التوجيه_والتنبيه')),
            'action': clean_str(er.get('الإجراء_المطلوب')),
            'city': clean_str(er.get('المدينة')),
            'gov': clean_str(er.get('المحافظة')),
            'district': clean_str(er.get('الحي')),
            'street': clean_str(er.get('الشارع')),
            'lon': float(er.get('خط_الطول')) if pd.notnull(er.get('خط_الطول')) and str(er.get('خط_الطول')) != '-' else None,
            'lat': float(er.get('خط_العرض')) if pd.notnull(er.get('خط_العرض')) and str(er.get('خط_العرض')) != '-' else None,
            'desc': clean_str(er.get('وصف_التعدي'))
        })

    sector = 'صرف صحي' if ('صرف' in sub_prog or 'صرف' in p_name) else ('مياه' if ('مياه' in sub_prog or 'مياه' in p_name) else 'مياه وصرف')

    projects_list.append({
        'id': p_id,
        'operation_number': op_num,
        'name': p_name,
        'location': scope,
        'po': po,
        'contractor': c_name,
        'consultant': consultant,
        'status': status,
        'sector': sector,
        'sub_program': sub_prog,
        'program_manager': pm_norm,
        'prog_phone': clean_str(r.get('جوال مدير برنامج NWC')),
        'prog_email': clean_str(r.get('ايميل مدير برنامج NWC')),
        'project_manager': clean_str(r.get('مدير مشروع NWC')),
        'proj_phone': clean_str(r.get('جوال مدير مشروع NWC')),
        'proj_email': clean_str(r.get('ايميل  مدير مشروع NWC')),
        'supervisor': clean_str(r.get('مهندس مقيم - استشاري')),
        'resident_phone': clean_str(r.get('جوال مهندس مقيم - استشاري')),
        'executive_manager': clean_str(r.get('المدير التنفيذي')),
        'encroachments_count': len(enc_items),
        'encroachments': enc_items
    })

kmz_dir = r"c:/antigravity files IDE/PROGRAM/KMZ"

def extract_geojson(kmz_file, keywords):
    features = []
    with zipfile.ZipFile(os.path.join(kmz_dir, kmz_file), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if (fname is not None and fname.text) else ''
            if any(k in fname_text for k in keywords):
                for pm in f.findall('.//kml:Placemark', ns):
                    pname = pm.find('kml:name', ns)
                    pname_text = pname.text.strip() if (pname is not None and pname.text) else fname_text
                    coords_nodes = pm.findall('.//kml:coordinates', ns)
                    for cn in coords_nodes:
                        raw = cn.text.strip()
                        points = []
                        for coord in raw.split():
                            parts = coord.split(',')
                            if len(parts) >= 2:
                                try:
                                    lon, lat = float(parts[0]), float(parts[1])
                                    points.append((lon, lat))
                                except: pass
                        if len(points) >= 3:
                            try:
                                poly = Polygon(points)
                                if poly.is_valid and not poly.is_empty:
                                    poly_s = poly.simplify(0.0001, preserve_topology=True)
                                    features.append({
                                        'type': 'Feature',
                                        'properties': {'name': pname_text, 'layer': fname_text},
                                        'geometry': mapping(poly_s)
                                    })
                            except: pass
                        elif len(points) == 2:
                            try:
                                line = LineString(points)
                                features.append({
                                    'type': 'Feature',
                                    'properties': {'name': pname_text, 'layer': fname_text},
                                    'geometry': mapping(line)
                                })
                            except: pass
    return {'type': 'FeatureCollection', 'features': features}

water_geojson = extract_geojson('مشاريع المياه بمدينة الرياض .kmz', ['الجاري', 'الجارية'])
sewer_geojson = extract_geojson('- مشاريع الصرف الصحي بمدينة الرياض.kmz', ['الجارية'])

projects_json_str = json.dumps(projects_list, ensure_ascii=False)
water_geojson_str = json.dumps(water_geojson, ensure_ascii=False)
sewer_geojson_str = json.dumps(sewer_geojson, ensure_ascii=False)

html_template = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>منصة بطاقات مشاريع NWC والتعديات الجارية</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Leaflet CSS & JS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS (XLSX Export) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- JSZip (KMZ Export) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <!-- Google Fonts: Cairo & Tajawal -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&family=Tajawal:wght@400;500;700;800&display=swap" rel="stylesheet">

    <style>
        :root {
            --color-primary: #153E4B;     /* العناوين والنصوص الرئيسية */
            --color-secondary: #168A89;   /* العناوين الفرعية والشريط الجانبي */
            --color-bg-page: #FBF8F0;     /* خلفية الصفحات */
            --color-bg-phrase: #E2EFEB;   /* خلفية العبارات */
            --color-bg-card: #EEF3EF;     /* خلفية بطاقات */
            --color-border: #D8E2DF;      /* الخطوط الفاصلة */
        }

        body {
            font-family: 'Cairo', 'Sakkal Majalla', sans-serif;
            background-color: var(--color-bg-page);
            color: var(--color-primary);
        }

        .font-tajawal { font-family: 'Tajawal', sans-serif; }
        .bg-nwc-primary { background-color: var(--color-primary); }
        .text-nwc-primary { color: var(--color-primary); }
        .bg-nwc-secondary { background-color: var(--color-secondary); }
        .text-nwc-secondary { color: var(--color-secondary); }
        .bg-nwc-page { background-color: var(--color-bg-page); }
        .bg-nwc-phrase { background-color: var(--color-bg-phrase); }
        .bg-nwc-card { background-color: var(--color-bg-card); }
        .border-nwc { border-color: var(--color-border); }

        ::-webkit-scrollbar { width: 7px; height: 7px; }
        ::-webkit-scrollbar-track { background: var(--color-bg-phrase); }
        ::-webkit-scrollbar-thumb { background: var(--color-secondary); border-radius: 4px; }

        .card-shadow {
            box-shadow: 0 4px 20px -2px rgba(21, 62, 75, 0.08);
            transition: all 0.25s ease-in-out;
        }
        .card-shadow:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 25px -4px rgba(21, 62, 75, 0.14);
        }

        #map { height: 580px; width: 100%; border-radius: 1rem; }
        
        .tab-btn.active {
            background-color: var(--color-secondary);
            color: #FFFFFF;
            font-weight: bold;
        }

        @media print {
            .no-print { display: none !important; }
            body { background: white !important; }
            .bg-nwc-card { background: #f9f9f9 !important; border: 1px solid #ccc !important; }
        }
    </style>
</head>
<body class="min-h-screen flex flex-col font-tajawal antialiased">

    <!-- Top Header -->
    <header class="bg-nwc-primary text-white shadow-lg sticky top-0 z-50 border-b border-nwc">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-4">
            <div class="flex items-center space-x-3 space-x-reverse">
                <div class="w-11 h-11 rounded-xl bg-nwc-secondary flex items-center justify-center text-white text-xl shadow-inner">
                    <i class="fa-solid fa-faucet-drip"></i>
                </div>
                <div>
                    <h1 class="text-xl sm:text-2xl font-black font-cairo tracking-wide flex items-center gap-2">
                        شركة المياه الوطنية <span class="text-xs bg-nwc-secondary px-2.5 py-0.5 rounded-full text-white font-bold">منصة المشاريع والتعديات</span>
                    </h1>
                    <p class="text-xs text-nwc-phrase opacity-90">لوحة المتابعة الميدانية وحوكمة التعديات الجارية بمدينة الرياض والمحافظات</p>
                </div>
            </div>

            <!-- Export Actions Toolbar -->
            <div class="flex items-center gap-2.5 no-print">
                <button onclick="exportToExcel()" class="bg-nwc-secondary hover:bg-emerald-700 text-white text-xs sm:text-sm px-3.5 py-2 rounded-lg font-bold flex items-center gap-2 shadow transition">
                    <i class="fa-solid fa-file-excel text-emerald-300"></i> تصدير Excel
                </button>
                <button onclick="exportToKMZ()" class="bg-teal-700 hover:bg-teal-800 text-white text-xs sm:text-sm px-3.5 py-2 rounded-lg font-bold flex items-center gap-2 shadow transition">
                    <i class="fa-solid fa-earth-americas text-cyan-300"></i> تصدير KMZ
                </button>
                <button onclick="window.print()" class="bg-slate-700 hover:bg-slate-800 text-white text-xs sm:text-sm px-3.5 py-2 rounded-lg font-bold flex items-center gap-2 shadow transition">
                    <i class="fa-solid fa-print"></i> طباعة PDF
                </button>
            </div>
        </div>
    </header>

    <!-- Main Container -->
    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-1 w-full space-y-6">

        <!-- Executive Metrics KPI Banner -->
        <section class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3.5">
            <div class="bg-nwc-card p-4 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-nwc-secondary">
                    <span class="text-xs font-bold">إجمالي المشاريع</span>
                    <i class="fa-solid fa-folder-tree text-lg"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-total-projects" class="text-2xl font-black text-nwc-primary">122</span>
                    <span class="text-[11px] text-gray-500 font-semibold">مشروعاً معتمداً</span>
                </div>
            </div>

            <div class="bg-nwc-card p-4 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-emerald-600">
                    <span class="text-xs font-bold">المشاريع الجارية</span>
                    <i class="fa-solid fa-person-digging text-lg"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-ongoing-projects" class="text-2xl font-black text-emerald-700">73</span>
                    <span class="text-[11px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold">تحت التنفيذ</span>
                </div>
            </div>

            <div class="bg-nwc-card p-4 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-sky-600">
                    <span class="text-xs font-bold">مسلم ابتدائي</span>
                    <i class="fa-solid fa-circle-check text-lg"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-delivered-projects" class="text-2xl font-black text-sky-700">41</span>
                    <span class="text-[11px] text-gray-500 font-semibold">منجز</span>
                </div>
            </div>

            <div class="bg-nwc-card p-4 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-rose-600">
                    <span class="text-xs font-bold">مشاريع مسحوبة</span>
                    <i class="fa-solid fa-ban text-lg"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-withdrawn-projects" class="text-2xl font-black text-rose-700">8</span>
                    <span class="text-[11px] text-rose-600 font-semibold">متوقف</span>
                </div>
            </div>

            <div class="bg-nwc-card p-4 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-amber-600">
                    <span class="text-xs font-bold">تعديات قيد المعالجة</span>
                    <i class="fa-solid fa-triangle-exclamation text-lg"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-encroachments" class="text-2xl font-black text-amber-700">64</span>
                    <span class="text-[11px] bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded font-bold">نطاق جاري</span>
                </div>
            </div>

            <div class="bg-nwc-card p-4 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-nwc-secondary">
                    <span class="text-xs font-bold">مدراء البرامج</span>
                    <i class="fa-solid fa-users-gear text-lg"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span class="text-2xl font-black text-nwc-primary">10</span>
                    <span class="text-[11px] text-gray-500 font-semibold">نطاقات رئيسية</span>
                </div>
            </div>
        </section>

        <!-- View Switcher & Search Bar -->
        <section class="bg-nwc-card p-4 rounded-2xl border border-nwc card-shadow space-y-4 no-print">
            <div class="flex flex-col md:flex-row items-center justify-between gap-4">
                <!-- Navigation Tabs -->
                <div class="flex items-center bg-nwc-phrase p-1.5 rounded-xl border border-nwc w-full md:w-auto overflow-x-auto">
                    <button onclick="switchView('cards')" id="tab-cards" class="tab-btn active px-4 py-2 rounded-lg text-xs sm:text-sm font-bold flex items-center gap-2 transition whitespace-nowrap">
                        <i class="fa-solid fa-grip"></i> بطاقات المشاريع
                    </button>
                    <button onclick="switchView('map')" id="tab-map" class="tab-btn px-4 py-2 rounded-lg text-xs sm:text-sm font-bold flex items-center gap-2 text-nwc-primary hover:text-nwc-secondary transition whitespace-nowrap">
                        <i class="fa-solid fa-map-location-dot"></i> الخريطة التفاعلية (KMZ)
                    </button>
                    <button onclick="switchView('charts')" id="tab-charts" class="tab-btn px-4 py-2 rounded-lg text-xs sm:text-sm font-bold flex items-center gap-2 text-nwc-primary hover:text-nwc-secondary transition whitespace-nowrap">
                        <i class="fa-solid fa-chart-pie"></i> الرسوم البيانية والإحصائيات
                    </button>
                </div>

                <!-- Instant Search Input -->
                <div class="relative w-full md:w-80">
                    <i class="fa-solid fa-magnifying-glass absolute right-3.5 top-3 text-gray-400"></i>
                    <input type="text" id="searchInput" oninput="applyFilters()" placeholder="بحث سريع بالمشروع، المقاول، الحي، PO..." 
                           class="w-full pl-4 pr-10 py-2 text-xs sm:text-sm bg-white border border-nwc rounded-xl focus:ring-2 focus:ring-nwc-secondary focus:outline-none transition shadow-sm">
                </div>
            </div>

            <!-- Filters Toolbar -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 pt-2 border-t border-nwc">
                <div>
                    <label class="block text-[11px] font-bold text-nwc-secondary mb-1">مدير البرنامج NWC:</label>
                    <select id="filterManager" onchange="applyFilters()" class="w-full text-xs bg-white border border-nwc rounded-lg p-2 font-semibold text-nwc-primary focus:ring-1 focus:ring-nwc-secondary">
                        <option value="all">جميع مدراء البرامج (10)</option>
                        <option value="م. عسكر لسلوم">م. عسكر لسلوم (صرف صحي - شمال الرياض)</option>
                        <option value="م. أمجد الفالح">م. أمجد الفالح (صرف صحي - غرب الرياض)</option>
                        <option value="م. تركي الاسمري">م. تركي الاسمري (صرف صحي - جنوب الرياض)</option>
                        <option value="م. سفر العتيبي">م. سفر العتيبي (مياه - شمال الرياض)</option>
                        <option value="م. عبدالله العنزي">م. عبدالله علي العنزي / م. فهد العنزي (مياه - غرب الرياض)</option>
                        <option value="م. علي الشهري">م. علي الشهري (مياه - جنوب الرياض)</option>
                        <option value="م. عبدالله الأسود">م. عبدالله الأسود العنزي (كامل الرياض - استثناء)</option>
                        <option value="م. علي القحطاني (الشمالية)">م. علي القحطاني (المحافظات الشمالية)</option>
                        <option value="م. شاكر الحقباني (الجنوبية)">م. شاكر الحقباني (المحافظات الجنوبية)</option>
                        <option value="م. سعيد الحارث (الغربية)">م. سعيد الحارث (المحافظات الغربية)</option>
                    </select>
                </div>

                <div>
                    <label class="block text-[11px] font-bold text-nwc-secondary mb-1">حالة المشروع:</label>
                    <select id="filterStatus" onchange="applyFilters()" class="w-full text-xs bg-white border border-nwc rounded-lg p-2 font-semibold text-nwc-primary focus:ring-1 focus:ring-nwc-secondary">
                        <option value="all">جميع الحالات</option>
                        <option value="جاري" selected>المشاريع الجارية فقط 🟢</option>
                        <option value="مسلم ابتدائي">مسلم ابتدائي 🔵</option>
                        <option value="مسحوب">مسحوب 🔴</option>
                    </select>
                </div>

                <div>
                    <label class="block text-[11px] font-bold text-nwc-secondary mb-1">القطاع والخدمة:</label>
                    <select id="filterSector" onchange="applyFilters()" class="w-full text-xs bg-white border border-nwc rounded-lg p-2 font-semibold text-nwc-primary focus:ring-1 focus:ring-nwc-secondary">
                        <option value="all">كافة القطاعات</option>
                        <option value="مياه">شبكات ومشاريع المياه 💧</option>
                        <option value="صرف صحي">شبكات ومشاريع الصرف الصحي 🌿</option>
                        <option value="مياه وصرف">مشاريع مياه وصرف مشتركة</option>
                    </select>
                </div>

                <div class="flex items-end">
                    <label class="flex items-center gap-2 cursor-pointer bg-nwc-phrase px-3 py-2 rounded-lg border border-nwc w-full text-xs font-bold text-nwc-primary hover:bg-emerald-100 transition">
                        <input type="checkbox" id="filterHasEncroachments" onchange="applyFilters()" class="w-4 h-4 text-nwc-secondary rounded focus:ring-nwc-secondary">
                        <span>المشاريع التي لديها تعديات معلقة فقط ⚠️</span>
                    </label>
                </div>
            </div>
        </section>

        <!-- VIEW 1: Cards View -->
        <section id="view-cards" class="space-y-4">
            <div class="flex items-center justify-between">
                <h2 class="text-base sm:text-lg font-extrabold text-nwc-primary flex items-center gap-2">
                    <i class="fa-solid fa-layer-group text-nwc-secondary"></i> قائمة بطاقات المشاريع المعروضة (<span id="showing-count">0</span> مشروعاً)
                </h2>
            </div>
            
            <div id="cardsGrid" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            </div>
        </section>

        <!-- VIEW 2: Interactive Map View -->
        <section id="view-map" class="hidden space-y-4">
            <div class="bg-nwc-card p-3 rounded-xl border border-nwc flex flex-wrap items-center justify-between gap-3">
                <div class="flex items-center gap-2 text-xs font-bold text-nwc-primary">
                    <i class="fa-solid fa-layer-group text-nwc-secondary text-sm"></i> الطبقات المكانية الثابتة:
                </div>
                <div class="flex flex-wrap items-center gap-2">
                    <button id="toggle-water" onclick="toggleMapLayer('water')" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-sky-700 text-white flex items-center gap-1.5 shadow transition">
                        <i class="fa-solid fa-water"></i> نطاقات المياه الجارية (#01579B)
                    </button>
                    <button id="toggle-sewer" onclick="toggleMapLayer('sewer')" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-700 text-white flex items-center gap-1.5 shadow transition">
                        <i class="fa-solid fa-seedling"></i> نطاقات الصرف الصحي الجارية (#097138)
                    </button>
                    <button id="toggle-enc" onclick="toggleMapLayer('enc')" class="px-3 py-1.5 rounded-lg text-xs font-bold bg-amber-600 text-white flex items-center gap-1.5 shadow transition">
                        <i class="fa-solid fa-triangle-exclamation"></i> نقاط البلاغات المعلقة
                    </button>
                    <button onclick="resetMapView()" class="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-gray-200 text-gray-700 hover:bg-gray-300 transition">
                        <i class="fa-solid fa-arrows-rotate"></i> إعادة الضبط
                    </button>
                </div>
            </div>

            <div class="bg-white p-2 rounded-2xl border border-nwc card-shadow">
                <div id="map"></div>
            </div>
        </section>

        <!-- VIEW 3: Charts & Analytics View -->
        <section id="view-charts" class="hidden space-y-6">
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div class="bg-nwc-card p-5 rounded-2xl border border-nwc card-shadow">
                    <h3 class="text-sm font-bold text-nwc-primary mb-4 flex items-center gap-2">
                        <i class="fa-solid fa-chart-column text-nwc-secondary"></i> توزيع المشاريع حسب مدراء البرامج
                    </h3>
                    <div class="h-72">
                        <canvas id="chartManagers"></canvas>
                    </div>
                </div>

                <div class="bg-nwc-card p-5 rounded-2xl border border-nwc card-shadow">
                    <h3 class="text-sm font-bold text-nwc-primary mb-4 flex items-center gap-2">
                        <i class="fa-solid fa-chart-pie text-nwc-secondary"></i> توزيع المشاريع حسب الحالة التشغيلية
                    </h3>
                    <div class="h-72 flex items-center justify-center">
                        <canvas id="chartStatus"></canvas>
                    </div>
                </div>

                <div class="bg-nwc-card p-5 rounded-2xl border border-nwc card-shadow lg:col-span-2">
                    <h3 class="text-sm font-bold text-nwc-primary mb-4 flex items-center gap-2">
                        <i class="fa-solid fa-chart-bar text-amber-600"></i> حصر بلاغات التعديات المعلقة على المقاولين الجاريين
                    </h3>
                    <div class="h-80">
                        <canvas id="chartContractors"></canvas>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- Project Details Modal -->
    <div id="projectModal" class="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 hidden flex items-center justify-center p-4">
        <div class="bg-white w-full max-w-3xl max-h-[90vh] rounded-2xl shadow-2xl border border-nwc flex flex-col overflow-hidden">
            <div class="bg-nwc-primary text-white p-4 flex items-center justify-between">
                <div class="flex items-center gap-3">
                    <span id="modalProjectId" class="bg-nwc-secondary text-white text-xs px-2.5 py-1 rounded-full font-bold">#0</span>
                    <h3 id="modalProjectName" class="text-base sm:text-lg font-bold">اسم المشروع</h3>
                </div>
                <button onclick="closeModal()" class="text-gray-300 hover:text-white text-xl font-bold p-1">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            </div>

            <div id="modalBody" class="p-6 overflow-y-auto space-y-5 bg-nwc-page text-nwc-primary text-xs sm:text-sm">
            </div>

            <div class="bg-gray-100 p-3.5 border-t border-nwc flex justify-between items-center">
                <button id="modalMapBtn" class="bg-nwc-secondary hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-xs font-bold flex items-center gap-2">
                    <i class="fa-solid fa-location-crosshairs"></i> تحديد الموقع على الخريطة
                </button>
                <button onclick="closeModal()" class="bg-gray-300 hover:bg-gray-400 text-gray-800 px-4 py-2 rounded-lg text-xs font-bold">
                    إغلاق
                </button>
            </div>
        </div>
    </div>

    <!-- Footer -->
    <footer class="bg-nwc-primary text-white text-center py-4 text-xs border-t border-nwc mt-auto">
        <p class="opacity-80">شركة المياه الوطنية (NWC) - نظام إدارة ومتابعة التعديات والمشاريع الجارية &copy; 2026</p>
    </footer>

    <!-- EMBEDDED DATA -->
    <script>
        const PROJECTS_DATA = __PROJECTS_PLACEHOLDER__;
        const WATER_GEOJSON = __WATER_PLACEHOLDER__;
        const SEWER_GEOJSON = __SEWER_PLACEHOLDER__;
    </script>

    <!-- CORE APPLICATION LOGIC -->
    <script>
        let currentView = 'cards';
        let map = null;
        let waterLayerGroup = null;
        let sewerLayerGroup = null;
        let encLayerGroup = null;
        let filteredProjects = [...PROJECTS_DATA];

        const STATUS_CONFIG = {
            'جاري': { bg: 'bg-emerald-100', text: 'text-emerald-800', border: 'border-emerald-300', dot: 'bg-emerald-500', label: 'جاري تحت التنفيذ' },
            'مسلم ابتدائي': { bg: 'bg-sky-100', text: 'text-sky-800', border: 'border-sky-300', dot: 'bg-sky-500', label: 'مسلم ابتدائي' },
            'مسحوب': { bg: 'bg-rose-100', text: 'text-rose-800', border: 'border-rose-300', dot: 'bg-rose-500', label: 'مشروع مسحوب' }
        };

        document.addEventListener('DOMContentLoaded', () => {
            applyFilters();
        });

        function switchView(viewName) {
            currentView = viewName;
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.getElementById('tab-' + viewName).classList.add('active');

            document.getElementById('view-cards').classList.add('hidden');
            document.getElementById('view-map').classList.add('hidden');
            document.getElementById('view-charts').classList.add('hidden');

            document.getElementById('view-' + viewName).classList.remove('hidden');

            if (viewName === 'map') {
                initMapIfNeeded();
                setTimeout(() => { if (map) map.invalidateSize(); }, 200);
            } else if (viewName === 'charts') {
                renderCharts();
            }
        }

        function applyFilters() {
            const search = document.getElementById('searchInput').value.trim().toLowerCase();
            const manager = document.getElementById('filterManager').value;
            const status = document.getElementById('filterStatus').value;
            const sector = document.getElementById('filterSector').value;
            const hasEnc = document.getElementById('filterHasEncroachments').checked;

            filteredProjects = PROJECTS_DATA.filter(p => {
                const matchSearch = !search || 
                    (p.name && p.name.toLowerCase().includes(search)) ||
                    (p.operation_number && p.operation_number.toLowerCase().includes(search)) ||
                    (p.po && p.po.toLowerCase().includes(search)) ||
                    (p.contractor && p.contractor.toLowerCase().includes(search)) ||
                    (p.consultant && p.consultant.toLowerCase().includes(search)) ||
                    (p.location && p.location.toLowerCase().includes(search));

                const matchManager = (manager === 'all') || (p.program_manager && p.program_manager.includes(manager));
                const matchStatus = (status === 'all') || (p.status === status);
                const matchSector = (sector === 'all') || (p.sector === sector);
                const matchEnc = !hasEnc || (p.encroachments_count > 0);

                return matchSearch && matchManager && matchStatus && matchSector && matchEnc;
            });

            document.getElementById('showing-count').innerText = filteredProjects.length;
            renderCards();

            if (map) {
                updateMapMarkers();
            }
        }

        function renderCards() {
            const grid = document.getElementById('cardsGrid');
            grid.innerHTML = '';

            if (filteredProjects.length === 0) {
                grid.innerHTML = `
                    <div class="col-span-full py-12 text-center text-gray-500 bg-nwc-card rounded-2xl border border-nwc">
                        <i class="fa-solid fa-box-open text-4xl mb-3 text-gray-400"></i>
                        <p class="font-bold">لا توجد مشاريع مطابقة لمعايير التصفية والبحث الحالية</p>
                    </div>
                `;
                return;
            }

            filteredProjects.forEach(p => {
                const st = STATUS_CONFIG[p.status] || { bg: 'bg-gray-100', text: 'text-gray-800', border: 'border-gray-300', dot: 'bg-gray-500', label: p.status };
                const progressWidth = (p.status === 'مسلم ابتدائي') ? '100%' : ((p.status === 'جاري') ? '65%' : '20%');
                const progressColor = (p.status === 'مسلم ابتدائي') ? 'bg-sky-600' : ((p.status === 'جاري') ? 'bg-emerald-600' : 'bg-rose-600');

                const cardHtml = `
                    <div class="bg-nwc-card border border-nwc rounded-2xl p-4 card-shadow flex flex-col justify-between">
                        <div>
                            <div class="flex items-center justify-between gap-2 mb-2">
                                <span class="text-xs font-bold bg-nwc-phrase text-nwc-secondary px-2.5 py-0.5 rounded-lg border border-nwc">
                                    PO: ${p.po}
                                </span>
                                <span class="text-xs px-2.5 py-0.5 rounded-full font-bold flex items-center gap-1.5 ${st.bg} ${st.text} border ${st.border}">
                                    <span class="w-2 h-2 rounded-full ${st.dot}"></span>
                                    ${st.label}
                                </span>
                            </div>

                            <h3 class="font-extrabold text-sm sm:text-base text-nwc-primary mb-1 line-clamp-2 leading-snug">
                                ${p.name}
                            </h3>
                            <p class="text-xs text-gray-500 mb-3 flex items-center gap-1">
                                <i class="fa-solid fa-location-dot text-nwc-secondary"></i> ${p.location}
                            </p>

                            <div class="w-full bg-gray-200 h-1.5 rounded-full mb-3 overflow-hidden">
                                <div class="${progressColor} h-1.5 rounded-full" style="width: ${progressWidth}"></div>
                            </div>

                            <div class="bg-white p-2.5 rounded-xl border border-nwc space-y-1.5 text-xs mb-3">
                                <div class="flex justify-between items-center">
                                    <span class="text-gray-500">المقاول:</span>
                                    <span class="font-bold text-nwc-primary truncate max-w-[170px]" title="${p.contractor}">${p.contractor}</span>
                                </div>
                                <div class="flex justify-between items-center">
                                    <span class="text-gray-500">الاستشاري:</span>
                                    <span class="font-semibold text-gray-800 truncate max-w-[170px]" title="${p.consultant}">${p.consultant}</span>
                                </div>
                                <div class="flex justify-between items-center">
                                    <span class="text-gray-500">مدير البرنامج:</span>
                                    <span class="font-bold text-nwc-secondary">${p.program_manager}</span>
                                </div>
                            </div>
                        </div>

                        <div class="pt-2 border-t border-nwc flex items-center justify-between gap-2">
                            ${
                                p.encroachments_count > 0 ?
                                `<span class="bg-amber-100 text-amber-900 border border-amber-300 text-[11px] font-bold px-2 py-1 rounded-lg flex items-center gap-1">
                                    <i class="fa-solid fa-triangle-exclamation text-amber-600"></i> ${p.encroachments_count} بلاغ معلق
                                </span>` :
                                `<span class="text-[11px] text-gray-400 font-semibold flex items-center gap-1">
                                    <i class="fa-solid fa-check text-emerald-500"></i> لا توجد تعديات
                                </span>`
                            }

                            <button onclick="openModal(${p.id})" class="bg-nwc-primary hover:bg-nwc-secondary text-white text-xs px-3 py-1.5 rounded-lg font-bold flex items-center gap-1 transition shadow-sm">
                                التفاصيل <i class="fa-solid fa-chevron-left text-[10px]"></i>
                            </button>
                        </div>
                    </div>
                `;
                grid.insertAdjacentHTML('beforeend', cardHtml);
            });
        }

        function openModal(projectId) {
            const p = PROJECTS_DATA.find(x => x.id === projectId);
            if (!p) return;

            document.getElementById('modalProjectId').innerText = '#' + p.id;
            document.getElementById('modalProjectName').innerText = p.name;

            const st = STATUS_CONFIG[p.status] || { bg: 'bg-gray-100', text: 'text-gray-800', border: 'border-gray-300', dot: 'bg-gray-500', label: p.status };

            let encHtml = '';
            if (p.encroachments && p.encroachments.length > 0) {
                encHtml = `
                    <div class="space-y-2">
                        <h4 class="font-bold text-xs text-amber-800 flex items-center gap-1.5">
                            <i class="fa-solid fa-triangle-exclamation text-amber-600"></i> التعديات المعلقة والنشطة المرتبطة بالمشروع (${p.encroachments.length} بلاغ):
                        </h4>
                        <div class="overflow-x-auto border border-nwc rounded-xl bg-white shadow-sm">
                            <table class="w-full text-right text-xs">
                                <thead class="bg-nwc-phrase text-nwc-primary font-bold border-b border-nwc">
                                    <tr>
                                        <th class="p-2">رقم البلاغ</th>
                                        <th class="p-2">التاريخ</th>
                                        <th class="p-2">الحي / الشارع</th>
                                        <th class="p-2">حالة البلاغ</th>
                                        <th class="p-2">الإجراء والتنبيه المطلوب</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-nwc">
                                    ${p.encroachments.map(e => `
                                        <tr class="hover:bg-gray-50">
                                            <td class="p-2 font-bold text-nwc-primary">#${e.id}</td>
                                            <td class="p-2 text-gray-500">${e.date}</td>
                                            <td class="p-2 font-semibold">${e.district} - ${e.street}</td>
                                            <td class="p-2">
                                                <span class="px-2 py-0.5 rounded font-bold ${e.status === 'تحت معالجة المقاول' ? 'bg-amber-100 text-amber-800' : 'bg-sky-100 text-sky-800'}">
                                                    ${e.status}
                                                </span>
                                            </td>
                                            <td class="p-2 font-bold text-nwc-secondary">${e.target}: ${e.action}</td>
                                        </tr>
                                    `).join('')}
                                </tbody>
                            </table>
                        </div>
                    </div>
                `;
            } else {
                encHtml = `
                    <div class="p-3 bg-emerald-50 border border-emerald-200 rounded-xl text-emerald-800 text-xs font-bold flex items-center gap-2">
                        <i class="fa-solid fa-circle-check text-emerald-600"></i> لا توجد بلاغات تعدي معلقة مسجلة على هذا المشروع.
                    </div>
                `;
            }

            const bodyHtml = `
                <div class="bg-white p-4 rounded-xl border border-nwc grid grid-cols-2 sm:grid-cols-4 gap-3">
                    <div>
                        <span class="block text-gray-500 text-[11px]">الرقم التشغيلي:</span>
                        <span class="font-bold text-nwc-primary">${p.operation_number}</span>
                    </div>
                    <div>
                        <span class="block text-gray-500 text-[11px]">رقم PO:</span>
                        <span class="font-bold text-nwc-primary">${p.po}</span>
                    </div>
                    <div>
                        <span class="block text-gray-500 text-[11px]">الحالة:</span>
                        <span class="inline-block mt-0.5 text-xs px-2 py-0.5 rounded font-bold ${st.bg} ${st.text}">${st.label}</span>
                    </div>
                    <div>
                        <span class="block text-gray-500 text-[11px]">القطاع:</span>
                        <span class="font-bold text-nwc-secondary">${p.sector} (${p.sub_program})</span>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-nwc space-y-3">
                    <h4 class="font-bold text-xs text-nwc-primary flex items-center gap-1.5 border-b border-nwc pb-2">
                        <i class="fa-solid fa-address-book text-nwc-secondary"></i> مصفوفة مدراء المشروع وفريق الإشراف
                    </h4>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                        <div class="bg-nwc-page p-2.5 rounded-lg border border-nwc">
                            <span class="text-gray-500 block text-[11px]">مدير برنامج NWC:</span>
                            <span class="font-bold text-nwc-primary">${p.program_manager}</span>
                            <div class="text-gray-600 mt-1 flex items-center gap-3">
                                <span><i class="fa-solid fa-phone text-nwc-secondary"></i> ${p.prog_phone}</span>
                            </div>
                        </div>
                        <div class="bg-nwc-page p-2.5 rounded-lg border border-nwc">
                            <span class="text-gray-500 block text-[11px]">مدير مشروع NWC:</span>
                            <span class="font-bold text-nwc-primary">${p.project_manager}</span>
                            <div class="text-gray-600 mt-1 flex items-center gap-3">
                                <span><i class="fa-solid fa-phone text-nwc-secondary"></i> ${p.proj_phone}</span>
                            </div>
                        </div>
                        <div class="bg-nwc-page p-2.5 rounded-lg border border-nwc">
                            <span class="text-gray-500 block text-[11px]">المقاول المنفذ:</span>
                            <span class="font-bold text-nwc-primary">${p.contractor}</span>
                        </div>
                        <div class="bg-nwc-page p-2.5 rounded-lg border border-nwc">
                            <span class="text-gray-500 block text-[11px]">استشاري التنفيذ / المشرف:</span>
                            <span class="font-bold text-nwc-primary">${p.consultant} (${p.supervisor})</span>
                        </div>
                    </div>
                </div>

                ${encHtml}
            `;

            document.getElementById('modalBody').innerHTML = bodyHtml;
            document.getElementById('projectModal').classList.remove('hidden');

            document.getElementById('modalMapBtn').onclick = () => {
                closeModal();
                switchView('map');
                if (p.encroachments && p.encroachments.length > 0 && p.encroachments[0].lat) {
                    map.flyTo([p.encroachments[0].lat, p.encroachments[0].lon], 14);
                }
            };
        }

        function closeModal() {
            document.getElementById('projectModal').classList.add('hidden');
        }

        function initMapIfNeeded() {
            if (map) return;

            map = L.map('map').setView([24.7136, 46.6753], 11);

            L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
                attribution: '&copy; NWC Riyadh GIS'
            }).addTo(map);

            waterLayerGroup = L.geoJSON(WATER_GEOJSON, {
                style: {
                    color: '#01579B',
                    weight: 2,
                    fillColor: '#0288D1',
                    fillOpacity: 0.25
                },
                onEachFeature: (feature, layer) => {
                    layer.bindPopup(`<div dir="rtl" class="text-right"><b>💧 مشروع مياه جاري:</b><br>${feature.properties.name}</div>`);
                }
            }).addTo(map);

            sewerLayerGroup = L.geoJSON(SEWER_GEOJSON, {
                style: {
                    color: '#097138',
                    weight: 2,
                    fillColor: '#2E7D32',
                    fillOpacity: 0.25
                },
                onEachFeature: (feature, layer) => {
                    layer.bindPopup(`<div dir="rtl" class="text-right"><b>🌿 مشروع صرف صحي جاري:</b><br>${feature.properties.name}</div>`);
                }
            }).addTo(map);

            encLayerGroup = L.markerClusterGroup();
            updateMapMarkers();
            map.addLayer(encLayerGroup);
        }

        function updateMapMarkers() {
            if (!encLayerGroup) return;
            encLayerGroup.clearLayers();

            filteredProjects.forEach(p => {
                if (p.encroachments) {
                    p.encroachments.forEach(e => {
                        if (e.lat && e.lon) {
                            const marker = L.circleMarker([e.lat, e.lon], {
                                radius: 7,
                                fillColor: e.status === 'تحت معالجة المقاول' ? '#F59E0B' : '#0EA5E9',
                                color: '#FFFFFF',
                                weight: 2,
                                opacity: 1,
                                fillOpacity: 0.9
                            });

                            marker.bindPopup(`
                                <div dir="rtl" class="text-right font-tajawal text-xs space-y-1.5 p-1">
                                    <div class="font-bold text-sm text-[#153E4B]">🚨 بلاغ تعدي #${e.id}</div>
                                    <div class="text-gray-600"><b>المشروع:</b> ${p.name}</div>
                                    <div class="text-gray-600"><b>المقاول:</b> ${p.contractor}</div>
                                    <div class="text-gray-600"><b>الموقع:</b> ${e.city} - ${e.district} (${e.street})</div>
                                    <div class="font-bold text-amber-700"><b>الحالة:</b> ${e.status}</div>
                                    <div class="p-1.5 bg-emerald-50 rounded text-emerald-900 border border-emerald-200">
                                        <b>التوجيه:</b> ${e.target} - ${e.action}
                                    </div>
                                </div>
                            `);
                            encLayerGroup.addLayer(marker);
                        }
                    });
                }
            });
        }

        function toggleMapLayer(type) {
            if (type === 'water') {
                if (map.hasLayer(waterLayerGroup)) {
                    map.removeLayer(waterLayerGroup);
                    document.getElementById('toggle-water').classList.replace('bg-sky-700', 'bg-gray-400');
                } else {
                    map.addLayer(waterLayerGroup);
                    document.getElementById('toggle-water').classList.replace('bg-gray-400', 'bg-sky-700');
                }
            } else if (type === 'sewer') {
                if (map.hasLayer(sewerLayerGroup)) {
                    map.removeLayer(sewerLayerGroup);
                    document.getElementById('toggle-sewer').classList.replace('bg-emerald-700', 'bg-gray-400');
                } else {
                    map.addLayer(sewerLayerGroup);
                    document.getElementById('toggle-sewer').classList.replace('bg-gray-400', 'bg-emerald-700');
                }
            } else if (type === 'enc') {
                if (map.hasLayer(encLayerGroup)) {
                    map.removeLayer(encLayerGroup);
                    document.getElementById('toggle-enc').classList.replace('bg-amber-600', 'bg-gray-400');
                } else {
                    map.addLayer(encLayerGroup);
                    document.getElementById('toggle-enc').classList.replace('bg-gray-400', 'bg-amber-600');
                }
            }
        }

        function resetMapView() {
            if (map) {
                map.setView([24.7136, 46.6753], 11);
            }
        }

        let chartManagers = null;
        let chartStatus = null;
        let chartContractors = null;

        function renderCharts() {
            const mgrCounts = {};
            PROJECTS_DATA.forEach(p => {
                mgrCounts[p.program_manager] = (mgrCounts[p.program_manager] || 0) + 1;
            });

            const ctx1 = document.getElementById('chartManagers').getContext('2d');
            if (chartManagers) chartManagers.destroy();
            chartManagers = new Chart(ctx1, {
                type: 'bar',
                data: {
                    labels: Object.keys(mgrCounts),
                    datasets: [{
                        label: 'عدد المشاريع',
                        data: Object.values(mgrCounts),
                        backgroundColor: '#168A89',
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } }
                }
            });

            const stCounts = { 'جاري': 0, 'مسلم ابتدائي': 0, 'مسحوب': 0 };
            PROJECTS_DATA.forEach(p => {
                if (stCounts[p.status] !== undefined) stCounts[p.status]++;
            });

            const ctx2 = document.getElementById('chartStatus').getContext('2d');
            if (chartStatus) chartStatus.destroy();
            chartStatus = new Chart(ctx2, {
                type: 'doughnut',
                data: {
                    labels: ['جاري تحت التنفيذ', 'مسلم ابتدائي', 'مسحوب'],
                    datasets: [{
                        data: Object.values(stCounts),
                        backgroundColor: ['#10B981', '#0EA5E9', '#F43F5E']
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false
                }
            });

            const conCounts = {};
            PROJECTS_DATA.forEach(p => {
                if (p.encroachments_count > 0 && p.contractor) {
                    conCounts[p.contractor] = (conCounts[p.contractor] || 0) + p.encroachments_count;
                }
            });

            const sortedCons = Object.entries(conCounts).sort((a, b) => b[1] - a[1]).slice(0, 8);

            const ctx3 = document.getElementById('chartContractors').getContext('2d');
            if (chartContractors) chartContractors.destroy();
            chartContractors = new Chart(ctx3, {
                type: 'bar',
                data: {
                    labels: sortedCons.map(x => x[0]),
                    datasets: [{
                        label: 'عدد البلاغات المعلقة',
                        data: sortedCons.map(x => x[1]),
                        backgroundColor: '#F59E0B',
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    indexAxis: 'y'
                }
            });
        }

        function exportToExcel() {
            const exportRows = [];
            filteredProjects.forEach(p => {
                if (p.encroachments && p.encroachments.length > 0) {
                    p.encroachments.forEach(e => {
                        exportRows.push({
                            'id': p.id,
                            'operation_number': p.operation_number,
                            'name': p.name,
                            'location': p.location,
                            'po': p.po,
                            'contractor': p.contractor,
                            'consultant': p.consultant,
                            'status': p.status,
                            'sub_program': p.sub_program,
                            'executive_manager': p.executive_manager,
                            'program_manager_nwc': p.program_manager,
                            'project_manager_nwc': p.project_manager,
                            'supervisor': p.supervisor,
                            'المدينة': e.city,
                            'المحافظة': e.gov,
                            'الحي': e.district,
                            'الشارع': e.street,
                            'خط_الطول': e.lon,
                            'خط_العرض': e.lat,
                            'رقم_بلاغ_التعدي': e.id,
                            'تاريخ_البلاغ': e.date,
                            'حالة_البلاغ': e.status,
                            'جهة_التوجيه_والتنبيه': e.target,
                            'الإجراء_المطلوب': e.action,
                            'الجهة_المالكة': 'الجهة التي تم التعدي عليها وتضررت أصولها',
                            'الجهة_المتعدية': 'إدارة الامتثال - شركة المياه الوطنية NWC',
                            'اعتماد_المركز': 'مركز مشاريع البنية التحتية بمنطقة الرياض (RIPC)',
                            'وصف_التعدي': e.desc
                        });
                    });
                } else {
                    exportRows.push({
                        'id': p.id,
                        'operation_number': p.operation_number,
                        'name': p.name,
                        'location': p.location,
                        'po': p.po,
                        'contractor': p.contractor,
                        'consultant': p.consultant,
                        'status': p.status,
                        'sub_program': p.sub_program,
                        'executive_manager': p.executive_manager,
                        'program_manager_nwc': p.program_manager,
                        'project_manager_nwc': p.project_manager,
                        'supervisor': p.supervisor,
                        'المدينة': '-',
                        'المحافظة': '-',
                        'الحي': p.location,
                        'الشارع': '-',
                        'خط_الطول': '-',
                        'خط_العرض': '-',
                        'رقم_بلاغ_التعدي': '-',
                        'تاريخ_البلاغ': '-',
                        'حالة_البلاغ': 'لا توجد تعديات',
                        'جهة_التوجيه_والتنبيه': '-',
                        'الإجراء_المطلوب': '-',
                        'الجهة_المالكة': '-',
                        'الجهة_المتعدية': 'إدارة الامتثال - شركة المياه الوطنية NWC',
                        'اعتماد_المركز': 'مركز مشاريع البنية التحتية بمنطقة الرياض (RIPC)',
                        'وصف_التعدي': '-'
                    });
                }
            });

            const ws = XLSX.utils.json_to_sheet(exportRows);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "مشاريع_NWC_والتعديات");
            XLSX.writeFile(wb, "تقرير_مشاريع_NWC_المعتمد.xlsx");
        }

        function exportToKMZ() {
            let kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>بلاغات ومشاريع NWC الجارية</name>`;

            filteredProjects.forEach(p => {
                if (p.encroachments) {
                    p.encroachments.forEach(e => {
                        if (e.lat && e.lon) {
                            kml += `
    <Placemark>
      <name>بلاغ #${e.id} - ${p.name}</name>
      <description><![CDATA[
        <b>المشروع:</b> ${p.name}<br/>
        <b>المقاول:</b> ${p.contractor}<br/>
        <b>مدير البرنامج:</b> ${p.program_manager}<br/>
        <b>الحالة:</b> ${e.status}<br/>
        <b>التوجيه:</b> ${e.target} - ${e.action}
      ]]></description>
      <Point>
        <coordinates>${e.lon},${e.lat},0</coordinates>
      </Point>
    </Placemark>`;
                        }
                    });
                }
            });

            kml += `
  </Document>
</kml>`;

            const zip = new JSZip();
            zip.file("doc.kml", kml);
            zip.generateAsync({ type: "blob" }).then(content => {
                const a = document.createElement("a");
                a.href = URL.createObjectURL(content);
                a.download = "NWC_Ongoing_Projects.kmz";
                a.click();
            });
        }
    </script>
</body>
</html>"""

final_html = html_template.replace('__PROJECTS_PLACEHOLDER__', projects_json_str)\
                          .replace('__WATER_PLACEHOLDER__', water_geojson_str)\
                          .replace('__SEWER_PLACEHOLDER__', sewer_geojson_str)

out_html_path = r"c:/antigravity files IDE/PROGRAM/index.html"
with open(out_html_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"SUCCESS: Generated standalone index.html ({len(final_html):,} bytes) at:\n{out_html_path}")
