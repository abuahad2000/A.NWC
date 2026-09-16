import json, os, pandas as pd
from datetime import datetime

# 1. Load Projects Master
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])

# 2. Load Strict Encroachments
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx"
df_enc = pd.read_excel(f_enc, header=3)

base_date = pd.to_datetime('2026-09-16')
df_enc['dt'] = pd.to_datetime(df_enc['تاريخ_البلاغ'], errors='coerce')
df_enc['age_days'] = (base_date - df_enc['dt']).dt.days

# Build clean projects and encroachments JSON
projects_list = []
for _, r in df_proj.iterrows():
    p_id = int(r['م'])
    c_name = str(r.get('إسم المقاول', '')).strip()
    p_name = str(r.get('إسم المشروع (Ar)', '')).strip()
    op_num = str(r.get('الرقم التشغيلي', '')).strip()
    scope = str(r.get('نطاق المشروع', '')).strip()
    po = str(r.get('PO', '')).replace('.0', '').strip()
    consultant = str(r.get('إسم استشاري التنفيذ', '')).strip()
    status = str(r.get('مرحلة المشروع', '')).strip()
    sub_prog = str(r.get('البرنامج الفرعي', '')).strip()
    
    pm = str(r.get('مدير برنامج NWC', ''))
    if 'تركي' in pm or 'الاسمري' in pm: pm_norm = 'م. تركي الاسمري'
    elif 'عسكر' in pm or 'لسلوم' in pm: pm_norm = 'م. عسكر لسلوم'
    elif 'عبدالله الاسود' in pm or 'عبدالله الأسود' in pm: pm_norm = 'م. عبدالله الأسود'
    elif 'سفر' in pm or 'العتيبي' in pm: pm_norm = 'م. سفر العتيبي'
    elif 'علي الشهري' in pm or ('الشهري' in pm and 'علي' in pm): pm_norm = 'م. علي الشهري'
    elif 'أمجد' in pm or 'امجد' in pm or 'الفالح' in pm: pm_norm = 'م. أمجد الفالح'
    elif 'فهد العنزي' in pm or 'عبدالله العنزي' in pm or 'عبدالله علي' in pm: pm_norm = 'م. عبدالله العنزي'
    elif 'الشمالية' in sub_prog or 'القحطاني' in pm: pm_norm = 'م. علي القحطاني (الشمالية)'
    elif 'الجنوبية' in sub_prog or 'الحقباني' in pm: pm_norm = 'م. شاكر الحقباني (الجنوبية)'
    elif 'الغربية' in sub_prog or 'الحارث' in pm: pm_norm = 'م. سعيد الحارث (الغربية)'
    else: pm_norm = 'مدير برنامج عام'

    proj_encs = df_enc[df_enc['id'] == p_id]
    
    enc_items = []
    for _, er in proj_encs.iterrows():
        age = int(er['age_days']) if pd.notnull(er['age_days']) else 0
        enc_items.append({
            'id': int(er['رقم_بلاغ_التعدي']),
            'date': str(er['تاريخ_البلاغ']),
            'age_days': age,
            'is_critical': (age > 60),
            'status': str(er['حالة_البلاغ']),
            'target': str(er['جهة_التوجيه_والتنبيه']),
            'action': str(er['الإجراء_المطلوب']),
            'city': str(er['المدينة']),
            'gov': str(er['المحافظة']),
            'district': str(er['الحي']),
            'street': str(er['الشارع']),
            'lon': float(er['خط_الطول']) if pd.notnull(er['خط_الطول']) and str(er['خط_الطول']) != '-' else None,
            'lat': float(er['خط_العرض']) if pd.notnull(er['خط_العرض']) and str(er['خط_العرض']) != '-' else None,
            'desc': str(er['وصف_التعدي'])
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
        'prog_phone': str(r.get('جوال مدير برنامج NWC', '-')).replace('.0', ''),
        'prog_email': str(r.get('ايميل مدير برنامج NWC', '-')),
        'project_manager': str(r.get('مدير مشروع NWC', '-')),
        'proj_phone': str(r.get('جوال مدير مشروع NWC', '-')).replace('.0', ''),
        'proj_email': str(r.get('ايميل  مدير مشروع NWC', '-')),
        'supervisor': str(r.get('مهندس مقيم - استشاري', '-')),
        'resident_phone': str(r.get('جوال مهندس مقيم - استشاري', '-')).replace('.0', ''),
        'executive_manager': str(r.get('المدير التنفيذي', '-')),
        'encroachments_count': len(enc_items),
        'encroachments': enc_items
    })

# Load GeoJSON
from update_all_strict import water_geojson, sewer_geojson

projects_json_str = json.dumps(projects_list, ensure_ascii=False)
water_geojson_str = json.dumps(water_geojson, ensure_ascii=False)
sewer_geojson_str = json.dumps(sewer_geojson, ensure_ascii=False)

# Build HTML with Infographic Section & PDF Pending Encroachments Print View
html_template = f"""<!DOCTYPE html>
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
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&family=Tajawal:wght@400;500;700;800&family=Sakkal+Majalla:wght@400;700&display=swap" rel="stylesheet">

    <style>
        :root {{
            --color-primary: #153E4B;     /* كحلي مائل للأخضر: العناوين والنصوص الرئيسية */
            --color-secondary: #168A89;   /* تركوازي: العناوين الفرعية والشريط الجانبي */
            --color-bg-page: #FBF8F0;     /* كريمي فاتح: خلفية الصفحات */
            --color-bg-phrase: #E2EFEB;   /* أخضر نعناعي فاتح: خلفية العبارات */
            --color-bg-card: #EEF3EF;     /* أخضر رمادي فاتح: خلفية بطاقات */
            --color-border: #D8E2DF;      /* رمادي مخضر: الخطوط الفاصلة */
            --danger: #DC2626;
            --danger-bg: #FEE2E2;
        }}

        body {{
            font-family: 'Cairo', 'Tajawal', sans-serif;
            background-color: var(--color-bg-page);
            color: var(--color-primary);
        }}

        .bg-nwc-primary {{ background-color: var(--color-primary); }}
        .text-nwc-primary {{ color: var(--color-primary); }}
        .bg-nwc-secondary {{ background-color: var(--color-secondary); }}
        .text-nwc-secondary {{ color: var(--color-secondary); }}
        .bg-nwc-page {{ background-color: var(--color-bg-page); }}
        .bg-nwc-phrase {{ background-color: var(--color-bg-phrase); }}
        .bg-nwc-card {{ background-color: var(--color-bg-card); }}
        .border-nwc {{ border-color: var(--color-border); }}

        .card-shadow {{
            box-shadow: 0 4px 20px -2px rgba(21, 62, 75, 0.08);
            transition: all 0.25s ease-in-out;
        }}
        .card-shadow:hover {{
            transform: translateY(-3px);
            box-shadow: 0 10px 25px -4px rgba(21, 62, 75, 0.14);
        }}

        #map {{ height: 560px; width: 100%; border-radius: 1rem; }}
        
        .tab-btn.active {{
            background-color: var(--color-secondary);
            color: #FFFFFF;
            font-weight: bold;
        }}

        /* Print Table Styles strictly for PDF Export */
        #print-report-container {{
            display: none;
        }}

        @media print {{
            .no-print, header, main, footer, .modal-backdrop {{
                display: none !important;
            }}
            body {{
                background: #FFFFFF !important;
                padding: 0 !important;
                margin: 0 !important;
                color: #000000 !important;
            }}
            #print-report-container {{
                display: block !important;
                width: 100% !important;
                padding: 10px 14px !important;
            }}
            .print-table {{
                width: 100% !important;
                border-collapse: collapse !important;
                font-family: 'Sakkal Majalla', 'Cairo', sans-serif !important;
                font-size: 11pt !important;
            }}
            .print-table th {{
                background-color: #153E4B !important;
                color: #FFFFFF !important;
                padding: 6px 8px !important;
                border: 1px solid #000 !important;
                text-align: center !important;
                font-weight: bold !important;
            }}
            .print-table td {{
                padding: 5px 8px !important;
                border: 1px solid #D9D9D9 !important;
                text-align: right !important;
                vertical-align: middle !important;
            }}
            .print-badge-danger {{
                background-color: #FEE2E2 !important;
                color: #DC2626 !important;
                font-weight: bold !important;
                padding: 2px 6px !important;
                border-radius: 4px !important;
                display: inline-block !important;
            }}
            .print-badge-warn {{
                background-color: #FEF3C7 !important;
                color: #D97706 !important;
                font-weight: bold !important;
                padding: 2px 6px !important;
                border-radius: 4px !important;
                display: inline-block !important;
            }}
            .print-badge-ok {{
                background-color: #E2EFEB !important;
                color: #153E4B !important;
                font-weight: bold !important;
                padding: 2px 6px !important;
                border-radius: 4px !important;
                display: inline-block !important;
            }}
            @page {{
                size: A4 portrait;
                margin: 8mm 6mm;
            }}
        }}
    </style>
</head>
<body class="min-h-screen flex flex-col antialiased">

    <!-- Top Header -->
    <header class="bg-nwc-primary text-white shadow-lg sticky top-0 z-50 border-b border-nwc no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3 flex flex-wrap items-center justify-between gap-3">
            <div class="flex items-center space-x-3 space-x-reverse">
                <div class="w-10 h-10 rounded-xl bg-nwc-secondary flex items-center justify-center text-white text-lg shadow-inner">
                    <i class="fa-solid fa-faucet-drip"></i>
                </div>
                <div>
                    <h1 class="text-base sm:text-lg font-black tracking-tight leading-none text-white">منظومة بطاقات مشاريع NWC والتعديات</h1>
                    <p class="text-[11px] text-teal-200 mt-1 font-semibold">لوحة المتابعة الجغرافية الشاملة للمشاريع الرأسمالية</p>
                </div>
            </div>

            <!-- Navigation Tabs & Actions -->
            <div class="flex items-center gap-2">
                <a href="infographic_report.html" target="_blank" class="bg-nwc-phrase text-nwc-primary hover:bg-emerald-200 text-xs sm:text-sm px-3.5 py-2 rounded-lg font-bold flex items-center gap-1.5 shadow transition border border-nwc">
                    <i class="fa-solid fa-chart-pie text-nwc-secondary"></i> إنفوجرافيك التعديات
                </a>
                <button onclick="exportToExcel()" class="bg-nwc-secondary hover:bg-emerald-700 text-white text-xs sm:text-sm px-3 py-2 rounded-lg font-bold flex items-center gap-1.5 shadow transition">
                    <i class="fa-solid fa-file-excel text-emerald-300"></i> Excel
                </button>
                <button onclick="exportToKMZ()" class="bg-teal-800 hover:bg-teal-900 text-white text-xs sm:text-sm px-3 py-2 rounded-lg font-bold flex items-center gap-1.5 shadow transition">
                    <i class="fa-solid fa-earth-americas text-cyan-300"></i> KMZ
                </button>
                <button onclick="printPendingTablePDF()" class="bg-rose-700 hover:bg-rose-800 text-white text-xs sm:text-sm px-3.5 py-2 rounded-lg font-bold flex items-center gap-1.5 shadow transition">
                    <i class="fa-solid fa-print"></i> طباعة PDF (المعلق فقط)
                </button>
            </div>
        </div>
    </header>

    <!-- Main Container -->
    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-1 w-full space-y-6 no-print">

        <!-- Executive Metrics KPI Banner -->
        <section class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3.5">
            <div class="bg-nwc-card p-3.5 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-nwc-secondary">
                    <span class="text-xs font-bold">إجمالي المشاريع</span>
                    <i class="fa-solid fa-folder-tree text-base"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-total-projects" class="text-2xl font-black text-nwc-primary">122</span>
                    <span class="text-[11px] text-gray-500 font-semibold">مشروعاً</span>
                </div>
            </div>

            <div class="bg-nwc-card p-3.5 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-emerald-700">
                    <span class="text-xs font-bold">المشاريع الجارية</span>
                    <i class="fa-solid fa-person-digging text-base"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-ongoing-projects" class="text-2xl font-black text-emerald-800">73</span>
                    <span class="text-[10px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold">تحت التنفيذ</span>
                </div>
            </div>

            <div class="bg-nwc-card p-3.5 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-sky-700">
                    <span class="text-xs font-bold">مسلم ابتدائي</span>
                    <i class="fa-solid fa-circle-check text-base"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-delivered-projects" class="text-2xl font-black text-sky-800">41</span>
                    <span class="text-[10px] text-gray-500 font-semibold">منجز</span>
                </div>
            </div>

            <div class="bg-nwc-card p-3.5 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-rose-700">
                    <span class="text-xs font-bold">مشاريع مسحوبة</span>
                    <i class="fa-solid fa-ban text-base"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-withdrawn-projects" class="text-2xl font-black text-rose-800">8</span>
                    <span class="text-[10px] text-rose-600 font-semibold">متوقف</span>
                </div>
            </div>

            <div class="bg-nwc-card p-3.5 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-amber-700">
                    <span class="text-xs font-bold">البلاغات المعلقة</span>
                    <i class="fa-solid fa-triangle-exclamation text-base"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span id="kpi-encroachments" class="text-2xl font-black text-amber-800">30</span>
                    <span class="text-[10px] bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded font-bold">نطاق جاري</span>
                </div>
            </div>

            <div class="bg-nwc-card p-3.5 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div class="flex items-center justify-between text-nwc-secondary">
                    <span class="text-xs font-bold">مدراء البرامج</span>
                    <i class="fa-solid fa-users-gear text-base"></i>
                </div>
                <div class="mt-2 flex items-baseline justify-between">
                    <span class="text-2xl font-black text-nwc-primary">10</span>
                    <span class="text-[10px] text-gray-500 font-semibold">مدراء</span>
                </div>
            </div>
        </section>

        <!-- Quick Infographic Spotlight Banner -->
        <section class="bg-gradient-to-r from-nwc-primary to-nwc-secondary p-5 rounded-2xl text-white shadow-md flex flex-wrap items-center justify-between gap-4 border border-nwc">
            <div class="flex items-center gap-3.5">
                <div class="w-12 h-12 rounded-xl bg-white/10 flex items-center justify-center text-2xl backdrop-blur">
                    📊
                </div>
                <div>
                    <h3 class="text-base font-black">إنفوجرافيك التعديات ومؤشر خطر إقفال المنصة (&gt; شهرين)</h3>
                    <p class="text-xs text-teal-100 mt-0.5">تحليل مكاني وزمني للبلاغات الـ 30 المعلقة، مع دليل التواصل المباشر لمدراء البرامج والاستشاريين.</p>
                </div>
            </div>
            <div class="flex items-center gap-2.5">
                <a href="infographic_report.html" target="_blank" class="bg-white text-nwc-primary hover:bg-teal-50 px-4 py-2 rounded-xl font-bold text-xs shadow transition flex items-center gap-2">
                    <i class="fa-solid fa-arrow-up-right-from-square text-nwc-secondary"></i> فتح الإنفوجرافيك A4
                </a>
                <button onclick="printPendingTablePDF()" class="bg-rose-600 hover:bg-rose-700 text-white px-4 py-2 rounded-xl font-bold text-xs shadow transition flex items-center gap-2">
                    <i class="fa-solid fa-file-pdf"></i> طباعة جدول المعلق وتأخيره (PDF)
                </button>
            </div>
        </section>

        <!-- Filter & Search Controls -->
        <section class="bg-nwc-card p-4 rounded-xl border border-nwc card-shadow space-y-3">
            <div class="grid grid-cols-1 md:grid-cols-4 gap-3">
                <div class="relative">
                    <i class="fa-solid fa-magnifying-glass absolute right-3 top-3 text-gray-400 text-xs"></i>
                    <input type="text" id="searchInput" oninput="applyFilters()" placeholder="بحث باسم المشروع، المقاول، الحي، أو PO..." class="w-full pl-3 pr-9 py-2 bg-white border border-nwc rounded-lg text-xs focus:ring-2 focus:ring-nwc-secondary focus:outline-none">
                </div>

                <div>
                    <select id="managerFilter" onchange="applyFilters()" class="w-full px-3 py-2 bg-white border border-nwc rounded-lg text-xs focus:ring-2 focus:ring-nwc-secondary focus:outline-none">
                        <option value="">جميع مدراء البرامج (10)</option>
                    </select>
                </div>

                <div>
                    <select id="statusFilter" onchange="applyFilters()" class="w-full px-3 py-2 bg-white border border-nwc rounded-lg text-xs focus:ring-2 focus:ring-nwc-secondary focus:outline-none">
                        <option value="">جميع حالات المشاريع</option>
                        <option value="جاري">المشاريع الجارية فقط</option>
                        <option value="مسلم ابتدائي">مسلم ابتدائي</option>
                        <option value="مسحوب">مسحوب</option>
                    </select>
                </div>

                <div>
                    <select id="encFilter" onchange="applyFilters()" class="w-full px-3 py-2 bg-white border border-nwc rounded-lg text-xs focus:ring-2 focus:ring-nwc-secondary focus:outline-none">
                        <option value="">جميع المشاريع (مع وبدون تعديات)</option>
                        <option value="has_enc">مشاريع عليها بلاغات تعدي معلقة (15)</option>
                        <option value="delayed_60">مشاريع بها بلاغات متأخرة > شهرين 🚨</option>
                        <option value="no_enc">مشاريع خالية من التعديات</option>
                    </select>
                </div>
            </div>

            <!-- Program Managers Tabs -->
            <div id="pmTabsContainer" class="flex gap-1.5 overflow-x-auto pb-1 text-xs"></div>
        </section>

        <!-- Map & Chart Section -->
        <section class="grid grid-cols-1 lg:grid-cols-3 gap-5">
            <div class="lg:col-span-2 bg-white p-3.5 rounded-xl border border-nwc card-shadow">
                <div class="flex items-center justify-between mb-2">
                    <h3 class="text-xs sm:text-sm font-black text-nwc-primary flex items-center gap-2">
                        <i class="fa-solid fa-map-location-dot text-nwc-secondary"></i> الخريطة التفاعلية للمشاريع الجارية والتعديات
                    </h3>
                    <div class="flex items-center gap-2 text-[10px] font-bold">
                        <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-[#01579B]"></span> مياه جارية</span>
                        <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-[#097138]"></span> صرف جاري</span>
                        <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-[#DC2626]"></span> بلاغ معلق</span>
                    </div>
                </div>
                <div id="map"></div>
            </div>

            <div class="bg-white p-4 rounded-xl border border-nwc card-shadow flex flex-col justify-between">
                <div>
                    <h3 class="text-xs sm:text-sm font-black text-nwc-primary mb-3 flex items-center gap-2 border-b border-nwc pb-2">
                        <i class="fa-solid fa-chart-pie text-nwc-secondary"></i> توزيع البلاغات المعلقة حسب مدير البرنامج
                    </h3>
                    <div class="h-64 relative">
                        <canvas id="pmChart"></canvas>
                    </div>
                </div>
                <div class="bg-nwc-card p-3 rounded-lg border border-nwc mt-3 text-[11px] text-nwc-primary">
                    <div class="font-bold flex items-center justify-between mb-1">
                        <span>خطر إقفال المنصة (&gt; شهرين):</span>
                        <span class="text-rose-600 font-black">16 بلاغاً (53%)</span>
                    </div>
                    <p class="text-gray-600 text-[10px]">يستوجب الإغلاق الفوري لجميع البلاغات المتجاوزة لـ 60 يوماً لتفادي إيقاف المنصة وتراخيص المقاولين.</p>
                </div>
            </div>
        </section>

        <!-- Projects Cards Grid -->
        <section class="space-y-3">
            <div class="flex items-center justify-between border-b border-nwc pb-2">
                <h3 class="text-sm font-black text-nwc-primary flex items-center gap-2">
                    <i class="fa-solid fa-cubes text-nwc-secondary"></i> بطاقات المشاريع (<span id="results-count">0</span> مشروع)
                </h3>
            </div>
            <div id="projectsGrid" class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4"></div>
        </section>

    </main>

    <!-- PRINT-ONLY PENDING ENCROACHMENTS TABLE VIEW -->
    <div id="print-report-container">
        <div style="border-bottom: 2px solid #153E4B; padding-bottom: 8px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: flex-end;">
            <div>
                <div style="font-size: 11pt; color: #168A89; font-weight: bold;">🏢 الشركة الوطنية للمياه (NWC) - القطاع الأوسط</div>
                <h2 style="font-size: 16pt; color: #153E4B; font-weight: 900; margin: 2px 0;">سجل البلاغات المعلقة للمشاريع الرأسمالية الجارية</h2>
                <div style="font-size: 10pt; color: #555;">مؤشر مدة التأخير وخطر إقفال المنصة (الحد الأقصى: 60 يوماً / شهرين) | تاريخ الاستخراج: 16 سبتمبر 2026</div>
            </div>
            <div style="text-align: left; font-size: 9.5pt; color: #153E4B; border: 1px solid #D8E2DF; background: #EEF3EF; padding: 4px 8px; border-radius: 6px;">
                <div>إجمالي البلاغات المعلقة: <strong>30 بلاغاً</strong></div>
                <div style="color: #DC2626;">متأخرة &gt; شهرين: <strong>16 بلاغاً (53%)</strong></div>
            </div>
        </div>

        <table class="print-table">
            <thead>
                <tr>
                    <th style="width: 5%;">م</th>
                    <th style="width: 8%;">رقم البلاغ</th>
                    <th style="width: 25%;">اسم المشروع الجاري</th>
                    <th style="width: 18%;">المقاول المنفذ</th>
                    <th style="width: 14%;">مدير البرنامج</th>
                    <th style="width: 12%;">الحي / الشارع</th>
                    <th style="width: 9%;">تاريخ البلاغ</th>
                    <th style="width: 9%;">أيام التأخير</th>
                </tr>
            </thead>
            <tbody id="printTableBody">
                <!-- Injected via JavaScript -->
            </tbody>
        </table>

        <div style="margin-top: 14px; border-top: 1px solid #D8E2DF; padding-top: 6px; display: flex; justify-content: space-between; font-size: 9pt; color: #555;">
            <div>نظام حوكمة وإدارة تعديات المشاريع الجارية - شركة المياه الوطنية NWC</div>
            <div style="font-weight: bold; color: #153E4B;">إعداد: عبدالله الزغيبي - التعديات والحوادث</div>
        </div>
    </div>

    <!-- Data Injection -->
    <script>
        const projects = {projects_json_str};
        const waterGeo = {water_geojson_str};
        const sewerGeo = {sewer_geojson_str};

        let map, markersCluster;
        let selectedPM = '';

        // Extract all pending encroachments flat list
        function getAllPendingEncroachments() {{
            const list = [];
            projects.forEach(p => {{
                if (p.encroachments && p.encroachments.length > 0) {{
                    p.encroachments.forEach(e => {{
                        list.push({{
                            ...e,
                            project_id: p.id,
                            project_name: p.name,
                            contractor: p.contractor,
                            program_manager: p.program_manager,
                            supervisor: p.supervisor,
                            location: p.location
                        }});
                    }});
                }}
            }});
            return list.sort((a, b) => b.age_days - a.age_days);
        }}

        // Render Print Table Body
        function renderPrintTable() {{
            const tbody = document.getElementById('printTableBody');
            if (!tbody) return;
            const encs = getAllPendingEncroachments();
            let html = '';
            encs.forEach((e, idx) => {{
                const isCrit = (e.age_days > 60);
                const badgeClass = isCrit ? 'print-badge-danger' : (e.age_days > 30 ? 'print-badge-warn' : 'print-badge-ok');
                const badgeText = isCrit ? `${{e.age_days}} يوم 🚨` : `${{e.age_days}} يوم`;
                
                html += `<tr>
                    <td style="text-align: center; font-weight: bold;">${{idx + 1}}</td>
                    <td style="text-align: center; font-weight: bold; color: #153E4B;">#${{e.id}}</td>
                    <td><strong style="color: #153E4B;">${{e.project_name}}</strong></td>
                    <td>${{e.contractor}}</td>
                    <td>${{e.program_manager}}</td>
                    <td>${{e.district}} - ${{e.street}}</td>
                    <td style="text-align: center;">${{e.date}}</td>
                    <td style="text-align: center;"><span class="${{badgeClass}}">${{badgeText}}</span></td>
                </tr>`;
            }});
            tbody.innerHTML = html;
        }}

        function printPendingTablePDF() {{
            renderPrintTable();
            window.print();
        }}

        document.addEventListener('DOMContentLoaded', () => {{
            initTabs();
            initMap();
            initChart();
            renderProjects(projects);
            renderPrintTable();
        }});

        function initTabs() {{
            const pms = Array.from(new Set(projects.map(p => p.program_manager))).filter(Boolean);
            const select = document.getElementById('managerFilter');
            const tabsCont = document.getElementById('pmTabsContainer');
            
            tabsCont.innerHTML = `<button onclick="filterByPM('')" class="tab-btn px-3 py-1.5 rounded-lg border border-nwc bg-white text-nwc-primary font-bold whitespace-nowrap active shadow-sm">الكل (122)</button>`;
            
            pms.forEach(pm => {{
                const opt = document.createElement('option');
                opt.value = pm;
                opt.textContent = pm;
                select.appendChild(opt);

                const count = projects.filter(p => p.program_manager === pm).length;
                const encCount = projects.filter(p => p.program_manager === pm).reduce((acc, p) => acc + (p.encroachments_count || 0), 0);
                const encBadge = encCount > 0 ? `<span class="bg-rose-500 text-white text-[10px] px-1 rounded-full mr-1">${{encCount}}</span>` : '';
                
                tabsCont.innerHTML += `<button onclick="filterByPM('${{pm}}')" class="tab-btn px-2.5 py-1.5 rounded-lg border border-nwc bg-white text-nwc-primary font-semibold whitespace-nowrap shadow-sm hover:bg-nwc-phrase">${{pm}} (${{count}}) ${{encBadge}}</button>`;
            }});
        }}

        function filterByPM(pm) {{
            selectedPM = pm;
            document.getElementById('managerFilter').value = pm;
            document.querySelectorAll('.tab-btn').forEach(b => {{
                if ((!pm && b.textContent.includes('الكل')) || (pm && b.textContent.includes(pm))) {{
                    b.classList.add('active');
                }} else {{
                    b.classList.remove('active');
                }}
            }});
            applyFilters();
        }}

        function initMap() {{
            map = L.map('map').setView([24.7136, 46.6753], 10);
            L.tileLayer('https://{{s}}.basemaps.cartocdn.com/rastertiles/voyager/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                attribution: '&copy; CartoDB'
            }}).addTo(map);

            if (waterGeo && waterGeo.features) {{
                L.geoJSON(waterGeo, {{
                    style: {{ color: '#01579B', weight: 2, fillOpacity: 0.12 }}
                }}).addTo(map);
            }}

            if (sewerGeo && sewerGeo.features) {{
                L.geoJSON(sewerGeo, {{
                    style: {{ color: '#097138', weight: 2, fillOpacity: 0.12 }}
                }}).addTo(map);
            }}

            markersCluster = L.markerClusterGroup();
            updateMapMarkers(projects);
            map.addLayer(markersCluster);
        }}

        function updateMapMarkers(projs) {{
            if (!markersCluster) return;
            markersCluster.clearLayers();
            const bounds = [];

            projs.forEach(p => {{
                if (p.encroachments) {{
                    p.encroachments.forEach(e => {{
                        if (e.lat && e.lon) {{
                            const marker = L.circleMarker([e.lat, e.lon], {{
                                radius: 6,
                                fillColor: e.is_critical ? '#DC2626' : '#D97706',
                                color: '#FFFFFF',
                                weight: 1.5,
                                fillOpacity: 0.95
                            }});
                            marker.bindPopup(`
                                <div style="direction: rtl; font-family: Cairo, sans-serif; font-size: 12px;">
                                    <div style="font-weight: 900; color: #153E4B; margin-bottom: 2px;">بلاغ تعدي #${{e.id}}</div>
                                    <div style="color: #168A89; font-weight: 700;">${{p.name}}</div>
                                    <div style="font-size: 11px; margin-top: 4px;"><strong>المقاول:</strong> ${{p.contractor}}</div>
                                    <div style="font-size: 11px;"><strong>الحي:</strong> ${{e.district}} - ${{e.street}}</div>
                                    <div style="font-size: 11px;"><strong>أيام التأخير:</strong> <span style="color: ${{e.is_critical ? '#DC2626' : '#D97706'}}; font-weight: bold;">${{e.age_days}} يوم</span></div>
                                </div>
                            `);
                            markersCluster.addLayer(marker);
                            bounds.push([e.lat, e.lon]);
                        }}
                    }});
                }}
            }});

            if (bounds.length > 0) {{
                map.fitBounds(bounds, {{ padding: [20, 20] }});
            }}
        }}

        function initChart() {{
            const ctx = document.getElementById('pmChart').getContext('2d');
            const dataMap = {{
                'م. أمجد الفالح': 12,
                'م. تركي الأسمري': 6,
                'م. عبدالله الأسود': 5,
                'م. شاكر الحقباني': 4,
                'م. عسكر لسلوم': 2,
                'م. عبدالله العنزي': 1
            }};

            new Chart(ctx, {{
                type: 'doughnut',
                data: {{
                    labels: Object.keys(dataMap),
                    datasets: [{{
                        data: Object.values(dataMap),
                        backgroundColor: ['#153E4B', '#168A89', '#0D9488', '#D97706', '#DC2626', '#64748B'],
                        borderWidth: 2,
                        borderColor: '#FFFFFF'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'bottom',
                            labels: {{ font: {{ family: 'Cairo', size: 10, weight: '700' }} }}
                        }}
                    }}
                }}
            }});
        }}

        function applyFilters() {{
            const q = document.getElementById('searchInput').value.toLowerCase();
            const pm = document.getElementById('managerFilter').value;
            const status = document.getElementById('statusFilter').value;
            const enc = document.getElementById('encFilter').value;

            const filtered = projects.filter(p => {{
                const matchQ = !q || p.name.toLowerCase().includes(q) || p.contractor.toLowerCase().includes(q) || p.location.toLowerCase().includes(q) || p.po.includes(q);
                const matchPM = !pm || p.program_manager === pm;
                const matchStatus = !status || p.status.includes(status);
                
                let matchEnc = true;
                if (enc === 'has_enc') matchEnc = (p.encroachments_count > 0);
                else if (enc === 'delayed_60') matchEnc = (p.encroachments && p.encroachments.some(e => e.age_days > 60));
                else if (enc === 'no_enc') matchEnc = (p.encroachments_count === 0);

                return matchQ && matchPM && matchStatus && matchEnc;
            }});

            renderProjects(filtered);
            updateMapMarkers(filtered);
        }}

        function renderProjects(projs) {{
            const grid = document.getElementById('projectsGrid');
            document.getElementById('results-count').textContent = projs.length;
            
            if (projs.length === 0) {{
                grid.innerHTML = `<div class="col-span-full py-12 text-center text-gray-500 font-bold">لا توجد مشاريع مطابقة لمعايير البحث الحالية.</div>`;
                return;
            }}

            let html = '';
            projs.forEach(p => {{
                const hasEnc = p.encroachments_count > 0;
                const delayedCrit = p.encroachments && p.encroachments.some(e => e.age_days > 60);
                
                let encBadge = '';
                if (delayedCrit) {{
                    encBadge = `<span class="bg-rose-100 text-rose-800 text-[10px] px-2 py-0.5 rounded-md font-black border border-rose-200"><i class="fa-solid fa-triangle-exclamation"></i> ${{p.encroachments_count}} بلاغ (متأخر &gt; شهرين 🚨)</span>`;
                }} else if (hasEnc) {{
                    encBadge = `<span class="bg-amber-100 text-amber-800 text-[10px] px-2 py-0.5 rounded-md font-bold border border-amber-200">${{p.encroachments_count}} بلاغ معلق</span>`;
                }}

                html += `
                <div class="bg-white rounded-xl border border-nwc p-4 card-shadow flex flex-col justify-between relative overflow-hidden">
                    <div class="space-y-2">
                        <div class="flex items-start justify-between gap-2">
                            <span class="text-[10px] bg-nwc-phrase text-nwc-primary px-2 py-0.5 rounded font-bold border border-nwc">#${{p.id}} | ${{p.sector}}</span>
                            <div>${{encBadge}}</div>
                        </div>

                        <h4 class="font-bold text-xs sm:text-sm text-nwc-primary leading-snug">${{p.name}}</h4>
                        
                        <div class="text-[11px] text-gray-600 space-y-1 pt-1 border-t border-nwc">
                            <div class="flex items-center gap-1.5">
                                <i class="fa-solid fa-hard-hat text-nwc-secondary w-3.5"></i>
                                <span><strong>المقاول:</strong> ${{p.contractor}}</span>
                            </div>
                            <div class="flex items-center gap-1.5">
                                <i class="fa-solid fa-user-tie text-nwc-secondary w-3.5"></i>
                                <span><strong>مدير البرنامج:</strong> ${{p.program_manager}}</span>
                            </div>
                            <div class="flex items-center gap-1.5">
                                <i class="fa-solid fa-location-dot text-nwc-secondary w-3.5"></i>
                                <span><strong>الموقع:</strong> ${{p.location}}</span>
                            </div>
                        </div>
                    </div>

                    <div class="mt-3 pt-2 border-t border-nwc flex items-center justify-between text-[11px]">
                        <span class="font-bold text-nwc-secondary">${{p.status}}</span>
                        <span class="text-gray-400 font-mono text-[10px]">PO: ${{p.po}}</span>
                    </div>
                </div>`;
            }});

            grid.innerHTML = html;
        }}

        function exportToExcel() {{
            window.location.href = 'XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx';
        }}

        function exportToKMZ() {{
            window.location.href = 'KMZ/- مشاريع الصرف الصحي بمدينة الرياض.kmz';
        }}
    </script>
</body>
</html>
"""

with open('c:/antigravity files IDE/PROGRAM/index.html', 'w', encoding='utf-8') as f:
    f.write(html_template)

print("SUCCESS: index.html updated with Infographic integration & PDF Print Table for Pending Encroachments only!")
