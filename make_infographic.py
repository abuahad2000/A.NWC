import json, pandas as pd

df = pd.read_excel('c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx', sheet_name=0, skiprows=3)

base_date = pd.to_datetime('2026-09-16')
df['dt'] = pd.to_datetime(df['تاريخ_البلاغ'], errors='coerce')
df['age_days'] = (base_date - df['dt']).dt.days

enc_data = []
for _, r in df.iterrows():
    enc_data.append({
        'id': int(r['رقم_بلاغ_التعدي']),
        'project': str(r['name']),
        'contractor': str(r['contractor']),
        'manager': str(r['program_manager_nwc']),
        'district': str(r['الحي']),
        'street': str(r['الشارع']),
        'status': str(r['حالة_البلاغ']),
        'target': str(r['جهة_التوجيه_والتنبيه']),
        'date': str(r['تاريخ_البلاغ']),
        'age_days': int(r['age_days']) if pd.notnull(r['age_days']) else 0,
        'is_critical': bool(r['age_days'] > 60) if pd.notnull(r['age_days']) else False,
        'lat': float(r['خط_العرض']) if pd.notnull(r['خط_العرض']) and str(r['خط_العرض']) != '-' else None,
        'lon': float(r['خط_الطول']) if pd.notnull(r['خط_الطول']) and str(r['خط_الطول']) != '-' else None,
    })

enc_json_str = json.dumps(enc_data, ensure_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>إنفوجرافيك التعديات للمشاريع الرأسمالية - شركة المياه الوطنية NWC</title>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@600;700;800;900&family=Sakkal+Majalla:wght@700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <style>
        :root {{
            --primary: #153E4B;       /* العناوين والنصوص الرئيسية: كحلي مائل للأخضر */
            --secondary: #168A89;     /* العناوين الفرعية والشريط الجانبي: تركوازي */
            --bg-page: #FBF8F0;       /* خلفية الصفحات: كريمي فاتح */
            --pill-bg: #E2EFEB;       /* خلفية العبارات: أخضر نعناعي فاتح */
            --card-bg: #EEF3EF;       /* خلفية البطاقات: أخضر رمادي فاتح */
            --border-color: #D8E2DF;  /* الخطوط الفاصلة: رمادي مخضر */
        }}
        
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }}

        body {{
            font-family: 'Cairo', 'Sakkal Majalla', sans-serif;
            background-color: var(--bg-page);
            color: var(--primary);
            line-height: 1.25;
            padding: 10px;
        }}

        .sheet-container {{
            max-width: 1080px;
            margin: 0 auto;
            background: #FFFFFF;
            border-radius: 12px;
            border: 1.5px solid var(--border-color);
            padding: 14px 18px;
            position: relative;
            box-shadow: 0 4px 16px rgba(21, 62, 75, 0.05);
        }}

        .print-btn {{
            position: absolute;
            top: 14px;
            left: 18px;
            background: var(--secondary);
            color: #FFFFFF;
            border: none;
            padding: 6px 14px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 800;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 5px;
            box-shadow: 0 2px 8px rgba(22, 138, 137, 0.25);
            transition: all 0.2s ease;
            z-index: 100;
        }}
        .print-btn:hover {{
            background: var(--primary);
        }}

        /* Header */
        .header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid var(--border-color);
            padding-bottom: 8px;
            margin-bottom: 10px;
        }}
        .header-title-box {{
            text-align: right;
        }}
        .corp-badge {{
            display: inline-block;
            background: var(--pill-bg);
            color: var(--primary);
            padding: 2px 10px;
            border-radius: 20px;
            font-size: 10.5px;
            font-weight: 800;
            margin-bottom: 3px;
            border: 1px solid var(--border-color);
        }}
        .main-title {{
            font-size: 17px;
            font-weight: 900;
            color: var(--primary);
            line-height: 1.2;
        }}
        .sub-title {{
            font-size: 11px;
            color: var(--secondary);
            font-weight: 700;
        }}
        .header-meta {{
            text-align: left;
            font-size: 10px;
            color: var(--primary);
            font-weight: 700;
            background: var(--card-bg);
            padding: 4px 10px;
            border-radius: 6px;
            border: 1px solid var(--border-color);
            line-height: 1.35;
        }}

        /* KPIs Row */
        .kpi-row {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 8px;
            margin-bottom: 10px;
        }}
        .kpi-card {{
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 6px 10px;
            text-align: center;
            position: relative;
        }}
        .kpi-card::before {{
            content: '';
            position: absolute;
            top: 0; right: 0; left: 0;
            height: 3px;
            background: var(--secondary);
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
        }}
        .kpi-num {{
            font-size: 19px;
            font-weight: 900;
            color: var(--primary);
            line-height: 1.1;
        }}
        .kpi-title {{
            font-size: 10px;
            font-weight: 800;
            color: var(--primary);
            margin-top: 1px;
        }}
        .kpi-sub {{
            font-size: 8.5px;
            font-weight: 700;
            color: var(--secondary);
        }}

        /* Middle Grid */
        .mid-grid {{
            display: grid;
            grid-template-columns: 1.1fr 1fr 1.05fr;
            gap: 8px;
            margin-bottom: 10px;
        }}
        .panel-box {{
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 8px 10px;
        }}
        .panel-head {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 4px;
            margin-bottom: 6px;
        }}
        .panel-head h4 {{
            font-size: 11px;
            font-weight: 800;
            color: var(--primary);
            display: flex;
            align-items: center;
            gap: 4px;
        }}
        .badge-tag {{
            background: var(--pill-bg);
            color: var(--primary);
            font-size: 8.5px;
            font-weight: 800;
            padding: 1.5px 5px;
            border-radius: 4px;
            border: 1px solid var(--border-color);
        }}

        #map {{
            height: 130px;
            border-radius: 6px;
            border: 1px solid var(--border-color);
            background: #FFFFFF;
        }}

        /* Top Districts */
        .district-item {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #FFFFFF;
            border: 1px solid var(--border-color);
            border-radius: 5px;
            padding: 4px 6px;
            margin-bottom: 4px;
            font-size: 10px;
            font-weight: 700;
        }}
        .dist-right {{
            display: flex;
            align-items: center;
            gap: 5px;
        }}
        .dist-num {{
            background: var(--primary);
            color: #FFFFFF;
            width: 15px;
            height: 15px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 8.5px;
            font-weight: 900;
        }}
        .dist-count {{
            background: var(--pill-bg);
            color: var(--primary);
            font-weight: 900;
            padding: 1px 6px;
            border-radius: 10px;
            font-size: 9px;
            border: 1px solid var(--border-color);
        }}

        /* Directory Table */
        .table-panel {{
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 8px 10px;
            margin-bottom: 8px;
        }}
        .table-title {{
            font-size: 11px;
            font-weight: 900;
            color: var(--primary);
            margin-bottom: 6px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .dir-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 9.5px;
            background: #FFFFFF;
            border-radius: 6px;
            overflow: hidden;
            border: 1px solid var(--border-color);
        }}
        .dir-table th {{
            background: var(--primary);
            color: #FFFFFF;
            padding: 4.5px 6px;
            text-align: right;
            font-weight: 800;
            font-size: 9.5px;
            border-left: 1px solid var(--secondary);
        }}
        .dir-table th:last-child {{ border-left: none; text-align: center; }}
        .dir-table td {{
            padding: 4px 6px;
            border-bottom: 1px solid var(--border-color);
            color: var(--primary);
            font-weight: 700;
            border-left: 1px solid var(--border-color);
        }}
        .dir-table td:last-child {{ border-left: none; text-align: center; }}
        .dir-table tr:nth-child(even) td {{
            background: var(--card-bg);
        }}
        .phone-tag {{
            direction: ltr;
            display: inline-block;
            background: var(--pill-bg);
            color: var(--primary);
            padding: 1px 4px;
            border-radius: 3px;
            font-family: monospace;
            font-weight: 800;
            font-size: 9.5px;
            border: 1px solid var(--border-color);
        }}
        .count-pill {{
            background: var(--pill-bg);
            color: var(--primary);
            padding: 1.5px 6px;
            border-radius: 8px;
            font-weight: 900;
            font-size: 9.5px;
            display: inline-block;
            border: 1px solid var(--border-color);
        }}

        /* Footer */
        .footer {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-top: 1px solid var(--border-color);
            padding-top: 6px;
            font-size: 9.5px;
            font-weight: 700;
            color: var(--secondary);
        }}
        .footer-author {{
            background: var(--pill-bg);
            color: var(--primary);
            padding: 2px 10px;
            border-radius: 10px;
            font-weight: 800;
            font-size: 10px;
            border: 1px solid var(--border-color);
        }}

        /* Strict Single Page Print Rules (A4) */
        @page {{
            size: A4 portrait;
            margin: 6mm 5mm;
        }}

        @media print {{
            html, body {{
                background: #FFFFFF !important;
                padding: 0 !important;
                margin: 0 !important;
                width: 100% !important;
                height: 100% !important;
            }}
            .sheet-container {{
                box-shadow: none !important;
                border: 1.5px solid var(--border-color) !important;
                padding: 10px 12px !important;
                max-width: 100% !important;
                margin: 0 !important;
                border-radius: 0 !important;
                page-break-inside: avoid !important;
                page-break-after: avoid !important;
            }}
            .print-btn {{
                display: none !important;
            }}
            #map {{
                height: 115px !important;
            }}
            .chart-box {{
                height: 115px !important;
            }}
        }}
    </style>
</head>
<body>

<div class="sheet-container">
    <button class="print-btn" onclick="window.print()">
        <span>🖨️</span> طباعة / حفظ كـ PDF (صفحة واحدة)
    </button>

    <!-- Header -->
    <div class="header">
        <div class="header-title-box">
            <div class="corp-badge">🏢 الشركة الوطنية للمياه (NWC) - القطاع الأوسط</div>
            <h1 class="main-title">📊 التعديات للمشاريع الرأسمالية لشبكات المياه والصرف الصحي</h1>
            <div class="sub-title">لوحة الحوكمة والتحليل المكاني ومؤشر إقفال المنصة (> شهرين) للبلاغات المعلقة المعتمدة</div>
        </div>
        <div class="header-meta">
            <div>تاريخ الاستخراج: 16 سبتمبر 2026</div>
            <div>نطاق الفحص: المشاريع الجارية (1-إلى-1)</div>
            <div>المعلق المعتمد: 30 بلاغاً</div>
        </div>
    </div>

    <!-- KPIs Row -->
    <div class="kpi-row">
        <div class="kpi-card">
            <div class="kpi-num">30</div>
            <div class="kpi-title">إجمالي البلاغات المعلقة</div>
            <div class="kpi-sub">المشاريع الجارية المعتمدة</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-num" style="color: var(--secondary);">16</div>
            <div class="kpi-title">متأخرة &gt; شهرين (53%)</div>
            <div class="kpi-sub">مهددة بإقفال المنصة (60 يوماً)</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-num">8</div>
            <div class="kpi-title">قيد المتابعة (30-60 يوم)</div>
            <div class="kpi-sub">ضمن مهلة المعالجة الميدانية</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-num">6</div>
            <div class="kpi-title">بلاغات حديثة (&lt; 30 يوم)</div>
            <div class="kpi-sub">ضمن الفترة المعتمدة (SLA)</div>
        </div>
    </div>

    <!-- Middle 3-Column Grid -->
    <div class="mid-grid">
        <!-- Col 1: Map -->
        <div class="panel-box">
            <div class="panel-head">
                <h4><span>📍</span> التوزيع الجغرافي للمواقع</h4>
                <span class="badge-tag">إحداثيات حية</span>
            </div>
            <div id="map"></div>
        </div>

        <!-- Col 2: Top Districts in Riyadh -->
        <div class="panel-box">
            <div class="panel-head">
                <h4><span>🏆</span> أحياء الرياض الأكثر تسجيلاً</h4>
                <span class="badge-tag">داخل الرياض فقط</span>
            </div>
            <div class="district-item">
                <div class="dist-right">
                    <span class="dist-num">1</span>
                    <div>
                        <strong>حي العوالي</strong>
                        <div style="font-size: 8.5px; color: var(--secondary);">صرف صحي (م. أمجد الفالح)</div>
                    </div>
                </div>
                <span class="dist-count">7 بلاغات</span>
            </div>
            <div class="district-item">
                <div class="dist-right">
                    <span class="dist-num">2</span>
                    <div>
                        <strong>حي بدر</strong>
                        <div style="font-size: 8.5px; color: var(--secondary);">صرف صحي (م. تركي الأسمري)</div>
                    </div>
                </div>
                <span class="dist-count">5 بلاغات</span>
            </div>
            <div class="district-item">
                <div class="dist-right">
                    <span class="dist-num">3</span>
                    <div>
                        <strong>حي الحائر</strong>
                        <div style="font-size: 8.5px; color: var(--secondary);">صرف صحي (م. تركي الأسمري)</div>
                    </div>
                </div>
                <span class="dist-count">3 بلاغات</span>
            </div>
        </div>

        <!-- Col 3: Delayed Contractors Chart -->
        <div class="panel-box">
            <div class="panel-head">
                <h4><span>⚠️</span> المقاولون المتأخرون &gt; شهرين</h4>
                <span class="badge-tag">حد المنصة 60 يوماً</span>
            </div>
            <div class="chart-box" style="height: 130px; position: relative;">
                <canvas id="contractorChart"></canvas>
            </div>
        </div>
    </div>

    <!-- Bottom: Program Managers & Supervising Consultants Table -->
    <div class="table-panel">
        <div class="table-title">
            <span>👥 سجل متابعة مدراء البرامج والاستشاريين المشرفين وأرقام التواصل</span>
            <span class="badge-tag">إجمالي المعلق: 30 بلاغاً</span>
        </div>
        <table class="dir-table">
            <thead>
                <tr>
                    <th>مدير البرنامج (NWC)</th>
                    <th>جوال مدير البرنامج</th>
                    <th>القطاع / النطاق</th>
                    <th>الاستشاري المشرف (المشروع)</th>
                    <th>المهندس المقيم للاستشاري</th>
                    <th>جوال الاستشاري</th>
                    <th style="text-align: center;">البلاغات المعلقة</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>م. أمجد الفالح</strong></td>
                    <td><span class="phone-tag">0555390254</span></td>
                    <td>غرب الرياض - صرف</td>
                    <td>مكتب الياردة للاستشارات الهندسية</td>
                    <td>م. محمد أنور رشوان</td>
                    <td><span class="phone-tag">0555390254</span></td>
                    <td style="text-align: center;"><span class="count-pill">12 بلاغاً</span></td>
                </tr>
                <tr>
                    <td><strong>م. تركي الأسمري</strong></td>
                    <td><span class="phone-tag">0582877792</span></td>
                    <td>جنوب الرياض - صرف</td>
                    <td>مكتب الياردة للاستشارات الهندسية</td>
                    <td>م. وليد حسين فرج</td>
                    <td><span class="phone-tag">0582877792</span></td>
                    <td style="text-align: center;"><span class="count-pill">6 بلاغات</span></td>
                </tr>
                <tr>
                    <td><strong>م. عبدالله الأسود</strong></td>
                    <td><span class="phone-tag">0509997997</span></td>
                    <td>مشاريع متفرقة (استثناء شامل)</td>
                    <td>مكتب الياردة للاستشارات الهندسية</td>
                    <td>م. محمد السيد تهامي</td>
                    <td><span class="phone-tag">0549266084</span></td>
                    <td style="text-align: center;"><span class="count-pill">5 بلاغات</span></td>
                </tr>
                <tr>
                    <td><strong>م. شاكر الحقباني</strong></td>
                    <td><span class="phone-tag">0500082508</span></td>
                    <td>المحافظات الجنوبية (الخرج)</td>
                    <td>فرع شركة أياسا انخيريا إي اركيتوركورا</td>
                    <td>م. عمرو صماده</td>
                    <td><span class="phone-tag">0543833828</span></td>
                    <td style="text-align: center;"><span class="count-pill">4 بلاغات</span></td>
                </tr>
                <tr>
                    <td><strong>م. عسكر لسلوم</strong></td>
                    <td><span class="phone-tag">0567448829</span></td>
                    <td>شمال الرياض - صرف</td>
                    <td>مكتب الياردة للاستشارات الهندسية</td>
                    <td>م. أشرف عبدالغفار</td>
                    <td><span class="phone-tag">0567448829</span></td>
                    <td style="text-align: center;"><span class="count-pill">بلاغان (2)</span></td>
                </tr>
                <tr>
                    <td><strong>م. عبدالله العنزي</strong></td>
                    <td><span class="phone-tag">0555197474</span></td>
                    <td>غرب الرياض - مياه</td>
                    <td>مكتب الياردة للاستشارات الهندسية</td>
                    <td>م. الصديق علي الحاج</td>
                    <td><span class="phone-tag">0502191621</span></td>
                    <td style="text-align: center;"><span class="count-pill">1 بلاغ</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <!-- Footer -->
    <div class="footer">
        <div>نظام حوكمة وإدارة تعديات البنية التحتية | شركة المياه الوطنية (NWC)</div>
        <div class="footer-author">إعداد: عبدالله الزغيبي - التعديات والحوادث</div>
    </div>
</div>

<script>
    const encroachments = {enc_json_str};

    const map = L.map('map', {{ zoomControl: false }}).setView([24.7136, 46.6753], 10);
    L.tileLayer('https://{{s}}.basemaps.cartocdn.com/rastertiles/voyager/{{z}}/{{x}}/{{y}}{{r}}.png', {{
        attribution: ''
    }}).addTo(map);

    const bounds = [];
    encroachments.forEach(e => {{
        if (e.lat && e.lon) {{
            const marker = L.circleMarker([e.lat, e.lon], {{
                radius: 4.5,
                fillColor: e.is_critical ? '#153E4B' : '#168A89',
                color: '#FFFFFF',
                weight: 1.2,
                fillOpacity: 0.95
            }}).addTo(map);
            bounds.push([e.lat, e.lon]);
        }}
    }});
    if (bounds.length > 0) {{
        map.fitBounds(bounds, {{ padding: [8, 8] }});
    }}

    new Chart(document.getElementById('contractorChart'), {{
        type: 'bar',
        data: {{
            labels: ['الخريف', 'المسار', 'سعد العيسى', 'نظم البيئة', 'الأعمال المدنية', 'الخط الذهبي', 'أخرى'],
            datasets: [{{
                label: 'بلاغات > شهرين',
                data: [4, 3, 3, 3, 2, 2, 3],
                backgroundColor: '#168A89',
                borderRadius: 3
            }}]
        }},
        options: {{
            responsive: true,
            maintainAspectRatio: false,
            scales: {{
                y: {{ beginAtZero: true, ticks: {{ stepSize: 1, font: {{ size: 8 }} }}, grid: {{ color: '#D8E2DF' }} }},
                x: {{ grid: {{ display: false }}, ticks: {{ font: {{ family: 'Cairo', size: 8, weight: '700' }} }} }}
            }},
            plugins: {{ legend: {{ display: false }} }}
        }}
    }});
</script>

</body>
</html>
"""

with open('c:/antigravity files IDE/PROGRAM/infographic_report.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("SUCCESS: 100% Palette Aligned & Strict Single-Page A4 Infographic Report generated!")
