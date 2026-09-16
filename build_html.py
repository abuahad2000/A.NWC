import os, zipfile, xml.etree.ElementTree as ET, json
import pandas as pd
from shapely.geometry import Polygon, LineString, mapping

print("Building NWC Web Portal index.html...")

# 1. Load Projects
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])

# 2. Load Encroachments (using header=3 due to title block)
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx"
try:
    df_enc = pd.read_excel(f_enc, header=3)
    if 'id' not in df_enc.columns:
        df_enc = pd.read_excel(f_enc)
except:
    df_enc = pd.read_excel(f_enc)

print("Encroachments columns:", df_enc.columns.tolist()[:5])

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

# Normalize Program Manager
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

# Build Projects JSON
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
    
    # Matching encroachments
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

# 3. Extract KMZ GeoJSON
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

print(f"Total projects processed: {len(projects_list)}")

# Read base HTML template and format
with open("c:/antigravity files IDE/PROGRAM/build_html.py", "r", encoding="utf-8") as f_prev:
    full_code = f_prev.read()

# Extract template from earlier
template_start = full_code.find('html_content = f"""') + len('html_content = f"""')
template_end = full_code.rfind('"""\n\n# Write to c:/antigravity')
html_template = full_code[template_start:template_end]

# Format template
final_html = html_template.replace('{projects_json_str}', projects_json_str)\
                          .replace('{water_geojson_str}', water_geojson_str)\
                          .replace('{sewer_geojson_str}', sewer_geojson_str)

out_html_path = r"c:/antigravity files IDE/PROGRAM/index.html"
with open(out_html_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"SUCCESS: Generated standalone index.html at:\n{out_html_path}")
