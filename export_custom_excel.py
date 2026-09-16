import os, zipfile, xml.etree.ElementTree as ET, re
import pandas as pd
from shapely.geometry import Point, Polygon, LineString
from datetime import datetime

# 1. Clean Dictionary
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

ASWAD_CONTRACTORS = [
    'مجموعة سعد علي العيسى للمقاولات',
    'مؤسسة العرين للمقاولات',
    'شركة الاومير للتجارة والمقاولات',
    'مؤسسة ثليل للمقاولات',
    'شركة صلت للمقاولات'
]

# Governorate Mapping helper
GOVERNORATES = {
    'الدرعية': 'محافظة الدرعية',
    'العيينة': 'محافظة الدرعية',
    'الجبيلة': 'محافظة الدرعية',
    'الخرج': 'محافظة الخرج',
    'الدلم': 'محافظة الدلم',
    'المجمعة': 'محافظة المجمعة',
    'الزلفي': 'محافظة الزلفي',
    'شقراء': 'محافظة شقراء',
    'مرات': 'محافظة مرات',
    'حوطة بني تميم': 'محافظة حوطة بني تميم',
    'الحريق': 'محافظة الحريق',
    'الأفلاج': 'محافظة الأفلاج',
    'السليل': 'محافظة السليل',
    'وادي الدواسر': 'محافظة وادي الدواسر',
    'القويعية': 'محافظة القويعية',
    'الدوادمي': 'محافظة الدوادمي',
    'عفيف': 'محافظة عفيف',
    'ضرما': 'محافظة ضرما',
    'المزاحمية': 'محافظة المزاحمية',
    'رماح': 'محافظة رماح',
    'ثادق': 'محافظة ثادق',
    'حريملاء': 'محافظة حريملاء',
    'الغاط': 'محافظة الغاط'
}

def extract_governorate(city, district, project_scope):
    text = f"{city} {district} {project_scope}"
    for k, v in GOVERNORATES.items():
        if k in text:
            return v
    return "مدينة الرياض"

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

# 2. Load Projects Master
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])
df_proj['المقاول_الموحد'] = df_proj['إسم المقاول'].apply(clean_contractor)
df_proj_ongoing = df_proj[df_proj['مرحلة المشروع'].astype(str).str.contains('جاري', na=False)].copy()
registered_contractors = set(df_proj['المقاول_الموحد'].dropna().unique())

# 3. Load KMZ Layers
kmz_dir = r"c:/antigravity files IDE/PROGRAM/KMZ"

def extract_kmz_polygons(kmz_file, ongoing_keywords, maintenance_keywords):
    ongoing_geoms = []
    maint_geoms = []
    with zipfile.ZipFile(os.path.join(kmz_dir, kmz_file), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if (fname is not None and fname.text) else ''
            is_ongoing = any(k in fname_text for k in ongoing_keywords)
            is_maint = any(k in fname_text for k in maintenance_keywords)
            for pm in f.findall('.//kml:Placemark', ns):
                pname = pm.find('kml:name', ns)
                pname_text = pname.text.strip() if (pname is not None and pname.text) else fname_text
                desc = pm.find('kml:description', ns)
                desc_text = desc.text if (desc is not None and desc.text) else ''
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
                                item = {'name': pname_text, 'layer': fname_text, 'desc': desc_text, 'geom': poly}
                                if is_ongoing: ongoing_geoms.append(item)
                                elif is_maint: maint_geoms.append(item)
                        except: pass
                    elif len(points) == 2:
                        try:
                            line = LineString(points).buffer(0.0005) # ~50m
                            item = {'name': pname_text, 'layer': fname_text, 'desc': desc_text, 'geom': line}
                            if is_ongoing: ongoing_geoms.append(item)
                            elif is_maint: maint_geoms.append(item)
                        except: pass
    return ongoing_geoms, maint_geoms

water_ongoing, water_maint = extract_kmz_polygons("مشاريع المياه بمدينة الرياض .kmz", ['الجاري', 'الجارية'], ['القائمة', 'قائمة'])
sewer_ongoing, sewer_maint = extract_kmz_polygons("- مشاريع الصرف الصحي بمدينة الرياض.kmz", ['الجارية'], ['القائمة', 'قائمة', 'المسلمة'])
all_ongoing_layers = water_ongoing + sewer_ongoing
all_maint_layers = water_maint + sewer_maint

# 4. Load Encroachments File
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
df_enc = pd.read_excel(f_enc)

TARGET_STATUSES = ['تحت معالجة المقاول', 'بانتظار اعتماد الجهة المتعدية']
final_export_rows = []

for idx, row in df_enc.iterrows():
    c_raw = row.get('اسم المقاول')
    c_clean = clean_contractor(c_raw)
    status = str(row.get('حالة البلاغ', '')).strip()
    lon = row.get('خط الطول')
    lat = row.get('خط العرض')
    enc_id = row.get('رقم بلاغ التعدي')
    city = clean_str(row.get('المدينة'))
    district = clean_str(row.get('الحي'))
    street = clean_str(row.get('الشارع'))
    
    # Gate 3: Target Statuses
    if status not in TARGET_STATUSES:
        continue
        
    # Gate 1: Registered Contractor
    if not c_clean or c_clean not in registered_contractors:
        has_match = any(c_clean in rc or rc in c_clean for rc in registered_contractors if len(c_clean) > 4)
        if not has_match:
            continue
            
    # Gate 2 & 4: Spatial Check
    matched_ongoing_project = None
    matched_maint_network = None
    
    if pd.notnull(lon) and pd.notnull(lat):
        try:
            pt = Point(float(lon), float(lat))
            for og in all_ongoing_layers:
                if og['geom'].contains(pt):
                    matched_ongoing_project = og
                    break
            if not matched_ongoing_project:
                for mg in all_maint_layers:
                    if mg['geom'].contains(pt):
                        matched_maint_network = mg
                        break
        except: pass

    is_aswad_exception = (c_clean in ASWAD_CONTRACTORS)
    
    # Exclude maintenance / existing networks
    if matched_maint_network and not matched_ongoing_project and not is_aswad_exception:
        continue

    # Match project record
    matched_proj_row = None
    matching_projects = df_proj_ongoing[df_proj_ongoing['المقاول_الموحد'] == c_clean]
    for _, p_row in matching_projects.iterrows():
        p_scope = str(p_row.get('نطاق المشروع', ''))
        if str(district) in p_scope or p_scope in str(district):
            matched_proj_row = p_row
            break
            
    if matched_proj_row is None and not matching_projects.empty:
        matched_proj_row = matching_projects.iloc[0]
    elif matched_proj_row is None:
        any_p = df_proj[df_proj['المقاول_الموحد'] == c_clean]
        if not any_p.empty:
            matched_proj_row = any_p.iloc[0]

    if matched_proj_row is not None or is_aswad_exception or matched_ongoing_project:
        # Schema fields requested by user
        p_id = clean_str(matched_proj_row.get('م')) if matched_proj_row is not None else '-'
        op_num = clean_str(matched_proj_row.get('الرقم التشغيلي')) if matched_proj_row is not None else '-'
        p_name = clean_str(matched_proj_row.get('إسم المشروع (Ar)')) if matched_proj_row is not None else (matched_ongoing_project['name'] if matched_ongoing_project else 'مشروع جاري')
        p_location = clean_str(matched_proj_row.get('نطاق المشروع')) if matched_proj_row is not None else district
        po = clean_str(matched_proj_row.get('PO')) if matched_proj_row is not None else '-'
        contractor = clean_str(matched_proj_row.get('إسم المقاول')) if matched_proj_row is not None else c_clean
        consultant = clean_str(matched_proj_row.get('إسم استشاري التنفيذ')) if matched_proj_row is not None else '-'
        p_status = clean_str(matched_proj_row.get('مرحلة المشروع')) if matched_proj_row is not None else 'جاري'
        sub_prog = clean_str(matched_proj_row.get('البرنامج الفرعي')) if matched_proj_row is not None else '-'
        exec_mgr = clean_str(matched_proj_row.get('المدير التنفيذي')) if matched_proj_row is not None else '-'
        prog_mgr = clean_str(matched_proj_row.get('مدير برنامج NWC')) if matched_proj_row is not None else ('م. عبدالله الأسود العنزي' if is_aswad_exception else '-')
        proj_mgr = clean_str(matched_proj_row.get('مدير مشروع NWC')) if matched_proj_row is not None else '-'
        supervisor = clean_str(matched_proj_row.get('مهندس مقيم - استشاري')) if matched_proj_row is not None else '-'
        
        # Spatial location resolution
        resolved_city = "مدينة الرياض" if ("الرياض" in city or city == '-') else city
        resolved_gov = extract_governorate(resolved_city, district, p_location)
        
        # Action guidance
        if status == 'تحت معالجة المقاول':
            action_title = "تنبيه لمدير البرنامج"
            action_desc = "إلزام المقاول بالمعالجة الميدانية الفورية وإغلاق البلاغ بنظام المركز"
        else:
            action_title = "تنبيه لمسؤول التعديات بالمشاريع"
            action_desc = "إلزام المقاول بإضافة نموذج مراجعة المعالجة لدى إدارة الامتثال بـ NWC"

        final_export_rows.append({
            # Project fields requested by user
            'id': p_id,
            'operation_number': op_num,
            'name': p_name,
            'location': p_location,
            'po': po,
            'contractor': contractor,
            'consultant': consultant,
            'status': p_status,
            'sub_program': sub_prog,
            'executive_manager': exec_mgr,
            'program_manager_nwc': prog_mgr,
            'project_manager_nwc': proj_mgr,
            'supervisor': supervisor,
            
            # Location fields extracted from coordinates & spatial matching
            'المدينة': resolved_city,
            'المحافظة': resolved_gov,
            'الحي': district,
            'الشارع': street,
            'خط_الطول': lon,
            'خط_العرض': lat,
            
            # Encroachment details & actions
            'رقم_بلاغ_التعدي': enc_id,
            'تاريخ_البلاغ': str(row.get('تاريخ البلاغ', ''))[:10],
            'حالة_البلاغ': status,
            'جهة_التوجيه_والتنبيه': action_title,
            'الإجراء_المطلوب': action_desc,
            'الجهة_المالكة': 'الجهة التي تم التعدي عليها وتضررت أصولها',
            'الجهة_المتعدية': 'إدارة الامتثال - شركة المياه الوطنية NWC',
            'اعتماد_المركز': 'مركز مشاريع البنية التحتية بمنطقة الرياض (RIPC)',
            'وصف_التعدي': str(row.get('وصف التعدي', ''))[:120].replace('\n', ' ')
        })

df_final_export = pd.DataFrame(final_export_rows)
out_excel = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx"
df_final_export.to_excel(out_excel, index=False)

print(f"Exported {len(df_final_export)} rows successfully to: {out_excel}")
print("\nSample Columns:")
print(df_final_export.columns.tolist())
