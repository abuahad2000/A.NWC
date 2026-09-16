import os, zipfile, xml.etree.ElementTree as ET, re
import pandas as pd
from shapely.geometry import Point, Polygon, LineString
from datetime import datetime

print("=== محرك الربط المكاني الحصري الصارم 1-إلى-1 ===")

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

# 2. Load Master Projects
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])
df_proj['المقاول_الموحد'] = df_proj['إسم المقاول'].apply(clean_contractor)
df_ongoing = df_proj[df_proj['مرحلة المشروع'].astype(str).str.contains('جاري', na=False)].copy()

print(f"Projects count: {len(df_proj)}, Ongoing: {len(df_ongoing)}")

# 3. Extract KMZ Ongoing Polygons and Match each Polygon to a Single Project in df_proj
kmz_dir = r"c:/antigravity files IDE/PROGRAM/KMZ"

def load_ongoing_polygons_with_project_match(kmz_file, keywords):
    matched_polys = []
    with zipfile.ZipFile(os.path.join(kmz_dir, kmz_file), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if (fname is not None and fname.text) else ''
            if any(k in fname_text for k in keywords):
                for pm in f.findall('.//kml:Placemark', ns):
                    pname = pm.find('kml:name', ns)
                    pname_text = pname.text.strip() if (pname is not None and pname.text) else ''
                    desc = pm.find('kml:description', ns)
                    desc_text = desc.text.strip() if (desc is not None and desc.text) else ''
                    
                    search_str = f"{pname_text} {desc_text}".replace('\xa0', ' ')
                    
                    # Match with single project in df_proj
                    matched_project_row = None
                    for _, p_row in df_ongoing.iterrows():
                        p_name_db = str(p_row.get('إسم المشروع (Ar)', '')).strip()
                        op_num_db = str(p_row.get('الرقم التشغيلي', '')).strip()
                        po_db = str(p_row.get('PO', '')).replace('.0', '').strip()
                        
                        if (op_num_db and op_num_db in search_str) or (po_db and len(po_db) > 4 and po_db in search_str):
                            matched_project_row = p_row
                            break
                        if p_name_db and len(p_name_db) > 10 and (p_name_db in search_str or search_str in p_name_db):
                            matched_project_row = p_row
                            break
                            
                    # Coordinates
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
                                    matched_polys.append({
                                        'kmz_name': pname_text,
                                        'layer': fname_text,
                                        'project_row': matched_project_row,
                                        'geom': poly
                                    })
                            except: pass
                        elif len(points) == 2:
                            try:
                                line = LineString(points).buffer(0.0005) # 50m
                                matched_polys.append({
                                    'kmz_name': pname_text,
                                    'layer': fname_text,
                                    'project_row': matched_project_row,
                                    'geom': line
                                })
                            except: pass
    return matched_polys

water_ongoing_polys = load_ongoing_polygons_with_project_match("مشاريع المياه بمدينة الرياض .kmz", ['الجاري', 'الجارية'])
sewer_ongoing_polys = load_ongoing_polygons_with_project_match("- مشاريع الصرف الصحي بمدينة الرياض.kmz", ['الجارية'])
all_ongoing_polys = water_ongoing_polys + sewer_ongoing_polys
print(f"Total verified ongoing KMZ polygons: {len(all_ongoing_polys)}")

# Load Maintenance Polygons to strictly exclude them
def load_maintenance_polygons(kmz_file, keywords):
    m_polys = []
    with zipfile.ZipFile(os.path.join(kmz_dir, kmz_file), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if (fname is not None and fname.text) else ''
            if any(k in fname_text for k in keywords):
                for pm in f.findall('.//kml:Placemark', ns):
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
                                    m_polys.append(poly)
                            except: pass
    return m_polys

water_maint = load_maintenance_polygons("مشاريع المياه بمدينة الرياض .kmz", ['القائمة', 'قائمة'])
sewer_maint = load_maintenance_polygons("- مشاريع الصرف الصحي بمدينة الرياض.kmz", ['القائمة', 'قائمة', 'المسلمة'])
all_maint_polys = water_maint + sewer_maint
print(f"Total maintenance polygons for exclusion: {len(all_maint_polys)}")

# 4. Strict Matching on Raw Encroachments File
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
df_enc = pd.read_excel(f_enc)

TARGET_STATUSES = ['تحت معالجة المقاول', 'بانتظار اعتماد الجهة المتعدية']

strict_results = []
seen_enc_ids = set()

for idx, row in df_enc.iterrows():
    enc_id = row.get('رقم بلاغ التعدي')
    status = str(row.get('حالة البلاغ', '')).strip()
    
    # Gate 3: Status
    if status not in TARGET_STATUSES:
        continue
        
    # Prevent any duplicate encroachment ID
    if enc_id in seen_enc_ids:
        continue
        
    c_raw = row.get('اسم المقاول')
    c_clean = clean_contractor(c_raw)
    lon = row.get('خط الطول')
    lat = row.get('خط العرض')
    district = clean_str(row.get('الحي'))
    street = clean_str(row.get('الشارع'))
    city = clean_str(row.get('المدينة'))
    
    is_aswad_contractor = (c_clean in ASWAD_CONTRACTORS)
    matched_polygon = None
    
    if pd.notnull(lon) and pd.notnull(lat):
        try:
            pt = Point(float(lon), float(lat))
            
            # Check ongoing polygons
            for op in all_ongoing_polys:
                if op['geom'].contains(pt):
                    matched_polygon = op
                    break
                    
            # Check maintenance exclusion (if not ongoing and not Aswad)
            if not matched_polygon and not is_aswad_contractor:
                in_maint = any(mp.contains(pt) for mp in all_maint_polys)
                if in_maint:
                    continue # Exclude Maintenance (Gate 4)
        except: pass

    # If inside ongoing polygon OR belongs to Aswad exception
    if matched_polygon is not None or is_aswad_contractor:
        assigned_proj_row = None
        
        if matched_polygon is not None and matched_polygon.get('project_row') is not None:
            assigned_proj_row = matched_polygon['project_row']
        elif is_aswad_contractor:
            # Find Aswad project
            aswad_projs = df_ongoing[df_ongoing['المقاول_الموحد'] == c_clean]
            if not aswad_projs.empty:
                assigned_proj_row = aswad_projs.iloc[0]
            else:
                aswad_projs_all = df_proj[df_proj['المقاول_الموحد'] == c_clean]
                if not aswad_projs_all.empty:
                    assigned_proj_row = aswad_projs_all.iloc[0]
        elif matched_polygon is not None:
            # Fallback polygon match by name
            for _, p_row in df_ongoing.iterrows():
                p_name_db = str(p_row.get('إسم المشروع (Ar)', '')).strip()
                if p_name_db and (p_name_db in matched_polygon['kmz_name'] or matched_polygon['kmz_name'] in p_name_db):
                    assigned_proj_row = p_row
                    break
                    
        # If still no project row found, match contractor's ongoing project in that district
        if assigned_proj_row is None and c_clean:
            cand = df_ongoing[df_ongoing['المقاول_الموحد'] == c_clean]
            for _, cp in cand.iterrows():
                if district != '-' and district in str(cp.get('نطاق المشروع', '')):
                    assigned_proj_row = cp
                    break
            if assigned_proj_row is None and not cand.empty:
                assigned_proj_row = cand.iloc[0]

        if assigned_proj_row is not None:
            seen_enc_ids.add(enc_id)
            
            p_id = int(assigned_proj_row['م'])
            p_name = clean_str(assigned_proj_row.get('إسم المشروع (Ar)'))
            op_num = clean_str(assigned_proj_row.get('الرقم التشغيلي'))
            scope = clean_str(assigned_proj_row.get('نطاق المشروع'))
            po = clean_str(assigned_proj_row.get('PO'))
            contractor_final = clean_str(assigned_proj_row.get('إسم المقاول'))
            consultant = clean_str(assigned_proj_row.get('إسم استشاري التنفيذ'))
            p_status = clean_str(assigned_proj_row.get('مرحلة المشروع'))
            sub_prog = clean_str(assigned_proj_row.get('البرنامج الفرعي'))
            exec_mgr = clean_str(assigned_proj_row.get('المدير التنفيذي'))
            prog_mgr = clean_str(assigned_proj_row.get('مدير برنامج NWC'))
            proj_mgr = clean_str(assigned_proj_row.get('مدير مشروع NWC'))
            supervisor = clean_str(assigned_proj_row.get('مهندس مقيم - استشاري'))
            
            # Action
            if status == 'تحت معالجة المقاول':
                action_title = "تنبيه لمدير البرنامج"
                action_desc = "إلزام المقاول بالمعالجة الميدانية الفورية وإغلاق البلاغ بنظام المركز"
            else:
                action_title = "تنبيه لمسؤول التعديات بالمشاريع"
                action_desc = "إلزام المقاول بإضافة نموذج مراجعة المعالجة لدى إدارة الامتثال بـ NWC"
                
            resolved_city = "مدينة الرياض" if ("الرياض" in city or city == '-') else city
            resolved_gov = "مدينة الرياض"
            if "شقراء" in resolved_city or "مرات" in scope or "شقراء" in scope:
                resolved_gov = "محافظة شقراء"
            elif "الدرعية" in resolved_city or "العيينة" in scope or "الدرعية" in scope:
                resolved_gov = "محافظة الدرعية"
            elif "حوطة" in scope:
                resolved_gov = "محافظة حوطة بني تميم"

            strict_results.append({
                'id': p_id,
                'operation_number': op_num,
                'name': p_name,
                'location': scope,
                'po': po,
                'contractor': contractor_final,
                'consultant': consultant,
                'status': p_status,
                'sub_program': sub_prog,
                'executive_manager': exec_mgr,
                'program_manager_nwc': prog_mgr,
                'project_manager_nwc': proj_mgr,
                'supervisor': supervisor,
                'المدينة': resolved_city,
                'المحافظة': resolved_gov,
                'الحي': district,
                'الشارع': street,
                'خط_الطول': lon,
                'خط_العرض': lat,
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

df_strict = pd.DataFrame(strict_results)
print(f"\nFinal Strict Unique Encroachments: {len(df_strict)}")
print(f"Unique Encroachment IDs: {df_strict['رقم_بلاغ_التعدي'].nunique()}")
print("Distribution across Program Managers:")
print(df_strict['program_manager_nwc'].value_counts())
