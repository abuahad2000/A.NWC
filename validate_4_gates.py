import os, zipfile, xml.etree.ElementTree as ET, re
import pandas as pd
from shapely.geometry import Point, Polygon, LineString, MultiPolygon
from shapely.ops import unary_union
from datetime import datetime

print("=== بناء محرك التدقيق المكاني الصارم وفق البوابات الأربع ===")

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
    'شركة ابداع الحياة': 'شركة ابداع الحياة للإستثمار',
    'شركة يالين العربية': 'شركة يالين العربية للمقاولات',
    'شركة بلر العربية': 'شركة بلر العربية للتجارة والمقاولات',
    'شركة المطوع للتجارة': 'شركة المطوع للتجارة والخدمات العامة',
    'شركة بن ماضي': 'شركة بن ماضي المحدودة (شركة شخص واحد)',
    'شركة أنظمة القياس': 'شركة أنظمة القياس والتحكم الصناعي',
    'شركة انشاءات حفر الانفاق': 'شركة انشاءات حفر الانفاق المحدوده',
    'شركة سيسرا': 'شركة سيسرا المحدودة'
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

# 2. Load Projects Master DB (Gate 1 Reference)
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])
df_proj['المقاول_الموحد'] = df_proj['إسم المقاول'].apply(clean_contractor)

# Only ongoing projects
df_proj_ongoing = df_proj[df_proj['مرحلة المشروع'].astype(str).str.contains('جاري', na=False)].copy()
registered_contractors = set(df_proj['المقاول_الموحد'].dropna().unique())
ongoing_contractors = set(df_proj_ongoing['المقاول_الموحد'].dropna().unique())

print(f"إجمالي المشاريع بالسجل: {len(df_proj)} | المشاريع الجارية: {len(df_proj_ongoing)}")
print(f"عدد المقاولين المعتمدين بالسجل: {len(registered_contractors)}")

# 3. Load KMZ Layers (Gate 2: Ongoing vs Maintenance)
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

print(f"طبقات المشاريع الجارية المستخلصة: {len(all_ongoing_layers)} مضلع/مسار")
print(f"طبقات شبكات الصيانة والتشغيل المسلمة: {len(all_maint_layers)} مضلع")

# 4. Load Encroachments File
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
df_enc = pd.read_excel(f_enc)
print(f"\nإجمالي البلاغات في الملف الخام: {len(df_enc)}")

# 5. Apply the 4 Gates
TARGET_STATUSES = ['تحت معالجة المقاول', 'بانتظار اعتماد الجهة المتعدية']

accepted_projects_rows = []
excluded_maintenance_rows = []
excluded_other_status_rows = []
excluded_contractor_rows = []

for idx, row in df_enc.iterrows():
    c_raw = row.get('اسم المقاول')
    c_clean = clean_contractor(c_raw)
    status = str(row.get('حالة البلاغ', '')).strip()
    lon = row.get('خط الطول')
    lat = row.get('خط العرض')
    enc_id = row.get('رقم بلاغ التعدي')
    district = row.get('الحي')
    street = row.get('الشارع')
    
    # Gate 3 Check: Status condition
    if status not in TARGET_STATUSES:
        excluded_other_status_rows.append(row)
        continue
        
    # Gate 1 Check: Contractor in projects registry
    if not c_clean or c_clean not in registered_contractors:
        # Check if contractor keyword exists in registry
        has_match = any(c_clean in rc or rc in c_clean for rc in registered_contractors if len(c_clean) > 4)
        if not has_match:
            excluded_contractor_rows.append(row)
            continue
            
    # Gate 2 & 4 Check: Spatial location (Ongoing Project vs Maintenance / Existing Network)
    matched_ongoing_project = None
    matched_maint_network = None
    
    if pd.notnull(lon) and pd.notnull(lat):
        try:
            pt = Point(float(lon), float(lat))
            # Check ongoing polygons
            for og in all_ongoing_layers:
                if og['geom'].contains(pt):
                    matched_ongoing_project = og
                    break
            # Check maintenance polygons
            if not matched_ongoing_project:
                for mg in all_maint_layers:
                    if mg['geom'].contains(pt):
                        matched_maint_network = mg
                        break
        except: pass

    # Exception rule: Eng. Abdullah Al-Aswad (covers entire Riyadh for his 5 contractors)
    is_aswad_exception = (c_clean in ASWAD_CONTRACTORS)
    
    # Decision:
    # If in maintenance and NOT ongoing project and NOT Aswad exception -> EXCLUDE (Gate 4: تشغيل وصيانة)
    if matched_maint_network and not matched_ongoing_project and not is_aswad_exception:
        excluded_maintenance_rows.append({
            'رقم_البلاغ': enc_id,
            'الحي': district,
            'المقاول': c_clean,
            'حالة_البلاغ': status,
            'السبب': f"يقع داخل شبكة قائمة ومسلمة ({matched_maint_network['name'][:30]}) - تشغيل وصيانة"
        })
        continue

    # Find the corresponding Project Record from df_proj
    matched_proj_row = None
    # 1. By contractor in ongoing projects
    matching_projects = df_proj_ongoing[df_proj_ongoing['المقاول_الموحد'] == c_clean]
    
    # If district matches scope
    for _, p_row in matching_projects.iterrows():
        p_scope = str(p_row.get('نطاق المشروع', ''))
        if str(district) in p_scope or p_scope in str(district):
            matched_proj_row = p_row
            break
            
    if matched_proj_row is None and not matching_projects.empty:
        matched_proj_row = matching_projects.iloc[0]
    elif matched_proj_row is None:
        # Fallback to any project of contractor in master
        any_p = df_proj[df_proj['المقاول_الموحد'] == c_clean]
        if not any_p.empty:
            matched_proj_row = any_p.iloc[0]

    # If it is an active project
    if matched_proj_row is not None or is_aswad_exception or matched_ongoing_project:
        # Action determination
        if status == 'تحت معالجة المقاول':
            action_title = "تنبيه لمدير البرنامج"
            action_desc = "إلزام المقاول بالمعالجة الميدانية العاجلة ورفع إثبات الإغلاق"
        else: # بانتظار اعتماد الجهة المتعدية
            action_title = "تنبيه لمسؤول التعديات بالمشاريع"
            action_desc = "إلزام المقاول بإضافة نموذج مراجعة المعالجة لدى إدارة الامتثال بـ NWC"
            
        p_id = clean_str(matched_proj_row.get('م')) if matched_proj_row is not None else '-'
        p_name = clean_str(matched_proj_row.get('إسم المشروع (Ar)')) if matched_proj_row is not None else (matched_ongoing_project['name'] if matched_ongoing_project else 'مشروع جاري')
        op_num = clean_str(matched_proj_row.get('الرقم التشغيلي')) if matched_proj_row is not None else '-'
        scope = clean_str(matched_proj_row.get('نطاق المشروع')) if matched_proj_row is not None else str(district)
        po = clean_str(matched_proj_row.get('PO')) if matched_proj_row is not None else '-'
        consultant = clean_str(matched_proj_row.get('إسم استشاري التنفيذ')) if matched_proj_row is not None else '-'
        
        prog_mgr = clean_str(matched_proj_row.get('مدير برنامج NWC')) if matched_proj_row is not None else ('م. عبدالله الأسود العنزي' if is_aswad_exception else '-')
        prog_phone = clean_str(matched_proj_row.get('جوال مدير برنامج NWC')) if matched_proj_row is not None else '-'
        prog_email = clean_str(matched_proj_row.get('ايميل مدير برنامج NWC')) if matched_proj_row is not None else '-'
        
        proj_mgr = clean_str(matched_proj_row.get('مدير مشروع NWC')) if matched_proj_row is not None else '-'
        proj_phone = clean_str(matched_proj_row.get('جوال مدير مشروع NWC')) if matched_proj_row is not None else '-'
        proj_email = clean_str(matched_proj_row.get('ايميل  مدير مشروع NWC')) if matched_proj_row is not None else '-'
        
        res_eng = clean_str(matched_proj_row.get('مهندس مقيم - استشاري')) if matched_proj_row is not None else '-'
        res_phone = clean_str(matched_proj_row.get('جوال مهندس مقيم - استشاري')) if matched_proj_row is not None else '-'
        res_email = clean_str(matched_proj_row.get('ايميل مهندس مقيم - استشاري')) if matched_proj_row is not None else '-'
        
        accepted_projects_rows.append({
            'رقم_بلاغ_التعدي': enc_id,
            'تاريخ_البلاغ': str(row.get('تاريخ البلاغ', ''))[:10],
            'حالة_البلاغ': status,
            'المقاول_المنفذ': c_clean,
            'الحي': district,
            'الشارع': street,
            'خط_الطول': lon,
            'خط_العرض': lat,
            'جهة_التوجيه_والتنبيه': action_title,
            'الإجراء_المطلوب': action_desc,
            'رقم_المشروع': p_id,
            'اسم_المشروع_الجاري': p_name,
            'الرقم_التشغيلي': op_num,
            'نطاق_المشروع': scope,
            'رقم_PO': po,
            'الاستشاري': consultant,
            'مدير_برنامج_NWC': prog_mgr,
            'جوال_مدير_البرنامج': prog_phone,
            'ايميل_مدير_البرنامج': prog_email,
            'مدير_مشروع_NWC': proj_mgr,
            'جوال_مدير_المشروع': proj_phone,
            'ايميل_مدير_المشروع': proj_email,
            'المهندس_المقيم_الاستشاري': res_eng,
            'جوال_المهندس_المقيم': res_phone,
            'ايميل_المهندس_المقيم': res_email,
            'الجهة_المالكة': 'الجهة التي تم التعدي عليها وتضررت أصولها',
            'الجهة_المتعدية': 'إدارة الامتثال - شركة المياه الوطنية NWC',
            'اعتماد_المركز': 'مركز مشاريع البنية التحتية بمنطقة الرياض (RIPC)',
            'وصف_التعدي': str(row.get('وصف التعدي', ''))[:100].replace('\n', ' ')
        })
    else:
        excluded_maintenance_rows.append({
            'رقم_البلاغ': enc_id,
            'الحي': district,
            'المقاول': c_clean,
            'حالة_البلاغ': status,
            'السبب': 'خارج نطاق المشاريع الجارية - تشغيل وصيانة'
        })

df_accepted = pd.DataFrame(accepted_projects_rows)
df_maint_ex = pd.DataFrame(excluded_maintenance_rows)

print("\n=== نتائج تطبيق الاشتراطات والبوابات الأربع ===")
print(f"1. بلاغات معتمدة لإدارة المشاريع (مشاريع جارية مستوفية للشروط): {len(df_accepted)}")
print(f"   - تحت معالجة المقاول (تنبيه لمدير البرنامج): {len(df_accepted[df_accepted['حالة_البلاغ'] == 'تحت معالجة المقاول'])}")
print(f"   - بانتظار اعتماد الجهة المتعدية (تنبيه لمسؤول التعديات): {len(df_accepted[df_accepted['حالة_البلاغ'] == 'بانتظار اعتماد الجهة المتعدية'])}")
print(f"2. بلاغات مستبعدة (تابعة للتشغيل والصيانة / شبكات قائمة ومسلمة): {len(df_maint_ex)}")
print(f"3. بلاغات مستبعدة (حالات أخرى مثل تمت المعالجة، معاد، إلخ): {len(excluded_other_status_rows)}")

# Export to clean Excel
out_file = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_كاملة_البيانات.xlsx"
df_accepted.to_excel(out_file, index=False)
print(f"\nتم تصدير ملف إدارة المشاريع المعتمد بنجاح إلى:\n{out_file}")
