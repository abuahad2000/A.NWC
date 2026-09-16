import os, zipfile, xml.etree.ElementTree as ET, json, re
import pandas as pd
from shapely.geometry import Point, Polygon, LineString, MultiPolygon
from shapely.ops import unary_union
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

def clean_contractor(val):
    if pd.isna(val) or val is None: return ''
    s = str(val).strip()
    for k, v in CONTRACTOR_CLEANUP.items():
        if k in s: return v
    return s

# 2. Parse KMZ Ongoing Geometries
kmz_dir = r"c:/antigravity files IDE/PROGRAM/KMZ"

def load_ongoing_geometries(kmz_path, folder_keywords):
    features = []
    with zipfile.ZipFile(kmz_path, 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if fname is not None else ''
            if any(k in fname_text for k in folder_keywords):
                for pm in f.findall('.//kml:Placemark', ns):
                    pname = pm.find('kml:name', ns)
                    pname_text = pname.text.strip() if pname is not None and pname.text else 'مشروع جاري'
                    desc = pm.find('kml:description', ns)
                    desc_text = desc.text if desc is not None and desc.text else ''
                    
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
                                    features.append({'name': pname_text, 'desc': desc_text, 'geom': poly})
                            except: pass
                        elif len(points) == 2:
                            try:
                                line = LineString(points).buffer(0.0005) # approx 50m
                                features.append({'name': pname_text, 'desc': desc_text, 'geom': line})
                            except: pass
    return features

water_geoms = load_ongoing_geometries(os.path.join(kmz_dir, "مشاريع المياه بمدينة الرياض .kmz"), ['الجاري', 'الجارية'])
sewer_geoms = load_ongoing_geometries(os.path.join(kmz_dir, "- مشاريع الصرف الصحي بمدينة الرياض.kmz"), ['الجارية'])
all_ongoing_geoms = water_geoms + sewer_geoms
print(f"Loaded {len(all_ongoing_geoms)} ongoing geometry features from KMZ.")

# 3. Load Excel Data
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
df_enc = pd.read_excel(f_enc)

# Filter Target Statuses ONLY
TARGET_STATUSES = ['تحت معالجة المقاول', 'بانتظار اعتماد الجهة المتعدية']
df_target = df_enc[df_enc['حالة البلاغ'].isin(TARGET_STATUSES)].copy()
df_target['المقاول_الموحد'] = df_target['اسم المقاول'].apply(clean_contractor)

print(f"Total target status encroachments: {len(df_target)}")
print(df_target['حالة البلاغ'].value_counts())

# 4. Spatial Matching & Exception Matching
matched_records = []
for idx, row in df_target.iterrows():
    c_name = row['المقاول_الموحد']
    lon = row.get('خط الطول')
    lat = row.get('خط العرض')
    status = row.get('حالة البلاغ')
    
    is_aswad_exception = (c_name in ASWAD_CONTRACTORS)
    matched_project = None
    
    # Point check
    if pd.notnull(lon) and pd.notnull(lat):
        try:
            pt = Point(float(lon), float(lat))
            for g in all_ongoing_geoms:
                if g['geom'].contains(pt):
                    matched_project = g['name']
                    break
        except: pass
        
    # If spatial match OR contractor is active ongoing contractor
    if matched_project or is_aswad_exception or c_name:
        action = ""
        action_target = ""
        if status == 'تحت معالجة المقاول':
            action = "تنبيه لمدير البرنامج لإلزام المقاول بالمعالجة الميدانية الفورية وإغلاق البلاغ"
            action_target = "مدير البرنامج"
        elif status == 'بانتظار اعتماد الجهة المتعدية':
            action = "تنبيه لمسؤول التعديات بالمشاريع لطلب إضافة نموذج مراجعة المعالجة من قبل المقاول (إدارة الامتثال - NWC)"
            action_target = "مسؤول التعديات بالمشاريع"
            
        matched_records.append({
            'رقم_البلاغ': row.get('رقم بلاغ التعدي'),
            'تاريخ_البلاغ': str(row.get('تاريخ البلاغ'))[:10],
            'حالة_البلاغ': status,
            'المقاول': c_name,
            'الحي': row.get('الحي'),
            'الشارع': row.get('الشارع'),
            'المشروع_المطابق': matched_project if matched_project else 'نطاق مشاريع المقاول الجارية',
            'الجهة_الموجه_لها_التنبيه': action_target,
            'الإجراء_المطلوب': action,
            'وصف_التعدي': str(row.get('وصف التعدي', ''))[:80].replace('\n', ' ')
        })

df_filtered_output = pd.DataFrame(matched_records)
print(f"\nFinal Filtered Actionable Encroachments: {len(df_filtered_output)}")
print(df_filtered_output['حالة_البلاغ'].value_counts())
print(df_filtered_output['الجهة_الموجه_لها_التنبيه'].value_counts())

# Export clean Excel
out_excel = r"c:/antigravity files IDE/PROGRAM/XLSX/البلاغات_المستهدفة_للمشاريع_الجارية.xlsx"
df_filtered_output.to_excel(out_excel, index=False)
print(f"Exported clean Excel to: {out_excel}")
