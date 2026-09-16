import os, zipfile, xml.etree.ElementTree as ET, re
import pandas as pd

kmz_dir = rc:/antigravity files IDE/PROGRAM/KMZ
f_proj = rC:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx
df_proj = pd.read_excel(f_proj).dropna(subset=[م])

def clean_txt(s):
    if not s or pd.isna(s): return "
 return str(s).replace('\xa0', ' ').replace('\u200e', '').replace('\u200f', '').strip()

# 1. Parse Ongoing Projects from Water KMZ
water_kmz = os.path.join(kmz_dir, مشاريع المياه بمدينة الرياض .kmz)
water_features = []
with zipfile.ZipFile(water_kmz, 'r') as z:
 tree = ET.fromstring(z.read('doc.kml'))
 ns = {'kml': 'http://www.opengis.net/kml/2.2'}
 for f in tree.findall('.//kml:Folder', ns):
 fname = clean_txt(f.find('kml:name', ns).text if f.find('kml:name', ns) is not None else '')
 if 'الجاري' in fname or 'الجارية' in fname:
 for pm in f.findall('.//kml:Placemark', ns):
 pname = clean_txt(pm.find('kml:name', ns).text if pm.find('kml:name', ns) is not None else '')
 desc = clean_txt(pm.find('kml:description', ns).text if pm.find('kml:description', ns) is not None else '')
 water_features.append({
 'sector': 'مياه',
 'layer': fname,
 'kmz_name': pname,
 'kmz_desc': desc
 })

# 2. Parse Ongoing Projects from Sewer KMZ
sewer_kmz = os.path.join(kmz_dir, - مشاريع الصرف الصحي بمدينة الرياض.kmz)
sewer_features = []
with zipfile.ZipFile(sewer_kmz, 'r') as z:
 tree = ET.fromstring(z.read('doc.kml'))
 ns = {'kml': 'http://www.opengis.net/kml/2.2'}
 for f in tree.findall('.//kml:Folder', ns):
 fname = clean_txt(f.find('kml:name', ns).text if f.find('kml:name', ns) is not None else '')
 if 'الجارية' in fname:
 for pm in f.findall('.//kml:Placemark', ns):
 pname = clean_txt(pm.find('kml:name', ns).text if pm.find('kml:name', ns) is not None else '')
 desc = clean_txt(pm.find('kml:description', ns).text if pm.find('kml:description', ns) is not None else '')
 sewer_features.append({
 'sector': 'صرف صحي',
 'layer': fname,
 'kmz_name': pname,
 'kmz_desc': desc
 })

print(fWater Ongoing Features: {len(water_features)})
print(fSewer Ongoing Features: {len(sewer_features)})

# Match Water Features with Excel
print(\n=== مطابقة مشاريع المياه الجارية (KMZ vs Excel) ===)
for idx, wf in enumerate(water_features):
 full_search = f{wf['kmz_name']} {wf['kmz_desc']}
 matched_row = None
 for _, row in df_proj.iterrows():
 pname = clean_txt(row.get('إسم المشروع (Ar)'))
 op = clean_txt(row.get('الرقم التشغيلي'))
 po = clean_txt(row.get('PO'))
 if (op and op in full_search) or (po and po in full_search) or (pname and (pname in full_search or full_search in pname)):
 matched_row = row
 break
 
 if matched_row is not None:
 print(f[{idx+1}] KMZ: {wf['kmz_name'][:45]})
 print(f    -> Excel ID: {matched_row['م']} | المشروع: {matched_row['إسم المشروع (Ar)'][:35]})
 print(f    -> مدير البرنامج: {matched_row.get('مدير برنامج NWC')} | مدير المشروع: {matched_row.get('مدير مشروع NWC')} ({matched_row.get('جوال مدير مشروع NWC')}))
 print(f    -> المقاول: {matched_row.get('إسم المقاول')} | الاستشاري: {matched_row.get('إسم استشاري التنفيذ')})
 else:
 print(f[{idx+1}] KMZ: {wf['kmz_name'][:45]} -> (لم يتطابق رقم تشغيلي مباشر))

