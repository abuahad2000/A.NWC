import os, zipfile, xml.etree.ElementTree as ET
import pandas as pd

kmz_dir = r'c:/antigravity files IDE/PROGRAM/KMZ'
f_proj = r'C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx'
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])

def clean_txt(s):
    if not s or pd.isna(s): return ''
    return str(s).replace('\xa0', ' ').replace('\u200e', '').replace('\u200f', '').strip()

# 1. Water
water_kmz = os.path.join(kmz_dir, 'مشاريع المياه بمدينة الرياض .kmz')
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

print('=== مطابقة مشاريع المياه الجارية (KMZ vs Excel) ===')
for idx, wf in enumerate(water_features):
    full_search = wf['kmz_name'] + ' ' + wf['kmz_desc']
    matched_row = None
    for _, row in df_proj.iterrows():
        pname = clean_txt(row['إسم المشروع (Ar)'])
        op = clean_txt(row['الرقم التشغيلي'])
        po = clean_txt(row['PO']).replace('.0', '')
        if (op and op in full_search) or (po and len(po) > 4 and po in full_search) or (pname and len(pname) > 10 and (pname in full_search or full_search in pname)):
            matched_row = row
            break
    
    if matched_row is not None:
        mgr = str(matched_row.get('مدير برنامج NWC', '-'))
        pmgr = str(matched_row.get('مدير مشروع NWC', '-'))
        phone = str(matched_row.get('جوال مدير مشروع NWC', '-')).replace('.0', '')
        email = str(matched_row.get('ايميل  مدير مشروع NWC', '-'))
        c_name = str(matched_row.get('إسم المقاول', '-'))
        scope = str(matched_row.get('نطاق المشروع', '-'))
        p_name = matched_row['إسم المشروع (Ar)']
        p_id = int(matched_row['م'])
        print(f'[{idx+1}] KMZ: {wf["kmz_name"][:40]} | الطبقة: {wf["layer"]}')
        print(f'    -> مشروع #{p_id}: {p_name}')
        print(f'    -> مدير البرنامج: {mgr} | مدير المشروع: {pmgr} (جوال: {phone} | بريد: {email})')
        print(f'    -> المقاول: {c_name} | النطاق: {scope}\n')
    else:
        print(f'[{idx+1}] KMZ: {wf["kmz_name"][:40]} -> لم يتطابق مباشرة\n')

# 2. Sewer
sewer_kmz = os.path.join(kmz_dir, '- مشاريع الصرف الصحي بمدينة الرياض.kmz')
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

print('=== مطابقة مشاريع الصرف الصحي الجارية (KMZ vs Excel) ===')
unique_sewer = {}
for sf in sewer_features:
    name = sf['kmz_name']
    if name not in unique_sewer:
        unique_sewer[name] = sf

for idx, (sname, sf) in enumerate(unique_sewer.items()):
    matched_row = None
    for _, row in df_proj.iterrows():
        pname = clean_txt(row['إسم المشروع (Ar)'])
        op = clean_txt(row['الرقم التشغيلي'])
        if (op and op in sname) or (pname and len(pname) > 8 and (pname in sname or sname in pname)):
            matched_row = row
            break
    if matched_row is not None:
        mgr = str(matched_row.get('مدير برنامج NWC', '-'))
        pmgr = str(matched_row.get('مدير مشروع NWC', '-'))
        phone = str(matched_row.get('جوال مدير مشروع NWC', '-')).replace('.0', '')
        c_name = str(matched_row.get('إسم المقاول', '-'))
        scope = str(matched_row.get('نطاق المشروع', '-'))
        p_name = matched_row['إسم المشروع (Ar)']
        p_id = int(matched_row['م'])
        print(f'[{idx+1}] KMZ: {sname[:45]}')
        print(f'    -> مشروع #{p_id}: {p_name}')
        print(f'    -> مدير البرنامج: {mgr} | مدير المشروع: {pmgr} ({phone}) | المقاول: {c_name} | النطاق: {scope}\n')
