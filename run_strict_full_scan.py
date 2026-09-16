import os, zipfile, xml.etree.ElementTree as ET, json
import pandas as pd
from shapely.geometry import Point, Polygon, LineString

kmz_dir = r"c:/antigravity files IDE/PROGRAM/KMZ"
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])

from update_all_strict import clean_contractor, ASWAD_CONTRACTORS, TARGET_STATUSES

df_proj['المقاول_الموحد'] = df_proj['إسم المقاول'].apply(clean_contractor)
df_ongoing = df_proj[df_proj['مرحلة المشروع'].astype(str).str.contains('جاري', na=False)].copy()

# 1. Load Ongoing Polygons with Exact Project Mapping
ongoing_polys = []
maint_polys = []

for kfile in ['- مشاريع الصرف الصحي بمدينة الرياض.kmz', 'مشاريع المياه بمدينة الرياض .kmz']:
    with zipfile.ZipFile(os.path.join(kmz_dir, kfile), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if fname is not None and fname.text else ''
            is_ong = any(k in fname_text for k in ['الجارية', 'الجاري'])
            is_maint = any(k in fname_text for k in ['القائمة', 'قائمة', 'المسلمة'])
            
            for pm in f.findall('.//kml:Placemark', ns):
                pname = pm.find('kml:name', ns)
                pname_text = pname.text.strip() if pname is not None and pname.text else ''
                desc = pm.find('kml:description', ns)
                desc_text = desc.text.strip() if desc is not None and desc.text else ''
                search_str = f"{pname_text} {desc_text}".replace('\xa0', ' ')
                
                # Match project
                matched_proj = None
                for _, prow in df_ongoing.iterrows():
                    pname_db = str(prow['إسم المشروع (Ar)'])
                    op_num = str(prow.get('الرقم التشغيلي', '')).strip()
                    po = str(prow.get('PO', '')).replace('.0', '').strip()
                    if (op_num and op_num in search_str) or (po and len(po)>4 and po in search_str):
                        matched_proj = prow
                        break
                    if len(pname_db) > 10 and (pname_db in search_str or search_str in pname_db):
                        matched_proj = prow
                        break

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
                                if is_ong:
                                    ongoing_polys.append({'layer': fname_text, 'name': pname_text, 'proj': matched_proj, 'poly': poly})
                                elif is_maint:
                                    maint_polys.append(poly)
                        except: pass
                    elif len(points) == 2 and is_ong:
                        try:
                            line = LineString(points)
                            line_buf = line.buffer(0.0005)
                            ongoing_polys.append({'layer': fname_text, 'name': pname_text, 'proj': matched_proj, 'poly': line_buf})
                        except: pass

print(f"Ongoing polys: {len(ongoing_polys)}, Maint polys: {len(maint_polys)}")

# 2. Audit all 4,078 raw rows
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
df_raw = pd.read_excel(f_enc)

verified_strict = []
seen = set()

for idx, row in df_raw.iterrows():
    eid = int(row.get('رقم بلاغ التعدي'))
    status = str(row.get('حالة البلاغ', '')).strip()
    if status not in TARGET_STATUSES or eid in seen:
        continue
        
    c_raw = row.get('اسم المقاول')
    c_clean = clean_contractor(c_raw)
    lon = row.get('خط الطول')
    lat = row.get('خط العرض')
    district = str(row.get('الحي', '')).strip()
    street = str(row.get('الشارع', '')).strip()
    
    # 1. Al-Aswad Exception (5 contractors city-wide)
    if c_clean in ASWAD_CONTRACTORS:
        aswad_projs = df_ongoing[df_ongoing['المقاول_الموحد'] == c_clean]
        if not aswad_projs.empty:
            p_row = aswad_projs.iloc[0]
            verified_strict.append({
                'id': eid, 'status': status, 'date': str(row.get('تاريخ البلاغ'))[:10],
                'contractor': c_clean, 'proj_id': int(p_row['م']), 'proj_name': p_row['إسم المشروع (Ar)'],
                'pm': p_row['مدير برنامج NWC'], 'district': district, 'street': street,
                'lat': lat, 'lon': lon, 'scope': p_row['نطاق المشروع'], 'type': 'استثناء م. عبدالله الأسود (شامل)'
            })
            seen.add(eid)
            continue
            
    # 2. Strict Spatial Point-in-Polygon Check
    if pd.notnull(lon) and pd.notnull(lat):
        try:
            pt = Point(float(lon), float(lat))
            
            # Must fall inside an ongoing polygon
            matched_poly = None
            for op in ongoing_polys:
                if op['poly'].contains(pt):
                    matched_poly = op
                    break
                    
            if matched_poly is not None:
                p_row = matched_poly.get('proj')
                if p_row is not None:
                    # Check contractor match: either contractor matches project OR contractor matches raw
                    if p_row['المقاول_الموحد'] == c_clean:
                        verified_strict.append({
                            'id': eid, 'status': status, 'date': str(row.get('تاريخ البلاغ'))[:10],
                            'contractor': c_clean, 'proj_id': int(p_row['م']), 'proj_name': p_row['إسم المشروع (Ar)'],
                            'pm': p_row['مدير برنامج NWC'], 'district': district, 'street': street,
                            'lat': lat, 'lon': lon, 'scope': p_row['نطاق المشروع'], 'type': 'مطابقة مكانية ومقاول مشروع جاري'
                        })
                        seen.add(eid)
        except: pass

df_res = pd.DataFrame(verified_strict)
print(f"\n=== TOTAL STRICT VERIFIED ENCROACHMENTS: {len(df_res)} ===")
print("\nBy Program Manager:")
print(df_res['pm'].value_counts())
print("\nBy Project:")
print(df_res['proj_name'].value_counts())
print("\nAll Verified Encroachments:")
print(df_res[['id', 'pm', 'proj_name', 'contractor', 'district', 'street', 'status', 'type']].to_string())
