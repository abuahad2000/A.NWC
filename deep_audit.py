import os, zipfile, xml.etree.ElementTree as ET, pandas as pd
from shapely.geometry import Point, Polygon

kmz_dir = r'c:/antigravity files IDE/PROGRAM/KMZ'
f_enc = r'c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx'
df = pd.read_excel(f_enc, header=3)

# Load ongoing polygons only
ongoing_polys = []
maint_polys = []

for kfile in ['- مشاريع الصرف الصحي بمدينة الرياض.kmz', 'مشاريع المياه بمدينة الرياض .kmz']:
    with zipfile.ZipFile(os.path.join(kmz_dir, kfile), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if fname is not None and fname.text else ''
            is_ongoing = any(k in fname_text for k in ['الجارية', 'الجاري'])
            is_maint = any(k in fname_text for k in ['القائمة', 'قائمة', 'المسلمة'])
            
            for pm in f.findall('.//kml:Placemark', ns):
                pname = pm.find('kml:name', ns)
                pname_text = pname.text.strip() if pname is not None and pname.text else ''
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
                                if is_ongoing:
                                    ongoing_polys.append({'layer': fname_text, 'name': pname_text, 'poly': poly})
                                elif is_maint:
                                    maint_polys.append({'layer': fname_text, 'name': pname_text, 'poly': poly})
                        except: pass

print(f"Ongoing polygons: {len(ongoing_polys)}, Maint polygons: {len(maint_polys)}")

print("\n--- Detailed Audit for each row ---")
for idx, r in df.iterrows():
    eid = r['رقم_بلاغ_التعدي']
    c_name = r['contractor']
    p_name = r['name']
    scope = r['location']
    dist = r['الحي']
    street = r['الشارع']
    lat, lon = r['خط_العرض'], r['خط_الطول']
    pm = r['program_manager_nwc']
    
    in_ong = []
    in_maint = []
    if pd.notnull(lat) and pd.notnull(lon):
        pt = Point(float(lon), float(lat))
        for op in ongoing_polys:
            if op['poly'].contains(pt):
                in_ong.append(op['name'])
        for mp in maint_polys:
            if mp['poly'].contains(pt):
                in_maint.append(mp['name'])
                
    ong_str = ' | '.join(in_ong) if in_ong else 'NONE'
    maint_str = ' | '.join(in_maint) if in_maint else 'NONE'
    
    print(f"#{eid} | Dist:[{dist}] St:[{street}] | PM:[{pm}] | Contractor:[{c_name}]")
    print(f"   Project: [{p_name}] (Scope: [{scope}])")
    print(f"   In Ongoing KMZ: {ong_str}")
    print(f"   In Maintenance KMZ: {maint_str}\n")
