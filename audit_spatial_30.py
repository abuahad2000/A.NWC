import os, zipfile, xml.etree.ElementTree as ET, pandas as pd
from shapely.geometry import Point, Polygon

kmz_dir = r'c:/antigravity files IDE/PROGRAM/KMZ'
f_enc = r'c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx'
df = pd.read_excel(f_enc, header=3)

# Load all KMZ placemarks
all_placemarks = []
for kfile in ['- مشاريع الصرف الصحي بمدينة الرياض.kmz', 'مشاريع المياه بمدينة الرياض .kmz']:
    with zipfile.ZipFile(os.path.join(kmz_dir, kfile), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_text = fname.text if fname is not None and fname.text else ''
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
                                all_placemarks.append({
                                    'file': kfile, 'folder': fname_text, 'placemark': pname_text, 'poly': poly
                                })
                        except: pass

print(f"Loaded {len(all_placemarks)} polygons from KMZ.\n")

print("=== AUDIT OF ALL 30 ENCROACHMENTS AGAINST KMZ POLYGONS ===")
for idx, r in df.iterrows():
    eid = r['رقم_بلاغ_التعدي']
    c_name = r['contractor']
    proj_name = r['name']
    scope = r['location']
    district = r['الحي']
    street = r['الشارع']
    lat, lon = r['خط_العرض'], r['خط_الطول']
    pm = r['program_manager_nwc']
    
    containing = []
    if pd.notnull(lat) and pd.notnull(lon):
        pt = Point(float(lon), float(lat))
        for p in all_placemarks:
            if p['poly'].contains(pt):
                containing.append(f"[{p['folder']}] {p['placemark']}")
                
    cont_str = ' | '.join(containing) if containing else 'OUTSIDE ALL POLYGONS'
    print(f"#{eid} | PM: {pm} | District: [{district}] | Street: [{street}]")
    print(f"   Project: [{proj_name}] (Scope: {scope}) | Contractor: [{c_name}]")
    print(f"   KMZ Match: {cont_str}\n")
