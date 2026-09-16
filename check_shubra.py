import pandas as pd
import zipfile, xml.etree.ElementTree as ET, os

print("=== 1. مشاريع شركة الأعمال المدنية في السجل ===")
f_proj = r"C:/AI-Workspace/projects/التعديات/excel/بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx"
df_proj = pd.read_excel(f_proj)

civ_proj = df_proj[df_proj['إسم المقاول'].astype(str).str.contains('المدنية', na=False)]
for idx, row in civ_proj.iterrows():
    print(f"- مشروع #{row.get('م')}: {row.get('إسم المشروع (Ar)')} | النطاق: {row.get('نطاق المشروع')} | الحالة: {row.get('مرحلة المشروع')}")

print("\n=== 2. البحث في طبقات KMZ لمشاريع المياه والصرف الجارية ===")
kmz_dir = r"c:/antigravity files IDE/PROGRAM/KMZ"
for kmz_name in ["مشاريع المياه بمدينة الرياض .kmz", "- مشاريع الصرف الصحي بمدينة الرياض.kmz"]:
    with zipfile.ZipFile(os.path.join(kmz_dir, kmz_name), 'r') as z:
        tree = ET.fromstring(z.read('doc.kml'))
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for f in tree.findall('.//kml:Folder', ns):
            fname = f.find('kml:name', ns)
            fname_t = fname.text if (fname is not None and fname.text) else ''
            pms = f.findall('.//kml:Placemark', ns)
            for pm in pms:
                pname = pm.find('kml:name', ns)
                pname_t = pname.text if (pname is not None and pname.text) else ''
                desc = pm.find('kml:description', ns)
                desc_t = desc.text if (desc is not None and desc.text) else ''
                if 'شبرا' in pname_t or 'شبرا' in desc_t:
                    print(f"[{kmz_name}] مجلد/طبقة: {fname_t} | عنصر: {pname_t}")

print("\n=== 3. بلاغات التعدي لشركة الأعمال المدنية في حي شبرا ===")
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
df_enc = pd.read_excel(f_enc)
shubra_enc = df_enc[(df_enc['الحي'].astype(str).str.contains('شبرا', na=False)) & (df_enc['اسم المقاول'].astype(str).str.contains('المدنية', na=False))]
print(f"عدد بلاغات التعدي المسجلة على شركة الأعمال المدنية في حي شبرا: {len(shubra_enc)}")
if len(shubra_enc) > 0:
    print(shubra_enc['حالة البلاغ'].value_counts())
    print("\nعينة من البلاغات:")
    print(shubra_enc[['رقم بلاغ التعدي', 'تاريخ البلاغ', 'حالة البلاغ', 'الشارع', 'وصف التعدي']].head(5))

print("\n=== 4. ما هو وضع شبكة حي شبرا في KMZ؟ ===")
# Let's check what layer covers Shubra in KMZ
