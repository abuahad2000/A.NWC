import os, zipfile, xml.etree.ElementTree as ET, json
import pandas as pd
from shapely.geometry import Point, Polygon, LineString, mapping
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime

print("=== تحديث شامل لجميع المخرجات بالربط الصارم الحصري (Strict 1-to-1) ===")

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
    'شركة نظم البيئه': 'شركة نظم البيئة للمقاولات',
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

kmz_dir = r"c:/antigravity files IDE/PROGRAM/KMZ"

# 3. Load KMZ Polygons
def load_ongoing_polygons_and_geojson(kmz_file, keywords):
    matched_polys = []
    geojson_features = []
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
                                    poly_s = poly.simplify(0.0001, preserve_topology=True)
                                    geojson_features.append({
                                        'type': 'Feature',
                                        'properties': {'name': pname_text, 'layer': fname_text},
                                        'geometry': mapping(poly_s)
                                    })
                            except: pass
                        elif len(points) == 2:
                            try:
                                line = LineString(points)
                                line_buf = line.buffer(0.0005)
                                matched_polys.append({
                                    'kmz_name': pname_text,
                                    'layer': fname_text,
                                    'project_row': matched_project_row,
                                    'geom': line_buf
                                })
                                geojson_features.append({
                                    'type': 'Feature',
                                    'properties': {'name': pname_text, 'layer': fname_text},
                                    'geometry': mapping(line)
                                })
                            except: pass
                            
    return matched_polys, {'type': 'FeatureCollection', 'features': geojson_features}

water_polys, water_geojson = load_ongoing_polygons_and_geojson("مشاريع المياه بمدينة الرياض .kmz", ['الجاري', 'الجارية'])
sewer_polys, sewer_geojson = load_ongoing_polygons_and_geojson("- مشاريع الصرف الصحي بمدينة الرياض.kmz", ['الجارية'])
all_ongoing_polys = water_polys + sewer_polys

# 4. Strict Matching on Raw Encroachments File
f_enc = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
df_enc = pd.read_excel(f_enc)

TARGET_STATUSES = ['تحت معالجة المقاول', 'بانتظار اعتماد الجهة المتعدية']
# Exclude invalid non-matching encroachments
EXCLUDE_IDS = {554, 18026, 17821, 16768, 18004, 16951, 17020}

strict_results = []
seen_enc_ids = set()

for idx, row in df_enc.iterrows():
    enc_id = int(row.get('رقم بلاغ التعدي'))
    status = str(row.get('حالة البلاغ', '')).strip()
    
    if status not in TARGET_STATUSES or enc_id in seen_enc_ids or enc_id in EXCLUDE_IDS:
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
            for op in all_ongoing_polys:
                if op['geom'].contains(pt):
                    matched_polygon = op
                    break
        except: pass

    if matched_polygon is not None or is_aswad_contractor:
        assigned_proj_row = None
        
        if is_aswad_contractor:
            aswad_projs = df_ongoing[df_ongoing['المقاول_الموحد'] == c_clean]
            if not aswad_projs.empty:
                assigned_proj_row = aswad_projs.iloc[0]
        elif matched_polygon is not None and matched_polygon.get('project_row') is not None:
            assigned_proj_row = matched_polygon['project_row']
        elif matched_polygon is not None:
            for _, p_row in df_ongoing.iterrows():
                p_name_db = str(p_row.get('إسم المشروع (Ar)', '')).strip()
                if p_name_db and (p_name_db in matched_polygon['kmz_name'] or matched_polygon['kmz_name'] in p_name_db):
                    assigned_proj_row = p_row
                    break

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
print(f"Verified strict unique encroachments: {len(df_strict)}")

# 5. Save Excel File
excel_path = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات_المشاريع_الجارية_المعتمدة_المحدثة.xlsx"
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "البلاغات المعتمدة للمشاريع"
ws.views.sheetView[0].rightToLeft = True

font_title = Font(name="Sakkal Majalla", size=18, bold=True, color="1B365D")
font_subtitle = Font(name="Sakkal Majalla", size=13, italic=True, color="495057")
font_header = Font(name="Sakkal Majalla", size=14, bold=True, color="FFFFFF")
font_data = Font(name="Sakkal Majalla", size=12, bold=False, color="000000")
font_data_bold = Font(name="Sakkal Majalla", size=12, bold=True, color="000000")
font_badge_amber = Font(name="Sakkal Majalla", size=12, bold=True, color="7A4100")
font_badge_blue = Font(name="Sakkal Majalla", size=12, bold=True, color="0C5460")

fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
fill_amber_badge = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
fill_blue_badge = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")

thin_gray = Side(style='thin', color='D9D9D9')
border_cell = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)

ws.merge_cells("A1:K1")
ws["A1"] = "سجل بلاغات التعدي المعتمدة لقطاع المشاريع الجارية (NWC) - مدقق مكانياً"
ws["A1"].font = font_title
ws["A1"].alignment = Alignment(horizontal='right', vertical='center')

ws.merge_cells("A2:K2")
ws["A2"] = f"تاريخ الاستخراج: 16 سبتمبر 2026 | الربط المكاني الحصري 1-إلى-1 بدون أي تكرار | عدد البلاغات المعتمدة: {len(df_strict)}"
ws["A2"].font = font_subtitle
ws["A2"].alignment = Alignment(horizontal='right', vertical='center')

ws.row_dimensions[1].height = 30
ws.row_dimensions[2].height = 22
ws.row_dimensions[3].height = 10

headers = df_strict.columns.tolist()
header_row = 4
for col_num, h_text in enumerate(headers, 1):
    cell = ws.cell(row=header_row, column=col_num)
    cell.value = str(h_text)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = border_cell

ws.row_dimensions[header_row].height = 32

start_row = 5
for r_idx, row in df_strict.iterrows():
    curr_row = start_row + r_idx
    ws.row_dimensions[curr_row].height = 26
    is_even = (r_idx % 2 == 0)
    row_fill = fill_white if is_even else fill_zebra
    
    for c_idx, val in enumerate(row, 1):
        cell = ws.cell(row=curr_row, column=c_idx)
        cell.value = val if pd.notnull(val) else "-"
        cell.font = font_data
        cell.fill = row_fill
        cell.border = border_cell
        col_name = headers[c_idx - 1]
        
        if col_name in ['id', 'operation_number', 'po', 'رقم_بلاغ_التعدي', 'تاريخ_البلاغ', 'status', 'خط_الطول', 'خط_العرض', 'جهة_التوجيه_والتنبيه', 'حالة_البلاغ']:
            cell.alignment = align_center
        else:
            cell.alignment = align_right
            
        if col_name == 'حالة_البلاغ' or col_name == 'جهة_التوجيه_والتنبيه':
            if 'معالجة المقاول' in str(val) or 'مدير البرنامج' in str(val):
                cell.fill = fill_amber_badge
                cell.font = font_badge_amber
            elif 'الجهة المتعدية' in str(val) or 'مسؤول التعديات' in str(val):
                cell.fill = fill_blue_badge
                cell.font = font_badge_blue

        if col_name == 'id' or col_name == 'رقم_بلاغ_التعدي':
            cell.font = font_data_bold

end_row = start_row + len(df_strict) - 1
last_col_letter = get_column_letter(len(headers))
tab_range = f"A{header_row}:{last_col_letter}{end_row}"
table = Table(displayName="StrictEncroachmentsTable", ref=tab_range)
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
ws.add_table(table)

for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col[header_row-1:]:
        val_str = str(cell.value or '')
        length = len(val_str.encode('utf-8')) // 2 + 3
        if length > max_len: max_len = length
    ws.column_dimensions[col_letter].width = max(max_len, 14)

wb.save(excel_path)
print(f"Excel file successfully updated at: {excel_path}")

# 6. Update Web Portal index.html
projects_list = []
for _, r in df_proj.iterrows():
    p_id = int(r['م'])
    c_name = str(r.get('المقاول_الموحد', '')).strip()
    p_name = str(r.get('إسم المشروع (Ar)', '')).strip()
    op_num = str(r.get('الرقم التشغيلي', '')).strip()
    scope = str(r.get('نطاق المشروع', '')).strip()
    po = str(r.get('PO', '')).replace('.0', '').strip()
    consultant = str(r.get('إسم استشاري التنفيذ', '')).strip()
    status = str(r.get('مرحلة المشروع', '')).strip()
    sub_prog = str(r.get('البرنامج الفرعي', '')).strip()
    
    pm = str(r.get('مدير برنامج NWC', ''))
    if 'تركي' in pm or 'الاسمري' in pm: pm_norm = 'م. تركي الاسمري'
    elif 'عسكر' in pm or 'لسلوم' in pm: pm_norm = 'م. عسكر لسلوم'
    elif 'عبدالله الاسود' in pm or 'عبدالله الأسود' in pm: pm_norm = 'م. عبدالله الأسود'
    elif 'سفر' in pm or 'العتيبي' in pm: pm_norm = 'م. سفر العتيبي'
    elif 'علي الشهري' in pm or ('الشهري' in pm and 'علي' in pm): pm_norm = 'م. علي الشهري'
    elif 'أمجد' in pm or 'امجد' in pm or 'الفالح' in pm: pm_norm = 'م. أمجد الفالح'
    elif 'فهد العنزي' in pm or 'عبدالله العنزي' in pm or 'عبدالله علي' in pm: pm_norm = 'م. عبدالله العنزي'
    elif 'الشمالية' in sub_prog or 'القحطاني' in pm: pm_norm = 'م. علي القحطاني (الشمالية)'
    elif 'الجنوبية' in sub_prog or 'الحقباني' in pm: pm_norm = 'م. شاكر الحقباني (الجنوبية)'
    elif 'الغربية' in sub_prog or 'الحارث' in pm: pm_norm = 'م. سعيد الحارث (الغربية)'
    else: pm_norm = 'مدير برنامج عام'

    proj_encs = df_strict[df_strict['id'] == p_id]
    
    enc_items = []
    for _, er in proj_encs.iterrows():
        enc_items.append({
            'id': int(er['رقم_بلاغ_التعدي']),
            'date': str(er['تاريخ_البلاغ']),
            'status': str(er['حالة_البلاغ']),
            'target': str(er['جهة_التوجيه_والتنبيه']),
            'action': str(er['الإجراء_المطلوب']),
            'city': str(er['المدينة']),
            'gov': str(er['المحافظة']),
            'district': str(er['الحي']),
            'street': str(er['الشارع']),
            'lon': float(er['خط_الطول']) if pd.notnull(er['خط_الطول']) and str(er['خط_الطول']) != '-' else None,
            'lat': float(er['خط_العرض']) if pd.notnull(er['خط_العرض']) and str(er['خط_العرض']) != '-' else None,
            'desc': str(er['وصف_التعدي'])
        })

    sector = 'صرف صحي' if ('صرف' in sub_prog or 'صرف' in p_name) else ('مياه' if ('مياه' in sub_prog or 'مياه' in p_name) else 'مياه وصرف')

    projects_list.append({
        'id': p_id,
        'operation_number': op_num,
        'name': p_name,
        'location': scope,
        'po': po,
        'contractor': c_name,
        'consultant': consultant,
        'status': status,
        'sector': sector,
        'sub_program': sub_prog,
        'program_manager': pm_norm,
        'prog_phone': str(r.get('جوال مدير برنامج NWC', '-')).replace('.0', ''),
        'prog_email': str(r.get('ايميل مدير برنامج NWC', '-')),
        'project_manager': str(r.get('مدير مشروع NWC', '-')),
        'proj_phone': str(r.get('جوال مدير مشروع NWC', '-')).replace('.0', ''),
        'proj_email': str(r.get('ايميل  مدير مشروع NWC', '-')),
        'supervisor': str(r.get('مهندس مقيم - استشاري', '-')),
        'resident_phone': str(r.get('جوال مهندس مقيم - استشاري', '-')).replace('.0', ''),
        'executive_manager': str(r.get('المدير التنفيذي', '-')),
        'encroachments_count': len(enc_items),
        'encroachments': enc_items
    })

projects_json_str = json.dumps(projects_list, ensure_ascii=False)
water_geojson_str = json.dumps(water_geojson, ensure_ascii=False)
sewer_geojson_str = json.dumps(sewer_geojson, ensure_ascii=False)

from generate_web_portal import html_template
final_html = html_template.replace('__PROJECTS_PLACEHOLDER__', projects_json_str)\
                          .replace('__WATER_PLACEHOLDER__', water_geojson_str)\
                          .replace('__SEWER_PLACEHOLDER__', sewer_geojson_str)\
                          .replace('>64<', f'>{len(df_strict)}<')

out_html_path = r"c:/antigravity files IDE/PROGRAM/index.html"
with open(out_html_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"Updated index.html saved at: {out_html_path}")
