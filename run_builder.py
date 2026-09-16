import os, json, re
import pandas as pd
from datetime import datetime

base_dir = r'c:\antigravity files IDE\PROGRAM\NWC_Projects'
os.makedirs(base_dir, exist_ok=True)

f_proj = r'C:\AI-Workspace\projects\التعديات\excel\بيانات_مقاولين_المشاريع_مع_نطاق_المشروع.xlsx'
df_proj = pd.read_excel(f_proj).dropna(subset=['م'])

f_enc = r'C:\AI-Workspace\projects\التعديات\excel\التعديات.xlsx'
df_enc = pd.read_excel(f_enc)
df_pending = df_enc[df_enc['حالة البلاغ'] != 'تمت المعالجة'].copy()

def clean_str(val):
    if pd.isna(val) or val is None: return '-'
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return '-' if s.lower() == 'nan' or not s else s

def get_pm_key(row):
    sub = str(row.get('البرنامج الفرعي', ''))
    pm = str(row.get('مدير برنامج NWC', ''))
    if 'تركي' in pm or 'الاسمري' in pm: return ('01_تركي_الاسمري.md', 'م. تركي ظافر يحيى الاسمري', 'صرف صحي - جنوب الرياض')
    if 'عسكر' in pm or 'لسلوم' in pm: return ('02_عسكر_لسلوم.md', 'م. عسكر لسلوم', 'صرف صحي - شمال الرياض')
    if 'عبدالله الاسود' in pm or 'عبدالله الأسود' in pm: return ('07_عبدالله_الأسود.md', 'م. عبدالله الأسود العنزي', 'مياه وصرف - مشاريع متفرقة (استثناء)')
    if 'سفر' in pm or 'العتيبي' in pm: return ('04_سفر_العتيبي.md', 'م. سفر العتيبي', 'مياه - شمال الرياض')
    if 'علي الشهري' in pm or ('الشهري' in pm and 'علي' in pm): return ('05_علي_الشهري.md', 'م. علي الشهري', 'مياه - جنوب الرياض')
    if 'أمجد' in pm or 'امجد' in pm or 'الفالح' in pm: return ('06_أمجد_الفالح.md', 'م. أمجد الفالح', 'صرف صحي - غرب الرياض')
    if 'فهد العنزي' in pm or 'عبدالله العنزي' in pm or 'عبدالله علي' in pm: return ('03_عبدالله_العنزي.md', 'م. عبدالله علي العنزي / م. فهد العنزي', 'مياه - غرب الرياض')
    if 'الشمالية' in sub or 'القحطاني' in pm: return ('08_المحافظات_الشمالية.md', 'م. علي القحطاني', 'مياه وصرف - المحافظات الشمالية')
    if 'الجنوبية' in sub or 'الحقباني' in pm: return ('09_المحافظات_الجنوبية.md', 'م. شاكر الحقباني', 'مياه وصرف - المحافظات الجنوبية')
    if 'الغربية' in sub or 'الحارث' in pm: return ('10_المحافظات_الغربية.md', 'م. سعيد الحارث', 'مياه وصرف - المحافظات الغربية')
    return ('غير_محدد.md', 'غير محدد', 'عام')

df_proj['file_tuple'] = df_proj.apply(get_pm_key, axis=1)
df_proj['filename'] = df_proj['file_tuple'].apply(lambda x: x[0])
df_proj['pm_display'] = df_proj['file_tuple'].apply(lambda x: x[1])
df_proj['pm_scope'] = df_proj['file_tuple'].apply(lambda x: x[2])

file_stats = []
for filename, group in df_proj.groupby('filename'):
    pm_name = group['pm_display'].iloc[0]
    pm_scope = group['pm_scope'].iloc[0]
    ongoing_group = group[group['مرحلة المشروع'].astype(str).str.contains('جاري', na=False)]
    
    out_lines = []
    out_lines.append(f'# بطاقات مشاريع {pm_name}')
    out_lines.append(f'> **النطاق التشغيلي:** {pm_scope}')
    out_lines.append(f'> **إجمالي المشاريع:** {len(group)} | **المشاريع الجارية:** {len(ongoing_group)}')
    out_lines.append(f'> **تاريخ التحديث:** {datetime.now().strftime("%Y-%m-%d")}\n')
    out_lines.append('---\n')
    
    mgr_enc_count = 0
    for _, row in group.iterrows():
        p_id = int(row['م']) if pd.notnull(row['م']) else 0
        p_name = clean_str(row.get('إسم المشروع (Ar)'))
        op_num = clean_str(row.get('الرقم التشغيلي'))
        scope = clean_str(row.get('نطاق المشروع'))
        po = clean_str(row.get('PO'))
        contractor = clean_str(row.get('إسم المقاول'))
        consultant = clean_str(row.get('إسم استشاري التنفيذ'))
        status = clean_str(row.get('مرحلة المشروع'))
        sub_prog = clean_str(row.get('البرنامج الفرعي'))
        
        exec_mgr = clean_str(row.get('المدير التنفيذي'))
        exec_phone = clean_str(row.get('جوال المدير التنفيذي'))
        exec_email = clean_str(row.get('ايميل المدير التنفيذي'))
        
        prog_mgr = clean_str(row.get('مدير برنامج NWC'))
        prog_phone = clean_str(row.get('جوال مدير برنامج NWC'))
        prog_email = clean_str(row.get('ايميل مدير برنامج NWC'))
        
        proj_mgr = clean_str(row.get('مدير مشروع NWC'))
        proj_phone = clean_str(row.get('جوال مدير مشروع NWC'))
        proj_email = clean_str(row.get('ايميل  مدير مشروع NWC'))
        
        res_eng = clean_str(row.get('مهندس مقيم - استشاري'))
        res_phone = clean_str(row.get('جوال مهندس مقيم - استشاري'))
        res_email = clean_str(row.get('ايميل مهندس مقيم - استشاري'))
        
        # Match pending encroachments
        matched_enc = pd.DataFrame()
        if contractor and contractor != '-':
            contractor_prefix = contractor.replace('شركة', '').replace('مؤسسة', '').replace('مجموعة', '').strip()[:8]
            matched_enc = df_pending[df_pending['اسم المقاول'].astype(str).str.contains(contractor_prefix, na=False)]
        
        mgr_enc_count += len(matched_enc)
        status_tag = '🟢 [جاري]' if 'جاري' in status else ('🔵 [مسلم ابتدائي]' if 'مسلم' in status else '🔴 [مسحوب]')
        
        out_lines.append(f'## 🏷️ بطاقة مشروع #{p_id}: {p_name}')
        out_lines.append('`yaml')
        out_lines.append(f'id: {p_id}')
        out_lines.append(f'operation_number: "{op_num}"')
        out_lines.append(f'name: "{p_name}"')
        out_lines.append(f'scope: "{scope}"')
        out_lines.append(f'po: "{po}"')
        out_lines.append(f'contractor: "{contractor}"')
        out_lines.append(f'consultant: "{consultant}"')
        out_lines.append(f'status: "{status}"')
        out_lines.append(f'sub_program: "{sub_prog}"')
        out_lines.append(f'executive_manager: "{exec_mgr}"')
        out_lines.append(f'exec_phone: "{exec_phone}"')
        out_lines.append(f'exec_email: "{exec_email}"')
        out_lines.append(f'program_manager: "{prog_mgr}"')
        out_lines.append(f'prog_phone: "{prog_phone}"')
        out_lines.append(f'prog_email: "{prog_email}"')
        out_lines.append(f'project_manager: "{proj_mgr}"')
        out_lines.append(f'proj_phone: "{proj_phone}"')
        out_lines.append(f'proj_email: "{proj_email}"')
        out_lines.append(f'resident_engineer: "{res_eng}"')
        out_lines.append(f'resident_phone: "{res_phone}"')
        out_lines.append(f'resident_email: "{res_email}"')
        out_lines.append('`\n')
        
        out_lines.append('### 📞 بيانات الاتصال والمسؤولين:')
        out_lines.append(f'- **مدير البرنامج:** {prog_mgr} | جوال: {prog_phone} | بريد: {prog_email}')
        out_lines.append(f'- **مدير المشروع NWC:** {proj_mgr} | جوال: {proj_phone} | بريد: {proj_email}')
        out_lines.append(f'- **المهندس المقيم (الاستشاري):** {res_eng} ({consultant}) | جوال: {res_phone} | بريد: {res_email}')
        out_lines.append(f'- **المقاول المنفذ:** {contractor} | **الحالة:** {status_tag}\n')
        
        out_lines.append(f'### ⚠️ التعديات المعلقة المرتبطة بالمشروع والمقاول ({len(matched_enc)} بلاغ معلق):')
        if not matched_enc.empty:
            out_lines.append('| رقم البلاغ | تاريخ البلاغ | الحي | الشارع | حالة البلاغ | وصف التعدي |')
            out_lines.append('|---|---|---|---|---|---|')
            for _, enc_row in matched_enc.head(10).iterrows():
                e_id = enc_row.get('رقم بلاغ التعدي', '-')
                e_dt = str(enc_row.get('تاريخ البلاغ', '-'))[:10]
                e_dis = enc_row.get('الحي', '-')
                e_str = str(enc_row.get('الشارع', '-'))[:25]
                e_st = enc_row.get('حالة البلاغ', '-')
                e_desc = str(enc_row.get('وصف التعدي', '-'))[:45].replace('\n', ' ')
                out_lines.append(f'| {e_id} | {e_dt} | {e_dis} | {e_str} | **{e_st}** | {e_desc} |')
            if len(matched_enc) > 10:
                out_lines.append(f'\n*(يوجد {len(matched_enc)-10} بلاغاً إضافياً في السجل...)*')
        else:
            out_lines.append('*لا توجد بلاغات تعدي معلقة مسجلة حالياً على هذا النطاق/المقاول.*')
        out_lines.append('\n---\n')
    
    with open(os.path.join(base_dir, filename), 'w', encoding='utf-8') as f:
        f.write('\n'.join(out_lines))
        
    file_stats.append({
        'file': filename,
        'pm': pm_name,
        'scope': pm_scope,
        'total': len(group),
        'ongoing': len(ongoing_group),
        'enc': mgr_enc_count
    })

total_ongoing = len(df_proj[df_proj['مرحلة المشروع'].astype(str).str.contains('جاري', na=False)])
readme_lines = [
    '# 📊 قاعدة بيانات مشاريع NWC والتعديات المعلقة (المشاريع الجارية)',
    f'> **إجمالي المشاريع في القاعدة:** {len(df_proj)} مشروعاً  ',
    f'> **المشاريع الجارية تحت التنفيذ:** {total_ongoing} مشروعاً  ',
    f'> **إجمالي التعديات المعلقة النشطة:** {len(df_pending)} بلاغاً  ',
    f'> **تاريخ التحديث:** {datetime.now().strftime("%Y-%m-%d")}  ',
    '> **التصنيف:** حسب مدراء برامج NWC والمشاريع الجارية\n',
    '## 📁 بطاقات مدراء البرامج والمشاريع\n',
    '| # | الملف | مدير البرنامج | النطاق | إجمالي المشاريع | المشاريع الجارية | التعديات المعلقة |',
    '|---|-------|--------------|--------|----------------|-----------------|-----------------|'
]

for idx, s in enumerate(sorted(file_stats, key=lambda x: x['file'])):
    fname = s['file']
    fpm = s['pm']
    fsc = s['scope']
    ftot = s['total']
    fong = s['ongoing']
    fenc = s['enc']
    readme_lines.append(f'| {idx+1} | [{fname}]({fname}) | {fpm} | {fsc} | {ftot} | **{fong}** | {fenc} |')

with open(os.path.join(base_dir, 'README.md'), 'w', encoding='utf-8') as f:
    f.write('\n'.join(readme_lines))

print('UPDATED_CLEANLY')
