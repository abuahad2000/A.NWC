import pandas as pd
import json

fpath = r"c:/antigravity files IDE/PROGRAM/XLSX/بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
xl = pd.ExcelFile(fpath)
print("Sheets:", xl.sheet_names)

df = pd.read_excel(fpath)
print("Total rows:", len(df))
print("Columns:", df.columns.tolist())

print("\n--- توزيع حالات البلاغات بالتفصيل ---")
status_counts = df['حالة البلاغ'].value_counts(dropna=False)
status_pct = df['حالة البلاغ'].value_counts(normalize=True, dropna=False) * 100

summary_df = pd.DataFrame({
    'العدد': status_counts,
    'النسبة المئوية': status_pct.round(2)
})
print(summary_df)

print("\n--- تصنيف البلاغات حسب الجهة المسؤولة عن الإجراء القادم ---")
# Let's inspect unique statuses
for status in df['حالة البلاغ'].unique():
    subset = df[df['حالة البلاغ'] == status]
    print(f"\nحالة: {status} (العدد: {len(subset)})")
    if 'اسم المقاول' in df.columns:
        top_contractors = subset['اسم المقاول'].value_counts().head(3).to_dict()
        print(f"  أبرز المقاولين: {top_contractors}")
    if 'تاريخ البلاغ' in df.columns:
        min_date = str(subset['تاريخ البلاغ'].min())[:10]
        max_date = str(subset['تاريخ البلاغ'].max())[:10]
        print(f"  النطاق الزمني: من {min_date} إلى {max_date}")
