#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================
🔹 NWC Encroachments Governance Analyzer
=============================================================
محلل بلاغات التعديات على مشاريع شركة المياه الوطنية
يقوم بتحليل ملف الإكسيل الخام وتطبيق قواعد الحوكمة الآلية
لإنتاج ملف JSON جاهز للبوابة التفاعلية.

📅 آخر تحديث: 17 سبتمبر 2026
👤 إعداد: إدارة التعديات والحوادث - القطاع الأوسط
=============================================================
"""

import json
import re
import os
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("❌ يجب تثبيت مكتبة pandas:")
    print("   pip install pandas openpyxl")
    exit(1)


# ============================================================
# 🔧 الإعدادات العامة
# ============================================================
BASE_DIR = Path(__file__).parent
INPUT_FILE = BASE_DIR / "XLSX" / "بلاغات تعدي مقاولي شركة المياه الوطنية 12 سبتمبر.xlsx"
OUTPUT_FILE = BASE_DIR / "data" / "encroachments_verified.json"
CATALOG_FILE = BASE_DIR / "data" / "projects_catalog.json"
BASE_DATE = datetime(2026, 9, 17)  # تاريخ الاستخراج المرجعي


# ============================================================
# 🌐 دوال تطبيع النصوص العربية
# ============================================================
def normalize_arabic(text: str) -> str:
    """تطبيع النص العربي لتوحيد الأشكال المختلفة"""
    if not text or pd.isna(text):
        return ""
    text = str(text).strip()
    text = re.sub(r'[أإآ]', 'ا', text)
    text = text.replace('ة', 'ه')
    text = text.replace('ى', 'ي')
    text = re.sub(r'\s+', ' ', text)
    return text.lower()


def safe_float(value, default=None):
    """تحويل القيمة إلى رقم عشري بأمان"""
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (ValueError, TypeError):
        return default


# ============================================================
# 📋 قواعد المطابقة (Governance Rules)
# ============================================================
RIYADH_RULES = [
    {"pm": "م. عبدالله الأسود", "id": 60, "name": "تنفيذ خطوط صرف صحي متفرقة بمدينة الرياض – عقد رقم 26 – المرحلة الثالثة",
     "contractor_keywords": ["سعد علي العيسي", "سعد العيسي", "العيسي"], "districts": None},
    {"pm": "م. تركي الاسمري", "id": 7, "name": "تنفيذ شبكة صرف صحي بأجزاء من احياء الحزم ونمار المرحلة الثالثة",
     "contractor_keywords": ["الخط الذهبي"], "districts": ["الحزم", "نمار"]},
    {"pm": "م. تركي الاسمري", "id": 2, "name": "عقد تنفيذ شبكات صرف صحي بحي الحائر",
     "contractor_keywords": ["المسار الحديث"], "districts": ["الحائر"]},
    {"pm": "م. عسكر لسوم", "id": 20, "name": "عقد تنفيذ شبكة صرف صحى بحى المعيزلية - المرحلة الأولى",
     "contractor_keywords": ["نظم البيئه", "نظم البيئة"], "districts": ["المعيزليه", "المعيزلية"]},
    {"pm": "م. عسكر لسوم", "id": 12, "name": "عقد تنفيذ شبكات الصرف الصحي بأجزاء من أحياء القدس والملك عبد الله",
     "contractor_keywords": ["ربوه التعمير", "راكو"], "districts": ["القدس", "الملك عبدالله"]},
    {"pm": "م. امجد الفالح", "id": 58, "name": "عقد تنفيذ شبكات الصرف الصحي بحي العوالي - مرحلة ثانية",
     "contractor_keywords": ["نظم البيئه", "نظم البيئة"], "districts": ["العوالي"]},
    {"pm": "م. امجد الفالح", "id": 57, "name": "عقد تنفيذ شبكات الصرف الصحي بحي العوالي (مرحلة أولى)",
     "contractor_keywords": ["الاعمال المدنيه", "الاعمال المدنية", "الأعمال المدنية"], "districts": ["العوالي"]},
    {"pm": "م. عبدالله العنزي", "id": 23, "name": "عقد تنفيذ شبكات المياه بحي المهدية (عقد رقم 3)بمدينة الرياض",
     "contractor_keywords": ["الدايل"], "districts": ["المهديه", "المهدية"]},
]

GOV_RULES = [
    {"pm": "م. سعيد الحارث", "id": 112, "name": "عقد استكمال مشاريع المياه بمحافظتي شقراء ومرات (المرحلة الاولي)",
     "contractor_keywords": ["ضيف الله العتيبي", "ضيف الله العتيبى"], "govs": ["شقراء", "مرات"]},
    {"pm": "م. سعيد الحارث", "id": 114, "name": "عقد تنفيذ شبكات الصرف الصحي بمدينة شقراء ( المرحلة السابعة )",
     "contractor_keywords": ["اضواء رتاج", "أضواء رتاج"], "govs": ["شقراء"]},
    {"pm": "م. سعيد الحارث", "id": 117, "name": "عقد تنفيذ و استكمال مشاريع المياه بمحافظة القويعية",
     "contractor_keywords": ["ماكسون"], "govs": ["القويعيه", "القويعية", "الرويضه", "الرويضة"]},
    {"pm": "م. سعيد الحارث", "id": 107, "name": "عقد استكمال مشاريع المياه بمدينة الدوادمي ومراكز البجادية ونفي",
     "contractor_keywords": ["مشروعات المياه والطاقه", "مشروعات المياة والطاقة", "مشروعات المياه والطاقة"],
     "govs": ["الدوادمي", "البجاديه", "البجادية", "نفي"]},
    {"pm": "م. شاكر الحقباني", "id": 95, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة الخرج (المرحلة السابعة)",
     "contractor_keywords": ["الخريف"], "govs": ["الخرج"]},
    {"pm": "م. شاكر الحقباني", "id": 96, "name": "عقد تنفيذ مشروع صرف صحي بحوطة بني تميم (المرحلة الثالثة )",
     "contractor_keywords": ["السبق العربي"], "govs": ["حوطه بني تميم", "حوطة بني تميم", "الحوطه", "الحوطة"]},
    {"pm": "م. شاكر الحقباني", "id": 97, "name": "عقد تنفيذ شبكات الصرف الصحي بحوطة بني تميم و الخرج (المرحلة الثانية)",
     "contractor_keywords": ["السبق العربي"], "govs": ["الخرج", "حوطه بني تميم", "حوطة بني تميم"]},
    {"pm": "م. شاكر الحقباني", "id": 99, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة الخرج",
     "contractor_keywords": ["مرامر"], "govs": ["الخرج"]},
    {"pm": "م. سعيد الحارث", "id": 113, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة المزاحمية (المرحلة الثانية )",
     "contractor_keywords": ["السبق العربي"], "govs": ["المزاحميه", "المزاحمية"]},
    {"pm": "م. سعيد الحارث", "id": 104, "name": "عقد تنفيذ شبكات الصرف الصحي بمحافظة ضرماء (المرحلة الثانية)",
     "contractor_keywords": ["مسره الوسطي", "مسرة الوسطى"], "govs": ["ضرماء", "ضرما"]},
    {"pm": "م. علي القحطاني", "id": 80, "name": "عقد تنفيذ خطوط نقل المياه بمراكز وقرى محافظة المجمعة (المرحلة الثانية)",
     "contractor_keywords": ["مشروعات المياه والطاقه"], "govs": ["المجمعه", "المجمعة"]},
    {"pm": "م. علي القحطاني", "id": 83, "name": "عقد تنفيذ شبكات الصرف الصحي بمدينة الزلفي ( المرحلة السادسة )",
     "contractor_keywords": ["بلر العربية"], "govs": ["الزلفي"]},
]


# ============================================================
# 🚪 البوابات الأربع للفلترة (4 Gates Filter)
# ============================================================
def gate_1_filter_status(status: str) -> bool:
    """البوابة 1: فلترة الحالات (البلاغات المعلقة فقط)"""
    valid_statuses = ["تحت معالجة المقاول", "بانتظار اعتماد الجهة المتعدية"]
    norm_status = normalize_arabic(status)
    return any(normalize_arabic(s) == norm_status for s in valid_statuses)


def gate_2_exclude_om(contractor: str, comment: str) -> bool:
    """البوابة 2: استبعاد التشغيل والصيانة (O&M)"""
    norm_contractor = normalize_arabic(contractor)
    norm_comment = normalize_arabic(comment)
    exclude_keywords = ["التشغيل والصيانه", "تشغيل وصيانه", "تاسي للتشغيل"]
    return not any(kw in norm_contractor or kw in norm_comment for kw in exclude_keywords)


def gate_3_match_govs(contractor: str, city: str, district: str, owner: str, comment: str):
    """البوابة 3: مطابقة محافظات القطاع الأوسط"""
    norm_contractor = normalize_arabic(contractor)
    norm_city = normalize_arabic(city)
    norm_dist = normalize_arabic(district)
    norm_owner = normalize_arabic(owner)
    norm_comment = normalize_arabic(comment)

    for rule in GOV_RULES:
        contractor_match = any(
            normalize_arabic(kw) in norm_contractor
            for kw in rule["contractor_keywords"]
        )
        if contractor_match:
            location_match = any(
                normalize_arabic(gov) in norm_city or
                normalize_arabic(gov) in norm_dist or
                normalize_arabic(gov) in norm_owner or
                normalize_arabic(gov) in norm_comment
                for gov in rule["govs"]
            )
            if location_match:
                return rule
    return None


def gate_4_match_riyadh(contractor: str, city: str, district: str):
    """البوابة 4: مطابقة مدينة الرياض"""
    norm_contractor = normalize_arabic(contractor)
    norm_city = normalize_arabic(city)
    norm_dist = normalize_arabic(district)

    if not (normalize_arabic("رياض") in norm_city or norm_city == "" or norm_city == "nan"):
        return None

    for rule in RİYADH_RULES:
        contractor_match = any(
            normalize_arabic(kw) in norm_contractor
            for kw in rule["contractor_keywords"]
        )
        if contractor_match:
            if rule["districts"] is None:
                return rule
            district_match = any(
                normalize_arabic(d) in norm_dist
                for d in rule["districts"]
            )
            if district_match:
                return rule
    return None


# ============================================================
# 📊 الدالة الرئيسية للتحليل
# ============================================================
def analyze_excel(input_path: Path, output_path: Path):
    """تحليل ملف الإكسيل وإنتاج ملف JSON معتمد"""
    print("=" * 70)
    print("🔹 NWC Encroachments Governance Analyzer")
    print("=" * 70)
    print(f"📂 ملف الإدخال: {input_path.name}")
    print(f"📅 تاريخ الاستخراج المرجعي: {BASE_DATE.strftime('%Y-%m-%d')}")
    print("-" * 70)

    # 1. قراءة ملف الإكسيل
    if not input_path.exists():
        print(f"❌ الملف غير موجود: {input_path}")
        return False

    try:
        df = pd.read_excel(input_path, engine='openpyxl')
    except Exception as e:
        print(f"❌ خطأ في قراءة الملف: {e}")
        return False

    total_rows = len(df)
    print(f"📊 إجمالي الصفوف في الملف: {total_rows}")

    # 2. تهيئة العدادات
    excluded_status = 0
    excluded_om = 0
    excluded_no_match = 0
    parsed_records = []

    # 3. معالجة كل صف
    for idx, row in df.iterrows():
        # استخراج الحقول بمرونة (دعم أسماء أعمدة مختلفة)
        rep_id = (row.get('رقم بلاغ التعدي') or row.get('رقم_بلاغ_التعدي') or
                  row.get('رقم البلاغ') or row.get('ID') or row.get('id'))
        if pd.isna(rep_id) or rep_id == '':
            continue

        status = str(row.get('حالة البلاغ') or row.get('حالة_البلاغ') or
                     row.get('الحالة') or 'تحت معالجة المقاول')

        # 🚪 البوابة 1: فلترة الحالة
        if not gate_1_filter_status(status):
            excluded_status += 1
            continue

        contractor = str(row.get('اسم المقاول') or row.get('المقاول المنفذ') or
                         row.get('contractor') or '')
        city = str(row.get('المدينة') or row.get('المحافظة') or 'مدينة الرياض')
        district = str(row.get('الحي') or row.get('الموقع') or '')
        comment = str(row.get('تعليق المركز') or row.get('وصف التعدي') or '')
        owner = str(row.get('الجهة المالكة') or '')
        date_val = str(row.get('تاريخ البلاغ') or row.get('تاريخ_البلاغ') or
                       row.get('تاريخ التعدي') or '2026-01-01')

        # تطبيع التاريخ
        if 'T' in date_val:
            date_val = date_val.split('T')[0]
        if ' ' in date_val:
            date_val = date_val.split(' ')[0]

        lat = safe_float(row.get('خط العرض') or row.get('خط_العرض') or
                         row.get('Lat') or row.get('latitude'))
        lng = safe_float(row.get('خط الطول') or row.get('خط_الطول') or
                         row.get('Lng') or row.get('longitude'))

        desc = str(row.get('وصف التعدي') or row.get('الإجراء المطلوب') or
                   'أعمال حفريات وتمديد شبكات بدون استكمال إجراءات إخلاء الطرف')

        # 🚪 البوابة 2: استبعاد التشغيل والصيانة
        if not gate_2_exclude_om(contractor, comment):
            excluded_om += 1
            continue

        # 🚪 البوابة 3: مطابقة المحافظات
        matched_rule = gate_3_match_govs(contractor, city, district, owner, comment)

        # 🚪 البوابة 4: مطابقة الرياض
        if not matched_rule:
            matched_rule = gate_4_match_riyadh(contractor, city, district)

        if not matched_rule:
            excluded_no_match += 1
            continue

        # حساب الإحداثيات الافتراضية
        if lat is None:
            lat = 24.7136 if city == 'مدينة الرياض' else 25.2388
        if lng is None:
            lng = 46.6753 if city == 'مدينة الرياض' else 45.2775

        # بناء السجل النهائي
        record = {
            "id": matched_rule["id"],
            "name": matched_rule["name"],
            "contractor": contractor if contractor != 'nan' else "مقاول معتمد",
            "program_manager_nwc": matched_rule["pm"],
            "المدينة": city if city != 'nan' else "مدينة الرياض",
            "المحافظة": city if city != 'nan' else "مدينة الرياض",
            "الحي": district if district != 'nan' else "موقع معتمد",
            "خط_العرض": lat,
            "خط_الطول": lng,
            "رقم_بلاغ_التعدي": int(rep_id),
            "تاريخ_البلاغ": date_val,
            "حالة_البلاغ": status,
            "وصف_التعدي": desc if desc != 'nan' else "أعمال حفريات وتمديد شبكات",
            "الإجراء_المطلوب": "إلزام المقاول بالمعالجة الميدانية الفورية وإغلاق البلاغ بنظام المركز"
        }
        parsed_records.append(record)

    # 4. عرض التقرير النهائي
    print("\n📈 تقرير التحليل:")
    print(f"   ✅ بلاغات معتمدة: {len(parsed_records)}")
    print(f"   ❌ مستبعدة (حالة منتهية): {excluded_status}")
    print(f"   ❌ مستبعدة (تشغيل وصيانة): {excluded_om}")
    print(f"   ❌ مستبعدة (لا تطابق قاعدة): {excluded_no_match}")
    print(f"   📊 المجموع: {len(parsed_records) + excluded_status + excluded_om + excluded_no_match}")

    # 5. إحصائيات إضافية
    if parsed_records:
        riyadh_count = sum(1 for r in parsed_records if r["المدينة"] == "مدينة الرياض")
        govs_count = len(parsed_records) - riyadh_count
        print(f"\n🗺️ التوزيع الجغرافي:")
        print(f"   🏙️ مدينة الرياض: {riyadh_count} بلاغاً")
        print(f"   🏘️ المحافظات: {govs_count} بلاغاً")

        pm_counts = {}
        for r in parsed_records:
            pm = r["program_manager_nwc"]
            pm_counts[pm] = pm_counts.get(pm, 0) + 1
        print(f"\n👔 التوزيع حسب مدير البرنامج:")
        for pm, count in sorted(pm_counts.items(), key=lambda x: -x[1]):
            print(f"   • {pm}: {count} بلاغاً")

    # 6. حفظ الملف
    if parsed_records:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(parsed_records, f, ensure_ascii=False, indent=2)
        print(f"\n💾 تم حفظ الملف بنجاح: {output_path}")
        print(f"📦 حجم الملف: {output_path.stat().st_size / 1024:.2f} KB")
        print("\n✨ يمكنك الآن رفع المشروع إلى Vercel!")
        return True
    else:
        print("\n⚠️ لم يتم العثور على أي بلاغات مطابقة للمعايير.")
        return False


# ============================================================
# 🚀 نقطة الدخول الرئيسية
# ============================================================
if __name__ == "__main__":
    # التحقق من وجود ملف الإدخال
    if not INPUT_FILE.exists():
        print(f"\n⚠️ الملف الافتراضي غير موجود: {INPUT_FILE}")
        print("🔍 جاري البحث عن ملفات XLSX في المجلد...")

        xlsx_dir = BASE_DIR / "XLSX"
        if xlsx_dir.exists():
            xlsx_files = list(xlsx_dir.glob("*.xlsx"))
            if xlsx_files:
                print("\n📋 الملفات المتاحة:")
                for i, f in enumerate(xlsx_files, 1):
                    print(f"   {i}. {f.name}")
                choice = input("\n🔢 اختر رقم الملف للتحليل (أو اتركه فارغاً للخروج): ")
                if choice.isdigit() and 1 <= int(choice) <= len(xlsx_files):
                    INPUT_FILE = xlsx_files[int(choice) - 1]
                else:
                    exit(0)
            else:
                print("❌ لا توجد ملفات XLSX في المجلد.")
                exit(1)
        else:
            print("❌ مجلد XLSX غير موجود.")
            exit(1)

    # تشغيل المحلل
    success = analyze_excel(INPUT_FILE, OUTPUT_FILE)
    exit(0 if success else 1)
