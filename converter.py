import pandas as pd
import json
import os
from pathlib import Path
from datetime import datetime

# المسار الحالي
BASE_DIR = Path(__file__).parent

def convert_timestamp_to_string(obj):
    """تحويل Timestamp إلى string"""
    if isinstance(obj, pd.Timestamp):
        return obj.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(obj, datetime):
        return obj.strftime('%Y-%m-%d %H:%M:%S')
    return obj

def convert_excel_to_json():
    """تحويل ملفات Excel إلى JSON"""
    
    # الملفات المراد تحويلها
    files_to_convert = {
        'pass.xlsx': 'data/passwords.json',
        'new-uplode.xlsx': 'data/branches.json'
    }
    
    # إنشاء مجلد data إذا لم يكن موجوداً
    data_dir = BASE_DIR / 'data'
    data_dir.mkdir(exist_ok=True)
    
    for excel_file, json_file in files_to_convert.items():
        excel_path = BASE_DIR / excel_file
        json_path = BASE_DIR / json_file
        
        if not excel_path.exists():
            print(f"❌ الملف {excel_file} غير موجود")
            continue
        
        try:
            print(f"🔄 جاري تحويل {excel_file}...")
            
            # قراءة ملف Excel
            df = pd.read_excel(excel_path)
            
            # ملء القيم الفارغة
            df = df.fillna('')
            
            # تحويل جميع الأعمدة إلى نصوص أو أنواع صحيحة
            for col in df.columns:
                if df[col].dtype == 'datetime64[ns]':
                    # تحويل التواريخ إلى string
                    df[col] = df[col].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '')
                else:
                    # تحويل أي Timestamp متبقي
                    df[col] = df[col].apply(lambda x: convert_timestamp_to_string(x) if pd.notna(x) else '')
            
            # تحويل إلى قائمة من الـ dictionaries
            data = df.to_dict('records')
            
            # تنظيف البيانات من Timestamp
            data = json.loads(json.dumps(data, default=str, ensure_ascii=False))
            
            # حفظ كـ JSON
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            print(f"✅ تم تحويل {excel_file} إلى {json_file}")
            print(f"   عدد الصفوف: {len(data)}")
            
        except Exception as e:
            print(f"❌ خطأ في تحويل {excel_file}: {str(e)}")

def display_json_preview():
    """عرض معاينة من البيانات المحولة"""
    print("\n" + "="*50)
    print("📊 معاينة البيانات المحولة:")
    print("="*50)
    
    data_dir = BASE_DIR / 'data'
    
    for json_file in data_dir.glob('*.json'):
        print(f"\n📄 ملف: {json_file.name}")
        
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            print(f"   عدد السجلات: {len(data)}")
            
            if data:
                print(f"   الأعمدة: {list(data[0].keys())}")
                print(f"\n   أول سجل:")
                for key, value in list(data[0].items())[:3]:
                    print(f"      {key}: {value}")
        except Exception as e:
            print(f"   ❌ خطأ في قراءة الملف: {e}")

if __name__ == '__main__':
    print("🚀 برنامج تحويل Excel إلى JSON\n")
    convert_excel_to_json()
    display_json_preview()
    print("\n✨ اكتمل التحويل! الملفات موجودة في مجلد 'data'")