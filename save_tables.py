import requests # برای ارسال درخواست‌ های HTTP
import os # ساخت پوشه
import datetime # کار با تاریخ
import jdatetime # برای تبدیل تاریخ ها
import pandas as pd # خواندن و پردازش داده‌های جدولی
import glob # جست‌ و جو در مسیرها برای پیدا کردن فایل‌ ها
from bs4 import BeautifulSoup # برای تجزیه و تحلیل محتوای HTML
from html2excel import ExcelParser # تبدیل جدول‌ های HTML به فایل اکسل
from deep_translator import GoogleTranslator # ترجمه متن 

url = "https://tradingeconomics.com/commodities"

# هدر مرورگر برای جلوگیری از بلوکه‌ شدن توسط سایت
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0 Safari/537.36"
}

response = requests.get(url, headers=headers)

def translate_text(text):
    """ترجمه ی متن به فارسی"""

    if pd.notna(text): # بررسی null نبودن مقدار
        try:
            translated = GoogleTranslator(source='en', target='fa').translate(str(text))
            return translated
        except Exception as e:
            print(f"خطا در ترجمه {text}: {e}")
            return text
    return text

def process_excel(file_path, folder_path):
        """گرفتن سلول های مورد نظر و ترجمه ی آنها"""

        print(f"پردازش فایل: {file_path}")

        df = pd.read_excel(file_path)

        # ترجمه عنوان ستون‌ها
        df.columns = [translate_text(col) for col in df.columns]

        column_9_name = df.columns[8]
        for row in df.index:
            for col in df.columns:
                value = df.at[row, col]
                if pd.notna(value) and isinstance(value, str): # ترجمه فقط سلول‌ های متنی
                    # ردیف های زیر عنوان تاریخ ترجمه نشوند چون بهتره که ماه ها انگلیسی باقی بمانند
                    if col == column_9_name:
                        continue
                    df.at[row, col] = translate_text(str(value))
        # پوشه‌ برای ذخیره فایل‌ های ترجمه‌ شده
        translated_folder = f"Translated_{os.path.basename(folder_path)}"
        os.makedirs(translated_folder, exist_ok=True)
        
        # تعیین مسیر خروجی فایل ترجمه‌ شده
        output_file = os.path.join(translated_folder, os.path.basename(file_path))
        df.to_excel(output_file, index=False)
        print(f"ذخیره فایل ترجمه‌شده: {output_file}")

# HTML -> BeautifulSoup
if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')
    tables = soup.find_all('table')

    print(f"تعداد جداول: {len(tables)}")

    # طبقه بندی کالاها
    names = ["انرژی", "فلزات", "کشاورزی", "صنعتی", "دام", "شاخص", "برق"]

    # گرفتن تاریخ امروز به صورت شمسی و استفاده از آن در ساخت اسم فولدر برای اکسل های اولیه
    now = jdatetime.date.fromgregorian(date=datetime.date.today())
    now_str = now.strftime(f'%Y-%m-%d')
    folder = f'Tables_{now_str}'
    os.makedirs(folder, exist_ok=True)
    
    # ذخیره HTML 
    for i, table in enumerate(tables):
        file_path = rf'{folder}\table_{i+1}.html'
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(str(table))

        # ذخیره هر جدول به صورت فایل اکسل قبل از ترجمه  
        output_file = rf'{folder}\{names[i]}.xlsx'

        # تبدیل HTML به جدول
        parser = ExcelParser(file_path)
        
        parser.to_excel(output_file)

        print(f"جدول {i+1} ذخیره شد")
       
        # حذف HTML ذخیره شده
        os.remove(file_path)
    
    # پیدا کردن همه ی اکسل های داخل پوشه
    excel_files = glob.glob(os.path.join(folder, "*.xlsx"))

    if not excel_files:
        print("هیچ فایل اکسلی یافت نشد")
    else:
        # پردازش و ترجمه هر فایل اکسل
        for file_path in excel_files:
            process_excel(file_path, folder)
else:
    print(f"خطا: {response.status_code}")
