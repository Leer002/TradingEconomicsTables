import requests
import os
import datetime
import jdatetime
import pandas as pd
import asyncio
import shutil
import glob
from googletrans import Translator
from bs4 import BeautifulSoup
from html2excel import ExcelParser

url = "https://tradingeconomics.com/commodities"

# https://www.whatismybrowser.com/detect/what-is-my-user-agent 
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0 Safari/537.36"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')
    tables = soup.find_all('table')
    first_headers = []
    for table in tables:
        first_th = table.find('th') 
        if first_th:
            first_headers.append(first_th.get_text(strip=True))
    
    print(f"تعداد جداول: {len(tables)}")

    now = jdatetime.date.fromgregorian(date=datetime.date.today())
    now_str = now.strftime(f'%Y-%m-%d')

    folder = f'Table{now_str}'
    os.makedirs(folder, exist_ok=True)

    for i, table in enumerate(tables):
        file_path = rf'{folder}\table_{i+1}.html'  
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(str(table))
        
        output_file = rf'{folder}\{first_headers[i]}.xlsx'  
        parser = ExcelParser(file_path)
        parser.to_excel(output_file)

        os.remove(file_path) 
        print(f"جدول {i+1}ذخیره")

        translator = Translator()

        async def translate_text(text):
            if pd.notna(text):
                translated = await translator.translate(text, src='en', dest='fa')
                return translated.text
            return text

        async def process_excel(file_path, folder_path):
            print(f"پردازش فایل: {file_path}") 

            df = pd.read_excel(file_path)

            header_tasks = {col: asyncio.create_task(translate_text(col)) for col in df.columns}
            header_results = await asyncio.gather(*header_tasks.values())
            df.columns = header_results 

            
            tasks = {}
            for row in df.index:
                for col in df.columns:
                    if pd.notna(df.at[row, col]) and col != df.columns[8]:
                        tasks[(row, col)] = asyncio.create_task(translate_text(df.at[row, col]))
                    elif pd.notna(df.at[row, col]) and col == df.columns[8] and df.index[0]:
                        tasks[(row, col)] = asyncio.create_task(translate_text(df.at[row, col]))
            
            results = await asyncio.gather(*tasks.values())
            
            for (row, col), translated_text in zip(tasks.keys(), results):
                df.at[row, col] = translated_text

            output_file = os.path.join(f"translated_{os.path.basename(folder_path)}", f"translated_{os.path.basename(file_path)}")
            os.makedirs(f"translated_{os.path.basename(folder_path)}", exist_ok=True)  
            df.to_excel(output_file, index=False)

            print(f"ذخیره فایل ترجمه‌شده: {output_file}")

        async def main():
            folder_path = rf"{folder}"
            
            excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

            if not excel_files:
                print(" هیچ فایل اکسلی یافت نشد")
                return

            for file_path in excel_files:
                await process_excel(file_path, folder_path)

        asyncio.run(main())
else:
    print(f"خطا: {response.status_code}")

# shutil.rmtree(folder)

