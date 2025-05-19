import requests
import os
import datetime
import jdatetime
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
else:
    print(f"خطا: {response.status_code}")

