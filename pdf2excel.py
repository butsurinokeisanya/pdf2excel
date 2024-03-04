import os
import re
import pdfplumber
from openpyxl import Workbook
##
def extract_pdf_content(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
        return text

def main():
    pdf_files = [file for file in os.listdir() if file.endswith('.pdf')]
    
    if not pdf_files:
        print("カレントディレクトリにPDFファイルが見つかりません。")
        return
    
    wb = Workbook()
    
    for pdf_file in pdf_files:
        sheet_name = os.path.splitext(pdf_file)[0]
        content = extract_pdf_content(pdf_file)
        ws = wb.create_sheet(title=sheet_name)
        
        # 「■」で始まる行ごとにセルに分割して書き込む
        lines = content.split('\n')
        current_row = 1
        for line in lines:
            if line.startswith('■'):
                # 項目名と内容を分割する
                parts = re.findall(r'■[^■]*', line)
                for part in parts:
                    parts_split = part.split()
                    # 円を削除してから書き込む
                    value = parts_split[1].replace('円', '')
                    ws.append([parts_split[0], value])
    
    del wb['Sheet']  # デフォルトのシートを削除
    
    wb.save('pdf_contents.xlsx')
    print("内容がExcelファイルに出力されました。")

if __name__ == "__main__":
    main()
