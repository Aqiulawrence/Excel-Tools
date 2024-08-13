import os
from openpyxl import load_workbook
from tkinter import messagebox

def find_data_in_single_excel(folder_path, data_to_find):
    try:
        file_path = folder_path
        if file_path.endswith('.xlsx'):
            workbook = load_workbook(file_path)
            for sheet in workbook.sheetnames:
                worksheet = workbook[sheet]
                for row in worksheet.iter_rows():
                    for cell in row:
                        if data_to_find.lower() in str(cell.value).lower():
                            print(f'在 "{file_path}" 中的 "{cell.coordinate}" 找到 "{data_to_find}"')
    except PermissionError:
        messagebox.showerror('错误', f'{folder_path} 拒绝访问。')
        return False
    return True

def find_data_in_multiple_excel(folder_path, data_to_find):
    try:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if not file.endswith('.xlsx'):
                    continue
                full_path = os.path.join(root, file)
                workbook = load_workbook(full_path)
                for sheet in workbook.sheetnames:
                    worksheet = workbook[sheet]
                    for row in worksheet.iter_rows():
                        for cell in row:
                            if data_to_find.lower() in str(cell.value).lower():
                                print(f'在 "{full_path}" 中的 "{cell.coordinate}" 找到 "{data_to_find}"')
    except PermissionError: pass

if __name__ == '__main__':
    folder_path = r'D:/'
    data_to_find = '111'
    find_data_in_multiple_excel(folder_path, data_to_find)