import os
from openpyxl import load_workbook
from tkinter import messagebox
from colorama import Fore, Style

def red_text(text):
    return Fore.RED+Style.BRIGHT+text+Style.RESET_ALL

def blue_text(text):
    return Fore.BLUE+Style.BRIGHT+text+Style.RESET_ALL

def yellow_text(text):
    return Fore.YELLOW+Style.BRIGHT+text+Style.RESET_ALL

def green_text(text):
    return Fore.GREEN+Style.BRIGHT+text+Style.RESET_ALL

REPLACE = {
    '%20': '+',
    ' ': '%20',
    '!': '%21',
    "'": '%27',
    '(': '%28',
    ')': '%29',
    '~': '%7E',
    '%00': r'\x00'
}

def find_data_in_single_excel(folder_path, data_to_find):
    try:
        file_path = folder_path
        if file_path.endswith('.xlsx'):
            workbook = load_workbook(file_path)
            for sheet in workbook.sheetnames:
                worksheet = workbook[sheet]
                for row in worksheet.iter_rows():
                    for cell in row:
                        if data_to_find.lower().strip() in str(cell.value).lower().strip():
                            print(f'在 "{file_path}" 中的 "{cell.coordinate}" 找到 "{data_to_find}"')

    except PermissionError:
        messagebox.showerror('错误', f'{folder_path} 拒绝访问。')
        return False
    return True

def find_data_in_multiple_excel(folder_path, data_to_find):
    # 遍历Excel文件
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if not file.endswith('.xlsx'):
                continue
            full_path = os.path.join(root, file)
            full_path = full_path.replace('\\', '/')

            try:
                workbook = load_workbook(full_path)
            except PermissionError as e:
                if "~$" in str(e):
                    continue
                print(yellow_text(str(e)))
                continue

            # 初始化
            data = {'rmb': None, 'usd': None, 'part': None, 'description': None}
            data2 = {'part': "件号", 'description': "品名", 'rmb': "RMB", 'usd': "USD"}
            # 遍历每个单元格
            for sheet in workbook.sheetnames:
                worksheet = workbook[sheet]
                for row in worksheet.iter_rows():
                    for cell in row:
                        for i in data: # i 为需求字符串
                            if data[i] is None and i in str(cell.value).lower().strip():
                                data[i] = cell.coordinate[0]
                                #if sum(1 for v in data.values() if v is not None) == 3: # 变量flag已废弃...

                        if data_to_find.lower().strip() in str(cell.value).lower().strip():
                            print(f'在 "{(full_path)}" "{(cell.coordinate)}" 找到目标')
                            is_first = True
                            for i in data:
                                if data[i] is None:
                                    continue
                                index = data[i]+cell.coordinate[1:]
                                if not is_first: # 判断是否要添加分隔符
                                    print(f"  {yellow_text("|")}  ", end="")

                                if i == 'usd' or i == 'rmb':
                                    output = red_text(str(worksheet[index].value))
                                else:
                                    output = blue_text(str(worksheet[index].value))

                                print(f"{data2[i]}: {output}", end="") # 可以更改颜色
                                is_first = False
                            print('\n')

if __name__ == '__main__':
    folder_path = r'D:/'
    data_to_find = '111'
    find_data_in_multiple_excel(folder_path, data_to_find)
    # find_data_in_single_excel(r'D:\Excel Project\excel\test.xlsx', data_to_find)