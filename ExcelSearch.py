import os
import threading
import multiprocessing
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from openpyxl import load_workbook
from colorama import Fore, Style

def red_text(text):
    return Fore.RED+Style.BRIGHT+text+Style.RESET_ALL

def blue_text(text):
    return Fore.BLUE+Style.BRIGHT+text+Style.RESET_ALL

def yellow_text(text):
    return Fore.YELLOW+Style.BRIGHT+text+Style.RESET_ALL

print_lock = threading.Lock()

def singleSearch(file_path, search_term):
    try:
        workbook = load_workbook(file_path, read_only=True)  # 使用只读模式提高性能
        results = []

        for sheet in workbook.sheetnames:
            data = {'rmb': None, 'usd': None, 'part': None, 'description': None}
            data2 = {'part': "件号", 'description': "品名", 'rmb': "RMB", 'usd': "USD"}

            worksheet = workbook[sheet]
            for row in worksheet.iter_rows():
                for cell in row:
                    cell_value = str(cell.value).lower().strip()
                    # 检测需要打印的数据在哪列
                    for i in data:
                        if data[i] is None and i in cell_value:
                            data[i] = cell.coordinate[0]

                    if search_term in cell_value:
                        results.append(f'在 "{file_path}" "{sheet}" "{cell.coordinate}" 找到目标')
                        output_parts = []
                        for i in data:
                            if data[i] is not None:
                                index = data[i]+cell.coordinate[1:]
                                value = str(worksheet[index].value)
                                colored = red_text(value) if i in ('usd', 'rmb') else blue_text(value)
                                output_parts.append(f"{data2[i]}: {colored}")

                        results.append(f"  {yellow_text('|')}  ".join(output_parts) + "\n")

        with print_lock:
            for result in results:
                print(result)

    finally:
        if 'workbook' in locals():
            workbook.close()


def multipleSearch(folder_path, data_to_find, max_workers=8):
    search_term = data_to_find.lower().strip()

    file_paths = [
        os.path.join(root, file).replace('\\', '/')
        for root, _, files in os.walk(folder_path)
        for file in files
        if file.endswith('.xlsx') and not file.startswith('~$')
    ]

    # 改用进程池
    with multiprocessing.Pool(processes=max_workers) as pool:
        pool.starmap(singleSearch, [(fp, search_term) for fp in file_paths])

if __name__ == '__main__':
    folder_path = r'D:\stress test'
    data_to_find = 'ak+q'
    multipleSearch(folder_path, data_to_find)