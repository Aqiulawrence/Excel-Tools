from tkinter import messagebox
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from colorama import Fore, Style

base_path = "./img/"

def red_text(text):
    return Fore.RED+Style.BRIGHT+text+Style.RESET_ALL

def insert(ws, image_path, cell_ref):
    # 如果w1为13，w2为None，使用w1; 如果w1为13，w2不为None，使用w2; 如果w1不为13，使用w1
    # 如果h1为None，使用h2; 如果h1不为None，使用h1
    w1 = ws.column_dimensions[cell_ref[0]].width
    h1 = ws.row_dimensions[cell_ref[1]].height
    w2 = ws.sheet_format.defaultColWidth
    h2 = ws.sheet_format.defaultRowHeight
    if w1 == 13:
        if w2 is None:
            width = w1
        else:
            width = w2
    else:
        width = w1

    if h1 is None:
        height = h2
    else:
        height = h1

    width *= 8  # 一个单元格为宽为9，像素为72（待确认？）
    height *= 1.3  # 一个单元格高为13.5，像素为18（待确认？）

    img = Image(image_path)
    original_width, original_height = img.width, img.height

    # 计算图片缩放比例（保持宽高比）
    scale_width = width / original_width
    scale_height = height / original_height
    scale = min(scale_width, scale_height)  # 选择最小的缩放比例

    # 调整图片的大小
    img.width = int(original_width * scale)
    img.height = int(original_height * scale)

    ws.add_image(img, f"{cell_ref[0]}{cell_ref[1]}")

def main(cell_ref, total_img, excel_path): # start_cell 为一个列表，比如['A', 1]
    wb = load_workbook(excel_path)
    ws = wb.active

    img_names = [f'{i:03d}.png' for i in range(total_img)]
    error_count = 0 # 没有被插入的图片

    for index, name in enumerate(img_names):
        img_path = base_path + name

        try:
            insert(ws, img_path, cell_ref)
        except Exception:
            error_count += 1
            print(red_text(f'[Failed]Index:{index}'))
        cell_ref[1] += 1

    while True:
        try:
            wb.save(excel_path)
            break
        except PermissionError:
            if messagebox.askretrycancel('警告', '保存失败！请确保文件未被占用且不是只读文件。\n点击“重试”再次尝试保存。'):
                continue
            raise

    return error_count

if __name__ == '__main__':
    main(['A', 1], 30, 'test.xlsx')
