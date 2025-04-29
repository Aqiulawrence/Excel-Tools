import warnings

from tkinter import messagebox
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

warnings.filterwarnings("ignore")

def resize_image(input_image_path, base_width, base_height):
    with Image.open(input_image_path) as img:
        orig_width, orig_height = img.size

        ratio = min(base_width / orig_width, base_height / orig_height)
        new_width = int(orig_width * ratio)
        new_height = int(orig_height * ratio)

        img = img.convert('RGB')

        img = img.resize((new_width, new_height), Image.LANCZOS)
        return img, new_width, new_height

def insert(excel_path, image_path, cell_ref):
    global wb, ws
    wb = load_workbook(excel_path)
    ws = wb.active

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
    img_original_width, img_original_height = img.width, img.height
    # 计算图片缩放比例（保持宽高比）
    scale_width = width / img_original_width
    scale_height = height / img_original_height
    scale = min(scale_width, scale_height)  # 选择最小的缩放比例

    # 调整图片的大小
    img.width = int(img_original_width * scale)
    img.height = int(img_original_height * scale)

    ws.add_image(img, f"{cell_ref[0]}{cell_ref[1]}")

    try:
        wb.save(excel_path)
    except PermissionError:
        messagebox.showerror('错误', '插入失败！请确保Excel不为只读文件并且Excel已关闭！')
        raise
    return True

def main(start_cell, total_num, excel_path, image_base_path=".\\img"):
    image_names = []
    image_base_path += '\\'

    for i in range(1, total_num + 1):
        image_names.append(f'{i:03d}.png')

    error_insert = 0 # 没有被插入的图片
    for name in image_names:
        img_path = image_base_path + name
        state = insert(excel_path, img_path, start_cell)
        if not state:
            error_insert += 1
            print('Inserted Failed')
        elif state == 'STOP':
            return 'STOP'
        else:
            print(f'Successfully inserted {img_path}')
        start_cell[1] += 1

    return error_insert

if __name__ == '__main__':
    main(['A', 1], 30, 'test.xlsx')
