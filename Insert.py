from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import shutil
from pathlib import Path
import os
import warnings
from tkinter import messagebox
import PIL
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor
from openpyxl.utils.units import pixels_to_EMU

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

def insert_resized_image_to_excel(excel_path, image_path, cell_ref):
    global wb, ws
    wb = load_workbook(excel_path)
    ws = wb.active

    width = ws.column_dimensions[cell_ref[0]].width * 72 // 9 # 一个单元格高为13.5，像素为18
    height = ws.row_dimensions[cell_ref[1]].height * 18 // 13.5 # 一个单元格为宽为9，像素为72

    try:
        resized_img, img_width, img_height = resize_image(image_path, width, height)
    except PIL.UnidentifiedImageError:
        return False

    x = (width-img_width) / 2
    y = (height-img_height) / 2

    temp_img_path = 'temp_resized_image.jpg'
    resized_img.save(temp_img_path)

    img = OpenpyxlImage(temp_img_path)

    ws.add_image(img, f"{cell_ref[0]}{cell_ref[1]}")

    try:
        wb.save(excel_path)
    except PermissionError:
        messagebox.showerror('错误', '插入失败！请确保Excel不为只读文件并且Excel已关闭！')
        raise
    return True


def main(start_cell, excel_path, image_base_path=".\\img"):
    image_names = []
    image_base_path += '\\'
    folder_path = Path(image_base_path)

    for file in folder_path.iterdir():
        if file.is_file():
            image_names.append(file.name)

    error_insert = 0
    for idx, image_name in enumerate(image_names):
        img_path = image_base_path + image_name
        if not insert_resized_image_to_excel(excel_path, img_path, start_cell):
            error_insert += 1
            print('Insert Failed')
        else:
            print(f'Successfully inserte {img_path}')
        start_cell[1] += 1

    os.remove("./temp_resized_image.jpg")
    return error_insert

if __name__ == '__main__':
    main(['B', 1], 'test.xlsx')
