import tkinter as tk
from tkinter import filedialog
import os
from ast import literal_eval
import json
import time

def get_time():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

CONFIG_DIR = rf'.\configs'
CONFIG_FILE2 = rf'{CONFIG_DIR}\cf2.json'
set_value_len = 4

def load():
    global data
    data = {}
    if os.path.exists(CONFIG_FILE2):
        with open(CONFIG_FILE2, 'r') as f:
            try:
                data = literal_eval(f.read())
                return data
            except: pass
    save(True)
    return data

def save(create=False):
    global data
    data = {}
    if create:
        data['path'] = r'.\img'
        data['auto_update'] = 1
        data['filter'] = 1
        data['auto_backup'] = 1
    else:
        data['path'] = Svar1.get()
        data['auto_update'] = var1.get()
        data['filter'] = var2.get()
        data['auto_backup'] = var3.get()

    with open(CONFIG_FILE2, 'w') as f:
        json.dump(data, f)

def select():
    file_path = filedialog.askdirectory()
    if file_path:
        Svar1.set(file_path)

def main():
    global Svar1, var1, var2, var3
    root = tk.Tk()
    root.title("设置（关闭自动保存）")
    root.geometry("310x200+400+200") # 一个选项+40px，添加后要更改上面的len
    root.resizable(width=False, height=False)
    root.attributes("-topmost", True)

    # 导出到main()函数set_value列表
    Svar1 = tk.StringVar() # 图片输出
    var1 = tk.IntVar() # 启用自动更新
    var2 = tk.IntVar() # 网站筛选
    var3 = tk.IntVar() # 移动前备份

    load()
    try:
        Svar1.set(data['path'])
        var1.set(data['auto_update'])
        var2.set(data['filter'])
        var3.set(data['auto_backup'])
    except KeyError:
       pass

    lf1 = tk.LabelFrame(root, text='基础设置')
    lf1.grid(row=0, column=0, padx=10, pady=5, sticky=tk.NW)
    lb1 = tk.Label(lf1, text='图片输出：')
    lb1.grid(row=0, column=0, padx=10, pady=5, sticky=tk.NW)
    et1 = tk.Entry(lf1, textvariable=Svar1, state=tk.DISABLED)
    et1.grid(row=0, column=1, padx=5, pady=5, sticky=tk.NW)
    bt1 = tk.Button(lf1, text='选择', command=select)
    bt1.grid(row=0, column=2, padx=5, pady=5, sticky=tk.NW)

    f1 = tk.Frame(lf1)
    f1.grid(row=1, column=0, padx=5, pady=5, sticky=tk.NW, columnspan=10)
    cb1 = tk.Checkbutton(f1, text='启用自动更新', variable=var1)
    cb1.grid(row=0, column=0, sticky=tk.NW, padx=5, pady=5)
    cb2 = tk.Checkbutton(f1, text='搜图启用网站筛选（效果未知）', variable=var2)
    cb2.grid(row=1, column=0, sticky=tk.NW, padx=5, pady=5)
    cb3 = tk.Checkbutton(f1, text='移动件号前备份文件', variable=var3)
    cb3.grid(row=2, column=0, sticky=tk.NW, padx=5, pady=5)

    root.mainloop()
    save()

if __name__ == '__main__':
    main()