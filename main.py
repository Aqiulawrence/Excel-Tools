import Extract
import GoogleSearch
import Insert
import CheckUpdate
import Settings
import MoveKey
import ExcelSearch
import PostData
import ExcelCheck

import tkinter as tk
import shutil
import traceback
import sys
import requests
import webbrowser
import winsound
import os
import subprocess
import tkinterdnd2
import json
import time
import re
import ctypes

from colorama import Fore, Style
from threading import Thread
from ast import literal_eval
from tkinter import messagebox
from tkinter import filedialog
from easygui import exceptionbox
from socket import gethostname

##### 正常更新记得删除兼容设置部分代码
update_content = '''
1. 优化了搜索显示。
'''

'''
# 删除旧版本配置文件，兼容部分代码
from glob import glob
ini_files = glob(os.path.join('./configs', "*.ini"))
for ini_file in ini_files:
    try:
        os.remove(ini_file)
    except Exception as e:
        pass
'''

def get_time():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

class MultiStream: # 多重错误流
    def __init__(self, *streams):
        self.streams = streams
        self.isWrite = False

    def write(self, message):
        for stream in self.streams:
            stream.write(message)
            stream.flush()
        self.isWrite = True

    def flush(self):
        for stream in self.streams:
            stream.flush()

CONFIG_DIR = rf'.\configs' # 此常量仅在 main.py 和 Settings.py 出现
CONFIG_FILE1 = rf'.\configs\cf1.json' # 此文件作用是保存主窗口内输入框的值
LOG_DIR = rf'.\logs' # 此常量在 GoogleSearch.py 也存在

if not os.path.exists(CONFIG_DIR):
    os.makedirs(CONFIG_DIR)
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# 设置多重错误流
err_log = open(rf'{LOG_DIR}\error.log', 'a')
multi_stream = MultiStream(sys.stderr, err_log)
sys.stderr = multi_stream

VERSION = "2.0" # 当前版本
NEW = None # 最新版本
id = None # 蓝奏云文件的id，爬取下载地址需要用到
top = True # 窗口是否置顶
record = [f'!{get_time()}'] # 记录程序的功能使用情况

def isFirstOpen(): # 获取是否为第一次打开程序
    path = rf'{CONFIG_DIR}\record.json' # 这个文件只记录打开过的版本号
    if not os.path.exists(path):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({VERSION: True}, f)
            return True
    else:
        with open(path, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
            except json.decoder.JSONDecodeError:
                return False
            if VERSION in data:
                return False
            else:
                data[VERSION] = True
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(data, f)
                return True

def deleteOld(): # 删除旧版本
    if os.getcwd().endswith('dist'):
        return None
    flag = False
    pattern = r'Excel Tools v(\d+.\d+).exe'
    for file in os.listdir(os.getcwd()):
        if os.path.isfile(file):
            ret = re.match(pattern, file)
            if ret:
                ver = float(ret.group(1))
                if ver < float(VERSION):
                    os.remove(file)
                    flag = True
    if flag:
        # print('已删除旧版本。')
        return True
    return False

def update(auto=False):  # auto表示该函数是否为自动更新调用的，如果是就不弹窗
    global NEW, id
    NEW, id = CheckUpdate.check_update()
    if NEW == 0:
        messagebox.showerror('错误', '无法连接至更新服务器！请检查网络状态。')
    elif VERSION >= NEW:
        if not auto:
            messagebox.showinfo('提示', '当前程序为最新版本，无需更新。')
        else:
            print(green_text('当前程序为最新版本！'))
    elif VERSION < NEW:
        if messagebox.askyesno('提示', f'发现新版本v{NEW}，当前版本v{VERSION}\n按下”是“进行自动更新，按下”否“忽略本次更新。\n注：下载可能需要一定的时间，请耐心等待，请勿关闭此程序！'):
            try:
                CheckUpdate.download_update(id, f'Excel Tools v{NEW}.exe')
            except:
                exceptionbox(title='错误', msg='更新失败！请检查网络状态，并将此错误报告发送至管理员Sam！')
                return
            messagebox.showinfo('提示', '更新成功！即将运行新版本。')
            root.destroy()
            subprocess.Popen(f'Excel Tools v{NEW}.exe')
            sys.exit()

def open_file():
    file_path = filedialog.askopenfilename()
    if check_path(file_path):
        var1.set(file_path)

def open_file2(): # 选择文件
    file_path = filedialog.askopenfilename()
    if check_path(file_path):
        var8.set(file_path)

def open_file3(): # 选择文件夹
    file_path = filedialog.askdirectory()
    if os.path.isdir(file_path):
        var8.set(file_path)

def extract():
    if not os.path.isfile(var1.get()):
        messagebox.showerror('错误', '选择的文件不存在！')
        record.append(f'{get_time()} Extract Failed - The selected file does not exist')
        return False
    t1.delete("1.0", tk.END)
    result = Extract.main(var2.get(), var3.get(), var1.get())
    if type(result[0]) == list:
        for i in result:
            t1.insert(tk.END, "\n".join(i))
            t1.insert(tk.END, "\n")
    else:
        t1.insert(tk.END, "\n".join(result))
    if 'None' in result:
        messagebox.showwarning('警告', '提取出空值，请检查选择的文件以及输入的单元格是否正确！')
    record.append(f'{get_time()} Extract Successfully')
    return True

def search():
    try:
        response = requests.get("https://www.google.com/")
    except:
        messagebox.showerror(title="错误", message="无法连接至谷歌服务器，请重试。")
        record.append(f'{get_time()} Search Failed')
        return False

    error_img = GoogleSearch.main(t1.get("1.0", tk.END).split("\n"), set_value['filter'], set_value['path'])

    if error_img: # 有图片没搜到
        messagebox.showwarning('警告', f'搜索完成！但有{error_img}个件号图片无法被搜到，请检查件号是否有误！')
        record.append(f'{get_time()} Search Successfully but {error_img} Not Found')

        '''if messagebox.askyesno('提示', '是否继续搜索英文名？'):
            for content in t1.get("1.0", tk.END).split("\n"):
                if content == '' or content == '\n':
                    continue
                GoogleSearch.searchName(content)
            messagebox.showinfo('提示', '英文名搜索完成。')
            record.append(f'{get_time()} English_name Search Successfully')'''

        return False
    # 正常状态
    record.append(f'{get_time()} Search Successfully')
    messagebox.showinfo('提示', '搜索完成！')

    '''if messagebox.askyesno("提示", "搜索完成！是否继续搜索英文名？"):
        for content in t1.get("1.0", tk.END).split("\n"):
            if content == '' or content == '\n':
                continue
            GoogleSearch.searchName(content)
        messagebox.showinfo('提示', '英文名搜索完成。')
        record.append(f'{get_time()} English_name Search Successfully')'''
    return True

def insert():
    try:
        num = int(var3.get()[1:]) - int(var2.get()[1:]) + 1 # 待插入的图片数量
        error_insert = Insert.main([var4.get()[0], int(var4.get()[1:])], num, var1.get(), set_value['path'])
    except FileNotFoundError:
        messagebox.showerror('错误', '未找到Excel文件或图片文件！')
        record.append(f'{get_time()} Insert Failed')
        return False
    '''if set_value['auto_delete']:  # 删除img文件夹
        shutil.rmtree(set_value['path'])'''
    if error_insert and error_insert != 'STOP':
        messagebox.showwarning('警告', f'有{error_insert}个图片插入失败！其余插入成功。\n注：请检查插入图片失败的件号是否正确！')
        record.append(f'{get_time()} Insert Successfully but {error_insert} Failed')
        return False
    elif error_insert == 'STOP':
        return False
    messagebox.showinfo("提示", "插入完成！")
    record.append(f'{get_time()} Insert Successfully')

def move_key():
    if not os.path.isfile(var1.get()):
        messagebox.showerror('错误', '选择的文件不存在！')
        record.append(f'{get_time()} Move keys Failed - The selected file does not exist')
        return False
    if set_value.get('auto_backup'):
        shutil.copyfile(var1.get(), 'backup.xlsx')
        print(blue_text('Excel文件已成功备份到程序目录下的 "backup.xlsx"！'))
    MoveKey.main(var1.get(), var5.get(), var6.get(), var7.get())
    messagebox.showinfo('提示', '移动完成！')
    record.append(f'{get_time()} Move keys Successfully')

def excel_search():
    data_list = t1.get('1.0', tk.END).split('\n')
    for data in data_list:
        if data.strip() == '':
            continue
        elif os.path.isfile(var8.get()):
            if not ExcelSearch.find_data_in_single_excel(var8.get(), data):
                record.append(f'{get_time()} Single Excel Search Failed')
                return False
            record.append(f'{get_time()} Single Excel Search Successfully')
        elif os.path.isdir(var8.get()):
            ExcelSearch.find_data_in_multiple_excel(var8.get(), data)
            record.append(f'{get_time()} Multiple Excel Search Successfully')
        print(f'---------------以上为 "{data}" 的搜索结果---------------')
    messagebox.showinfo('提示', '搜索完成！请到命令行查看搜索结果。')

def excel_check():
    if os.path.isdir(var8.get()):
        ExcelCheck.checkSum(var8.get())
        print(f'---------------以上为 "SUM函数" 的检查结果---------------')
        messagebox.showinfo('提示', '检查完成！请到命令行查看检查结果。')
        record.append(f'{get_time()} CheckSum Successfully')
    else:
        messagebox.showwarning('警告', '请选择一个文件夹而不是一个文件。')
        record.append(f'{get_time()} CheckSum Failed - Selected a file instead of a folder')

def exit_():
    root.destroy()

    save()
    postLog()
    sys.exit()

def about():
    if NEW == None:
        messagebox.showinfo("关于", f'当前版本：v{VERSION}\n最新版本：未知\n（当前自动更新已关闭，请点击 "关于->检查更新" 手动获取最新版本）\n作者：Sam')
    elif NEW == 0:
        messagebox.showinfo("关于", f'当前版本：v{VERSION}\n最新版本：未知（获取更新失败，请点击 "关于->检查更新" 重新获取更新）\n作者：Sam')
    else:
        messagebox.showinfo("关于", f"当前版本：v{VERSION}\n最新版本：v{NEW}\n作者：Sam")

def postLog():
    # print(yellow_text('正在上传日志，请不要关闭程序！'))
    # 先处理文件
    with open(rf'{LOG_DIR}\operation.log', 'a') as f:
        record.append(f'{get_time()}!')
        f.write(json.dumps(record)+'\n')

    try:
        # 网站已废弃，暂时关闭postlog功能，何时恢复未知...（检查注释）
        return False
        response = requests.get(r'http://techxi.us.kg/post', timeout=3)
        if response.status_code != 200:
            return False
    except:
        return False

    if multi_stream.isWrite:
        with open(rf'{LOG_DIR}\error.log', 'a') as f:
            f.write(f'----------Above {get_time()}----------\n\n')
        with open(rf'{LOG_DIR}\error.log', 'r') as f:
            data = f.read()
            if PostData.postData(data, 'errorLog', gethostname()):
                print(blue_text('日志上传成功！'))
            else:
                print(red_text('日志上传失败！'))
                sys.exit()

    with open(rf'{LOG_DIR}\operation.log', 'r') as f:
        state = PostData.postData(f.read(), 'recordLog', gethostname())
        if not multi_stream.isWrite:
            if state:
                print(blue_text('日志上传成功！'))
            else:
                print(red_text('日志上传失败！'))

def easydo(): # 一键操作
    if not extract():
        return
    if search() == 0:
        return
    insert()

def settings():
    record.append(f'{get_time()} Open Settings')
    save()
    root.destroy()
    Settings.main()
    main(False)

def load(): # 加载上次填充的内容
    global data
    data = {}
    if os.path.exists(CONFIG_FILE1):
        with open(CONFIG_FILE1, 'r') as f:
            try:
                data = json.load(f)
                return True
            except json.decoder.JSONDecodeError:
                data = {}
                with open(CONFIG_FILE1, 'w'): pass
    else:
        with open(CONFIG_FILE1, 'w') as f: pass
        return False

def save():
    data = {}
    data['file'] = var1.get()
    data['start1'] = var2.get()
    data['end1'] = var3.get()
    data['start2'] = var4.get()
    data['move_start'] = var5.get()
    data['move_end'] = var6.get()
    data['move_target'] = var7.get()
    data['search_file'] = var8.get()
    data['isTop'] = top

    with open(CONFIG_FILE1, 'w') as f:
        json.dump(data, f)

def top_switch():
    global top
    top = not top
    root.attributes("-topmost", top)

def on_drop(event):
    if event.data.count('{') == 1:
        path = event.data[1:-1]
    elif event.data.count('{') == 0:
        path = event.data
    elif event.data.count('{') > 1:
        return
    if check_path(path):
        var1.set(path)

def on_drop2(event):
    if event.data.count('{') == 1:
        path = event.data[1:-1]
    elif event.data.count('{') == 0:
        path = event.data
    elif event.data.count('{') > 1:
        return
    if check_path(path):
        var8.set(path)
    elif os.path.isdir(path):
        var8.set(path)

def check_path(path):
    if path[-4:] == '.xls':
        messagebox.showerror('错误', '暂不支持对.xls文件操作，请转换为.xlsx版本')
        return False
    elif path[-5:] != '.xlsx':
        return False
    return True

def rc_popup(event): # 右键弹出粘贴菜单
    rc_menu.post(event.x_root, event.y_root)

def rc_paste(event=None):
    t1.event_generate('<<Paste>>')

def rc_copy(event=None):
    t1.event_generate('<<Copy>>')

def red_text(text):
    return Fore.RED+Style.BRIGHT+text+Style.RESET_ALL

def blue_text(text):
    return Fore.BLUE+Style.BRIGHT+text+Style.RESET_ALL

def green_text(text):
    return Fore.GREEN+Style.BRIGHT+text+Style.RESET_ALL

def yellow_text(text):
    return Fore.YELLOW+Style.BRIGHT+text+Style.RESET_ALL

def main(check=True): # check为是否检查更新以及是否输出提示文本
    global root, t1, var1, var2, var3, var4, var5, var6, var7, var8, top, NEW, id, set_value, data, t1, rc_menu

    deleteOld()

    set_value = Settings.load()
    if len(set_value) != Settings.set_value_len:
        Settings.save(True)
        set_value = Settings.load()

    root = tkinterdnd2.Tk()
    root.title(f"Excel Tools by Sam v{VERSION}")
    root.geometry("515x405+400+200")
    root.resizable(width=False, height=False)

    var1 = tk.StringVar() # 文件
    var2 = tk.StringVar() # 开始
    var3 = tk.StringVar() # 结束
    var4 = tk.StringVar() # 插入开始
    var5 = tk.StringVar() # 移动件号开始
    var6 = tk.StringVar() # 移动件号结束
    var7 = tk.StringVar() # 移动件号目标
    var8 = tk.StringVar() # 数据搜索目标文件

    if load():
        if type(data) != dict:
            data = {}
        try:
            var1.set(data['file'])
            var2.set(data['start1'])
            var3.set(data['end1'])
            var4.set(data['start2'])
            var5.set(data['move_start'])
            var6.set(data['move_end'])
            var7.set(data['move_target'])
            var8.set(data['search_file'])
            top = data['isTop']
        except KeyError:
            pass

    if top:
        root.attributes("-topmost", top)

    f1 = tk.Frame(root)
    f1.grid(row=0, column=0)

    lb1 = tk.Label(f1, text="选择文件:", fg="red")
    lb1.grid(row=0, column=0, sticky=tk.W, padx=5)
    et1 = tk.Entry(f1, textvariable=var1, width=45, state="disabled")
    et1.grid(row=0, column=1, padx=5, sticky=tk.W)
    root.drop_target_register(tkinterdnd2.DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)
    bt1 = tk.Button(f1, text="选择", command=open_file)
    bt1.grid(row=0, column=2, padx=5, sticky=tk.W)

    f2 = tk.Frame(root)
    f2.grid(row=1, column=0, sticky=tk.NW)
    f3 = tk.Frame(f2)
    f3.grid(row=0, column=1, sticky=tk.NW)

    # 提取文字
    lf1 = tk.LabelFrame(f2, text="提取文字")
    lf1.grid(row=0, column=0, padx=5, sticky=tk.NW)
    lb2 = tk.Label(lf1, text="开始：", fg="red")
    lb2.grid(row=0, column=0, sticky=tk.E)
    et2 = tk.Entry(lf1, textvariable=var2, width=7)
    et2.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
    lb3 = tk.Label(lf1, text="结束：", fg="red")
    lb3.grid(row=1, column=0, sticky=tk.E)
    et3 = tk.Entry(lf1, textvariable=var3, width=7)
    et3.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
    t1 = tk.Text(lf1, width=22, height=12)
    scroll = tk.Scrollbar(lf1, orient=tk.VERTICAL)
    scroll.grid(row=2, column=1, sticky='nse')
    scroll.config(command=t1.yview)
    t1.insert("1.0", "此处为待搜索的内容")
    t1.config(yscrollcommand=scroll.set)
    t1.grid(row=2, column=0, columnspan=2, padx=10)
    t1.bind('<Button-3>', rc_popup) # 绑定右键菜单

    bt2 = tk.Button(lf1, text="提取", command=extract)
    bt2.grid(row=3, column=0, columnspan=2, pady=5)

    # 谷歌搜索
    lf2 = tk.LabelFrame(f3, text="谷歌搜索")
    lf2.grid(row=0, column=0, padx=5, sticky=tk.N)
    lb4 = tk.Label(lf2, text="进度请查看命令行")
    lb4.grid(row=0, column=0, padx=5)
    bt3 = tk.Button(lf2, text="搜索", command=search)
    bt3.grid(row=1, column=0, pady=5)

    # 插入图片
    lf3 = tk.LabelFrame(f3, text="插入图片")
    lf3.grid(row=1, column=0, padx=5, sticky=tk.N)
    lb5 = tk.Label(lf3, text="开始：", fg="red")
    lb5.grid(row=0, column=0, sticky=tk.E)
    et4 = tk.Entry(lf3, textvariable=var4, width=7)
    et4.grid(row=0, column=1, padx=8, pady=5, sticky=tk.W)
    bt4 = tk.Button(lf3, text="插入", command=insert)
    bt4.grid(row=1, column=0, pady=5, columnspan=2)

    # 一键操作
    lf4 = tk.LabelFrame(f3, text="一键操作")
    lf4.grid(row=2, column=0, padx=5, sticky=tk.N)
    bt5 = tk.Button(lf4, text="一键操作", command=easydo)
    bt5.grid(padx=27, pady=5)

    # 移动件号
    lf5 = tk.LabelFrame(f3, text="移动件号")
    lf5.grid(row=0, column=1, padx=5, sticky=tk.N, rowspan=2)
    lb6 = tk.Label(lf5, text="开始：", fg="red")
    lb6.grid(row=0, column=0, sticky=tk.E)
    et5 = tk.Entry(lf5, textvariable=var5, width=7)
    et5.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
    lb7 = tk.Label(lf5, text="结束：", fg="red")
    lb7.grid(row=1, column=0, sticky=tk.E)
    et6 = tk.Entry(lf5, textvariable=var6, width=7)
    et6.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
    lb8 = tk.Label(lf5, text="目标：", fg="red")
    lb8.grid(row=2, column=0, sticky=tk.E)
    et7 = tk.Entry(lf5, textvariable=var7, width=7)
    et7.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
    bt6 = tk.Button(lf5, text="移动", command=move_key)
    bt6.grid(row=3, column=0, padx=5, pady=5, columnspan=2)

    # 数据搜索&函数检查
    lf6 = tk.LabelFrame(root, text='数据搜索&函数检查(在上方文本框输入待搜索的内容)')
    lf6.grid(row=2, column=0, padx=5, sticky=tk.NW)
    lf6.drop_target_register(tkinterdnd2.DND_FILES)
    lf6.dnd_bind('<<Drop>>', on_drop2)
    lb9 = tk.Label(lf6, text="目标文件：", fg="red")
    lb9.grid(row=0, column=0, sticky=tk.W, padx=5)
    et8 = tk.Entry(lf6, textvariable=var8, width=45, state="disabled")
    et8.grid(row=0, column=1, padx=5, sticky=tk.W)
    f4 = tk.Frame(lf6)
    f4.grid(row=1, column=0, columnspan=3, padx=10)
    bt7 = tk.Button(f4, text="选择文件", command=open_file2)
    bt7.grid(row=0, column=0, padx=5, pady=5, sticky=tk.N)
    bt8 = tk.Button(f4, text="选择文件夹", command=open_file3)
    bt8.grid(row=0, column=1, padx=5, pady=5, sticky=tk.N)
    bt9 = tk.Button(f4, text='搜索', command=excel_search)
    bt9.grid(row=0, column=2, padx=5, pady=5, sticky=tk.N)
    bt10 = tk.Button(f4, text='检查', command=excel_check)
    bt10.grid(row=0, column=3, padx=5, pady=5, sticky=tk.N)

    # 右键Menu
    rc_menu = tk.Menu(root, tearoff=0)
    rc_menu.add_command(label='复制', command=rc_copy)
    rc_menu.add_command(label='粘贴', command=rc_paste)

    # 顶部Menu
    menubar = tk.Menu(root)

    about_menu = tk.Menu(menubar, tearoff=False)
    about_menu.add_command(label="关于", command=about)
    about_menu.add_command(label='检查更新', command=update)
    about_menu.add_command(label='更新公告', command=lambda: messagebox.showinfo('更新内容', update_content))
    about_menu.add_command(label='官网', command=lambda:webbrowser.open(r'https://aqiulawrence.github.io/'))

    set_menu = tk.Menu(menubar, tearoff=False)
    set_menu.add_command(label="设置", command=settings)

    window_menu = tk.Menu(menubar, tearoff=False)
    window_menu.add_command(label='切换置顶', command=top_switch)

    exit_menu = tk.Menu(menubar, tearoff=False)
    exit_menu.add_command(label='退出', command=exit_)

    menubar.add_cascade(label="关于", menu=about_menu)
    menubar.add_cascade(label="设置", menu=set_menu)
    menubar.add_cascade(label='窗口', menu=window_menu)
    menubar.add_cascade(label="退出", menu=exit_menu)
    root.config(menu=menubar)

    if check and set_value['auto_update']: # check在第一次进入时为True，从设置进入时为False
        thread = Thread(target=update, args=(True,))
        thread.start()

    if check:
        if ctypes.windll.shell32.IsUserAnAdmin():
            messagebox.showinfo('提示', '若非特殊情况，请勿使用管理员身份运行此程序！')
        else:
            print(blue_text('[提示]文件无需手动选择，可拖拽进入窗口！'))

    if isFirstOpen():
        # Settings.save(True) # 防止设置不兼容，某些版本更新需要设置，属于兼容部分代码
        messagebox.showinfo('更新内容', '注：请先关闭此更新公告再使用主程序！！\n'+update_content)

    root.mainloop()

    # 主界面关闭后的处理，记得还要同步到 exit_() 函数上！
    save()
    postLog()
    sys.exit()

if __name__ == "__main__":
    try:
        import pyi_splash
        pyi_splash.update_text('Loading...')
        pyi_splash.close()
    except: pass

    try:
        main()
    except SystemExit: pass
    except KeyboardInterrupt: pass
    except Exception:
        traceback.print_exc() # stderr
        input()