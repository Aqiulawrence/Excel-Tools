import ExcelExtract
import ExcelInsert
import ExcelSearch
import GoogleSearch
import CheckUpdate
import MoveKey
from Logger import AppLogger, log

import tkinter as tk
import shutil
import traceback
import sys
import requests
import webbrowser
import os
import subprocess
import tkinterdnd2
import json
import time
import re
import ctypes
import multiprocessing

from colorama import Fore, Style
from threading import Thread
from tkinter import messagebox, filedialog
from easygui import exceptionbox

VERSION = "2.2" # 当前版本
NEW = None # 最新版本
UPDATE_CONTENT = '''1. 搜图部分改用多线程，大幅提高了搜图效率。
2. 优化了插入图片部分代码，大幅提高了插入图片的效率。'''

top = False # 窗口是否置顶
CONFIG_DIR = './config'
USERDATA_FILE = f'{CONFIG_DIR}/userdata.json'
if not os.path.exists(CONFIG_DIR):
    os.makedirs(CONFIG_DIR)

# 初始化日志系统
app_logger = AppLogger()

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

# 设置多重错误流
err_log = open(f'{CONFIG_DIR}/error.log', 'a')
multi_stream = MultiStream(sys.stderr, err_log)
sys.stderr = multi_stream

class SettingsApp:
    CONFIG_FILE = f"{CONFIG_DIR}/config.json"

    # 定义所有设置项及其属性
    SETTINGS_SCHEMA = {
        'auto_update': {
            'type': 'checkbox',
            'label': '启用自动更新：',
            'default': 1
        },
        'filter': {
            'type': 'checkbox',
            'label': '搜图启用网站筛选：',
            'default': 1
        },
        'display_failure': {
            'type': 'checkbox',
            'label': '搜图显示失败信息：',
            'default': 0
        },
        'excel_search_max_workers': {
            'type': 'entry',
            'label': '搜索最大进程数：',
            'default': 8,
            'validate': 'int'  # 验证类型
        },
        'google_search_max_workers': {
            'type': 'entry',
            'label': '搜图最大线程数：',
            'default': 20,
            'validate': 'int'  # 验证类型
        }
    }

    def __init__(self):
        self._settings_window = None
        self._current_settings = self._load_default_settings()
        self._load_saved_settings()

    @property
    def settings(self):
        return self._current_settings.copy()

    def get(self, name: str, default=None):
        return self._current_settings.get(name, default)

    def show(self, parent):
        if self._settings_window and self._settings_window.winfo_exists(): # 双重判断
            self._settings_window.lift() # 跳到前面
            return

        self._settings_window = tk.Toplevel(parent)
        self._settings_window.title("设置")
        self._settings_window.geometry("260x245+400+200")
        self._settings_window.resizable(False, False)
        self._settings_window.attributes("-topmost", True)

        self._create_settings_ui()
        self._settings_window.protocol("WM_DELETE_WINDOW", self._close_settings)

    def _load_default_settings(self):
        return {name: config['default'] for name, config in self.SETTINGS_SCHEMA.items()}

    def _load_saved_settings(self):
        try:
            if not os.path.exists(CONFIG_DIR):
                os.makedirs(CONFIG_DIR)

            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r') as f:
                    saved_settings = json.load(f)
                    for name, value in saved_settings.items():
                        if name in self._current_settings:
                            self._current_settings[name] = value
        except Exception as e:
            print(f"加载设置失败: {e}")

    def _save_from_ui(self):
        try:
            for name, var in self._setting_widgets.items():
                config = self.SETTINGS_SCHEMA[name]

                if config['type'] == 'checkbox':
                    self._current_settings[name] = var.get()
                elif config['type'] == 'entry':
                    value = var.get()
                    if config.get('validate') == 'int':
                        value = int(value)
                    self._current_settings[name] = value

            if self._save_settings():
                self._close_settings()

        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")

    def _save_settings(self) -> bool:
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(self._current_settings, f, indent=2)
            return True
        except Exception as e:
            print(f"保存设置失败: {e}")
            return False

    def _create_settings_ui(self):
        main_frame = tk.Frame(self._settings_window, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建设置项控件
        self._setting_widgets = {}
        for row, (name, config) in enumerate(self.SETTINGS_SCHEMA.items()):
            frame = tk.Frame(main_frame)
            frame.pack(fill=tk.X, pady=5)

            label = tk.Label(frame, text=config['label'], width=20, anchor='w')
            label.pack(side=tk.LEFT)

            if config['type'] == 'checkbox':
                var = tk.IntVar(value=self.get(name))
                widget = tk.Checkbutton(frame, variable=var)
                self._setting_widgets[name] = var
            elif config['type'] == 'entry':
                var = tk.StringVar(value=str(self.get(name)))
                widget = tk.Entry(frame, textvariable=var, width=10)
                self._setting_widgets[name] = var
            widget.pack(side=tk.LEFT, padx=5)

        # 添加按钮
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        tk.Button(
            btn_frame, text="保存",
            command=self._save_from_ui, width=10
        ).pack(side=tk.RIGHT, padx=5)

        tk.Button(
            btn_frame, text="取消",
            command=self._close_settings, width=10
        ).pack(side=tk.RIGHT)

    def _close_settings(self):
        if self._settings_window:
            self._settings_window.destroy()
            self._settings_window = None

app = SettingsApp()

def get_time():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

def isFirstOpen(): # 获取是否为第一次打开程序
    path = f'{CONFIG_DIR}/record.json' # 这个文件只记录打开过的版本号
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

def start_update():
    thread = Thread(target=update)
    thread.start()

def update(auto=False):
    global NEW
    if auto == False:
        print('检查更新中...')
    state = CheckUpdate.update(VERSION) # 返回值：0表示获取更新失败（没开VPN）、1表示新版本下载完了、版本号、False表示获取更新失败
    if state == 1: # 更新成功
        messagebox.showinfo('提示', '更新成功！请您运行新版本。')
    else:
        NEW = state

@log
def extract():
    if not os.path.isfile(var1.get()):
        messagebox.showerror('错误', '选择的文件不存在！')
        return False
    t1.delete("1.0", tk.END)
    result = ExcelExtract.main(var2.get(), var3.get(), var1.get())

    if type(result[0]) == list:
        for i in result:
            t1.insert(tk.END, "\n".join(i))
            t1.insert(tk.END, "\n")
    else:
        t1.insert(tk.END, "\n".join(result))
    if 'None' in result:
        messagebox.showwarning('警告', '提取出空值，请检查选择的文件以及输入的单元格是否正确！')

    return True

@log
def search():
    try: # 先尝试连接
        response = requests.get("https://www.google.com/")
    except:
        messagebox.showerror(title="错误", message="无法连接至谷歌服务器，请重试。")
        return False

    start = time.time()
    error_count = GoogleSearch.main(t1.get("1.0", tk.END).split("\n"), app.settings['google_search_max_workers'], app.settings['filter'], app.settings['display_failure'])
    print(f'用时：{time.time()-start:.2f}s')

    if error_count: # 有图片没搜到
        messagebox.showwarning('警告', f'搜索完成！但有{error_count}个件号图片无法被搜到，请检查件号是否有误！')
        return False

    # 正常状态
    messagebox.showinfo('提示', '搜索完成！')
    return True

@log
def insert():
    try:
        num = int(var3.get()[1:]) - int(var2.get()[1:]) + 1 # 待插入的图片数量
        error_count = ExcelInsert.main([var4.get()[0], int(var4.get()[1:])], num, var1.get())
    except FileNotFoundError:
        messagebox.showerror('错误', '未找到Excel文件！')
        return False

    if error_count:
        messagebox.showwarning('警告', f'有{error_count}个图片插入失败！其余插入成功。')
        return False
    messagebox.showinfo("提示", "插入完成！")

@log
def move_key():
    if not os.path.isfile(var1.get()):
        messagebox.showerror('错误', '选择的文件不存在！')
        return False
    MoveKey.main(var1.get(), var5.get(), var6.get(), var7.get())
    messagebox.showinfo('提示', '移动完成！')

@log
def excel_search():
    search_term = t1.get('1.0', tk.END).split('\n')
    for term in search_term:
        start = time.time()
        if term.strip() == '':
            continue

        elif os.path.isfile(var8.get()): # 搜索单个文件
            ExcelSearch.single_search(var8.get(), term)

        elif os.path.isdir(var8.get()): # 搜索多个文件
            ExcelSearch.batch_search(var8.get(), term, app.settings['excel_search_max_workers'])
            if os.path.exists('./temp'):
                shutil.rmtree('./temp')

        end = time.time()
        print(f'---------------以上为 "{term}" 的搜索结果，用时：{end-start:.2f}s---------------')
    messagebox.showinfo('提示', '搜索完成！请到命令行查看搜索结果。')

def about(): # 关于
    if NEW:
        messagebox.showinfo("关于", f"当前版本：v{VERSION}\n最新版本：v{NEW}\n作者QQ：2418131303\n有bug请联系作者！")
    else:
        messagebox.showinfo("关于", f"当前版本：v{VERSION}\n最新版本：未知\n作者QQ：2418131303\n有bug请联系作者！")

@log
def easydo(): # 一键操作
    if not extract():
        return
    if search() == 0:
        return
    insert()

def load(): # 加载上次填充的内容
    global data
    data = {}
    if os.path.exists(USERDATA_FILE):
        with open(USERDATA_FILE, 'r') as f:
            try:
                data = json.load(f)
                return True
            except json.decoder.JSONDecodeError:
                data = {}
                with open(USERDATA_FILE, 'w'): pass
    else:
        with open(USERDATA_FILE, 'w') as f: pass
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

    with open(USERDATA_FILE, 'w') as f:
        json.dump(data, f)

def top_switch():
    global top
    top = not top
    root.attributes("-topmost", top)

def check_path(path):
    if path.lower().endswith(('.xlsx', '.xls')):
        return True
    return False

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

def main():
    global root, t1, var1, var2, var3, var4, var5, var6, var7, var8, top, NEW, data, t1, rc_menu

    deleteOld()

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
    bt_open = tk.Button(f1, text='打开', command=lambda: os.startfile(var1.get()))
    bt_open.grid(row=0, column=3, padx=5, sticky=tk.W)

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
    lf6 = tk.LabelFrame(root, text='数据搜索(在上方文本框输入待搜索的内容)')
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

    # 右键Menu
    rc_menu = tk.Menu(root, tearoff=0)
    rc_menu.add_command(label='复制', command=rc_copy)
    rc_menu.add_command(label='粘贴', command=rc_paste)

    # 创建菜单栏
    menubar = tk.Menu(root)
    root.config(menu=menubar)

    # 创建下拉菜单和菜单项
    about_menu = tk.Menu(menubar, tearoff=False)
    menubar.add_cascade(label="关于", menu=about_menu)
    about_menu.add_command(label="关于", command=about)
    about_menu.add_command(label='官网', command=lambda: webbrowser.open(r'https://aqiulawrence.github.io/'))
    about_menu.add_command(label='检查更新', command=start_update)
    about_menu.add_command(label='更新公告', command=lambda: messagebox.showinfo('更新内容', UPDATE_CONTENT))

    set_menu = tk.Menu(menubar, tearoff=False)
    menubar.add_cascade(label="设置", menu=set_menu)
    set_menu.add_command(label="设置", command=lambda: app.show(root))
    
    window_menu = tk.Menu(menubar, tearoff=False)
    menubar.add_cascade(label='窗口', menu=window_menu)
    var_topmost = tk.BooleanVar(value=top) # 这里懒得改了直接用top变量
    window_menu.add_checkbutton(label='总在最前', command=top_switch, variable=var_topmost)
    window_menu.add_command(label="保存数据", command=save)

    if app.settings['auto_update']:
        thread = Thread(target=update, args=(True,))
        thread.start()

    if ctypes.windll.shell32.IsUserAnAdmin():
        messagebox.showinfo('提示', '以管理员身份运行此程序将无法得到彩色输出！')
    else:
        print(blue_text('[提示]文件无需手动选择，可拖拽进入窗口！'))

    if isFirstOpen():
        messagebox.showinfo('更新内容', UPDATE_CONTENT)

    root.mainloop()

    save()
    if multi_stream.isWrite:
        with open(f'{CONFIG_DIR}/error.log', 'a') as f:
            f.write(f'----------Above {get_time()}----------\n\n')
    sys.exit()

if __name__ == "__main__":
    try:
        import pyi_splash
        pyi_splash.update_text('Loading...')
        pyi_splash.close()
    except: pass

    try:
        multiprocessing.freeze_support()
        main()
    except SystemExit: pass
    except KeyboardInterrupt: pass
    except Exception:
        traceback.print_exc() # stderr
        input()