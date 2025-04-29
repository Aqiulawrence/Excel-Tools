import tkinter as tk
import os
import json
import time

from tkinter import filedialog, messagebox

def get_time():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

CONFIG_DIR = rf'.\configs'
CONFIG_FILE = rf'{CONFIG_DIR}\cf2.json'

class AppSettings:
    CONFIG_DIR = "./configs"
    CONFIG_FILE = "./configs/app_settings.json"

    # 定义所有设置项及其属性
    SETTINGS_SCHEMA = {
        'auto_update': {
            'type': 'checkbox',
            'label': '启用自动更新',
            'default': 1
        },
        'filter': {
            'type': 'checkbox',
            'label': '搜图启用网站筛选',
            'default': 1
        },
        'auto_backup': {
            'type': 'checkbox',
            'label': '移动前自动备份',
            'default': 1
        },
        'max_workers': {
            'type': 'entry',
            'label': '搜索最大进程数：',
            'default': 8,
            'validate': 'int'  # 添加验证类型
        }
    }

    def __init__(self):
        self._settings_window = None
        self._current_settings = self._load_default_settings()
        self._load_saved_settings()

    @property
    def settings(self):
        """获取当前所有设置的副本"""
        return self._current_settings.copy()

    def get(self, name: str, default=None):
        """获取单个设置值"""
        return self._current_settings.get(name, default)

    def show(self, parent):
        """显示设置对话框"""
        if self._settings_window and self._settings_window.winfo_exists(): # 双重判断
            self._settings_window.lift() # 跳到前面
            return

        self._settings_window = tk.Toplevel(parent)
        self._settings_window.title("设置")
        self._settings_window.geometry("450x300")
        self._settings_window.resizable(False, False)

        self._create_settings_ui()
        self._settings_window.protocol("WM_DELETE_WINDOW", self._close_settings)

    def _load_default_settings(self):
        return {name: config['default'] for name, config in self.SETTINGS_SCHEMA.items()}

    def _load_saved_settings(self):
        try:
            if not os.path.exists(self.CONFIG_DIR):
                os.makedirs(self.CONFIG_DIR)

            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r') as f:
                    saved_settings = json.load(f)
                    for name, value in saved_settings.items():
                        if name in self._current_settings:
                            self._current_settings[name] = value
        except Exception as e:
            print(f"加载设置失败: {e}")

    def _save_settings(self) -> bool:
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(self._current_settings, f, indent=2)
            return True
        except Exception as e:
            print(f"保存设置失败: {e}")
            return False

    def _create_settings_ui(self):
        """创建设置界面UI"""
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
                messagebox.showinfo("成功", "设置已保存！")
                self._close_settings()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")

    def _close_settings(self):
        if self._settings_window:
            self._settings_window.destroy()
            self._settings_window = None

if __name__ == '__main__':
    root = tk.Tk()
    app_settings = AppSettings()
    app_settings.show(root)
    root.mainloop()