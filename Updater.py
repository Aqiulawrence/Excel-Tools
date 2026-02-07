import sys
import os
import requests
import hashlib
import shutil
import time
import subprocess
import tempfile
from pathlib import Path
import zipfile
import winreg
from time import sleep

VERSION = "v2026.02"
SERVER_URL = "http://www.wublog.site/update"
APP_NAME = "Excel-Tools"
UPDATE_ZIP = "update.zip"
exes = ["ExcelCompare.exe", "ExcelSearch.exe", "ImageSearch.exe"]


def add_startup():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0,
                         winreg.KEY_SET_VALUE)
    winreg.SetValueEx(key, "Updater", 0, winreg.REG_SZ, f'"{os.path.abspath(sys.argv[0])}"')
    winreg.CloseKey(key)

def check_update():
    try:
        resp = requests.post(SERVER_URL, json={
            "version": VERSION,
            "device": os.environ.get("COMPUTERNAME", "Unknown"),
            "information": "Get Updates"
        }, timeout=30)
        if resp.status_code == 200:
            return resp.json()  # 返回 {"version": "...", "url": "...", "hash": "..."}
    except:
        pass
    return None


def download_file(url, path):
    resp = requests.get(url, stream=True, timeout=60)
    with open(path, 'wb') as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)


def verify_hash(filepath, expected_hash):
    with open(filepath, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest() == expected_hash


def is_process_running(exe_name):
    try:
        output = subprocess.check_output('tasklist /fi "imagename eq %s"' % exe_name, shell=True)
        return exe_name.encode() in output
    except:
        return False


def wait_for_close():
    while any(is_process_running(exe) for exe in exes):
        time.sleep(5)


def run_new_updater(new_updater_path, temp_dir):
    app_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    cmd = [new_updater_path, "--apply-update", temp_dir, app_dir]

    # 启动新的updater
    subprocess.Popen(cmd)
    sys.exit(0)


def apply_update():
    if len(sys.argv) < 4 or sys.argv[1] != "--apply-update":
        return False

    temp_dir = sys.argv[2]
    target_dir = sys.argv[3]

    # 等待其他进程关闭
    wait_for_close()

    # 替换文件
    target_path = Path(target_dir)
    for item in Path(temp_dir).rglob("*"):
        if item.is_file():
            rel_path = item.relative_to(temp_dir)
            target = target_path / rel_path
            target.parent.mkdir(parents=True, exist_ok=True)

            # 尝试复制文件
            for _ in range(3600):  # 最多尝试5小时
                try:
                    shutil.copy2(item, target)
                    break
                except:
                    time.sleep(5)

    requests.post(SERVER_URL, json={
        "version": VERSION,
        "device": os.environ.get("COMPUTERNAME", "Unknown"),
        "information": "Update Successfully"
    }, timeout=30)

    # 删除临时目录
    shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    # 如果收到--apply-update参数，执行更新替换
    if len(sys.argv) > 1 and sys.argv[1] == "--apply-update":
        apply_update()
        sys.exit()

    # 检查更新
    update_info = check_update()
    if not update_info or update_info["version"] <= VERSION:
        sys.exit(0)

    # 下载并校验
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, UPDATE_ZIP)
    download_file(update_info["url"], zip_path)
    if not verify_hash(zip_path, update_info["hash"]):
        sys.exit(1)

    # 解压
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(temp_dir)
    os.remove(zip_path)

    new_updater_path = os.path.join(temp_dir, "Updater.exe")
    run_new_updater(new_updater_path, temp_dir)


if __name__ == "__main__":
    add_startup() # 添加启动项
    sleep(60)
    main()
