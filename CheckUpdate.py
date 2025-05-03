import requests
import os
import time

from colorama import Fore, Style
from tqdm import tqdm

def red_text(text):
    return Fore.RED+Style.BRIGHT+text+Style.RESET_ALL

def blue_text(text):
    return Fore.BLUE+Style.BRIGHT+text+Style.RESET_ALL

owner = "Aqiulawrence"
repo = "Excel-Tools"

def replace(filename):
    v_index = filename.find('v')
    if v_index == -1:
        return filename.replace('.', ' ')

    before_v = filename[:v_index].replace('.', ' ')
    after_v = filename[v_index:]
    return before_v + after_v

def update(ver, owner=owner, repo=repo, output_dir="./", token=None):
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/vnd.github.v3+json"
    }
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    if token:
        headers["Authorization"] = f"token {token}"

    try:
        response = requests.get(api_url, headers=headers)
        release = response.json()
    except Exception as e:
        print(f'获取更新失败：{e}')
        return False

    try:
        new = release['tag_name'][1:] # github release为空
    except:
        print('找不到可用更新。')
        return 0

    if new <= ver:
        print('当前为最新版本。')
        return new

    print(f'发现新版本v{new}，正在进行后台更新！')

    max_retries = 3
    retry_count = 0
    base_delay = 12

    while retry_count < max_retries:
        try:
            # 查找.exe文件
            exe_assets = [asset for asset in release.get("assets", [])
                          if asset["name"].lower().endswith(".exe")]

            # 下载第一个找到的.exe文件(如果有多个，可以修改这里)
            exe_asset = exe_assets[0]
            download_url = exe_asset["browser_download_url"]
            file_name = exe_asset["name"]
            file_name = replace(file_name) # 替换点为空格
            file_path = os.path.join(output_dir, file_name)

            response = requests.get(download_url, headers=headers, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            with open(file_path, 'wb') as file, tqdm(
                    desc=file_path,
                    total=total_size,
                    unit='iB',
                    unit_scale=True,
                    unit_divisor=1024,
            ) as bar:
                for data in response.iter_content(chunk_size=1024):
                    size = file.write(data)
                    bar.update(size)

            print(blue_text('下载更新成功！请您运行新版本。'))
            return 1

        except requests.exceptions.SSLError:
            retry_count += 1
            print(f'SSLError-{retry_count}')
            if retry_count == max_retries:
                print(red_text('下载更新失败！请重试。--SSLError'))
                return new

            time.sleep(base_delay)
            # delay = base_delay * (2 ** (retry_count - 1)) # 指数退避

        except requests.exceptions.ConnectionError: # 未开启代理
            print(red_text('下载更新失败！请开启VPN并重试。--ConnectionError'))
            return 0

if __name__ == "__main__":
    update('2.0')