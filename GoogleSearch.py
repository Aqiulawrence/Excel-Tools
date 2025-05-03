import requests
import os
import time
import random

from bs4 import BeautifulSoup
from colorama import Fore, Style
from threading import Lock
from urllib.parse import urlparse
from concurrent.futures.thread import ThreadPoolExecutor

from ExcelSearch import yellow_text

# 使用旧版User-Agent以防止显示base64图片，并且更加方便进行网站信息处理
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:50.0) Gecko/20100101 Firefox/50.0'}

CONFIG_DIR = rf'.\config'
output_directory=".\\img"
pri_path = f'./{CONFIG_DIR}/priority.txt'
black_path = f'./{CONFIG_DIR}/blacklist.txt'
priority = ['ebay.com', 'amazon.com', 'cat.com', 'alibaba.com']
blacklist = ['farfetch.com']
error_count = 0
completed = []

print_lock = Lock()
lock = Lock()

def delay(wide=True):
    if wide:
        t = random.uniform(2, 5)
        time.sleep(t)
    else:
        t = random.uniform(0, 2)
        time.sleep(t)

def safe_print(*args, **kwargs):
    with print_lock:  # 自动获取和释放锁
        print(*args, **kwargs)

def blue_text(text):
    return Fore.BLUE + Style.BRIGHT + text + Style.RESET_ALL

def green_text(text):
    return Fore.GREEN+Style.BRIGHT+text+Style.RESET_ALL

def red_text(text):
    return Fore.RED+Style.BRIGHT+text+Style.RESET_ALL

def load_config(path, content):  # 如果不存在这个文件，就创建这个文件并写入content
    if os.path.exists(path):
        with open(path, 'r') as f:
            return f.read().split()
    else:
        with open(path, 'w') as f:
            f.write('\n'.join(content))
            return content

def search_website(tag):  # 搜索图片所对应的网站
    span_count = 0
    current_tag = tag
    while current_tag and span_count < 4:
        current_tag = current_tag.find_next()
        if current_tag and current_tag.name == 'span':
            span_count += 1
    if current_tag and span_count == 4:
        website = current_tag.get_text(strip=True)
        return website

def download_image(url, index, is_filter):
    global error_count, completed
    file_name = os.path.join(output_directory, f'{index:03d}.png')

    delay(False)
    while True:
        try:
            response = requests.get(url, headers=headers)
        except Exception as e:
            parsed = urlparse(url)
            result = f"{parsed.scheme}://{parsed.netloc}"
            safe_print(yellow_text(f'[Failed]Index:{index}'), result)
            delay()
            continue
        break

    html = response.text
    soup = BeautifulSoup(html, "html.parser")
    img_tags = soup.find_all("img")
    del img_tags[0]  # 删除第0个，第0个固定为Google logo
    # with open(rf'debug.html', 'w', encoding='utf-8') as f: f.write(soup.prettify())

    # 优先使用特定网站的图片
    record = []
    if is_filter:
        for img in img_tags:
            website = search_website(img)
            for pri in priority:
                if pri in website:
                    record.append(img)
    if not record:
        record = img_tags
        for img in img_tags:
            website = search_website(img)
            for black in blacklist:
                if black in website:
                    record.remove(img) # 删除黑名单网站

    for img in record:
        try:
            src = img.attrs["src"]
        except Exception:
            continue

        if src[:4] == "http":  # 防止base64图片
            while True:
                try:
                    img_response = requests.get(src, headers=headers)
                except Exception as e:
                    parsed = urlparse(src)
                    result = f"{parsed.scheme}://{parsed.netloc}"
                    safe_print(yellow_text(f'[Failed]Index:{index}'), result)
                    delay()
                    continue
                break

            with open(file_name, 'wb') as handler:  # 先下载图片
                handler.write(img_response.content)
            with lock:
                completed.append(index)
            safe_print(green_text(f'[Successful]Index:{index}'), search_website(img))
            return True

    # 创建一个空文件，插入的时候不会篡位
    with open(file_name, 'w'):
        pass
    error_count += 1
    safe_print('No images found')

def main(search_term, max_workers=1024, is_filter=True):
    global priority, blacklist, error_count, completed
    priority = load_config(pri_path, priority)
    blacklist = load_config(black_path, blacklist)
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    if is_filter:
        print(blue_text('网站筛选功能已启用！'))

    search_term = [x.strip() for x in search_term if x.strip()]

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        completed = []
        for index, term in enumerate(search_term):
            url = f'https://www.google.com.hk/search?q={term}&udm=2'
            future = executor.submit(download_image, url, index, is_filter)
            futures.append(future)

    while len(completed) != len(search_term):
        correct = set(range(len(search_term)))
        missing = correct - set(completed)
        print(red_text('Missing:'), missing)
        for index in missing:
            url = f'https://www.google.com.hk/search?q={search_term[index]}&udm=2'
            download_image(url, index, is_filter)

    if error_count:
        return error_count
    return 0

if __name__ == '__main__':
    main(['JHOAT163475'])
