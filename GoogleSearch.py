import requests
import os
import time

from bs4 import BeautifulSoup
from colorama import Fore, Style

# 使用旧版User-Agent以防止显示base64图片，并且更加方便进行网站信息处理
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:50.0) Gecko/20100101 Firefox/50.0'}

pri_path = './configs/priority.txt'
black_path = './configs/blacklist.txt'
priority = ['ebay.com', 'amazon.com', 'cat.com', 'alibaba.com']
blacklist = ['farfetch.com']

# 全局延迟
delay = 2

def load(path, content):  # 如果不存在这个文件，就创建这个文件并写入content
    if os.path.exists(path):
        with open(path, 'r') as f:
            return f.read().split()
    else:
        with open(path, 'w') as f:
            f.write('\n'.join(content))
            return content


def blue_text(text):
    return Fore.BLUE + Style.BRIGHT + text + Style.RESET_ALL

def search(tag):  # 搜索图片所对应的网站
    span_count = 0
    current_tag = tag
    while current_tag and span_count < 4:
        current_tag = current_tag.find_next()
        if current_tag and current_tag.name == 'span':
            span_count += 1
    if current_tag and span_count == 4:
        website = current_tag.get_text(strip=True)
        return website

def download_images(url, output_folder, num, is_filter):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_name = os.path.join(output_folder, f'{num:03d}.png')

    while True:
        try:
            response = requests.get(url, headers=headers)
        except Exception as e:
            print('retry-1', str(e).split(':')[0])
            time.sleep(delay)
            continue
        break

    html = response.text
    soup = BeautifulSoup(html, "html.parser")
    img_tags = soup.find_all("img")
    del img_tags[0]  # 删除第0个，第0个固定为Google logo
    # with open(rf'debug.html', 'w', encoding='utf-8') as f: f.write(soup.prettify())

    # 优先使用特定网站的图片
    priority = ['ebay.com', 'amazon.com', 'cat.com', 'alibaba.com']
    record = []
    if is_filter:
        for img in img_tags:
            website = search(img)
            for pri in priority:
                if pri in website:
                    record.append(img)
    if not record:
        record = img_tags
        for img in img_tags:
            website = search(img)
            for black in blacklist:
                if black in website:
                    record.remove(img) # 删除黑名单网站

    for img in record:
        try:
            data = img.attrs["src"]
        except Exception:
            continue

        if data[:4] == "http":  # 防止base64图片
            while True:
                try:
                    img_response = requests.get(data, headers=headers)
                except Exception as e:
                    print('retry-2', str(e).split(':')[0])
                    time.sleep(delay)
                    continue
                break
            with open(file_name, 'wb') as handler:  # 先下载图片
                handler.write(img_response.content)

            # 已删除判断图片过小的代码，可能会起到反作用

            print(f"Downloaded image to {file_name}", f"({search(img)})")
            return True

    print('No images found')

    # 创建一个空文件，插入的时候不会篡位
    with open(file_name, 'w'):
        pass

    return False

def main(data, is_filter, output_directory=".\\img"):
    global priority, blacklist
    error_img = 0
    index = 1
    # 网站筛选
    priority = load(pri_path, priority)
    blacklist = load(black_path, blacklist)
    if is_filter:
        print(blue_text('网站筛选功能已启用！'))
    for content in data:
        if content == '' or content == '\n':
            continue
        url_to_scrape = f'https://www.google.com.hk/search?q={content}&udm=2'
        if not download_images(url_to_scrape, output_directory, index, is_filter):
            error_img += 1
        index += 1
    if error_img:
        return error_img
    return 0

if __name__ == '__main__':
    main(['JHOAT163475'], True)
