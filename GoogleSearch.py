import requests
import base64
import re
import os
import time
from bs4 import BeautifulSoup
from PIL import Image

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:50.0) Gecko/20100101 Firefox/50.0'}

def download_images(url, output_folder, num):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    if num < 10:
        file_name = os.path.join(output_folder, f'00{num}.png')
    elif num < 100:
        file_name = os.path.join(output_folder, f'0{num}.png')
    elif num < 1000:
        file_name = os.path.join(output_folder, f'{num}.png')
    else:
        print('错误！搜寻的图片数量过大！请联系Sam来修改搜图数量上限！')
        return False

    while True:
        try:
            response = requests.get(url, headers=headers, timeout=3)
        except:
            time.sleep(1)
            continue
        break

    html = response.text
    soup = BeautifulSoup(html, "html.parser")
    img_tags = soup.find_all("img")

    with open('test.html', 'w', encoding='utf-8') as f:
        f.write(html)

    for img in img_tags:
        try:
            data = img.attrs["src"]
        except Exception:
            continue

        if data[:4] == "http":
            while True:
                try:
                    img_response = requests.get(data, headers=headers, timeout=3)
                except:
                    time.sleep(0.7)
                    continue
                break
            with open(file_name, 'wb') as handler: # 先下载图片
                handler.write(img_response.content)

            with Image.open(file_name) as img: # 图片过小时下载新的图片
                if min(img.size[0], img.size[1]) < 100:
                    continue
            print(f"Downloaded image to {file_name}")
            return True

    print('No images found')
    with open('debug.html', 'w', encoding='utf-8') as f:
        f.write(soup.prettify())

    with open(file_name, 'w'): # 创建一个空文件，插入的时候不会篡位
        pass
    return False

def searchName(key):
    url = 'https://www.google.com/search?q='
    while True:
        try:
            response = requests.get(url+key)
        except:
            time.sleep(0.7)
            continue
        break
    soup = BeautifulSoup(response.text, "html.parser")
    print(soup.find_all('h3')[0].text)

def main(data, output_directory=".\\img"):
    error_img = 0
    index = 1
    for content in data:
        if content == '' or content == '\n':
            continue
        url_to_scrape = f'https://www.google.com.hk/search?q={content}&udm=2'
        if not download_images(url_to_scrape, output_directory, index):
            error_img += 1
        index += 1
    if error_img:
        return error_img
    return 0

if __name__ == '__main__':
    main(['349-7059'])