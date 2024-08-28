import requests  
from bs4 import BeautifulSoup  
import base64  
import re
from PIL import Image
import os

def download_images(url, output_folder, num):
    if not os.path.exists(output_folder):  
        os.makedirs(output_folder)

    while True:
        try:
            response = requests.get(url, verify=False)
        except:
            continue
        break

    response.encoding = "utf-8"
    html = response.text

    soup = BeautifulSoup(html, "html.parser")
    img_tags = soup.find_all("img")

    count = 1

    if num < 10:
        file_name = os.path.join(output_folder, f'00{num}.png')
    elif num < 100:
        file_name = os.path.join(output_folder, f'0{num}.png')
    elif num < 1000:
        file_name = os.path.join(output_folder, f'{num}.png')
    else:
        print('错误！搜寻的图片数量过大！请联系Sam来修改搜图数量上限！')
        return False

    for img in img_tags:
        if count >= 2:
            break
        data = img.attrs["src"]
        if data[:4] != "http":
            continue
        while True:
            try:
                img_response = requests.get(data, stream=True, verify=False)
            except:
                continue
            break

        with open(file_name, 'wb') as handler:  
            handler.write(img_response.content)
            
        with Image.open(file_name) as img:
            if min(img.size[0], img.size[1]) < 100:
                continue
        count += 1
        print(f"Downloaded image to {file_name}")
        return True

    print('No images found')
    with open(file_name, 'w'):
        pass
    return False

def searchName(key):
    url = 'https://www.google.com/search?q='
    while True:
        try:
            response = requests.get(url + key)
        except:
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
    searchName('208-70-72170')