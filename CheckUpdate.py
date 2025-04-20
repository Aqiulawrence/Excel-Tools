import requests
import json
import re

prefix = r'https://aqiulawrence.github.io/downloads/'

def check_update():
    response = requests.get('https://api.github.com/repos/aqiulawrence/aqiulawrence.github.io/contents/downloads')
    name = json.loads(response.text)[0]['name']
    return name

def download_update(name):
    response = requests.get(prefix+name)
    with open('111.exe', 'wb') as f:
        f.write(response.content)

if __name__ == '__main__':
    name = check_update()
    print(name)
    download_update(name)