import requests
import re

headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'}

def check_update():
    pattern = r"var\s+.+\s+=\s+'(.+)';"
    url = 'https://whx2008.lanzouw.com/b0hc6wule'
    url_post = 'https://whx2008.lanzouw.com/filemoreajax.php?file=10596674'

    try:
        response = requests.get(url=url, headers=headers)
        k = re.findall(pattern, response.text)
        data_post = {'lx': '2', 'fid': '10596674', 'uid': '1158568', 'pg': '1', 'rep': '0', 't': k[0], 'k': k[1], 'up': '1', 'ls': '1', 'pwd': '3g3t'}
        res = requests.post(url=url_post, data=data_post, headers=headers)
    except:
        return 0, 0

    p_ver = r'.*v(\d\.\d+).exe'
    value = res.json()['text']
    versions = []

    for i in value:
        name = i["name_all"]
        ver = re.findall(p_ver, name)[0]
        versions.append(ver)

    return max(versions), value[0]['id']  # 一般最新都是第0个，这里直接输出第0个的id

def download_update(id, name):
    pattern = r'var\s+fid\s+=\s+(.*);'
    url = f'https://whx2008.lanzouw.com/{id}'
    res = requests.get(url=url, headers=headers)
    result = re.findall(pattern, res.text)[0]

    url_post = f'https://whx2008.lanzouw.com/ajaxm.php?file={result}'

    pattern = r'<div class="ifr">\n<iframe class=".*"\sname=".*" src="([^"]*)"'
    src = re.findall(pattern, res.text)
    urlsrc = f'https://whx2008.lanzouw.com/{src}'
    res = requests.get(url=urlsrc, headers=headers)

    h = headers
    h['referer'] = urlsrc

    pattern = r"data : { 'action':'downprocess','signs':ajaxdata,'sign':'([^']*)'"
    print(res.text)
    sign = re.findall(pattern, res.text)[0]
    pattern = r"var aihidcms = '([^']*)';"
    websignkey = re.findall(pattern, res.text)[0]
    data = {'action': 'downprocess', 'signs': '?ctdf', 'sign': sign, 'websignkey': websignkey, 'ves': '1', 'kd': '1'}
    res = requests.post(url=url_post, data=data, headers=h)
    value = res.json()
    download_url = f'{value['dom']}/file/{value['url']}'
    h={'accept-language': 'zh-CN,zh;q=0.9', 'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'}
    res=requests.get(download_url, headers=h)

    with open(name, 'wb') as f:
        f.write(res.content)

if __name__ == '__main__':
    NEW, id = check_update()
    download_update(id, f'Excel Tools v{NEW}.exe')