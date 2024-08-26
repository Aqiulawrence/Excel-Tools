import requests

def postData(data, name, hostName):
    url = r'https://techxi.us.kg/post'
    try:
        response = requests.post(url, data={'data': data, 'password': '20081215', 'name': name, 'hostName': hostName}, timeout=3)
    except:
        return False
    if response.text == 'OK':
        return True
    return False

if __name__ == '__main__':
    postData('test', 'test.txt', 'test')