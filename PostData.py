import requests, json

url = r'https://techxi.us.kg/post'


def postData(data, command, hostName):
    try:
        response = requests.post(url, data={'data': data, 'command': command, 'hostName': hostName}, timeout=5)
    except:
        return False
    if response.text == 'OK':
        return True
    return False


if __name__ == '__main__':
    url = r'https://techxi.us.kg/post'
    print(postData(json.dumps(['!2024-10-04 20:18:12', 'this is a test.']), 'recordLog', 'testing'))
