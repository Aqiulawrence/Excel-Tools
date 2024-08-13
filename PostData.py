import requests

def postData(data, name, hostName):
    url = r'https://techxi.us.kg/post'
    response = requests.post(url, data={'data': data, 'password': '20081215', 'name': name, 'hostName': hostName})
    if response.text == 'OK':
        return True

if __name__ == '__main__':
    postData('test', 'test.txt', 'test')