import pyperclip

def data():
    result = {}
    while True:
        text = input()
        if text == "":
            break
        result[text.split(": ")[0]] = text.split(": ")[1]

    pyperclip.copy(result)
    print(result)

def headers():
    result = {}
    flag = 0
    key = None
    while True:
        text = input()
        if text == "":
            break
        if flag == 0:
            key = text[:len(text) - 1]
            flag = 1
        else:
            result[key] = text
            flag = 0

    pyperclip.copy(result)
    print(result)

headers()