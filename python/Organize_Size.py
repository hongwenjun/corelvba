# python 读写剪切板内容, 先用下行命令安装运行库
# python -m pip install pywin32

import win32clipboard as w
import win32con
import re

def getText():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_TEXT)
    w.CloseClipboard()
    return(d).decode('GBK')

def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardText(aString)
    w.CloseClipboard()

def list_to_clipboard(list):
    text = " ".join(list)
    print(text)
    setText(text)

# 单位in或cm 换算mm
def convert_mm(text, unit):
    list = re.findall(r"[1-9][\d\.]*\d*", text)
    if (unit == 'in') :
        print("单位英寸")
        for i, ch in enumerate(list):
            list[i] = str((int(float(ch) * 25.4 + 0.5)))
    elif(unit == 'cm')  :
        print("单位厘米")
        for i, ch in enumerate(list):
             list[i] = str((int(float(ch) * 10 + 0.5)))         
    list_to_clipboard(list)

# 获取剪贴板文本
text = getText()
# print(text)

# 正则搜索数字，写回剪贴板
list = re.findall(r"[1-9][\d\.]*\d*[cmin\"]*", text)
list_to_clipboard(list)

# 判断厘米和英寸换算mm
match  = re.search(r"cm|in|\"", text)
if match:
    # print(match.group())
    unit = match.group()
    if (unit == 'in') or (unit == '\"') :
        convert_mm(text, 'in')
    elif(unit == 'cm')  :
        convert_mm(text, 'cm')