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

text = getText()

# 修改一行 编程提取条码数字
list = re.findall(r"\d{12,14}", text)
if len(list) == 0 :
    list = re.findall(r"X00[0-9a-zA-Z]+", text)
text = "\n".join(list)
setText(text)
print(text)