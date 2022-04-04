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

# 获取剪贴板文本
text = getText()
# print(text)

# 正则搜索数字，写回剪贴板
list = re.findall(r"[1-9][\d\.]*\d*[cmin\"]*", text)
text = " ".join(list)
print(text)
setText(text)