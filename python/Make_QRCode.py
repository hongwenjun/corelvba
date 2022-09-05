# 输入命令安装所需库  python -m pip install pywin32 qrcode
import win32clipboard as w
import win32con, qrcode, sys

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

text = getText()   # 从剪贴板得到文字
img = qrcode.make(text)  # 把文字转成图片
qrcode_file = sys.path[0] + '\\qrcode.png'   # 组合保存结果的文件名
img.save(qrcode_file)   # 把图片保存成文件
setText(qrcode_file)   # 把文件名写到剪贴板