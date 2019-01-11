# ライブラリインポート
import urllib.request
import re
import os
import ctypes
import pythoncom
from win32com.shell import shell, shellcon
import win32gui

# html取得 type:str
html = str(urllib.request.urlopen("https://www.bing.com").read())

# 各種データ取得
available = re.findall(                                                 # 利用可能か
    "[^\":,\\s]+", re.findall("\"wp\":.*?,", html)[0])[1]

if available == "false":                                                # 壁紙使用不可の画像の場合、ここでプログラム終了
    print("can't download!")
    exit()

filename = re.findall(                                                  # 画像名
    "[^=\"\\s]+", re.findall("imgName.*?\".*?\"", html)[0])[1] + ".jpg"

url = "https://www.bing.com/hpwp/" + re.findall(                        # URL
    "[^\":\\s]+", re.findall("\"hsh\":\".*?\"", html)[0])[1]

# 保存先フォルダー作成
if os.path.isdir('./pictures') == False:
    os.mkdir('./pictures')

# ダウンロード
urllib.request.urlretrieve(url, "./pictures/" + filename)

# 壁紙のパス
path = str(os.path.dirname(os.path.abspath(__file__))) + \
    "\\pictures\\" + filename

# progman 再起動
win32gui.SendMessageTimeout(win32gui.FindWindow(
    'Progman', None), 0x52c, None, None, 0, 1000)

# 壁紙の変更 1
# SPI_SETDESKWALLPAPER = 20
# ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, path , 0)

# 壁紙の変更 2
iad = pythoncom.CoCreateInstance(shell.CLSID_ActiveDesktop, None,
                                 pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IActiveDesktop)
iad.SetWallpaper(path, 0)
iad.ApplyChanges(shellcon.AD_APPLY_ALL)
