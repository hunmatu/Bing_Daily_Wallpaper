# ライブラリインポート
import urllib.request
import re
import os
import ctypes 
import pythoncom
from win32com.shell import shell, shellcon
import win32gui

# url指定
url ="https://www.bing.com" + re.findall("/az.*?.jpg",re.findall("g_img=.*?\".*?\"",str(urllib.request.urlopen("https://www.bing.com").read()))[0])[0]

# 保存名の指定
img = url.lstrip("https://www.bing.com/az/hprichbg/rb/")

# 保存先フォルダー作成
if os.path.isdir('./picture') == False:
    os.mkdir('./picture')

# ダウンロード
urllib.request.urlretrieve(url, "./picture/" + img)

# 壁紙のパス
path = str(os.path.dirname(os.path.abspath(__file__)))+"\\picture\\" + img

# progman 再起動
win32gui.SendMessageTimeout(win32gui.FindWindow('Progman', None),0x52c, None, None, 0, 1000)

# 壁紙の変更 1
# SPI_SETDESKWALLPAPER = 20 
# ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, path , 0) 

# 壁紙の変更 2
iad = pythoncom.CoCreateInstance(shell.CLSID_ActiveDesktop, None,
          pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IActiveDesktop)
iad.SetWallpaper(path, 0)
iad.ApplyChanges(shellcon.AD_APPLY_ALL)



