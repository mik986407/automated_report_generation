import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver import Remote
import time
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pyautogui as pyg
import pyperclip as pycp

check1 = r'https://hcdweb.cdc.gov.tw/HASWeb/Security/Public/LoginForm.aspx?ReturnUrl=%2fHasweb'
for i in range(10):
    options = webdriver.ChromeOptions()
    options.add_argument('ignore-certificate-errors')
    s = Service(r"D:\Users\mik986407\python_code\chromedriver.exe")
    driver = webdriver.Chrome(service=s,options=options)
    driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/Security/Public/LoginForm.aspx?ReturnUrl=%2fHasweb")
    driver.maximize_window()
    time.sleep(0.1)
    # 按下 CTRL + L 選擇網址欄
    pyg.hotkey('ctrl', 'l')

    # 按下 CTRL + C 複製當前網頁的連結
    test1=pyg.hotkey('ctrl', 'c')

    # 從剪貼板中獲取連結
    current_url = pycp.paste()
    current_url2 = str(current_url)
    
    if current_url2 == check1:
        print('抓到連結!')
        driver.close()
        break
    else:
        print('抓取失敗')
        driver.close()
#============================================================================================================
options = webdriver.ChromeOptions()
options.add_argument('ignore-certificate-errors')
s = Service(r"D:\Users\mik986407\python_code\chromedriver.exe")
driver = webdriver.Chrome(service=s,options=options)
driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/Security/Public/LoginForm.aspx?ReturnUrl=%2fHasweb")
driver.maximize_window()

#=============================================================================================================
# 獲取所有視窗
windows = pyg.getAllWindows()

# 獲取螢幕的寬度
screen_width = pyg.size().width

# 初始化最右邊的視窗
rightmost_window = None
rightmost_x = 0

# 迭代每個視窗
for window in windows:
    # 獲取視窗的左上角座標
    x, y, width, height = window.left, window.top, window.width, window.height
    # 判斷視窗是否位於最右邊
    if x + width > rightmost_x:
        rightmost_window = window
        rightmost_x = x + width

# 如果找到了最右邊的視窗，則點擊其標題欄
if rightmost_window:
    pyg.click(rightmost_window.left + 10, rightmost_window.top + 10)


#改換 pyautogui 操作點擊

# pyg.moveTo(801, 1057, duration = 0.1) 
# pyg.click(clicks=2, interval=0.5, button='left')
# 登入追管系統
#點選自然人憑證卡
pyg.moveTo(1013, 218, duration = 0.1) 
pyg.click(clicks=2, interval=0.5, button='left')
#點選自然人憑證卡
pyg.moveTo(966, 299, duration = 0.1) 
pyg.click(clicks=2, interval=0.5, button='left')
#點選登入
pyg.moveTo(959, 302, duration = 0.1) 
pyg.click(clicks=2, interval=0.5, button='left')
#確保密碼輸入位置正確
pyg.moveTo(939, 559, duration = 0.1) 
pyg.click(clicks=2, interval=0.5, button='left')
time.sleep(0.1)
#輸入密碼
text1 = list('870716')
for i in text1:
    pyg.press(i)

#    
pyg.moveTo(880, 587, duration = 0.1) 
pyg.click(clicks=2, interval=0.5, button='left')
print(str(driver.current_url))# 顯示當前的url
time.sleep(2)
# 個案基本資料匯出
pyg.moveTo(386, 209, duration = 0.1) 
pyg.moveTo(365, 234, duration = 0.1) 
pyg.click(clicks=1, interval=0.5, button='left')
# pyg.moveTo(632, 346, duration = 0.1)
pyg.moveTo(273, 313, duration = 0.1) 
pyg.click(clicks=1, interval=0.5, button='left')
pyg.press('num2')
pyg.press('num0')
pyg.press('num2')
pyg.press('num3')
pyg.press('num0')
pyg.press('num1')
pyg.press('num3')
pyg.press('num1')
    
pyg.moveTo(227, 471, duration = 0.1) 
pyg.click(clicks=1, interval=0.5, button='left')
pyg.moveTo(239, 505, duration = 0.1) 
pyg.click(clicks=1, interval=0.5, button='left')
pyg.moveTo(164, 575, duration = 0.1) 
pyg.click(clicks=1, interval=0.5, button='left')

pyg.moveTo(36, 734, duration = 0.1) 
pyg.click(clicks=1, interval=0.5, button='left')

pyg.moveTo(1063, 341, duration = 0.1) 
pyg.click(clicks=1, interval=0.5, button='left')
time.sleep(900)

driver.close()