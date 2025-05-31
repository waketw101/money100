"""
from pynput.mouse import Listener, Button

def on_click(x, y, button, pressed):
    if pressed and button == Button.right:
        print(f"Right click at ({x}, {y})")

# 啟動滑鼠監聽器
with Listener(on_click=on_click) as listener:
    listener.join()
"""
#------------------------------------------

import pyautogui
import time
import winsound
count = 0
y=415
examine= [58,29,29]
okgo=0
area='taipai'
Date_range=47
from PIL import ImageGrab

def get_pixel_color_at(x, y, a ,c,d,f):
    screen = ImageGrab.grab()           # 擷取整個螢幕
    color = list(screen.getpixel((x, y)))     # 取得 (x, y) 點的 RGB 值
    if f==1:
        if color == a :
            print(color)
            print(c)
        else :

            print(str(color) + str(x) + ',' + str(y))
            print(d)
            # 發出頻率 1000Hz，持續 500 毫秒
            winsound.Beep(1000, 3000)
            exit()
    if f == 2:
        if color == a:
            print(str(c)+str(d)+str(color) + str(x) + ',' + str(y)+'讀取正常')
            return 1
        else :
            print(str(c)+str(d)+str(color) + str(x) + ',' + str(y)+'讀取失敗')
            if c>=30 :
                winsound.Beep(1000, 3000)
                exit()
            return 0

if area=='taipai':
    x1,y1=386, 1044
    x2,y2,x11 = 54, 414,1041
    x3,y3 = 445, 1044
    x4,y4 = 1025, 52
    x5,y5 = 1886, 77
    x6, y6,x7, y7,x8,y8,x9,y9=1156, 783,1130, 401,1156, 576,1405, 781
    x10, y10 = 930, 915
    x12, y12 = 1901, 935
    xy1=170
if area == 'yilan':
    x1,y1 = 255,1053
    x2,y2,x11 = 38, 274,762
    x3,y3 = 296, 1063
    x4,y4 = 795, 37
    x5, y5 = 209, 94
    x6, y6, x7, y7, x8, y8, x9, y9 = 789, 660, 762, 699, 765, 816, 951, 658
    x10, y10 = 934, 759
    x12, y12 = 1907, 974
    xy1=130
while count < 1000:
# 移動滑鼠到 (500, 300) excel
    if  count==0:
        pyautogui.moveTo(x1, y1, duration=0.1)  # duration=1 表示 1 秒內平滑移動
# 模擬滑鼠左鍵點擊目前位置
        pyautogui.click()
# 移動滑鼠到 (500, 300)股票號
    pyautogui.moveTo(x2, y2, duration=0.1)  # duration=1 表示 1 秒內平滑移動
    get_pixel_color_at(x2, y2, [192, 230, 245], count,'excel股票號錯誤',1)
#pyautogui.click()
    pyautogui.doubleClick()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
# 移動滑鼠到 (500, 300)國泰
    pyautogui.moveTo(x3, y3, duration=0.1)  # duration=1 表示 1 秒內平滑移動
# 模擬滑鼠左鍵點擊目前位置
    pyautogui.click()
# 移動滑鼠到 (500, 300)股票號
    pyautogui.moveTo(x4,y4 , duration=0.1)  # duration=1 表示 1 秒內平滑移動
    get_pixel_color_at(x5, y5, [42, 42, 42], count,'國泰股票號錯誤',1)
#模擬 Ctrl + V
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('.')
    pyautogui.press('T')
    pyautogui.press('W')
    pyautogui.press('enter')
    time.sleep(0.5)
    okgo = 0
    okgo_count = 0
    while okgo== 0:
        okgo= get_pixel_color_at(x5, y5, [42, 42, 42], okgo_count, '輸入資料中1', 2)
        okgo_count = okgo_count+1
        time.sleep(0.3)
# 移動滑鼠到 (500, 300)選擇股票
#  pyautogui.moveTo(1025, 82, duration=1)  # duration=1 表示 1 秒內平滑移動
# 模擬滑鼠左鍵點擊目前位置
#  pyautogui.click()
    if Date_range==47:
    # 移動滑鼠到 (500, 300)調整日期
        pyautogui.moveTo(x6, y6, duration=0.1)  # duration=1 表示 1 秒內平滑移動
        pyautogui.click()
        pyautogui.moveTo(x7, y7, duration=0.1)  # duration=1 表示 1 秒內平滑移動
        pyautogui.click()
        pyautogui.moveTo(x8, y8, duration=0.1)  # duration=1 表示 1 秒內平滑移動
        pyautogui.click()
    if Date_range==522:
    # 移動滑鼠到 (500, 300)調整日期
        pyautogui.moveTo(x6, y6, duration=0.1)  # duration=1 表示 1 秒內平滑移動
        pyautogui.click()
        pyautogui.moveTo(897, 889, duration=0.1)  # duration=1 表示 1 秒內平滑移動
        pyautogui.click()
        pyautogui.moveTo(889,662, duration=0.1)  # duration=1 表示 1 秒內平滑移動//結束日
        pyautogui.click()
        pyautogui.moveTo(990, 878, duration=0.1)  # duration=1 表示 1 秒內平滑移動
        pyautogui.click()
    if Date_range != 1 :
        pyautogui.moveTo(x9, y9, duration=0.2)  # duration=1 表示 1 秒內平滑移動收尋
        pyautogui.click()
    time.sleep(0.5)
    okgo = 0
    okgo_count = 0
    while okgo == 0:
        okgo = get_pixel_color_at(x5, y5, [42, 42, 42], okgo_count, '輸入資料中2', 2)
        okgo_count = okgo_count + 1
        time.sleep(0.3)

# 移動滑鼠到 (500, 300)複製平均價
    pyautogui.moveTo( x10, y10, duration=0.1)  # duration=1 表示 1 秒內平滑移動
# 測試：抓取座標 (100, 200) 的顏色
    get_pixel_color_at(x10,y10,[58,29,29],count,'國泰平均價錯誤',1)
# 從當下位置拖曳到 x=200, y=200, duration為拖曳速度
    pyautogui.dragTo(x10+xy1, y10, duration=0.4)
    pyautogui.mouseUp(button='left')
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'c')
# 移動滑鼠到 (500, 300)
    pyautogui.moveTo(x1,y1, duration=0.1)  # duration=1 表示 1 秒內平滑移動
# 模擬滑鼠左鍵點擊目前位置
    pyautogui.click()
# 移動滑鼠到 (500, 300)
    pyautogui.moveTo(x11, y2, duration=0.1)  # duration=1 表示 1 秒內平滑移動
# 模擬滑鼠左鍵點擊目前位置
    pyautogui.doubleClick()
    pyautogui.press('backspace')
    pyautogui.hotkey('ctrl', 'v')
    count += 1
    # y=y+34
# 移動滑鼠到 (500, 300)向下一筆
    pyautogui.moveTo(x12, y12, duration=0.1)  # duration=1 表示 1 秒內平滑移動
# 模擬滑鼠左鍵點擊目前位置
    pyautogui.click()