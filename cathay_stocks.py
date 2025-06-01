import pyautogui
import time
import pandas as pd
import winsound
#today = datetime.today().strftime('%Y%m%d')
today="20250529"
count = 0
y=415
examine= [58,29,29]
okgo=0
area='yilan'
#-----載入資料
pyautogui.moveTo(294, 1056, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#國泰工具列
pyautogui.moveTo(9, 37, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#國泰首頁
time.sleep(1)
pyautogui.moveTo(970, 83, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#修改商品
pyautogui.moveTo(673, 314, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下匯入
pyautogui.moveTo(849, 872, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下匯入
time.sleep(1)
pyautogui.write('ddd.csv', interval=0.05)
pyautogui.moveTo(1243, 868, duration=0.1)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#選擇檔案類型
pyautogui.moveTo(1254, 903, duration=0.1)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#選擇csv
time.sleep(0.5)
pyautogui.moveTo(1144, 898, duration=0.1)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點擊開啟檔案
pyautogui.moveTo(991, 595, duration=0.1)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下確定
pyautogui.moveTo(1389, 763, duration=0.1)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#儲存
pyautogui.moveTo(1226, 764, duration=0.1)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下完成
#----選擇組合
pyautogui.moveTo(147, 88, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下完成
pyautogui.moveTo(70, 270, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下完成
time.sleep(0.5)
#----輸出excel
pyautogui.moveTo(14, 147, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click(button="right")#點下完成
pyautogui.moveTo(113, 314, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.moveTo(287, 319, duration=0.5)  # duration=1 表示 1 秒內平滑移動輸出
pyautogui.click()#組合內所有商品
pyautogui.moveTo(976, 689, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下完成
time.sleep(10)
#----儲存csv
pyautogui.moveTo(259, 1048, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#excel
pyautogui.moveTo(12, 229, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click(button="right")#複製選單
pyautogui.moveTo(52, 301, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#點下完成
time.sleep(1)
# 讀取剪貼簿內容（支援 Excel 格式複製的表格）
df = pd.read_clipboard()

# 存成 csv 檔
df.to_csv(rf"C:\Users\lanst\Desktop\股市\國台資料\{today}.csv", index=False, encoding="utf-8-sig")
print(f"已存檔 {today}.csv")
#----關閉excel
pyautogui.moveTo(1894, 30, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#excel
pyautogui.moveTo(1011, 629, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#excel
pyautogui.moveTo(973, 596, duration=0.5)  # duration=1 表示 1 秒內平滑移動
pyautogui.click()#excel

