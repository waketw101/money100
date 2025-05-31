from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from io import StringIO
import time
import os
from datetime import datetime

# 獲取今天的日期
today = datetime.today().strftime('%Y%m%d')

options = Options()
options.add_argument("--headless")
driver = webdriver.Chrome(options=options)

url = 'https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E6%88%90%E4%BA%A4%E5%BC%B5%E6%95%B8+%28%E9%AB%98%E2%86%92%E4%BD%8E%29%40%40%E6%88%90%E4%BA%A4%E5%BC%B5%E6%95%B8%40%40%E7%94%B1%E9%AB%98%E2%86%92%E4%BD%8E'
driver.get(url)
time.sleep(3)

wanted_ranges = ["1~300", "301~600", "601~900", "901~1200"]
all_dfs = []

for wanted in wanted_ranges:
    select = driver.find_element("id", "selRANK")
    options_list = select.find_elements("tag name", "option")
    for option in options_list:
        if option.text.strip() == wanted:
            option.click()
            break

    # 等待 select value 變更，或表格行數明顯變化
    try:
        WebDriverWait(driver, 10).until(
            lambda d: d.find_element(By.ID, "selRANK").get_attribute("value") == wanted
        )
        time.sleep(1.5)  # 再保險一點，讓表格刷新
    except Exception as e:
        print(f"等待分頁資料變動超時: {e}")

    html = driver.page_source
    dfs = pd.read_html(StringIO(html))
    target_df = None
    for df in dfs:
        cols = list(df.columns)
        if "排 名" in cols and "代號" in cols and "名稱" in cols:
            target_df = df
            break
    if target_df is not None and len(target_df) > 0:
        header = list(target_df.columns)
        rows_to_drop = []
        for idx, row in target_df.iterrows():
            if list(row.values) == header:
                rows_to_drop.append(idx)
        target_df = target_df.drop(rows_to_drop).reset_index(drop=True)
        if isinstance(target_df.columns, pd.MultiIndex):
            target_df.columns = target_df.columns.get_level_values(-1)
        all_dfs.append(target_df)
        print(f"{wanted} 取得 {len(target_df)} 筆資料")
    else:
        print(f"{wanted} 沒有資料，停止往下抓")
        break

driver.quit()

if all_dfs:
    final_df = pd.concat(all_dfs, ignore_index=True)
    os.makedirs('excel/goodinfo', exist_ok=True)
    final_df.to_csv(f'excel/goodinfo/{today}_goodinfo.csv', index=False, encoding="utf-8-sig")
    print(f"已儲存 {today}_goodinfo.csv，共{len(final_df)}筆")
else:
    print("沒有任何資料可儲存")