import pandas as pd
from datetime import datetime
import os

today = datetime.today().strftime('%Y%m%d')
base_path = r"C:\Users\lanst\Desktop\股市\排行榜標的"

file1 = os.path.join("excel/csv1", f"{today}_成量排行.csv")
file2 = os.path.join("excel/html_data", f"{today}_stocks_data.csv")
output_file = os.path.join("excel/csv2", f"{today}_成量排行2.csv")

# 讀入檔案
df1 = pd.read_csv(file1, encoding="utf-8-sig")
df2 = pd.read_csv(file2, encoding="utf-8-sig")

# 1. 用股票代碼 <-> stock_id 對應，覆蓋成本價
if "成本價" in df1.columns and "成本價" in df2.columns:
    # 建立 mapping: 股票代碼 -> 成本價
    cost_map = df1.set_index("股票代碼")["成本價"].to_dict()
    # 用 stock_id 替換 df2["成本價"]
    df2["成本價"] = df2["stock_id"].map(cost_map).combine_first(df2["成本價"])
else:
    print("找不到成本價欄位，請檢查欄位名稱。")

# 2. 刪除成量排行的成本價
if "成本價" in df1.columns:
    df1 = df1.drop(columns=["成本價"])
# 3. 合併: 股票代碼 <-> stock_id
merged = pd.merge(
    df1,
    df2,
    left_on="股票代碼",
    right_on="stock_id",
    suffixes=('', '_from_stocks_data')
)

# 4. 存檔
# 取出「代號」欄位，不要標題
merged["股票代碼"].to_csv(r"C:\Users\lanst\Desktop\股市\ddd.csv", index=False, header=False)
merged = merged.drop(columns=["股票代碼"])
# 確保欄位存在
if '成交' in merged.columns and '昨量' in merged.columns:
    merged['均價'] = merged['成交'] - merged['昨量']
else:
    print("缺少 '成交' 或 '成本價' 欄位，無法計算成本比")
merged.to_csv(output_file, index=False, encoding="utf-8-sig")
print("國泰專用碼tw以寫入ddd", output_file)
print("合併完成，檔案位置：", output_file)