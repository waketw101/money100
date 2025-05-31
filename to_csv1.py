import pandas as pd
import os
import pymysql
from datetime import datetime
# 獲取今天的日期
today = datetime.today().strftime('%Y%m%d')

FOLDER_PATH = "excel/goodinfo"
CSV_FILE = os.path.join("excel/csv1", "3.csv")
FILTERED_CSV = os.path.join("excel/csv1", f"{today}_成量排行.csv")
DB_CONFIG = dict(
    host="localhost",
    user="root",
    password="1234567",
    database="stocks_data",
    charset="utf8mb4"
)

INDUSTRY_LIST = [
    "鋼鐵", "其他", "數位雲端", "電機機械", "電器電纜", "電腦週邊", "電子通路", "電子組件", "資訊服務", "塑膠",
    "通信網路", "航運業", "其他電子", "光電", "半導體"
]

def get_db_connection():
    return pymysql.connect(**DB_CONFIG)

def merge_and_label_industry():
    file_list = [f for f in os.listdir(FOLDER_PATH) if f.endswith(".csv")]
    df_list = [pd.read_csv(os.path.join(FOLDER_PATH, file), encoding="utf-8-sig") for file in file_list]
    merged_df = pd.concat(df_list, ignore_index=True)
    # 清理欄位名稱
    merged_df.columns = merged_df.columns.str.replace(" ", "").str.replace("'", "")
    merged_df = merged_df[~merged_df["代號"].astype(str).str.contains("[a-zA-Z]", regex=True, na=False)]
    merged_df["代號"] = merged_df["代號"].astype(str).str.replace('"', '').str.replace('=', '')
    merged_df["成交張數"] = merged_df["成交張數"].astype(str).str.replace(",", "", regex=False)
    merged_df = merged_df[merged_df["代號"].str.len() <= 4]
    merged_df["代號"] = pd.to_numeric(merged_df["代號"], errors="coerce")
    merged_df["成交"] = pd.to_numeric(merged_df["成交"], errors="coerce")
    merged_df["成交張數"] = pd.to_numeric(merged_df["成交張數"], errors="coerce")
    merged_df = merged_df[merged_df["代號"] >= 100]
    merged_df = merged_df[(merged_df["成交"] <= 300) & (merged_df["成交"] > 10) & (merged_df["成交張數"] > 200)]
    drop_cols = [
        '排名', '市場', '股價日期', 'K線', '漲跌價', '成交額(百萬)', '昨收', '振幅(%)',
        'PER', 'PBR', '折溢價％', '1個月走勢圖', '3個月走勢圖', '1年走勢圖', '3年走勢圖', '10年走勢圖'
    ]
    merged_df.drop(columns=[col for col in drop_cols if col in merged_df.columns], inplace=True)
    merged_df["股票代碼"] = merged_df["代號"].astype(int).astype(str) + ".TW"

    # 連接資料庫抓產業對應
    db = get_db_connection()
    cursor = db.cursor()
    cursor.execute("SELECT stocks_id, class_name FROM stocks_class")
    stocks_class = {int(k): v for k, v in cursor.fetchall()}
    merged_df["產業"] = merged_df["代號"].astype(int).map(stocks_class)
    # 產業欄插入到第三欄
    col = merged_df.pop("產業")
    merged_df.insert(2, "產業", col)
    merged_df = merged_df.sort_values(by="股票代碼", ascending=True)
    merged_df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    print('資料合併與產業標註完成')
    cursor.close()
    db.close()

def update_stocks_class_and_filter():
    db = get_db_connection()
    cursor = db.cursor()
    df = pd.read_csv(CSV_FILE)
    # 取得現有 stocks_id
    cursor.execute("SELECT stocks_id FROM stocks_class")
    existing_ids = set(row[0] for row in cursor.fetchall())
    # 準備要新增的資料
    df_sql = df[["代號", "產業"]].dropna().drop_duplicates()
    df_sql["代號"] = df_sql["代號"].astype(int)
    to_insert = [
        (row["代號"], row["產業"])
        for idx, row in df_sql.iterrows()
        if row["代號"] not in existing_ids and pd.notnull(row["產業"])
    ]
    if to_insert:
        sql = "INSERT INTO stocks_class (stocks_id, class_name) VALUES (%s, %s)"
        cursor.executemany(sql, to_insert)
        db.commit()
        print(f"已新增 {len(to_insert)} 筆資料到 stocks_class")
    else:
        print("沒有需要新增的資料。")
    cursor.close()
    db.close()
    # 產業類別過濾
    filtered_df = df[df["產業"].isin(INDUSTRY_LIST)].copy()
    filtered_df["成本價"] = 0
    filtered_df.to_csv(FILTERED_CSV, index=False, encoding="utf-8-sig")
    print('產業類別完成')
    # 取出「代號」欄位，不要標題
    filtered_df["代號"].to_csv(r"C:\Users\lanst\Desktop\股市\ddd.csv", index=False, header=False)

if not os.path.exists(CSV_FILE):
    merge_and_label_industry()
    update_stocks_class_and_filter()
else:
    update_stocks_class_and_filter()