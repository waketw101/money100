import requests
from bs4 import BeautifulSoup
import csv
import datetime
import time
# 獲取今天的日期
today = datetime.today().strftime('%Y%m%d')
def money_fun(x, y, z):
    try:
        a = float(str(x).replace('%', '').strip())
    except Exception:
        a = 0
    try:
        b = float(str(y).replace('%', '').strip())
    except Exception:
        b = 0
    try:
        c = float(str(z).replace('%', '').strip())
    except Exception:
        c = 0
    kk = 0
    if a <= 0:
        kk = -1
    elif 0 < a < 100:
        kk = 1
    elif 100 <= a < 200:
        kk = 2
    elif 200 <= a < 300:
        kk = 3
    elif 300 <= a < 400:
        kk = 4
    elif a >= 400:
        kk = 5
    if b <= 0:
        kk = kk - 0.5
    elif 0 < b < 100:
        kk = kk + 0.5
    elif 100 <= b < 200:
        kk = kk + 1
    elif 200 <= b < 300:
        kk = kk + 1.5
    elif b >= 300:
        kk = kk + 2
    elif b >= 400:
        kk = kk + 2.5
    if c <= 0:
        kk = kk - 0.25
    elif 0 < c < 100:
        kk = kk + 0.25
    elif 100 <= c < 200:
        kk = kk + 0.5
    elif 200 <= c < 300:
        kk = kk + 0.75
    elif 300 <= c < 400:
        kk = kk + 1
    elif c >= 400:
        kk = kk + 1.25
    if a > 0 and b > 0 and c > 0:
        kk = kk + 1
    elif a > 0 and b > 0 and c < 0:
        kk = kk + 0.5
    return kk

def get_response(url, encoding='utf-8'):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    try:
        r = requests.get(url, headers=headers, timeout=10)
        r.encoding = encoding
        return r.text
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return ""

def get_stock_data(stock_list):
    stock_data = {}
    for idx, (xx, name) in enumerate(stock_list, start=1):
        xx = xx.strip()
        if not xx:
            continue
        data = [None] * 16  # 0~15
        # 收盤價, 本益比, 最近交易日
        url = f'https://fubon-ebrokerdj.fbs.com.tw/Z/ZC/ZCX/ZCX_{xx}.djhtm'
        html = get_response(url, encoding='big5')
        txt = BeautifulSoup(html, "html.parser").get_text()
        try:
            stock_v = txt.split("收盤價")[1].split("\n")[1].strip()
        except Exception:
            stock_v = ''
        try:
            pe_split = txt.split("本益比")[1].split("\n")[1].strip()
            P_E_ratio = 0 if 'N/A' in pe_split else pe_split
        except Exception:
            P_E_ratio = ''
        try:
            date_str = txt.split("最近交易日:")[1].split("&nbsp;")[0].split("/")
            last_transaction_date = f"{datetime.datetime.now().year}-{date_str[0]}-{date_str[1]}"
        except Exception:
            last_transaction_date = ''
        data[0] = stock_v
        data[12] = P_E_ratio
        # 外資持有 & 外資評分
        url = f'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcj/zcj_{xx}.djhtm'
        txt = BeautifulSoup(get_response(url, encoding='big5'), "html.parser").get_text()
        try:
            stock_foreign_investment_1 = txt.split("外資持股")[1].split("\n")[1].strip()
        except Exception:
            stock_foreign_investment_1 = ''
        url = f'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zcl/zcl.djhtm?a={xx}&b=2'
        txt = BeautifulSoup(get_response(url, encoding='big5'), "html.parser").get_text()
        try:
            arr = txt.split("三大法人")[1].split("\n")
            gg = 2
            foreign_score = 0
            for i in range(4, min(122, len(arr)), 13):
                try:
                    foreign_score += (int(arr[i].strip().replace(",", "") or 0) / 10) * gg
                except Exception:
                    pass
                gg -= 0.1
        except Exception:
            foreign_score = 0
        try:
            stock_foreign_investment_2 = arr[1].strip()
        except:
            stock_foreign_investment_2 = ''
        data[6] = round(foreign_score, 2)
        data[7] = stock_foreign_investment_1
        try:
            data[8] = f"{arr[4]}_{arr[17]}_{arr[30]}"
        except Exception:
            data[8] = ''
        # 營收
        url = f'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zch/zch_{xx}.djhtm'
        html = get_response(url, encoding='big5')
        html = html.replace("</td>", ' ').replace('<td class="t3n1">', ' ').replace('<td class="t3r1">', ' ')
        arr = html.split('<td class="t3n0">')
        try:
            m1 = arr[1].split()
            m2 = arr[2].split()
            m3 = arr[3].split()
            scor = money_fun(m1[4], m2[4], m3[4])
            financial_report = f"{m1[4]}_{m2[4]}_{m3[4]}"
            data[9] = scor
            data[10] = m1[4]
            data[13] = financial_report
        except Exception:
            pass
        # eps淨利
        url = f'https://fubon-ebrokerdj.fbs.com.tw/z/zc/zce/zce_{xx}.djhtm'
        html = get_response(url, encoding='big5')
        html = html.replace("</td>", ' ').replace('<td class="t3n1">', ' ').replace('<td class="t3r1">', ' ')
        arr = html.split('<td class="t3n0">')
        try:
            m1 = arr[1].split()
            m2 = arr[2].split()
            m3 = arr[3].split()
            financial_report_eps = f"{m1[10]}_{m2[10]}_{m3[10]}"
            data[11] = m1[10]
            data[14] = financial_report_eps
        except Exception:
            pass
        # 15. stock_id.TW
        data[15] = f"{xx}.TW"
        stock_data[xx] = data

        # 即時輸出完成訊息
        print(f"完成第{idx}筆資料")

        time.sleep(1)  # 避免被封鎖
    return stock_data

def main():
    stock_list = []
    # 讀取 ddd.csv，第一欄股票代號，第二欄名稱
    with open(r'C:\Users\lanst\Desktop\股市\ddd.csv', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if row and row[0].strip():
                name = row[1].strip() if len(row) > 1 else ""
                stock_list.append((row[0].strip(), name))

    data_name = [
        '昨量', '成本價', '均價', '融資餘額', '融券餘額', '成本比', '外資評分', '外資持有',
        '外資近三月', '營收評分', '近期營收', '近期esp', '本益比', '近三營收', '近三季esp'
    ]
    stock_data = get_stock_data(stock_list)

    # 輸出到CSV，使用 utf-8-sig 讓 Excel 顯示中文正常
    with open(f'excel/html_data/{today}_stocks_data.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(data_name + ['stock_id', 'name'])
        for stock_id, name in stock_list:
            row = stock_data.get(stock_id, [''] * 16)
            writer.writerow(row[:15] + [row[15], name])

if __name__ == '__main__':
    main()