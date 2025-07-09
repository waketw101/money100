import sys
import os
import pandas as pd
import requests
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QStyledItemDelegate
)
from PyQt6.QtCore import Qt, QRect, QTimer, QUrl
from PyQt6.QtGui import QGuiApplication, QFont, QColor, QPen, QPainter, QBrush
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput

yesterday = '20250708'

def round2(x):
    try:
        return round(float(x), 2)
    except Exception:
        if isinstance(x, str) and x.strip() in ["-", ""]:
            return 0
        return x

def safe_float(x, default=0.0):
    try:
        # 處理 ['-', '', 0, 0.0] 等異常情況
        if isinstance(x, str) and x.strip() in ["-", ""]:
            return default
        return float(x)
    except Exception:
        return default

def get_stocks_info_batch(stock_ids):
    print(f"[DEBUG] 查詢 {stock_ids} 股票")
    markets = ["tse", "otc"]
    result = {stock_id: None for stock_id in stock_ids}
    stock_ids = [str(sid) for sid in stock_ids]
    for market in markets:
        ex_chs = [f"{market}_{stock_id}.tw" for stock_id in stock_ids]
        for i in range(0, len(ex_chs), 50):
            part = ex_chs[i:i+50]
            url = f"http://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch={'|'.join(part)}&json=1&delay=0"
            headers = {
                "User-Agent": "Mozilla/5.0",
                "Referer": "http://mis.twse.com.tw/stock/index.jsp"
            }
            try:
                resp = requests.get(url, headers=headers, timeout=5)
                data = resp.json()
                if 'msgArray' in data:
                    for info in data['msgArray']:
                        stock_id = info.get("c")
                        open_price = info.get("o")
                        if stock_id in result and open_price not in [None, "", "-", '0', 0, 0.0]:
                            result[stock_id] = info

            except Exception as e:
                print(f"[DEBUG] 查詢 {market} 股票失敗: {e}")
    return result

def load_and_prepare_df():
    folder_path = r'C:\Users\lanst\Desktop\股市\排行榜標的'
    file_path = os.path.join(folder_path, f"{yesterday}_成量排行x.xlsx")
    df = pd.read_excel(file_path)
    df = df[df[9999].notnull() & (df[9999] != 0) & (df[9999] != '')]
    columns = ["代號", "名稱", "成交", "漲跌幅"]
    df = df[columns]
    df = df.fillna(0)
    df["開盤"] = 0
    df["即時成交"] = 0
    df["即時漲幅"] = 0
    float_cols = []
    # 產生 0.90 ~ 1.10（0.01間隔，包含1.00）
    n_start, n_end = 90, 111  # 0.90 ~ 1.10
    for n in range(n_start, n_end):
        val = round(n * 0.01, 2)
        col_name = f"{val:.2f}"
        df[col_name] = df["成交"].apply(lambda x: round2(x * val))
        float_cols.append(col_name)
    final_cols = ["代號", "名稱", "漲跌幅", "開盤", "即時成交", "即時漲幅"] + float_cols
    df = df[final_cols]
    for col in df.columns:
        if col not in ["代號", "名稱"]:
            df[col] = df[col].apply(round2)
    return df

def update_realtime_info(df):
    stock_ids = df["代號"].astype(str).tolist()
    info_dict = get_stocks_info_batch(stock_ids)
    for idx, row in df.iterrows():
        sid = str(row["代號"])
        info = info_dict.get(sid)
        if not info:
            continue  # 若查無即時資料，直接跳過
        try:
            aa = round2(safe_float(info.get("z", 0)))
            if aa == 0.0:
                continue  # 跳過成交價為0的股票，不更新即時資料

            df.at[idx, "開盤"] = round2(safe_float(info.get("o", 0)))
            df.at[idx, "即時成交"] = round2(safe_float(info.get("z", 0)))
            base_price = safe_float(row["1.00"])
            realtime_price = safe_float(info.get("z", 0))
            if base_price > 0:
                df.at[idx, "即時漲幅"] = round2((realtime_price - base_price) / base_price * 100)
            else:
                df.at[idx, "即時漲幅"] = 0
        except Exception as e:
            print(f"[DEBUG] 代號 {sid} 更新失敗: {e}")
    return df

class RedBorderDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        QStyledItemDelegate.paint(self, painter, option, index)
        if index.data(Qt.ItemDataRole.UserRole) == "border-red":
            pen = QPen(QColor(200, 0, 0), 2)
            painter.save()
            painter.setPen(pen)
            # 左線
            painter.drawLine(option.rect.left()+1, option.rect.top()+1, option.rect.left()+1, option.rect.bottom()-2)
            # 右線
            painter.drawLine(option.rect.right()-2, option.rect.top()+1, option.rect.right()-2, option.rect.bottom()-2)
            painter.restore()

class TableWindow(QMainWindow):
    def __init__(self, df):
        super().__init__()
        self.df = df
        self.init_ui()
        # 警報邏輯
        self.alert_cells = set()  # 儲存被警戒的cell (row, col)
        self.alert_active_cells = set()  # 觸發中
        # 改用 QMediaPlayer 做長警報
        self.alert_player = QMediaPlayer()
        self.audio_output = QAudioOutput()
        self.alert_player.setAudioOutput(self.audio_output)
        audio_path = os.path.join(os.path.dirname(__file__), "alert.wav")
        self.alert_player.setSource(QUrl.fromLocalFile(audio_path))
        self.alert_player.setLoops(-1)  # -1 代表無限循環
        self.alert_timer = QTimer(self)
        self.alert_timer.setInterval(30000)  # 30秒
        self.alert_timer.timeout.connect(self.stop_alert_sound)
        self.table.setItemDelegate(RedBorderDelegate(self.table))

    def init_ui(self):
        screen = QGuiApplication.primaryScreen().geometry()
        width = int(screen.width() * 0.95)
        height = int(screen.height() * 0.70)
        x = 50
        y = 70
        self.setWindowFlags(Qt.WindowType.Window)
        self.setGeometry(QRect(x, y, width, height))
        self.setWindowTitle("股票資料表格")
        self.table = QTableWidget(self)
        self.table.setRowCount(self.df.shape[0])
        self.table.setColumnCount(self.df.shape[1])
        self.table.setHorizontalHeaderLabels(self.df.columns.tolist())

        # 設定 1.00 欄位標題的背景色為深紅色，字體為白
        header_item = self.table.horizontalHeaderItem(self.df.columns.get_loc("1.00"))
        header_item.setBackground(QBrush(QColor(180, 0, 0)))  # 深紅色
        header_item.setForeground(QBrush(QColor(255, 255, 255)))  # 白色字

        font = QFont()
        font.setPointSize(12)
        font_red_bold = QFont()
        font_red_bold.setPointSize(12)
        font_red_bold.setBold(True)
        font_green_bold = QFont()
        font_green_bold.setPointSize(12)
        font_green_bold.setBold(True)
        deep_green = QColor(0, 100, 0)
        self.font = font
        self.font_red_bold = font_red_bold
        self.font_green_bold = font_green_bold
        self.deep_green = deep_green

        # float_cols: 0.90 ~ 1.10，包含1.00
        self.float_cols = [f"{round(n*0.01,2):.2f}" for n in range(90,111)]

        for i in range(self.df.shape[0]):
            for j in range(self.df.shape[1]):
                self.set_table_item(i, j)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        self.table.verticalHeader().setVisible(False)
        # 設定 0.90~1.10 欄位寬度為 65px
        for col_name in self.float_cols:
            col_idx = self.df.columns.get_loc(col_name)
            self.table.setColumnWidth(col_idx, 65)
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        # 取消 Qt 預設選取背景色，這樣背景色完全由你 setBackground 控制
        self.table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self.show()
        self.setGeometry(QRect(x, y, width, height))

        # cell點擊事件
        self.table.cellClicked.connect(self.handle_cell_clicked)

        # 啟動定時器，每3秒更新即時數據
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_realtime_columns)
        self.timer.start(3000)

    def set_table_item(self, i, j):
        val = self.df.iloc[i, j]
        item = QTableWidgetItem(str(val))
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        col_name = self.df.columns[j]
        # 1.00 欄位都加紅框
        if col_name == "1.00":
            item.setData(Qt.ItemDataRole.UserRole, "border-red")
        # 判斷漲跌幅/即時漲幅欄位並設置顏色及粗體
        if col_name in ["漲跌幅", "即時漲幅"]:
            try:
                fval = float(val)
                if fval > 0:
                    item.setForeground(QColor(220, 0, 0))  # 紅色
                    item.setFont(self.font_red_bold)
                elif fval < 0:
                    item.setForeground(QColor(0, 255, 0))  # 深綠
                    item.setFont(self.font_green_bold)
                else:
                    item.setFont(self.font)
            except Exception:
                item.setFont(self.font)
        else:
            item.setFont(self.font)

        # 預警紅色邊框
        if (i, j) in getattr(self, 'alert_cells', set()):
            item.setForeground(QColor(255, 0, 0))
            item.setFont(self.font_red_bold)
            item.setBackground(QColor(255, 255, 255))
            item.setData(Qt.ItemDataRole.UserRole, "alert")
            item.setToolTip("預警中，再次點擊可關閉")
        else:
            # 背景色維持與原本refresh邏輯一致
            if col_name in self.float_cols:
                try:
                    realtime_price = safe_float(self.df.iloc[i, self.df.columns.get_loc("即時成交")])
                except Exception:
                    realtime_price = None
                try:
                    cell_val = safe_float(self.df.iloc[i, j])
                    if realtime_price is not None and realtime_price > (cell_val*0.995):
                        item.setBackground(QColor(0, 0, 0))
                        item.setForeground(QColor(255, 255, 255))
                    else:
                        item.setBackground(QColor(255, 255, 255))
                        item.setForeground(QColor(0, 0, 0))
                except Exception:
                    item.setBackground(QColor(255, 255, 255))
                    item.setForeground(QColor(0, 0, 0))
        self.table.setItem(i, j, item)

    def handle_cell_clicked(self, row, col):
        col_name = self.df.columns[col]
        if col_name not in self.float_cols:
            return
        if (row, col) in self.alert_cells:
            # 已有預警，則移除紅框、取消預警、停止警報
            self.alert_cells.remove((row, col))
            self.set_table_item(row, col)
            self.stop_alert_sound()
        else:
            # 設定預警
            self.alert_cells.add((row, col))
            self.set_table_item(row, col)
            # 若已觸發警報，立刻啟動警報聲
            self.check_alert_triggered(row, col)

    def refresh_realtime_columns(self):
        # 只更新「即時成交」與「即時漲幅」兩欄及0.90~1.10欄位背景
        update_realtime_info(self.df)
        target_cols = ["即時成交", "即時漲幅"]
        for i in range(self.df.shape[0]):
            # 先更新即時成交與即時漲幅數值/顏色
            for col in target_cols:
                j = self.df.columns.get_loc(col)
                self.set_table_item(i, j)

            # 取得當前即時成交價
            try:
                realtime_price = safe_float(self.df.iloc[i, self.df.columns.get_loc("即時成交")])
            except Exception:
                realtime_price = None

            # 更新0.90~1.10欄位背景色
            for col in self.float_cols:
                j = self.df.columns.get_loc(col)
                self.set_table_item(i, j)

        # 檢查警報
        for cell in list(self.alert_cells):
            self.check_alert_triggered(*cell)

    def check_alert_triggered(self, row, col):
        col_name = self.df.columns[col]
        try:
            cell_val = safe_float(self.df.iloc[row, col])
            realtime_price = safe_float(self.df.iloc[row, self.df.columns.get_loc("即時成交")])
        except Exception:
            return
        if realtime_price > (cell_val*0.995) and (row, col) in self.alert_cells:
            if self.alert_player.playbackState() != QMediaPlayer.PlaybackState.PlayingState:
                self.start_alert_sound()
        else:
            if (row, col) in self.alert_cells and self.alert_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
                self.stop_alert_sound()

    def start_alert_sound(self):
        # 只會有一個警報聲
        if self.alert_player.playbackState() != QMediaPlayer.PlaybackState.PlayingState:
            self.alert_player.play()
            self.alert_timer.start()

    def stop_alert_sound(self):
        if self.alert_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self.alert_player.stop()
        self.alert_timer.stop()

def main():
    df = load_and_prepare_df()
    df = update_realtime_info(df)
    app = QApplication(sys.argv)
    win = TableWindow(df)
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()