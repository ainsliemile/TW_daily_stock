import pandas as pd
import yfinance as yf
import json
import os
from datetime import datetime
import pytz

tw_tz = pytz.timezone('Asia/Taipei')
now = datetime.now(tw_tz)
today_str = now.strftime('%Y-%m-%d')
current_hour = now.hour

# 檔案路徑設定
state_file = 'dashboard_state.json'
history_file = 'historical_momentum_data.csv'
current_snapshot = '早盤(09:05)' if current_hour < 11 else '尾盤(13:20)'

# 1. 讀取與儲存網頁目前的歷史狀態
if os.path.exists(state_file):
    try:
        with open(state_file, 'r', encoding='utf-8') as f:
            state = json.load(f)
    except:
        state = {}
else:
    state = {}

# 每天換日時清空當日網頁顯示紀錄
if state.get('date') != today_str:
    state = {'date': today_str, 'sox_pass': False, 'sox_msg': '', 'sox_csv_val': '未知',
             'morning_update_time': '', 'morning_top30': [],
             'afternoon_update_time': '', 'afternoon_top30': []}

# 2. 計算 SOX 濾網 (費城半導體 5日均線)
try:
    sox_data = yf.download('^SOX', period='10d', progress=False)
    if not sox_data.empty:
        sox_close = sox_data['Close']
        if isinstance(sox_close, pd.DataFrame): sox_close = sox_close.squeeze()
            
        latest_sox = float(sox_close.iloc[-1])
        ma5_sox = float(sox_close.rolling(window=5).mean().iloc[-1])
        
        if latest_sox >= ma5_sox:
            state['sox_pass'] = True
            state['sox_msg'] = f"<div class='status pass'>✅ 【濾網通過】SOX大於5日均線 ({latest_sox:.2f} >= {ma5_sox:.2f}) - 今日允許進場</div>"
            # 🌟 格式化存入 Excel 的文字：例如 "通過(5120.5/5080.2)"
            state['sox_csv_val'] = f"通過({latest_sox:.2f}/{ma5_sox:.2f})"
        else:
            state['sox_pass'] = False
            state['sox_msg'] = f"<div class='status fail'>❌ 【濾網未通過】SOX小於5日均線 ({latest_sox:.2f} < {ma5_sox:.2f}) - 今日策略全數出清</div>"
            # 🌟 格式化存入 Excel 的文字：例如 "未通過(4950.1/5010.4)"
            state['sox_csv_val'] = f"未通過({latest_sox:.2f}/{ma5_sox:.2f})"
except Exception as e:
    if not state['sox_msg']:
        state['sox_msg'] = f"<div class='status fail'>⚠️ SOX 數據抓取失敗: {e}</div>"
        state['sox_csv_val'] = "抓取失敗"

# 3. 讀取 Excel 中的 股票 與 ETF (雙頁合併)
try:
    df_stocks = pd.read_excel('TrackingList-TW.xlsx', sheet_name=0, header=None, dtype=str)
    df_etfs = pd.read_excel('TrackingList-TW.xlsx', sheet_name=1, header=None, dtype=str)
    df_excel = pd.concat([df_stocks, df_etfs], ignore_index=True)
    
    # 🌟 核心防呆：強制轉字串，並清除 Pandas 自動加上的 '.0' 尾巴
    raw_tickers = df_excel.iloc[:, 0].dropna().astype(str).str.strip()
    tickers = raw_tickers.str.replace(r'\.0$', '', regex=True).tolist()
    
    names = df_excel.iloc[:, 1].dropna().astype(str).str.strip().tolist()
except Exception as e:
    print(f"Excel 讀取失敗: {e}")
    tickers, names = [], []
    

# 🌟 為了對照實驗：即便 SOX 沒通過，我們依然下載資料並記錄，讓大數據不中斷！
if tickers:
    tw_tickers = [f"{t}.TW" for t in tickers]
    two_tickers = [f"{t}.TWO" for t in tickers]
    all_tickers = tw_tickers + two_tickers
    
    try:
        data = yf.download(all_tickers, period='2d', progress=False)
        prices = data['Close']
        if isinstance(prices, pd.Series): prices = prices.to_frame()
            
        results = []
        for t, n in zip(tickers, names):
            tw_t, two_t = f"{t}.TW", f"{t}.TWO"
            valid_t = None
            if tw_t in prices.columns and not pd.isna(prices[tw_t].iloc[-1]): valid_t = tw_t
            elif two_t in prices.columns and not pd.isna(prices[two_t].iloc[-1]): valid_t = two_t
                
            if valid_t:
                s = prices[valid_t].dropna()
                if len(s) >= 2:
                    y_close, c_price = float(s.iloc[-2]), float(s.iloc[-1])
                    momentum = ((c_price - y_close) / y_close) * 100
                    results.append({'rank': 0, 'ticker': t, 'name': n, 'yest_close': y_close, 'curr_price': c_price, 'momentum': momentum, 'status': ''})
                    
        # 排序並取前 30 名
        results.sort(key=lambda x: x['momentum'], reverse=True)
        top30 = results[:30]
        
        # 標記狀態 (前5名未漲停的標記為⭐️買進，其餘保留原樣供對照實驗)
        buy_count = 0
        for i, r in enumerate(top30):
            r['rank'] = i + 1
            if r['momentum'] >= 9.5:
                r['status'] = "⚠️ 漲停跳過"
            elif buy_count < 5:
                r['status'] = "⭐️ 買進標的"
                buy_count += 1
            else:
                r['status'] = "觀察中"
                
        # 紀錄當前時間戳記並寫入網格狀態
        current_time_str = now.strftime('%Y-%m-%d %H:%M:%S')
        if current_hour < 11:
            state['morning_top30'] = top30
            state['morning_update_time'] = current_time_str
        else:
            state['afternoon_top30'] = top30
            state['afternoon_update_time'] = current_time_str
            
        # 🌟 寫入歷史 CSV 檔案 (新增 SOX濾網 欄位)
        history_rows = []
        for r in top30:
            history_rows.append({
                '日期': today_str,
                '時段': current_snapshot,
                'SOX濾網': state['sox_csv_val'], # 👈 新增此欄位記錄 SOX 狀況
                '排名': r['rank'],
                '代號': r['ticker'],
                '名稱': r['name'],
                '昨收': round(r['yest_close'], 2),
                '即時價': round(r['curr_price'], 2),
                '漲幅(%)': round(r['momentum'], 2),
                '狀態': r['status']
            })
        df_new_history = pd.DataFrame(history_rows)
        # 使用 utf-8-sig 確保 Excel 開啟中文不會亂碼
        df_new_history.to_csv(history_file, mode='a', header=not os.path.exists(history_file), index=False, encoding='utf-8-sig')
        print(f"📊 成功將 30 筆 {current_snapshot} 數據與 SOX 狀態累積寫入歷史資料庫！")

    except Exception as e:
        print(f"股價下載失敗: {e}")

# 儲存狀態進 JSON
with open(state_file, 'w', encoding='utf-8') as f:
    json.dump(state, f, ensure_ascii=False, indent=2)

# 4. 產生網頁 HTML
def generate_table_html(top30_data, update_time, title, subtitle):
    if not top30_data:
        return f"<div class='box'><h3>{title} <br><span style='font-size: 13px; color: #aaa;'>{subtitle}</span></h3><p style='padding: 20px; text-align: center; color: #888;'>尚未到達擷取時間，或資料等待中...</p></div>"
    
    html = f"<div class='box'><h3>{title} <br><span style='font-size: 13px; color: #aaa;'>{subtitle} (擷取於: {update_time})</span></h3>"
    html += "<table><tr><th>排名</th><th>代號</th><th>名稱</th><th>昨收</th><th>即時價</th><th>漲幅(%)</th><th>狀態</th></tr>"
    for r in top30_data:
        row_class = "buy-target" if "買進" in r['status'] else ""
        if "漲停" in r['status']: row_class = "limit-up"
            
        html += f"<tr class='{row_class}'><td>{r['rank']}</td><td>{r['ticker']}</td><td>{r['name']}</td><td>{r['yest_close']:.2f}</td><td>{r['curr_price']:.2f}</td><td>{r['momentum']:.2f}%</td><td>{r['status']}</td></tr>"
    html += "</table></div>"
    return html

html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>動能量化實驗室：早盤 vs 尾盤 (Top 30)</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #121212; color: #ffffff; padding: 20px; margin: 0; }}
        h1 {{ text-align: center; color: #00ffff; }}
        .header-info {{ text-align: center; margin-bottom: 20px; color: #aaa; }}
        .status {{ padding: 15px; border-radius: 8px; font-weight: bold; text-align: center; margin-bottom: 20px; max-width: 800px; margin: 0 auto 20px auto; }}
        .pass {{ background-color: #1e4620; color: #4caf50; border: 1px solid #4caf50; }}
        .fail {{ background-color: #4a1919; color: #f44336; border: 1px solid #f44336; }}
        .container {{ display: flex; flex-wrap: wrap; gap: 20px; max-width: 1400px; margin: 0 auto; }}
        .box {{ flex: 1; min-width: 350px; background: #1e1e1e; border-radius: 8px; overflow: hidden; border: 1px solid #333; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }}
        .box h3 {{ margin: 0; padding: 15px; background: #2c2c2c; text-align: center; color: #00ffff; line-height: 1.4; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border-bottom: 1px solid #333; padding: 10px 5px; text-align: center; font-size: 14px; }}
        th {{ background-color: #333; color: #aaa; }}
        .buy-target {{ background-color: #1a365d; font-weight: bold; color: #63b3ed; border-left: 4px solid #3182ce; }}
        .limit-up {{ background-color: rgba(255, 152, 0, 0.15); color: #ffb74d; font-style: italic; border-left: 4px solid #ff9800; }}
    </style>
</head>
<body>
    <h1>🔬 動能量化實驗室：進場時機與大數據對照 (Top 30)</h1>
    <div class="header-info">網頁最後刷新時間：{now.strftime('%Y-%m-%d %H:%M:%S')}</div>
    
    {state.get('sox_msg', '')}
    
    <div class="container">
        {generate_table_html(state.get('morning_top30', []), state.get('morning_update_time', ''), "🌅 實驗組 A：早盤爆發力 (Top 30)", "檢測開盤 09:05 的動能延續性")}
        {generate_table_html(state.get('afternoon_top30', []), state.get('afternoon_update_time', ''), "🌇 實驗組 B：尾盤穩定度 (Top 30)", "檢測尾盤 13:20 的實體K線確認")}
    </div>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
