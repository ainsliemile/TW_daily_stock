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

# 1. 讀取與儲存歷史狀態
state_file = 'dashboard_state.json'
if os.path.exists(state_file):
    try:
        with open(state_file, 'r', encoding='utf-8') as f:
            state = json.load(f)
    except:
        state = {}
else:
    state = {}

# 每天換日時清空昨日紀錄
if state.get('date') != today_str:
    state = {'date': today_str, 'sox_pass': False, 'sox_msg': '', 
             'morning_update_time': '', 'morning_top20': [],
             'afternoon_update_time': '', 'afternoon_top20': []}

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
        else:
            state['sox_pass'] = False
            state['sox_msg'] = f"<div class='status fail'>❌ 【濾網未通過】SOX小於5日均線 ({latest_sox:.2f} < {ma5_sox:.2f}) - 今日策略全數出清</div>"
except Exception as e:
    if not state['sox_msg']:
        state['sox_msg'] = f"<div class='status fail'>⚠️ SOX 數據抓取失敗: {e}</div>"

# 3. 讀取 Excel 中的 股票 與 ETF
try:
    df_stocks = pd.read_excel('TrackingList-TW.xlsx', sheet_name=0, header=None, dtype=str)
    df_etfs = pd.read_excel('TrackingList-TW.xlsx', sheet_name=1, header=None, dtype=str)
    df_excel = pd.concat([df_stocks, df_etfs], ignore_index=True)
    
    tickers = df_excel.iloc[:, 0].dropna().str.strip().tolist()
    names = df_excel.iloc[:, 1].dropna().str.strip().tolist()
except Exception as e:
    print(f"Excel 讀取失敗: {e}")
    tickers, names = [], []

if state['sox_pass'] and tickers:
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
                    
        results.sort(key=lambda x: x['momentum'], reverse=True)
        top20 = results[:20]
        
        buy_count = 0
        for i, r in enumerate(top20):
            r['rank'] = i + 1
            if r['momentum'] >= 9.5:
                r['status'] = "⚠️ 漲停跳過"
            elif buy_count < 5:
                r['status'] = "⭐️ 買進標的"
                buy_count += 1
            else:
                r['status'] = "觀察中"
                
        # 🌟 判斷現在是早上還是下午，並立刻存檔
        current_time_str = now.strftime('%Y-%m-%d %H:%M:%S')
        if current_hour < 11:
            state['morning_top20'] = top20
            state['morning_update_time'] = current_time_str
        else:
            state['afternoon_top20'] = top20
            state['afternoon_update_time'] = current_time_str
            
    except Exception as e:
        print(f"股價下載失敗: {e}")

with open(state_file, 'w', encoding='utf-8') as f:
    json.dump(state, f, ensure_ascii=False, indent=2)

# 4. 產生視覺化比較 HTML 網頁 (實驗室風格)
def generate_table_html(top20_data, update_time, title, subtitle):
    if not top20_data:
        return f"<div class='box'><h3>{title} <br><span style='font-size: 13px; color: #aaa;'>{subtitle}</span></h3><p style='padding: 20px; text-align: center; color: #888;'>尚未到達擷取時間，或資料等待中...</p></div>"
    
    html = f"<div class='box'><h3>{title} <br><span style='font-size: 13px; color: #aaa;'>{subtitle} (擷取於: {update_time})</span></h3>"
    html += "<table><tr><th>排名</th><th>代號</th><th>名稱</th><th>昨收</th><th>即時價</th><th>漲幅(%)</th><th>狀態</th></tr>"
    for r in top20_data:
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
    <title>動能量化實驗室：早盤 vs 尾盤</title>
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
    <h1>🔬 動能量化實驗室：進場時機測試</h1>
    <div class="header-info">網頁最後刷新時間：{now.strftime('%Y-%m-%d %H:%M:%S')}</div>
    
    {state.get('sox_msg', '')}
    
    <div class="container">
        {generate_table_html(state.get('morning_top20', []), state.get('morning_update_time', ''), "🌅 實驗組 A：早盤爆發力", "檢測開盤 09:05 的動能延續性")}
        {generate_table_html(state.get('afternoon_top20', []), state.get('afternoon_update_time', ''), "🌇 實驗組 B：尾盤穩定度", "檢測尾盤 13:20 的實體K線確認")}
    </div>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
