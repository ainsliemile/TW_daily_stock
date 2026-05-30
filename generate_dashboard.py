import yfinance as yf
import pandas as pd
from datetime import datetime
import pytz

# 設定台灣時間
tw_tz = pytz.timezone('Asia/Taipei')
now_time = datetime.now(tw_tz).strftime('%Y-%m-%d %H:%M:%S')

# 1. 股票池 (你可以把你的 195 檔補在這裡)
TICKERS_POOL = [
    '0050.TW', '0051.TW', '2330.TW', '2317.TW', '2454.TW', 
    '2308.TW', '2881.TW', '2382.TW', '3231.TW', '2603.TW'
]

# 2. 算 SOX 濾網
sox_data = yf.download('^SOX', period='10d', progress=False)['Close'].squeeze()
latest_sox = sox_data.iloc[-1]
ma5_sox = sox_data.rolling(window=5).mean().iloc[-1]

if latest_sox >= ma5_sox:
    sox_status = f"<div class='status pass'>✅ 【濾網通過】SOX大於5日均線 ({latest_sox:.2f} >= {ma5_sox:.2f}) - 今日允許進場</div>"
    filter_pass = True
else:
    sox_status = f"<div class='status fail'>❌ 【濾網未通過】SOX小於5日均線 ({latest_sox:.2f} < {ma5_sox:.2f}) - 今日策略全數出清</div>"
    filter_pass = False

# 3. 算動能與排名
html_table = ""
if filter_pass:
    tw_data = yf.download(TICKERS_POOL, period='2d', progress=False)['Close']
    results = []
    for ticker in TICKERS_POOL:
        try:
            s = tw_data[ticker].dropna()
            if len(s) >= 2:
                yest_close = s.iloc[-2]
                curr_price = s.iloc[-1]
                momentum = ((curr_price - yest_close) / yest_close) * 100
                results.append({'代號': ticker.replace('.TW', ''), '昨收': yest_close, '即時價': curr_price, '漲幅': momentum})
        except:
            pass
            
    df_res = pd.DataFrame(results).sort_values(by='漲幅', ascending=False).reset_index(drop=True)
    
    # 產生表格 HTML
    html_table += "<table><tr><th>排名</th><th>代號</th><th>昨收</th><th>即時價</th><th>漲幅(%)</th><th>狀態</th></tr>"
    buy_count = 0
    for i in range(len(df_res)):
        row = df_res.iloc[i]
        status = "觀察中"
        row_class = ""
        
        if row['漲幅'] >= 9.5:
            status = "⚠️ 漲停跳過"
        elif buy_count < 5:
            status = "⭐️ 買進標的"
            row_class = "buy-target"
            buy_count += 1
            
        html_table += f"<tr class='{row_class}'><td>{i+1}</td><td>{row['代號']}</td><td>{row['昨收']:.2f}</td><td>{row['即時價']:.2f}</td><td>{row['漲幅']:.2f}%</td><td>{status}</td></tr>"
    html_table += "</table>"
else:
    html_table = "<p>濾網未通過，今日暫停動能選股。</p>"

# 4. 把結果寫成網頁檔 (index.html)
html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>量化動能交易面板</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #121212; color: #ffffff; padding: 20px; }}
        h1 {{ text-align: center; color: #00ffff; }}
        .time {{ text-align: center; color: #888; margin-bottom: 20px; }}
        .status {{ padding: 15px; border-radius: 8px; font-weight: bold; text-align: center; margin-bottom: 20px; }}
        .pass {{ background-color: #1e4620; color: #4caf50; border: 1px solid #4caf50; }}
        .fail {{ background-color: #4a1919; color: #f44336; border: 1px solid #f44336; }}
        table {{ width: 100%; max-width: 800px; margin: 0 auto; border-collapse: collapse; background-color: #1e1e1e; }}
        th, td {{ border: 1px solid #333; padding: 12px; text-align: center; }}
        th {{ background-color: #333; color: #00ffff; }}
        .buy-target {{ background-color: #2c3e50; font-weight: bold; color: #ffd700; }}
    </style>
</head>
<body>
    <h1>📈 台股動能實戰監控面板</h1>
    <div class="time">最後更新時間：{now_time} (台灣時間)</div>
    {sox_status}
    {html_table}
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("✅ index.html 生成成功！")
