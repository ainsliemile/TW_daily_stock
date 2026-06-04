import pandas as pd
import yfinance as yf
import os
import requests
import warnings
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import pytz
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

warnings.filterwarnings('ignore')

# 1. 建立超穩定安全連線引擎
session = requests.Session()
retry = Retry(connect=3, backoff_factor=0.5)
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
})

tw_tz = pytz.timezone('Asia/Taipei')
now = datetime.now(tw_tz)
today_str = now.strftime('%Y-%m-%d %H:%M:%S')

# ----------------- 寄信通知功能設定 -----------------
def send_email_notify(msg_body):
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_password = os.environ.get("GMAIL_APP_PASSWORD")
    
    if not gmail_user or not gmail_password:
        print("未設定 GMAIL_USER 或 GMAIL_APP_PASSWORD，無法發送 Email 通知。")
        return
        
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = gmail_user  # 預設寄給自己
    msg['Subject'] = f"📊 量化面板通知 ({now.strftime('%m/%d')})"
    
    msg.attach(MIMEText(msg_body, 'plain', 'utf-8'))
    
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(gmail_user, gmail_password)
        server.send_message(msg)
        server.quit()
        print("Email 通知發送成功！")
    except Exception as e:
        print(f"Email 通知發送失敗: {e}")

# ----------------- 濾網與固定標的 -----------------
# 1. 銅價 (HG=F) 濾網
cu_msg_html = ""
email_msg_cu = ""
try:
    cu_data = yf.download('HG=F', period='20d', progress=False)
    if not cu_data.empty:
        cu_close = cu_data['Close']
        if isinstance(cu_close, pd.DataFrame): cu_close = cu_close.squeeze()
        
        latest_cu = float(cu_close.iloc[-1])
        ma10_cu = float(cu_close.rolling(window=10).mean().iloc[-1])
        
        if latest_cu > ma10_cu:
            cu_msg_html = f"<div class='status pass'>✅ 【銅價濾網】現價 > 10日均線 ({latest_cu:.4f} > {ma10_cu:.4f}) - 買入前5名標的</div>"
            email_msg_cu = f"\n✅【銅價濾網】買入前5名標的 (現價:{latest_cu:.4f} > MA10:{ma10_cu:.4f})"
        else:
            cu_msg_html = f"<div class='status fail'>❌ 【銅價濾網】現價 < 10日均線 ({latest_cu:.4f} < {ma10_cu:.4f}) - 賣出前5名標的，保持現金</div>"
            email_msg_cu = f"\n❌【銅價濾網】賣出前5名標的，保持現金 (現價:{latest_cu:.4f} < MA10:{ma10_cu:.4f})"
except Exception as e:
    cu_msg_html = f"<div class='status fail'>⚠️ 銅(HG=F) 數據抓取失敗: {e}</div>"

# 2. 標普500 (^GSPC) 濾網
sp_msg_html = ""
email_msg_sp = ""
try:
    sp_data = yf.download('^GSPC', period='10d', progress=False)
    if not sp_data.empty:
        sp_close = sp_data['Close']
        if isinstance(sp_close, pd.DataFrame): sp_close = sp_close.squeeze()
        
        latest_sp = float(sp_close.iloc[-1])
        ma5_sp = float(sp_close.rolling(window=5).mean().iloc[-1])
        
        if latest_sp > ma5_sp:
            sp_msg_html = f"<span style='color:#4caf50;'>✅ 買入正2</span> (現價: {latest_sp:.2f} > MA5: {ma5_sp:.2f})"
            email_msg_sp = f"\n✅【標普500】買入正2 (現價:{latest_sp:.2f} > MA5:{ma5_sp:.2f})"
        else:
            sp_msg_html = f"<span style='color:#f44336;'>❌ 賣出正2</span> (現價: {latest_sp:.2f} < MA5: {ma5_sp:.2f})"
            email_msg_sp = f"\n❌【標普500】賣出正2 (現價:{latest_sp:.2f} < MA5:{ma5_sp:.2f})"
except Exception as e:
    sp_msg_html = "標普500數據抓取失敗"

# 3. 固定標的現價
p_631L, p_675L = 0.0, 0.0
try: p_631L = float(yf.Ticker('00631L.TW', session=session).history(period='1d')['Close'].iloc[-1])
except: pass
try: p_675L = float(yf.Ticker('00675L.TW', session=session).history(period='1d')['Close'].iloc[-1])
except: pass


# ----------------- 股票標的抓取與計算 -----------------
try:
    df_stocks = pd.read_excel('TrackingList-TW.xlsx', sheet_name=0, header=None, dtype=str)
    df_etfs = pd.read_excel('TrackingList-TW.xlsx', sheet_name=1, header=None, dtype=str)
    df_excel = pd.concat([df_stocks, df_etfs], ignore_index=True)
    df_valid = df_excel.iloc[:, 0:2].dropna()
    tickers = [str(t).strip().replace('.0', '') for t in df_valid.iloc[:, 0]]
    names = df_valid.iloc[:, 1].astype(str).str.strip().tolist()
except Exception as e:
    print(f"Excel 讀取失敗: {e}")
    tickers, names = [], []

results = []
if tickers:
    print(f"準備下載 {len(tickers)} 檔標的資料 (計算 1月+3月 平均報酬)...")
    for t, n in zip(tickers, names):
        try:
            tkr = yf.Ticker(f"{t}.TW", session=session)
            hist = tkr.history(period="100d", auto_adjust=True)
            
            if hist.empty or 'Close' not in hist.columns:
                tkr = yf.Ticker(f"{t}.TWO", session=session)
                hist = tkr.history(period="100d", auto_adjust=True)
                
            if not hist.empty and len(hist) > 0:
                hist_clean = hist['Close'].dropna()
                if len(hist_clean) >= 2:
                    c_price = float(hist_clean.iloc[-1])
                    
                    p_1m = float(hist_clean.iloc[-21]) if len(hist_clean) >= 21 else float(hist_clean.iloc[0])
                    p_3m = float(hist_clean.iloc[-63]) if len(hist_clean) >= 63 else float(hist_clean.iloc[0])
                    
                    ret_1m = ((c_price - p_1m) / p_1m) * 100
                    ret_3m = ((c_price - p_3m) / p_3m) * 100
                    avg_ret = (ret_1m + ret_3m) / 2
                    
                    results.append({
                        'ticker': t, 'name': n, 'curr_price': c_price, 
                        'ret_1m': ret_1m, 'ret_3m': ret_3m, 'avg_ret': avg_ret, 'status': ''
                    })
        except Exception:
            pass

results.sort(key=lambda x: x['avg_ret'], reverse=True)
top15 = results[:15]

email_msg_top5 = "\n\n📋【前5名買進標的】:"
for i, r in enumerate(top15):
    r['rank'] = i + 1
    if i < 5:
        r['status'] = "⭐️ 買進標的"
        email_msg_top5 += f"\n{r['name']}({r['ticker']}): 均報酬 {r['avg_ret']:.1f}%, 現價 {r['curr_price']:.2f}"
    else:
        r['status'] = "觀察中"

# 發送 Email 通知
final_email_msg = f"這是一封由量化腳本自動發送的通知\n\n📊 量化面板 ({now.strftime('%m/%d %H:%M')})" + email_msg_cu + email_msg_sp + email_msg_top5
send_email_notify(final_email_msg)

# 將今日 Top15 寫入歷史 CSV
history_file = 'historical_momentum_data.csv'
history_rows = []
for r in top15:
    history_rows.append({
        '日期': now.strftime('%Y-%m-%d'),
        '排名': r['rank'],
        '代號': r['ticker'],
        '名稱': r['name'],
        '即時價': round(r['curr_price'], 2),
        '1月報酬(%)': round(r['ret_1m'], 2),
        '3月報酬(%)': round(r['ret_3m'], 2),
        '平均報酬(%)': round(r['avg_ret'], 2),
        '狀態': r['status']
    })
if history_rows:
    df_new_history = pd.DataFrame(history_rows)
    df_new_history.to_csv(history_file, mode='a', header=not os.path.exists(history_file), index=False, encoding='utf-8-sig')


# ----------------- 生成 HTML 網頁 -----------------
table_html = "<table><tr><th>排名</th><th>代號</th><th>名稱</th><th>即時價</th><th>1月報酬(%)</th><th>3月報酬(%)</th><th>平均報酬(%)</th><th>狀態</th></tr>"
for r in top15:
    row_class = "buy-target" if "買進" in r['status'] else ""
    table_html += f"<tr class='{row_class}'><td>{r['rank']}</td><td>{r['ticker']}</td><td>{r['name']}</td><td>{r['curr_price']:.2f}</td><td>{r['ret_1m']:.2f}%</td><td>{r['ret_3m']:.2f}%</td><td><b>{r['avg_ret']:.2f}%</b></td><td>{r['status']}</td></tr>"
table_html += "</table>"

html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>動能量化實驗室 (每日 07:00 更新)</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #121212; color: #ffffff; padding: 20px; margin: 0; }}
        h1 {{ text-align: center; color: #00ffff; }}
        .header-info {{ text-align: center; margin-bottom: 20px; color: #aaa; }}
        .status {{ padding: 15px; border-radius: 8px; font-weight: bold; text-align: center; margin-bottom: 20px; max-width: 1000px; margin: 0 auto 20px auto; }}
        .pass {{ background-color: #1e4620; color: #4caf50; border: 1px solid #4caf50; }}
        .fail {{ background-color: #4a1919; color: #f44336; border: 1px solid #f44336; }}
        .fixed-box {{ max-width: 1000px; margin: 0 auto 20px auto; background: #1e1e1e; padding: 20px; border-radius: 8px; border: 1px solid #333; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }}
        .fixed-box h3 {{ margin-top: 0; color: #ffeb3b; border-bottom: 1px solid #333; padding-bottom: 10px; text-align: center; }}
        .fixed-item {{ font-size: 18px; margin: 15px 0; padding: 10px; background-color: #2c2c2c; border-radius: 5px; }}
        .container {{ display: flex; flex-wrap: wrap; gap: 20px; max-width: 1000px; margin: 0 auto; }}
        .box {{ flex: 1; background: #1e1e1e; border-radius: 8px; overflow: hidden; border: 1px solid #333; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }}
        .box h3 {{ margin: 0; padding: 15px; background: #2c2c2c; text-align: center; color: #00ffff; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border-bottom: 1px solid #333; padding: 12px 5px; text-align: center; font-size: 14px; }}
        th {{ background-color: #333; color: #aaa; }}
        .buy-target {{ background-color: #1a365d; font-weight: bold; color: #63b3ed; border-left: 4px solid #3182ce; }}
    </style>
</head>
<body>
    <h1>🔬 動能量化實驗室 (每月1號加碼、換股)</h1>
    <div class="header-info">網頁最後擷取時間：{today_str}</div>
    
    {cu_msg_html}
    
    <div class="fixed-box">
        <h3>📌 固定追蹤標的與大盤指示</h3>
        <div class="fixed-item">1. 元大台灣50正2 (00631L) 現價: <b>{p_631L:.2f}</b></div>
        <div class="fixed-item">2. 富邦臺灣加權正2 (00675L) 現價: <b>{p_675L:.2f}</b></div>
        <div class="fixed-item">3. 標普500指數 (^GSPC) 指示: <b>{sp_msg_html}</b></div>
    </div>

    <div class="container">
        <div class="box">
            <h3>🏆 綜合動能 Top 15 (1月+3月平均報酬)</h3>
            {table_html}
        </div>
    </div>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
