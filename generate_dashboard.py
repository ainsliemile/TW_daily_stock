import pandas as pd
import yfinance as yf
import json
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

# ==========================================
# 📧 寄信通知功能設定
# ==========================================
def send_email_notify(subject, html_body):
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_password = os.environ.get("GMAIL_APP_PASSWORD")
    
    # 預設寄給自己，也可以設定另一個收件人環境變數
    recipient = os.environ.get("GMAIL_RECIPIENT", gmail_user) 
    
    if not gmail_user or not gmail_password:
        print("⚠️ 未設定 GMAIL_USER 或 GMAIL_APP_PASSWORD 環境變數，跳過 Email 通知發送。")
        return
        
    try:
        msg = MIMEMultipart("alternative")
        msg['Subject'] = subject
        msg['From'] = gmail_user
        msg['To'] = recipient
        
        # 將網頁 HTML 直接作為信件內容
        part = MIMEText(html_body, 'html', 'utf-8')
        msg.attach(part)
        
        # 連線至 Gmail SMTP 伺服器
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, recipient, msg.as_string())
        server.quit()
        print("📧 Email 通知發送成功！請檢查你的信箱。")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# ==========================================
# 🌟 超穩定安全連線引擎
# ==========================================
session = requests.Session()
retry = Retry(connect=3, backoff_factor=0.5)
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
})

# 時間與狀態設定
tw_tz = pytz.timezone('Asia/Taipei')
now = datetime.now(tw_tz)
today_str = now.strftime('%Y-%m-%d')
current_hour = now.hour

state_file = 'master_dashboard_state.json'
history_file = 'master_historical_data.csv'

# 1. 讀取歷史狀態機（跨時段保留記憶的核心機制）
if os.path.exists(state_file):
    try:
        with open(state_file, 'r', encoding='utf-8') as f:
            state = json.load(f)
    except:
        state = {}
else:
    state = {}

# 🌟 換日防呆修正：保留歷史數據，只更新日期標記
if state.get('date') != today_str:
    state['date'] = today_str
    if 'filters' not in state: state['filters'] = {}
    if 'tw_data' not in state: state['tw_data'] = []
    if 'us_data' not in state: state['us_data'] = []
    if 'tw_time' not in state: state['tw_time'] = '等待今日計算...'
    if 'us_time' not in state: state['us_time'] = '等待今日計算...'

# ==========================================
# 📊 濾網與數據計算引擎
# ==========================================
def get_ma_filter(ticker, window, period="50d"):
    try:
        df = yf.download(ticker, period=period, progress=False)
        if not df.empty:
            s = df['Close']
            if isinstance(s, pd.DataFrame): s = s.squeeze()
            curr = float(s.iloc[-1])
            ma = float(s.rolling(window).mean().iloc[-1])
            return curr > ma, curr, ma
    except:
        pass
    return False, 0, 0

def get_sox_momentum():
    try:
        df = yf.download("^SOX", period="100d", progress=False)
        if not df.empty:
            s = df['Close']
            if isinstance(s, pd.DataFrame): s = s.squeeze()
            m1 = (s.iloc[-1] / s.iloc[-22] - 1) if len(s) > 22 else 0
            m3 = (s.iloc[-1] / s.iloc[-64] - 1) if len(s) > 64 else 0
            avg_mom = (m1 + m3) / 2
            return avg_mom > 0, avg_mom * 100
    except:
        pass
    return False, 0

def fetch_close_series(ticker):
    try:
        tkr = yf.Ticker(ticker, session=session)
        hist = tkr.history(period="1y", auto_adjust=True)
        return hist['Close'].dropna() if not hist.empty else pd.Series()
    except:
        return pd.Series()

def calc_mom_tw(s):
    if len(s) < 65: return -999
    m1 = (s.iloc[-1] / s.iloc[-22]) - 1
    m3 = (s.iloc[-1] / s.iloc[-64]) - 1
    return ((m1 + m3) / 2) * 100

def calc_mom_us(s):
    if len(s) < 130: return -999
    m1 = (s.iloc[-1] / s.iloc[-22]) - 1
    m3 = (s.iloc[-1] / s.iloc[-64]) - 1
    m6 = (s.iloc[-1] / s.iloc[-127]) - 1
    return ((m1 + m3 + m6) / 3) * 100

# ==========================================
# 🌞 早上 7 點前：執行台股模組
# ==========================================
is_morning_run = current_hour < 11

if is_morning_run:
    print("🌅 觸發早上時段：正在更新 IXIC 濾網與台股動能池...")
    ixic_pass, ixic_curr, ixic_ma20 = get_ma_filter("^IXIC", 20)
    state['filters']['IXIC'] = f"IXIC 20MA: {ixic_curr:.2f} {'大於' if ixic_pass else '小於'} {ixic_ma20:.2f}"
    
    try:
        df_tw1 = pd.read_excel('TrackingList-TW.xlsx', sheet_name=0, header=None, dtype=str)
        df_tw2 = pd.read_excel('TrackingList-TW.xlsx', sheet_name=1, header=None, dtype=str)
        df_tw = pd.concat([df_tw1, df_tw2], ignore_index=True).dropna(subset=[0, 1])
        tw_pool = []
        for t in df_tw.iloc[:, 0]:
            t_str = str(t).strip()
            if t_str.endswith('.0'): t_str = t_str[:-2]
            tw_pool.append(t_str)
        tw_names = df_tw.iloc[:, 1].astype(str).str.strip().tolist()
    except:
        tw_pool, tw_names = [], []

    tw_results = []
    tw_fixed_tickers = ['00631L', '00675L']
    
    for t, n in zip(tw_pool, tw_names):
        actual_t = f"{t}.TW"
        s = fetch_close_series(actual_t)
        if s.empty:
            actual_t = f"{t}.TWO"
            s = fetch_close_series(actual_t)
            
        if not s.empty:
            mom = calc_mom_tw(s)
            if mom > -900:
                is_fixed = t in tw_fixed_tickers
                status = "⭐️ 買進標的"
                if is_fixed and not ixic_pass:
                    status = "❌ 跌破IXIC濾網(強制賣出)"
                
                tw_results.append({
                    'ticker': actual_t, 'name': n, 'price': float(s.iloc[-1]), 
                    'momentum': mom, 'status': status, 'is_fixed': is_fixed
                })
                
    fixed_tw_data = [r for r in tw_results if r['is_fixed']]
    dynamic_tw_data = [r for r in tw_results if not r['is_fixed']]
    dynamic_tw_data.sort(key=lambda x: x['momentum'], reverse=True)
    top10_tw = dynamic_tw_data[:10]
    
    state['tw_data'] = fixed_tw_data + top10_tw
    state['tw_time'] = now.strftime('%Y-%m-%d %H:%M:%S')

# ==========================================
# 🌇 下午 2 點後：執行美股模組
# ==========================================
else:
    print("🌇 觸發下午時段：正在更新 TWII、SOX 濾網與美股動能池...")
    twii_pass, twii_curr, twii_ma10 = get_ma_filter("^TWII", 10, period="30d")
    sox_pass, sox_mom_val = get_sox_momentum()
    
    state['filters']['TWII'] = f"TWII 10MA: {twii_curr:.2f} {'大於' if twii_pass else '小於'} {twii_ma10:.2f}"
    state['filters']['SOX'] = f"SOX (1M+3M) 動能: {sox_mom_val:.2f}% ({'多頭' if sox_pass else '空頭'})"

    try:
        df_us1 = pd.read_excel('TrackingList-US.xlsx', sheet_name=0, header=None, dtype=str)
        df_us2 = pd.read_excel('TrackingList-US.xlsx', sheet_name=1, header=None, dtype=str)
        df_us = pd.concat([df_us1, df_us2], ignore_index=True).dropna(subset=[0, 1])
        us_pool = []
        for t in df_us.iloc[:, 0]:
            t_str = str(t).strip().upper().replace('.', '-')
            if t_str.endswith('.0'): t_str = t_str[:-2]
            us_pool.append(t_str)
        us_names = df_us.iloc[:, 1].astype(str).str.strip().tolist()
    except:
        us_pool, us_names = [], []

    us_results = []
    us_fixed_tickers = ['SOXL', 'USD']
    
    for ft in us_fixed_tickers:
        if ft not in us_pool:
            us_pool.append(ft)
            us_names.append(ft)

    for t, n in zip(us_pool, us_names):
        s = fetch_close_series(t)
        if not s.empty:
            mom = calc_mom_us(s)
            if mom > -900:
                is_fixed = t in us_fixed_tickers
                
                if t == 'SOXL':
                    status = "⭐️ 買進標的 (釘住)" if twii_pass else "❌ 跌破TWII濾網(強制賣出)"
                elif t == 'USD':
                    status = "⭐️ 買進標的 (無濾網釘住)"
                else:
                    status = "⭐️ 買進標的" if sox_pass else "❌ SOX動能轉弱(強制賣出)"
                
                us_results.append({
                    'ticker': t, 'name': n, 'price': float(s.iloc[-1]), 
                    'momentum': mom, 'status': status, 'is_fixed': is_fixed
                })

    fixed_us_data = [r for r in us_results if r['is_fixed']]
    dynamic_us_data = [r for r in us_results if not r['is_fixed']]
    dynamic_us_data.sort(key=lambda x: x['momentum'], reverse=True)
    top10_us = dynamic_us_data[:10]
    
    fixed_us_data.sort(key=lambda x: x['ticker'])
    state['us_data'] = fixed_us_data + top10_us
    state['us_time'] = now.strftime('%Y-%m-%d %H:%M:%S')

# 儲存狀態檔案
with open(state_file, 'w', encoding='utf-8') as f:
    json.dump(state, f, ensure_ascii=False, indent=2)

# ==========================================
# 🌐 網頁 HTML 自動即時渲染
# ==========================================
def build_html_table(data_list):
    if not data_list: return "<tr><td colspan='5' style='text-align:center; color:#666;'>歷史數據加載中或尚無資料...</td></tr>"
    rows = ""
    for idx, r in enumerate(data_list):
        row_class = "buy-target" if "買進" in r['status'] else "sell-target" if "賣出" in r['status'] else ""
        pin_icon = "📌 釘住" if r.get('is_fixed') else f"{idx+1 - len([x for x in data_list if x.get('is_fixed')])}"
        rows += f"<tr class='{row_class}'><td>{pin_icon}</td><td><strong>{r['ticker']}</strong></td><td>{r['name']}</td><td>{r['momentum']:.2f}%</td><td>{r['status']}</td></tr>"
    return rows

ixic_txt = state.get('filters', {}).get('IXIC', '等待早上 7 點刷新...')
twii_txt = state.get('filters', {}).get('TWII', '等待下午 2 點刷新...')
sox_txt = state.get('filters', {}).get('SOX', '等待下午 2 點刷新...')

html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🌐 跨市場多因子動能實驗室</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #0d1117; color: #c9d1d9; padding: 20px; margin: 0; }}
        h1 {{ text-align: center; color: #58a6ff; margin-bottom: 10px; }}
        .header-panel {{ background: #161b22; border: 1px solid #30363d; padding: 15px; border-radius: 8px; margin-bottom: 20px; text-align: center; }}
        .filter-tag {{ display: inline-block; background: #21262d; padding: 8px 15px; margin: 5px; border-radius: 20px; border: 1px solid #58a6ff; font-size: 14px; }}
        .container {{ display: flex; flex-wrap: wrap; gap: 20px; max-width: 1600px; margin: 0 auto; }}
        .box {{ flex: 1; min-width: 500px; background: #161b22; border-radius: 10px; border: 1px solid #30363d; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.5); }}
        .box h3 {{ margin: 0; padding: 15px; background: #21262d; text-align: center; color: #fff; border-bottom: 1px solid #30363d; }}
        .box h3 span {{ font-size: 13px; color: #8b949e; display: block; margin-top: 5px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 12px; text-align: center; font-size: 14px; border-bottom: 1px solid #21262d; }}
        th {{ background-color: #1f242c; color: #8b949e; }}
        .buy-target {{ background-color: rgba(35, 78, 156, 0.2); font-weight: bold; color: #79c0ff; border-left: 4px solid #58a6ff; }}
        .sell-target {{ background-color: rgba(248, 81, 73, 0.1); color: #ff7b72; border-left: 4px solid #f85149; text-decoration: line-through; opacity: 0.8; }}
    </style>
</head>
<body>
    <h1>🔬 跨市場多因子動能實驗室</h1>
    
    <div class="header-panel">
        <div style="margin-bottom: 10px; color: #8b949e; font-size: 14px;">大盤避險濾網狀態（跨時段自動同步）</div>
        <div class="filter-tag">🇺🇸 納斯達克 (控台股) | {ixic_txt}</div>
        <div class="filter-tag">🇹🇼 加權指數 (控SOXL) | {twii_txt}</div>
        <div class="filter-tag">💻 費城半導體 (控美股) | {sox_txt}</div>
    </div>
    
    <div class="container">
        <div class="box">
            <h3>🇹🇼 台股避險動能池 (共12檔) 
                <span>早上 07:00 刷新 | 動能算法: (1M+3M)/2 | 最後同步: {state.get('tw_time')}</span>
            </h3>
            <table>
                <tr><th>屬性/排名</th><th>代號</th><th>名稱</th><th>(1M+3M)動能</th><th>策略狀態</th></tr>
                {build_html_table(state.get('tw_data', []))}
            </table>
        </div>

        <div class="box">
            <h3>🇺🇸 美股避險動能池 (共12檔)
                <span>下午 14:00 刷新 | 動能算法: (1M+3M+6M)/3 | 最後同步: {state.get('us_time')}</span>
            </h3>
            <table>
                <tr><th>屬性/排名</th><th>代號</th><th>名稱</th><th>(1+3+6M)動能</th><th>策略狀態</th></tr>
                {build_html_table(state.get('us_data', []))}
            </table>
        </div>
    </div>
</body>
</html>"""

# 寫入 HTML 檔案
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("🌐 網頁發布引擎成功運作！最新數據已順利匯出至 index.html。")

# ==========================================
# 📧 觸發 Gmail 發送通知
# ==========================================
# 判斷信件主旨
if is_morning_run:
    mail_subject = f"🌅 晨間量化通知 (台股模組更新完畢) - {today_str}"
else:
    mail_subject = f"🌇 午後量化通知 (美股模組更新完畢) - {today_str}"

# 將剛剛產生的網頁內容，直接寄到你的信箱
send_email_notify(mail_subject, html_content)
