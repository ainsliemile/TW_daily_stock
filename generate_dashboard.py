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
