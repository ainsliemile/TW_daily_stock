import pandas as pd
import json
import os
import requests
import warnings
import smtplib
import time
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
    recipient = os.environ.get("GMAIL_RECIPIENT", gmail_user) 
    
    if not gmail_user or not gmail_password:
        print("⚠️ 未設定 Email 環境變數，跳過發送。")
        return
        
    try:
        msg = MIMEMultipart("alternative")
        msg['Subject'] = subject
        msg['From'] = gmail_user
        msg['To'] = recipient
        part = MIMEText(html_body, 'html', 'utf-8')
        msg.attach(part)
        
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, recipient, msg.as_string())
        server.quit()
        print("📧 Email 通知發送成功！")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# ==========================================
# 🌟 初始化設定 & 極速防護 Session
# ==========================================
tw_tz = pytz.timezone('Asia/Taipei')
now = datetime.now(tw_tz)
today_str = now.strftime('%Y-%m-%d')

state_file = 'master_dashboard_state.json'
history_file = 'master_historical_data.csv'
excel_file = 'TrackingList.xlsx'

# 建立極速與防封鎖 Session
session = requests.Session()
retry = Retry(connect=5, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)

if os.path.exists(state_file):
    try:
        with open(state_file, 'r', encoding='utf-8') as f: state = json.load(f)
    except: state = {}
else: state = {}

if state.get('date') != today_str:
    state['date'] = today_str
    for key in ['filters', 'tw_data', 'us_data']:
        if key not in state: state[key] = {} if key == 'filters' else []
    state['tw_time'] = '等待今日計算...'
    state['us_time'] = '等待今日計算...'

# ==========================================
# 🎯 核心武器：直接打 Yahoo 台灣的背後 API (不依賴 yfinance)
# ==========================================
def fetch_raw_yahoo_api(ticker):
    """
    偽裝成從 tw.stock.yahoo.com 瀏覽網頁，直接獲取 JSON 圖表歷史報價
    """
    url = f"https://query2.finance.yahoo.com/v8/finance/chart/{ticker}"
    params = {
        "range": "1y",
        "interval": "1d",
        "includePrePost": "false",
        "events": "div,splits"
    }
    # 🔥 強制掛上 tw.stock.yahoo.com 的來源標頭，讓伺服器以為是合法網頁點擊
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Referer": f"https://tw.stock.yahoo.com/quote/{ticker}",
        "Origin": "https://tw.stock.yahoo.com",
        "Accept": "application/json"
    }
    
    try:
        res = session.get(url, params=params, headers=headers, timeout=10)
        res.raise_for_status()
        data = res.json()
        
        # 解析 Yahoo 回傳的 JSON 結構
        result = data.get('chart', {}).get('result', [])
        if not result: return pd.Series()
        
        timestamps = result[0].get('timestamp', [])
        # 優先抓取還原權息的收盤價 (adjclose)，若無則抓一般收盤價 (close)
        indicators = result[0].get('indicators', {})
        try:
            closes = indicators['adjclose'][0]['adjclose']
        except KeyError:
            closes = indicators['quote'][0]['close']
            
        if not timestamps or not closes: return pd.Series()
        
        # 轉成 Pandas Series，過濾掉 None 的無效資料
        df = pd.DataFrame({'close': closes}, index=pd.to_datetime(timestamps, unit='s')).dropna()
        return df['close']
    
    except Exception as e:
        return pd.Series()

def download_robustly_custom(tickers, group_name):
    print(f"\n[{now.strftime('%H:%M:%S')}] 🚀 準備自訂抓取 {group_name} 共 {len(tickers)} 檔標的...")
    all_prices = {}
    total = len(tickers)
    
    for i, ticker in enumerate(tickers, 1):
        s = fetch_raw_yahoo_api(ticker)
        actual_ticker = ticker
        
        # 智能後綴轉換與備援機制 (.TW 找不到換 .TWO)
        if s.empty:
            fallback = None
            if ticker.endswith('.TW'): fallback = ticker.replace('.TW', '.TWO')
            elif not ticker.endswith('.TW') and not ticker.endswith('.TWO') and '.' in ticker: fallback = ticker.replace('.', '-')
            
            if fallback:
                time.sleep(0.2)
                s = fetch_raw_yahoo_api(fallback)
                if not s.empty: actual_ticker = fallback
                
        if not s.empty:
            all_prices[actual_ticker] = s
            print(f"   ({i}/{total}) ✅ 成功: {actual_ticker}")
        else:
            print(f"   ({i}/{total}) ❌ 失敗: {ticker} (查無資料或遭阻擋)")
            
        time.sleep(0.2) # 極短延遲防封鎖
        
    if not all_prices: return pd.DataFrame()
    return pd.DataFrame(all_prices)

# ==========================================
# 📊 動能與濾網計算函數
# ==========================================
def get_ma_from_series(s, window):
    s = s.dropna()
    if len(s) >= window:
        curr, ma = float(s.iloc[-1]), float(s.rolling(window).mean().iloc[-1])
        return True, curr > ma, curr, ma
    return False, False, 0, 0

def calc_mom_tw(s):
    s = s.dropna()
    if len(s) < 65: return -999
    return (((s.iloc[-1] / s.iloc[-22]) - 1 + (s.iloc[-1] / s.iloc[-64]) - 1) / 2) * 100

def calc_mom_us(s):
    s = s.dropna()
    if len(s) < 130: return -999
    return (((s.iloc[-1] / s.iloc[-22]) - 1 + (s.iloc[-1] / s.iloc[-64]) - 1 + (s.iloc[-1] / s.iloc[-127]) - 1) / 3) * 100

# ==========================================
# 🚀 執行主流程：大盤與總經
# ==========================================
macro_tickers = ["^IXIC", "^TWII", "^SOX", "BTC-USD", "GC=F"]
df_macro = download_robustly_custom(macro_tickers, "大盤與總經")

# 1. 大盤濾網
success_ixic, ix_pass, ixic_curr, ixic_ma20 = get_ma_from_series(df_macro.get("^IXIC", pd.Series()), 20)
if success_ixic: state['filters']['IXIC'] = f"20MA: {ixic_curr:.2f} {'大於' if ix_pass else '小於'} {ixic_ma20:.2f}"

success_twii, tw_pass, twii_curr, twii_ma10 = get_ma_from_series(df_macro.get("^TWII", pd.Series()).tail(30), 10)
if success_twii: state['filters']['TWII'] = f"10MA: {twii_curr:.2f} {'大於' if tw_pass else '小於'} {twii_ma10:.2f}"

s_sox = df_macro.get("^SOX", pd.Series()).dropna()
if len(s_sox) >= 64:
    sox_mom = ((s_sox.iloc[-1] / s_sox.iloc[-22] - 1) + (s_sox.iloc[-1] / s_sox.iloc[-64] - 1)) / 2
    sox_pass = sox_mom > 0
    state['filters']['SOX'] = f"動能: {sox_mom * 100:.2f}% ({'多頭' if sox_pass else '空頭'})"
else: sox_pass = '多頭' in state.get('filters', {}).get('SOX', '')

# 2. 總經與熊市警訊
success_btc, btc_pass, btc_curr, btc_ma = get_ma_from_series(df_macro.get("BTC-USD", pd.Series()), 120)
if success_btc: state['filters']['BTC'] = f"現價 {btc_curr:.1f} vs 120MA {btc_ma:.1f} ({'✅ 安全' if btc_pass else '⚠️ 熊市警訊'})"

success_gold, gold_pass, gold_curr, gold_ma = get_ma_from_series(df_macro.get("GC=F", pd.Series()), 120)
if success_gold: state['filters']['GOLD'] = f"現價 {gold_curr:.1f} vs 120MA {gold_ma:.1f} ({'✅ 安全' if gold_pass else '⚠️ 熊市警訊'})"

state['filters']['STLFSI4'] = '<a href="https://fred.stlouisfed.org/series/STLFSI4" target="_blank" style="color:#79c0ff; text-decoration:underline;">🔗 點擊查詢</a> (⚠️警戒: >0 | 🚨熊市: >0.5)'
state['filters']['CDS'] = '<a href="https://hk.investing.com/rates-bonds/united-states-cds-5-years-usd" target="_blank" style="color:#79c0ff; text-decoration:underline;">🔗 點擊查詢</a> (⚠️警戒: 月漲>20% | 🚨熊市: >40%)'

# ==========================================
# 🌞 台股模組
# ==========================================
tw_pool, tw_names_map = [], {}
try:
    xls = pd.ExcelFile(excel_file)
    for sheet in ['台灣ETF', '台灣股票']:
        if sheet in xls.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet, header=None, dtype=str).dropna(subset=[0])
            for t, n in zip(df.iloc[:, 0], df[1].fillna("")):
                clean_t = str(t).strip().upper()
                if clean_t.endswith('.0'): clean_t = clean_t[:-2]
                clean_t = clean_t.replace('.TW0', '.TWO')
                if not (clean_t.endswith('.TW') or clean_t.endswith('.TWO')): clean_t += ".TW"
                if clean_t not in tw_pool:
                    tw_pool.append(clean_t)
                    tw_names_map[clean_t] = n
except: pass

tw_fixed_tickers = ['00631L.TW', '00675L.TW']
for ft in tw_fixed_tickers:
    if ft not in tw_pool:
        tw_pool.append(ft)
        tw_names_map[ft] = ft

df_tw = download_robustly_custom(tw_pool, "台股池")

tw_results = []
for t in tw_pool:
    actual_t = t if t in df_tw.columns else (t.replace('.TW', '.TWO') if t.replace('.TW', '.TWO') in df_tw.columns else None)
    if actual_t:
        s = df_tw[actual_t]
        mom = calc_mom_tw(s)
        if mom > -900:
            is_fixed = t in tw_fixed_tickers
            status = "⭐️ 買進標的"
            if is_fixed and not ix_pass: status = "❌ 跌破IXIC濾網"
            tw_results.append({'ticker': actual_t, 'name': tw_names_map.get(t, actual_t), 'price': float(s.dropna().iloc[-1]), 'momentum': mom, 'status': status, 'is_fixed': is_fixed})

if tw_results:
    fixed = [r for r in tw_results if r['is_fixed']]
    dynamic = sorted([r for r in tw_results if not r['is_fixed']], key=lambda x: x['momentum'], reverse=True)
    state['tw_data'] = fixed + dynamic[:10]
    state['tw_time'] = now.strftime('%Y-%m-%d %H:%M:%S')

# ==========================================
# 🌇 美股模組
# ==========================================
us_pool, us_names_map = [], {}
try:
    xls = pd.ExcelFile(excel_file)
    for sheet in ['美國ETF', '美國股票']:
        if sheet in xls.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet, header=None, dtype=str).dropna(subset=[0])
            for t, n in zip(df.iloc[:, 0], df[1].fillna("")):
                clean_t = str(t).strip().upper()
                if clean_t.endswith('-0') or clean_t.endswith('.0'): clean_t = clean_t[:-2]
                clean_t = clean_t.replace('.', '-')
                if clean_t not in us_pool:
                    us_pool.append(clean_t)
                    us_names_map[clean_t] = n
except: pass

us_fixed_tickers = ['SOXL', 'USD']
for ft in us_fixed_tickers:
    if ft not in us_pool:
        us_pool.append(ft)
        us_names_map[ft] = ft

df_us = download_robustly_custom(us_pool, "美股池")

us_results = []
for t in us_pool:
    if t in df_us.columns:
        s = df_us[t]
        mom = calc_mom_us(s)
        if mom > -900:
            is_fixed = t in us_fixed_tickers
            if t == 'SOXL': status = "⭐️ 買進標的 (釘住)" if tw_pass else "❌ 跌破TWII濾網"
            elif t == 'USD': status = "⭐️ 買進標的"
            else: status = "⭐️ 買進標的" if sox_pass else "❌ SOX動能轉弱"
            us_results.append({'ticker': t, 'name': us_names_map.get(t, t), 'price': float(s.dropna().iloc[-1]), 'momentum': mom, 'status': status, 'is_fixed': is_fixed})

if us_results:
    fixed = [r for r in us_results if r['is_fixed']]
    dynamic = sorted([r for r in us_results if not r['is_fixed']], key=lambda x: x['momentum'], reverse=True)
    state['us_data'] = fixed + dynamic[:10]
    state['us_time'] = now.strftime('%Y-%m-%d %H:%M:%S')

# ==========================================
# 💾 儲存檔案與渲染 HTML
# ==========================================
with open(state_file, 'w', encoding='utf-8') as f: json.dump(state, f, ensure_ascii=False, indent=2)

history_rows = []
for phase, data_key in [('台股模組', 'tw_data'), ('美股模組', 'us_data')]:
    for r in state.get(data_key, []):
        history_rows.append({'日期': today_str, '時段': phase, '代號': r['ticker'], '名稱': r['name'], '當前股價': round(r.get('price', 0), 2), '動能(%)': round(r['momentum'], 2), '狀態': r['status']})
if history_rows: pd.DataFrame(history_rows).to_csv(history_file, mode='a', header=not os.path.exists(history_file), index=False, encoding='utf-8-sig')

def build_html_table(data_list):
    if not data_list: return "<tr><td colspan='6' style='text-align:center; color:#666;'>無資料...</td></tr>"
    rows = ""
    for idx, r in enumerate(data_list):
        row_class = "buy-target" if "買進" in r['status'] else ("sell-pinned" if r.get('is_fixed') and "賣出" in r['status'] else "sell-target" if "賣出" in r['status'] else "")
        pin_icon = "📌 釘住" if r.get('is_fixed') else f"{idx+1 - len([x for x in data_list if x.get('is_fixed')])}"
        rows += f"<tr class='{row_class}'><td>{pin_icon}</td><td><strong>{r['ticker']}</strong></td><td>{r['name']}</td><td>{r.get('price', 0):.2f}</td><td>{r['momentum']:.2f}%</td><td>{r['status']}</td></tr>"
    return rows

ixic_txt = state.get('filters', {}).get('IXIC', '等待更新...')
twii_txt = state.get('filters', {}).get('TWII', '等待更新...')
sox_txt = state.get('filters', {}).get('SOX', '等待更新...')
btc_txt = state.get('filters', {}).get('BTC', '等待更新...')
gold_txt = state.get('filters', {}).get('GOLD', '等待更新...')
stlfsi4_txt = state.get('filters', {}).get('STLFSI4', '等待更新...')
cds_txt = state.get('filters', {}).get('CDS', '等待更新...')

web_html = f"""<!DOCTYPE html>
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
        .macro-tag {{ border: 1px solid #f85149; background: #3c1818; color: #ff7b72; }}
        .container {{ display: flex; flex-wrap: wrap; gap: 20px; max-width: 1600px; margin: 0 auto; }}
        .box {{ flex: 1; min-width: 500px; background: #161b22; border-radius: 10px; border: 1px solid #30363d; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.5); }}
        .box h3 {{ margin: 0; padding: 15px; background: #21262d; text-align: center; color: #fff; border-bottom: 1px solid #30363d; }}
        .box h3 span {{ font-size: 13px; color: #8b949e; display: block; margin-top: 5px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 12px; text-align: center; font-size: 14px; border-bottom: 1px solid #21262d; }}
        th {{ background-color: #1f242c; color: #8b949e; }}
        .buy-target {{ background-color: rgba(35, 78, 156, 0.2); font-weight: bold; color: #79c0ff; border-left: 4px solid #58a6ff; }}
        .sell-target {{ background-color: rgba(248, 81, 73, 0.1); color: #ff7b72; border-left: 4px solid #f85149; text-decoration: line-through; opacity: 0.7; }}
        .sell-pinned {{ background-color: rgba(248, 81, 73, 0.15); color: #ff7b72; border-left: 4px solid #f85149; font-weight: bold; }}
    </style>
</head>
<body>
    <h1>🔬 跨市場多因子動能實驗室</h1>
    <div class="header-panel">
        <div style="margin-bottom: 10px; color: #8b949e; font-size: 14px;">大盤避險濾網狀態</div>
        <div class="filter-tag">🇺🇸 納斯達克 | {ixic_txt}</div>
        <div class="filter-tag">🇹🇼 加權指數 | {twii_txt}</div>
        <div class="filter-tag">💻 費半指數 | {sox_txt}</div>
    </div>
    <div class="header-panel" style="border: 1px solid #f85149;">
        <div style="margin-bottom: 10px; color: #ff7b72; font-size: 15px; font-weight: bold;">🚨 總體經濟與熊市警訊指標</div>
        <div class="filter-tag macro-tag">₿ BTC 120MA | {btc_txt}</div>
        <div class="filter-tag macro-tag">🥇 黃金 120MA | {gold_txt}</div>
        <div class="filter-tag macro-tag">🏦 金融壓力 (STLFSI4) | {stlfsi4_txt}</div>
        <div class="filter-tag macro-tag">🛡️ 美國 5 年期 CDS | {cds_txt}</div>
    </div>
    <div class="container">
        <div class="box">
            <h3>🇹🇼 台股避險動能池 <span>最後同步: {state.get('tw_time')}</span></h3>
            <table>
                <tr><th>屬性/排名</th><th>代號</th><th>名稱</th><th>當前股價</th><th>(1M+3M)動能</th><th>策略狀態</th></tr>
                {build_html_table(state.get('tw_data', []))}
            </table>
        </div>
        <div class="box">
            <h3>🇺🇸 美股避險動能池 <span>最後同步: {state.get('us_time')}</span></h3>
            <table>
                <tr><th>屬性/排名</th><th>代號</th><th>名稱</th><th>當前股價</th><th>(1+3+6M)動能</th><th>策略狀態</th></tr>
                {build_html_table(state.get('us_data', []))}
            </table>
        </div>
    </div>
</body>
</html>"""
with open("index.html", "w", encoding="utf-8") as f: f.write(web_html)

def build_email_table_html(data_list):
    if not data_list: return "<tr><td colspan='6' style='padding:15px; text-align:center; color:#888;'>無資料...</td></tr>"
    rows = ""
    for idx, r in enumerate(data_list):
        bg_color, text_color, border_left, font_weight, text_decor = "#161b22", "#c9d1d9", "1px solid #30363d", "normal", "none"
        if "買進" in r['status']: bg_color, text_color, border_left, font_weight = "#142c4f", "#58a6ff", "5px solid #58a6ff", "bold"
        elif "賣出" in r['status']:
            if r.get('is_fixed'): bg_color, text_color, border_left, font_weight = "#3c1818", "#ff7b72", "5px solid #f85149", "bold"
            else: bg_color, text_color, border_left, text_decor = "#211616", "#8b949e", "5px solid #da3633", "line-through"
        pin_icon = "📌 釘住" if r.get('is_fixed') else f"第 {idx+1 - len([x for x in data_list if x.get('is_fixed')])} 名"
        rows += f"""<tr style="background-color: {bg_color}; color: {text_color}; text-decoration: {text_decor};">
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px;">{pin_icon}</td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px; border-left: {border_left};"><strong>{r['ticker']}</strong></td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px;">{r['name']}</td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px; color: #ffffff; font-weight: bold;">${r.get('price', 0):.2f}</td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px;">{r['momentum']:.2f}%</td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 14px;">{r['status']}</td>
        </tr>"""
    return rows

email_html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="background-color: #0d1117; color: #c9d1d9; font-family: sans-serif; padding: 10px; margin: 0;">
    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 800px; background-color: #0d1117; margin: 0 auto;">
        <tr><td style="padding: 10px 0; text-align: center;"><h1 style="color: #58a6ff; font-size: 24px; margin-bottom: 5px;">🔬 跨市場多因子動能實驗室</h1><p style="color: #8b949e; font-size: 13px; margin: 0;">報告產生時間: {now.strftime('%Y-%m-%d %H:%M:%S')}</p></td></tr>
        <tr>
            <td style="padding: 10px 0;">
                <div style="background-color: #161b22; border: 2px solid #30363d; padding: 15px; border-radius: 8px;">
                    <div style="color: #58a6ff; font-size: 16px; font-weight: bold; margin-bottom: 12px; text-align: center; border-bottom: 1px solid #30363d; padding-bottom: 8px;">📊 大盤避險濾網狀態</div>
                    <div style="padding: 12px; margin-bottom: 10px; background-color: #21262d; border-radius: 6px; color: #ffffff; font-size: 15px; border-left: 6px solid #58a6ff;">🇺🇸 納斯達克 | <span style="color: #79c0ff;">{ixic_txt}</span></div>
                    <div style="padding: 12px; margin-bottom: 10px; background-color: #21262d; border-radius: 6px; color: #ffffff; font-size: 15px; border-left: 6px solid #34d058;">🇹🇼 加權指數 | <span style="color: #56d44f;">{twii_txt}</span></div>
                    <div style="padding: 12px; background-color: #21262d; border-radius: 6px; color: #ffffff; font-size: 15px; border-left: 6px solid #ffab70;">💻 費城半導體 | <span style="color: #ff9b57;">{sox_txt}</span></div>
                </div>
            </td>
        </tr>
        <tr>
            <td style="padding: 15px 0;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #30363d; border-radius: 8px; overflow: hidden; background-color: #161b22;">
                    <tr><td style="background-color: #21262d; padding: 15px; text-align: center; color: #ffffff; font-size: 18px; font-weight: bold; border-bottom: 1px solid #30363d;">🇹🇼 台股避險動能池</td></tr>
                    <tr><td>
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr style="background-color: #1f242c; color: #8b949e;"><th style="padding: 12px 4px; font-size: 13px;">屬性</th><th style="padding: 12px 4px; font-size: 13px;">代號</th><th style="padding: 12px 4px; font-size: 13px;">名稱</th><th style="padding: 12px 4px; font-size: 13px;">現價</th><th style="padding: 12px 4px; font-size: 13px;">動能</th><th style="padding: 12px 4px; font-size: 13px;">狀態</th></tr>
                            {build_email_table_html(state.get('tw_data', []))}
                        </table>
                    </td></tr>
                </table>
            </td>
        </tr>
        <tr>
            <td style="padding: 15px 0;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #30363d; border-radius: 8px; overflow: hidden; background-color: #161b22;">
                    <tr><td style="background-color: #21262d; padding: 15px; text-align: center; color: #ffffff; font-size: 18px; font-weight: bold; border-bottom: 1px solid #30363d;">🇺🇸 美股避險動能池</td></tr>
                    <tr><td>
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr style="background-color: #1f242c; color: #8b949e;"><th style="padding: 12px 4px; font-size: 13px;">屬性</th><th style="padding: 12px 4px; font-size: 13px;">代號</th><th style="padding: 12px 4px; font-size: 13px;">名稱</th><th style="padding: 12px 4px; font-size: 13px;">現價</th><th style="padding: 12px 4px; font-size: 13px;">動能</th><th style="padding: 12px 4px; font-size: 13px;">狀態</th></tr>
                            {build_email_table_html(state.get('us_data', []))}
                        </table>
                    </td></tr>
                </table>
            </td>
        </tr>
    </table>
</body></html>"""

send_email_notify(f"🌍 雙市場合併通知 - {today_str}", email_html)
print(f"\n🎉 執行完畢！所有報價皆已安全抓取並更新完成。")
