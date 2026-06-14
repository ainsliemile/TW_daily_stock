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
    recipient = os.environ.get("GMAIL_RECIPIENT", gmail_user) 
    
    if not gmail_user or not gmail_password:
        print("⚠️ 未設定 GMAIL_USER 或 GMAIL_APP_PASSWORD 環境變數，跳過 Email 通知發送。")
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

tw_tz = pytz.timezone('Asia/Taipei')
now = datetime.now(tw_tz)
today_str = now.strftime('%Y-%m-%d')
current_hour = now.hour

state_file = 'master_dashboard_state.json'
history_file = 'master_historical_data.csv'
excel_file = 'TrackingList.xlsx'

if os.path.exists(state_file):
    try:
        with open(state_file, 'r', encoding='utf-8') as f:
            state = json.load(f)
    except:
        state = {}
else:
    state = {}

if state.get('date') != today_str:
    state['date'] = today_str
    if 'filters' not in state: state['filters'] = {}
    if 'tw_data' not in state: state['tw_data'] = []
    if 'us_data' not in state: state['us_data'] = []
    if 'tw_time' not in state: state['tw_time'] = '等待今日計算...'
    if 'us_time' not in state: state['us_time'] = '等待今日計算...'

# ==========================================
# 📊 濾網與數據計算引擎 (防呆升級版)
# ==========================================
def get_ma_filter(ticker, window, period="50d"):
    try:
        df = yf.download(ticker, period=period, session=session, progress=False)
        if df.empty:
            print(f"⚠️ {ticker} 抓取為空值！(Yahoo 可能阻擋連線)")
            return False, False, 0, 0
            
        s = df['Close']
        if isinstance(s, pd.DataFrame): s = s.squeeze()
        s = s.dropna()
        
        if len(s) >= window:
            curr = float(s.iloc[-1])
            ma = float(s.rolling(window).mean().iloc[-1])
            return True, curr > ma, curr, ma
        else:
            print(f"⚠️ {ticker} 歷史資料不足 ({len(s)} < {window}天)")
    except Exception as e:
        print(f"⚠️ {ticker} 濾網抓取發生錯誤: {e}")
    return False, False, 0, 0

def get_sox_momentum():
    try:
        df = yf.download("^SOX", period="100d", session=session, progress=False)
        if df.empty:
            return False, False, 0
            
        s = df['Close']
        if isinstance(s, pd.DataFrame): s = s.squeeze()
        s = s.dropna()
        
        if len(s) >= 64:
            m1 = (s.iloc[-1] / s.iloc[-22] - 1) 
            m3 = (s.iloc[-1] / s.iloc[-64] - 1) 
            avg_mom = (m1 + m3) / 2
            return True, avg_mom > 0, avg_mom * 100
    except:
        pass
    return False, False, 0

# 新增：FRED 聖路易斯聯儲金融壓力指數
def get_stlfsi4_signal():
    try:
        url = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=STLFSI4"
        df = pd.read_csv(url, parse_dates=['DATE'], index_col='DATE')
        if not df.empty:
            df['STLFSI4'] = pd.to_numeric(df['STLFSI4'], errors='coerce')
            df = df.dropna()
            curr = float(df['STLFSI4'].iloc[-1])
            if curr > 0.5:
                status = "🚨 確定熊市"
            elif curr > 0:
                status = "⚠️ 警戒"
            else:
                status = "✅ 安全"
            return True, curr, status
    except Exception as e:
        print(f"⚠️ STLFSI4 抓取錯誤: {e}")
    return False, 0, "未知"

# 新增：美國 5 年期 CDS (MacroMicro)
def get_us_cds_signal():
    try:
        # 由於 MacroMicro 會阻擋自動化爬蟲，若失敗將顯示防呆提醒
        # 若未來你有 MacroMicro API Key，可將下方替換為正式 API Request
        url = "https://www.macromicro.me/collections/7840/global-cds/68239/us-5year-cds"
        # 這裡僅配置基礎框架，實際未授權環境下極大概率回傳 403 Forbidden
        response = session.get(url, timeout=5)
        if response.status_code == 200:
            # 假設我們成功解析(此處為邏輯展示，實際需解出圖表 JSON)
            # pct_change = ((curr - past_1m) / past_1m) * 100
            # return True, pct_change
            return False, "網站阻擋自動抓取" 
        else:
            return False, "無法直接連線(建議使用API)"
    except:
        pass
    return False, "抓取失敗"

def fetch_close_series(ticker):
    try:
        tkr = yf.Ticker(ticker, session=session)
        hist = tkr.history(period="1y", auto_adjust=True)
        if not hist.empty and 'Close' in hist.columns:
            s = hist['Close'].dropna()
            if not s.empty:
                last_date = s.index[-1]
                now_utc = pd.Timestamp.utcnow()
                if last_date.tz is None: last_date = last_date.tz_localize('UTC')
                else: last_date = last_date.tz_convert('UTC')
                
                if (now_utc - last_date).days > 15:
                    return pd.Series()
                return s
    except:
        pass
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

print("🔄 正在檢查並更新避險濾網與總經警訊...")
# 1. IXIC 濾網
success_ixic, ix_pass, ixic_curr, ixic_ma20 = get_ma_filter("^IXIC", 20)
if success_ixic:
    ixic_pass = ix_pass
    state['filters']['IXIC'] = f"20MA: {ixic_curr:.2f} {'大於' if ixic_pass else '小於'} {ixic_ma20:.2f}"
else:
    ixic_pass = '大於' in state.get('filters', {}).get('IXIC', '大於')

# 2. TWII 濾網
success_twii, tw_pass, twii_curr, twii_ma10 = get_ma_filter("^TWII", 10, period="30d")
if success_twii:
    twii_pass = tw_pass
    state['filters']['TWII'] = f"10MA: {twii_curr:.2f} {'大於' if twii_pass else '小於'} {twii_ma10:.2f}"
else:
    twii_pass = '大於' in state.get('filters', {}).get('TWII', '大於')

# 3. SOX 濾網
success_sox, sx_pass, sox_mom_val = get_sox_momentum()
if success_sox:
    sox_pass = sx_pass
    state['filters']['SOX'] = f"動能: {sox_mom_val:.2f}% ({'多頭' if sox_pass else '空頭'})"
else:
    sox_pass = '多頭' in state.get('filters', {}).get('SOX', '多頭')

# --- 新增：總經與熊市警訊指標 ---
# 1. BTC vs 120MA
success_btc, btc_pass, btc_curr, btc_ma = get_ma_filter("BTC-USD", 120, period="1y")
if success_btc:
    btc_status = "✅ 安全" if btc_pass else "⚠️ 熊市警訊"
    state['filters']['BTC'] = f"現價 {btc_curr:.1f} vs 120MA {btc_ma:.1f} ({btc_status})"

# 2. 黃金 vs 120MA
success_gold, gold_pass, gold_curr, gold_ma = get_ma_filter("GC=F", 120, period="1y")
if success_gold:
    gold_status = "✅ 安全" if gold_pass else "⚠️ 熊市警訊(半年~1年內)"
    state['filters']['GOLD'] = f"現價 {gold_curr:.1f} vs 120MA {gold_ma:.1f} ({gold_status})"

# 3. STLFSI4 金融壓力指數
success_stl, stl_curr, stl_status = get_stlfsi4_signal()
if success_stl:
    state['filters']['STLFSI4'] = f"數值 {stl_curr:.4f} ({stl_status})"

# 4. US CDS
success_cds, cds_msg = get_us_cds_signal()
state['filters']['CDS'] = cds_msg

run_tw = (current_hour < 11) or (len(state.get('tw_data', [])) == 0)
run_us = (current_hour >= 11) or (len(state.get('us_data', [])) == 0)

# ==========================================
# 🌞 台股模組
# ==========================================
if run_tw:
    print("🌅 觸發台股模組：正在更新台股動能池...")
    tw_pool, tw_names = [], []
    try:
        xls = pd.ExcelFile(excel_file)
        for sheet in ['台灣ETF', '台灣股票']:
            if sheet in xls.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet, header=None, dtype=str).dropna(subset=[0])
                df[1] = df[1].fillna("")
                for t in df.iloc[:, 0]:
                    t_str = str(t).strip().upper()
                    if t_str.endswith('.0'): t_str = t_str[:-2]
                    t_str = t_str.replace('.TW0', '.TWO')
                    tw_pool.append(t_str)
                tw_names.extend(df.iloc[:, 1].astype(str).str.strip().tolist())
    except Exception as e:
        print(f"台股 Excel 讀取失敗: {e}")

    tw_results = []
    tw_fixed_tickers = ['00631L', '00675L']
    
    for t, n in zip(tw_pool, tw_names):
        actual_t = t if t.endswith('.TW') or t.endswith('.TWO') else f"{t}.TW"
        s = fetch_close_series(actual_t)
        if s.empty and not (t.endswith('.TW') or t.endswith('.TWO')):
            actual_t = f"{t}.TWO"
            s = fetch_close_series(actual_t)
            
        if not s.empty:
            mom = calc_mom_tw(s)
            if mom > -900:
                is_fixed = t in tw_fixed_tickers
                status = "⭐️ 買進標的"
                if is_fixed and not ixic_pass: status = "❌ 跌破IXIC濾網(強制賣出)"
                
                tw_results.append({
                    'ticker': actual_t, 'name': n, 'price': float(s.iloc[-1]), 
                    'momentum': mom, 'status': status, 'is_fixed': is_fixed
                })
                
    fixed_tw_data = [r for r in tw_results if r['is_fixed']]
    dynamic_tw_data = [r for r in tw_results if not r['is_fixed']]
    dynamic_tw_data.sort(key=lambda x: x['momentum'], reverse=True)
    state['tw_data'] = fixed_tw_data + dynamic_tw_data[:10]
    state['tw_time'] = now.strftime('%Y-%m-%d %H:%M:%S')

# ==========================================
# 🌇 美股模組
# ==========================================
if run_us:
    print("🌇 觸發美股模組：正在更新美股動能池...")
    us_pool, us_names = [], []
    try:
        xls = pd.ExcelFile(excel_file)
        for sheet in ['美國ETF', '美國股票']:
            if sheet in xls.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet, header=None, dtype=str).dropna(subset=[0])
                df[1] = df[1].fillna("")
                for t in df.iloc[:, 0]:
                    t_str = str(t).strip().upper()
                    if t_str.endswith('-0') or t_str.endswith('.0'): t_str = t_str[:-2]
                    t_str = t_str.replace('.', '-')
                    us_pool.append(t_str)
                us_names.extend(df.iloc[:, 1].astype(str).str.strip().tolist())
    except Exception as e:
        print(f"美股 Excel 讀取失敗: {e}")

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
                if t == 'SOXL': status = "⭐️ 買進標的 (釘住)" if twii_pass else "❌ 跌破TWII濾網(強制賣出)"
                elif t == 'USD': status = "⭐️ 買進標的 (無濾網釘住)"
                else: status = "⭐️ 買進標的" if sox_pass else "❌ SOX動能轉弱(強制賣出)"
                
                us_results.append({
                    'ticker': t, 'name': n, 'price': float(s.iloc[-1]), 
                    'momentum': mom, 'status': status, 'is_fixed': is_fixed
                })

    fixed_us_data = [r for r in us_results if r['is_fixed']]
    dynamic_us_data = [r for r in us_results if not r['is_fixed']]
    dynamic_us_data.sort(key=lambda x: x['momentum'], reverse=True)
    fixed_us_data.sort(key=lambda x: x['ticker'])
    state['us_data'] = fixed_us_data + dynamic_us_data[:10]
    state['us_time'] = now.strftime('%Y-%m-%d %H:%M:%S')

with open(state_file, 'w', encoding='utf-8') as f:
    json.dump(state, f, ensure_ascii=False, indent=2)

# 💾 寫入歷史 CSV 庫
history_rows = []
for phase, data_key in [('台股模組', 'tw_data'), ('美股模組', 'us_data')]:
    for r in state.get(data_key, []):
        history_rows.append({'日期': today_str, '時段': phase, '代號': r['ticker'], '名稱': r['name'], '當前股價': round(r.get('price', 0), 2), '動能(%)': round(r['momentum'], 2), '狀態': r['status']})
if history_rows:
    pd.DataFrame(history_rows).to_csv(history_file, mode='a', header=not os.path.exists(history_file), index=False, encoding='utf-8-sig')

# ==========================================
# 🌐 網頁 HTML 自動即時渲染 (用於 GitHub Pages)
# ==========================================
def build_html_table(data_list):
    if not data_list: return "<tr><td colspan='6' style='text-align:center; color:#666;'>歷史數據加載中或尚無資料...</td></tr>"
    rows = ""
    for idx, r in enumerate(data_list):
        row_class = "buy-target" if "買進" in r['status'] else ("sell-pinned" if r.get('is_fixed') and "賣出" in r['status'] else "sell-target" if "賣出" in r['status'] else "")
        pin_icon = "📌 釘住" if r.get('is_fixed') else f"{idx+1 - len([x for x in data_list if x.get('is_fixed')])}"
        price_str = f"{r.get('price', 0):.2f}" if r.get('price', 0) > 0 else "N/A"
        rows += f"<tr class='{row_class}'><td>{pin_icon}</td><td><strong>{r['ticker']}</strong></td><td>{r['name']}</td><td>{price_str}</td><td>{r['momentum']:.2f}%</td><td>{r['status']}</td></tr>"
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
        <div style="margin-bottom: 10px; color: #8b949e; font-size: 14px;">大盤避險濾網狀態（網頁端同步）</div>
        <div class="filter-tag">🇺🇸 納斯達克 | {ixic_txt}</div>
        <div class="filter-tag">🇹🇼 加權指數 | {twii_txt}</div>
        <div class="filter-tag">💻 費半指數 | {sox_txt}</div>
    </div>

    <div class="header-panel" style="border: 1px solid #f85149;">
        <div style="margin-bottom: 10px; color: #ff7b72; font-size: 15px; font-weight: bold;">🚨 總體經濟與熊市警訊指標</div>
        <div class="filter-tag macro-tag">₿ BTC 120日均線 | {btc_txt}</div>
        <div class="filter-tag macro-tag">🥇 黃金 120日均線 | {gold_txt}</div>
        <div class="filter-tag macro-tag">🏦 金融壓力 (STLFSI4) | {stlfsi4_txt}</div>
        <div class="filter-tag macro-tag">🛡️ 美國 5年期 CDS | {cds_txt}</div>
    </div>

    <div class="container">
        <div class="box">
            <h3>🇹🇼 台股避險動能池 (共12檔) <span>最後同步: {state.get('tw_time')}</span></h3>
            <table>
                <tr><th>屬性/排名</th><th>代號</th><th>名稱</th><th>當前股價</th><th>(1M+3M)動能</th><th>策略狀態</th></tr>
                {build_html_table(state.get('tw_data', []))}
            </table>
        </div>
        <div class="box">
            <h3>🇺🇸 美股避險動能池 (共12檔) <span>最後同步: {state.get('us_time')}</span></h3>
            <table>
                <tr><th>屬性/排名</th><th>代號</th><th>名稱</th><th>當前股價</th><th>(1+3+6M)動能</th><th>策略狀態</th></tr>
                {build_html_table(state.get('us_data', []))}
            </table>
        </div>
    </div>
</body>
</html>"""

with open("index.html", "w", encoding="utf-8") as f: f.write(web_html)
print("🌐 網頁發布引擎成功運作！最新數據已順利匯出至 index.html。")

# ==========================================
# 📧 🚀 專門為 Email 打造的高對比、高清晰度渲染引擎
# ==========================================
def build_email_table_html(data_list):
    if not data_list: return "<tr><td colspan='6' style='padding:15px; text-align:center; color:#888;'>等待今日模組數據補齊...</td></tr>"
    rows = ""
    for idx, r in enumerate(data_list):
        bg_color, text_color, border_left, font_weight, text_decor = "#161b22", "#c9d1d9", "1px solid #30363d", "normal", "none"
        if "買進" in r['status']:
            bg_color, text_color, border_left, font_weight = "#142c4f", "#58a6ff", "5px solid #58a6ff", "bold"
        elif "賣出" in r['status']:
            if r.get('is_fixed'): bg_color, text_color, border_left, font_weight = "#3c1818", "#ff7b72", "5px solid #f85149", "bold"
            else: bg_color, text_color, border_left, text_decor = "#211616", "#8b949e", "5px solid #da3633", "line-through"
                
        pin_icon = "📌 釘住" if r.get('is_fixed') else f"第 {idx+1 - len([x for x in data_list if x.get('is_fixed')])} 名"
        price_str = f"{r.get('price', 0):.2f}" if r.get('price', 0) > 0 else "N/A"
        
        rows += f"""<tr style="background-color: {bg_color}; color: {text_color}; text-decoration: {text_decor};">
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px;">{pin_icon}</td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px; border-left: {border_left};"><strong>{r['ticker']}</strong></td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px;">{r['name']}</td>
            <td style="padding: 14px 6px; border-bottom: 1px solid #30363d; text-align: center; font-size: 15px; color: #ffffff; font-weight: bold;">${price_str}</td>
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
                    <div style="padding: 12px; margin-bottom: 10px; background-color: #21262d; border-radius: 6px; color: #ffffff; font-size: 15px; border-left: 6px solid #58a6ff; font-weight: bold;">🇺🇸 納斯達克 <span style="color: #8b949e; margin: 0 5px;">|</span> <span style="color: #79c0ff;">{ixic_txt}</span></div>
                    <div style="padding: 12px; margin-bottom: 10px; background-color: #21262d; border-radius: 6px; color: #ffffff; font-size: 15px; border-left: 6px solid #34d058; font-weight: bold;">🇹🇼 加權指數 <span style="color: #8b949e; margin: 0 5px;">|</span> <span style="color: #56d44f;">{twii_txt}</span></div>
                    <div style="padding: 12px; background-color: #21262d; border-radius: 6px; color: #ffffff; font-size: 15px; border-left: 6px solid #ffab70; font-weight: bold;">💻 費城半導體 <span style="color: #8b949e; margin: 0 5px;">|</span> <span style="color: #ff9b57;">{sox_txt}</span></div>
                </div>

                <div style="background-color: #2a1515; border: 2px solid #f85149; padding: 15px; border-radius: 8px; margin-top: 15px;">
                    <div style="color: #ff7b72; font-size: 16px; font-weight: bold; margin-bottom: 12px; text-align: center; border-bottom: 1px solid #f85149; padding-bottom: 8px;">🚨 總體經濟與熊市警訊指標</div>
                    <div style="padding: 10px; margin-bottom: 8px; background-color: #3c1818; border-radius: 6px; color: #ffffff; font-size: 14px; border-left: 6px solid #f85149;">₿ BTC 120MA <span style="color: #8b949e; margin: 0 5px;">|</span> <span style="color: #ff7b72;">{btc_txt}</span></div>
                    <div style="padding: 10px; margin-bottom: 8px; background-color: #3c1818; border-radius: 6px; color: #ffffff; font-size: 14px; border-left: 6px solid #f85149;">🥇 黃金 120MA <span style="color: #8b949e; margin: 0 5px;">|</span> <span style="color: #ff7b72;">{gold_txt}</span></div>
                    <div style="padding: 10px; margin-bottom: 8px; background-color: #3c1818; border-radius: 6px; color: #ffffff; font-size: 14px; border-left: 6px solid #f85149;">🏦 金融壓力 (STLFSI4) <span style="color: #8b949e; margin: 0 5px;">|</span> <span style="color: #ff7b72;">{stlfsi4_txt}</span></div>
                    <div style="padding: 10px; background-color: #3c1818; border-radius: 6px; color: #ffffff; font-size: 14px; border-left: 6px solid #f85149;">🛡️ 美國5年期CDS <span style="color: #8b949e; margin: 0 5px;">|</span> <span style="color: #ff7b72;">{cds_txt}</span></div>
                </div>
            </td>
        </tr>
        
        <tr>
            <td style="padding: 15px 0;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #30363d; border-radius: 8px; overflow: hidden; background-color: #161b22;">
                    <tr><td style="background-color: #21262d; padding: 15px; text-align: center; color: #ffffff; font-size: 18px; font-weight: bold; border-bottom: 1px solid #30363d;">🇹🇼 台股避險動能池 <br><span style="font-size: 12px; color: #8b949e; font-weight: normal;">更新時間: {state.get('tw_time')}</span></td></tr>
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
                    <tr><td style="background-color: #21262d; padding: 15px; text-align: center; color: #ffffff; font-size: 18px; font-weight: bold; border-bottom: 1px solid #30363d;">🇺🇸 美股避險動能池 <br><span style="font-size: 12px; color: #8b949e; font-weight: normal;">更新時間: {state.get('us_time')}</span></td></tr>
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

if run_tw and run_us: mail_subject = f"🌍 雙市場合併通知 (含總經警訊) - {today_str}"
elif run_tw: mail_subject = f"🌅 晨間量化通知 (含總經警訊) - {today_str}"
else: mail_subject = f"🌇 午後量化通知 (含總經警訊) - {today_str}"

send_email_notify(mail_subject, email_html)
