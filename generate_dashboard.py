import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import json
import re
import requests
import io

import platform as _platform
_script_dir  = os.path.dirname(os.path.abspath(__file__))
if _platform.system() == 'Windows':
    _default_out = r'C:\Users\309se\OneDrive\Desktop\클로드 폴더\자동_자금일보(일별)\자금일보_대쉬보드.html'
    out_path = _default_out if os.path.exists(os.path.dirname(_default_out)) else os.path.join(_script_dir, '자금일보_대쉬보드.html')
else:
    # GitHub Actions (Linux) - index.html 로 직접 저장
    out_path = os.path.join(_script_dir, 'index.html')

# ── Google Sheets 다운로드 ──────────────────────────────────
SHEET_ID = '1pLycRFzbv-mmK0Xdrm59rjpB8wKoPCS9Lp2iQqxiMEQ'
EXPORT_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx'

print('Google Sheets 다운로드 중...')
resp = requests.get(EXPORT_URL, timeout=30)
if resp.status_code != 200 or resp.content[:2] != b'PK':
    print(f'  오류: 시트 다운로드 실패 (status={resp.status_code})')
    sys.exit(1)
xl = pd.ExcelFile(io.BytesIO(resp.content))
all_sheet_names = xl.sheet_names
print(f'  다운로드 완료 - 시트: {all_sheet_names}')

# ──────────────────────────────────────────────────────────
# 1. 일별 자금일보 시트 파서 (포맷 자동 감지)
# ──────────────────────────────────────────────────────────
TARGET_ACCOUNTS = [
    '기업017','우리577','신한709','신한812','신한220',
    '국민930','국민323','국민962','하나104','하나652',
    '토스(시오레)','토스(친한스토어)','토스(박약다식몰)','토스(b2b)',
    '기업013','하나438',
]
INACTIVE_ACCOUNTS = ['국민312','기업031','토스','하나014']
USD_ACCOUNTS = {'기업013', '하나438'}

def parse_daily_sheet(df):
    """
    자금일보 형식 시트에서 계좌별 (기초, 입금, 출금, 잔액) 추출.
    헤더 행을 자동 탐색해 컬럼 위치를 찾는다.
    """
    begin_col = dep_col = wdr_col = end_col = None
    header_row_idx = None

    for i in range(min(15, len(df))):
        row = df.iloc[i]
        cols = {}
        for j, val in enumerate(row):
            s = str(val).replace(' ', '')
            if '전일잔액' in s:
                cols['begin'] = j
            elif s in ('입금', '입  금', '입금액') and 'begin' in cols:
                cols['dep'] = j
            elif s in ('출금', '출  금', '출금액') and 'dep' in cols:
                cols['wdr'] = j
            elif '당일잔액' in s and 'wdr' in cols:
                cols['end'] = j
        if all(k in cols for k in ('begin','dep','wdr','end')):
            begin_col, dep_col, wdr_col, end_col = cols['begin'], cols['dep'], cols['wdr'], cols['end']
            header_row_idx = i
            break

    if header_row_idx is None:
        return {}

    result = {}
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        # 이미 모든 계좌를 찾았으면 중단
        if len(result) == len(TARGET_ACCOUNTS):
            break
        for j, val in enumerate(row):
            acct = str(val).strip()
            if acct in TARGET_ACCOUNTS and acct not in result:
                # $ 기호·콤마 제거 후 숫자 변환 (USD 계좌: '$ 233,709.82' 형식 처리)
                def to_num(v):
                    if isinstance(v, str):
                        v = v.replace('$','').replace(',','').replace('-','0').strip()
                    return pd.to_numeric(v, errors='coerce')

                end_num = to_num(row.iloc[end_col])
                if pd.isna(end_num):
                    break  # 이 행은 숫자가 없는 다른 섹션, 건너뜀
                begin = to_num(row.iloc[begin_col])
                dep   = to_num(row.iloc[dep_col])
                wdr   = to_num(row.iloc[wdr_col])
                result[acct] = {
                    'begin': float(begin) if pd.notna(begin) else 0.0,
                    'dep':   float(dep)   if pd.notna(dep)   else 0.0,
                    'wdr':   float(wdr)   if pd.notna(wdr)   else 0.0,
                    'end':   float(end_num),
                }
                break
    return result

# ──────────────────────────────────────────────────────────
# 2. 모든 일별 자금일보 시트 수집 ('N월 M일' 패턴)
# ──────────────────────────────────────────────────────────
daily_sheets = {}  # {date_str: {account: {begin,dep,wdr,end}}}

# 연도 기준: 시트1 거래 데이터의 가장 최근 날짜에서 추출
def detect_year_from_sheet1(xl):
    try:
        df_tmp = xl.parse('시트1', header=None)
        dates = df_tmp.iloc[2:, 0].dropna().astype(str)
        dates = [d.strip() for d in dates if re.match(r'\d{4}\.\d{2}\.\d{2}', d.strip())]
        return int(max(dates).split('.')[0]) if dates else 2026
    except Exception:
        return 2026

jg_year = detect_year_from_sheet1(xl)
print(f'연도 기준: {jg_year}년')

# 'N월 M일' 패턴 시트들
for sheet in all_sheet_names:
    m = re.match(r'^(\d+)월\s*(\d+)일$', sheet.strip())
    if not m:
        continue
    s_month, s_day = int(m.group(1)), int(m.group(2))
    s_year = jg_year
    date_str = f'{s_year}.{s_month:02d}.{s_day:02d}'
    df_s = xl.parse(sheet, header=None)
    daily_sheets[date_str] = parse_daily_sheet(df_s)
    print(f'  일별시트 로드: {sheet} → {date_str}, 계좌수={len(daily_sheets[date_str])}')

print(f'일별시트 날짜: {sorted(daily_sheets.keys())}')

# ──────────────────────────────────────────────────────────
# 3. 시트1 읽기 (월누적 거래내역)
# ──────────────────────────────────────────────────────────
df_raw = xl.parse('시트1', header=None)
# 0행: 헤더, 1행: 전기이월(있을 수도 없을 수도) → 전기이월 필터로 처리
df = df_raw.iloc[1:].copy()
df = df.rename(columns={0:'date',1:'account',2:'category',3:'description',4:'counterparty',5:'deposit',6:'withdrawal'})
df = df[['date','account','category','description','counterparty','deposit','withdrawal']].copy()
df = df[df['date'].notna() & (df['date'].astype(str).str.strip() != '전기이월')].copy()
df['deposit']     = pd.to_numeric(df['deposit'],     errors='coerce').fillna(0)
df['withdrawal']  = pd.to_numeric(df['withdrawal'],  errors='coerce').fillna(0)
df['date']        = df['date'].astype(str).str.strip()
df['account']     = df['account'].astype(str).str.strip()
df['category']    = df['category'].fillna('').astype(str)
df['description'] = df['description'].fillna('').astype(str)
df['counterparty']= df['counterparty'].fillna('').astype(str)

월누적_dates = set(df['date'].unique())
all_dates = sorted(월누적_dates)

# ──────────────────────────────────────────────────────────
# 4. 앵커 결정: 월누적에 포함된 날짜 중 가장 최근 일별시트
# ──────────────────────────────────────────────────────────
anchor_candidates = [d for d in daily_sheets if d in 월누적_dates]
if anchor_candidates:
    anchor_date = max(anchor_candidates)
elif daily_sheets:
    anchor_date = max(daily_sheets.keys())  # 월누적에 없어도 가장 최근 일별시트 사용
else:
    print('오류: 앵커로 쓸 일별시트가 없습니다.')
    sys.exit(1)

anchor_sheet = daily_sheets[anchor_date]
current_balances = {}
for acct in TARGET_ACCOUNTS:
    val = anchor_sheet.get(acct, {}).get('end', 0.0)
    current_balances[acct] = float(val)
for acct in INACTIVE_ACCOUNTS:
    current_balances[acct] = 0.0

all_accounts = list(current_balances.keys())
print(f'앵커 날짜: {anchor_date}')

# ──────────────────────────────────────────────────────────
# 5. 계좌별 일별 잔액 계산 (앵커 기준 역산)
# ──────────────────────────────────────────────────────────
acct_daily = {}
for acct in all_accounts:
    acct_df = df[df['account'] == acct]
    if len(acct_df) == 0:
        acct_daily[acct] = {}
        continue
    daily = acct_df.groupby('date').agg(dep=('deposit','sum'), wdr=('withdrawal','sum')).to_dict('index')
    acct_daily[acct] = daily

# 표시 순서 (자금일보 원본과 동일)
ACCOUNT_ORDER = [
    '기업017','우리577','신한709','신한812','신한220',
    '국민930','국민323','국민962','하나104','하나652',
    '토스(시오레)','토스(친한스토어)','토스(박약다식몰)','토스(b2b)',
    '기업013','하나438',
]

acct_end_bal    = {}
acct_initial_bal = {}  # 거래 이전 초기 잔액

for acct in all_accounts:
    daily   = acct_daily[acct]
    cur_bal = current_balances[acct]
    if not daily:
        acct_end_bal[acct]     = {}
        acct_initial_bal[acct] = cur_bal
        continue
    sorted_dates = sorted(daily.keys())
    if sorted_dates and sorted_dates[0] > anchor_date:
        # 앵커가 모든 거래보다 이전 → 앵커 잔액이 곧 초기 잔액 (순방향 계산)
        initial_bal = cur_bal
    else:
        # 앵커가 거래 중간 또는 이후 → 역산 (기존 방식)
        cum_net = sum(daily[d]['dep'] - daily[d]['wdr'] for d in sorted_dates)
        initial_bal = cur_bal - cum_net
    acct_initial_bal[acct] = initial_bal
    running = initial_bal
    end_bals = {}
    for d in sorted_dates:
        running += daily[d]['dep'] - daily[d]['wdr']
        end_bals[d] = running
    acct_end_bal[acct] = end_bals

def get_end_bal_on(acct, date):
    """해당 날짜 당일잔액 반환 (거래 없는 날은 직전 잔액 유지)"""
    end_bals = acct_end_bal[acct]
    if date in end_bals:
        return end_bals[date]
    earlier = [d for d in end_bals if d <= date]
    if earlier:
        return end_bals[max(earlier)]
    return acct_initial_bal[acct]  # 첫 거래 이전 시점

# ──────────────────────────────────────────────────────────
# 5b. USD/KRW 환율 데이터 (서울외국환중개 매매기준율)
# ──────────────────────────────────────────────────────────
def fetch_smbs_rates(all_dates):
    """SMBS에서 월별 USD/KRW 매매기준율 가져오기 (Playwright 사용)"""
    import calendar as _cal
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print('  playwright 미설치 → pip install playwright && playwright install chromium')
        return {}

    if not all_dates:
        return {}

    # 필요한 월 범위 계산
    start_date = min(all_dates)
    end_date   = max(all_dates)
    sy, sm = int(start_date[:4]), int(start_date[5:7])
    ey, em = int(end_date[:4]),   int(end_date[5:7])
    # 현재 월도 포함 (오늘 환율)
    import datetime
    today = datetime.date.today()
    if (today.year, today.month) > (ey, em):
        ey, em = today.year, today.month

    rates = {}
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            pg = browser.new_page()
            yr, mo = sy, sm
            while (yr, mo) <= (ey, em):
                last_day = _cal.monthrange(yr, mo)[1]
                url = (f'http://www.smbs.biz/ExRate/StdExRate.jsp'
                       f'?StrSch_sYear={yr}&StrSch_sMonth={mo:02d}&StrSch_sDay=01'
                       f'&StrSch_eYear={yr}&StrSch_eMonth={mo:02d}&StrSch_eDay={last_day}'
                       f'&tongwha_code=USD')
                pg.goto(url, timeout=30000)
                pg.wait_for_load_state('networkidle')
                pg.wait_for_timeout(1500)

                for row in pg.query_selector_all('table tr'):
                    cols = [c.strip() for c in row.inner_text().split('\t') if c.strip()]
                    if len(cols) >= 3 and re.match(r'\d{4}\.\d{2}\.\d{2}', cols[0]):
                        try:
                            rates[cols[0]] = float(cols[2].replace(',', ''))
                        except:
                            pass
                print(f'  SMBS {yr}.{mo:02d}: {sum(1 for d in rates if d.startswith(f"{yr}.{mo:02d}"))}건')
                mo += 1
                if mo > 12:
                    yr, mo = yr + 1, 1
            browser.close()
    except Exception as e:
        print(f'  SMBS 오류: {e}')
    return rates

print('USD/KRW 환율 데이터 가져오는 중 (SMBS 매매기준율)...')
usd_krw_rates = fetch_smbs_rates(all_dates)

# Fallback: 네이버 금융
if not usd_krw_rates:
    print('  SMBS 실패 → 네이버 금융으로 대체')
    min_needed = min(all_dates) if all_dates else '2026.01.01'
    h_nav = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    for pg in range(1, 50):
        try:
            r = requests.get('https://finance.naver.com/marketindex/exchangeDailyQuote.naver',
                             params={'marketindexCd':'FX_USDKRW','page':str(pg)},
                             headers=h_nav, timeout=10)
            r.encoding = 'utf-8'
            nums = re.findall(r'(\d{4}\.\d{2}\.\d{2})[^<]*</td>\s*<td[^>]*>([\d,\.]+)', r.text)
            if not nums: break
            for ds, rs in nums:
                usd_krw_rates[ds] = float(rs.replace(',',''))
            if usd_krw_rates and min(usd_krw_rates.keys()) <= min_needed: break
        except: break

print(f'환율 데이터: {len(usd_krw_rates)}건')
if usd_krw_rates:
    for sd in sorted(usd_krw_rates.keys(), reverse=True)[:3]:
        print(f'  {sd}: {usd_krw_rates[sd]:,.2f}')

def get_rate_on(date_str):
    """해당 날짜의 USD/KRW 환율 반환. 주말·공휴일은 직전 거래일 환율."""
    if date_str in usd_krw_rates:
        return usd_krw_rates[date_str]
    earlier = [d for d in usd_krw_rates if d <= date_str]
    if earlier:
        return usd_krw_rates[max(earlier)]
    return None

def calc_krw(balance, acct, date_str):
    """잔액을 원화로 환산. USD 계좌는 환율 적용, KRW 계좌는 그대로."""
    if acct not in USD_ACCOUNTS:
        return round(balance)
    rate = get_rate_on(date_str)
    if rate is None:
        return None
    return round(balance * rate)

# ──────────────────────────────────────────────────────────
# 6. 월누적 기반 대시보드 데이터 생성 (모든 계좌 표시)
# ──────────────────────────────────────────────────────────
dashboard_data = {}
for date in all_dates:
    summary = []
    total_dep_krw = 0
    total_wdr_krw = 0
    for acct in ACCOUNT_ORDER:
        daily    = acct_daily.get(acct, {})
        dep      = daily.get(date, {}).get('dep', 0)
        wdr      = daily.get(date, {}).get('wdr', 0)
        end_bal  = get_end_bal_on(acct, date)
        begin_bal = end_bal - dep + wdr

        # 잔액이 0이고 거래도 없으면 비활성 계좌 → 제외
        if begin_bal == 0 and end_bal == 0 and dep == 0 and wdr == 0:
            continue

        is_usd = acct in USD_ACCOUNTS
        krw_end = calc_krw(end_bal, acct, date)

        # 원화 환산 입출금 (총계용)
        if is_usd:
            rate = get_rate_on(date)
            if rate:
                total_dep_krw += round(dep * rate)
                total_wdr_krw += round(wdr * rate)
        else:
            total_dep_krw += round(dep)
            total_wdr_krw += round(wdr)

        summary.append({
            'account':    acct,
            'begin':      round(begin_bal),
            'deposit':    round(dep),
            'withdrawal': round(wdr),
            'end':        round(end_bal),
            'is_usd':     is_usd,
            'krw_end':    krw_end,
        })

    day_df   = df[df['date'] == date]
    dep_list = day_df[day_df['deposit']   > 0][['account','category','description','counterparty','deposit']].copy()
    wdr_list = day_df[day_df['withdrawal'] > 0][['account','category','description','counterparty','withdrawal']].copy()
    dep_list['deposit']    = dep_list['deposit'].apply(lambda x: int(round(x)))
    wdr_list['withdrawal'] = wdr_list['withdrawal'].apply(lambda x: int(round(x)))

    dashboard_data[date] = {
        'summary':          summary,
        'deposits':         dep_list.to_dict('records'),
        'withdrawals':      wdr_list.to_dict('records'),
        'total_deposit':    total_dep_krw,
        'total_withdrawal': total_wdr_krw,
        'usd_rate':         get_rate_on(date),
    }

# ──────────────────────────────────────────────────────────
# 7. 일별시트 전용 날짜 추가 (월누적에 없는 날짜)
# ──────────────────────────────────────────────────────────
for date_str, sheet_data in daily_sheets.items():
    if date_str in 월누적_dates:
        continue
    if not sheet_data:
        continue
    summary = []
    total_dep_krw = 0
    total_wdr_krw = 0
    for acct in ACCOUNT_ORDER:
        vals = sheet_data.get(acct)
        if not vals:
            continue
        begin = vals.get('begin', 0)
        dep   = vals.get('dep',   0)
        wdr   = vals.get('wdr',   0)
        end   = vals.get('end',   0)
        # 모두 0이면 제외
        if begin == 0 and end == 0 and dep == 0 and wdr == 0:
            continue
        is_usd = acct in USD_ACCOUNTS
        krw_end = calc_krw(end, acct, date_str)

        if is_usd:
            rate = get_rate_on(date_str)
            if rate:
                total_dep_krw += round(dep * rate)
                total_wdr_krw += round(wdr * rate)
        else:
            total_dep_krw += round(dep)
            total_wdr_krw += round(wdr)

        summary.append({
            'account':    acct,
            'begin':      round(begin),
            'deposit':    round(dep),
            'withdrawal': round(wdr),
            'end':        round(end),
            'is_usd':     is_usd,
            'krw_end':    krw_end,
        })
    if not summary:
        continue
    dashboard_data[date_str] = {
        'summary': summary,
        'deposits': [],
        'withdrawals': [],
        'total_deposit':    total_dep_krw,
        'total_withdrawal': total_wdr_krw,
        'usd_rate':         get_rate_on(date_str),
        'summary_only': True,
    }
    print(f'  일별시트 전용 날짜 추가: {date_str}')

json_str = json.dumps(dashboard_data, ensure_ascii=False)

# 거래 있는 날짜 목록 (월누적 + 일별시트)
dates_with_tx = sorted(dashboard_data.keys())

USD_ACCOUNTS_JS = json.dumps(sorted(USD_ACCOUNTS))
LAST_DATE = dates_with_tx[-1] if dates_with_tx else '2026.05.27'

HTML = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>자금일보 대시보드</title>
<style>
  /* ── 인증 오버레이 ── */
  #auth-overlay {{
    display: flex; position: fixed; inset: 0; z-index: 9999;
    background: #0F0F10; flex-direction: column;
    align-items: center; justify-content: center; gap: 14px;
  }}
  .auth-logo {{ width: 56px; height: 56px; background: #3182F6; border-radius: 16px;
    display: flex; align-items: center; justify-content: center;
    font-size: 26px; font-weight: 800; color: white; margin-bottom: 4px; }}
  .auth-title {{ font-size: 20px; font-weight: 700; color: #F2F2F7; }}
  .auth-sub {{ font-size: 13px; color: #636366; margin-bottom: 4px; }}
  .auth-input {{ background: #1C1C1E; border: 1px solid #3A3A3C; border-radius: 10px;
    padding: 13px 16px; font-size: 14px; color: #F2F2F7; width: 280px;
    font-family: inherit; outline: none; transition: border 0.15s; text-align: center; letter-spacing: 2px; }}
  .auth-input:focus {{ border-color: #3182F6; }}
  .auth-input.error {{ border-color: #FF453A; animation: shake 0.3s ease; }}
  @keyframes shake {{
    0%,100% {{ transform: translateX(0); }}
    25% {{ transform: translateX(-8px); }}
    75% {{ transform: translateX(8px); }}
  }}
  .auth-btn {{ background: #3182F6; color: white; border: none; border-radius: 12px;
    padding: 13px 0; font-size: 14px; font-weight: 700; cursor: pointer;
    font-family: inherit; transition: all 0.15s; width: 280px; }}
  .auth-btn:hover {{ background: #2563d4; transform: translateY(-1px);
    box-shadow: 0 4px 16px rgba(49,130,246,0.45); }}
  .auth-error {{ color: #FF453A; font-size: 12px; height: 16px; }}
  #main-content {{ display: none; }}
  .lock-btn {{ background: transparent; color: #636366; border: 1px solid #3A3A3C;
    border-radius: 20px; padding: 5px 12px; font-size: 11px; font-weight: 600;
    cursor: pointer; font-family: inherit; transition: all 0.15s; }}
  .lock-btn:hover {{ color: #F2F2F7; border-color: #636366; }}

  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700&display=swap');
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif; background: #0F0F10; color: #F2F2F7; font-size: 12px; }}

  /* ── 헤더 ── */
  .header {{ background: #0F0F10; border-bottom: 1px solid #1C1C1E; padding: 12px 20px; display: flex; align-items: center; justify-content: space-between; }}
  .header-left {{ display: flex; align-items: center; gap: 10px; }}
  .header-logo {{ width: 28px; height: 28px; background: #3182F6; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 14px; font-weight: 800; color: white; letter-spacing: -1px; }}
  .header h1 {{ font-size: 14px; font-weight: 700; color: #F2F2F7; }}
  .header .sub {{ font-size: 11px; color: #636366; }}

  /* 데이터 불러오기 버튼 */
  .refresh-btn {{ display: flex; align-items: center; gap: 7px; background: #3182F6; color: white; border: none; border-radius: 20px; padding: 7px 16px; font-size: 12px; font-weight: 700; cursor: pointer; font-family: inherit; transition: all 0.15s; white-space: nowrap; }}
  .refresh-btn:hover:not(:disabled) {{ background: #2563d4; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(49,130,246,0.4); }}
  .refresh-btn:disabled {{ opacity: 0.5; cursor: not-allowed; transform: none; }}
  .refresh-btn .btn-icon {{ font-size: 13px; }}
  .refresh-btn.loading .btn-icon {{ display: inline-block; animation: spin 0.8s linear infinite; }}
  @keyframes spin {{ from {{ transform: rotate(0deg); }} to {{ transform: rotate(360deg); }} }}

  /* 토스트 알림 */
  .toast {{ position: fixed; bottom: 24px; right: 24px; background: #1C1C1E; color: #F2F2F7; padding: 12px 18px; border-radius: 12px; font-size: 12px; font-weight: 600; box-shadow: 0 8px 24px rgba(0,0,0,0.4); z-index: 9999; opacity: 0; transform: translateY(10px); transition: all 0.25s; pointer-events: none; display: flex; align-items: center; gap: 8px; }}
  .toast.show {{ opacity: 1; transform: translateY(0); }}
  .toast.success {{ border-left: 3px solid #05C072; }}
  .toast.error {{ border-left: 3px solid #FF453A; }}

  /* ── 레이아웃 ── */
  .container {{ display: flex; gap: 12px; padding: 12px 20px; height: calc(100vh - 53px); overflow: hidden; }}

  /* ── 캘린더 ── */
  .calendar-panel {{ width: 236px; flex-shrink: 0; display: flex; flex-direction: column; gap: 8px; overflow-y: auto; }}
  .calendar-card {{ background: #1C1C1E; border-radius: 14px; overflow: hidden; }}
  .year-selector {{ padding: 10px 12px 0; display: flex; gap: 5px; flex-wrap: wrap; }}
  .year-btn {{ padding: 3px 9px; border: 1px solid #3A3A3C; border-radius: 20px; cursor: pointer; font-size: 11px; color: #8E8E93; background: transparent; font-family: inherit; transition: all 0.15s; }}
  .year-btn.active {{ background: #3182F6; color: white; border-color: #3182F6; font-weight: 600; }}
  .year-btn:hover:not(.active) {{ border-color: #636366; color: #F2F2F7; }}
  .cal-header {{ padding: 10px 12px 8px; display: flex; align-items: center; justify-content: space-between; }}
  .cal-header .month-title {{ font-size: 13px; font-weight: 700; color: #F2F2F7; }}
  .cal-nav {{ background: #2C2C2E; border: none; color: #8E8E93; font-size: 14px; cursor: pointer; width: 24px; height: 24px; border-radius: 6px; display: flex; align-items: center; justify-content: center; transition: background 0.15s; }}
  .cal-nav:hover {{ background: #3A3A3C; color: #F2F2F7; }}
  .cal-weekdays {{ display: grid; grid-template-columns: repeat(7,1fr); padding: 0 6px; }}
  .cal-weekday {{ text-align: center; padding: 4px 0; font-size: 10px; font-weight: 600; color: #48484A; }}
  .cal-weekday.sun {{ color: #FF453A; }}
  .cal-weekday.sat {{ color: #3182F6; }}
  .cal-days {{ display: grid; grid-template-columns: repeat(7,1fr); padding: 4px 6px 10px; gap: 1px; }}
  .cal-day {{ text-align: center; padding: 4px 2px; font-size: 11px; border-radius: 7px; cursor: default; color: #48484A; min-height: 28px; display: flex; align-items: center; justify-content: center; flex-direction: column; transition: background 0.12s; }}
  .cal-day.has-data {{ cursor: pointer; color: #EBEBF5; font-weight: 600; }}
  .cal-day.has-data:hover {{ background: #2C2C2E; }}
  .cal-day.selected {{ background: #3182F6 !important; color: white !important; font-weight: 700; }}
  .cal-day.today:not(.selected) {{ background: #1C3A5E; }}
  .cal-day.sunday.has-data {{ color: #FF6B6B; }}
  .cal-day.saturday.has-data {{ color: #5E9CF5; }}
  .cal-day .dot {{ width: 3px; height: 3px; background: #3182F6; border-radius: 50%; margin-top: 2px; }}
  .cal-day.selected .dot {{ background: rgba(255,255,255,0.6); }}
  .cal-day.sunday.selected, .cal-day.saturday.selected {{ color: white !important; }}

  /* ── 리포트 패널 ── */
  .report-panel {{ flex: 1; display: flex; flex-direction: column; gap: 10px; overflow-y: auto; min-width: 0; }}

  .section-card {{ background: #1C1C1E; border-radius: 14px; overflow: hidden; flex-shrink: 0; }}
  .section-header {{ padding: 12px 14px 10px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #2C2C2E; }}
  .section-header h2 {{ font-size: 12px; font-weight: 700; color: #EBEBF5; }}
  .section-header .date-badge {{ background: #1C3A5E; color: #5E9CF5; padding: 2px 9px; border-radius: 20px; font-size: 11px; font-weight: 600; }}
  .section-body {{ padding: 10px 14px; }}

  /* 환율 배지 */
  .rate-badge {{ background: #2C2C2E; color: #8E8E93; padding: 2px 9px; border-radius: 20px; font-size: 10px; font-weight: 500; }}
  .rate-badge span {{ color: #F5A623; font-weight: 700; }}

  /* 안내 메시지 */
  .empty-msg {{ text-align: center; padding: 50px 20px; color: #48484A; }}
  .empty-msg .icon {{ font-size: 36px; margin-bottom: 10px; }}
  .empty-msg p {{ font-size: 12px; line-height: 1.6; }}

  /* 요약 통계 카드 */
  .stat-row {{ display: flex; gap: 8px; }}
  .stat-box {{ flex: 1; background: #2C2C2E; border-radius: 10px; padding: 10px 12px; }}
  .stat-box .label {{ font-size: 10px; color: #636366; margin-bottom: 5px; font-weight: 500; }}
  .stat-box .value {{ font-size: 15px; font-weight: 700; letter-spacing: -0.3px; }}
  .stat-box.dep .value {{ color: #05C072; }}
  .stat-box.wdr .value {{ color: #FF453A; }}
  .stat-box.net .value {{ color: #3182F6; }}

  /* ── 테이블 공통 ── */
  table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
  th {{ background: #2C2C2E; color: #8E8E93; padding: 6px 9px; text-align: center; font-weight: 600; white-space: nowrap; font-size: 10px; text-transform: uppercase; letter-spacing: 0.3px; }}
  td {{ padding: 6px 9px; border-bottom: 1px solid #1C1C1E; color: #EBEBF5; }}
  td.num {{ text-align: right; white-space: nowrap; }}
  td.center {{ text-align: center; }}
  tr:hover td {{ background: #2C2C2E; }}
  tr.total-row td {{ background: #1C3A5E; color: #5E9CF5; font-weight: 700; border-bottom: none; }}
  tr.total-row:hover td {{ background: #243F6A; }}

  .deposit-amount {{ color: #05C072; font-weight: 600; }}
  .withdrawal-amount {{ color: #FF453A; font-weight: 600; }}
  .balance-amount {{ color: #EBEBF5; font-weight: 700; }}
  .krw-amount {{ color: #F5A623; font-weight: 600; }}
  .usd-badge {{ display: inline-block; background: #2C3E50; color: #5E9CF5; border-radius: 3px; padding: 0px 4px; font-size: 9px; font-weight: 700; margin-left: 3px; vertical-align: middle; }}

  /* 탭 */
  .tabs {{ display: flex; gap: 0; border-bottom: 1px solid #2C2C2E; margin-bottom: 10px; }}
  .tab {{ padding: 7px 14px; cursor: pointer; font-size: 11px; font-weight: 600; color: #636366; border-bottom: 2px solid transparent; margin-bottom: -1px; transition: color 0.15s; font-family: inherit; background: none; border-top: none; border-left: none; border-right: none; }}
  .tab.active {{ color: #3182F6; border-bottom-color: #3182F6; }}
  .tab:hover:not(.active) {{ color: #8E8E93; }}
  .tab-content {{ display: none; }}
  .tab-content.active {{ display: block; }}
  .no-data {{ text-align: center; padding: 20px; color: #48484A; font-size: 11px; }}

  /* 입출금 테이블 */
  .tx-table {{ table-layout: fixed; }}
  .tx-table td {{ overflow: hidden; white-space: nowrap; text-overflow: ellipsis; max-width: 0; }}
  .tx-table td.desc {{ white-space: normal; word-break: break-word; max-width: 0; line-height: 1.4; }}
  .tx-table td.center {{ text-align: center; }}

  /* 정렬 헤더 */
  .sortable-th {{ cursor: pointer; user-select: none; }}
  .sortable-th:hover {{ background: #3A3A3C; color: #F2F2F7; }}
  .sort-arrow {{ font-size: 8px; margin-left: 3px; opacity: 0.3; }}
  .sort-arrow.active {{ opacity: 1; color: #3182F6; }}

  /* 스크롤바 */
  ::-webkit-scrollbar {{ width: 4px; height: 4px; }}
  ::-webkit-scrollbar-track {{ background: transparent; }}
  ::-webkit-scrollbar-thumb {{ background: #3A3A3C; border-radius: 2px; }}

  @media (max-width: 900px) {{
    .container {{ flex-direction: column; height: auto; overflow: visible; }}
    .calendar-panel {{ width: 100%; overflow-y: visible; }}
  }}
</style>
</head>
<body>

<!-- 인증 오버레이 -->
<div id="auth-overlay">
  <div class="auth-logo">자</div>
  <div class="auth-title">자금일보 대시보드</div>
  <div class="auth-sub">파마브로스 임직원 전용</div>
  <input class="auth-input" type="password" id="pwInput"
         placeholder="비밀번호를 입력하세요"
         onkeydown="if(event.key==='Enter') checkPw()">
  <button class="auth-btn" onclick="checkPw()">입장하기</button>
  <div class="auth-error" id="pwError"></div>
</div>

<!-- 메인 콘텐츠 (로그인 후 표시) -->
<div id="main-content">

<div class="header">
  <div class="header-left">
    <div class="header-logo">자</div>
    <div>
      <h1>자금일보</h1>
      <div class="sub">파마브로스 · 일별 자금 현황</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:8px;">
    <button class="lock-btn" onclick="lockDashboard()">🔒 잠금</button>
    <button class="refresh-btn" id="refreshBtn" onclick="refreshData()">
      <span class="btn-icon">↻</span>
      <span class="btn-label">데이터 불러오기</span>
    </button>
  </div>
</div>
<div class="toast" id="toast"></div>

<div class="container">
  <!-- 캘린더 -->
  <div class="calendar-panel">
    <div class="calendar-card">
      <div class="year-selector" id="yearSelector"></div>
      <div class="cal-header">
        <button class="cal-nav" onclick="changeMonth(-1)">&#8249;</button>
        <span class="month-title" id="monthTitle"></span>
        <button class="cal-nav" onclick="changeMonth(1)">&#8250;</button>
      </div>
      <div class="cal-weekdays">
        <div class="cal-weekday sun">일</div>
        <div class="cal-weekday">월</div>
        <div class="cal-weekday">화</div>
        <div class="cal-weekday">수</div>
        <div class="cal-weekday">목</div>
        <div class="cal-weekday">금</div>
        <div class="cal-weekday sat">토</div>
      </div>
      <div class="cal-days" id="calDays"></div>
    </div>
  </div>

  <!-- 리포트 -->
  <div class="report-panel" id="reportPanel">
    <div class="section-card">
      <div class="empty-msg">
        <div class="icon">📅</div>
        <p>날짜를 선택하면<br>해당일 자금 현황이 표시됩니다.</p>
      </div>
    </div>
  </div>
</div>

<script>
const DATA = {json_str};
const TX_DATES = new Set({json.dumps(dates_with_tx)});
const USD_ACCOUNTS = new Set({USD_ACCOUNTS_JS});

let currentYear, currentMonth, selectedDate = null;

// 초기 날짜: 마지막 거래일
const lastDate = '{LAST_DATE}';
const parts = lastDate.split('.');
currentYear = parseInt(parts[0]);
currentMonth = parseInt(parts[1]);

function fmt(n) {{
  if (n === null || n === undefined) return '-';
  return Math.round(n).toLocaleString('ko-KR');
}}
function fmtDate(d) {{
  const p = d.split('.');
  return p[0] + '년 ' + p[1] + '월 ' + p[2] + '일';
}}
function fmtAmt(n, isUsd) {{
  if (n === null || n === undefined) return '-';
  const prefix = isUsd ? '$' : '₩';
  return prefix + Math.abs(Math.round(n)).toLocaleString('ko-KR');
}}
function fmtKrw(n) {{
  if (n === null || n === undefined) return '<span style="color:#48484A">-</span>';
  return '₩' + Math.round(n).toLocaleString('ko-KR');
}}

function renderYearSelector() {{
  const years = [...new Set([...TX_DATES].map(d => parseInt(d.split('.')[0])))].sort();
  const el = document.getElementById('yearSelector');
  el.innerHTML = years.map(y =>
    `<button class="year-btn ${{y===currentYear?'active':''}}" onclick="selectYear(${{y}})">${{y}}년</button>`
  ).join('');
}}

function selectYear(y) {{
  currentYear = y;
  renderYearSelector();
  renderCalendar();
}}

function changeMonth(delta) {{
  currentMonth += delta;
  if (currentMonth < 1) {{ currentMonth = 12; currentYear--; }}
  if (currentMonth > 12) {{ currentMonth = 1; currentYear++; }}
  renderYearSelector();
  renderCalendar();
}}

function renderCalendar() {{
  document.getElementById('monthTitle').textContent = currentYear + '년 ' + String(currentMonth).padStart(2,'0') + '월';
  const firstDay = new Date(currentYear, currentMonth-1, 1).getDay();
  const daysInMonth = new Date(currentYear, currentMonth, 0).getDate();
  const today = new Date();
  const todayStr = today.getFullYear() + '.' + String(today.getMonth()+1).padStart(2,'0') + '.' + String(today.getDate()).padStart(2,'0');

  let html = '';
  for (let i = 0; i < firstDay; i++) html += '<div class="cal-day"></div>';

  for (let d = 1; d <= daysInMonth; d++) {{
    const dateStr = currentYear + '.' + String(currentMonth).padStart(2,'0') + '.' + String(d).padStart(2,'0');
    const dow = new Date(currentYear, currentMonth-1, d).getDay();
    const hasTx = TX_DATES.has(dateStr);
    const isSelected = dateStr === selectedDate;
    const isToday = dateStr === todayStr;

    let cls = 'cal-day';
    if (hasTx) cls += ' has-data';
    if (isSelected) cls += ' selected';
    if (isToday) cls += ' today';
    if (dow === 0) cls += ' sunday';
    if (dow === 6) cls += ' saturday';

    const dot = hasTx && !isSelected ? '<div class="dot"></div>' : '';
    const clickFn = hasTx ? `onclick="selectDate('${{dateStr}}')"` : '';
    html += `<div class="${{cls}}" ${{clickFn}}>${{d}}${{dot}}</div>`;
  }}

  document.getElementById('calDays').innerHTML = html;
}}

function selectDate(date) {{
  selectedDate = date;
  sortState = {{ dep: {{ key: null, dir: 1 }}, wdr: {{ key: null, dir: 1 }} }};
  renderCalendar();
  renderReport(date);
}}

let sortState = {{ dep: {{ key: null, dir: 1 }}, wdr: {{ key: null, dir: 1 }} }};

function renderTxTable(tableId, rows, amountField, amountClass, date) {{
  const state = sortState[tableId];
  let sorted = [...rows];
  if (state.key) {{
    sorted.sort((a, b) => {{
      let va = a[state.key] ?? '';
      let vb = b[state.key] ?? '';
      if (typeof va === 'number' && typeof vb === 'number') return state.dir * (va - vb);
      return state.dir * String(va).localeCompare(String(vb), 'ko');
    }});
  }}
  const total = rows.reduce((s, r) => s + (r[amountField] || 0), 0);

  const cols = [
    {{ key: 'account',      label: '계좌',    width: '80px'  }},
    {{ key: 'category',     label: '계정과목', width: '100px' }},
    {{ key: 'description',  label: '적 요',   width: '160px' }},
    {{ key: amountField,    label: '금 액',   width: '130px' }},
    {{ key: 'counterparty', label: '거래처',  width: '150px' }},
  ];

  const headerCols = cols.map(c => {{
    const isActive = state.key === c.key;
    const arrow = isActive ? (state.dir === 1 ? '▲' : '▼') : '⇅';
    const wStyle = c.width ? `width:${{c.width}};` : '';
    return `<th class="sortable-th" style="${{wStyle}}" onclick="sortTx('${{tableId}}','${{c.key}}','${{date}}')">${{c.label}} <span class="sort-arrow${{isActive?' active':''}}">${{arrow}}</span></th>`;
  }}).join('');

  let bodyRows = '';
  const isSummaryOnly = DATA[date] && DATA[date].summary_only;
  if (sorted.length === 0) {{
    bodyRows = isSummaryOnly
      ? `<tr><td colspan="5" class="no-data" style="color:#636366;">개별 거래내역은 월누적 시트 반영 후 확인됩니다.</td></tr>`
      : `<tr><td colspan="5" class="no-data">${{amountField === 'deposit' ? '입금' : '출금'}} 내역 없음</td></tr>`;
  }} else {{
    sorted.forEach(row => {{
      const desc = row.description || '';
      const isUsdAcct = USD_ACCOUNTS.has(row.account);
      const amtDisplay = isUsdAcct ? '$' + Math.round(row[amountField]).toLocaleString('ko-KR') : '₩' + Math.round(row[amountField]).toLocaleString('ko-KR');
      bodyRows += `<tr>
        <td class="center" title="${{row.account||''}}">${{row.account||''}}</td>
        <td class="center" title="${{row.category||''}}">${{row.category||''}}</td>
        <td class="desc" title="${{desc}}">${{desc}}</td>
        <td class="num ${{amountClass}}">${{amtDisplay}}</td>
        <td title="${{row.counterparty||''}}">${{row.counterparty||''}}</td>
      </tr>`;
    }});
    bodyRows += `<tr class="total-row">
      <td class="center" colspan="3">합 계 (원화환산)</td>
      <td class="num">₩${{fmt(total)}}</td>
      <td></td>
    </tr>`;
  }}

  return `<table class="tx-table">
    <thead><tr>${{headerCols}}</tr></thead>
    <tbody>${{bodyRows}}</tbody>
  </table>`;
}}

function sortTx(tableId, key, date) {{
  const state = sortState[tableId];
  state.dir = state.key === key ? -state.dir : 1;
  state.key = key;
  const d = DATA[date];
  if (!d) return;
  const rows       = tableId === 'dep' ? d.deposits    : d.withdrawals;
  const amtField   = tableId === 'dep' ? 'deposit'     : 'withdrawal';
  const amtClass   = tableId === 'dep' ? 'deposit-amount' : 'withdrawal-amount';
  document.getElementById('table-' + tableId).innerHTML = renderTxTable(tableId, rows, amtField, amtClass, date);
}}

function renderReport(date) {{
  const d = DATA[date];
  if (!d) return;

  const totalDep = d.total_deposit;
  const totalWdr = d.total_withdrawal;
  const netFlow = totalDep - totalWdr;

  // 환율 배지 - 모든 날짜에 표시
  const usdRate = d.usd_rate;
  const rateBadge = usdRate
    ? `<span class="rate-badge">USD/KRW <span>${{usdRate.toFixed(2)}}</span></span>`
    : '';

  // 계좌 요약 테이블
  let summaryRows = '';
  let sumKrwBegin = 0, sumKrwDep = 0, sumKrwWdr = 0, sumKrwEnd = 0;

  (d.summary || []).forEach(row => {{
    const isUsd = row.is_usd;
    const krwEnd = row.krw_end;

    // 원화 합계 계산
    if (!isUsd) {{
      sumKrwBegin += row.begin || 0;
      sumKrwDep   += row.deposit || 0;
      sumKrwWdr   += row.withdrawal || 0;
      sumKrwEnd   += row.end || 0;
    }} else if (krwEnd !== null) {{
      // USD 계좌: 원화환산 잔액만 원화잔액 합계에 포함
      sumKrwEnd += krwEnd;
    }}

    const usdBadge = isUsd ? '<span class="usd-badge">USD</span>' : '';
    const krwCell = `<td class="num krw-amount">${{fmtKrw(krwEnd)}}</td>`;

    summaryRows += `<tr>
      <td class="center">${{row.account}}${{usdBadge}}</td>
      <td class="num balance-amount">${{fmtAmt(row.begin, isUsd)}}</td>
      <td class="num deposit-amount">${{fmtAmt(row.deposit, isUsd)}}</td>
      <td class="num withdrawal-amount">${{fmtAmt(row.withdrawal, isUsd)}}</td>
      <td class="num balance-amount">${{fmtAmt(row.end, isUsd)}}</td>
      ${{krwCell}}
    </tr>`;
  }});

  if (summaryRows === '') {{
    summaryRows = '<tr><td colspan="6" class="no-data">계좌 데이터 없음</td></tr>';
  }} else {{
    summaryRows += `<tr class="total-row">
      <td class="center">합 계</td>
      <td class="num">₩${{fmt(sumKrwBegin)}}</td>
      <td class="num">₩${{fmt(sumKrwDep)}}</td>
      <td class="num">₩${{fmt(sumKrwWdr)}}</td>
      <td class="num">-</td>
      <td class="num">₩${{fmt(sumKrwEnd)}}</td>
    </tr>`;
  }}

  const depTableHtml = renderTxTable('dep', d.deposits||[], 'deposit', 'deposit-amount', date);
  const wdrTableHtml = renderTxTable('wdr', d.withdrawals||[], 'withdrawal', 'withdrawal-amount', date);

  document.getElementById('reportPanel').innerHTML = `
    <div class="section-card">
      <div class="section-header">
        <h2>자금 요약</h2>
        <span class="date-badge">${{fmtDate(date)}}</span>
      </div>
      <div class="section-body">
        <div class="stat-row">
          <div class="stat-box dep">
            <div class="label">입금 (원화환산)</div>
            <div class="value">₩${{fmt(totalDep)}}</div>
          </div>
          <div class="stat-box wdr">
            <div class="label">출금 (원화환산)</div>
            <div class="value">₩${{fmt(totalWdr)}}</div>
          </div>
          <div class="stat-box net">
            <div class="label">순유입</div>
            <div class="value">₩${{fmt(netFlow)}}</div>
          </div>
        </div>
      </div>
    </div>

    <div class="section-card">
      <div class="section-header">
        <h2>계좌별 현황</h2>
        ${{rateBadge}}
      </div>
      <div class="section-body" style="padding:0 0 8px 0;">
        <table>
          <thead>
            <tr>
              <th style="width:110px">계 좌</th>
              <th>기초잔액</th>
              <th>입 금</th>
              <th>출 금</th>
              <th>잔 액</th>
              <th style="color:#F5A623">원화잔액</th>
            </tr>
          </thead>
          <tbody>${{summaryRows}}</tbody>
        </table>
      </div>
    </div>

    <div class="section-card">
      <div class="section-header">
        <h2>입출금 내역</h2>
      </div>
      <div class="section-body">
        <div class="tabs">
          <div class="tab active" onclick="switchTab(this,'tab-dep')">입금 내역 (${{d.summary_only ? '집계만' : (d.deposits||[]).length + '건'}})</div>
          <div class="tab" onclick="switchTab(this,'tab-wdr')">출금 내역 (${{d.summary_only ? '집계만' : (d.withdrawals||[]).length + '건'}})</div>
        </div>
        <div id="tab-dep" class="tab-content active">
          <div id="table-dep">${{depTableHtml}}</div>
        </div>
        <div id="tab-wdr" class="tab-content">
          <div id="table-wdr">${{wdrTableHtml}}</div>
        </div>
      </div>
    </div>
  `;
}}

function switchTab(el, tabId) {{
  el.closest('.section-body').querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  el.closest('.section-body').querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  el.classList.add('active');
  document.getElementById(tabId).classList.add('active');
}}

// ── 데이터 불러오기 ──────────────────────────────
function showToast(msg, type) {{
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast ' + type + ' show';
  setTimeout(() => t.className = 'toast ' + type, 3000);
}}

async function refreshData() {{
  const btn   = document.getElementById('refreshBtn');
  const icon  = btn.querySelector('.btn-icon');
  const label = btn.querySelector('.btn-label');

  btn.disabled = true;
  btn.classList.add('loading');
  label.textContent = '업데이트 중...';

  // 1) 로컬 서버 우선 시도 (자동감시 실행 중일 때)
  const LOCAL  = 'http://localhost:8765/update';
  const WEBHOOK = 'https://script.google.com/macros/s/AKfycbyPntu89UWl3G-BaklTRq1Sx_yJAeh1lrD6W8BT-HaOXkrunGDZDkNFFi33lwc00Oad/exec';
  const isLocal = location.hostname === 'localhost';

  let success = false;

  // 로컬 서버 시도
  if (!success) {{
    try {{
      const res = await fetch(LOCAL, {{ method: 'POST', signal: AbortSignal.timeout(5000) }});
      if (res.ok) {{
        success = true;
        if (isLocal) {{
          showToast('✓ 업데이트 완료!', 'success');
          setTimeout(() => location.reload(), 800);
          return;
        }} else {{
          showToast('✓ 업데이트 요청 완료! 2~3분 후 F5 새로고침하세요.', 'success');
        }}
      }}
    }} catch(e) {{ /* 로컬 서버 없음 → 웹훅으로 */ }}
  }}

  // 웹훅(Google Apps Script) 시도
  if (!success) {{
    try {{
      await fetch(WEBHOOK, {{ mode: 'no-cors' }});
      success = true;
      showToast('✓ GitHub 업데이트 요청! 2~3분 후 F5 새로고침하세요.', 'success');
    }} catch(e) {{
      showToast('⚠ 업데이트 요청 실패. 네트워크를 확인해 주세요.', 'error');
    }}
  }}

  btn.disabled = false;
  btn.classList.remove('loading');
  label.textContent = '데이터 불러오기';
}}

// 초기화
renderYearSelector();
renderCalendar();

// ── 비밀번호 인증 ──────────────────────────────────────
const _K = 'cGhhcm1hLWJyb3M=';

function checkPw() {{
  const inp = document.getElementById('pwInput');
  const err = document.getElementById('pwError');
  if (btoa(inp.value) === _K) {{
    sessionStorage.setItem('_pbd', _K);
    inp.classList.remove('error');
    err.textContent = '';
    showDashboard();
  }} else {{
    inp.classList.remove('error');
    void inp.offsetWidth;
    inp.classList.add('error');
    err.textContent = '비밀번호가 올바르지 않습니다.';
    inp.value = '';
    inp.focus();
  }}
}}

function showDashboard() {{
  document.getElementById('auth-overlay').style.display = 'none';
  document.getElementById('main-content').style.display = 'block';
}}

function lockDashboard() {{
  sessionStorage.removeItem('_pbd');
  document.getElementById('auth-overlay').style.display = 'flex';
  document.getElementById('main-content').style.display = 'none';
  document.getElementById('pwInput').value = '';
  document.getElementById('pwInput').focus();
}}

if (sessionStorage.getItem('_pbd') === _K) {{
  showDashboard();
}} else {{
  document.getElementById('pwInput').focus();
}}
</script>

</div><!-- /main-content -->
</body>
</html>"""

with open(out_path, 'w', encoding='utf-8') as f:
    f.write(HTML)

print("완료:", out_path)
print("처리된 날짜 수:", len(dashboard_data))
print("계좌 수:", len(all_accounts))
