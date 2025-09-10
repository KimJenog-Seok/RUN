from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os
import pandas as pd
import re
import base64
import gspread
import json
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# === 환경변수에서 자격 증명 읽기 ===
ECOMM_ID = os.environ.get("ID1", "")
ECOMM_PW = os.environ.get("PW1", "")
GSVC_JSON_B64 = os.environ.get("KEY1", "")
SERVICE_ACCOUNT_INFO = {}
if GSVC_JSON_B64:
    try:
        SERVICE_ACCOUNT_INFO = json.loads(base64.b64decode(GSVC_JSON_B64).decode("utf-8"))
    except Exception as e:
        print("[WARN] 서비스계정 Base64 디코딩 실패:", e)

# === 환경변수 검증 ===
if not ECOMM_ID or not ECOMM_PW:
    raise RuntimeError("환경변수 ID1/PW1 가 설정되어야 합니다")

# =========================
# 환경/설정
# =========================

options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

# 👇 추가: 헤드리스/자동화 탐지 회피
options.add_argument("--lang=ko-KR")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-infobars")
options.add_argument("--start-maximized")
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")

driver = webdriver.Chrome(options=options)

# 👇 추가: navigator.webdriver 숨기기 (탐지 우회)
driver.execute_cdp_cmd(
    "Page.addScriptToEvaluateOnNewDocument",
    {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
)

# Google Sheets
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1kravfzRDMhArlt-uqEYjMIn0BVCY4NtRZekswChLTzo/edit?usp=sharing'
WORKSHEET_NAME = '홈쇼핑TOP100'
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"  # ★ 서식/필터/색/폭 변경에 필요
]
creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scope)
gc = gspread.authorize(creds)
sh = gc.open_by_url(SPREADSHEET_URL)
worksheet = sh.worksheet(WORKSHEET_NAME)

# =========================
# 플랫폼 ↔ 홈쇼핑구분 매핑
# =========================
PLATFORM_MAP = {
    "CJ온스타일": "Live",
    "CJ온스타일 플러스": "TC",
    "GS홈쇼핑": "Live",
    "GS홈쇼핑 마이샵": "TC",
    "KT알파쇼핑": "TC",
    "NS홈쇼핑": "Live",
    "NS홈쇼핑 샵플러스": "TC",
    "SK스토아": "TC",
    "공영쇼핑": "Live",
    "롯데원티비": "TC",
    "롯데홈쇼핑": "Live",
    "쇼핑엔티": "TC",
    "신세계쇼핑": "TC",
    "현대홈쇼핑": "Live",
    "현대홈쇼핑 플러스샵": "TC",
    "홈앤쇼핑": "Live",
}
PLATFORMS_BY_LEN = sorted(PLATFORM_MAP.keys(), key=len, reverse=True)  # 긴 이름 우선

def make_yesterday_title_kst():
    KST = timezone(timedelta(hours=9))
    today = datetime.now(KST).date()
    yday = today - timedelta(days=1)
    return f"{yday.month}/{yday.day}"  # 예: "8/22"

def unique_sheet_title(base):
    """동일 이름이 있으면 -1, -2 붙여서 유일한 시트명 반환"""
    title = base
    n = 1
    while True:
        try:
            sh.worksheet(title)
            n += 1
            title = f"{base}-{n}"
        except gspread.exceptions.WorksheetNotFound:
            return title

def split_company_from_broadcast(text):
    """방송정보 끝의 회사명을 찾아 제거하고 (cleaned, company, gubun) 반환"""
    if not text:
        return text, "", ""
    t = text.rstrip()
    for key in PLATFORMS_BY_LEN:
        pattern = r"\s*" + re.escape(key) + r"\s*$"
        if re.search(pattern, t):
            cleaned = re.sub(pattern, "", t).rstrip()
            return cleaned, key, PLATFORM_MAP[key]
    return text, "", ""

# =========================
# 1) 로그인 (헤드리스/지연 대응)
# =========================
driver.get("https://live.ecomm-data.com")

# 로그인 링크 대기 후 클릭 (가시성+클릭가능 대기)
login_link = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.LINK_TEXT, "로그인"))
)
driver.execute_script("arguments[0].click();", login_link)

# 로그인 페이지 진입 대기
WebDriverWait(driver, 20).until(lambda d: "/user/sign_in" in d.current_url)
print("✅ 로그인 페이지 진입 완료:", driver.current_url)

# 폼 요소 대기
email_input = WebDriverWait(driver, 20).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "input[name='email']"))
)
password_input = WebDriverWait(driver, 20).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "input[name='password']"))
)

# ↘️ 시크릿(환경변수) 사용
email_input.clear();    email_input.send_keys(ECOMM_ID)
password_input.clear(); password_input.send_keys(ECOMM_PW)

# 버튼 클릭 (form 내부의 '로그인' 버튼)
form = driver.find_element(By.TAG_NAME, "form")
login_button = form.find_element(By.XPATH, ".//button[contains(text(), '로그인')]")
driver.execute_script("arguments[0].click();", login_button)
print("✅ 로그인 시도!")

# =========================
# 1-1) 로그인 후 페이지 이동 및 세션 정리
# =========================
try:
    # 로그인 성공 후 랭킹 페이지의 테이블이 나타날 때까지 기다림
    # 👇 대기 시간 40초로 변경
    WebDriverWait(driver, 40).until(
        EC.visibility_of_element_located((By.TAG_NAME, "table"))
    )
    print("✅ 로그인 후 랭킹 페이지 진입 완료!")
except TimeoutException:
    print("⚠️ 랭킹 페이지 진입 실패. 세션 팝업 또는 기타 오류 확인 중...")
    try:
        # 랭킹 페이지로 이동하지 않았다면, 세션 팝업이 떴을 가능성을 확인
        # 👇 대기 시간 20초로 변경
        session_items = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.jsx-6ce14127fb5f1929 > li"))
        )
        if session_items:
            print(f"[INFO] 세션 초과: {len(session_items)}개 → 맨 아래 선택 후 '종료 후 접속'")
            driver.execute_script("arguments[0].click();", session_items[-1])
            close_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='종료 후 접속']"))
            )
            driver.execute_script("arguments[0].click();", close_btn)
            print("✅ '종료 후 접속' 버튼 클릭 완료")
            # 세션 처리 후 다시 랭킹 페이지 진입을 기다림
            # 👇 대기 시간 30초로 변경
            WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located((By.TAG_NAME, "table"))
            )
            print("✅ 세션 처리 후 랭킹 페이지 재진입 성공!")
    except Exception as e:
        print(f"⚠️ 세션 처리 실패 또는 다른 오류 발생: {e}")
        # 세션 처리 실패 시에는 오류를 다시 발생시켜서 원인 파악을 돕습니다.
        raise TimeoutException("로그인 후 랭킹 페이지나 세션 팝업을 찾을 수 없습니다. 환경변수를 확인하거나 사이트 상태를 점검해주세요.")

print("✅ 로그인 절차 완료!")


# =========================
# 2) 랭킹 페이지 크롤링
# =========================
ranking_url = "https://live.ecomm-data.com/ranking?period=1&cid=&date="
driver.get(ranking_url)
time.sleep(3)

table = driver.find_element(By.TAG_NAME, 'table')
tbody = table.find_element(By.TAG_NAME, 'tbody')
rows = tbody.find_elements(By.TAG_NAME, 'tr')

data = []
for row in rows:
    cols = row.find_elements(By.TAG_NAME, 'td')
    if len(cols) >= 8:
        item = {
            "랭킹": cols[0].text.strip(),
            "방송정보": cols[1].text.strip(),
            "분류": cols[2].text.strip(),
            "방송시간": cols[3].text.strip(),
            "시청률": cols[4].text.strip(),
            "판매량": cols[5].text.strip(),
            "매출액": cols[6].text.strip(),
            "상품수": cols[7].text.strip(),
        }
        data.append(item)

columns = ["랭킹", "방송정보", "분류", "방송시간", "시청률", "판매량", "매출액", "상품수"]
df = pd.DataFrame(data, columns=columns)
print(df.head())
print(f"총 {len(df)}개 상품 정보 추출 완료")

# =========================
# 3) '홈쇼핑TOP100' 시트 갱신
# =========================
data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
worksheet.clear()
worksheet.update(values=data_to_upload, range_name='A1')
print("✅ 구글시트 업로드 완료!")

# =========================
# 4) 어제 날짜 새 시트 생성 & 값 복사
# =========================
base_title = make_yesterday_title_kst()           # 예: "8/22"
target_title = unique_sheet_title(base_title)     # 중복 시 -1, -2…

source_values = worksheet.get_all_values() or [[""]]
rows_cnt = max(2, len(source_values))
cols_cnt = max(2, max(len(r) for r in source_values))
new_ws = sh.add_worksheet(title=target_title, rows=rows_cnt, cols=cols_cnt)
new_ws.update('A1', source_values)
print(f"✅ 새 시트 생성 및 값 붙여넣기 완료 → 시트명: {target_title}")

# =========================
# 5) 방송정보에서 회사명 제거 + 회사명/구분 열 추가
# =========================
values = new_ws.get_all_values() or [[""]]
header = values[0]
data_rows = values[1:]

final_header = header + ["회사명", "홈쇼핑구분"]
final_rows = []

for r in data_rows:
    padded = r + [""] * (len(header) - len(r))
    broadcast = padded[1].strip() if len(padded) > 1 else ""
    cleaned, company, gubun = split_company_from_broadcast(broadcast)
    if len(padded) > 1:
        padded[1] = cleaned
    final_rows.append(padded + [company, gubun])

final_data = [final_header] + final_rows
new_ws.clear()
new_ws.update('A1', final_data)
print("✅ 방송정보 말미 회사명 제거 + 회사명/홈쇼핑구분 열 추가 완료")

# =========================
# 6) 인사이트(단일 시트: INS_전일)
# =========================
def _to_int_kor(s):
    if s is None:
        return 0
    t = str(s).strip()
    if t == "" or t == "-":
        return 0
    t = t.replace(",", "").replace(" ", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", t):
        return int(float(t))
    unit_map = {"억": 100_000_000, "만": 10_000}
    m = re.fullmatch(r"(-?\d+(?:\.\d+)?)(억|만)", t)
    if m:
        return int(float(m.group(1)) * unit_map[m.group(2)])
    total = 0
    rest = t
    if "억" in rest:
        parts = rest.split("억")
        try: total += int(float(parts[0]) * unit_map["억"])
        except: pass
        rest = parts[1] if len(parts) > 1 else ""
    if "만" in rest:
        parts = rest.split("만")
        try: total += int(float(parts[0]) * unit_map["만"])
        except: pass
        rest = parts[1] if len(parts) > 1 else ""
    if re.fullmatch(r"-?\d+", rest):
        total += int(rest)
    if total == 0:
        nums = re.findall(r"-?\d+", t)
        return int(nums[0]) if nums else 0
    return total

def format_sales(v):
    try: v = int(v)
    except: return str(v)
    return f"{v/100_000_000:.2f}억"

def format_num(v):
    try: v = int(v)
    except: return str(v)
    return f"{v:,}"

# 6-1) 날짜 시트 new_ws 기준 DF
values = new_ws.get_all_values() or [[""]]
if not values or len(values) < 2:
    raise Exception("INS_전일 생성 실패: 데이터 행이 없습니다.")
header = values[0]; body = values[1:]
df_ins = pd.DataFrame(body, columns=header)

for col in ["판매량","매출액","홈쇼핑구분","회사명","분류"]:
    if col not in df_ins.columns: df_ins[col] = ""
df_ins["판매량_int"] = df_ins["판매량"].apply(_to_int_kor)
df_ins["매출액_int"] = df_ins["매출액"].apply(_to_int_kor)

# 6-2) 집계 → 포맷
def _agg_two(group_cols):
    g = (df_ins.groupby(group_cols, dropna=False)
                .agg(매출합=("매출액_int","sum"),
                     판매량합=("판매량_int","sum"))
                .reset_index()
                .sort_values("매출합", ascending=False))
    return g

gubun_tbl = _agg_two(["홈쇼핑구분"])
plat_tbl  = _agg_two(["회사명"])
cat_tbl   = _agg_two(["분류"])

def _format_df(df):
    d = df.copy()
    d["매출합"] = d["매출합"].apply(format_sales)
    d["판매량합"] = d["판매량합"].apply(format_num)
    return [d.columns.tolist()] + d.astype(str).values.tolist()

gubun_table = _format_df(gubun_tbl)
plat_table  = _format_df(plat_tbl)
cat_table   = _format_df(cat_tbl)

# 6-3) 기본 섹션(A/B/C)
sheet_data = []
sheet_data.append(["[LIVE/TC 집계]"])
sheet_data += gubun_table
sheet_data.append([""])

sheet_data.append(["[플랫폼(회사명) 집계]"])
sheet_data += plat_table
sheet_data.append([""])

sheet_data.append(["[상품분류(분류) 집계]"])
sheet_data += cat_table

# 6-4) 신규 진입 상품(최신 날짜 전체 비교)
def _norm_text(s: str) -> str:
    if s is None: return ""
    t = str(s).replace("\n"," ").replace("\r"," ").replace("\t"," ")
    t = re.sub(r"[·/【】\[\]\(\)]", " ", t)
    return re.sub(r"\s+"," ", t).strip()

def _make_key(df):
    for c in ["방송정보","회사명"]:
        if c not in df.columns: df[c] = ""
    a = df["방송정보"].apply(_norm_text).astype(str)
    b = df["회사명"].apply(_norm_text).astype(str)
    return a + "||" + b

all_ws_objs = sh.worksheets()
date_ws_objs = [w for w in all_ws_objs if re.match(r"^\d{1,2}/\d{1,2}(-\d+)?$", w.title)]

if date_ws_objs:
    def _parse_md_suffix(title: str):
        base = title.split("-")[0]
        m, d = map(int, base.split("/"))
        suf = 1
        mobj = re.search(r"-(\d+)$", title)
        if mobj: suf = int(mobj.group(1))
        return (m, d, suf)

    latest_ws_obj = max(date_ws_objs, key=lambda w: _parse_md_suffix(w.title))
    latest_title = latest_ws_obj.title
    latest_m, latest_d, _ = _parse_md_suffix(latest_title)

    latest_vals = latest_ws_obj.get_all_values() or [[""]]
    latest_header = latest_vals[0] if latest_vals else []
    latest_df = pd.DataFrame(latest_vals[1:], columns=latest_header) if len(latest_vals) >= 2 else pd.DataFrame(columns=["방송정보","회사명","분류","판매량","매출액"])
    for c in ["방송정보","회사명","분류","판매량","매출액"]:
        if c not in latest_df.columns: latest_df[c] = ""
    latest_df = latest_df.fillna("")
    latest_df["__KEY__"] = _make_key(latest_df)
    latest_keys = set(latest_df["__KEY__"])

    hist_keys = set()
    for w in date_ws_objs:
        m, d, _s = _parse_md_suffix(w.title)
        if (m == latest_m and d == latest_d):
            continue
        prev_vals = w.get_all_values() or [[""]]
        if not prev_vals or len(prev_vals) < 2:
            continue
        prev_df = pd.DataFrame(prev_vals[1:], columns=prev_vals[0])
        for c in ["방송정보","회사명"]:
            if c not in prev_df.columns: prev_df[c] = ""
        prev_df = prev_df.fillna("")
        hist_keys |= set(_make_key(prev_df))

    new_keys = latest_keys - hist_keys
    new_items = latest_df[latest_df["__KEY__"].isin(new_keys)].copy()

    new_items["판매량"] = new_items["판매량"].apply(_to_int_kor).apply(format_num)
    new_items["매출액"] = new_items["매출액"].apply(_to_int_kor).apply(format_sales)

    sheet_data.append([""])
    sheet_data.append([f"[신규 진입 상품] ({latest_title} 기준)"])
    new_table = [["방송정보","회사명","분류","판매량","매출액"]]
    if len(new_items) == 0:
        new_table.append(["(신규 진입 없음)", "", "", "", ""])
    else:
        tmp = new_items.copy()
        tmp["__매출액_int"] = tmp["매출액"].apply(_to_int_kor)
        tmp = tmp.sort_values("__매출액_int", ascending=False)
        new_table += tmp[["방송정보","회사명","분류","판매량","매출액"]].astype(str).values.tolist()
    sheet_data += new_table
else:
    sheet_data.append([""])
    sheet_data.append(["[신규 진입 상품] (날짜 시트 없음)"])
    sheet_data += [["방송정보","회사명","분류","판매량","매출액"],
                   ["(비교 불가)", "", "", "", ""]]

# 6-5) INS_전일 upsert
TARGET_TITLE = "INS_전일"
try:
    ws = sh.worksheet(TARGET_TITLE)
    ws.clear()
except gspread.exceptions.WorksheetNotFound:
    rows_cnt = max(2, len(sheet_data))
    cols_cnt = max(2, max(len(r) for r in sheet_data))
    ws = sh.add_worksheet(title=TARGET_TITLE, rows=rows_cnt, cols=cols_cnt)
ws.update("A1", sheet_data)

# =========================
# (마무리) 시트 순서 재배치 + 탭 색상 세팅
# =========================
try:
    all_ws_now = sh.worksheets()
    new_order = []
    new_order.append(ws)
    if new_ws.id != ws.id:
        new_order.append(new_ws)
    for w in all_ws_now:
        if w.id not in (ws.id, new_ws.id):
            new_order.append(w)
    sh.reorder_worksheets(new_order)

    requests = []
    for w in new_order:
        requests.append({
            "updateSheetProperties": {
                "properties": {"sheetId": w.id, "tabColor": None},
                "fields": "tabColor"
            }
        })
    red = {"red": 1.0, "green": 0.0, "blue": 0.0}
    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": ws.id, "tabColor": red},
            "fields": "tabColor"
        }
    })
    if new_ws.id != ws.id:
        requests.append({
            "updateSheetProperties": {
                "properties": {"sheetId": new_ws.id, "tabColor": red},
                "fields": "tabColor"
            }
        })
    sh.batch_update({"requests": requests})
    print("✅ 시트 순서/색상: INS_전일=1번째, 어제시트=2번째, 두 탭=빨강 적용 완료")

except Exception as e:
    print("⚠️ 시트 순서/색상 처리 중 오류:", e)
