from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import os
import pandas as pd
import re
import base64
import gspread
import json
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone

# === 환경변수에서 자격 증명 읽기 ===
ECOMM_ID = os.environ.get("ECOMM_ID", "")
ECOMM_PW = os.environ.get("ECOMM_PW", "")
GSVC_JSON_B64 = os.environ.get("GSVC_JSON_B64", "")
SERVICE_ACCOUNT_INFO = {}
if GSVC_JSON_B64:
    try:
        SERVICE_ACCOUNT_INFO = json.loads(base64.b64decode(GSVC_JSON_B64).decode("utf-8"))
    except Exception as e:
        print("[WARN] 서비스계정 Base64 디코딩 실패:", e)



# =========================
# 환경/설정
# =========================

options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

# 👇 추가: 헤드리스/자동화 탐지 우회
options.add_argument("--lang=ko-KR")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-infobars")
options.add_argument("--start-maximized")
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")

driver = webdriver.Chrome(options=options)

# 👇 추가: navigator.webdriver 숨기기
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


# 환경변수 검증
if not ECOMM_ID or not ECOMM_PW:
    raise RuntimeError("환경변수 ECOMM_ID/ECOMM_PW 가 설정되어야 합니다")

# =========================
# 1) 로그인 (헤드리스/지연 대응)
# =========================
driver.get("https://live.ecomm-data.com")

# 로그인 링크 대기 후 클릭 (가시성+클릭가능 대기)
login_link = WebDriverWait(driver, 15).until(
    EC.element_to_be_clickable((By.LINK_TEXT, "로그인"))
)
driver.execute_script("arguments[0].click();", login_link)

# 로그인 페이지 진입 대기
WebDriverWait(driver, 15).until(lambda d: "/user/sign_in" in d.current_url)
print("✅ 로그인 페이지 진입 완료:", driver.current_url)

# 폼 요소 대기
email_input = WebDriverWait(driver, 15).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "input[name='email']"))
)
password_input = WebDriverWait(driver, 15).until(
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

# URL이 /user/sign_in 에서 벗어날 때까지 대기
WebDriverWait(driver, 20).until(lambda d: "/user/sign_in" not in d.current_url)

# =========================
# 1-1) 동시 접속 세션 정리(맨 아래 선택 → '종료 후 접속')
# =========================
try:
    # 세션 리스트가 뜰 경우를 최대 8초 대기
    session_items = WebDriverWait(driver, 8).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.jsx-6ce14127fb5f1929 > li"))
    )
    if session_items:
        print(f"[INFO] 세션 초과: {len(session_items)}개 → 맨 아래 선택 후 '종료 후 접속'")
        driver.execute_script("arguments[0].click();", session_items[-1])
        close_btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='종료 후 접속']"))
        )
        driver.execute_script("arguments[0].click();", close_btn)
        # 세션 처리 후 홈으로 복귀
        WebDriverWait(driver, 10).until(lambda d: "/user/sign_in" not in d.current_url)
        time.sleep(1)
    else:
        print("[INFO] 세션 초과 안내창 없음")
except Exception:
    # 안내창 자체가 없거나 셀렉터 변경 시에도 흐름 계속
    print("[INFO] 세션 초과 안내창 없음(또는 스킵)")

print("✅ 로그인 절차 완료!")


    # =========================
    # 3) 랭킹 페이지 크롤링
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
# 4) '홈쇼핑TOP100' 시트 갱신
# =========================
data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
worksheet.clear()
worksheet.update(values=data_to_upload, range_name='A1')
print("✅ 구글시트 업로드 완료!")



