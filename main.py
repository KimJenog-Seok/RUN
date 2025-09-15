#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ====== 표준/외부 모듈 ======
import os, sys, time, re, json, base64
from pathlib import Path
from datetime import datetime, timedelta, timezone

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ====== Selenium ======
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ------------------------------------------------------------
# 환경 설정
# ------------------------------------------------------------
WAIT = 5
ARTIFACT_DIR = Path("artifacts")
ARTIFACT_DIR.mkdir(exist_ok=True)

# 로그인 계정 (요청에 따라 하드코딩 유지)
ECOMM_ID = "smt@trncompany.co.kr"
ECOMM_PW = "sales4580!!"

RANKING_URL = "https://live.ecomm-data.com/ranking?period=1&cid=&date="

# 구글 시트 설정
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1kravfzRDMhArlt-uqEYjMIn0BVCY4NtRZekswChLTzo/edit?usp=sharing"
WORKSHEET_NAME = "홈쇼핑TOP100"

# ------------------------------------------------------------
# 유틸
# ------------------------------------------------------------
def make_driver():
    """GitHub Actions/서버/로컬 공용 크롬 드라이버 (Headless)."""
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--lang=ko-KR")
    opts.add_argument("user-agent=Mozilla/5.0 Chrome/122.0.0.0 Safari/537.36")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=opts)
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
        )
    except Exception:
        pass
    driver.set_page_load_timeout(60)
    return driver

def save_debug(driver, tag: str):
    ts = int(time.time())
    png = ARTIFACT_DIR / f"{ts}_{tag}.png"
    html = ARTIFACT_DIR / f"{ts}_{tag}.html"
    try:
        driver.save_screenshot(str(png))
        with open(html, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"[DEBUG] 저장: {png.name}, {html.name}")
    except Exception as e:
        print(f"[WARN] 디버그 저장 실패: {e}")

# ------------------------------------------------------------
# 로그인 + 세션 초과 팝업 처리
# ------------------------------------------------------------
def login_and_handle_session(driver):
    driver.get("https://live.ecomm-data.com")
    print("[STEP] 메인 페이지 진입 완료")

    login_link = WebDriverWait(driver, WAIT).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "로그인"))
    )
    driver.execute_script("arguments[0].click();", login_link)
    print("[STEP] 로그인 링크 클릭 완료")

    t0 = time.time()
    while "/user/sign_in" not in driver.current_url:
        if time.time() - t0 > WAIT:
            raise Exception("로그인 페이지 진입 실패 (타임아웃)")
        time.sleep(0.5)
    print("✅ 로그인 페이지 진입 완료:", driver.current_url)

    time.sleep(1)
    email_input = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='email']") if e.is_displayed()][0]
    pw_input    = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='password']") if e.is_displayed()][0]
    email_input.clear(); email_input.send_keys(ECOMM_ID)
    pw_input.clear(); pw_input.send_keys(ECOMM_PW)
    time.sleep(0.5)

    form = driver.find_element(By.TAG_NAME, "form")
    login_button = form.find_element(By.XPATH, ".//button[contains(text(), '로그인')]")
    driver.execute_script("arguments[0].click();", login_button)
    print("✅ 로그인 시도!")

    # 세션 초과 팝업 처리
    time.sleep(2)
    try:
        session_items = [li for li in driver.find_elements(By.CSS_SELECTOR, "ul > li") if li.is_displayed()]
        if session_items:
            print(f"[INFO] 세션 초과: {len(session_items)}개 → 맨 아래 세션 선택 후 '종료 후 접속'")
            session_items[-1].click()
            time.sleep(1)
            close_btn = driver.find_element(By.XPATH, "//button[text()='종료 후 접속']")
            if close_btn.is_enabled():
                driver.execute_script("arguments[0].click();", close_btn)
                print("✅ '종료 후 접속' 버튼 클릭 완료")
                time.sleep(2)
        else:
            print("[INFO] 세션 초과 안내창 없음")
    except Exception as e:
        print("[WARN] 세션 처리 중 예외(무시):", e)

    # 성공 판정
    time.sleep(2)
    curr = driver.current_url
    email_inputs = driver.find_elements(By.CSS_SELECTOR, "input[name='email']")
    if "/sign_in" in curr and any(e.is_displayed() for e in email_inputs):
        print("❌ 로그인 실패 (폼 그대로 존재함)")
        save_debug(driver, "login_fail")
        raise RuntimeError("로그인 실패")
    print("✅ 로그인 성공 판정! 현재 URL:", curr)
    save_debug(driver, "login_success")

# ------------------------------------------------------------
# 랭킹 페이지 크롤링
# ------------------------------------------------------------
def crawl_ranking(driver):
    driver.get(RANKING_URL)
    time.sleep(3)
    table = driver.find_element(By.TAG_NAME, "table")
    tbody = table.find_element(By.TAG_NAME, "tbody")
    rows = tbody.find_elements(By.TAG_NAME, "tr")

    data = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) >= 8:
            item = {
                "랭킹":   cols[0].text.strip(),
                "방송정보": cols[1].text.strip(),
                "분류":   cols[2].text.strip(),
                "방송시간": cols[3].text.strip(),
                "시청률":  cols[4].text.strip(),
                "판매량":  cols[5].text.strip(),
                "매출액":  cols[6].text.strip(),
                "상품수":  cols[7].text.strip(),
            }
            data.append(item)

    columns = ["랭킹","방송정보","분류","방송시간","시청률","판매량","매출액","상품수"]
    df = pd.DataFrame(data, columns=columns)
    print(df.head())
    print(f"총 {len(df)}개 상품 정보 추출 완료")
    return df

# ------------------------------------------------------------
# Google Sheets 인증 (KEY1: Base64 JSON)
# ------------------------------------------------------------
def gs_client_from_env():
    GSVC_JSON_B64 = os.environ.get("KEY1", "")
    if not GSVC_JSON_B64:
        raise RuntimeError("환경변수 KEY1이 비어있습니다(Base64 인코딩된 서비스계정 JSON 필요).")
    try:
        svc_info = json.loads(base64.b64decode(GSVC_JSON_B64).decode("utf-8"))
    except Exception as e:
        print("[WARN] 서비스계정 Base64 디코딩 실패:", e)
        raise

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds = Credentials.from_service_account_info(svc_info, scopes=scope)
    return gspread.authorize(creds)

# ------------------------------------------------------------
# 플랫폼 매핑 및 유틸
# ------------------------------------------------------------
PLATFORM_MAP = {
    "CJ온스타일":"Live","CJ온스타일 플러스":"TC","GS홈쇼핑":"Live","GS홈쇼핑 마이샵":"TC",
    "KT알파쇼핑":"TC","NS홈쇼핑":"Live","NS홈쇼핑 샵플러스":"TC","SK스토아":"TC",
    "공영쇼핑":"Live","롯데원티비":"TC","롯데홈쇼핑":"Live","쇼핑엔티":"TC",
    "신세계쇼핑":"TC","현대홈쇼핑":"Live","현대홈쇼핑 플러스샵":"TC","홈앤쇼핑":"Live",
}
PLATFORMS_BY_LEN = sorted(PLATFORM_MAP.keys(), key=len, reverse=True)

def make_yesterday_title_kst():
    KST = timezone(timedelta(hours=9))
    today = datetime.now(KST).date()
    yday = today - timedelta(days=1)
    return f"{yday.month}/{yday.day}"  # 예: "9/10"

def unique_sheet_title(sh, base):
    title = base; n = 1
    while True:
        try:
            sh.worksheet(title)
            n += 1; title = f"{base}-{n}"
        except gspread.exceptions.WorksheetNotFound:
            return title

def split_company_from_broadcast(text):
    if not text:
        return text, "", ""
    t = text.rstrip()
    for key in PLATFORMS_BY_LEN:
        pattern = r"\s*" + re.escape(key) + r"\s*$"
        if re.search(pattern, t):
            cleaned = re.sub(pattern, "", t).rstrip()
            return cleaned, key, PLATFORM_MAP[key]
    return text, "", ""

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
    total = 0; rest = t
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

def _agg_two(df, group_cols):
    g = (df.groupby(group_cols, dropna=False)
            .agg(매출합=("매출액_int","sum"),
                 판매량합=("판매량_int","sum"))
            .reset_index()
            .sort_values("매출합", ascending=False))
    return g

def _format_df_table(df):
    d = df.copy()
    d["매출합"] = d["매출합"].apply(format_sales)
    d["판매량합"] = d["판매량합"].apply(format_num)
    return [d.columns.tolist()] + d.astype(str).values.tolist()

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

# ------------------------------------------------------------
# 메인
# ------------------------------------------------------------
def main():
    driver = make_driver()
    sh = None
    worksheet = None
    new_ws = None
    try:
        # 1) 로그인 + 세션 팝업 처리
        login_and_handle_session(driver)

        # 2) 랭킹 페이지 크롤링
        df = crawl_ranking(driver)

        # 3) 구글 시트 인증
        gc = gs_client_from_env()
        sh = gc.open_by_url(SPREADSHEET_URL)
        print("[GS] 스프레드시트 열기 OK")

        # 4) '홈쇼핑TOP100' 시트 확보(없으면 생성)
        try:
            worksheet = sh.worksheet(WORKSHEET_NAME)
            print("[GS] 기존 워크시트 찾음:", WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows=2, cols=8)
            print("[GS] 워크시트 생성:", WORKSHEET_NAME)

        # 5) 메인 시트 업로드
        data_to_upload = [df.columns.tolist()] + df.astype(str).values.tolist()
        worksheet.clear()
        worksheet.update(values=data_to_upload, range_name="A1")
        print(f"✅ 구글시트 업로드 완료 (행수: {len(data_to_upload)})")

        # 6) 어제 날짜 새 시트 생성 & 값 복사 (반드시 생성되도록 가드)
        base_title = make_yesterday_title_kst()     # 예: "9/10"
        target_title = unique_sheet_title(sh, base_title)
        source_values = worksheet.get_all_values() or [[""]]
        rows_cnt = max(2, len(source_values))
        cols_cnt = max(2, max(len(r) for r in source_values))
        new_ws = sh.add_worksheet(title=target_title, rows=rows_cnt, cols=cols_cnt)
        new_ws.update("A1", source_values)
        print(f"✅ 어제 날짜 시트 생성/복사 완료 → {target_title}")

        # 7) 방송정보 말미 회사명 제거 + 회사명/구분 열 추가
        values = new_ws.get_all_values() or [[""]]
        header = values[0] if values else []
        data_rows = values[1:] if len(values) >= 2 else []
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
        new_ws.update("A1", final_data)
        print("✅ 방송정보 말미 회사명 제거 + 회사명/홈쇼핑구분 열 추가 완료")

                # --- 어제 시트 표 서식 지정 (A1:J101) ---
        try:
            reqs = [
                # 1) A1:J101 모든 방향 테두리
                {
                    "updateBorders": {
                        "range": {
                            "sheetId": new_ws.id,
                            "startRowIndex": 0,   # A1 시작
                            "endRowIndex": 101,   # 101행까지
                            "startColumnIndex": 0,# A열 시작
                            "endColumnIndex": 10  # J열(0-based 10) 직전 인덱스
                        },
                        "top":    {"style": "SOLID"},
                        "bottom": {"style": "SOLID"},
                        "left":   {"style": "SOLID"},
                        "right":  {"style": "SOLID"},
                        "innerHorizontal": {"style": "SOLID"},
                        "innerVertical":   {"style": "SOLID"},
                    }
                },
                # 2) 헤더 A1:J1 가운데 정렬 + 회색 배경
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": new_ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 10
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "CENTER",
                                "backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}
                            }
                        },
                        "fields": "userEnteredFormat(horizontalAlignment,backgroundColor)"
                    }
                },
                # 3) A2:A101 가운데 정렬
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": new_ws.id,
                            "startRowIndex": 1,
                            "endRowIndex": 101,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1
                        },
                        "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                        "fields": "userEnteredFormat.horizontalAlignment"
                    }
                },
                # 3-2) C1:J101 가운데 정렬 (헤더 포함 C~J 전 범위)
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": new_ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": 101,
                            "startColumnIndex": 2,  # C열
                            "endColumnIndex": 10    # J열(0-based 10)
                        },
                        "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                        "fields": "userEnteredFormat.horizontalAlignment"
                    }
                },
                # 4) B2:B101 왼쪽 정렬
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": new_ws.id,
                            "startRowIndex": 1,
                            "endRowIndex": 101,
                            "startColumnIndex": 1,
                            "endColumnIndex": 2
                        },
                        "cell": {"userEnteredFormat": {"horizontalAlignment": "LEFT"}},
                        "fields": "userEnteredFormat.horizontalAlignment"
                    }
                },
                # 5) B열 전체 열폭 650px
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": new_ws.id,
                            "dimension": "COLUMNS",
                            "startIndex": 1,  # B열
                            "endIndex": 2
                        },
                        "properties": { "pixelSize": 650 },
                        "fields": "pixelSize"
                    }
                },
                # 6) I열 전체 열폭 120px
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": new_ws.id,
                            "dimension": "COLUMNS",
                            "startIndex": 8,  # I열
                            "endIndex": 9
                        },
                        "properties": { "pixelSize": 120 },
                        "fields": "pixelSize"
                    }
                }
            ]
            sh.batch_update({"requests": reqs})
            print("✅ 어제 시트 서식 지정 완료 (A1:J101 + B열 650px)")
        except Exception as e:
            print("⚠️ 어제 시트 서식 지정 실패:", e)



        # 8) INS_전일 생성/갱신
        values = new_ws.get_all_values() or [[""]]
        if not values or len(values) < 2:
            raise Exception("INS_전일 생성 실패: 데이터 행이 없습니다.")
        header = values[0]; body = values[1:]
        df_ins = pd.DataFrame(body, columns=header)

        for col in ["판매량","매출액","홈쇼핑구분","회사명","분류"]:
            if col not in df_ins.columns: df_ins[col] = ""
        df_ins["판매량_int"] = df_ins["판매량"].apply(_to_int_kor)
        df_ins["매출액_int"] = df_ins["매출액"].apply(_to_int_kor)

        gubun_tbl = _agg_two(df_ins, ["홈쇼핑구분"])
        plat_tbl  = _agg_two(df_ins, ["회사명"])
        cat_tbl   = _agg_two(df_ins, ["분류"])

        sheet_data = []
        sheet_data.append(["[LIVE/TC 집계]"]); sheet_data += _format_df_table(gubun_tbl); sheet_data.append([""])
        sheet_data.append(["[플랫폼(회사명) 집계]"]); sheet_data += _format_df_table(plat_tbl); sheet_data.append([""])
        sheet_data.append(["[상품분류(분류) 집계]"]); sheet_data += _format_df_table(cat_tbl)

        # 신규 진입 상품 (최신 날짜 vs 과거 전체)
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
                if w.id == latest_ws_obj.id:
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

        # INS_전일 upsert
        TARGET_TITLE = "INS_전일"
        try:
            ins_ws = sh.worksheet(TARGET_TITLE)
            ins_ws.clear()
            print("[GS] INS_전일 기존 워크시트 찾음 → 초기화")
        except gspread.exceptions.WorksheetNotFound:
            rows_cnt = max(2, len(sheet_data))
            cols_cnt = max(2, max(len(r) for r in sheet_data))
            ins_ws = sh.add_worksheet(title=TARGET_TITLE, rows=rows_cnt, cols=cols_cnt)
            print("[GS] INS_전일 워크시트 생성")

        ins_ws.update("A1", sheet_data)
        print("✅ INS_전일 생성/갱신 완료")

        # 9) 탭 순서 재배치: INS_전일 1번째, 어제시트 2번째
        try:
            all_ws_now = sh.worksheets()
            new_order = [ins_ws]
            if new_ws.id != ins_ws.id:
                new_order.append(new_ws)
            for w in all_ws_now:
                if w.id not in (ins_ws.id, new_ws.id):
                    new_order.append(w)
            sh.reorder_worksheets(new_order)
            print("✅ 시트 순서 재배치 완료: INS_전일=1번째, 어제시트=2번째")
        except Exception as e:
            print("⚠️ 시트 순서 재배치 오류:", e)

        print("🎉 전체 파이프라인 완료")

    except Exception as e:
        import traceback
        print("❌ 전체 자동화 과정 중 에러 발생:", e)
        print(traceback.format_exc())
        raise
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
