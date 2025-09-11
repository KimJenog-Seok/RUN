#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ECOMM 로그인 전용 스크립트 (v1.0)
- 목표: 로그인 성공 판정까지 완료하고 종료
- 특징:
  * 의미 기반 셀렉터 사용 (해시형/빌드 의존 클래스 지양)
  * SPA 환경 대응: URL 변화 대신 '성공 후에만 보이는 요소'로 판정
  * 동시 접속 팝업(ant-modal 계열) 자동 처리
  * 최대 3회 리트라이
  * 실패 시 artifacts/ 에 스크린샷/HTML 저장 (CI에서 업로드 권장)
환경 변수:
  - ID1: ECOMM 로그인 이메일
  - PW1: ECOMM 로그인 비밀번호
  - (KEY1: 이후 단계에서 구글 인증용 Base64 키가 필요할 수 있으나, 본 버전에서는 미사용)
종속성:
  - selenium
  - webdriver_manager (로컬 실행 편의. CI에서는 사전 설치된 크롬/드라이버 사용 가능)
"""
import os
import re
import sys
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# -------- 설정 값 --------
LOGIN_URL = "https://live.ecomm-data.com/user/sign_in"
POST_LOGIN_PROBES = [
    ("css", "a[href='/ranking']"),
    ("css", "header"),
    ("css", "nav"),
]
RANKING_URL = "https://live.ecomm-data.com/ranking?period=1&cid=&date="
MAX_RETRY = 3
WAIT_SHORT = 5
WAIT_MED = 15
WAIT_LONG = 25

ARTIFACT_DIR = Path("artifacts")
ARTIFACT_DIR.mkdir(exist_ok=True)

def getenv_required(key: str) -> str:
    val = os.getenv(key, "").strip()
    if not val:
        print(f"[FATAL] 환경변수 {key}가 비어있습니다.", file=sys.stderr)
        sys.exit(2)
    return val

ECOMM_ID = getenv_required("ID1")
ECOMM_PW = getenv_required("PW1")

def _is_ci_env() -> bool:
    return os.getenv("GITHUB_ACTIONS", "") == "true"

def make_driver():
    """크롬 드라이버 생성 (CI/로컬 공용)."""
    opts = webdriver.ChromeOptions()
    # CI & 일반 헤드리스 옵션
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1280,2000")
    opts.add_argument("--lang=ko-KR")
    # 크롤링 탐지 완화용 UA
    opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/122.0.0.0 Safari/537.36")
    # 일부 사이트에서 자동화 플래그 차단 회피
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=opts)
    # navigator.webdriver 제거 시도
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
        print(f"[DEBUG] saved {png.name}, {html.name}")
    except Exception as e:
        print(f"[WARN] save_debug failed: {e}")

def find_css(driver, selector: str, wait: int = WAIT_MED):
    return WebDriverWait(driver, wait).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
    )

def click_js(driver, elem):
    driver.execute_script("arguments[0].click();", elem)

def handle_session_popup(driver, max_wait: int = WAIT_SHORT) -> bool:
    """동시 접속 안내창(ant modal) 감지 시 '종료 후 접속' 처리."""
    try:
        popup = WebDriverWait(driver, max_wait).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.ant-modal-content"))
        )
    except TimeoutException:
        return False

    # 버튼 텍스트 기준으로 탐색
    buttons = popup.find_elements(By.TAG_NAME, "button")
    for b in buttons:
        txt = (b.text or "").strip()
        if re.search(r"종료\s*후\s*접속", txt):
            try:
                click_js(driver, b)
                time.sleep(1)
                print("✅ 동시 접속 팝업: '종료 후 접속' 클릭")
                return True
            except Exception:
                pass

    # 대안 경로: 리스트 아이템이 있는 경우 마지막 항목 → 확인 버튼
    lis = popup.find_elements(By.TAG_NAME, "li")
    if lis:
        try:
            click_js(driver, lis[-1])
            time.sleep(0.5)
            for b in buttons:
                if re.search(r"종료\s*후\s*접속", (b.text or "")):
                    click_js(driver, b)
                    time.sleep(1)
                    print("✅ 동시 접속 팝업(대안): '종료 후 접속' 클릭")
                    return True
        except Exception:
            pass
    return False

def wait_logged_in(driver, max_wait: int = WAIT_LONG) -> bool:
    for kind, sel in POST_LOGIN_PROBES:
        try:
            if kind == "css":
                WebDriverWait(driver, max_wait).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, sel))
                )
                return True
        except TimeoutException:
            continue
    return False

def login_once(driver) -> bool:
    driver.get(LOGIN_URL)
    print(f"[INFO] open login: {driver.current_url}")
    try:
        email_input = find_css(driver, "input[name='email']", wait=WAIT_LONG)
        pw_input = find_css(driver, "input[name='password']", wait=WAIT_LONG)
    except TimeoutException:
        print("[ERROR] 로그인 입력창을 찾지 못했습니다.")
        save_debug(driver, "no_login_fields")
        return False

    email_input.clear(); email_input.send_keys(ECOMM_ID)
    time.sleep(0.2)
    pw_input.clear(); pw_input.send_keys(ECOMM_PW)

    # 제출
    try:
        submit_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        click_js(driver, submit_btn)
    except NoSuchElementException:
        pw_input.send_keys(Keys.ENTER)

    # 성공 판정: 사후 요소 대기
    if not wait_logged_in(driver, max_wait=WAIT_LONG):
        print("[ERROR] 로그인 성공 요소가 나타나지 않았습니다.")
        save_debug(driver, "login_fail")
        return False

    # 로그인 직후 팝업 처리(있을 때만)
    handle_session_popup(driver, max_wait=WAIT_SHORT)
    print("[INFO] 로그인 성공(추정)")
    return True

def ensure_login_with_retry(driver, tries: int = MAX_RETRY) -> bool:
    for i in range(1, tries + 1):
        print(f"[TRY] 로그인 시도 {i}/{tries}")
        ok = login_once(driver)
        if ok:
            return True
        # 재시도 준비
        try:
            driver.delete_all_cookies()
        except Exception:
            pass
        time.sleep(2)
    return False

def main():
    driver = make_driver()
    try:
        ok = ensure_login_with_retry(driver, tries=MAX_RETRY)
        if not ok:
            print("[FATAL] 로그인에 3회 실패했습니다. artifacts/ 내 캡처와 HTML을 확인하세요.", file=sys.stderr)
            save_debug(driver, "final_fail")
            sys.exit(1)

        # 랭킹 페이지 진입(팝업 2차 처리 + 요소 로드 확인)
        driver.get(RANKING_URL)
        handle_session_popup(driver, max_wait=WAIT_SHORT)
        try:
            WebDriverWait(driver, WAIT_LONG).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
            print("[INFO] 랭킹 페이지 테이블 로드 확인")
        except TimeoutException:
            print("[WARN] 랭킹 테이블을 찾지 못했습니다. 로그인은 성공했을 수 있습니다.")
            save_debug(driver, "ranking_table_missing")

        print("✅ 최종 상태: 로그인 성공. (다음 단계 작업 가능)")
        sys.exit(0)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
