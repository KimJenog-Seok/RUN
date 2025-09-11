#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ECOMM 로그인 + 세션 초과 팝업 처리
- 로그인 성공 + 세션 초과 안내 팝업이 있을 경우 자동으로 기존 세션 종료 처리
"""

import sys
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

ARTIFACT_DIR = Path("artifacts")
ARTIFACT_DIR.mkdir(exist_ok=True)

WAIT = 5

ECOMM_ID = "smt@trncompany.co.kr"
ECOMM_PW = "sales4580!!"

def make_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1280,2000")
    opts.add_argument("--lang=ko-KR")
    opts.add_argument("user-agent=Mozilla/5.0 Chrome/122.0.0.0 Safari/537.36")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=opts)
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
    )
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
        print(f"[DEBUG] 저장된 {png.name}, {html.name}")
    except Exception as e:
        print(f"[WARN] 디버그 저장 실패: {e}")

def handle_session_popup(driver):
    """동시 접속 초과 안내창이 있으면 → 기존 세션 종료 후 접속."""
    time.sleep(2)
    try:
        session_items = driver.find_elements(By.CSS_SELECTOR, "ul > li")
        session_items = [li for li in session_items if li.is_displayed()]
        if session_items:
            print(f"[INFO] 세션 초과: {len(session_items)}개 → 맨 아래 세션 선택 후 '종료 후 접속'")
            session_items[-1].click()
            time.sleep(1)
            close_btns = driver.find_elements(By.XPATH, "//button[text()='종료 후 접속']")
            for btn in close_btns:
                if btn.is_enabled():
                    driver.execute_script("arguments[0].click();", btn)
                    print("✅ '종료 후 접속' 버튼 클릭 완료")
                    time.sleep(2)
                    return True
        else:
            print("[INFO] 세션 초과 안내창 없음")
    except Exception as e:
        print("[WARN] 세션 처리 중 예외(무시):", e)
    return False

def main():
    driver = make_driver()
    try:
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
        pw_input = [e for e in driver.find_elements(By.CSS_SELECTOR, "input[name='password']") if e.is_displayed()][0]
        email_input.clear(); email_input.send_keys(ECOMM_ID)
        pw_input.clear(); pw_input.send_keys(ECOMM_PW)
        time.sleep(0.5)

        form = driver.find_element(By.TAG_NAME, "form")
        login_button = form.find_element(By.XPATH, ".//button[contains(text(), '로그인')]")
        driver.execute_script("arguments[0].click();", login_button)
        print("✅ 로그인 시도!")

        time.sleep(2)
        handle_session_popup(driver)

        current_url = driver.current_url
        email_inputs = driver.find_elements(By.CSS_SELECTOR, "input[name='email']")
        if "/sign_in" in current_url and any(e.is_displayed() for e in email_inputs):
            print("❌ 로그인 실패 (폼 그대로 존재함)")
            save_debug(driver, "login_fail")
            sys.exit(1)
        else:
            print("✅ 로그인 성공 판정! 현재 URL:", current_url)
            save_debug(driver, "login_success")
            sys.exit(0)

    except Exception as e:
        print(f"[FATAL] 예외 발생: {e}")
        save_debug(driver, "fatal_error")
        sys.exit(1)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
