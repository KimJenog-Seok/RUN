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
        base_title = make_yesterday_title_kst()
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
                            "startRowIndex": 0,
                            "endRowIndex": 101,
                            "startColumnIndex": 0,
                            "endColumnIndex": 10
                        },
                        "top":    {"style": "SOLID"},
                        "bottom": {"style": "SOLID"},
                        "left":   {"style": "SOLID"},
                        "right":  {"style": "SOLID"},
                        "innerHorizontal": {"style": "SOLID"},
                        "innerVertical":   {"style": "SOLID"},
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
                            "startColumnIndex": 2,
                            "endColumnIndex": 10
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
                            "startIndex": 1,
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
                            "startIndex": 8,
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
        plat_tbl = _agg_two(df_ins, ["회사명"])
        cat_tbl = _agg_two(df_ins, ["분류"])

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

        # --- INS_전일 시트 서식 지정: START ---
        try:
            ins_ws.update("A1", sheet_data)
            print("✅ INS_전일 시트 데이터 업로드 완료")

            # A24:E24에 해당하는 첫 번째 테이블의 헤더 행 찾기
            header_row_idx = -1
            for i, row in enumerate(sheet_data):
                if len(row) > 0 and row[0] == '분류':
                    header_row_idx = i
                    break
            
            # A35:E35에 해당하는 두 번째 테이블의 헤더 행 찾기
            new_items_header_row_idx = -1
            found_first_header = False
            for i, row in enumerate(sheet_data):
                if row and row[0] == '분류':
                    if not found_first_header:
                        found_first_header = True
                    else: # 두 번째 '분류' 헤더(신규 진입 상품)
                        # 신규 진입 상품 표의 헤더는 '방송정보'이므로, '방송정보'를 기준으로 찾기
                        pass
                if row and row[0] == '방송정보':
                    new_items_header_row_idx = i
                    break

            format_requests = []
            # 첫 번째 표(상품분류 집계) 헤더 서식 지정
            if header_row_idx != -1:
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ins_ws.id,
                            "startRowIndex": header_row_idx,
                            "endRowIndex": header_row_idx + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 3
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}, # 밝은 회색 음영
                                "textFormat": {"bold": True} # 굵은 글씨
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat)"
                    }
                })
            
            # 신규 진입 상품 표 헤더 서식 지정
            if new_items_header_row_idx != -1:
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ins_ws.id,
                            "startRowIndex": new_items_header_row_idx,
                            "endRowIndex": new_items_header_row_idx + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 5
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}, # 밝은 회색 음영
                                "textFormat": {"bold": True} # 굵은 글씨
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat)"
                    }
                })
            
            sh.batch_update({"requests": format_requests})
            print("✅ INS_전일 시트 서식 지정 완료")
        except Exception as e:
            print("⚠️ INS_전일 시트 서식 지정 실패:", e)
        # --- INS_전일 시트 서식 지정: END ---

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
