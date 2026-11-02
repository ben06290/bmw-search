import streamlit as st
import pandas as pd
import re
import io
import urllib.request

st.set_page_config(page_title="BMW 預約車輛快速查詢", layout="centered", page_icon="🚗")

st.title("🚗 BMW 預約車輛快速查詢")

def safe_str(val):
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s

def clean_phone(raw_val):
    s = safe_str(raw_val)
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    if digits.startswith("0"):
        return digits
    if digits.startswith("9") and len(digits) == 9:
        return "0" + digits
    return digits

def clean_plate(plate_str):
    s = safe_str(plate_str).upper().replace(" ", "").replace("-", "")
    if not s or len(s) < 5:
        return ""
    # 標準車牌格式化 (帶連字符)
    if len(s) == 7 and s[:3].isalnum() and s[3:].isdigit():
        return f"{s[:3]}-{s[3:]}"
    elif len(s) == 6 and s[:2].isalnum() and s[2:].isdigit():
        return f"{s[:2]}-{s[2:]}"
    elif len(s) >= 5:
        return s
    return ""

def format_time_val(t_val):
    s = safe_str(t_val)
    if ":" in s:
        parts = s.split(":")
        if len(parts) >= 2:
            return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    return s

def parse_all_records(excel_bytes):
    xls = pd.ExcelFile(excel_bytes)
    all_records = []

    for sheet_name in xls.sheet_names:
        if "空白" in sheet_name or "副本" in sheet_name:
            continue

        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            # 雙行掃描（跳過前 6 列表頭/設定列）
            for i in range(6, len(df) - 1):
                r1 = df.iloc[i]
                r2 = df.iloc[i + 1]

                # --- 1. 快保區塊（左側） ---
                row1_text_kb = " ".join([safe_str(v) for v in r1.iloc[:16]])
                for col_idx in range(2, min(5, len(r1))):
                    candidate_plate = clean_plate(r1.iloc[col_idx])
                    if candidate_plate:
                        if "取消" not in row1_text_kb:
                            all_records.append({
                                "sheet": sheet_name,
                                "type": "快保",
                                "seq": safe_str(r1.iloc[0]),
                                "time": format_time_val(r1.iloc[1]),
                                "plate": candidate_plate,
                                "phone": clean_phone(r2.iloc[col_idx + 1] if (col_idx + 1) < len(r2) else ""),
                                "name": safe_str(r1.iloc[3]),
                                "model": safe_str(r2.iloc[4] if 4 < len(r2) else ""),
                                "advisor": safe_str(r1.iloc[16] if 16 < len(r1) else safe_str(r1.iloc[-1])),
                                "service": safe_str(r1.iloc[5] if 5 < len(r1) else "")
                            })
                        break

                # --- 2. 一般區塊（右側） ---
                if len(r1) > 20:
                    row1_text_yb = " ".join([safe_str(v) for v in r1.iloc[18:]])
                    for col_idx in range(19, min(23, len(r1))):
                        candidate_plate_yb = clean_plate(r1.iloc[col_idx])
                        if candidate_plate_yb:
                            if "取消" not in row1_text_yb:
                                all_records.append({
                                    "sheet": sheet_name,
                                    "type": "一般",
                                    "seq": safe_str(r1.iloc[18]),
                                    "time": format_time_val(r1.iloc[19]),
                                    "plate": candidate_plate_yb,
                                    "phone": clean_phone(r2.iloc[col_idx + 1] if (col_idx + 1) < len(r2) else ""),
                                    "name": safe_str(r1.iloc[21] if 21 < len(r1) else ""),
                                    "model": safe_str(r2.iloc[22] if 22 < len(r2) else ""),
                                    "advisor": safe_str(r1.iloc[34] if 34 < len(r1) else safe_str(r1.iloc[-1])),
                                    "service": safe_str(r1.iloc[23] if 23 < len(r1) else "")
                                })
                            break

        except Exception:
            continue

    return all_records

# 介面流程
input_url = st.text_input(
    "請貼上 Google 預約表網址：", 
    value="https://docs.google.com/spreadsheets/d/1cUzfaS_SN7sYhmWDVZv__yB92JY8pXjfeb1PjrvB8dc/edit?gid=566275317#gid=566275317"
)

if input_url:
    # 純化 URL，精準抽取 Spreadsheet ID
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", input_url)
    if not match:
        st.error("❌ 無法辨識此 Google 試算表網址，請確認網址格式。")
    else:
        sheet_id = match.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

        try:
            req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as response:
                content = response.read()

            records = parse_all_records(io.BytesIO(content))

            if not records:
                st.warning("⚠️ 試算表讀取完畢，但未解析出有效資料，請檢查試算表權限是否已設為「知道連結的人皆可檢視」。")
            else:
                st.success(f"✅ 預約表載入成功！全表共解析出 {len(records)} 筆預約紀錄。")

                raw_query = st.text_input("🔍 請輸入車牌號碼進行查詢 (例如：EBE-5698、ebe5698、5698)：").strip()

                if raw_query:
                    # 將搜尋字串純化 (轉大寫、去空格、去連字符)
                    clean_q = raw_query.upper().replace(" ", "").replace("-", "")

                    results = []
                    for r in records:
                        # 嚴格只針對車牌欄位 (rec['plate']) 進行純化比對
                        rec_plate_clean = r["plate"].upper().replace(" ", "").replace("-", "")
                        
                        # 符合任一條件即命中 (包含純化後車牌或標準格式車牌)
                        if clean_q in rec_plate_clean or raw_query.upper() in r["plate"].upper():
                            results.append(r)

                    if results:
                        st.markdown(f"### 🎯 找到 {len(results)} 筆相符車牌紀錄：")
                        for r in results:
                            with st.expander(f"📅 預約日期：{r['sheet']} | 車牌：{r['plate']} ({r['type']})", expanded=True):
                                c1, c2 = st.columns(2)
                                with c1:
                                    st.write(f"⏱️ **進廠時間**：{r['time']}")
                                    st.write(f"🚘 **車牌號碼**：{r['plate']}")
                                    st.write(f"👤 **車主姓名**：{r['name']}")
                                    st.write(f"📞 **電話號碼**：{r['phone']}")
                                with c2:
                                    st.write(f"🏎️ **車型**：{r['model']}")
                                    st.write(f"👨‍💼 **接待人員**：{r['advisor']}")
                                    st.write(f"🛠️ **類別與序號**：{r['type']} #{r['seq']}")
                                st.write(f"📋 **維修項目**：{r['service']}")
                    else:
                        st.warning(f"查無車牌包含「{raw_query}」的預約紀錄！")

        except Exception as e:
            st.error(f"❌ 下載試算表失敗，請確認網路連線或權限設定：{e}")
