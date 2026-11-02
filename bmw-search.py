import streamlit as st
import pandas as pd
import re
import io
import urllib.request

st.set_page_config(page_title="BMW 預約車輛快速查詢", layout="centered", page_icon="🚗")

st.title("🚗 BMW 預約車輛快速查詢")

def safe_get(row, col_idx):
    if col_idx in row.index and pd.notna(row[col_idx]):
        val = str(row[col_idx]).strip()
        if val.lower() != "nan":
            return val
    return ""

def clean_phone(raw_val):
    if raw_val is None or pd.isna(raw_val):
        return ""
    s = str(raw_val).strip()
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    if digits.startswith("0"):
        return digits
    if digits.startswith("9") and len(digits) == 9:
        return "0" + digits
    return digits

def clean_plate(plate_str):
    if not plate_str:
        return ""
    s = str(plate_str).strip().upper().replace(" ", "").replace("-", "")
    if s == "NAN":
        return ""
    if len(s) == 7 and s[:3].isalpha() and s[3:].isdigit():
        return f"{s[:3]}-{s[3:]}"
    elif len(s) == 6 and s[:2].isalpha() and s[2:].isdigit():
        return f"{s[:2]}-{s[2:]}"
    return s

def format_time_val(t_val):
    if not t_val:
        return ""
    s = str(t_val).strip()
    if ":" in s:
        parts = s.split(":")
        if len(parts) >= 2:
            return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    return s

def is_valid_seq(seq_raw):
    if not seq_raw:
        return False
    clean = str(seq_raw).replace(".", "").strip()
    return clean.isdigit()

def parse_all_records(excel_bytes):
    xls = pd.ExcelFile(excel_bytes)
    all_records = []
    errors = []

    for sheet_name in xls.sheet_names:
        if "空白" in sheet_name or "副本" in sheet_name:
            continue

        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            max_kb, max_yb = 46, 14
            for r_idx in range(3, min(7, len(df))):
                row = df.iloc[r_idx]
                row_str_kb = " ".join([str(v) for v in row.iloc[:12] if pd.notna(v)])
                row_str_yb = " ".join([str(v) for v in row.iloc[16:26] if pd.notna(v)])
                m_kb = re.search(r"(\d+)\s*FUXING", row_str_kb, re.I)
                if m_kb:
                    max_kb = int(m_kb.group(1))
                m_yb = re.search(r"(\d+)\s*FUXING", row_str_yb, re.I)
                if m_yb:
                    max_yb = int(m_yb.group(1))

            for i in range(7, len(df) - 1, 2):
                r1 = df.iloc[i]
                r2 = df.iloc[i + 1]

                # 1. 快保 (左側)
                seq_raw_kb = safe_get(r1, 0)
                if is_valid_seq(seq_raw_kb):
                    seq_num = int(float(seq_raw_kb))
                    if 1 <= seq_num <= max_kb:
                        plate_kb = clean_plate(safe_get(r1, 2))
                        advisor_kb = safe_get(r1, 16)
                        service_kb = safe_get(r1, 5)
                        raw_phone_kb = r2[3] if 3 in r2.index else None

                        is_cancelled = ("取消" in advisor_kb) or ("取消" in service_kb)
                        if not is_cancelled and plate_kb:
                            all_records.append({
                                "sheet": sheet_name,
                                "type": "快保",
                                "seq": str(seq_num),
                                "time": format_time_val(safe_get(r1, 1)),
                                "plate": plate_kb,
                                "phone": clean_phone(raw_phone_kb),
                                "name": safe_get(r1, 3),
                                "model": safe_get(r2, 4),
                                "advisor": advisor_kb,
                                "service": service_kb
                            })

                # 2. 一般 (右側)
                seq_raw_yb = safe_get(r1, 18)
                if is_valid_seq(seq_raw_yb):
                    seq_num_yb = int(float(seq_raw_yb))
                    if 1 <= seq_num_yb <= max_yb:
                        plate_yb = clean_plate(safe_get(r1, 20))
                        advisor_yb = safe_get(r1, 34)
                        service_yb = safe_get(r1, 23)
                        raw_phone_yb = r2[21] if 21 in r2.index else None

                        is_cancelled = ("取消" in advisor_yb) or ("取消" in service_yb)
                        if not is_cancelled and plate_yb:
                            all_records.append({
                                "sheet": sheet_name,
                                "type": "一般",
                                "seq": str(seq_num_yb),
                                "time": format_time_val(safe_get(r1, 19)),
                                "plate": plate_yb,
                                "phone": clean_phone(raw_phone_yb),
                                "name": safe_get(r1, 21),
                                "model": safe_get(r2, 22),
                                "advisor": advisor_yb,
                                "service": service_yb
                            })
        except Exception as err:
            errors.append(f"工作表 {sheet_name} 解析失敗: {str(err)}")

    return all_records, errors

# 介面流程
sheet_url = st.text_input("請貼上 Google 預約表網址：", value="https://docs.google.com/spreadsheets/d/1obpjM1JB8c5wSyC0WNL6qGc-xyo5qDhhF1TiuFPtAxo/edit")

if sheet_url:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        st.error("無法辨識此 Google 試算表網址，請確認格式。")
    else:
        sheet_id = match.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

        try:
            req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as response:
                content = response.read()

            records, parse_errors = parse_all_records(io.BytesIO(content))
            
            if parse_errors:
                with st.expander("⚠️ 系統解析紀錄與警告"):
                    for e in parse_errors:
                        st.write(e)

            if len(records) == 0:
                st.error("⚠️ 預約表載入失敗，沒有解析到任何資料，請檢查上面展開的詳細錯誤！")
            else:
                st.success(f"預約表載入成功！共解析出 {len(records)} 筆車輛預約紀錄。")

            query = st.text_input("🔍 請輸入車牌號碼進行查詢 (例如：BFG-0989 或 0989)：").strip().upper().replace(" ", "").replace("-", "")

            if query:
                results = []
                for rec in records:
                    plate_clean = rec["plate"].replace("-", "")
                    if query in plate_clean or query in rec["plate"]:
                        results.append(rec)

                if results:
                    st.markdown(f"### 🎯 找到 {len(results)} 筆相符車牌紀錄：")
                    for r in results:
                        with st.expander(f"📅 預約日期：{r['sheet']} | 車牌：{r['plate']} ({r['type']})", expanded=True):
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write(f"⏱️ **進廠時間**：{r['time']}")
                                st.write(f"🚘 **車牌號碼**：{r['plate']}")
                                st.write(f"👤 **車主姓名**：{r['name']}")
                                st.write(f"📞 **電話號碼**：{r['phone']}")
                            with col2:
                                st.write(f"🏎️ **車型**：{r['model']}")
                                st.write(f"👨‍💼 **接待人員**：{r['advisor']}")
                                st.write(f"🛠️ **類別與序號**：{r['type']} #{r['seq']}")
                            st.write(f"📋 **維修項目**：{r['service']}")
                else:
                    st.warning("查無此車牌的預約紀錄！")

        except Exception as e:
            st.error(f"下載或讀取資料失敗：{e}")
