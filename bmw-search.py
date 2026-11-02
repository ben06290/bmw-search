import streamlit as st
import pandas as pd
import re
import io
import urllib.request

st.set_page_config(page_title="BMW 預約車輛快速查詢", layout="centered", page_icon="🚗")

st.markdown("""
    <style>
    .stTextInput input {
        font-size: 19px !important;
        font-weight: bold !important;
        padding: 10px !important;
    }
    .stButton>button {
        font-size: 18px !important;
        font-weight: bold !important;
        width: 100% !important;
        padding: 10px !important;
        border-radius: 8px !important;
        background-color: #0066cc !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 BMW 預約車輛快速查詢")

MONTH_URLS = {
    "2026年10月預約表": "https://docs.google.com/spreadsheets/d/1obpjM1JB8c5wSyC0WNL6qGc-xyo5qDhhF1TiuFPtAxo/edit?gid=566275317#gid=566275317",
    "2026年11月預約表": "https://docs.google.com/spreadsheets/d/1cUzfaS_SN7sYhmWDVZv__yB92JY8pXjfeb1PjrvB8dc/edit?gid=566275317#gid=566275317",
    "2026年12月預約表": "https://docs.google.com/spreadsheets/d/15oNLBZ89FMkPtD5Wn9sslYSwF0szpO2ggmp0YVqWyHI/edit?gid=1276743763#gid=1276743763",
    "自訂其他網址": ""
}

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

def is_plate_format(val_str):
    s = re.sub(r"[^A-Za-z0-9]", "", safe_str(val_str)).upper()
    if 5 <= len(s) <= 8 and any(c.isalpha() for c in s) and any(c.isdigit() for c in s):
        if len(s) == 7 and s[:3].isalpha() and s[3:].isdigit():
            return True, f"{s[:3]}-{s[3:]}"
        elif len(s) == 6 and s[:2].isalpha() and s[2:].isdigit():
            return True, f"{s[:2]}-{s[2:]}"
        elif len(s) == 6 and s[:4].isdigit() and s[4:].isalpha():
            return True, f"{s[:4]}-{s[4:]}"
        return True, safe_str(val_str).upper()
    return False, ""

def format_time_val(t_val):
    s = safe_str(t_val)
    if ":" in s:
        parts = s.split(":")
        if len(parts) >= 2:
            return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    return s

def extract_month_from_sheets(sheet_names):
    for name in sheet_names:
        if "空白" in name or "副本" in name:
            continue
        m = re.search(r"(\d{1,2})(\d{2})", name)
        if m:
            month = int(m.group(1))
            if 1 <= month <= 12:
                return f"{month} 月"
    return "當月"

def fetch_and_parse(target_url):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", target_url)
    if not match:
        return [], "未知", "網址格式不正確，找不到試算表 ID"

    sheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    try:
        req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as response:
            content = response.read()

        xls = pd.ExcelFile(io.BytesIO(content))
        all_records = []
        detected_month = extract_month_from_sheets(xls.sheet_names)

        for sheet_name in xls.sheet_names:
            if "空白" in sheet_name or "副本" in sheet_name:
                continue
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                if len(df) < 2:
                    continue

                for i in range(len(df)):
                    row = df.iloc[i]
                    for col_idx in range(len(row)):
                        cell_val = row.iloc[col_idx]
                        is_plate, plate_num = is_plate_format(cell_val)
                        if is_plate:
                            seq_val = safe_str(row.iloc[0]) if col_idx > 0 else ""
                            time_val = format_time_val(row.iloc[1]) if len(row) > 1 else ""
                            
                            # 尋找電話號碼
                            phone_val = ""
                            next_row = df.iloc[i + 1] if i + 1 < len(df) else None
                            
                            for c in range(len(row)):
                                p = clean_phone(row.iloc[c])
                                if len(p) >= 9:
                                    phone_val = p
                                    break
                            if not phone_val and next_row is not None:
                                for c in range(len(next_row)):
                                    p = clean_phone(next_row.iloc[c])
                                    if len(p) >= 9:
                                        phone_val = p
                                        break

                            # 💡【精準修正】依照一般預表格式鎖定欄位與清洗接待人員名稱
                            is_general = col_idx >= 15 # 判斷為一般維修或快保
                            
                            if not is_general:
                                # 快保區塊
                                advisor_raw = safe_str(row.iloc[16]) if len(row) > 16 else ""
                                service_raw = safe_str(row.iloc[5]) if len(row) > 5 else ""
                                model_val = safe_str(row.iloc[4]) if len(row) > 4 else ""
                            else:
                                # 一般維修區塊
                                advisor_raw = safe_str(row.iloc[34]) if len(row) > 34 else ""
                                service_raw = safe_str(row.iloc[23]) if len(row) > 23 else ""
                                model_val = safe_str(row.iloc[22]) if len(row) > 22 else ""

                            name_val = safe_str(row.iloc[3] if not is_general else row.iloc[21])

                            # 清洗接待人員欄位（如果抓到備註，只取名字或特定代號，如：豐、豪、嘉、孟...等單字或常見名字）
                            advisor = advisor_raw.split("/")[0].strip() if advisor_raw else ""
                            if len(advisor) > 4: # 如果抓出來太長，可能抓錯欄位，嘗試從特定常用欄位再找一次
                                advisor = advisor_raw[0:2]

                            if "取消" not in advisor_raw and "取消" not in service_raw:
                                all_records.append({
                                    "sheet": sheet_name,
                                    "type": "一般" if is_general else "快保",
                                    "seq": seq_val,
                                    "time": time_val,
                                    "plate": plate_num,
                                    "phone": phone_val,
                                    "name": name_val,
                                    "model": model_val,
                                    "advisor": advisor,
                                    "service": service_raw
                                })
                            break 
            except Exception:
                continue

        return all_records, detected_month, ""
    except Exception as e:
        return [], "未知", str(e)

# Session State
if "records" not in st.session_state:
    st.session_state.records = []
if "detected_month" not in st.session_state:
    st.session_state.detected_month = ""
if "load_status" not in st.session_state:
    st.session_state.load_status = ""

# --- 1. 載入區塊 ---
selected_option = st.selectbox("📅 請選擇要查詢的月份預約表：", list(MONTH_URLS.keys()))

with st.form(key="url_form"):
    if selected_option == "自訂其他網址":
        input_url = st.text_input("請貼上 Google 預約表網址：", value="")
    else:
        default_url = MONTH_URLS[selected_option]
        input_url = st.text_input("請確認預約表網址：", value=default_url)

    btn_load_url = st.form_submit_button(label="📥 確定載入試算表")

if btn_load_url:
    if not input_url.strip():
        st.session_state.load_status = "error: 請先輸入或選擇有效的 Google 預約表網址！"
    else:
        with st.spinner("正在下載並解析預約表資料，請稍候..."):
            recs, mon, err = fetch_and_parse(input_url.strip())
            if err:
                st.session_state.load_status = f"error: {err}"
                st.session_state.records = []
            elif not recs:
                st.session_state.load_status = "empty"
                st.session_state.records = []
            else:
                st.session_state.load_status = "success"
                st.session_state.records = recs
                st.session_state.detected_month = mon

if st.session_state.load_status == "success":
    st.success(f"✅ 預約表載入成功！這是【{st.session_state.detected_month}】的預約表，共解析出 {len(st.session_state.records)} 筆紀錄。")
elif st.session_state.load_status == "empty":
    st.warning("⚠️ 試算表讀取完畢，但未解析出有效資料。")
elif st.session_state.load_status.startswith("error"):
    st.error(f"❌ 讀取失敗：{st.session_state.load_status.replace('error: ', '')}")

st.markdown("---")

# --- 2. 搜尋區塊 ---
st.subheader("🔍 車牌號碼快速搜尋")

with st.form(key="search_form"):
    raw_query = st.text_input("請輸入車牌號碼：", placeholder="例如：BXZ-9508、bxz9508、9508").strip()
    submit_button = st.form_submit_button(label="🔍 確定搜尋")

if submit_button or raw_query:
    if not st.session_state.records:
        st.warning("⚠️ 請先在上方點擊「📥 確定載入試算表」！")
    elif raw_query:
        clean_q = re.sub(r"[^A-Za-z0-9]", "", raw_query).upper()

        results = []
        for r in st.session_state.records:
            rec_plate_clean = re.sub(r"[^A-Za-z0-9]", "", r["plate"]).upper()
            if clean_q in rec_plate_clean:
                results.append(r)

        if results:
            st.success(f"🎯 成功找到 {len(results)} 筆相符車牌紀錄：")
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
                        st.write(f"🛠️ **類別**：{r['type']}")
                    st.write(f"📋 **維修項目**：{r['service']}")
        else:
            st.info(f"ℹ️ 查無車牌包含「{raw_query}」的預約紀錄。")
