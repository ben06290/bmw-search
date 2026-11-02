import streamlit as st
import pandas as pd
import re
import io
import urllib.request

st.set_page_config(page_title="BMW 預約車輛快速查詢", layout="centered", page_icon="🚗")

# 自訂 CSS：放大輸入框字體與按鈕
st.markdown("""
    <style>
    .stTextInput input {
        font-size: 20px !important;
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

def is_valid_plate(plate_str):
    s = safe_str(plate_str).upper().replace(" ", "").replace("-", "")
    if len(s) < 5 or len(s) > 8:
        return False, ""
    
    has_alpha = any(c.isalpha() for c in s)
    has_digit = any(c.isdigit() for c in s)
    if not (has_alpha and has_digit):
        return False, ""

    if len(s) == 7 and s[:3].isalnum() and s[3:].isdigit():
        return True, f"{s[:3]}-{s[3:]}"
    elif len(s) == 6 and s[:2].isalnum() and s[2:].isdigit():
        return True, f"{s[:2]}-{s[2:]}"
    elif len(s) == 6 and s[:4].isdigit() and s[4:].isalnum():
        return True, f"{s[:4]}-{s[4:]}"
    return True, s

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
        m2 = re.search(r"(\d{1,2})[/月]", name)
        if m2:
            month = int(m2.group(1))
            if 1 <= month <= 12:
                return f"{month} 月"
    return "當月"

@st.cache_data(ttl=300)
def fetch_and_parse(target_url):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", target_url)
    if not match:
        return [], "未知", "網址格式不正確"

    sheet_id = match.group(1)
    
    # 抽取網址中的 gid
    gid_match = re.search(r"[?&]gid=(\d+)", target_url)
    if gid_match:
        gid_str = gid_match.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={gid_str}"
    else:
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
                for i in range(6, len(df) - 1, 2):
                    r1 = df.iloc[i]
                    r2 = df.iloc[i + 1]

                    # 1. 快保
                    seq_kb = safe_str(r1.iloc[0])
                    if seq_kb.replace(".", "").isdigit():
                        for col_idx in range(2, min(6, len(r1))):
                            ok, formatted_plate = is_valid_plate(r1.iloc[col_idx])
                            if ok:
                                advisor_kb = safe_str(r1.iloc[16] if 16 < len(r1) else "")
                                service_kb = safe_str(r1.iloc[5] if 5 < len(r1) else "")
                                if "取消" not in advisor_kb and "取消" not in service_kb:
                                    all_records.append({
                                        "sheet": sheet_name,
                                        "type": "快保",
                                        "seq": seq_kb,
                                        "time": format_time_val(r1.iloc[1]),
                                        "plate": formatted_plate,
                                        "phone": clean_phone(r2.iloc[col_idx + 1] if (col_idx + 1) < len(r2) else ""),
                                        "name": safe_str(r1.iloc[3]),
                                        "model": safe_str(r2.iloc[4] if 4 < len(r2) else ""),
                                        "advisor": advisor_kb,
                                        "service": service_kb
                                    })
                                break

                    # 2. 一般
                    if len(r1) > 18:
                        seq_yb = safe_str(r1.iloc[18])
                        if seq_yb.replace(".", "").isdigit():
                            for col_idx in range(19, min(23, len(r1))):
                                ok, formatted_plate = is_valid_plate(r1.iloc[col_idx])
                                if ok:
                                    advisor_yb = safe_str(r1.iloc[34] if 34 < len(r1) else "")
                                    service_yb = safe_str(r1.iloc[23] if 23 < len(r1) else "")
                                    if "取消" not in advisor_yb and "取消" not in service_yb:
                                        all_records.append({
                                            "sheet": sheet_name,
                                            "type": "一般",
                                            "seq": seq_yb,
                                            "time": format_time_val(r1.iloc[19]),
                                            "plate": formatted_plate,
                                            "phone": clean_phone(r2.iloc[col_idx + 1] if (col_idx + 1) < len(r2) else ""),
                                            "name": safe_str(r1.iloc[21] if 21 < len(r1) else ""),
                                            "model": safe_str(r2.iloc[22] if 22 < len(r2) else ""),
                                            "advisor": advisor_yb,
                                            "service": service_yb
                                        })
                                    break
            except Exception:
                continue

        return all_records, detected_month, ""
    except Exception as e:
        return [], "未知", str(e)

# --- 1. 選擇月份/網址區塊 ---
selected_option = st.selectbox("📅 請選擇要查詢的月份預約表：", list(MONTH_URLS.keys()))

if selected_option == "自訂其他網址":
    input_url = st.text_input("請貼上 Google 預約表網址：", value="")
else:
    input_url = MONTH_URLS[selected_option]

records = []
detected_month = ""

if input_url:
    records, detected_month, err_msg = fetch_and_parse(input_url)
    if err_msg:
        st.error(f"❌ 讀取失敗：{err_msg}，請確認試算表共用權限為「知道連結的人皆可檢視」。")
    elif not records:
        st.warning("⚠️ 試算表已讀取，但未解析出有效資料，請點擊上方重新選擇月份。")
    else:
        st.success(f"✅ 預約表載入成功！這是【{detected_month}】的預約表，共解析出 {len(records)} 筆紀錄。")

st.markdown("---")

# --- 2. 搜尋區塊 (表單與確定按鈕) ---
st.subheader("🔍 車牌號碼快速搜尋")

with st.form(key="search_form"):
    raw_query = st.text_input("請輸入車牌號碼：", placeholder="例如：EBE-5698、ebe5698、5698").strip()
    submit_button = st.form_submit_button(label="🔍 確定搜尋")

if submit_button or raw_query:
    if not records:
        st.warning("請先確認上方預約表已成功載入！")
    elif raw_query:
        clean_q = raw_query.upper().replace(" ", "").replace("-", "")

        results = []
        for r in records:
            rec_plate_clean = r["plate"].upper().replace(" ", "").replace("-", "")
            if clean_q in rec_plate_clean or raw_query.upper() in r["plate"].upper():
                results.append(r)

        if results:
            st.success(f"🎯 找到 {len(results)} 筆相符車牌紀錄：")
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
            st.info(f"ℹ️ 查無車牌包含「{raw_query}」的預約紀錄。")
