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
                                "seq": str(seq
