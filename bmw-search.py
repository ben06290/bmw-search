import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="BMW 預約車輛查詢", layout="centered")
st.title("🚗 BMW 預約車輛快速查詢")

# 1. 輸入 Google 試算表連結或貼上資料
sheet_url = st.text_input("請貼上 Google 預約表網址：")

if sheet_url:
    try:
        # 自動轉換成匯出網址
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if match:
            sheet_id = match.group(1)
            export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
            
            # 讀取 Excel
            xls = pd.ExcelFile(export_url)
            st.success("預約表載入成功！")
            
            # 搜尋框
            search_plate = st.text_input("🔍 請輸入車牌號碼進行查詢 (例如：CBN-0939 或 0939)：").strip().upper()
            
            if search_plate:
                found = False
                # 遍歷所有日期工作表進行搜尋
                for sheet_name in xls.sheet_names:
                    if "空白" in sheet_name or "副本" in sheet_name:
                        continue
                    
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    
                    # 進行簡易關鍵字比對
                    for idx, row in df.iterrows():
                        row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
                        if search_plate in row_str.replace("-", ""):
                            found = True
                            st.info(f"📅 **預約日期**：{sheet_name}")
                            st.write(f"📋 **詳細紀錄**：{row_str}")
                            st.markdown("---")
                
                if not found:
                    st.warning("查無此車牌的預約紀錄！")
    except Exception as e:
        st.error(f"讀取資料失敗，請確認試算表權限已開啟：{e}")