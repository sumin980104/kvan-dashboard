# C:\Users\USER\Documents\개발 폴더\kvan-dashboard\pages\3_Data_Table.py
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

st.title("📋 Data Table")

# ==============================
# Google Sheet 연결
# ==============================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES,
)

gc = gspread.authorize(creds)

SPREADSHEET_ID = "1hFJB0D2L64m3h3pp7R5j9Bj0cZYinkM-dsA6twE1fZA"
SHEET_NAME = "data"

sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
records = sheet.get_all_records()

if not records:
    st.warning("저장된 데이터가 없습니다.")
    st.stop()

df = pd.DataFrame(records)

# 숫자 컬럼 정리
NUMERIC_COLS = [
    "gross_sales",
    "vendor_fee",
    "fx_fee",
    "exchange_rate",
    "net_sales",
    "ride_count",
]

for col in NUMERIC_COLS:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# -------------------------
# 필터
# -------------------------
col1, col2 = st.columns(2)

with col1:
    vendor_filter = st.multiselect(
        "업체 필터",
        options=sorted(df["vendor"].unique()),
        default=list(df["vendor"].unique())
    )

with col2:
    month_filter = st.multiselect(
        "월 필터",
        options=sorted(df["month"].unique()),
        default=list(df["month"].unique())
    )

filtered = df[
    (df["vendor"].isin(vendor_filter)) &
    (df["month"].isin(month_filter))
]

st.subheader("저장된 데이터")

st.dataframe(
    filtered,
    use_container_width=True,
    hide_index=True
)
