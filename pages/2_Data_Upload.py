# C:\Users\USER\Documents\개발 폴더\kvan-dashboard\pages\2_Data_Upload.py
import streamlit as st
import pandas as pd
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

from parsers.mk import parse_mk

# =========================
# Google Sheets 설정
# =========================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES,
)

gc = gspread.authorize(creds)

# 👉 실제 사용할 스프레드시트 이름
SPREADSHEET_NAME = "kvan-dashboard-data"
SHEET_NAME = "data"

sheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_NAME)

# =========================
# 저장 경로
# =========================

st.title("📥 Data Upload")

# =========================
# 업체 선택
# =========================
vendor = st.selectbox(
    "업체 선택",
    ["MK", "Klook", "Mozio", "Tripadvisor", "Kvanlimo", "Linkro"]
)

# =========================
# 공통 입력
# =========================
month = st.text_input(
    "정산 월 (YYYY-MM)",
    placeholder="예: 2025-11"
)

results = []

# ==================================================
# MK: 엑셀 업로드 방식 (기존 로직 그대로)
# ==================================================
if vendor == "MK":
    files = st.file_uploader(
        "MK 엑셀 파일 업로드",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

# ==================================================
# Klook: 수동 입력 방식
# ==================================================
if vendor == "Klook":
    st.subheader("Klook 수동 입력")

    gross_krw = st.number_input("매출액 (원화)", min_value=0, step=1000)
    usd_amount = st.number_input("이체 통화 금액 (USD)", min_value=0.0, step=10.0)
    exchange_rate = st.number_input("적용 환율", min_value=0.0, value=1350.0)
    net_krw = st.number_input("입금액 (원화)", min_value=0, step=1000)
    ride_count = st.number_input("운행 건수", min_value=0, step=1)

# ==================================================
# Mozio (수동 입력)
# ==================================================
if vendor == "Mozio":
    st.subheader("Mozio 수동 입력")

    gross_usd = st.number_input(
        "달러 매출액 (USD)",
        min_value=0.0,
        step=10.0
    )

    exchange_rate = st.number_input(
        "적용 환율",
        min_value=0.0,
        value=1350.0
    )

    net_krw = st.number_input(
        "입금액 (원화, 실매출)",
        min_value=0,
        step=1000
    )

    ride_count = st.number_input(
        "운행 건수",
        min_value=0,
        step=1
    )

# ==================================================
# Tripadvisor (주 단위 고정 입력 5행)
# ==================================================
if vendor == "Tripadvisor":
    st.subheader("Tripadvisor 주별 입력")
    st.caption("※ 최대 5주 입력 / 환전일은 참고용이며 월 계산에는 사용하지 않습니다.")

    trip_df = pd.DataFrame(
        {
            "환전일": [None] * 5,
            "달러 매출액 (USD)": [None] * 5,
            "환율": [None] * 5,
            "운행 건수": [None] * 5,
        }
    )

    edited_df = st.data_editor(
        trip_df,
        num_rows="fixed",          # 🔥 자동 행 추가 완전 차단
        use_container_width=True,
        key="tripadvisor_fixed"
    )
# ==================================================
# Kvanlimo 
# ==================================================
if vendor == "Kvanlimo":
    st.subheader("Kvanlimo 정산 입력")
    st.caption("※ 고정 20행 / 빈 줄은 자동 무시됩니다.")

    kvan_df = pd.DataFrame(
        {
            "환전일": [None] * 20,
            "달러 매출액 (USD)": [None] * 20,
            "수수료 (USD)": [None] * 20,
            "환율": [None] * 20,
            "운행 건수": [None] * 20,
        }
    )

    edited_kvan_df = st.data_editor(
        kvan_df,
        num_rows="fixed",
        use_container_width=True,
        key="kvanlimo_fixed"
    )
# ==================================================
# Linkro (통화 선택 수동 입력)
# ==================================================
if vendor == "Linkro":
    st.subheader("Linkro 정산 입력")

    currency_type = st.radio(
        "입금 통화 선택",
        ["KRW (원화)", "USD (달러)"],
        horizontal=True
    )

    fx_date = st.date_input("환전일 / 결제일")

    # ---------------------------
    # KRW 입금
    # ---------------------------
    if currency_type == "KRW (원화)":
        gross_krw = st.number_input(
            "매출액 (KRW)",
            min_value=0,
            step=1000
        )

        fee_krw = st.number_input(
            "수수료 (KRW, 미입력 시 0)",
            min_value=0,
            step=1000
        )

        ride_count = st.number_input(
            "운행 건수 (미입력 시 1)",
            min_value=0,
            step=1,
            value=1
        )

    # ---------------------------
    # USD 입금
    # ---------------------------
    else:
        gross_usd = st.number_input(
            "매출액 (USD)",
            min_value=0.0,
            step=10.0
        )

        fee_usd = st.number_input(
            "수수료 (USD, 미입력 시 0)",
            min_value=0.0,
            step=1.0
        )

        exchange_rate = st.number_input(
            "적용 환율",
            min_value=0.0,
            value=1350.0
        )

        ride_count = st.number_input(
            "운행 건수 (미입력 시 1)",
            min_value=0,
            step=1,
            value=1
        )


# =========================
# 저장 버튼
# =========================
if st.button("저장"):
    if not month:
        st.warning("정산 월을 입력하세요.")
        st.stop()

    # ----------------------
    # MK 처리
    # ----------------------
    if vendor == "MK":
        if not files:
            st.warning("MK 엑셀 파일을 업로드하세요.")
            st.stop()

        for f in files:
            parsed = parse_mk(f, month)
            results.append(parsed)

    # ----------------------
    # Klook 처리 (수동)
    # ----------------------
    if vendor == "Klook":
        fee = gross_krw - net_krw

        row = {
            "month": month,
            "vendor": "Klook",
            "currency": "USD",
            "gross_sales": gross_krw,
            "vendor_fee": fee,
            "fx_fee": 0,
            "exchange_rate": exchange_rate,
            "net_sales": net_krw,
            "ride_count": ride_count
        }

        results.append(pd.DataFrame([row]))

    # ----------------------
    # Mozio 저장
    # ----------------------
    if vendor == "Mozio":
        if ride_count == 0:
            st.warning("운행 건수를 입력하세요.")
            st.stop()

        MOZIO_FEE_USD = 3.0

        gross_krw = gross_usd * exchange_rate
        mozio_fee_krw = MOZIO_FEE_USD * exchange_rate
        fx_fee = gross_krw - net_krw
        total_fee = mozio_fee_krw + fx_fee

        row = {
            "month": month,
            "vendor": "Mozio",
            "currency": "USD",
            "gross_sales": gross_krw,
            "vendor_fee": total_fee,
            "fx_fee": fx_fee,
            "exchange_rate": exchange_rate,
            "net_sales": net_krw,
            "ride_count": ride_count
        }

        results.append(pd.DataFrame([row]))
    
    # ----------------------
    # Tripadvisor 처리 (고정 5행)
    # ----------------------
    if vendor == "Tripadvisor":

        rows = []

        for _, r in edited_df.iterrows():
            usd = pd.to_numeric(r["달러 매출액 (USD)"], errors="coerce")
            rate = pd.to_numeric(r["환율"], errors="coerce")
            ride = pd.to_numeric(r["운행 건수"], errors="coerce") or 0

        # 🔥 빈 줄은 그냥 무시
            if pd.isna(usd) or pd.isna(rate):
                continue

            gross_krw = usd * rate

            rows.append({
                "month": month,                # ✅ 상단 입력 월만 사용
                "vendor": "Tripadvisor",
                "currency": "USD",
                "gross_sales": gross_krw,
                "vendor_fee": 0,
                "fx_fee": 0,
                "exchange_rate": rate,
                "net_sales": gross_krw,
                "ride_count": ride,
                "fx_date": r["환전일"],        # ✅ 환전일은 참고용으로 저장
            })

        if not rows:
            st.warning("입력된 Tripadvisor 데이터가 없습니다.")
            st.stop()

        results.append(pd.DataFrame(rows))

        # ----------------------
    # Kvanlimo 처리 (고정 20행)
    # ----------------------
    if vendor == "Kvanlimo":

        rows = []

        for _, r in edited_kvan_df.iterrows():
            usd = pd.to_numeric(r["달러 매출액 (USD)"], errors="coerce")
            fee_usd = pd.to_numeric(r["수수료 (USD)"], errors="coerce")
            rate = pd.to_numeric(r["환율"], errors="coerce")
            ride = pd.to_numeric(r["운행 건수"], errors="coerce") or 0

            # 빈 줄 무시
            if pd.isna(usd) or pd.isna(rate) or pd.isna(fee_usd):
                continue

            gross_krw = usd * rate
            fee_krw = fee_usd * rate
            net_krw = gross_krw - fee_krw

            rows.append({
                "month": month,
                "vendor": "Kvanlimo",
                "currency": "USD",
                "gross_sales": gross_krw,
                "vendor_fee": fee_krw,
                "fx_fee": 0,
                "exchange_rate": rate,
                "net_sales": net_krw,
                "ride_count": ride,
                "fx_date": r["환전일"],
            })

        if not rows:
            st.warning("입력된 Kvanlimo 데이터가 없습니다.")
            st.stop()

        results.append(pd.DataFrame(rows))

    # ----------------------
# Linkro 저장
# ----------------------
if vendor == "Linkro":

    # 기본값 보정
    ride = ride_count if ride_count > 0 else 1

    if currency_type == "KRW (원화)":
        fee = fee_krw if fee_krw else 0
        net_krw = gross_krw - fee

        row = {
            "month": month,
            "vendor": "Linkro",
            "currency": "KRW",
            "gross_sales": gross_krw,
            "vendor_fee": fee,
            "fx_fee": 0,
            "exchange_rate": 1,
            "net_sales": net_krw,
            "ride_count": ride,
            "fx_date": fx_date.strftime("%Y-%m-%d") if fx_date else "",

        }

        results.append(pd.DataFrame([row]))

    else:
        fee = fee_usd if fee_usd else 0

        gross_krw = gross_usd * exchange_rate
        fee_krw = fee * exchange_rate
        net_krw = gross_krw - fee_krw

        row = {
            "month": month,
            "vendor": "Linkro",
            "currency": "USD",
            "gross_sales": gross_krw,
            "vendor_fee": fee_krw,
            "fx_fee": 0,
            "exchange_rate": exchange_rate,
            "net_sales": net_krw,
            "ride_count": ride,
            "fx_date": fx_date.strftime("%Y-%m-%d") if fx_date else "",

        }

        results.append(pd.DataFrame([row]))

    

    # ======================
# Google Sheets 저장
# ======================
    new_df = pd.concat(results)

# NaN → 빈 문자열 (Sheets 오류 방지)
    new_df = new_df.fillna("")

# DataFrame → list of lists
    rows = new_df.values.tolist()

# 헤더가 비어 있으면 헤더 먼저 추가
    if not sheet.get_all_values():
        sheet.append_row(list(new_df.columns))

# 데이터 append
    sheet.append_rows(rows)

    st.success("Google Sheets에 저장 완료")
    st.dataframe(new_df, use_container_width=True)

