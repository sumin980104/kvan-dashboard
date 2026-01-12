import streamlit as st
import pandas as pd
from pathlib import Path

DATA_PATH = Path("data/processed.parquet")

st.title("📋 Data Table")

if not DATA_PATH.exists():
    st.warning("저장된 데이터가 없습니다.")
    st.stop()

df = pd.read_parquet(DATA_PATH)

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

# -------------------------
# 편집 가능한 테이블
# -------------------------
edited_df = st.data_editor(
    filtered,
    use_container_width=True,
    num_rows="dynamic"
)

# -------------------------
# 저장 버튼
# -------------------------
if st.button("✏️ 수정 내용 저장"):
    # 원본에서 해당 행 제거 후 다시 합치기
    remaining = df.drop(filtered.index)
    final_df = pd.concat([remaining, edited_df])

    final_df.to_parquet(DATA_PATH, index=False)
    st.success("수정 내용이 저장되었습니다.")
