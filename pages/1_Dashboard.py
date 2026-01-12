# C:\Users\USER\Documents\개발 폴더\kvan-dashboard\pages\1_Dashboard.py
import io
import streamlit as st
import pandas as pd
from pathlib import Path
import plotly.express as px
from reports.excel_report import build_monthly_report


# ==============================
# 업체별 색상 팔레트 (고정)
# ==============================
VENDOR_COLORS = {
    "Klook": "#E74C3C",        # 빨강
    "Tripadvisor": "#2ECC71",  # 초록
    "Mozio": "#F39C12",        # 주황
    "MK": "#FF69B4",           # 핑크
    "Kvanlimo": "#1F2A44",     # 네이비
    "Linkro": "#3498DB",       # 파랑 (예정)
}

st.set_page_config(
    layout="wide",
    page_title="KVAN Dashboard",
)

DATA_PATH = Path("data/processed.parquet")

st.markdown("""
<style>

/* ===============================
   Multiselect 태그 색상 통일
   =============================== */

/* 선택된 태그 배경 */
.stMultiSelect [data-baseweb="tag"] {
    background-color: #1F2A44 !important;
    color: #ffffff !important;
    border-radius: 8px;
    font-weight: 600;
}

/* 태그 안 X 아이콘 */
.stMultiSelect [data-baseweb="tag"] svg {
    fill: #ffffff !important;
}

/* hover 상태 */
.stMultiSelect [data-baseweb="tag"]:hover {
    background-color: #243356 !important;
}

/* ===============================
   Section Card
   =============================== */
.section-card {
    background: #ffffff;
    border-radius: 14px;
    padding: 24px;
    box-shadow: 0 6px 18px rgba(0,0,0,0.06);
    margin-bottom: 32px;
}

.section-title {
    font-size: 1.15rem;
    font-weight: 700;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 8px;
}

</style>
""", unsafe_allow_html=True)

# ----------------------------

st.title("📊 Dashboard")

# 데이터 존재 여부
if not DATA_PATH.exists():
    st.warning("아직 업로드된 정산 데이터가 없습니다.")
    st.stop()

df = pd.read_parquet(DATA_PATH)

# -----------------------------
# 필터 영역
# -----------------------------
all_months = sorted(df["month"].unique())

col1, col2, col3 = st.columns([2, 1, 1])

with col1:
    vendors = st.multiselect(
        "업체 선택",
        options=sorted(df["vendor"].unique()),
        default=list(df["vendor"].unique())
    )

with col2:
    start_month = st.selectbox(
        "시작 정산월",
        options=all_months,
        index=0
    )

with col3:
    end_month = st.selectbox(
        "종료 정산월",
        options=all_months,
        index=len(all_months) - 1
    )

filtered = df[
    (df["vendor"].isin(vendors)) &
    (df["month"] >= start_month) &
    (df["month"] <= end_month)
]


if filtered.empty:
    st.info("선택된 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# ==============================
# KPI 스타일 (CSS)
# ==============================
st.markdown(
    """
    <style>
    /* KPI 카드 */
    .kpi-card {
        background-color: #ffffff;
        border-radius: 14px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.06);
        overflow: hidden;
        min-height: 120px;
    }

    /* 상단 컬러 바 */
    .kpi-header {
        background-color: #1F2A44; /* 네이비 */
        padding: 10px 16px;
        text-align: center;
    }

    .kpi-header span {
        color: #ffffff;
        font-size: 0.9rem;
        font-weight: 600;
    }

    /* 값 영역 */
    .kpi-body {
        padding: 22px 16px;
        text-align: center;
    }

    .kpi-value {
        font-size: 1.9rem;
        font-weight: 700;
        color: #1f2937;
        white-space: nowrap;
    }
    </style>
    """,
    unsafe_allow_html=True
)



# KPI 위 여백
st.markdown("<div style='margin-top:20px'></div>", unsafe_allow_html=True)

# ==============================
# KPI 값 계산
# ==============================
total_gross = filtered["gross_sales"].sum()
total_fee = filtered["vendor_fee"].sum()
total_net = filtered["net_sales"].sum()
total_rides = int(filtered["ride_count"].sum())

# ==============================
# KPI 카드 출력
# ==============================
k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-header">
                <span>총 매출액</span>
            </div>
            <div class="kpi-body">
                <div class="kpi-value">{total_gross:,.0f} 원</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with k2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-header">
                <span>총 수수료</span>
            </div>
            <div class="kpi-body">
                <div class="kpi-value">{total_fee:,.0f} 원</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with k3:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-header">
                <span>실 입금액</span>
            </div>
            <div class="kpi-body">            
                <div class="kpi-value">{total_net:,.0f} 원</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with k4:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-header">
                <span>운행 건수</span>
            </div>
            <div class="kpi-body">             
                <div class="kpi-value">{total_rides:,}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ==============================
# 📊 차트 영역 (3열 핵심!)
# ==============================
st.markdown("<div style='margin-top:30px'></div>", unsafe_allow_html=True)

vendor_sum = (
    filtered.groupby("vendor", as_index=False)
    .agg(gross_sales=("gross_sales", "sum"))
)

vendor_unit_price = (
    filtered.groupby("vendor", as_index=False)
    .agg(
        gross_sales=("gross_sales", "sum"),
        ride_count=("ride_count", "sum"),
    )
)
vendor_unit_price["unit_sales"] = (
    vendor_unit_price["gross_sales"] / vendor_unit_price["ride_count"]
)

# 🔥 여기 핵심
col1, col2, col3 = st.columns(3)

# --- 업체별 매출 비교 ---
with col1:
    st.subheader("🏢 업체별 매출 비교")
    fig = px.bar(
        vendor_sum,
        x="vendor",
        y="gross_sales",
        color="vendor",
        color_discrete_map=VENDOR_COLORS,
    )
    fig.update_layout(height=360, showlegend=False, xaxis_title=None, yaxis_title=None)
    fig.update_yaxes(tickformat=",")
    fig.update_layout(annotations=[
        dict(
            x=row.vendor,
            y=row.gross_sales,
            text=f"{row.gross_sales:,.0f} 원",
            showarrow=False,
            yshift=20,
            bgcolor="#F3F4F6",
            bordercolor="#E5E7EB",
            borderpad=4
        )
        for _, row in vendor_sum.iterrows()
    ])
    st.plotly_chart(fig, use_container_width=True)

# --- 업체별 매출 비중 ---
with col2:
    st.subheader("🏢 업체별 매출 비중")
    fig = px.pie(
        vendor_sum,
        names="vendor",
        values="gross_sales",
        hole=0.45,
        color="vendor",
        color_discrete_map=VENDOR_COLORS,
    )
    fig.update_traces(textinfo="percent+label", textposition="inside")
    fig.update_layout(height=360, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

# --- 업체별 건당 매출 ---
with col3:
    st.subheader("🏷️ 업체별 건당 매출")
    fig = px.bar(
        vendor_unit_price,
        x="vendor",
        y="unit_sales",
        color="vendor",
        color_discrete_map=VENDOR_COLORS,
    )
    fig.update_layout(height=360, showlegend=False, xaxis_title=None, yaxis_title=None)
    fig.update_yaxes(tickformat=",")
    fig.update_layout(annotations=[
        dict(
            x=row.vendor,
            y=row.unit_sales,
            text=f"{row.unit_sales:,.0f} 원",
            showarrow=False,
            yshift=16,
            bgcolor="#F3F4F6",
            bordercolor="#E5E7EB",
            borderpad=3
        )
        for _, row in vendor_unit_price.iterrows()
    ])
    st.plotly_chart(fig, use_container_width=True)


# ==============================
# 📈 월별 매출 추이 (개선 버전)
# ==============================
st.subheader("📈 월별 매출 추이")

monthly_total = (
    filtered
    .groupby("month", as_index=False)
    .agg(total_sales=("gross_sales", "sum"))
    .sort_values("month")
)

fig = px.line(
    monthly_total,
    x="month",
    y="total_sales",
)

fig.update_traces(
    mode="lines+markers",
    line=dict(color="#1F2A44", width=3),
    marker=dict(
        size=10,
        color="#1F2A44",
        line=dict(width=2, color="white")
    ),
    hovertemplate="%{x}<br><b>%{y:,.0f} 원</b><extra></extra>"
)


fig.update_layout(
    height=420,
    margin=dict(l=40, r=40, t=10, b=40),
    showlegend=False,
    plot_bgcolor="white",
    paper_bgcolor="white",

    # X축: 월을 균등 간격으로
    xaxis=dict(
        title=None,
        type="category",
        tickangle=0,
        tickfont=dict(size=13)
    ),

    # Y축: 0 기준 고정 (중요)
    yaxis=dict(
        title=None,
        tickformat=",",
        rangemode="tozero",
        gridcolor="#E5E7EB"
    ),
)
annotations = []

for _, row in monthly_total.iterrows():
    annotations.append(
        dict(
            x=row["month"],
            y=row["total_sales"],
            text=f"{row['total_sales']:,.0f} 원",
            showarrow=False,
            yshift=23,  # ⬅ 점 위로 띄우기 (중요)
            font=dict(
                size=14,
                color="#111827",   # 진한 네이비/블랙
                family="Arial Black"
            ),
            bgcolor="#F3F4F6",     # ⬅ 연한 회색 박스
            bordercolor="#E5E7EB",
            borderwidth=1,
            borderpad=4,           # ⬅ padding
            opacity=0.95
        )
    )

fig.update_layout(annotations=annotations)
max_y = monthly_total["total_sales"].max()

fig.update_yaxes(
    range=[0, max_y * 1.15]  # ⬅ 상단 15% 여백
)

st.plotly_chart(fig, use_container_width=True)
st.markdown("</div>", unsafe_allow_html=True)

# ==============================
# 📅 월별 요약 (좌측 정렬 통일)
# ==============================
st.subheader("📅 월별 요약")

monthly = (
    filtered
    .groupby(["month", "vendor"], as_index=False)
    .agg(
        gross_sales=("gross_sales", "sum"),
        vendor_fee=("vendor_fee", "sum"),
        net_sales=("net_sales", "sum"),
        ride_count=("ride_count", "sum"),
    )
    .sort_values("month")
)

display_df = monthly.rename(columns={
    "month": "정산월",
    "vendor": "업체",
    "gross_sales": "매출액",
    "net_sales": "실입금액",
    "vendor_fee": "수수료",
    "ride_count": "운행 건수",
})

# 🔹 모든 숫자 컬럼 → 문자열로 변환 (좌측 정렬 핵심)
for col in ["매출액", "실입금액", "수수료"]:
    display_df[col] = display_df[col].map(lambda x: f"{x:,.0f}")

display_df["운행 건수"] = display_df["운행 건수"].map(lambda x: f"{int(x)}")

# 🔹 dataframe 출력 (좌측 정렬 통일)
st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True
)


# -----------------------------
# 📥 엑셀 다운로드 (보고용 포맷)
# -----------------------------
excel_buffer = build_monthly_report(
    df=monthly,
    vendors=vendors,
    start_month=start_month,
    end_month=end_month,
)

st.download_button(
    label="📥 현재 조회 조건 엑셀 다운로드",
    data=excel_buffer,
    file_name=f"KVAN_Report_{start_month}_{end_month}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
