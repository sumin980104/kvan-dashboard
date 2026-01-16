# C:\Users\USER\Documents\개발 폴더\kvan-dashboard\reports\excel_report.py
import io
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList

today_str = date.today().strftime("%Y-%m-%d")


def build_monthly_report(df, vendors, start_month, end_month):
    wb = Workbook()

    # =========================================================
    # 공통 스타일
    # =========================================================
    NAVY = "1F2A44"
    BG_GRAY = "F2F2F2"
    CARD_GRAY = "EDEDED"
    TEXT_GRAY = "555555"

    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin")
    soft_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # =========================================================
    # Dashboard 시트
    # =========================================================
    ws = wb.active
    ws.title = "Dashboard"

    # --- 전체 배경 연한 회색 ---
    for r in range(1, 80):
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=BG_GRAY)

    # ---------------------------------------------------------
    # 제목
    # ---------------------------------------------------------
    ws.merge_cells("A1:H2")
    ws["A1"] = "해외부 매출 Dashboard"
    ws["A1"].font = Font(bold=True, size=22)
    ws["A1"].alignment = center

    ws.merge_cells("A3:H3")
    ws["A3"] = f"기간: {start_month} ~ {end_month}"
    ws["A3"].alignment = center
    ws["A3"].font = Font(size=12, color=TEXT_GRAY)

    # ---------------------------------------------------------
    # KPI 계산
    # ---------------------------------------------------------
    total_gross = df["gross_sales"].sum()
    total_net = df["net_sales"].sum()
    total_fee = df["vendor_fee"].sum()
    total_rides = int(df["ride_count"].sum())

    kpis = [
        ("총 매출액", f"{total_gross:,.0f} 원"),
        ("실 입금액", f"{total_net:,.0f} 원"),
        ("총 수수료", f"{total_fee:,.0f} 원"),
        ("운행 건수", f"{total_rides:,} 건"),
    ]

    kpi_cols = ["A", "C", "E", "G"]

    for i, (title, value) in enumerate(kpis):
        col = kpi_cols[i]

        # 카드 영역
        ws.merge_cells(f"{col}5:{col}9")
        c = ws[f"{col}5"]
        c.fill = PatternFill("solid", fgColor=CARD_GRAY)
        c.border = soft_border
        c.alignment = center
        c.value = f"{title}\n\n{value}"
        c.font = Font(bold=True, size=15)

    # ---------------------------------------------------------
    # 섹션 : 업체별 매출 분석
    # ---------------------------------------------------------
    ws.merge_cells("A11:H11")
    ws["A11"] = "📊 업체별 매출 분석"
    ws["A11"].font = Font(bold=True, size=14)
    ws["A11"].alignment = Alignment(horizontal="left", vertical="center")

    # ---------------------------------------------------------
    # 업체별 매출 집계 (차트용 데이터)
    # ---------------------------------------------------------
    table_row = 12
    ws.cell(row=table_row, column=1, value="업체").font = bold_font
    ws.cell(row=table_row, column=2, value="매출액").font = bold_font

    vendor_total = (
        df.groupby("vendor", as_index=False)
        .agg(gross_sales=("gross_sales", "sum"))
    )

    r = table_row + 1
    for _, row in vendor_total.iterrows():
        ws.cell(row=r, column=1, value=row["vendor"])
        ws.cell(row=r, column=2, value=row["gross_sales"]).number_format = "#,##0"
        r += 1

    # ---------------------------------------------------------
    # Bar Chart (좌)
    # ---------------------------------------------------------
    bar = BarChart()
    bar.title = "업체별 매출 비교"
    bar.legend = None
    bar.y_axis.majorGridlines = None
    bar.style = 10
    bar.width = 18
    bar.height = 9

    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True

    data = Reference(ws, min_col=2, min_row=table_row,
                     max_row=table_row + len(vendor_total))
    cats = Reference(ws, min_col=1, min_row=table_row + 1,
                     max_row=table_row + len(vendor_total))

    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)

    ws.add_chart(bar, "A13")

    # ---------------------------------------------------------
    # Pie Chart (우)
    # ---------------------------------------------------------
    pie = PieChart()
    pie.title = "업체별 매출 비중"
    pie.width = 18
    pie.height = 9

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)

    ws.add_chart(pie, "E13")

    # ---------------------------------------------------------
    # 섹션 : 월별 매출 추이
    # ---------------------------------------------------------
    ws.merge_cells("A29:H29")
    ws["A29"] = "📈 월별 매출 추이"
    ws["A29"].font = Font(bold=True, size=14)
    ws["A29"].alignment = Alignment(horizontal="left", vertical="center")

    # ---------------------------------------------------------
    # 월별 매출 데이터
    # ---------------------------------------------------------
    line_row = 30
    ws.cell(row=line_row, column=1, value="월").font = bold_font
    ws.cell(row=line_row, column=2, value="매출액").font = bold_font

    monthly = (
        df.groupby("month", as_index=False)
        .agg(gross_sales=("gross_sales", "sum"))
        .sort_values("month")
    )

    r = line_row + 1
    for _, row in monthly.iterrows():
        ws.cell(row=r, column=1, value=row["month"])
        ws.cell(row=r, column=2, value=row["gross_sales"]).number_format = "#,##0"
        r += 1

    # ---------------------------------------------------------
    # Line Chart (🔥 혼자라서 크게)
    # ---------------------------------------------------------
    line = LineChart()
    line.title = "월별 매출 추이"
    line.smooth = True
    line.legend = None
    line.y_axis.majorGridlines = None
    line.width = 36     # ← 핵심
    line.height = 12

    line.dataLabels = DataLabelList()
    line.dataLabels.showVal = True

    data = Reference(ws, min_col=2, min_row=line_row,
                     max_row=line_row + len(monthly))
    cats = Reference(ws, min_col=1, min_row=line_row + 1,
                     max_row=line_row + len(monthly))

    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)

    ws.add_chart(line, "A31")

    # ---------------------------------------------------------
    # 컬럼 너비
    # ---------------------------------------------------------
    for col in ["A","B","C","D","E","F","G","H"]:
        ws.column_dimensions[col].width = 22



    # =========================================================
    # 3️⃣ 시트 : 업체별 월매출 (🔥 완전 수정 🔥)
    # =========================================================
    ws3 = wb.create_sheet(title="업체별 월매출")

    ws3.merge_cells("A1:M1")
    ws3["A1"] = "해외부 월별 업체 매출"
    ws3["A1"].font = Font(bold=True, size=18)
    ws3["A1"].alignment = center

    ws3.merge_cells("A2:M2")
    ws3["A2"] = f"업체: {', '.join(vendors)} | 기간: {start_month} ~ {end_month}"
    ws3["A2"].alignment = center

    ws3["A3"] = f"작성일: {today_str}"
    ws3["A4"] = "담당자: 이수민"

    current_row = 6
    months = sorted(df["month"].unique())

    # --- 헤더 (한 번만)
    headers = ["업체", "구분"] + months + ["합계"]
    for col_idx, h in enumerate(headers, start=1):
        c = ws3.cell(row=current_row, column=col_idx, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = soft_border


    current_row += 1

    metrics = [
        ("매출액", "gross_sales"),
        ("업체 수수료", "vendor_fee"),
        ("실 입금액", "net_sales"),
        ("운행건수", "ride_count"),
    ]

    for vendor in vendors:
        vendor_df = df[df["vendor"] == vendor]
        start_vendor_row = current_row

        for label, col in metrics:
            ws3.cell(row=current_row, column=2, value=label).alignment = center
            ws3.cell(row=current_row, column=2).border = soft_border


            row_sum = 0
            for i, m in enumerate(months, start=3):
                v = vendor_df[vendor_df["month"] == m][col].sum()
                c = ws3.cell(row=current_row, column=i, value=v)
                c.border = soft_border
                c.alignment = center
                if col != "ride_count":
                    c.number_format = "#,##0"
                row_sum += v

            total_col = len(months) + 3
            c = ws3.cell(row=current_row, column=total_col, value=row_sum)
            c.font = bold_font
            c.border = soft_border
            c.alignment = center
            if col != "ride_count":
                c.number_format = "#,##0"

            current_row += 1

        # 업체명 세로 병합 (A열)
        ws3.merge_cells(
            start_row=start_vendor_row,
            start_column=1,
            end_row=current_row - 1,
            end_column=1
        )
        c = ws3.cell(row=start_vendor_row, column=1, value=vendor)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = soft_border

        current_row += 1  # 업체 간 여백

    # =========================
    # 🔥 총계 블록 (모든 업체 합산)
    # =========================
    total_start_row = current_row

    for label, col in metrics:
        ws3.cell(row=current_row, column=2, value=label).alignment = center
        ws3.cell(row=current_row, column=2).border = soft_border


        row_sum = 0
        for i, m in enumerate(months, start=3):
            v = df[df["month"] == m][col].sum()
            c = ws3.cell(row=current_row, column=i, value=v)
            c.border = soft_border
            c.alignment = center
            if col != "ride_count":
                c.number_format = "#,##0"
            row_sum += v

        total_col = len(months) + 3
        c = ws3.cell(row=current_row, column=total_col, value=row_sum)
        c.font = bold_font
        c.border = soft_border
        c.alignment = center
        if col != "ride_count":
            c.number_format = "#,##0"

        current_row += 1

    ws3.merge_cells(
        start_row=total_start_row,
        start_column=1,
        end_row=current_row - 1,
        end_column=1
    )
    c = ws3.cell(row=total_start_row, column=1, value="총계")
    c.fill = header_fill
    c.font = header_font
    c.alignment = center
    c.border = soft_border


    # 컬럼 너비
    ws3.column_dimensions["A"].width = 14
    ws3.column_dimensions["B"].width = 14
    for i in range(3, len(months) + 4):
        ws3.column_dimensions[get_column_letter(i)].width = 18

    # =========================
    # 저장
    # =========================
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
