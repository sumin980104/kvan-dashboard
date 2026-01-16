# C:\Users\USER\Documents\개발 폴더\kvan-dashboard\reports\excel_report.py
import io
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker

today_str = date.today().strftime("%Y-%m-%d")


def build_monthly_report(df, vendors, start_month, end_month):
    wb = Workbook()

    # =========================================================
    # 공통 스타일
    # =========================================================
    NAVY = "1F2A44"    
    GRAY_BG = "F3F4F6"
    WHITE = "FFFFFF"
    BORDER_GRAY = "D1D5DB"

    bold_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    thin = Side(style="thin", color=BORDER_GRAY)
    soft_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(color="FFFFFF", bold=True)

    # =========================================================
    # 1️⃣ Dashboard 시트
    # =========================================================
    ws = wb.active
    ws.title = "Dashboard"

    # ---- 전체 배경 연회색
    for r in range(1, 300):
        for c in range(1, 20):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=GRAY_BG)

    # ---- 제목
    ws.merge_cells("A1:H2")
    ws["A1"] = "해외부 매출 Dashboard"
    ws["A1"].font = Font(bold=True, size=22)
    ws["A1"].alignment = center

    ws.merge_cells("A3:H3")
    ws["A3"] = f"기간: {start_month} ~ {end_month}"
    ws["A3"].font = Font(size=12)
    ws["A3"].alignment = center

    # =========================================================
    # KPI 카드 (텍스트/숫자 완전 분리)
    # =========================================================
    total_gross = float(df["gross_sales"].sum())
    total_net = float(df["net_sales"].sum())
    total_fee = float(df["vendor_fee"].sum())
    total_rides = int(df["ride_count"].sum())

    kpis = [
        ("총 매출액", f"{total_gross:,.0f} 원"),
        ("실 입금액", f"{total_net:,.0f} 원"),
        ("총 수수료", f"{total_fee:,.0f} 원"),
        ("운행 건수", f"{total_rides:,} 건"),
    ]

    cols = ["A", "C", "E", "G"]

    for i, (title, value) in enumerate(kpis):
        col = cols[i]

        # ── 타이틀 영역
        ws.merge_cells(f"{col}5:{col}6")
        t = ws[f"{col}5"]
        t.value = title
        t.font = Font(bold=True, size=12)
        t.alignment = center
        t.fill = PatternFill("solid", fgColor=GRAY_BG)
        t.border = soft_border

        # ── 값 영역
        ws.merge_cells(f"{col}7:{col}9")
        v = ws[f"{col}7"]
        v.value = value
        v.font = Font(bold=True, size=20)
        v.alignment = center
        v.fill = PatternFill("solid", fgColor=WHITE)
        v.border = soft_border

        # =========================================================
    # 업체별 매출 데이터
    # =========================================================
    ws.merge_cells("A11:H11")
    ws["A11"] = "업체별 매출 분석"
    ws["A11"].font = Font(bold=True, size=14)

    base_row = 12
    vendor_sum = df.groupby("vendor", as_index=False).agg(
        gross_sales=("gross_sales", "sum")
    )

    for i, r in vendor_sum.iterrows():
        ws.cell(row=base_row + i, column=1, value=r["vendor"])
        ws.cell(row=base_row + i, column=2, value=r["gross_sales"]).number_format = "#,##0"

    data = Reference(ws, min_col=2, min_row=base_row,
                     max_row=base_row + len(vendor_sum) - 1)
    cats = Reference(ws, min_col=1, min_row=base_row,
                     max_row=base_row + len(vendor_sum) - 1)

    # =========================================================
    # Bar Chart
    # =========================================================
    bar = BarChart()
    bar.legend = None
    bar.y_axis.majorGridlines = None
    bar.width = 18
    bar.height = 8

    bar.add_data(data, titles_from_data=False)
    bar.set_categories(cats)

    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True
    bar.dataLabels.showCatName = False
    bar.dataLabels.showSerName = False

    ws.add_chart(bar, "A13")

    # =========================================================
    # Donut Chart
    # =========================================================
    pie = PieChart()
    pie.holeSize = 60
    pie.legend = None
    pie.width = 18
    pie.height = 8

    pie.add_data(data, titles_from_data=False)
    pie.set_categories(cats)

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showCatName = True
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showSerName = False

    ws.add_chart(pie, "E13")

    # =========================================================
    # 월별 매출 추이
    # =========================================================
    ws.merge_cells("A29:H29")
    ws["A29"] = "월별 매출 추이"
    ws["A29"].font = Font(bold=True, size=14)

    line_row = 30
    monthly = df.groupby("month", as_index=False).agg(
        gross_sales=("gross_sales", "sum")
    ).sort_values("month")

    for i, r in monthly.iterrows():
        ws.cell(row=line_row + i, column=1, value=r["month"])
        ws.cell(row=line_row + i, column=2, value=r["gross_sales"]).number_format = "#,##0"

    data_line = Reference(ws, min_col=2, min_row=line_row,
                          max_row=line_row + len(monthly) - 1)
    cats_line = Reference(ws, min_col=1, min_row=line_row,
                          max_row=line_row + len(monthly) - 1)

    line = LineChart()
    line.legend = None
    line.smooth = True
    line.width = 36
    line.height = 12

    line.x_axis.tickLblPos = "low"
    line.y_axis.majorGridlines = None

    line.add_data(data_line, titles_from_data=False)
    line.set_categories(cats_line)

    line.dataLabels = DataLabelList()
    line.dataLabels.showVal = True
    line.dataLabels.showCatName = True
    line.dataLabels.showSerName = False

    for s in line.series:
        s.marker = Marker(symbol="circle", size=7)

    ws.add_chart(line, "A31")

    # 컬럼 너비
    for c in ["A","B","C","D","E","F","G","H"]:
        ws.column_dimensions[c].width = 22

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
