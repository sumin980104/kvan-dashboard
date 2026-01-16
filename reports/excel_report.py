# C:\Users\USER\Documents\개발 폴더\kvan-dashboard\reports\excel_report.py
import io
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference


today_str = date.today().strftime("%Y-%m-%d")


def build_monthly_report(df, vendors, start_month, end_month):
    wb = Workbook()

    # =========================================================
    # 공통 스타일
    # =========================================================
    header_fill = PatternFill("solid", fgColor="1F2A44")  # 네이비
    header_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    center = Alignment(horizontal="center", vertical="center")

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # =========================================================
    # 1️⃣ Dashboard 시트 (보고용)
    # =========================================================
    ws_dash = wb.active
    ws_dash.title = "Dashboard"

    ws_dash.merge_cells("A1:H1")
    ws_dash["A1"] = "📊 해외부 매출 Dashboard"
    ws_dash["A1"].font = Font(bold=True, size=20)
    ws_dash["A1"].alignment = center

    ws_dash.merge_cells("A2:H2")
    ws_dash["A2"] = f"기간: {start_month} ~ {end_month}"
    ws_dash["A2"].alignment = center

    # -------------------------
    # KPI 계산
    # -------------------------
    total_gross = df["gross_sales"].sum()
    total_net = df["net_sales"].sum()
    total_fee = df["vendor_fee"].sum()
    total_rides = int(df["ride_count"].sum())
    avg_unit = total_gross / total_rides if total_rides else 0

    kpis = [
        ("총 매출액", total_gross),
        ("실 입금액", total_net),
        ("총 수수료", total_fee),
        ("운행 건수", total_rides),
        ("평균 건당 매출", avg_unit),
    ]

    row = 4
    for title, value in kpis:
        ws_dash.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        ws_dash.merge_cells(start_row=row, start_column=4, end_row=row, end_column=8)

        h = ws_dash.cell(row=row, column=1, value=title)
        h.fill = header_fill
        h.font = header_font
        h.alignment = center
        h.border = border

        v = ws_dash.cell(row=row, column=4, value=value)
        v.font = Font(bold=True, size=15)
        v.alignment = center
        v.border = border
        if title != "운행 건수":
            v.number_format = "#,##0"

        row += 1

    # =========================================================
    # 업체별 매출 집계 (차트용 테이블)
    # =========================================================
    chart_table_row = row + 2
    ws_dash.cell(row=chart_table_row, column=1, value="업체").font = bold_font
    ws_dash.cell(row=chart_table_row, column=2, value="매출액").font = bold_font

    vendor_total = (
        df.groupby("vendor", as_index=False)
        .agg(gross_sales=("gross_sales", "sum"))
    )

    r = chart_table_row + 1
    for _, vr in vendor_total.iterrows():
        ws_dash.cell(row=r, column=1, value=vr["vendor"])
        ws_dash.cell(row=r, column=2, value=vr["gross_sales"]).number_format = "#,##0"
        r += 1

    # -------------------------
    # 업체별 매출 Bar 차트
    # -------------------------
    bar = BarChart()
    bar.title = "업체별 매출 비교"
    bar.y_axis.title = "매출액"
    bar.x_axis.title = "업체"

    data = Reference(
        ws_dash,
        min_col=2,
        min_row=chart_table_row,
        max_row=chart_table_row + len(vendor_total),
    )
    cats = Reference(
        ws_dash,
        min_col=1,
        min_row=chart_table_row + 1,
        max_row=chart_table_row + len(vendor_total),
    )

    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)

    ws_dash.add_chart(bar, "J4")

    # -------------------------
    # 업체별 매출 비중 Pie 차트
    # -------------------------
    pie = PieChart()
    pie.title = "업체별 매출 비중"
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)

    ws_dash.add_chart(pie, "J20")

    # =========================================================
    # 월별 매출 추이 테이블
    # =========================================================
    line_table_row = chart_table_row + len(vendor_total) + 4
    ws_dash.cell(row=line_table_row, column=1, value="월").font = bold_font
    ws_dash.cell(row=line_table_row, column=2, value="총 매출액").font = bold_font

    monthly = (
        df.groupby("month", as_index=False)
        .agg(gross_sales=("gross_sales", "sum"))
        .sort_values("month")
    )

    r = line_table_row + 1
    for _, mr in monthly.iterrows():
        ws_dash.cell(row=r, column=1, value=mr["month"])
        ws_dash.cell(row=r, column=2, value=mr["gross_sales"]).number_format = "#,##0"
        r += 1

    # -------------------------
    # 월별 매출 추이 Line 차트
    # -------------------------
    line = LineChart()
    line.title = "월별 매출 추이"
    line.y_axis.title = "매출액"

    data = Reference(
        ws_dash,
        min_col=2,
        min_row=line_table_row,
        max_row=line_table_row + len(monthly),
    )
    cats = Reference(
        ws_dash,
        min_col=1,
        min_row=line_table_row + 1,
        max_row=line_table_row + len(monthly),
    )

    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)

    ws_dash.add_chart(line, "A20")

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
        c.border = border

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
            ws3.cell(row=current_row, column=2).border = border

            row_sum = 0
            for i, m in enumerate(months, start=3):
                v = vendor_df[vendor_df["month"] == m][col].sum()
                c = ws3.cell(row=current_row, column=i, value=v)
                c.border = border
                c.alignment = center
                if col != "ride_count":
                    c.number_format = "#,##0"
                row_sum += v

            total_col = len(months) + 3
            c = ws3.cell(row=current_row, column=total_col, value=row_sum)
            c.font = bold_font
            c.border = border
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
        c.border = border

        current_row += 1  # 업체 간 여백

    # =========================
    # 🔥 총계 블록 (모든 업체 합산)
    # =========================
    total_start_row = current_row

    for label, col in metrics:
        ws3.cell(row=current_row, column=2, value=label).alignment = center
        ws3.cell(row=current_row, column=2).border = border

        row_sum = 0
        for i, m in enumerate(months, start=3):
            v = df[df["month"] == m][col].sum()
            c = ws3.cell(row=current_row, column=i, value=v)
            c.border = border
            c.alignment = center
            if col != "ride_count":
                c.number_format = "#,##0"
            row_sum += v

        total_col = len(months) + 3
        c = ws3.cell(row=current_row, column=total_col, value=row_sum)
        c.font = bold_font
        c.border = border
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
    c.border = border


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
