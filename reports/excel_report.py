# C:\Users\USER\Documents\개발 폴더\kvan-dashboard\reports\excel_report.py
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date

today_str = date.today().strftime("%Y-%m-%d")

def build_monthly_report(df, vendors, start_month, end_month):
    wb = Workbook()
    # =========================
    # 1️⃣ 시트 1 : 월별 요약 (기존)
    # =========================
    ws = wb.active
    ws.title = "월별 업체 매출"
    

    # =========================
    # 스타일 정의
    # =========================
    header_fill = PatternFill("solid", fgColor="2F3A4A")  # 네이비
    header_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # =========================
    # 제목
    # =========================
    ws.merge_cells("A1:F1")
    ws["A1"] = "해외부 월별 업체 매출 "
    ws["A1"].font = Font(bold=True, size=20)
    ws["A1"].alignment = center

    ws.merge_cells("A2:F2")
    ws["A2"] = f"업체: {', '.join(vendors)} | 기간: {start_month} ~ {end_month}"
    ws["A2"].alignment = center
    
    ws["A3"] = f"작성일: {today_str}"
    ws["A3"].alignment = Alignment(horizontal="left", vertical="center")

    ws["A4"] = "담당자: 이수민"
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center")


    # =========================
    # 헤더 (직접 작성)
    # =========================
    headers = ["월", "업체", "매출액", "업체 수수료", "실 입금액", "운행건수"]
    ws.append([])
    ws.append(headers)

    header_row_idx = ws.max_row

    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(row=header_row_idx, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    # =========================
    # 데이터 행
    # =========================
    for _, r in df.iterrows():
        ws.append([
            r["month"],
            r["vendor"],
            r["gross_sales"],
            r["vendor_fee"],
            r["net_sales"],
            r["ride_count"],
        ])

        row_idx = ws.max_row

        for col_idx in range(1, 7):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = border

            if col_idx >= 3:
                cell.number_format = "#,##0"
                cell.alignment = center
            else:
                cell.alignment = center

    # =========================
    # Grand Total
    # =========================
    ws.append([
        "합계",
        "TOTAL",
        df["gross_sales"].sum(),
        df["vendor_fee"].sum(),
        df["net_sales"].sum(),
        df["ride_count"].sum(),
    ])

    total_row_idx = ws.max_row

    for col_idx in range(1, 7):
        cell = ws.cell(row=total_row_idx, column=col_idx)
        cell.font = bold_font
        cell.border = border

        if col_idx >= 3:
            cell.number_format = "#,##0"
            cell.alignment = center
        else:
            cell.alignment = center

    # =========================
    # 컬럼 너비 고정
    # =========================
    COLUMN_WIDTHS = {
        "A": 20,  # month
        "B": 20,  # vendor
        "C": 25,  # gross_sales
        "D": 25,  # vendor_fee
        "E": 25,  # net_sales
        "F": 20,  # ride_count
    }

    for col, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width

    # =========================
    # 2️⃣ 시트 2 : 월 통합 매출 
    # =========================
    ws2 = wb.create_sheet(title="월 통합 매출")

    # 제목
    ws2.merge_cells("A1:E1")
    ws2["A1"] = "해외부 월 통합 매출"
    ws2["A1"].font = Font(bold=True, size=18)
    ws2["A1"].alignment = center

    ws2.merge_cells("A2:E2")
    ws2["A2"] = f"기간: {start_month} ~ {end_month}"
    ws2["A2"].alignment = center

    # 헤더
    headers = ["월", "매출액", "업체 수수료", "실 입금액", "운행 건수"]
    ws2.append([])
    ws2.append(headers)

    header_row = ws2.max_row
    for col_idx in range(1, 6):
        cell = ws2.cell(row=header_row, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    # 🔥 월 통합 데이터
    monthly_total = (
        df.groupby("month", as_index=False)
        .agg(
            gross_sales=("gross_sales", "sum"),
            vendor_fee=("vendor_fee", "sum"),
            net_sales=("net_sales", "sum"),
            ride_count=("ride_count", "sum"),
        )
        .sort_values("month")
    )

    for _, r in monthly_total.iterrows():
        ws2.append([
            r["month"],
            r["gross_sales"],
            r["vendor_fee"],
            r["net_sales"],
            r["ride_count"],
        ])

        row_idx = ws2.max_row
        for col_idx in range(1, 6):
            cell = ws2.cell(row=row_idx, column=col_idx)
            cell.border = border
            cell.alignment = center
            if col_idx >= 2:
                cell.number_format = "#,##0"

    # 합계
    ws2.append([
        "합계",
        monthly_total["gross_sales"].sum(),
        monthly_total["vendor_fee"].sum(),
        monthly_total["net_sales"].sum(),
        monthly_total["ride_count"].sum(),
    ])

    total_row = ws2.max_row
    for col_idx in range(1, 6):
        cell = ws2.cell(row=total_row, column=col_idx)
        cell.font = bold_font
        cell.border = border
        cell.alignment = center
        if col_idx >= 2:
            cell.number_format = "#,##0"

        # 컬럼 너비
    widths = [20, 25, 25, 25, 20]
    for i, w in enumerate(widths, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

        # =========================
    # 3️⃣ 시트 : 업체별 월매출
    # =========================
    ws3 = wb.create_sheet(title="업체별 월매출")

    # -------------------------
    # 제목
    # -------------------------
    ws3.merge_cells("A1:G1")
    ws3["A1"] = "해외부 월별 업체 매출"
    ws3["A1"].font = Font(bold=True, size=18)
    ws3["A1"].alignment = center

    ws3.merge_cells("A2:G2")
    ws3["A2"] = f"업체: {', '.join(vendors)} | 기간: {start_month} ~ {end_month}"
    ws3["A2"].alignment = center

    ws3["A3"] = f"작성일: {today_str}"
    ws3["A4"] = "담당자: 이수민"

    current_row = 6

    months = sorted(df["month"].unique())

    # -------------------------
    # 업체별 블록
    # -------------------------
    for vendor in vendors:
        vendor_df = df[df["vendor"] == vendor]

        # 업체 헤더
        ws3.merge_cells(start_row=current_row, start_column=1,
                        end_row=current_row, end_column=len(months) + 2)
        ws3.cell(row=current_row, column=1, value=vendor)
        ws3.cell(row=current_row, column=1).fill = header_fill
        ws3.cell(row=current_row, column=1).font = header_font
        ws3.cell(row=current_row, column=1).alignment = center

        current_row += 1

        # 월 헤더
        headers = ["구분"] + months + ["합계"]
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

        for label, col in metrics:
            ws3.cell(row=current_row, column=1, value=label)
            ws3.cell(row=current_row, column=1).alignment = center
            ws3.cell(row=current_row, column=1).border = border

            row_sum = 0

            for i, m in enumerate(months, start=2):
                v = vendor_df[vendor_df["month"] == m][col].sum()
                ws3.cell(row=current_row, column=i, value=v)
                ws3.cell(row=current_row, column=i).border = border
                ws3.cell(row=current_row, column=i).alignment = center
                if col != "ride_count":
                    ws3.cell(row=current_row, column=i).number_format = "#,##0"
                row_sum += v

            ws3.cell(row=current_row, column=len(months) + 2, value=row_sum)
            ws3.cell(row=current_row, column=len(months) + 2).font = bold_font
            ws3.cell(row=current_row, column=len(months) + 2).alignment = center
            if col != "ride_count":
                ws3.cell(row=current_row, column=len(months) + 2).number_format = "#,##0"

            current_row += 1

        current_row += 1  # 업체 간 여백

    # 컬럼 너비
    ws3.column_dimensions["A"].width = 18
    for i in range(2, len(months) + 3):
        ws3.column_dimensions[get_column_letter(i)].width = 18


    # =========================
    # 저장
    # =========================
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
