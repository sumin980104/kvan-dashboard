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
    header_fill = PatternFill("solid", fgColor="1F2A44")  # 네이비
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
