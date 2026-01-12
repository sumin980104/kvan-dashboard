import pandas as pd

def parse_mk(file, month: str):
    df = pd.read_excel(file, sheet_name=0)

    # 실제 데이터 시작 (정산서 구조 기준)
    df = df.iloc[9:].copy()

    df = df.rename(columns={
        "Unnamed: 6": "gross_sales",
        "Unnamed: 7": "vendor_fee"
    })

    df = df[["gross_sales", "vendor_fee"]].dropna()

    df["gross_sales"] = pd.to_numeric(df["gross_sales"], errors="coerce")
    df["vendor_fee"] = pd.to_numeric(df["vendor_fee"], errors="coerce")

    df = df.dropna()

    result = {
        "month": month,
        "vendor": "MK",
        "currency": "KRW",
        "gross_sales": df["gross_sales"].sum(),
        "vendor_fee": df["vendor_fee"].sum(),
        "fx_fee": 0,
        "exchange_rate": 1,
        "net_sales": (df["gross_sales"] - df["vendor_fee"]).sum(),
        "ride_count": len(df)
    }

    return pd.DataFrame([result])
