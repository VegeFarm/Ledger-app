import io
import re
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st

PASSWORD = "0000"

try:
    import msoffcrypto  # pip install msoffcrypto-tool
except Exception:
    msoffcrypto = None


def _is_zip_xlsx(file_bytes: bytes) -> bool:
    return file_bytes[:4] == b"PK\x03\x04"


def decrypt_excel_bytes(file_bytes: bytes, password: str = PASSWORD) -> io.BytesIO:
    if _is_zip_xlsx(file_bytes):
        return io.BytesIO(file_bytes)

    if msoffcrypto is None:
        raise RuntimeError(
            "이 엑셀은 비밀번호로 암호화되어 있어요. requirements.txt에 'msoffcrypto-tool'을 추가해 설치해 주세요."
        )

    decrypted = io.BytesIO()
    office = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
    office.load_key(password=password)
    office.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(r"[^\d\.-]", "", regex=True),
        errors="coerce",
    )


def normalize_text_series(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .replace({"nan": "", "None": ""})
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def _norm_no_space(x: str) -> str:
    return re.sub(r"\s+", "", str(x or "")).strip()


def find_col(cols: List[str], candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in cols:
            return c

    cols_norm = {_norm_no_space(c): c for c in cols}
    for cand in candidates:
        n = _norm_no_space(cand)
        if n in cols_norm:
            return cols_norm[n]

    for cand in candidates:
        for col in cols:
            if str(cand) and str(cand) in str(col):
                return col

    return None


def detect_header_row(df: pd.DataFrame, max_scan: int = 30) -> int:
    must_have = {_norm_no_space("구매자명"), _norm_no_space("수취인명")}
    scan_n = min(max_scan, len(df))
    for r in range(scan_n):
        row_vals = df.iloc[r].astype(str).tolist()
        row_norm_set = set(_norm_no_space(v) for v in row_vals if str(v).strip() != "")
        if must_have.issubset(row_norm_set):
            return r
    return 0


def read_excel_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    bio = decrypt_excel_bytes(file_bytes, PASSWORD)
    raw = pd.read_excel(bio, sheet_name=None, header=None, engine="openpyxl")

    sheets: Dict[str, pd.DataFrame] = {}
    for name, df in raw.items():
        if df is None or df.empty:
            continue

        header_row = detect_header_row(df, max_scan=30)
        header = df.iloc[header_row].astype(str).str.strip().tolist()

        seen = {}
        new_cols = []
        for h in header:
            h2 = (h or "").strip()
            if h2.lower() == "nan" or h2 == "":
                h2 = "col"
            cnt = seen.get(h2, 0)
            new_cols.append(h2 if cnt == 0 else f"{h2}_{cnt}")
            seen[h2] = cnt + 1

        data = df.iloc[header_row + 1 :].copy()
        data.columns = new_cols
        data = data.dropna(how="all").reset_index(drop=True)
        sheets[name] = data

    return sheets


def compute_from_sheets(
    sheets: Dict[str, pd.DataFrame],
    amount_mode: str = "rows",  # "rows" or "order_first"
) -> Tuple[float, Set[str]]:
    AMOUNT_CANDS = ["최종 상품별 총 주문금액"]
    SHIP_CANDS = ["배송비 합계"]
    BUYER_CANDS = ["구매자명"]
    RECIP_CANDS = ["수취인명"]
    ADDR_CANDS = ["통합배송지", "주소", "배송지", "수취인주소", "수령인주소", "수취인 주소", "수령인 주소"]
    ORDER_CANDS = ["주문번호"]

    total_amount = 0.0
    nonzero_people_keys: Set[str] = set()

    for _, df in sheets.items():
        cols = [str(c).strip() for c in df.columns]

        amount_col = find_col(cols, AMOUNT_CANDS)
        ship_col = find_col(cols, SHIP_CANDS)
        buyer_col = find_col(cols, BUYER_CANDS)
        recip_col = find_col(cols, RECIP_CANDS)
        addr_col = find_col(cols, ADDR_CANDS)
        order_col = find_col(cols, ORDER_CANDS)

        # ---- amount ----
        if amount_col is not None:
            amt = to_number(df[amount_col]).fillna(0)

            if amount_mode == "order_first" and order_col is not None:
                tmp = df.copy()
                tmp["_amt"] = amt
                total_amount += float(tmp.groupby(order_col)["_amt"].first().sum() or 0.0)
            else:
                total_amount += float(amt.sum(skipna=True) or 0.0)

        # ---- shipping keys ----
        if ship_col is not None:
            ship = to_number(df[ship_col]).fillna(0)
            nonzero_mask = ship != 0

            buyer = normalize_text_series(df[buyer_col]) if buyer_col else pd.Series([""] * len(df))
            recip = normalize_text_series(df[recip_col]) if recip_col else pd.Series([""] * len(df))
            addr = normalize_text_series(df[addr_col]) if addr_col else pd.Series([""] * len(df))

            keys = (buyer + "||" + recip + "||" + addr)
            keys = keys[nonzero_mask].dropna()
            keys = keys[keys.str.replace("||", "", regex=False).str.strip() != ""]
            nonzero_people_keys.update(keys.tolist())

    return total_amount, nonzero_people_keys


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="매출 합계 계산기", layout="wide")
st.title("📊 네이버 매출 엑셀 합계 계산기")

uploaded_files = st.file_uploader(
    "엑셀 파일 업로드 (비밀번호 0000 고정) — 여러 개 업로드 가능",
    type=["xlsx"],
    accept_multiple_files=True,
)

opt1, opt2 = st.columns([1, 1])

with opt1:
    # ✅ 디폴트 변경: index=1 -> 파일별 합산이 기본값
    shipping_dedupe_mode = st.radio(
        "배송비 인원 계산 방식",
        ["전체 업로드 파일 통합(중복 제거)", "파]()
