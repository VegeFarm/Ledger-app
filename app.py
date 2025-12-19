import io
import re
from typing import Dict, Tuple, List, Optional, Set

import pandas as pd
import streamlit as st

PASSWORD = "0000"

# ----------------------------
# Optional dependency (for password-protected Excel)
# ----------------------------
try:
    import msoffcrypto  # pip install msoffcrypto-tool
except Exception:
    msoffcrypto = None


def _is_zip_xlsx(file_bytes: bytes) -> bool:
    # normal xlsx starts with PK.. (zip)
    return file_bytes[:4] == b"PK\x03\x04"


def decrypt_excel_bytes(file_bytes: bytes, password: str = PASSWORD) -> io.BytesIO:
    """
    Returns a BytesIO that can be read by pandas/openpyxl.
    - If file is normal xlsx(zip), returns as-is.
    - If file is encrypted (OLE), decrypts using msoffcrypto.
    """
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


def to_number(s: pd.Series) -> pd.Series:
    # 숫자/문자 섞여 있어도 안전하게 숫자로 변환 (콤마, 원, 공백 제거)
    return pd.to_numeric(
        s.astype(str).str.replace(r"[^\d\.-]", "", regex=True),
        errors="coerce"
    )


def normalize_text_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .replace({"nan": "", "None": ""})
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def find_col(cols: List[str], candidates: List[str]) -> Optional[str]:
    # 1) exact match
    for c in candidates:
        if c in cols:
            return c
    # 2) normalized match (remove spaces/newlines)
    def norm(x: str) -> str:
        return re.sub(r"\s+", "", str(x or ""))

    cols_norm = {norm(c): c for c in cols}
    for cand in candidates:
        n = norm(cand)
        if n in cols_norm:
            return cols_norm[n]

    # 3) substring match
    for cand in candidates:
        for col in cols:
            if str(cand) and str(cand) in str(col):
                return col
    return None


def read_excel_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    bio = decrypt_excel_bytes(file_bytes, PASSWORD)
    # raw read (no header), then we will set header from first row
    raw = pd.read_excel(bio, sheet_name=None, header=None, engine="openpyxl")
    sheets = {}
    for name, df in raw.items():
        if df.empty:
            continue
        # header is usually first row in your file
        header = df.iloc[0].astype(str).str.strip().tolist()

        # make header unique (avoid duplicate col names)
        seen = {}
        new_cols = []
        for h in header:
            h2 = (h or "").strip()
            if h2.lower() == "nan" or h2 == "":
                h2 = "col"
            cnt = seen.get(h2, 0)
            if cnt == 0:
                new_cols.append(h2)
            else:
                new_cols.append(f"{h2}_{cnt}")
            seen[h2] = cnt + 1

        data = df.iloc[1:].copy()
        data.columns = new_cols
        # drop fully empty rows
        data = data.dropna(how="all")
        sheets[name] = data.reset_index(drop=True)
    return sheets


def compute_from_sheets(sheets: Dict[str, pd.DataFrame]) -> Tuple[int, int]:
    """
    Returns:
      (sum_of_final_order_amount, unique_nonzero_shipping_people_count)
    Across all matched sheets.
    """
    AMOUNT_CANDS = ["최종 상품별 총 주문금액"]
    SHIP_CANDS = ["배송비 합계"]
    BUYER_CANDS = ["구매자명"]
    RECIP_CANDS = ["수취인명"]
    ADDR_CANDS = ["통합배송지", "주소", "배송지", "수취인주소", "수령인주소", "수취인 주소", "수령인 주소"]

    total_amount = 0
    nonzero_people_keys: Set[str] = set()

    for sheet_name, df in sheets.items():
        cols = [str(c).strip() for c in df.columns]

        amount_col = find_col(cols, AMOUNT_CANDS)
        ship_col = find_col(cols, SHIP_CANDS)
        buyer_col = find_col(cols, BUYER_CANDS)
        recip_col = find_col(cols, RECIP_CANDS)
        addr_col = find_col(cols, ADDR_CANDS)

        # this sheet doesn't contain needed metrics
        if amount_col is None and ship_col is None:
            continue

        if amount_col is not None:
            total_amount += int(to_number(df[amount_col]).sum(skipna=True) or 0)

        # unique people with nonzero shipping
        if ship_col is not None:
            ship = to_number(df[ship_col]).fillna(0)
            nonzero_mask = ship != 0

            # Dedup key needs buyer/recipient/address (fallback to empty if missing)
            buyer = normalize_text_series(df[buyer_col]) if buyer_col else pd.Series([""] * len(df))
            recip = normalize_text_series(df[recip_col]) if recip_col else pd.Series([""] * len(df))
            addr = normalize_text_series(df[addr_col]) if addr_col else pd.Series([""] * len(df))

            keys = (buyer + "||" + recip + "||" + addr)
            keys = keys[nonzero_mask].dropna()

            # skip empty keys
            keys = keys[keys.str.replace("||", "", regex=False).str.strip() != ""]
            nonzero_people_keys.update(keys.tolist())

    return total_amount, len(nonzero_people_keys)


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="매출 합계 계산기", layout="wide")
st.title("📊 네이버 매출 엑셀 합계 계산기")

uploaded = st.file_uploader(
    "엑셀 파일 업로드 (비밀번호 0000 고정) — 여러 개 업로드 가능",
    type=["xlsx"],
    accept_multiple_files=True
)

colA, colB = st.columns([1, 2])
with colA:
    calc = st.button("✅ 계산", use_container_width=True)

if calc:
    if not uploaded:
        st.warning("먼저 엑셀 파일을 업로드해 주세요.")
    else:
        per_file_rows = []
        grand_amount = 0
        grand_keys_union: Set[str] = set()
        grand_unique_count = 0  # will recompute by union logic via compute_from_sheets on merged sets? We'll do per-file union by keys.

        # We compute per file, but also global union for shipping keys by recomputing key-sets.
        # For simplicity and accuracy: compute per file totals, and also union keys inside the same pass.
        st.session_state["result"] = None

        progress = st.progress(0)
        for i, f in enumerate(uploaded, start=1):
            try:
                sheets = read_excel_sheets(f.getvalue())

                # compute totals
                amount_sum, unique_count = compute_from_sheets(sheets)

                # For global union, we need actual keys; easiest: re-run compute with key collection:
                # We'll collect keys by a lightweight function inside here.
                # (Keeps code self-contained and accurate across multiple files.)
                # ---- key collection ----
                SHIP_CANDS = ["배송비 합계"]
                BUYER_CANDS = ["구매자명"]
                RECIP_CANDS = ["수취인명"]
                ADDR_CANDS = ["통합배송지", "주소", "배송지", "수취인주소", "수령인주소", "수취인 주소", "수령인 주소"]

                for _, df in sheets.items():
                    cols = [str(c).strip() for c in df.columns]
                    ship_col = find_col(cols, SHIP_CANDS)
                    if ship_col is None:
                        continue
                    buyer_col = find_col(cols, BUYER_CANDS)
                    recip_col = find_col(cols, RECIP_CANDS)
                    addr_col = find_col(cols, ADDR_CANDS)

                    ship = to_number(df[ship_col]).fillna(0)
                    nonzero_mask = ship != 0

                    buyer = normalize_text_series(df[buyer_col]) if buyer_col else pd.Series([""] * len(df))
                    recip = normalize_text_series(df[recip_col]) if recip_col else pd.Series([""] * len(df))
                    addr = normalize_text_series(df[addr_col]) if addr_col else pd.Series([""] * len(df))

                    keys = (buyer + "||" + recip + "||" + addr)
                    keys = keys[nonzero_mask].dropna()
                    keys = keys[keys.str.replace("||", "", regex=False).str.strip() != ""]
                    grand_keys_union.update(keys.tolist())
                # ------------------------

                per_file_rows.append({
                    "파일명": f.name,
                    "최종 상품별 총 주문금액 합계": amount_sum,
                    "배송비≠0 (중복제거 인원수)": unique_count,
                    "배송비 계산(인원×3,500)": unique_count * 3500,
                })
                grand_amount += amount_sum

            except Exception as e:
                per_file_rows.append({
                    "파일명": f.name,
                    "최종 상품별 총 주문금액 합계": None,
                    "배송비≠0 (중복제거 인원수)": None,
                    "배송비 계산(인원×3,500)": None,
                    "오류": str(e),
                })

            progress.progress(i / len(uploaded))

        grand_unique_count = len(grand_keys_union)
        grand_shipping_calc = grand_unique_count * 3500

        summary_df = pd.DataFrame(per_file_rows)
        st.session_state["result"] = (summary_df, grand_amount, grand_unique_count, grand_shipping_calc)

if "result" in st.session_state and st.session_state["result"] is not None:
    summary_df, grand_amount, grand_unique_count, grand_shipping_calc = st.session_state["result"]

    st.subheader("✅ 전체 결과")
    m1, m2, m3 = st.columns(3)
    m1.metric("최종 상품별 총 주문금액 총합", f"{grand_amount:,} 원")
    m2.metric("배송비≠0 중복제거 인원수", f"{grand_unique_count:,} 명")
    m3.metric("인원×3,500 합계", f"{grand_shipping_calc:,} 원")

    st.subheader("파일별 상세")
    st.dataframe(summary_df, use_container_width=True)
