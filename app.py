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
                # 주문번호당 1번만 합산(첫 값) - 검증용 옵션
                tmp = df.copy()
                tmp["_amt"] = amt
                total_amount += float(tmp.groupby(order_col)["_amt"].first().sum() or 0.0)
            else:
                # 모든 행 합산(기본)
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
    shipping_dedupe_mode = st.radio(
        "배송비 인원 계산 방식",
        ["전체 업로드 파일 통합(중복 제거)", "파일별로 계산 후 합산(파일 간 중복은 유지)"],
        index=0,
    )
with opt2:
    amount_mode = st.radio(
        "주문금액 합계 방식(검증용)",
        ["모든 행 합산(기본)", "주문번호당 1번만 합산(첫 값)"],
        index=0,
    )

amount_mode_key = "rows" if amount_mode.startswith("모든 행") else "order_first"

left, _ = st.columns([1, 2])
with left:
    calc_btn = st.button("✅ 계산", use_container_width=True)

if calc_btn:
    if not uploaded_files:
        st.warning("먼저 엑셀 파일을 업로드해 주세요.")
    else:
        per_file_rows = []
        grand_amount = 0.0

        all_keys_union: Set[str] = set()
        sum_of_file_unique_counts = 0

        progress = st.progress(0)

        for i, f in enumerate(uploaded_files, start=1):
            try:
                sheets = read_excel_sheets(f.getvalue())
                amount_sum, keyset = compute_from_sheets(sheets, amount_mode=amount_mode_key)

                unique_count = len(keyset)
                shipping_calc = unique_count * 3500

                per_file_rows.append({
                    "파일명": f.name,
                    "최종 상품별 총 주문금액 합계": amount_sum,
                    "배송비≠0 (중복제거 인원수)": unique_count,
                    "인원×3,500 합계": shipping_calc,
                })

                grand_amount += amount_sum
                all_keys_union.update(keyset)
                sum_of_file_unique_counts += unique_count

            except Exception as e:
                per_file_rows.append({
                    "파일명": f.name,
                    "최종 상품별 총 주문금액 합계": None,
                    "배송비≠0 (중복제거 인원수)": None,
                    "인원×3,500 합계": None,
                    "오류": str(e),
                })

            progress.progress(i / len(uploaded_files))

        # 선택한 방식에 따라 전체 배송비 인원 결정
        union_count = len(all_keys_union)
        if shipping_dedupe_mode.startswith("전체 업로드"):
            shipping_people_count = union_count
        else:
            shipping_people_count = sum_of_file_unique_counts

        grand_shipping_calc = shipping_people_count * 3500
        summary_df = pd.DataFrame(per_file_rows)

        st.session_state["result"] = {
            "summary_df": summary_df,
            "grand_amount": grand_amount,
            "shipping_people_count": shipping_people_count,
            "grand_shipping_calc": grand_shipping_calc,
            "union_count": union_count,
            "sum_of_file_unique_counts": sum_of_file_unique_counts,
        }

if "result" in st.session_state:
    res = st.session_state["result"]
    summary_df = res["summary_df"]
    grand_amount = res["grand_amount"]
    shipping_people_count = res["shipping_people_count"]
    grand_shipping_calc = res["grand_shipping_calc"]

    union_count = res["union_count"]
    sum_of_file_unique_counts = res["sum_of_file_unique_counts"]

    st.subheader("✅ 전체 결과")

    amount_view = f"{grand_amount:,.0f}" if float(grand_amount).is_integer() else f"{grand_amount:,}"
    shipping_view = f"{grand_shipping_calc:,}"

    m1, m2, m3, m4 = st.columns([1, 1, 1, 1.35])

    m1.metric("최종 상품별 총 주문금액 총합", f"{amount_view} 원")
    m2.metric("배송비≠0 인원수", f"{shipping_people_count:,} 명")
    m3.metric("인원×3,500 합계", f"{shipping_view} 원")

    with m4:
        st.caption("📋 엑셀 복사용 (클릭 → Ctrl+C)")
        st.text_input(
            "최종 상품별 총 주문금액 총합 (표시용 / 콤마)",
            value=amount_view,
            key="copy_total_amount_fmt_only",
        )
        st.text_input(
            "인원×3,500원 합계 (표시용 / 콤마)",
            value=shipping_view,
            key="copy_shipping_fmt_only",
        )

        # 검증용: 두 방식 차이도 바로 확인 가능하게
        if union_count != sum_of_file_unique_counts:
            st.caption(f"※ 파일 간 중복 {sum_of_file_unique_counts - union_count}명 감지됨 (통합 중복제거 시 인원 감소)")

    st.subheader("파일별 상세")
    st.dataframe(summary_df, use_container_width=True)
