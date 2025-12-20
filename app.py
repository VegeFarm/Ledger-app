import io
import re
from typing import Dict, List, Optional, Set, Tuple

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
    # Normal xlsx starts with PK.. (zip)
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


def to_number(series: pd.Series) -> pd.Series:
    # 숫자/문자 섞여 있어도 안전하게 숫자로 변
