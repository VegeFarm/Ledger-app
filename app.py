import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
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


def get_excel_modified_dt(file_bytes: bytes) -> Optional[datetime]:
    """
    엑셀 문서 속성(docProps/core.xml)의 modified 날짜/시간을 읽어옵니다.
    없으면 None.
    """
    try:
        bio = decrypt_excel_bytes(file_bytes, PASSWORD)
        bio.seek(0)

        with zipfile.ZipFile(bio) as z:
            if "docProps/core.xml" not in z.namelist():
                return None
            core_xml = z.read("docProps/core.xml")

        root = ET.fromstring(core_xml)
        ns = {
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
            "xsi": "http://www.w3.org/2001/XMLSchema-instance",
        }

        mod_el = root.find("dcterms:modified", ns)
        if mod_el is None or not (mod_el.text or "").strip():
            return None

        txt = mod_el.text.strip()
        # 예: 2025-12-20T12:34:56Z
        if txt.endswith("Z"):
            txt = txt[:-1] + "+00:00"

        return datetime.fromisoformat(txt)
    except Exception:
        return None


def to_number(series: pd.Series) -> pd.Series:
    # 숫자/문자 섞여 있어도 안전하게 숫자로 변환 (콤마, 원, 공백 등 제거)
    return pd.to_numeric(
        series.astype(str).str.replace(r"
