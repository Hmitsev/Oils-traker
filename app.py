import os
import re
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st


# ======================================================
# APP CONFIG
# ======================================================

st.set_page_config(
    page_title="Differences Portal Inter Cars",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ======================================================
# FILES / CONSTANTS
# ======================================================

MASTER_DB_FILE = "master_database.csv"
DIFFERENCES_DB_FILE = "differences_database.csv"

MASTER_COLUMNS = [
    "Доставчик",
    "Vendor No",
    "Вътрешен номер",
    "Активен номер",
    "Ref. Number SUPPLIER",
    "Price",
]

DIFFERENCES_COLUMNS = [
    "Склад за дост.",
    "Доставчик",
    "Vendor No",
    "Начин на подаване",
    "Номер прием",
    "Вътрешен номер",
    "Активен номер",
    "Delivery No",
    "Invoice No",
    "Invoice Date",
    "Ref. Number SUPPLIER",
    "Price",
    "QTY",
    "Received QTY",
    "Difference",
    "Стойност тотал в евро",
    "Дата на подаване",
    "СТАТУС - Попълва се от централата!",
    "№ документа за разлики",
    "Дата на обработка на докумнет",
    "Допълнителен коментар",
]

STATUS_OPTIONS = [
    "",
    "Нова",
    "Подадена",
    "Изпратена",
    "Чака отговор",
    "Получено КИ",
    "Получено ДИ/фактура",
    "Отказани",
    "Анулирана",
    "МАСЛА - обработва се от РМ",
    "Обработва се",
    "Приключена",
]


# ======================================================
# CSS
# ======================================================

st.markdown(
    """
<style>
.block-container {
    padding-top: 0.8rem;
    padding-left: 1.6rem;
    padding-right: 1.6rem;
    max-width: 100%;
}

[data-testid="stHeader"] {
    background: rgba(255,255,255,0);
}

.portal-header {
    width: 100%;
    border-bottom: 4px solid #111;
    padding-bottom: 18px;
    margin-bottom: 18px;
}

.logo-box {
    border: 3px solid #111;
    height: 86px;
    display:flex;
    align-items:center;
    justify-content:center;
    text-align:center;
    font-weight:900;
    font-size:18px;
    background:#fff;
    color:#111;
}

.title-box {
    border: 4px solid #111;
    height: 86px;
    display:flex;
    align-items:center;
    justify-content:center;
    text-align:center;
    font-weight:900;
    font-size:27px;
    letter-spacing:1px;
    background:#fff;
    color:#111;
}

.menu-wrapper {
    border: 4px solid #111;
    padding: 18px;
    margin-bottom: 22px;
    background:#fff;
}

.stButton > button {
    width:100%;
    min-height:58px;
    border:3px solid #111;
    border-radius:0px;
    background:#fff;
    color:#111;
    font-weight:900;
    font-size:15px;
}

.stButton > button:hover {
    border:3px solid #ff4b4b;
    color:#ff4b4b;
    background:#fff7f7;
}

.section-title {
    font-size:26px;
    font-weight:900;
    margin-top:10px;
    margin-bottom:10px;
}

.info-box {
    border-left:6px solid #111;
    background:#f8fafc;
    padding:14px 18px;
    margin-bottom:16px;
    color:#111;
}

.small-muted {
    color:#6b7280;
    font-size:13px;
}

.warning-box {
    border-left:6px solid #f59e0b;
    background:#fffbeb;
    padding:13px 16px;
    margin-bottom:14px;
    color:#111;
}

.success-box {
    border-left:6px solid #16a34a;
    background:#f0fdf4;
    padding:13px 16px;
    margin-bottom:14px;
    color:#111;
}
</style>
""",
    unsafe_allow_html=True
)


# ======================================================
# BASIC HELPERS
# ======================================================

def clean_text(value):
    if pd.isna(value):
        return ""

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()

    if isinstance(value, int):
        return str(value)

    value = str(value)
    value = value.replace("\n", " ").replace("\r", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def clean_dataframe_as_text(df):
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].map(clean_text)
    return df


def normalize_columns(df):
    df = df.copy()
    df.columns = [clean_text(c) for c in df.columns]
    return df


def normalize_header(value):
    value = clean_text(value).lower()
    value = value.replace(".", "")
    value = value.replace("№", "no")
    value = value.replace("номер", "no")
    value = value.replace("_", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_key(value):
    value = clean_text(value).upper()
    value = value.replace(" ", "")
    value = value.replace(".", "")
    value = value.replace("-", "")
    value = value.replace("/", "")
    value = value.replace("\\", "")
    value = value.replace(":", "")
    return value.strip()


def safe_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0)


def dataframe_hash(df):
    if df is None or df.empty:
        return "empty"

    raw = df.to_csv(index=False).encode("utf-8", errors="ignore")
    return hashlib.md5(raw).hexdigest()


def get_excel_files():
    files = []

    for file in os.listdir("."):
        if file.startswith("~$"):
            continue

        if file.lower().endswith(".xlsx"):
            files.append(file)

    return files


def to_tsv(df):
    if df is None or df.empty:
        return ""

    return df.to_csv(index=False, sep="\t")


# ======================================================
# MASTER DATABASE BUILDING
# ======================================================

def column_aliases():
    return {
        "Доставчик": [
            "доставчик",
            "supplier",
            "vendor name",
        ],
        "Vendor No": [
            "vendor no",
            "vendor number",
            "vendor",
            "доставчик no",
            "доставчик no.",
            "доставчик номер",
            "доставчик no",
            "доставчик №",
            "vendor no.",
        ],
        "Вътрешен номер": [
            "вътрешен номер",
            "вътрешен no",
            "вътрешен no.",
            "вътрешен №",
            "internal no",
            "internal no.",
            "internal number",
            "internal",
            "вътрешен",
        ],
        "Активен номер": [
            "активен номер",
            "активен no",
            "активен no.",
            "активен №",
            "active no",
            "active no.",
            "active number",
            "active",
            "активен",
        ],
        "Ref. Number SUPPLIER": [
            "ref number supplier",
            "ref number",
            "ref no supplier",
            "ref no",
            "ref. number supplier",
            "ref. number",
            "supplier ref",
            "ref",
            "номер supplier",
        ],
        "Price": [
            "price",
            "цена",
            "unit price",
            "purchase price",
        ],
    }


def resolve_master_columns(df):
    """
    Намира реалните имена на колоните към стандартните MASTER_COLUMNS.
    Работи както със sheet 'взимане на данни от този шиит',
    така и със sheet 'Разлики'.
    """

    normalized_map = {}

    for col in df.columns:
        normalized_map[normalize_header(col)] = col

    result = {}

    aliases = column_aliases()

    for standard_col, alias_list in aliases.items():
        found = None

        for alias in alias_list:
            alias_norm = normalize_header(alias)

            if alias_norm in normalized_map:
                found = normalized_map[alias_norm]
                break

        if found is not None:
            result[standard_col] = found

    return result


def sheet_priority(sheet_name):
    name = clean_text(sheet_name).lower()

    if "взимане" in name:
        return 1

    if "данни" in name:
        return 2

    if "разлики" in name:
        return 3

    if "sheet" in name:
        return 4

    return 9


def extract_master_from_sheet(file_name, sheet_name, df):
    df = normalize_columns(df)
    col_map = resolve_master_columns(df)

    required_basic = [
        "Доставчик",
        "Vendor No",
        "Вътрешен номер",
        "Активен номер",
    ]

    if not all(col in col_map for col in required_basic):
        return pd.DataFrame(columns=MASTER_COLUMNS)

    result = pd.DataFrame()

    for col in MASTER_COLUMNS:
        if col in col_map:
            result[col] = df[col_map[col]].map(clean_text)
        else:
            result[col] = ""

    result = clean_dataframe_as_text(result)

    result = result[
        (result["Вътрешен номер"].astype(str).str.strip() != "") |
        (result["Активен номер"].astype(str).str.strip() != "")
    ]

    if result.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    result["source_file"] = file_name
    result["source_sheet"] = sheet_name
    result["source_priority"] = sheet_priority(sheet_name)

    return result


def build_master_database_from_excels():
    """
    Строи Master Database от всички налични .xlsx файлове.
    Приоритет:
    1. sheet с име 'взимане на данни...'
    2. други sheet-ове с нужните колони
    3. sheet 'Разлики'

    Ако има дубликат по Вътрешен + Активен номер,
    пази реда с цена и Ref Number.
    """

    all_rows = []
    source_log = []

    excel_files = get_excel_files()

    for file_name in excel_files:
        try:
            xls = pd.ExcelFile(file_name, engine="openpyxl")
        except Exception:
            continue

        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(
                    file_name,
                    sheet_name=sheet_name,
                    dtype=object,
                    engine="openpyxl"
                )
            except Exception:
                continue

            if df.empty:
                continue

            extracted = extract_master_from_sheet(
                file_name=file_name,
                sheet_name=sheet_name,
                df=df
            )

            if not extracted.empty:
                all_rows.append(extracted)
                source_log.append(f"{file_name} / {sheet_name}")

    if not all_rows:
        empty = pd.DataFrame(columns=MASTER_COLUMNS + ["source_file", "source_sheet", "source_priority"])
        return empty, source_log

    master = pd.concat(all_rows, ignore_index=True)
    master = clean_dataframe_as_text(master)

    # НЕ махаме дублиращите записи.
    # Искаме целият master да остане.

    master = master.reset_index(drop=True)

    return master, source_log


def save_master_database(df):
    df = df.copy()

    for col in MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    extra_cols = [c for c in df.columns if c not in MASTER_COLUMNS]
    df = df[MASTER_COLUMNS + extra_cols]

    df = clean_dataframe_as_text(df)
    df.to_csv(MASTER_DB_FILE, index=False, encoding="utf-8-sig")


def load_master_database():
    if os.path.exists(MASTER_DB_FILE):
        df = pd.read_csv(
            MASTER_DB_FILE,
            dtype=str,
            keep_default_na=False
        )

        df = clean_dataframe_as_text(df)

        for col in MASTER_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        return df

    master, _ = build_master_database_from_excels()
    save_master_database(master)
    return master


# ======================================================
# DIFFERENCES DATABASE
# ======================================================

def save_differences_database(df):
    df = df.copy()

    for col in DIFFERENCES_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[DIFFERENCES_COLUMNS]
    df = clean_dataframe_as_text(df)
    df.to_csv(DIFFERENCES_DB_FILE, index=False, encoding="utf-8-sig")


def load_differences_database():
    if os.path.exists(DIFFERENCES_DB_FILE):
        df = pd.read_csv(
            DIFFERENCES_DB_FILE,
            dtype=str,
            keep_default_na=False
        )

        df = clean_dataframe_as_text(df)

        for col in DIFFERENCES_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        return df[DIFFERENCES_COLUMNS]

    df = pd.DataFrame(columns=DIFFERENCES_COLUMNS)
    save_differences_database(df)
    return df


# ======================================================
# MASTER LOOKUP
# ======================================================

def make_master_lookup(master_df):

    lookup_internal = {}
    lookup_active = {}

    if master_df.empty:
        return lookup_internal, lookup_active

    for _, row in master_df.iterrows():

        internal_key = normalize_key(
            row.get("Вътрешен номер", "")
        )

        active_key = normalize_key(
            row.get("Активен номер", "")
        )

        record = {
            "Доставчик": clean_text(
                row.get("Доставчик", "")
            ),
            "Vendor No": clean_text(
                row.get("Vendor No", "")
            ),
            "Вътрешен номер": clean_text(
                row.get("Вътрешен номер", "")
            ),
            "Активен номер": clean_text(
                row.get("Активен номер", "")
            ),
            "Ref. Number SUPPLIER": clean_text(
                row.get("Ref. Number SUPPLIER", "")
            ),
            "Price": clean_text(
                row.get("Price", "")
            ),
        }

        # Взимаме ПОСЛЕДНИЯ ред от Master
        if internal_key:
            lookup_internal[internal_key] = record

        if active_key:
            lookup_active[active_key] = record

    return lookup_internal, lookup_active

# ======================================================
# NEW CLAIMS UPLOAD - DIRECT FROM SHEET1
# ======================================================

def read_new_claims_upload(uploaded_file):

    try:
        raw = pd.read_excel(
            uploaded_file,
            sheet_name="Sheet1",
            dtype=object,
            engine="openpyxl"
        )

    except Exception as e:
        st.error(f"Грешка при четене на файла: {e}")
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    if raw.empty:
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    raw = normalize_columns(raw)
    raw = clean_dataframe_as_text(raw)

    result = pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    # ==================================================
    # ДАННИ ДИРЕКТНО ОТ SHEET1
    # ==================================================

    result["Склад за дост."] = raw["Location Code"]

    result["Доставчик"] = raw["Vendor Name"]

    result["Vendor No"] = raw["Buy-from Vendor No_"]

    result["Начин на подаване"] = "Upload"

    result["Номер прием"] = raw["Receipt No"]

    result["Вътрешен номер"] = raw["Item No"]

    result["Активен номер"] = raw["Active No"]

    result["Delivery No"] = raw["Delivery No"]

    result["Invoice No"] = raw["Invoice No"]

    result["Invoice Date"] = raw["InvoiceDate"]

    result["Ref. Number SUPPLIER"] = raw["Supplier Ref Num"]

    result["Price"] = raw["Price per Invoice"]

    result["QTY"] = raw["Quantity"]

    result["Received QTY"] = ""

    result["Difference"] = raw["Quantity"]

    # ==================================================
    # ТОТАЛ В ЕВРО
    # ==================================================

    price_num = pd.to_numeric(
        raw["Price per Invoice"],
        errors="coerce"
    ).fillna(0)

    qty_num = pd.to_numeric(
        raw["Quantity"],
        errors="coerce"
    ).fillna(0)

    result["Стойност тотал в евро"] = (
        price_num * qty_num
    ).round(2)

    result["Дата на подаване"] = datetime.now().strftime(
        "%d.%m.%Y"
    )

    result["СТАТУС - Попълва се от централата!"] = "Нова"

    result["№ документа за разлики"] = ""

    result["Дата на обработка на докумнет"] = ""

    result["Допълнителен коментар"] = ""

    result = result[DIFFERENCES_COLUMNS]

    result = clean_dataframe_as_text(result)

    result = result[
        (result["Вътрешен номер"].astype(str).str.strip() != "") |
        (result["Активен номер"].astype(str).str.strip() != "")
    ]

    return result.reset_index(drop=True)
