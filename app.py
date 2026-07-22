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

    master["_key_internal"] = master["Вътрешен номер"].map(normalize_key)
    master["_key_active"] = master["Активен номер"].map(normalize_key)
    master["_dedupe_key"] = master["_key_internal"] + "||" + master["_key_active"]

    master["_has_price"] = master["Price"].astype(str).str.strip().ne("").astype(int)
    master["_has_ref"] = master["Ref. Number SUPPLIER"].astype(str).str.strip().ne("").astype(int)

    master["source_priority"] = pd.to_numeric(
        master["source_priority"],
        errors="coerce"
    ).fillna(9)

    master = master.sort_values(
        by=["source_priority", "_has_price", "_has_ref"],
        ascending=[True, False, False]
    )

    master = master.drop_duplicates(
        subset=["_dedupe_key"],
        keep="first"
    )

    master = master.drop(
        columns=[
            "_key_internal",
            "_key_active",
            "_dedupe_key",
            "_has_price",
            "_has_ref",
        ],
        errors="ignore"
    )

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
        internal_key = normalize_key(row.get("Вътрешен номер", ""))
        active_key = normalize_key(row.get("Активен номер", ""))

        record = {
            "Доставчик": clean_text(row.get("Доставчик", "")),
            "Vendor No": clean_text(row.get("Vendor No", "")),
            "Вътрешен номер": clean_text(row.get("Вътрешен номер", "")),
            "Активен номер": clean_text(row.get("Активен номер", "")),
            "Ref. Number SUPPLIER": clean_text(row.get("Ref. Number SUPPLIER", "")),
            "Price": clean_text(row.get("Price", "")),
        }

        if internal_key and internal_key not in lookup_internal:
            lookup_internal[internal_key] = record

        if active_key and active_key not in lookup_active:
            lookup_active[active_key] = record

    return lookup_internal, lookup_active


def autofill_claims_from_master(claims_df, master_df):
    """
    Попълва New Claims по Master Database.

    Ако има match:
    - налива Ref. Number SUPPLIER
    - налива Price

    Ако няма match:
    - Ref. Number SUPPLIER = 0
    - Price = 0
    """

    claims_df = claims_df.copy()

    for col in DIFFERENCES_COLUMNS:
        if col not in claims_df.columns:
            claims_df[col] = ""

    claims_df = claims_df[DIFFERENCES_COLUMNS]
    claims_df = clean_dataframe_as_text(claims_df)

    lookup_internal, lookup_active = make_master_lookup(master_df)

    filled_rows = []

    for _, row in claims_df.iterrows():
        row = row.copy()

        internal_key = normalize_key(row.get("Вътрешен номер", ""))
        active_key = normalize_key(row.get("Активен номер", ""))

        match = None

        if internal_key in lookup_internal:
            match = lookup_internal[internal_key]

        elif active_key in lookup_active:
            match = lookup_active[active_key]

        if match:
            for field in [
                "Доставчик",
                "Vendor No",
                "Вътрешен номер",
                "Активен номер",
            ]:
                if clean_text(row.get(field, "")) == "":
                    row[field] = match.get(field, "")

            supplier_ref = clean_text(
                match.get("Ref. Number SUPPLIER", "")
            )

            supplier_price = clean_text(
                match.get("Price", "")
            )

            if supplier_ref == "":
                supplier_ref = "0"

            if supplier_price == "":
                supplier_price = "0"

            row["Ref. Number SUPPLIER"] = supplier_ref
            row["Price"] = supplier_price

        else:
            row["Ref. Number SUPPLIER"] = "0"
            row["Price"] = "0"

        if clean_text(row.get("Склад за дост.", "")) == "":
            row["Склад за дост."] = "B01"

        if clean_text(row.get("Дата на подаване", "")) == "":
            row["Дата на подаване"] = datetime.now().strftime("%d.%m.%Y")

        if clean_text(row.get("СТАТУС - Попълва се от централата!", "")) == "":
            row["СТАТУС - Попълва се от централата!"] = "Нова"

        filled_rows.append(row)

    result = pd.DataFrame(filled_rows)
    result = result[DIFFERENCES_COLUMNS]

    return result


# ======================================================
# NEW CLAIMS UPLOAD
# ======================================================

def read_new_claims_upload(uploaded_file):
    """
    Upload файлът е с фиксирана структура.

    Взимаме:
    C = Доставчик
    D = Vendor No / Доставчик №
    E = Номер прием
    F = Вътрешен номер
    G = Активен номер
    I = Quantity -> Difference
    """

    try:
        raw = pd.read_excel(
            uploaded_file,
            sheet_name=0,
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

    if raw.shape[1] < 9:
        st.error("Файлът трябва да има минимум 9 колони, защото използваме C, D, E, F, G и I.")
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    col_c_supplier = raw.columns[2]
    col_d_vendor = raw.columns[3]
    col_e_receipt = raw.columns[4]
    col_f_internal = raw.columns[5]
    col_g_active = raw.columns[6]
    col_i_quantity = raw.columns[8]

    result = pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    result["Склад за дост."] = "B01"
    result["Доставчик"] = raw[col_c_supplier].map(clean_text)
    result["Vendor No"] = raw[col_d_vendor].map(clean_text)
    result["Начин на подаване"] = "Upload"
    result["Номер прием"] = raw[col_e_receipt].map(clean_text)
    result["Вътрешен номер"] = raw[col_f_internal].map(clean_text)
    result["Активен номер"] = raw[col_g_active].map(clean_text)

    result["Delivery No"] = ""
    result["Invoice No"] = ""
    result["Invoice Date"] = ""

    result["Ref. Number SUPPLIER"] = ""
    result["Price"] = ""

    result["QTY"] = ""
    result["Received QTY"] = ""
    result["Difference"] = raw[col_i_quantity].map(clean_text)

    result["Стойност тотал в евро"] = ""
    result["Дата на подаване"] = datetime.now().strftime("%d.%m.%Y")
    result["СТАТУС - Попълва се от централата!"] = "Нова"
    result["№ документа за разлики"] = ""
    result["Дата на обработка на докумнет"] = ""
    result["Допълнителен коментар"] = ""

    result = result[DIFFERENCES_COLUMNS]
    result = clean_dataframe_as_text(result)

    result = result[
        (result["Вътрешен номер"].astype(str).str.strip() != "") |
        (result["Активен номер"].astype(str).str.strip() != "") |
        (result["Номер прием"].astype(str).str.strip() != "")
    ]

    return result.reset_index(drop=True)


# ======================================================
# STATE
# ======================================================

if "page" not in st.session_state:
    st.session_state.page = "Разлики"

if "master_df" not in st.session_state:
    st.session_state.master_df = load_master_database()

if "differences_df" not in st.session_state:
    st.session_state.differences_df = load_differences_database()

if "preview_claims_df" not in st.session_state:
    st.session_state.preview_claims_df = pd.DataFrame(columns=DIFFERENCES_COLUMNS)


# ======================================================
# HEADER
# ======================================================

st.markdown('<div class="portal-header">', unsafe_allow_html=True)

h1, h2, h3 = st.columns([1.2, 5.2, 1.2])

with h1:
    st.markdown(
        """
        <div class="logo-box">
            INTER<br>CARS
        </div>
        """,
        unsafe_allow_html=True
    )

with h2:
    st.markdown(
        """
        <div class="title-box">
            DIFFERENCES PORTAL INTER CARS
        </div>
        """,
        unsafe_allow_html=True
    )

with h3:
    st.markdown(
        """
        <div class="logo-box">
            B01<br>SUPPLIERS
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("</div>", unsafe_allow_html=True)


# ======================================================
# MENU
# ======================================================

st.markdown('<div class="menu-wrapper">', unsafe_allow_html=True)

m1, m2, m3, m4, m5 = st.columns(5)

with m1:
    if st.button("📋 Разлики"):
        st.session_state.page = "Разлики"

with m2:
    if st.button("📦 Master Database"):
        st.session_state.page = "Master Database"

with m3:
    if st.button("➕ New Claims"):
        st.session_state.page = "New Claims"

with m4:
    if st.button("💰 Счетоводство"):
        st.session_state.page = "Счетоводство"

with m5:
    if st.button("⚙️ Админ"):
        st.session_state.page = "Админ"

st.markdown("</div>", unsafe_allow_html=True)


# ======================================================
# CURRENT DATA
# ======================================================

master_df = st.session_state.master_df.copy()
differences_df = st.session_state.differences_df.copy()


# ======================================================
# PAGE: DIFFERENCES
# ======================================================

if st.session_state.page == "Разлики":

    st.markdown('<div class="section-title">📋 Основен списък с разлики</div>', unsafe_allow_html=True)

    total_rows = len(differences_df)

    if "Difference" in differences_df.columns:
        diff_num = safe_numeric(differences_df["Difference"])
        plus_rows = int((diff_num > 0).sum())
        minus_rows = int((diff_num < 0).sum())
    else:
        plus_rows = 0
        minus_rows = 0

    if "СТАТУС - Попълва се от централата!" in differences_df.columns:
        open_rows = int(
            differences_df["СТАТУС - Попълва се от централата!"]
            .astype(str)
            .str.lower()
            .isin(["", "нова", "подадена", "изпратена", "чака отговор", "обработва се"])
            .sum()
        )
    else:
        open_rows = 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Всички редове", total_rows)
    k2.metric("Плюсове", plus_rows)
    k3.metric("Минуси", minus_rows)
    k4.metric("Отворени / необработени", open_rows)

    st.markdown(
        """
        <div class="info-box">
        Основният списък пази същата структура и ред на колоните като Excel файла.
        Редактирай директно в таблицата. Автозаписът работи, когато няма активен филтър.
        </div>
        """,
        unsafe_allow_html=True
    )

    search = st.text_input(
        "Търсене в основния списък",
        placeholder="Доставчик, вътрешен номер, активен номер, номер прием..."
    )

    view_df = differences_df.copy()

    if search.strip():
        query = search.strip().lower()

        mask = view_df.apply(
            lambda row: row.astype(str).str.lower().str.contains(query, na=False).any(),
            axis=1
        )

        view_df = view_df[mask]

    edited_df = st.data_editor(
        view_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "СТАТУС - Попълва се от централата!": st.column_config.SelectboxColumn(
                "СТАТУС - Попълва се от централата!",
                options=STATUS_OPTIONS,
                required=False,
            )
        },
        key="differences_editor"
    )

    if search.strip():
        st.warning("Редактираш филтриран изглед. За запис на масови промени изчисти търсенето.")
    else:
        old_hash = dataframe_hash(differences_df)
        new_hash = dataframe_hash(edited_df)

        if old_hash != new_hash:
            st.session_state.differences_df = edited_df.copy()
            save_differences_database(edited_df)
            st.success("Промените са записани автоматично.")

    c1, c2, c3 = st.columns([1, 1, 5])

    with c1:
        st.download_button(
            "⬇️ Export CSV",
            data=st.session_state.differences_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name="differences_database.csv",
            mime="text/csv"
        )

    with c2:
        if st.button("🔄 Презареди"):
            st.session_state.differences_df = load_differences_database()
            st.rerun()


# ======================================================
# PAGE: MASTER DATABASE
# ======================================================

elif st.session_state.page == "Master Database":

    st.markdown('<div class="section-title">📦 Master Database / Вземане на данни</div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="info-box">
        Master Database се използва за автоматично наливане на 
        <b>Ref. Number SUPPLIER</b> и <b>Price</b> при New Claims.
        Дубликатите се чистят по комбинация <b>Вътрешен номер + Активен номер</b>.
        </div>
        """,
        unsafe_allow_html=True
    )

    price_not_empty = int(master_df["Price"].astype(str).str.strip().ne("").sum()) if "Price" in master_df.columns else 0
    price_empty = int(master_df["Price"].astype(str).str.strip().eq("").sum()) if "Price" in master_df.columns else 0
    ref_not_empty = int(master_df["Ref. Number SUPPLIER"].astype(str).str.strip().ne("").sum()) if "Ref. Number SUPPLIER" in master_df.columns else 0

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Master редове", len(master_df))
    c2.metric("С цена", price_not_empty)
    c3.metric("Без цена", price_empty)
    c4.metric("С Ref.", ref_not_empty)
    c5.metric("Excel файлове", len(get_excel_files()))

    b1, b2, b3 = st.columns([1.3, 1.3, 5])

    with b1:
        if st.button("🔄 Пресъздай Master"):
            new_master, source_log = build_master_database_from_excels()
            save_master_database(new_master)
            st.session_state.master_df = new_master

            st.success("Master Database е пресъздадена.")
            if source_log:
                st.caption("Източници: " + " | ".join(source_log))

            st.rerun()

    with b2:
        st.download_button(
            "⬇️ Свали Master CSV",
            data=master_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name="master_database.csv",
            mime="text/csv"
        )

    search_master = st.text_input(
        "Търсене в Master Database",
        placeholder="Вътрешен номер / активен номер / доставчик / vendor / цена..."
    )

    master_view = master_df.copy()

    if search_master.strip():
        query = search_master.strip().lower()

        mask = master_view.apply(
            lambda row: row.astype(str).str.lower().str.contains(query, na=False).any(),
            axis=1
        )

        master_view = master_view[mask]

    edited_master = st.data_editor(
        master_view,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="master_editor"
    )

    if search_master.strip():
        st.warning("Редактираш филтриран изглед. За пълна редакция изчисти търсенето.")
    else:
        old_hash = dataframe_hash(master_df)
        new_hash = dataframe_hash(edited_master)

        if old_hash != new_hash:
            cleaned = clean_dataframe_as_text(edited_master)

            if all(col in cleaned.columns for col in MASTER_COLUMNS):
                cleaned["_key_internal"] = cleaned["Вътрешен номер"].map(normalize_key)
                cleaned["_key_active"] = cleaned["Активен номер"].map(normalize_key)
                cleaned["_dedupe_key"] = cleaned["_key_internal"] + "||" + cleaned["_key_active"]

                cleaned = cleaned.drop_duplicates(
                    subset=["_dedupe_key"],
                    keep="first"
                )

                cleaned = cleaned.drop(
                    columns=["_key_internal", "_key_active", "_dedupe_key"],
                    errors="ignore"
                )

            save_master_database(cleaned)
            st.session_state.master_df = cleaned
            st.success("Master Database е записана автоматично.")


# ======================================================
# PAGE: NEW CLAIMS
# ======================================================

elif st.session_state.page == "New Claims":

    st.markdown('<div class="section-title">➕ New Claims / Upload нови разлики</div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="info-box">
        Качи Excel файл с нови разлики. Приложението взима:
        <br><br>
        <b>C</b> = Доставчик<br>
        <b>D</b> = Vendor No / Доставчик №<br>
        <b>E</b> = Номер прием<br>
        <b>F</b> = Вътрешен номер<br>
        <b>G</b> = Активен номер<br>
        <b>I</b> = Quantity, което засега влиза в <b>Difference</b>
        <br><br>
        След това търси по <b>Вътрешен номер</b> или <b>Активен номер</b> в Master Database.
        Ако намери артикул — налива цена и Ref. Ако не намери — слага <b>0</b>.
        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_claims_file = st.file_uploader(
        "Качи Excel файл с нови разлики",
        type=["xlsx"],
        key="new_claims_upload"
    )

    c1, c2, c3 = st.columns([1.2, 1.2, 5])

    with c1:
        preview_clicked = st.button("👁️ Прегледай файла")

    with c2:
        confirm_clicked = st.button("✅ Потвърди и добави")

    if uploaded_claims_file is not None and preview_clicked:
        parsed = read_new_claims_upload(uploaded_claims_file)

        if parsed.empty:
            st.error("Не са намерени валидни редове във файла.")
        else:
            filled = autofill_claims_from_master(
                parsed,
                st.session_state.master_df
            )

            st.session_state.preview_claims_df = filled.copy()

            st.success(f"Подготвени редове за добавяне: {len(filled)}")

    preview_df = st.session_state.preview_claims_df.copy()

    if not preview_df.empty:
        st.subheader("Преглед преди добавяне в основния списък")

        diff_values = safe_numeric(preview_df["Difference"])
        price_values = preview_df["Price"].astype(str).str.strip()
        ref_values = preview_df["Ref. Number SUPPLIER"].astype(str).str.strip()

        rows_with_price = int(price_values.ne("").sum())
        rows_price_zero = int(price_values.eq("0").sum())
        rows_with_ref = int(ref_values.ne("").sum())
        rows_ref_zero = int(ref_values.eq("0").sum())

        s1, s2, s3, s4, s5, s6 = st.columns(6)

        s1.metric("Редове", len(preview_df))
        s2.metric("Плюсове", int((diff_values > 0).sum()))
        s3.metric("Минуси", int((diff_values < 0).sum()))
        s4.metric("С цена", rows_with_price)
        s5.metric("Price = 0", rows_price_zero)
        s6.metric("Ref = 0", rows_ref_zero)

        edited_preview = st.data_editor(
            preview_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "СТАТУС - Попълва се от централата!": st.column_config.SelectboxColumn(
                    "СТАТУС - Попълва се от централата!",
                    options=STATUS_OPTIONS,
                    required=False,
                )
            },
            key="preview_claims_editor"
        )

        st.session_state.preview_claims_df = edited_preview.copy()

    if confirm_clicked:
        preview = st.session_state.preview_claims_df.copy()

        if preview.empty:
            st.error("Няма подготвени редове. Първо натисни 'Прегледай файла'.")
        else:
            current = load_differences_database()

            updated = pd.concat(
                [current, preview],
                ignore_index=True
            )

            updated = clean_dataframe_as_text(updated)
            updated = updated[DIFFERENCES_COLUMNS]

            save_differences_database(updated)

            st.session_state.differences_df = updated
            st.session_state.preview_claims_df = pd.DataFrame(columns=DIFFERENCES_COLUMNS)

            st.success(f"Добавени са {len(preview)} реда в основния списък Разлики.")
            st.info("Отвори меню 'Разлики', за да ги видиш в основната таблица.")


# ======================================================
# PAGE: ACCOUNTING
# ======================================================

elif st.session_state.page == "Счетоводство":

    st.markdown('<div class="section-title">💰 Счетоводство</div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="info-box">
        Тук търсиш по документ, номер прием, доставчик, вътрешен номер или активен номер.
        Резултатът може да се копира директно.
        </div>
        """,
        unsafe_allow_html=True
    )

    accounting_search = st.text_input(
        "Въведи номер / документ / прием / доставчик / артикул",
        placeholder="Например номер прием, вътрешен номер или доставчик..."
    )

    result_df = differences_df.copy()

    if accounting_search.strip():
        query = accounting_search.strip().lower()

        mask = result_df.apply(
            lambda row: row.astype(str).str.lower().str.contains(query, na=False).any(),
            axis=1
        )

        result_df = result_df[mask]
    else:
        result_df = result_df.iloc[0:0]

    st.metric("Намерени редове", len(result_df))

    st.dataframe(
        result_df,
        use_container_width=True,
        hide_index=True
    )

    if not result_df.empty:
        copy_text = to_tsv(result_df)

        st.text_area(
            "Готов текст за копиране",
            value=copy_text,
            height=230,
            key="accounting_copy_text"
        )

        st.download_button(
            "⬇️ Свали резултата CSV",
            data=result_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name="accounting_export.csv",
            mime="text/csv"
        )


# ======================================================
# PAGE: ADMIN
# ======================================================

elif st.session_state.page == "Админ":

    st.markdown('<div class="section-title">⚙️ Администрация</div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="info-box">
        Технически действия: проверка на файлове, ръчно добавяне, изчистване и пресъздаване на бази.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.subheader("Файлове в проекта")

    files_df = pd.DataFrame({"Файл": os.listdir(".")})
    st.dataframe(files_df, use_container_width=True, hide_index=True)

    st.subheader("Ръчно добавяне в Master Database")

    manual = st.data_editor(
        pd.DataFrame(columns=MASTER_COLUMNS),
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="manual_master_add"
    )

    if st.button("➕ Добави към Master Database"):
        if manual.empty:
            st.warning("Няма въведени редове.")
        else:
            current_master = load_master_database()

            updated_master = pd.concat(
                [current_master, manual],
                ignore_index=True
            )

            updated_master = clean_dataframe_as_text(updated_master)

            updated_master["_key_internal"] = updated_master["Вътрешен номер"].map(normalize_key)
            updated_master["_key_active"] = updated_master["Активен номер"].map(normalize_key)
            updated_master["_dedupe_key"] = updated_master["_key_internal"] + "||" + updated_master["_key_active"]

            updated_master = updated_master.drop_duplicates(
                subset=["_dedupe_key"],
                keep="first"
            )

            updated_master = updated_master.drop(
                columns=["_key_internal", "_key_active", "_dedupe_key"],
                errors="ignore"
            )

            save_master_database(updated_master)
            st.session_state.master_df = updated_master

            st.success("Ръчно добавените редове са записани в Master Database.")

    st.divider()

    danger = st.checkbox("Покажи опасни действия")

    if danger:

        c1, c2, c3 = st.columns(3)

        with c1:
            if st.button("🧹 Изчисти основния списък Разлики"):
                empty = pd.DataFrame(columns=DIFFERENCES_COLUMNS)

                save_differences_database(empty)
                st.session_state.differences_df = empty

                st.success("Основният списък е изчистен.")

        with c2:
            if st.button("🧹 Изтрий Master CSV"):
                if os.path.exists(MASTER_DB_FILE):
                    os.remove(MASTER_DB_FILE)

                st.session_state.master_df = pd.DataFrame(columns=MASTER_COLUMNS)

                st.success("Master CSV е изтрит.")

        with c3:
            if st.button("🔄 Изтрий и пресъздай Master"):
                if os.path.exists(MASTER_DB_FILE):
                    os.remove(MASTER_DB_FILE)

                new_master, source_log = build_master_database_from_excels()

                save_master_database(new_master)
                st.session_state.master_df = new_master

                st.success("Master Database е пресъздадена.")
                if source_log:
                    st.caption("Източници: " + " | ".join(source_log))

                st.rerun()


# ======================================================
# FOOTER
# ======================================================

st.markdown("---")
st.markdown(
    """
    <div class="small-muted">
    Differences Portal Inter Cars · Professional v3 · Master Database + Upload New Claims + Accounting Export
    </div>
    """,
    unsafe_allow_html=True
)
