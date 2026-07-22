import os
import re
import io
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

MASTER_REQUIRED_COLUMNS = [
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
    "Diff Status",
    "Стойност тотал от Нави",
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
    "Чака отговор",
    "Получено КИ",
    "Отказани",
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
    padding-left: 1.8rem;
    padding-right: 1.8rem;
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
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    font-size: 18px;
    background: #fff;
}

.title-box {
    border: 4px solid #111;
    height: 86px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 900;
    font-size: 26px;
    letter-spacing: 1px;
    background: #fff;
}

.menu-wrapper {
    border: 4px solid #111;
    padding: 18px;
    margin-bottom: 22px;
    background: #fff;
}

.stButton > button {
    width: 100%;
    min-height: 60px;
    border: 3px solid #111;
    border-radius: 0px;
    background: #ffffff;
    color: #111111;
    font-weight: 900;
    font-size: 16px;
}

.stButton > button:hover {
    border: 3px solid #ff4b4b;
    color: #ff4b4b;
    background: #fff7f7;
}

.info-box {
    border-left: 6px solid #111;
    background: #f8fafc;
    padding: 14px 18px;
    margin-bottom: 16px;
}

.section-title {
    font-size: 26px;
    font-weight: 900;
    margin-top: 10px;
    margin-bottom: 10px;
}

.small-muted {
    color: #6b7280;
    font-size: 13px;
}

hr {
    margin-top: 12px;
    margin-bottom: 16px;
}
</style>
""",
    unsafe_allow_html=True
)


# ======================================================
# HELPERS
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


def normalize_key(value):
    value = clean_text(value)
    value = value.upper()
    value = value.replace(" ", "")
    value = value.replace(".", "")
    value = value.replace("-", "")
    value = value.replace("/", "")
    return value


def normalize_columns(df):
    df = df.copy()
    df.columns = [clean_text(c) for c in df.columns]
    return df


def clean_dataframe_as_text(df):
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].map(clean_text)
    return df


def get_excel_files():
    files = []
    for file in os.listdir("."):
        if file.startswith("~$"):
            continue
        if file.lower().endswith(".xlsx"):
            files.append(file)
    return files


def find_sheet_with_columns(excel_file, required_columns):
    try:
        xls = pd.ExcelFile(excel_file, engine="openpyxl")
    except Exception:
        return None, None

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(
                excel_file,
                sheet_name=sheet,
                dtype=object,
                engine="openpyxl"
            )

            df = normalize_columns(df)

            if all(col in df.columns for col in required_columns):
                return sheet, df

        except Exception:
            continue

    return None, None


def build_master_database_from_excels():
    all_master_rows = []
    source_log = []

    excel_files = get_excel_files()

    for file in excel_files:
        sheet, df = find_sheet_with_columns(file, MASTER_REQUIRED_COLUMNS)

        if df is None:
            continue

        df = df[MASTER_REQUIRED_COLUMNS].copy()
        df = clean_dataframe_as_text(df)

        df = df[
            (df["Вътрешен номер"].astype(str).str.strip() != "") |
            (df["Активен номер"].astype(str).str.strip() != "")
        ]

        df["source_file"] = file
        df["source_sheet"] = sheet

        all_master_rows.append(df)
        source_log.append(f"{file} / {sheet}")

    if not all_master_rows:
        empty = pd.DataFrame(columns=MASTER_REQUIRED_COLUMNS + ["source_file", "source_sheet"])
        return empty, source_log

    master = pd.concat(all_master_rows, ignore_index=True)
    master = clean_dataframe_as_text(master)

    master["_key_internal"] = master["Вътрешен номер"].map(normalize_key)
    master["_key_active"] = master["Активен номер"].map(normalize_key)
    master["_dedupe_key"] = master["_key_internal"] + "||" + master["_key_active"]

    master = master.drop_duplicates(subset=["_dedupe_key"], keep="first")

    master = master.drop(columns=["_key_internal", "_key_active", "_dedupe_key"])
    master = master.reset_index(drop=True)

    return master, source_log


def save_master_database(df):
    df = df.copy()

    for col in MASTER_REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    ordered_cols = MASTER_REQUIRED_COLUMNS + [c for c in df.columns if c not in MASTER_REQUIRED_COLUMNS]
    df = df[ordered_cols]
    df = clean_dataframe_as_text(df)
    df.to_csv(MASTER_DB_FILE, index=False, encoding="utf-8-sig")


def load_master_database():
    if os.path.exists(MASTER_DB_FILE):
        df = pd.read_csv(MASTER_DB_FILE, dtype=str, keep_default_na=False)
        df = clean_dataframe_as_text(df)

        for col in MASTER_REQUIRED_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        return df

    master, _ = build_master_database_from_excels()
    save_master_database(master)
    return master


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
        df = pd.read_csv(DIFFERENCES_DB_FILE, dtype=str, keep_default_na=False)
        df = clean_dataframe_as_text(df)

        for col in DIFFERENCES_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        return df[DIFFERENCES_COLUMNS]

    df = pd.DataFrame(columns=DIFFERENCES_COLUMNS)
    save_differences_database(df)
    return df


def dataframe_hash(df):
    if df is None or df.empty:
        return "empty"

    raw = df.to_csv(index=False).encode("utf-8", errors="ignore")
    return hashlib.md5(raw).hexdigest()


def safe_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0)


def make_master_lookup(master_df):
    lookup_internal = {}
    lookup_active = {}

    if master_df.empty:
        return lookup_internal, lookup_active

    for _, row in master_df.iterrows():
        internal = normalize_key(row.get("Вътрешен номер", ""))
        active = normalize_key(row.get("Активен номер", ""))

        record = {
            "Доставчик": clean_text(row.get("Доставчик", "")),
            "Vendor No": clean_text(row.get("Vendor No", "")),
            "Вътрешен номер": clean_text(row.get("Вътрешен номер", "")),
            "Активен номер": clean_text(row.get("Активен номер", "")),
            "Ref. Number SUPPLIER": clean_text(row.get("Ref. Number SUPPLIER", "")),
            "Price": clean_text(row.get("Price", "")),
        }

        if internal and internal not in lookup_internal:
            lookup_internal[internal] = record

        if active and active not in lookup_active:
            lookup_active[active] = record

    return lookup_internal, lookup_active


def autofill_claims_from_master(claims_df, master_df):
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
                "Ref. Number SUPPLIER",
                "Price",
            ]:
                if clean_text(row.get(field, "")) == "":
                    row[field] = match.get(field, "")

            # Важно: Ref. Number SUPPLIER и Price винаги ги обновяваме от Master,
            # защото това е точната база за приложението.
            row["Ref. Number SUPPLIER"] = match.get("Ref. Number SUPPLIER", "")
            row["Price"] = match.get("Price", "")

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


def read_new_claims_upload(uploaded_file):
    """
    Чете директно два листа от upload файла:

    1) Sheet1
       - съдържа основните данни от Нави
       - Quantity = QTY по Нави

    2) Разлики
       - съдържа реалната разлика
       - Quantity = Difference

    Логика:
    QTY = Sheet1 Quantity
    Difference = Разлики Quantity
    Received QTY = QTY + Difference

    Начин на подаване:
    Взима се от Sharepoint_Разлики_2026.xlsx по доставчик.
    Ако няма намерена информация - остава празно.
    """

    try:
        sheet1 = pd.read_excel(
            uploaded_file,
            sheet_name="Sheet1",
            dtype=object,
            engine="openpyxl"
        )

        diff_sheet = pd.read_excel(
            uploaded_file,
            sheet_name="Разлики",
            dtype=object,
            engine="openpyxl"
        )

    except Exception as e:
        st.error(f"Грешка при четене на файла: {e}")
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    if sheet1.empty:
        st.error("Sheet1 е празен.")
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    if diff_sheet.empty:
        st.error("Sheet 'Разлики' е празен.")
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    sheet1 = normalize_columns(sheet1)
    sheet1 = clean_dataframe_as_text(sheet1)

    diff_sheet = normalize_columns(diff_sheet)
    diff_sheet = clean_dataframe_as_text(diff_sheet)

    # ==================================================
    # SHAREPOINT LOOKUP - НАЧИН НА ПОДАВАНЕ
    # ==================================================

    submit_lookup = {}

    try:
        sharepoint_df = pd.read_excel(
            "Sharepoint_Разлики_2026.xlsx",
            dtype=str,
            engine="openpyxl"
        ).fillna("")

        sharepoint_df = normalize_columns(sharepoint_df)

        if (
            "Name" in sharepoint_df.columns
            and "линк към сайт или спец.бланка" in sharepoint_df.columns
        ):
            submit_lookup = dict(
                zip(
                    sharepoint_df["Name"].astype(str).str.strip().str.upper(),
                    sharepoint_df["линк към сайт или спец.бланка"].astype(str)
                )
            )

    except Exception:
        submit_lookup = {}

    required_sheet1_columns = [
        "Location Code",
        "Vendor Name",
        "Buy-from Vendor No_",
        "Receipt No",
        "Item No",
        "Active No",
        "Delivery No",
        "Invoice No",
        "InvoiceDate",
        "Supplier Ref Num",
        "Price per Invoice",
        "Quantity",
    ]

    required_diff_columns = [
        "Вътрешен №",
        "Активен №",
        "Quantity",
    ]

    missing_sheet1 = [
        col for col in required_sheet1_columns
        if col not in sheet1.columns
    ]

    missing_diff = [
        col for col in required_diff_columns
        if col not in diff_sheet.columns
    ]

    if missing_sheet1:
        st.error(
            "В Sheet1 липсват колони: "
            + ", ".join(missing_sheet1)
        )
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    if missing_diff:
        st.error(
            "В sheet 'Разлики' липсват колони: "
            + ", ".join(missing_diff)
        )
        return pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    # ==================================================
    # LOOKUP ЗА DIFFERENCE ОТ SHEET "Разлики"
    # ==================================================

    diff_lookup_full = {}
    diff_lookup_internal = {}

    for _, diff_row in diff_sheet.iterrows():

        diff_internal = normalize_key(
            diff_row.get("Вътрешен №", "")
        )

        diff_active = normalize_key(
            diff_row.get("Активен №", "")
        )

        diff_qty = clean_text(
            diff_row.get("Quantity", "")
        )

        full_key = diff_internal + "||" + diff_active

        if diff_internal or diff_active:
            diff_lookup_full[full_key] = diff_qty

        if diff_internal:
            diff_lookup_internal[diff_internal] = diff_qty

    def get_difference_for_row(row):

        internal_key = normalize_key(
            row.get("Item No", "")
        )

        active_key = normalize_key(
            row.get("Active No", "")
        )

        full_key = internal_key + "||" + active_key

        if full_key in diff_lookup_full:
            return diff_lookup_full[full_key]

        if internal_key in diff_lookup_internal:
            return diff_lookup_internal[internal_key]

        return "0"

    # ==================================================
    # ОСНОВНА ТАБЛИЦА
    # ==================================================

    result = pd.DataFrame(columns=DIFFERENCES_COLUMNS)

    result["Склад за дост."] = sheet1["Location Code"]

    result["Доставчик"] = sheet1["Vendor Name"]

    result["Vendor No"] = sheet1["Buy-from Vendor No_"]

    # Начин на подаване от Sharepoint_Разлики_2026.xlsx
    # Ако няма информация за доставчика - остава празно
    result["Начин на подаване"] = result["Доставчик"].apply(
        lambda x: submit_lookup.get(
            str(x).strip().upper(),
            ""
        )
    )

    result["Номер прием"] = sheet1["Receipt No"]

    result["Вътрешен номер"] = sheet1["Item No"]

    result["Активен номер"] = sheet1["Active No"]

    result["Delivery No"] = sheet1["Delivery No"]

    result["Invoice No"] = sheet1["Invoice No"]

    result["Invoice Date"] = sheet1["InvoiceDate"]

    result["Ref. Number SUPPLIER"] = sheet1["Supplier Ref Num"]

    result["Price"] = sheet1["Price per Invoice"]

    # QTY идва от Sheet1 Quantity
    result["QTY"] = sheet1["Quantity"]

    # Difference идва от лист Разлики -> Quantity
    result["Difference"] = sheet1.apply(
        get_difference_for_row,
        axis=1
    )

    # Визуална колона за плюс/минус
    result["Diff Status"] = result["Difference"].apply(
        lambda x: "🔴 минус" if str(x).strip().startswith("-") else "🟢 плюс"
    )

    # ==================================================
    # Received QTY = QTY + Difference
    # ==================================================

    qty_num = pd.to_numeric(
        result["QTY"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0)

    diff_num = pd.to_numeric(
        result["Difference"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0)

    result["Received QTY"] = (
        qty_num + diff_num
    )

    # ==================================================
# Стойност тотал от Нави = Price × ABS(Difference)
# ==================================================

price_num = pd.to_numeric(
    result["Price"].astype(str).str.replace(",", ".", regex=False),
    errors="coerce"
).fillna(0)

difference_abs = pd.to_numeric(
    result["Difference"].astype(str).str.replace(",", ".", regex=False),
    errors="coerce"
).fillna(0).abs()

result["Стойност тотал от Нави"] = (
    price_num * difference_abs
).round(2)

    result["Дата на подаване"] = datetime.now().strftime("%d.%m.%Y")

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
# ======================================================
# STATE
# ======================================================

if "page" not in st.session_state:
    st.session_state.page = "Разлики"

if "differences_df" not in st.session_state:
    st.session_state.differences_df = load_differences_database()

if "preview_claims_df" not in st.session_state:
    st.session_state.preview_claims_df = pd.DataFrame(
        columns=DIFFERENCES_COLUMNS
    )

# UNDO / REDO

if "undo_stack" not in st.session_state:
    st.session_state.undo_stack = []

if "redo_stack" not in st.session_state:
    st.session_state.redo_stack = []


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
# TOP MENU
# ======================================================

st.markdown('<div class="menu-wrapper">', unsafe_allow_html=True)

m1, m3, m4, m5 = st.columns(4)

with m1:
    if st.button("📋 Разлики"):
        st.session_state.page = "Разлики"

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
# GLOBAL DATA
# ======================================================

differences_df = st.session_state.differences_df.copy()


# ======================================================
# PAGE: DIFFERENCES
# ======================================================

if st.session_state.page == "Разлики":

    st.markdown(
        '<div class="section-title">📋 Основен списък с разлики</div>',
        unsafe_allow_html=True
    )

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
            .isin(["", "нова", "подадена", "чака отговор", "обработва се"])
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
        Основният прозорец пази колоните в същия ред като структурата.
        Можеш да редактираш директно таблицата.
        Включени са Undo / Redo.
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

        search_value = search.strip().lower()

        mask = view_df.apply(
            lambda row: row.astype(str)
            .str.lower()
            .str.contains(search_value, na=False)
            .any(),
            axis=1
        )

        view_df = view_df[mask]

    u1, u2, u3 = st.columns([1, 1, 8])

    with u1:
        if st.button("⬅️ Undo"):

            if st.session_state.undo_stack:

                st.session_state.redo_stack.append(
                    st.session_state.differences_df.copy()
                )

                st.session_state.differences_df = (
                    st.session_state.undo_stack.pop()
                )

                save_differences_database(
                    st.session_state.differences_df
                )

                st.rerun()

    with u2:
        if st.button("➡️ Redo"):

            if st.session_state.redo_stack:

                st.session_state.undo_stack.append(
                    st.session_state.differences_df.copy()
                )

                st.session_state.differences_df = (
                    st.session_state.redo_stack.pop()
                )

                save_differences_database(
                    st.session_state.differences_df
                )

                st.rerun()

    edited_df = st.data_editor(
        view_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "СТАТУС - Попълва се от централата!":
            st.column_config.SelectboxColumn(
                "СТАТУС - Попълва се от централата!",
                options=STATUS_OPTIONS,
                required=False,
            )
        },
        key="differences_editor"
    )

    if search.strip():

        st.warning(
            "Редактираш филтриран изглед. За реално записване изчисти търсенето."
        )

    else:

        old_hash = dataframe_hash(differences_df)
        new_hash = dataframe_hash(edited_df)

        if old_hash != new_hash:

            st.session_state.undo_stack.append(
                differences_df.copy()
            )

            st.session_state.redo_stack = []

            st.session_state.differences_df = edited_df.copy()

            save_differences_database(
                st.session_state.differences_df
            )

            st.success("Промените са записани автоматично.")

    c1, c2, c3 = st.columns([1, 1, 5])

    with c1:
        st.download_button(
            "⬇️ Export CSV",
            data=st.session_state.differences_df.to_csv(
                index=False,
                encoding="utf-8-sig"
            ),
            file_name="differences_database.csv",
            mime="text/csv"
        )

    with c2:
        if st.button("🔄 Презареди"):
            st.session_state.differences_df = load_differences_database()
            st.rerun()
# ======================================================
# PAGE: NEW CLAIMS
# ======================================================

elif st.session_state.page == "New Claims":

    st.markdown('<div class="section-title">➕ New Claims / Upload нови разлики</div>', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="info-box">
        Тук качваш Excel файл със структура като файла <b>разлики.xlsx</b>.
        Приложението взима автоматично:
        <br><br>
        <b>C</b> = Доставчик<br>
        <b>D</b> = Vendor No / Доставчик №<br>
        <b>E</b> = Номер прием<br>
        <b>F</b> = Вътрешен номер<br>
        <b>G</b> = Активен номер<br>
        <b>I</b> = Quantity, което засега влиза в <b>Difference</b>
        <br><br>
        След това търси по <b>Вътрешен номер</b> или <b>Активен номер</b> в Master Database
        и попълва автоматично <b>Ref. Number SUPPLIER</b> и <b>Price</b>.
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
            st.session_state.preview_claims_df = parsed.copy()
            st.success(
                f"Подготвени редове за добавяне: {len(parsed)}")   

    preview_df = st.session_state.preview_claims_df.copy()

    if not preview_df.empty:
        st.subheader("Преглед преди добавяне в основния списък")

        st.caption(
            "Тук можеш да направиш последна проверка. След потвърждение редовете ще се добавят в основния списък Разлики."
        )

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

        diff_values = safe_numeric(edited_preview["Difference"]) if "Difference" in edited_preview.columns else pd.Series(dtype=float)

        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Редове за добавяне", len(edited_preview))
        s2.metric("Плюсове", int((diff_values > 0).sum()))
        s3.metric("Минуси", int((diff_values < 0).sum()))
        s4.metric("Нула / празни", int((diff_values == 0).sum()))

    if confirm_clicked:
        preview = st.session_state.preview_claims_df.copy()

        if preview.empty:
            st.error("Няма подготвени редове. Първо натисни 'Прегледай файла'.")
        else:
            current = load_differences_database()

            updated = pd.concat([current, preview], ignore_index=True)
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
        Тук търсиш по номер документ за разлики, номер прием, доставчик или артикул.
        Резултатът можеш да копираш директно към счетоводство.
        </div>
        """,
        unsafe_allow_html=True
    )

    accounting_search = st.text_input(
        "Въведи номер / документ / прием / доставчик / артикул",
        placeholder="Например номер прием или вътрешен номер..."
    )

    result_df = differences_df.copy()

    if accounting_search.strip():
        search_value = accounting_search.strip().lower()
        mask = result_df.apply(
            lambda row: row.astype(str).str.lower().str.contains(search_value, na=False).any(),
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
            "Готов текст за копиране към счетоводство",
            value=copy_text,
            height=240,
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
        Тук са техническите действия: проверка на файлове, бази, изчистване и ръчно добавяне.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.subheader("Файлове в проекта")

    files_df = pd.DataFrame({"Файл": os.listdir(".")})
    st.dataframe(files_df, use_container_width=True, hide_index=True)

    st.subheader("Ръчно добавяне в Master Database")

    manual = st.data_editor(
        pd.DataFrame(columns=MASTER_REQUIRED_COLUMNS),
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
            updated_master = pd.concat([current_master, manual], ignore_index=True)
            updated_master = clean_dataframe_as_text(updated_master)

            updated_master["_key_internal"] = updated_master["Вътрешен номер"].map(normalize_key)
            updated_master["_key_active"] = updated_master["Активен номер"].map(normalize_key)
            updated_master["_dedupe_key"] = updated_master["_key_internal"] + "||" + updated_master["_key_active"]
            updated_master = updated_master.drop_duplicates(subset=["_dedupe_key"], keep="first")
            updated_master = updated_master.drop(columns=["_key_internal", "_key_active", "_dedupe_key"])

            save_master_database(updated_master)
            st.session_state.master_df = updated_master
            st.success("Ръчно добавените редове са записани в Master Database.")

    st.divider()

    danger = st.checkbox("Покажи опасни действия")

    if danger:
        c1, c2 = st.columns(2)

        with c1:
            if st.button("🧹 Изчисти основния списък Разлики"):
                empty = pd.DataFrame(columns=DIFFERENCES_COLUMNS)
                save_differences_database(empty)
                st.session_state.differences_df = empty
                st.success("Основният списък е изчистен.")

        with c2:
            if st.button("🧹 Изтрий Master Database и пресъздай"):
                if os.path.exists(MASTER_DB_FILE):
                    os.remove(MASTER_DB_FILE)

                new_master, source_log = build_master_database_from_excels()
                save_master_database(new_master)
                st.session_state.master_df = new_master

                st.success("Master Database е пресъздадена.")
                if source_log:
                    st.caption("Източници: " + " | ".join(source_log))


# ======================================================
# FOOTER
# ======================================================

st.markdown("---")
st.markdown(
    """
    <div class="small-muted">
    Differences Portal Inter Cars · v2 · Master Database + Upload New Claims + Accounting Export
    </div>
    """,
    unsafe_allow_html=True
)
