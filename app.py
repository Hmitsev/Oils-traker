import streamlit as st
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import date, datetime

# ======================================================
# ✅ CONFIG
# ======================================================

st.set_page_config(
    page_title="Differences Suppliers",
    page_icon="📦",
    layout="wide"
)

EXCEL_PATH = Path("Sharepoint_Разлики_2026.xlsx")
DB_PATH = Path("differences.sqlite")
SHEET_NAME = "Разлики"

# ======================================================
# ✅ COLUMNS - EXACT ORDER
# ======================================================

COLUMNS = [
    "Delivery No",
    "Invoice No",
    "Invoice Date",
    "Ref. Number SUPPLIER",
    "Other Supplier Ref Num",
    "Price",
    "QTY",
    "Received QTY",
    "Difference",
    "Подал разликата",
    "Стойност във валутата на доставчика",
    "Стойност (в лева)",
    "Дата на подаване",
    "СТАТУС - Попълва се от централата!",
    "№ документа за разлики",
    "Дата на документа за разлики",
    "При фактура - дата на приема в Навижън",
    "Обработени в B01",
    "Номер на клетка за минуси",
    "Намерени БРОЙКИ (след подаването им)",
    "Подадена информация към дост. за намерени бройки след подаването им (ДАТА)",
    "Допълнителен коментар",
    "ДАТА на намиране в В01 на липсващи артикули или коментар от склада",
    "дата на прием в Навижън",
    "working days after goods receipt in navision",
    "БРАНД",
    "РМ",
    "формула Ники",
    "дата на прием от таблицата на Тони",
    "номер на склада - програмата",
    "системен номер на доставка на склада"
]

# ======================================================
# ✅ STYLE - NEW DESIGN
# ======================================================

st.markdown(
    """
    <style>
        .stApp {
            background: linear-gradient(135deg, #f5f7fb 0%, #eef2f8 45%, #f8fafc 100%);
        }

        .main-title {
            font-size: 34px;
            font-weight: 800;
            color: #1f2937;
            margin-bottom: 0px;
        }

        .subtitle {
            font-size: 15px;
            color: #6b7280;
            margin-bottom: 25px;
        }

        .metric-card {
            background: white;
            padding: 22px;
            border-radius: 18px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
            border: 1px solid rgba(226, 232, 240, 0.9);
        }

        .metric-label {
            font-size: 13px;
            color: #6b7280;
            margin-bottom: 6px;
        }

        .metric-value {
            font-size: 28px;
            color: #111827;
            font-weight: 800;
        }

        .section-card {
            background: white;
            padding: 24px;
            border-radius: 20px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.07);
            border: 1px solid rgba(226, 232, 240, 0.9);
            margin-bottom: 18px;
        }

        .success-box {
            background: #ecfdf5;
            color: #065f46;
            padding: 14px 18px;
            border-radius: 14px;
            border: 1px solid #a7f3d0;
            font-weight: 600;
        }

        .warning-box {
            background: #fffbeb;
            color: #92400e;
            padding: 14px 18px;
            border-radius: 14px;
            border: 1px solid #fde68a;
            font-weight: 600;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        }

        section[data-testid="stSidebar"] {
            background: #111827;
        }

        section[data-testid="stSidebar"] * {
            color: white;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# ======================================================
# ✅ HELPERS
# ======================================================

def normalize_number(value):
    if value is None:
        return 0.0

    if pd.isna(value):
        return 0.0

    if isinstance(value, str):
        value = (
            value
            .replace("лв.", "")
            .replace(" ", "")
            .replace(",", ".")
            .strip()
        )

    if value == "":
        return 0.0

    try:
        return float(value)
    except Exception:
        return 0.0


def calculate_difference(qty, received_qty):
    qty_num = normalize_number(qty)
    received_num = normalize_number(received_qty)
    return received_num - qty_num


def calculate_supplier_value(price, difference):
    price_num = normalize_number(price)
    diff_num = normalize_number(difference)
    return round(price_num * diff_num, 2)


def make_niki_formula(invoice_no, ref_number):
    invoice_clean = (
        str(invoice_no)
        .strip()
        .replace("/", "")
        .replace("-", "")
        .replace(" ", "")
    )

    ref_clean = (
        str(ref_number)
        .strip()
        .replace("/", "")
        .replace("-", "")
        .replace(" ", "")
    )

    if invoice_clean and invoice_clean.lower() not in ["none", "nan"] and ref_clean and ref_clean.lower() not in ["none", "nan"]:
        return f"{invoice_clean}|{ref_clean}"

    return ""


def prepare_dataframe(df):
    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNS]
    df = df.fillna("")

    return df


def remove_empty_rows(df):
    df = df.copy()

    df = df[
        df.astype(str).apply(
            lambda row: any(
                str(x).strip().lower() not in ["", "none", "nan", "nat"]
                for x in row
            ),
            axis=1
        )
    ]

    return df


def apply_auto_calculations(df):
    df = df.copy()

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNS]

    # ✅ Difference = Received QTY - QTY
    df["Difference"] = df.apply(
        lambda row: calculate_difference(
            row.get("QTY", 0),
            row.get("Received QTY", 0)
        ),
        axis=1
    )

    # ✅ Стойност във валутата = Price * Difference
    df["Стойност във валутата на доставчика"] = df.apply(
        lambda row: calculate_supplier_value(
            row.get("Price", 0),
            row.get("Difference", 0)
        ),
        axis=1
    )

    # ✅ Дата на подаване - ако е празна
    df["Дата на подаване"] = df["Дата на подаване"].apply(
        lambda x: date.today().strftime("%Y-%m-%d")
        if pd.isna(x) or str(x).strip().lower() in ["", "none", "nan", "nat"]
        else x
    )

    # ✅ Формула Ники
    df["формула Ники"] = df.apply(
        lambda row: make_niki_formula(
            row.get("Invoice No", ""),
            row.get("Ref. Number SUPPLIER", "")
        ),
        axis=1
    )

    return df


def hash_dataframe(df):
    try:
        df_string = df.astype(str)
        return str(pd.util.hash_pandas_object(df_string, index=True).sum())
    except Exception:
        return str(datetime.now())


# ======================================================
# ✅ SQLITE DATABASE
# ======================================================

def get_connection():
    return sqlite3.connect(DB_PATH)


def save_db(df):
    df = prepare_dataframe(df)
    df = remove_empty_rows(df)

    conn = get_connection()
    df.to_sql("differences", conn, if_exists="replace", index=False)
    conn.close()


@st.cache_data(show_spinner=False)
def load_data():
    try:
        # ✅ Ако вече има SQLite база - чете от нея
        if DB_PATH.exists():
            conn = get_connection()
            df = pd.read_sql_query("SELECT * FROM differences", conn)
            conn.close()

            df = prepare_dataframe(df)
            return df

        # ✅ Ако няма SQLite база, но има Excel - мигрира Excel към SQLite
        if EXCEL_PATH.exists():
            try:
                df = pd.read_excel(
                    EXCEL_PATH,
                    sheet_name=SHEET_NAME,
                    engine="openpyxl"
                )
            except Exception:
                df = pd.read_excel(
                    EXCEL_PATH,
                    engine="openpyxl"
                )

            df = prepare_dataframe(df)
            df = remove_empty_rows(df)
            save_db(df)

            return df

        # ✅ Ако няма нищо
        return pd.DataFrame(columns=COLUMNS)

    except Exception as e:
        st.error(f"Грешка при зареждане на данните: {e}")
        return pd.DataFrame(columns=COLUMNS)


def refresh_data():
    st.cache_data.clear()


# ======================================================
# ✅ LOAD DATA
# ======================================================

df = load_data()

# ======================================================
# ✅ SIDEBAR
# ======================================================

with st.sidebar:
    st.markdown("## 📦 Differences")
    st.markdown("### Suppliers")
    st.divider()

    page = st.radio(
        "Меню",
        [
            "📊 Dashboard",
            "📋 Всички разлики",
            "📝 Въвеждане / редакция",
            "🔎 Проверка / търсене"
        ],
        label_visibility="collapsed"
    )

    st.divider()

    st.caption("База данни:")
    st.write(str(DB_PATH))

    if DB_PATH.exists():
        st.success("SQLite базата е активна")
    elif EXCEL_PATH.exists():
        st.info("Excel файлът е намерен - ще бъде мигриран")
    else:
        st.warning("Няма намерена база или Excel файл")

# ======================================================
# ✅ HEADER
# ======================================================

st.markdown(
    '<div class="main-title">Differences Suppliers</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Приложение за подаване и проследяване на разлики към доставчици</div>',
    unsafe_allow_html=True
)

# ======================================================
# ✅ DASHBOARD
# ======================================================

if page == "📊 Dashboard":
    total_rows = len(df)

    plus_count = 0
    minus_count = 0
    open_count = 0
    total_bgn = 0

    if not df.empty:
        plus_count = (
            df["Подал разликата"]
            .astype(str)
            .str.lower()
            .eq("плюс")
            .sum()
        )

        minus_count = (
            df["Подал разликата"]
            .astype(str)
            .str.lower()
            .eq("минус")
            .sum()
        )

        status_col = "СТАТУС - Попълва се от централата!"

        open_count = (
            df[status_col]
            .astype(str)
            .isin(["", "Нова", "Подадена", "Очаква отговор"])
            .sum()
        )

        total_bgn = pd.to_numeric(
            df["Стойност (в лева)"]
            .astype(str)
            .str.replace("лв.", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False),
            errors="coerce"
        ).fillna(0).sum()

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Общо записи</div>
                <div class="metric-value">{total_rows}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Плюсове</div>
                <div class="metric-value">{plus_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Минуси</div>
                <div class="metric-value">{minus_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Отворени</div>
                <div class="metric-value">{open_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.write("")
    st.markdown("### Последни 20 записа")

    if df.empty:
        st.info("Все още няма данни.")
    else:
        st.dataframe(
            df.tail(20),
            use_container_width=True,
            hide_index=True
        )

# ======================================================
# ✅ ALL DIFFERENCES
# ======================================================

elif page == "📋 Всички разлики":
    st.markdown('<div class="section-card">', unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)

    with f1:
        search_text = st.text_input(
            "Търсене",
            placeholder="артикул, фактура, доставка..."
        )

    with f2:
        status_filter = st.selectbox(
            "Статус",
            ["Всички"] + sorted(
                [
                    x for x in df["СТАТУС - Попълва се от централата!"]
                    .dropna()
                    .astype(str)
                    .unique()
                    if x.strip()
                ]
            ) if not df.empty else ["Всички"]
        )

    with f3:
        brand_filter = st.selectbox(
            "Бранд",
            ["Всички"] + sorted(
                [
                    x for x in df["БРАНД"]
                    .dropna()
                    .astype(str)
                    .unique()
                    if x.strip()
                ]
            ) if not df.empty else ["Всички"]
        )

    with f4:
        rm_filter = st.selectbox(
            "РМ",
            ["Всички"] + sorted(
                [
                    x for x in df["РМ"]
                    .dropna()
                    .astype(str)
                    .unique()
                    if x.strip()
                ]
            ) if not df.empty else ["Всички"]
        )

    filtered_df = df.copy()

    if search_text:
        mask = filtered_df.astype(str).apply(
            lambda row: row.str.contains(
                search_text,
                case=False,
                na=False
            ).any(),
            axis=1
        )
        filtered_df = filtered_df[mask]

    if status_filter != "Всички":
        filtered_df = filtered_df[
            filtered_df["СТАТУС - Попълва се от централата!"].astype(str) == status_filter
        ]

    if brand_filter != "Всички":
        filtered_df = filtered_df[
            filtered_df["БРАНД"].astype(str) == brand_filter
        ]

    if rm_filter != "Всички":
        filtered_df = filtered_df[
            filtered_df["РМ"].astype(str) == rm_filter
        ]

    st.markdown(f"**Показани редове:** {len(filtered_df)} / {len(df)}")

    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        height=620
    )

    st.markdown('</div>', unsafe_allow_html=True)

# ======================================================
# ✅ EXCEL-LIKE EDITOR WITH ADD / DELETE / UNDO / REDO + AUTOSAVE
# ======================================================

elif page == "📝 Въвеждане / редакция":
    st.markdown("### 📝 Въвеждане / редакция на разлики")
    st.caption(
        "Работа като в Excel - добавяш ред най-отдолу, триеш редове, редактираш клетки и можеш да върнеш промени с Назад / Напред."
    )

    if df.empty:
        editable_df = pd.DataFrame(columns=COLUMNS)
    else:
        editable_df = df.copy()

    editable_df = prepare_dataframe(editable_df)

    # ======================================================
    # ✅ INIT UNDO / REDO STATE
    # ======================================================

    if "undo_stack" not in st.session_state:
        st.session_state["undo_stack"] = []

    if "redo_stack" not in st.session_state:
        st.session_state["redo_stack"] = []

    if "last_saved_hash" not in st.session_state:
        st.session_state["last_saved_hash"] = hash_dataframe(editable_df)

    if "last_saved_df" not in st.session_state:
        st.session_state["last_saved_df"] = editable_df.copy()

    if "editor_version" not in st.session_state:
        st.session_state["editor_version"] = 0

    st.markdown(
        """
        <div class="success-box">
            ✅ Таблицата работи като Excel: добавяне на редове, изтриване на редове, редакция на клетки, Copy/Paste и автоматичен запис.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")

    # ======================================================
    # ✅ UNDO / REDO BUTTONS
    # ======================================================

    b1, b2, b3, b4 = st.columns([1, 1, 2, 4])

    with b1:
        undo_clicked = st.button(
            "⬅️ Назад",
            disabled=len(st.session_state["undo_stack"]) == 0,
            use_container_width=True
        )

    with b2:
        redo_clicked = st.button(
            "➡️ Напред",
            disabled=len(st.session_state["redo_stack"]) == 0,
            use_container_width=True
        )

    with b3:
        st.caption(
            f"Undo: {len(st.session_state['undo_stack'])} | Redo: {len(st.session_state['redo_stack'])}"
        )

    # ======================================================
    # ✅ UNDO ACTION
    # ======================================================

    if undo_clicked:
        if len(st.session_state["undo_stack"]) > 0:
            current_df = st.session_state["last_saved_df"].copy()
            previous_df = st.session_state["undo_stack"].pop()

            st.session_state["redo_stack"].append(current_df)

            previous_df = prepare_dataframe(previous_df)
            save_db(previous_df)

            refresh_data()

            st.session_state["last_saved_df"] = previous_df.copy()
            st.session_state["last_saved_hash"] = hash_dataframe(previous_df)
            st.session_state["editor_version"] += 1

            st.toast("⬅️ Върната е последната промяна", icon="⬅️")
            st.rerun()

    # ======================================================
    # ✅ REDO ACTION
    # ======================================================

    if redo_clicked:
        if len(st.session_state["redo_stack"]) > 0:
            current_df = st.session_state["last_saved_df"].copy()
            next_df = st.session_state["redo_stack"].pop()

            st.session_state["undo_stack"].append(current_df)

            next_df = prepare_dataframe(next_df)
            save_db(next_df)

            refresh_data()

            st.session_state["last_saved_df"] = next_df.copy()
            st.session_state["last_saved_hash"] = hash_dataframe(next_df)
            st.session_state["editor_version"] += 1

            st.toast("➡️ Върната е отменената промяна", icon="➡️")
            st.rerun()

    st.write("")

    # ======================================================
    # ✅ DATA EDITOR
    # ======================================================

    edited_df = st.data_editor(
        editable_df,
        use_container_width=True,
        hide_index=False,
        num_rows="dynamic",
        height=720,
        column_order=COLUMNS,
        key=f"differences_excel_editor_{st.session_state['editor_version']}",
        column_config={
            "Delivery No": st.column_config.TextColumn("Delivery No", width="medium"),
            "Invoice No": st.column_config.TextColumn("Invoice No", width="medium"),
            "Invoice Date": st.column_config.TextColumn("Invoice Date", width="medium"),
            "Ref. Number SUPPLIER": st.column_config.TextColumn("Ref. Number SUPPLIER", width="medium"),
            "Other Supplier Ref Num": st.column_config.TextColumn("Other Supplier Ref Num", width="medium"),
            "Price": st.column_config.TextColumn("Price", width="small"),
            "QTY": st.column_config.TextColumn("QTY", width="small"),
            "Received QTY": st.column_config.TextColumn("Received QTY", width="small"),
            "Difference": st.column_config.NumberColumn("Difference", width="small"),

            "Подал разликата": st.column_config.SelectboxColumn(
                "Подал разликата",
                options=["", "Плюс", "Минус"],
                width="small"
            ),

            "Стойност във валутата на доставчика": st.column_config.NumberColumn(
                "Стойност във валутата на доставчика",
                width="medium"
            ),

            "Стойност (в лева)": st.column_config.TextColumn("Стойност (в лева)", width="medium"),
            "Дата на подаване": st.column_config.TextColumn("Дата на подаване", width="medium"),

            "СТАТУС - Попълва се от централата!": st.column_config.SelectboxColumn(
                "СТАТУС - Попълва се от централата!",
                options=[
                    "",
                    "Нова",
                    "Подадена",
                    "Очаква отговор",
                    "Получено КИ",
                    "Отказани",
                    "МАСЛА - обработва се от РМ",
                    "Затворена"
                ],
                width="medium"
            ),

            "№ документа за разлики": st.column_config.TextColumn("№ документа за разлики", width="medium"),
            "Дата на документа за разлики": st.column_config.TextColumn("Дата на документа за разлики", width="medium"),

            "При фактура - дата на приема в Навижън": st.column_config.TextColumn(
                "При фактура - дата на приема в Навижън",
                width="medium"
            ),

            "Обработени в B01": st.column_config.TextColumn("Обработени в B01", width="medium"),
            "Номер на клетка за минуси": st.column_config.TextColumn("Номер на клетка за минуси", width="medium"),

            "Намерени БРОЙКИ (след подаването им)": st.column_config.TextColumn(
                "Намерени БРОЙКИ (след подаването им)",
                width="medium"
            ),

            "Подадена информация към дост. за намерени бройки след подаването им (ДАТА)": st.column_config.TextColumn(
                "Подадена информация към дост. за намерени бройки след подаването им (ДАТА)",
                width="large"
            ),

            "Допълнителен коментар": st.column_config.TextColumn("Допълнителен коментар", width="large"),

            "ДАТА на намиране в В01 на липсващи артикули или коментар от склада": st.column_config.TextColumn(
                "ДАТА на намиране в В01 на липсващи артикули или коментар от склада",
                width="large"
            ),

            "дата на прием в Навижън": st.column_config.TextColumn("дата на прием в Навижън", width="medium"),

            "working days after goods receipt in navision": st.column_config.TextColumn(
                "working days after goods receipt in navision",
                width="medium"
            ),

            "БРАНД": st.column_config.TextColumn("БРАНД", width="medium"),
            "РМ": st.column_config.TextColumn("РМ", width="medium"),
            "формула Ники": st.column_config.TextColumn("формула Ники", width="medium"),

            "дата на прием от таблицата на Тони": st.column_config.TextColumn(
                "дата на прием от таблицата на Тони",
                width="medium"
            ),

            "номер на склада - програмата": st.column_config.TextColumn(
                "номер на склада - програмата",
                width="medium"
            ),

            "системен номер на доставка на склада": st.column_config.TextColumn(
                "системен номер на доставка на склада",
                width="medium"
            )
        }
    )

    # ======================================================
    # ✅ CLEAN / CALCULATE / AUTOSAVE
    # ======================================================

    edited_df = edited_df.copy()

    for col in COLUMNS:
        if col not in edited_df.columns:
            edited_df[col] = ""

    edited_df = edited_df[COLUMNS]

    # ✅ Премахва напълно празни редове
    edited_df = remove_empty_rows(edited_df)

    # ✅ Автоматични формули
    edited_df = apply_auto_calculations(edited_df)

    # ✅ Подрежда колоните пак в правилния ред
    edited_df = prepare_dataframe(edited_df)

    current_hash = hash_dataframe(edited_df)

    # ======================================================
    # ✅ IF CHANGED - PUSH TO UNDO STACK AND SAVE
    # ======================================================

    if current_hash != st.session_state["last_saved_hash"]:
        previous_df = st.session_state["last_saved_df"].copy()

        # ✅ Пази предишното състояние за Назад
        st.session_state["undo_stack"].append(previous_df)

        # ✅ След нова промяна историята за Напред се чисти
        st.session_state["redo_stack"] = []

        # ✅ Пази максимум последните 30 промени
        if len(st.session_state["undo_stack"]) > 30:
            st.session_state["undo_stack"] = st.session_state["undo_stack"][-30:]

        save_db(edited_df)
        refresh_data()

        st.session_state["last_saved_df"] = edited_df.copy()
        st.session_state["last_saved_hash"] = current_hash

        st.toast("✅ Промяната е записана автоматично", icon="✅")

    st.caption(
        "Добавяне на ред: използвай празния ред най-отдолу. "
        "Изтриване на ред: маркирай реда отляво и използвай опцията за изтриване. "
        "Назад / Напред: използвай бутоните над таблицата."
    )
