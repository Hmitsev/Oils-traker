import streamlit as st
import pandas as pd
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

STATUS_OPTIONS = [
    "",
    "Нова",
    "Подадена",
    "Очаква отговор",
    "Получено КИ",
    "Отказани",
    "МАСЛА - обработва се от РМ",
    "Затворена"
]

DIFF_TYPE_OPTIONS = [
    "",
    "Плюс",
    "Минус"
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

        .stButton > button {
            border-radius: 14px;
            height: 45px;
            font-weight: 700;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: white;
            border: none;
            box-shadow: 0 8px 16px rgba(37, 99, 235, 0.25);
        }

        .stButton > button:hover {
            background: linear-gradient(135deg, #1d4ed8, #1e40af);
            color: white;
            border: none;
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
    if value is None or value == "":
        return 0.0

    if isinstance(value, str):
        value = value.replace(" ", "").replace(",", ".")

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
    invoice_clean = str(invoice_no).strip().replace("/", "").replace("-", "").replace(" ", "")
    ref_clean = str(ref_number).strip().replace("/", "").replace("-", "").replace(" ", "")
    if invoice_clean and ref_clean:
        return f"{invoice_clean}|{ref_clean}"
    return ""


@st.cache_data(show_spinner=False)
def load_excel():
    if not EXCEL_PATH.exists():
        df_empty = pd.DataFrame(columns=COLUMNS)
        return df_empty

    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
    except Exception:
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl")

    df.columns = [str(c).strip() for c in df.columns]

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNS]

    return df


def save_excel(df):
    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)


def refresh_data():
    st.cache_data.clear()


def format_money(value):
    try:
        return f"{float(value):,.2f}".replace(",", " ")
    except Exception:
        return ""


# ======================================================
# ✅ LOAD DATA
# ======================================================

df = load_excel()

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

    st.caption("Файл:")
    st.write(str(EXCEL_PATH))

    if EXCEL_PATH.exists():
        st.success("Excel файлът е намерен")
    else:
        st.warning("Excel файлът липсва")

# ======================================================
# ✅ HEADER
# ======================================================

st.markdown('<div class="main-title">Differences Suppliers</div>', unsafe_allow_html=True)
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
        plus_count = (df["Подал разликата"].astype(str).str.lower() == "плюс").sum()
        minus_count = (df["Подал разликата"].astype(str).str.lower() == "минус").sum()

        status_col = "СТАТУС - Попълва се от централата!"
        open_count = df[status_col].astype(str).isin(["", "Нова", "Подадена", "Очаква отговор"]).sum()

        total_bgn = pd.to_numeric(
            df["Стойност (в лева)"].astype(str)
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

    st.markdown("### Последни 20 записа")
    if df.empty:
        st.info("Все още няма данни.")
    else:
        st.dataframe(df.tail(20), use_container_width=True, hide_index=True)

# ======================================================
# ✅ ALL DIFFERENCES
# ======================================================

elif page == "📋 Всички разлики":
    st.markdown('<div class="section-card">', unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)

    with f1:
        search_text = st.text_input("Търсене", placeholder="артикул, фактура, доставка...")

    with f2:
        status_filter = st.selectbox(
            "Статус",
            ["Всички"] + sorted([x for x in df["СТАТУС - Попълва се от централата!"].dropna().astype(str).unique() if x.strip()])
            if not df.empty else ["Всички"]
        )

    with f3:
        brand_filter = st.selectbox(
            "Бранд",
            ["Всички"] + sorted([x for x in df["БРАНД"].dropna().astype(str).unique() if x.strip()])
            if not df.empty else ["Всички"]
        )

    with f4:
        rm_filter = st.selectbox(
            "РМ",
            ["Всички"] + sorted([x for x in df["РМ"].dropna().astype(str).unique() if x.strip()])
            if not df.empty else ["Всички"]
        )

    filtered_df = df.copy()

    if search_text:
        mask = filtered_df.astype(str).apply(
            lambda row: row.str.contains(search_text, case=False, na=False).any(),
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
# ✅ NEW DIFFERENCE
# ======================================================

# ======================================================
# ✅ EXCEL-LIKE EDITOR
# ======================================================

elif page == "📝 Въвеждане / редакция":
    st.markdown("### 📝 Въвеждане / редакция на разлики")
    st.caption("Работа като в Excel - добавяш нов ред най-отдолу и попълваш колона по колона.")

    if df.empty:
        editable_df = pd.DataFrame(columns=COLUMNS)
    else:
        editable_df = df.copy()

    st.markdown(
        """
        <div class="warning-box">
            Важно: След като добавиш или редактираш редове, натисни бутона <b>💾 Запази промените</b>.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")

    edited_df = st.data_editor(
        editable_df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        height=650,
        column_order=COLUMNS,
        key="differences_editor"
    )

    c1, c2, c3 = st.columns([1, 1, 4])

    with c1:
        save_button = st.button("💾 Запази промените")

    with c2:
        refresh_button = st.button("🔄 Презареди")

    if refresh_button:
        refresh_data()
        st.rerun()

    if save_button:
        edited_df = edited_df.copy()

        for col in COLUMNS:
            if col not in edited_df.columns:
                edited_df[col] = ""

        edited_df = edited_df[COLUMNS]

        # ✅ Автоматично изчисляване на Difference
        edited_df["Difference"] = edited_df.apply(
            lambda row: calculate_difference(
                row.get("QTY", 0),
                row.get("Received QTY", 0)
            ),
            axis=1
        )

        # ✅ Автоматично изчисляване на стойност във валутата на доставчика
        edited_df["Стойност във валутата на доставчика"] = edited_df.apply(
            lambda row: calculate_supplier_value(
                row.get("Price", 0),
                row.get("Difference", 0)
            ),
            axis=1
        )

        # ✅ Автоматично попълване на Дата на подаване, ако е празна
        if "Дата на подаване" in edited_df.columns:
            edited_df["Дата на подаване"] = edited_df["Дата на подаване"].apply(
                lambda x: date.today() if pd.isna(x) or str(x).strip() == "" or str(x).strip().lower() == "none" else x
            )

        # ✅ Автоматично формула Ники: Invoice No | Ref. Number SUPPLIER
        edited_df["формула Ники"] = edited_df.apply(
            lambda row: make_niki_formula(
                row.get("Invoice No", ""),
                row.get("Ref. Number SUPPLIER", "")
            ),
            axis=1
        )

        # ✅ Премахване на напълно празни редове
        edited_df = edited_df.dropna(how="all")

        # ✅ Запис в Excel файла
        save_excel(edited_df)
        refresh_data()

        st.success("✅ Промените са записани успешно в Excel файла.")
        st.rerun()

# ======================================================
# ✅ SEARCH
# ======================================================

elif page == "🔎 Проверка / търсене":
    st.markdown("### 🔎 Проверка за вече подадена разлика")

    search_value = st.text_input(
        "Въведи Invoice No, Ref. Number SUPPLIER, Delivery No или формула Ники"
    )

    if search_value:
        result = df[
            df.astype(str).apply(
                lambda row: row.str.contains(search_value, case=False, na=False).any(),
                axis=1
            )
        ]

        if result.empty:
            st.warning("Няма намерен запис.")
        else:
            st.success(f"Намерени записи: {len(result)}")
            st.dataframe(result, use_container_width=True, hide_index=True)
