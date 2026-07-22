import streamlit as st

st.set_page_config(
    page_title="Differences Portal",
    page_icon="📦",
    layout="wide"
)

# =====================================================
# CSS
# =====================================================

st.markdown("""
<style>

.block-container{
    padding-top:1rem;
    max-width:100%;
}

.main-title{
    border:3px solid #000;
    padding:20px;
    text-align:center;
    font-size:28px;
    font-weight:bold;
    margin-bottom:30px;
}

.top-menu{
    border:2px solid #000;
    padding:15px;
    margin-bottom:25px;
}

.stButton>button{
    width:100%;
    height:60px;
    font-size:18px;
    font-weight:bold;
}

.section-box{
    border:2px solid #d9d9d9;
    padding:20px;
    border-radius:10px;
    background:white;
}

.big-table{
    border:1px solid #ddd;
    padding:10px;
    border-radius:10px;
}

</style>
""", unsafe_allow_html=True)

# =====================================================
# HEADER
# =====================================================

col1, col2 = st.columns([1,4])

with col1:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/thumb/6/61/Inter_Cars_logo.svg/512px-Inter_Cars_logo.svg.png",
        width=150,
    )

with col2:

    st.markdown("""
    <div class="main-title">
    DIFFERENCES PORTAL INTER CARS
    </div>
    """, unsafe_allow_html=True)

# =====================================================
# MENU
# =====================================================

st.markdown('<div class="top-menu">', unsafe_allow_html=True)

c1,c2,c3,c4 = st.columns(4)

with c1:
    if st.button("📋 Разлики"):
        st.session_state.page = "differences"

with c2:
    if st.button("📦 Вземане на данни"):
        st.session_state.page = "master"

with c3:
    if st.button("💰 Счетоводство"):
        st.session_state.page = "accounting"

with c4:
    if st.button("⚙️ Администрация"):
        st.session_state.page = "admin"

st.markdown('</div>', unsafe_allow_html=True)

# default page
if "page" not in st.session_state:
    st.session_state.page = "home"

# =====================================================
# HOME
# =====================================================

if st.session_state.page == "home":

    st.info(
        "Изберете секция от менюто."
    )

# =====================================================
# DIFFERENCES
# =====================================================

elif st.session_state.page == "differences":

    st.header("Разлики")

    st.write(
        "Тук ще бъде основната таблица на всички разлики."
    )

    st.data_editor(
        [],
        use_container_width=True,
        num_rows="dynamic"
    )

# =====================================================
# MASTER DATA
# =====================================================

elif st.session_state.page == "master":

    st.header("Вземане на данни")

    st.success(
        "Тук ще се зарежда базата от уникални номера."
    )

    search = st.text_input(
        "Търси вътрешен номер / активен номер"
    )

    st.dataframe(
        [],
        use_container_width=True
    )

# =====================================================
# ACCOUNTING
# =====================================================

elif st.session_state.page == "accounting":

    st.header("Счетоводство")

    number = st.text_input(
        "Въведи номер на разлика"
    )

    st.write(
        "Тук ще се показват всички свързани редове."
    )

    st.dataframe(
        [],
        use_container_width=True
    )

# =====================================================
# ADMIN
# =====================================================

elif st.session_state.page == "admin":

    st.header("Администрация")

    st.write(
        "Управление на доставчици и база данни."
    )

    st.data_editor(
        [],
        use_container_width=True,
        num_rows="dynamic"
    )
