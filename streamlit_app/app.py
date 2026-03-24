import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Teko - BDA Daily Report", layout="wide")

# Header
col1, col2 = st.columns([1, 5])
with col1:
    st.markdown("## 🟠 TEKO")
with col2:
    st.title("BDA Daily Market Report")
    st.caption(f"Report Date: {datetime.now().strftime('%d/%m/%Y')}")

st.divider()

# 3 Tabs
tab1, tab2, tab3 = st.tabs(["📋 Internal Data Input", "📊 Data Preview", "📥 Generate & Download"])

with tab1:
    st.header("Internal Bank Data")
    st.info("Enter today's internal data below")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Liquidez - Moeda Nacional (Kz)")
        reservas = st.number_input("Reservas Livres BNA", value=0.0)
        do_comerciais = st.number_input("DO B. Comerciais", value=0.0)
        dp_comerciais = st.number_input("DP B. Comerciais", value=0.0)
        omas = st.number_input("OMAs", value=0.0)

    with col2:
        st.subheader("Liquidez - Moeda Estrangeira (USD)")
        saldo_do = st.number_input("SALDO D.O Estrangeiros", value=0.0)
        dps_me = st.number_input("DPs ME", value=0.0)
        colateral = st.number_input("COLATERAL CDI", value=0.0)

    st.divider()
    st.subheader("Taxas de Câmbio")
    col3, col4, col5 = st.columns(3)
    with col3:
        usd_akz = st.number_input("USD/AKZ", value=912.43)
    with col4:
        eur_akz = st.number_input("EUR/AKZ", value=1057.69)
    with col5:
        eur_usd = st.number_input("EUR/USD", value=1.16)

    if st.button("💾 Save Internal Data", type="primary"):
        st.success("✅ Internal data saved!")

with tab2:
    st.header("Data Preview")
    st.info("Click the button below to load and combine all data")
    if st.button("🔄 Load All Data"):
        st.warning("⏳ Scraping external sources... (coming soon)")

with tab3:
    st.header("Generate Report")
    st.info("Once data is loaded, generate your PowerPoint report here")
    if st.button("📊 Generate PowerPoint", type="primary"):
        st.warning("⏳ Report generation coming soon...")
