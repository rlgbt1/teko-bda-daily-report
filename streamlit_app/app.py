import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="Teko - BDA Daily Report", layout="wide", page_icon="🟠")

# Header
col1, col2 = st.columns([1, 6])
with col1:
    st.markdown("## 🟠 TEKO")
with col2:
    st.title("BDA Daily Market Report")
    st.caption(f"Report Date: {datetime.now().strftime('%d/%m/%Y')}")

st.divider()

# Session state init
if "internal_data" not in st.session_state:
    st.session_state.internal_data = {}
if "external_data" not in st.session_state:
    st.session_state.external_data = {}
if "report_path" not in st.session_state:
    st.session_state.report_path = None

tab1, tab2, tab3 = st.tabs(["📋 Internal Data Input", "📊 Data Preview", "📥 Generate & Download"])

# ── TAB 1: INTERNAL DATA ──────────────────────────────────────────────────────
with tab1:
    st.header("Internal Bank Data")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Liquidez - Moeda Nacional (Kz milhares)")
        reservas = st.number_input("Reservas Livres BNA", value=0.0, format="%.2f")
        do_com = st.number_input("DO B. Comerciais", value=0.0, format="%.2f")
        dp_com = st.number_input("DP B. Comerciais", value=0.0, format="%.2f")
        omas = st.number_input("OMAs", value=0.0, format="%.2f")

    with col2:
        st.subheader("Liquidez - Moeda Estrangeira (USD milhões)")
        saldo_do = st.number_input("SALDO D.O Estrangeiros", value=0.0, format="%.2f")
        dps_me = st.number_input("DPs ME", value=0.0, format="%.2f")
        colateral = st.number_input("COLATERAL CDI", value=0.0, format="%.2f")

    st.divider()
    st.subheader("Taxas de Câmbio")
    col3, col4, col5 = st.columns(3)
    with col3:
        usd_akz = st.number_input("USD/AKZ", value=912.43, format="%.2f")
    with col4:
        eur_akz = st.number_input("EUR/AKZ", value=1057.69, format="%.2f")
    with col5:
        eur_usd = st.number_input("EUR/USD", value=1.16, format="%.4f")

    st.divider()
    st.subheader("Operações Vivas")
    num_ops = st.number_input("Número de operações", min_value=0, max_value=20, value=0, step=1)
    operations = []
    for i in range(int(num_ops)):
        with st.expander(f"Operação {i+1}"):
            c1, c2, c3, c4 = st.columns(4)
            op = {
                "Contraparte": c1.text_input("Contraparte", key=f"cp_{i}"),
                "Montante": c2.number_input("Montante", key=f"amt_{i}", value=0.0),
                "Taxa (%)": c3.number_input("Taxa %", key=f"rate_{i}", value=0.0),
                "Maturidade (dias)": c4.number_input("Dias", key=f"mat_{i}", value=0, step=1),
            }
            operations.append(op)

    if st.button("💾 Save Internal Data", type="primary"):
        st.session_state.internal_data = {
            "liquidez_mn": {
                "Reservas Livres BNA": reservas,
                "DO B. Comerciais": do_com,
                "DP B. Comerciais": dp_com,
                "OMAs": omas,
                "LIQUIDEZ TOTAL": reservas + do_com + dp_com + omas,
            },
            "liquidez_me": {
                "SALDO D.O Estrangeiros": saldo_do,
                "DPs ME": dps_me,
                "COLATERAL CDI": colateral,
                "LIQUIDEZ TOTAL ME": saldo_do + dps_me + colateral,
            },
            "fx_rates": {"USD/AKZ": usd_akz, "EUR/AKZ": eur_akz, "EUR/USD": eur_usd},
            "operations": operations,
        }
        st.success("✅ Internal data saved!")

# ── TAB 2: DATA PREVIEW ───────────────────────────────────────────────────────
with tab2:
    st.header("Data Preview")

    if st.button("🔄 Scrape External Data"):
        with st.spinner("Fetching market data... (30-60 seconds)"):
            try:
                from src.scrapers.market_aggregator import scrape_all_external_data
                st.session_state.external_data = scrape_all_external_data()
                st.success("✅ External data loaded!")
            except Exception as e:
                st.error(f"Error scraping data: {e}")

    if st.session_state.external_data:
        ed = st.session_state.external_data

        if "markets" in ed and not ed["markets"].empty:
            st.subheader("🌍 Mercados Globais")
            st.dataframe(ed["markets"], use_container_width=True)

        if "commodities" in ed and not ed["commodities"].empty:
            st.subheader("🛢️ Commodities")
            st.dataframe(ed["commodities"], use_container_width=True)

        if "crypto" in ed and not ed["crypto"].empty:
            st.subheader("₿ Criptomoedas")
            st.dataframe(ed["crypto"], use_container_width=True)

        if "luibor" in ed and not ed["luibor"].empty:
            st.subheader("🏦 Taxas LUIBOR (BNA)")
            st.dataframe(ed["luibor"], use_container_width=True)

    if st.session_state.internal_data:
        st.subheader("🏦 Liquidez Interna")
        mn = st.session_state.internal_data.get("liquidez_mn", {})
        me = st.session_state.internal_data.get("liquidez_me", {})
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Moeda Nacional (Kz)**")
            st.dataframe(pd.DataFrame(mn.items(), columns=["Descrição", "Valor"]))
        with col2:
            st.write("**Moeda Estrangeira (USD)**")
            st.dataframe(pd.DataFrame(me.items(), columns=["Descrição", "Valor"]))

# ── TAB 3: GENERATE & DOWNLOAD ────────────────────────────────────────────────
with tab3:
    st.header("Generate Report")
    report_date = st.date_input("Report Date", datetime.now())

    col1, col2 = st.columns(2)
    with col1:
        use_ai = st.checkbox("Use AI Summaries (requires Ollama running)", value=False)
    with col2:
        st.info("💡 Start Ollama first: run `ollama serve` in terminal")

    if st.button("📊 Generate PowerPoint Report", type="primary"):
        if not st.session_state.external_data and not st.session_state.internal_data:
            st.warning("⚠️ Please load data first in the Data Preview tab.")
        else:
            with st.spinner("Building report..."):
                try:
                    from src.report_generator.pptx_builder import ReportBuilder
                    from src.agents.ai_agent import DailyReportAgent

                    date_str = report_date.strftime("%d/%m/%Y")
                    builder = ReportBuilder()
                    ed = st.session_state.external_data

                    # AI summaries
                    markets_summary = ""
                    commodities_summary = ""
                    crypto_summary = ""

                    if use_ai:
                        agent = DailyReportAgent()
                        if "markets" in ed:
                            markets_summary = agent.summarize_markets(ed.get("markets"))
                        if "commodities" in ed:
                            commodities_summary = agent.summarize_commodities(ed.get("commodities"))
                        if "crypto" in ed:
                            crypto_summary = agent.summarize_crypto(ed.get("crypto"))

                    # Build slides
                    builder.add_title_slide(date_str)

                    if "markets" in ed and not ed["markets"].empty:
                        builder.add_dataframe_slide("Mercados Globais - Capital Markets", date_str, ed["markets"], markets_summary)

                    if "commodities" in ed and not ed["commodities"].empty:
                        builder.add_dataframe_slide("Commodities & Minerais", date_str, ed["commodities"], commodities_summary)

                    if "crypto" in ed and not ed["crypto"].empty:
                        builder.add_dataframe_slide("Criptomoedas", date_str, ed["crypto"], crypto_summary)

                    if "luibor" in ed and not ed["luibor"].empty:
                        builder.add_dataframe_slide("Taxas LUIBOR - BNA", date_str, ed["luibor"])

                    # Internal data slides
                    id = st.session_state.internal_data
                    if id.get("liquidez_mn"):
                        mn_df = pd.DataFrame(id["liquidez_mn"].items(), columns=["Descrição", "Valor (Kz)"])
                        builder.add_dataframe_slide("Liquidez - Moeda Nacional", date_str, mn_df)

                    if id.get("liquidez_me"):
                        me_df = pd.DataFrame(id["liquidez_me"].items(), columns=["Descrição", "Valor (USD)"])
                        builder.add_dataframe_slide("Liquidez - Moeda Estrangeira", date_str, me_df)

                    # Save
                    os.makedirs("reports/generated", exist_ok=True)
                    path = f"reports/generated/BDA_Report_{report_date.strftime('%Y%m%d')}.pptx"
                    builder.save(path)
                    st.session_state.report_path = path
                    st.success("✅ Report generated!")

                except Exception as e:
                    st.error(f"Error generating report: {e}")
                    import traceback
                    st.code(traceback.format_exc())

    # Download buttons
    if st.session_state.report_path and os.path.exists(st.session_state.report_path):
        st.divider()
        with open(st.session_state.report_path, "rb") as f:
            st.download_button(
                label="📥 Download PPTX",
                data=f.read(),
                file_name=os.path.basename(st.session_state.report_path),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

