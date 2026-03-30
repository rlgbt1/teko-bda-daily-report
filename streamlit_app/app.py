"""
streamlit_app/app.py — BDA Daily Report — Streamlit frontend.

Tabs:
  1. Internal Data Input  — treasury / internal KPIs entered by the user
  2. Data Preview         — scrape external data & inspect combined tables
  3. Generate & Download  — build the PPTX, optional AI summaries, download
"""

import os
import sys
import traceback
from datetime import datetime

# Ensure the project root is on sys.path so `src.*` imports work regardless
# of how Streamlit is launched (streamlit run, python -m streamlit, etc.)
_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Teko – BDA Daily Report",
    layout="wide",
    page_icon="🟠",
)

# ── Header ────────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 7])
with col_logo:
    st.markdown("## 🟠 TEKO")
with col_title:
    st.title("BDA — Resumo Diário dos Mercados")
    st.caption(f"Data: {datetime.now().strftime('%d/%m/%Y')}")

st.divider()

# ── Session state init ────────────────────────────────────────────────────────
for key in ("internal_data", "external_data", "report_path", "pdf_path", "ai_summaries"):
    if key not in st.session_state:
        st.session_state[key] = {} if key not in ("report_path", "pdf_path") else None

# ─────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📋 Dados Internos", "📊 Pré-visualização", "📥 Gerar & Descarregar"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — INTERNAL DATA INPUT
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.header("Entrada de Dados Internos (Tesouraria)")

    c1, c2 = st.columns(2)

    # ── Liquidez MN ──────────────────────────────────────────────────────────
    with c1:
        st.subheader("Liquidez – Moeda Nacional (Kz milhares)")
        reservas  = st.number_input("Posição Reservas Livres BNA", value=0.0, format="%.2f")
        do_com    = st.number_input("Posição DO B. Comerciais",    value=0.0, format="%.2f")
        dp_com    = st.number_input("Posição DP B. Comerciais",    value=0.0, format="%.2f")
        omas      = st.number_input("Posição OMAs",                value=0.0, format="%.2f")

    # ── Liquidez ME ──────────────────────────────────────────────────────────
    with c2:
        st.subheader("Liquidez – Moeda Estrangeira (USD milhões)")
        saldo_do  = st.number_input("SALDO D.O Estrangeiros", value=0.0, format="%.2f")
        dps_me    = st.number_input("DPs ME",                 value=0.0, format="%.2f")
        colateral = st.number_input("COLATERAL CDI",          value=0.0, format="%.2f")

    st.divider()

    # ── Taxas de Câmbio ───────────────────────────────────────────────────────
    st.subheader("Taxas de Câmbio")
    cx1, cx2, cx3 = st.columns(3)
    with cx1:
        usd_akz = st.number_input("USD/AKZ", value=912.43, format="%.4f")
    with cx2:
        eur_akz = st.number_input("EUR/AKZ", value=1057.69, format="%.4f")
    with cx3:
        eur_usd = st.number_input("EUR/USD", value=1.16, format="%.4f")

    st.divider()

    # ── Operações Vivas ───────────────────────────────────────────────────────
    st.subheader("Operações Vivas (DP / OMA)")
    num_ops = st.number_input("Número de operações", min_value=0, max_value=30, value=0, step=1)
    operations: list[dict] = []
    for i in range(int(num_ops)):
        with st.expander(f"Operação {i + 1}"):
            oc1, oc2, oc3, oc4, oc5 = st.columns(5)
            op = {
                "tipo":        oc1.selectbox("Tipo",  ["DP", "OMA", "REPO"], key=f"tipo_{i}"),
                "contraparte": oc2.text_input("Contraparte", key=f"cp_{i}"),
                "montante":    oc3.number_input("Montante", key=f"amt_{i}", value=0.0),
                "taxa":        oc4.number_input("Taxa %", key=f"rate_{i}", value=0.0),
                "residual":    oc5.number_input("Dias Residuais", key=f"mat_{i}", value=0, step=1),
            }
            operations.append(op)

    st.divider()

    # ── P&L Control ──────────────────────────────────────────────────────────
    st.subheader("P&L Control")
    pl_c1, pl_c2, pl_c3 = st.columns(3)
    reembolso_n   = pl_c1.number_input("Reembolso Crédito — Nº Ops",  value=0, step=1)
    reembolso_val = pl_c2.number_input("Reembolso Crédito — Montante", value=0.0, format="%.2f")
    desembolso_n  = pl_c3.number_input("Desembolso Crédito — Nº Ops",  value=0, step=1)

    if st.button("💾 Guardar Dados Internos", type="primary"):
        total_mn = reservas + do_com + dp_com + omas
        total_me = saldo_do + dps_me + colateral

        st.session_state.internal_data = {
            # rows expected by pptx_builder
            "liquidez_mn_rows": [
                {"label": "Posição Reservas Livres BNA", "values": ["—", "—", "—", "—", f"{reservas:,.0f}"]},
                {"label": "Posição DO B. Comerciais",    "values": ["—", "—", "—", "—", f"{do_com:,.0f}"]},
                {"label": "Posição DP B. Comerciais",    "values": ["—", "—", "—", "—", f"{dp_com:,.0f}"]},
                {"label": "Posição OMAs",                "values": ["—", "—", "—", "—", f"{omas:,.0f}"]},
                {"label": "LIQUIDEZ BDA",                "values": ["—", "—", "—", "—", f"{total_mn:,.2f}"]},
            ],
            "liquidez_me_rows": [
                {"label": "SALDO D.O Estrangeiros", "values": ["—", "—", "—", "—", f"{saldo_do:,.2f}"]},
                {"label": "DPs ME",                 "values": ["—", "—", "—", "—", f"{dps_me:,.2f}"]},
                {"label": "COLATERAL CDI",           "values": ["—", "—", "—", "—", f"{colateral:,.2f}"]},
                {"label": "LIQUIDEZ BDA",            "values": ["—", "—", "—", "—", f"{total_me:,.2f}"]},
            ],
            "cambial": {
                "usd_akz": f"{usd_akz:,.4f}",
                "eur_akz": f"{eur_akz:,.4f}",
                "eur_usd": f"{eur_usd:,.4f}",
            },
            "cambial_rows": [
                {"par": "USD/AKZ", "anterior2": "—", "anterior": "—",
                 "atual": f"{usd_akz:,.4f}", "variacao": "—"},
                {"par": "EUR/AKZ", "anterior2": "—", "anterior": "—",
                 "atual": f"{eur_akz:,.4f}", "variacao": "—"},
                {"par": "EUR/USD", "anterior2": "—", "anterior": "—",
                 "atual": f"{eur_usd:,.4f}", "variacao": "—"},
            ],
            "operacoes_vivas": operations,
            "pl_summary": [
                {"label": "Reembolso de Crédito", "n_ops": reembolso_n,  "montante": f"{reembolso_val:,.2f}"},
                {"label": "Fornecedores",          "n_ops": "—",          "montante": "—"},
                {"label": "Desembolso de Crédito", "n_ops": desembolso_n, "montante": "—"},
            ],
            "kpis": [
                {"label": "Liquidez MN (mM Kz)",   "value": f"{total_mn/1000:,.2f}",  "variation_str": ""},
                {"label": "Liquidez ME (M USD)",    "value": f"{total_me:,.2f}",        "variation_str": ""},
                {"label": "USD/AKZ",                "value": f"{usd_akz:,.2f}",         "variation_str": ""},
                {"label": "EUR/AKZ",                "value": f"{eur_akz:,.2f}",         "variation_str": ""},
                {"label": "Reembolso Crédito (Kz)", "value": f"{reembolso_val:,.2f}",   "variation_str": ""},
                {"label": "DO B. Comerciais",        "value": f"{do_com:,.0f}",          "variation_str": ""},
                {"label": "DP B. Comerciais",        "value": f"{dp_com:,.0f}",          "variation_str": ""},
                {"label": "OMAs",                    "value": f"{omas:,.0f}",            "variation_str": ""},
            ],
            "reembolso_credito": f"{reembolso_val:,.2f} M Kz",
        }
        st.success("✅ Dados internos guardados!")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — DATA PREVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header("Pré-visualização de Dados")

    if st.button("🔄 Actualizar Dados Externos (scrape)"):
        with st.spinner("A recolher dados de mercado… (30-60 s)"):
            try:
                from src.scrapers.market_aggregator import scrape_all_external_data
                st.session_state.external_data = scrape_all_external_data()
                st.success("✅ Dados externos carregados!")
            except Exception as e:
                st.error(f"Erro ao recolher dados: {e}")
                st.code(traceback.format_exc())

    ed = st.session_state.external_data

    if ed:
        maps = [
            ("markets",     "🌍 Mercados Globais"),
            ("commodities", "🛢️ Commodities"),
            ("crypto",      "₿ Criptomoedas"),
            ("luibor",      "🏦 Taxas LUIBOR (BNA)"),
            ("fx_rates",    "💱 Taxas de Câmbio (BNA)"),
        ]
        for key, label in maps:
            item = ed.get(key)
            if item is not None and not (hasattr(item, "empty") and item.empty):
                st.subheader(label)
                if isinstance(item, pd.DataFrame):
                    st.dataframe(item, use_container_width=True)
                else:
                    st.json(item)

    id_ = st.session_state.internal_data
    if id_:
        st.subheader("🏦 Liquidez Interna (entrada manual)")
        c_mn, c_me = st.columns(2)
        with c_mn:
            st.write("**Moeda Nacional (Kz)**")
            mn_rows = id_.get("liquidez_mn_rows", [])
            if mn_rows:
                st.dataframe(
                    pd.DataFrame(
                        [{"Descrição": r["label"], "D (actual)": r["values"][-1]} for r in mn_rows]
                    ),
                    use_container_width=True,
                )
        with c_me:
            st.write("**Moeda Estrangeira (USD)**")
            me_rows = id_.get("liquidez_me_rows", [])
            if me_rows:
                st.dataframe(
                    pd.DataFrame(
                        [{"Descrição": r["label"], "D (actual)": r["values"][-1]} for r in me_rows]
                    ),
                    use_container_width=True,
                )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — GENERATE & DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header("Gerar Relatório")

    report_date = st.date_input("Data do Relatório", datetime.now())

    col_ai, col_hint = st.columns(2)
    with col_ai:
        use_ai = st.checkbox(
            "Usar Resumos IA (Gemini)",
            value=False,
            help="Requer GEMINI_API_KEY no ficheiro .env",
        )
    with col_hint:
        st.info(
            "💡 Para activar a IA: adicione `GEMINI_API_KEY=<chave>` ao ficheiro `.env` "
            "na raiz do projecto."
        )

    if st.button("📊 Gerar Relatório PowerPoint", type="primary"):
        ed  = st.session_state.external_data
        id_ = st.session_state.internal_data

        if not ed and not id_:
            st.warning("⚠️  Carregue primeiro os dados externos (Tab 2) ou introduza dados internos (Tab 1).")
        else:
            with st.spinner("A construir relatório…"):
                try:
                    from src.report_generator.pptx_builder import BDAReportGenerator
                    from src.agents.ai_agent import DailyReportAgent

                    date_str = report_date.strftime("%d.%m.%Y")

                    # ── Build AI summaries ────────────────────────────────────
                    summaries: dict = {}
                    if use_ai:
                        agent = DailyReportAgent()
                        if ed.get("markets") is not None:
                            summaries["cm_commentary"] = agent.summarize_markets(ed["markets"])
                        if ed.get("crypto") is not None:
                            summaries["crypto_commentary"] = agent.summarize_crypto(ed["crypto"])
                        if ed.get("commodities") is not None:
                            summaries["commodities_commentary"] = agent.summarize_commodities(ed["commodities"])

                    # ── Map external data → market_info dict ─────────────────
                    def _df_to_rows(df, name_col, val_col, var_col=None):
                        """Convert a Yahoo-style DataFrame to the pptx row schema."""
                        if df is None or df.empty:
                            return []
                        rows = []
                        for _, r in df.iterrows():
                            rows.append({
                                "indice" if "indice" not in r else "nome": str(r.get(name_col, "—")),
                                "nome":     str(r.get(name_col, "—")),
                                "anterior": str(r.get("Anterior", r.get("anterior", "—"))),
                                "atual":    str(r.get("Atual",    r.get("atual",    "—"))),
                                "variacao": str(r.get("Var (%)",  r.get("variacao", "—"))),
                            })
                        return rows

                    markets_df     = ed.get("markets")
                    commodities_df = ed.get("commodities")
                    crypto_df      = ed.get("crypto")

                    # Detect the name column automatically
                    def _first_str_col(df):
                        if df is None or df.empty:
                            return None
                        for col in df.columns:
                            if df[col].dtype == object:
                                return col
                        return df.columns[0]

                    cm_rows  = []
                    if markets_df is not None and not markets_df.empty:
                        nc = _first_str_col(markets_df)
                        for _, r in markets_df.iterrows():
                            cm_rows.append({
                                "indice":   str(r.get(nc, "—")),
                                "anterior": str(r.get("Anterior", "—")),
                                "atual":    str(r.get("Atual",    "—")),
                                "variacao": str(r.get("Var (%)",  "—")),
                            })

                    cr_rows  = []
                    if crypto_df is not None and not crypto_df.empty:
                        nc = _first_str_col(crypto_df)
                        for _, r in crypto_df.iterrows():
                            cr_rows.append({
                                "moeda":    str(r.get(nc, "—")),
                                "anterior": str(r.get("Anterior", "—")),
                                "atual":    str(r.get("Atual",    "—")),
                                "variacao": str(r.get("Var (%)",  "—")),
                            })

                    cmd_rows = []
                    if commodities_df is not None and not commodities_df.empty:
                        nc = _first_str_col(commodities_df)
                        for _, r in commodities_df.iterrows():
                            cmd_rows.append({
                                "nome":     str(r.get(nc, "—")),
                                "anterior": str(r.get("Anterior", "—")),
                                "atual":    str(r.get("Atual",    "—")),
                                "variacao": str(r.get("Var (%)",  "—")),
                            })

                    market_info = {
                        "capital_markets":         cm_rows  or [],
                        "crypto":                  cr_rows  or [],
                        "commodities":             cmd_rows or [],
                        "minerais":                [],
                        "cm_commentary":           summaries.get("cm_commentary",           ""),
                        "crypto_commentary":       summaries.get("crypto_commentary",       ""),
                        "commodities_commentary":  summaries.get("commodities_commentary",  ""),
                        "minerais_commentary":     summaries.get("minerais_commentary",     ""),
                    }

                    # ── LUIBOR → luibor dict ──────────────────────────────────
                    luibor: dict = {}
                    luibor_var: dict = {}
                    luibor_df = ed.get("luibor")
                    if luibor_df is not None and not luibor_df.empty:
                        for _, r in luibor_df.iterrows():
                            mat = str(r.get("Maturidade", r.get(luibor_df.columns[0], "")))
                            rate = str(r.get("Taxa (%)", r.get(luibor_df.columns[1] if len(luibor_df.columns) > 1 else "Taxa", "—")))
                            luibor[mat]     = rate
                            luibor_var[mat] = str(r.get("Var (%)", "—"))

                    # ── Combine everything into the data dict ─────────────────
                    data = {
                        "report_date": date_str,
                        "market_info": market_info,
                        "luibor":      luibor,
                        "luibor_variation": luibor_var,
                        **id_,          # internal data keys (kpis, cambial, etc.)
                    }

                    # ── Generate PPTX ─────────────────────────────────────────
                    os.makedirs("output", exist_ok=True)
                    date_tag  = report_date.strftime("%Y%m%d")
                    pptx_path = f"output/BDA_Report_{date_tag}.pptx"
                    gen  = BDAReportGenerator(data)
                    gen.build(pptx_path)
                    st.session_state.report_path = pptx_path

                    # ── Generate PDF ───────────────────────────────────────────
                    try:
                        from src.report_generator.pdf_builder import BDAReportPDF
                        pdf_path = f"output/BDA_Report_{date_tag}.pdf"
                        BDAReportPDF(data).build(pdf_path)
                        st.session_state.pdf_path = pdf_path
                        st.success(f"✅ Relatório gerado (PPTX + PDF): `output/BDA_Report_{date_tag}.*`")
                    except Exception as pdf_err:
                        st.session_state.pdf_path = None
                        st.success(f"✅ PPTX gerado: `{pptx_path}`")
                        st.warning(f"PDF não gerado: {pdf_err}")

                except Exception as e:
                    st.error(f"Erro ao gerar relatório: {e}")
                    st.code(traceback.format_exc())

    # ── Download buttons ──────────────────────────────────────────────────────
    rp  = st.session_state.report_path
    pdp = st.session_state.pdf_path

    if (rp and os.path.exists(rp)) or (pdp and os.path.exists(pdp)):
        st.divider()
        st.markdown("**Descarregar Relatório**")
        dl_col1, dl_col2 = st.columns(2)

        with dl_col1:
            if rp and os.path.exists(rp):
                with open(rp, "rb") as f:
                    st.download_button(
                        label="📊 Descarregar PPTX",
                        data=f.read(),
                        file_name=os.path.basename(rp),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
            else:
                st.button("📊 PPTX não disponível", disabled=True, use_container_width=True)

        with dl_col2:
            if pdp and os.path.exists(pdp):
                with open(pdp, "rb") as f:
                    st.download_button(
                        label="📄 Descarregar PDF",
                        data=f.read(),
                        file_name=os.path.basename(pdp),
                        mime="application/pdf",
                        use_container_width=True,
                    )
            else:
                st.button("📄 PDF não disponível", disabled=True, use_container_width=True)
