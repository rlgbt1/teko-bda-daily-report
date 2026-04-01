"""
streamlit_app/app.py — BDA Daily Report — Streamlit frontend.

Tabs:
  1. Dados Internos     — all manual treasury / KPI / operations data
  2. Pré-visualização  — scrape external data & inspect combined tables
  3. Gerar & Descarregar — build the PPTX, optional AI summaries, download
"""

import os
import sys
import traceback
from datetime import datetime

_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

import pandas as pd
import streamlit as st

from src.llm.llm_client import get_provider_name

st.set_page_config(
    page_title="Teko – BDA Daily Report",
    layout="wide",
    page_icon="🟠",
)

# ── Header ─────────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 7])
with col_logo:
    st.markdown("## 🟠 TEKO")
with col_title:
    st.title("BDA — Resumo Diário dos Mercados")
    st.caption(f"Data: {datetime.now().strftime('%d/%m/%Y')}")

st.divider()

# ── Session state init ─────────────────────────────────────────────────────────
for key in ("internal_data", "external_data", "report_path", "pdf_path"):
    if key not in st.session_state:
        st.session_state[key] = {} if key not in ("report_path", "pdf_path") else None

# ─────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📋 Dados Internos", "📊 Pré-visualização", "📥 Gerar & Descarregar"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — INTERNAL DATA INPUT
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.header("Entrada de Dados Internos (Tesouraria)")

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: LIQUIDEZ MN — 5 days of data
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("📊 Liquidez – Moeda Nacional (Kz Milhares)", expanded=True):
        st.caption("Introduza os valores para cada dia (D-4 → D). Deixe 0 para dias sem dados.")
        day_cols = st.columns(5)
        day_labels = ["D-4", "D-3", "D-2", "D-1", "D"]

        lmn_fields = [
            ("reservas",  "Posição Reservas Livres BNA"),
            ("do_com",    "Posição DO B. Comerciais"),
            ("dp_com",    "Posição DP B. Comerciais"),
            ("omas",      "Posição OMAs"),
        ]
        lmn_vals = {}
        for key, label in lmn_fields:
            st.markdown(f"**{label}**")
            row_cols = st.columns(5)
            lmn_vals[key] = []
            for j, (col, dlbl) in enumerate(zip(row_cols, day_labels)):
                v = col.number_input(dlbl, value=0.0, format="%.2f",
                                     key=f"lmn_{key}_{j}", label_visibility="visible")
                lmn_vals[key].append(v)

        juros_diario_mn = st.number_input("Juros Diário MN (Kz Milhares)", value=0.0, format="%.2f",
                                          key="juros_diario_mn")

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: LIQUIDEZ ME — 5 days
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("💵 Liquidez – Moeda Estrangeira (USD Milhões)", expanded=True):
        st.caption("Introduza os valores para cada dia (D-4 → D). Deixe 0 para dias sem dados.")

        lme_fields = [
            ("saldo_do",  "SALDO D.O Estrangeiros"),
            ("dps_me",    "DPs ME"),
            ("colateral", "COLATERAL CDI"),
        ]
        lme_vals = {}
        for key, label in lme_fields:
            st.markdown(f"**{label}**")
            row_cols = st.columns(5)
            lme_vals[key] = []
            for j, (col, dlbl) in enumerate(zip(row_cols, day_labels)):
                v = col.number_input(dlbl, value=0.0, format="%.4f",
                                     key=f"lme_{key}_{j}", label_visibility="visible")
                lme_vals[key].append(v)

        juros_diario_me = st.number_input("Juros Diário ME (USD)", value=0.0, format="%.2f",
                                          key="juros_diario_me")

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: TAXAS DE CÂMBIO (3 days: D-2, D-1, D)
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("💱 Taxas de Câmbio", expanded=True):
        fx_pairs = ["USD/AKZ", "EUR/AKZ", "EUR/USD"]
        fx_defaults = [(912.43, 912.43, 912.43), (1057.69, 1057.69, 1057.69), (1.16, 1.16, 1.16)]
        fx_vals = {}
        for pair, defaults in zip(fx_pairs, fx_defaults):
            st.markdown(f"**{pair}**")
            fc1, fc2, fc3, fc4 = st.columns(4)
            k = pair.replace("/", "_")
            fx_vals[k] = {
                "anterior2": fc1.number_input("D-2", value=defaults[0], format="%.4f",
                                              key=f"fx_{k}_d2"),
                "anterior":  fc2.number_input("D-1", value=defaults[1], format="%.4f",
                                              key=f"fx_{k}_d1"),
                "atual":     fc3.number_input("D",   value=defaults[2], format="%.4f",
                                              key=f"fx_{k}_d"),
            }
            prev = fx_vals[k]["anterior"]
            curr = fx_vals[k]["atual"]
            var_pct = ((curr - prev) / prev * 100) if prev else 0.0
            fx_vals[k]["variacao"] = f"{var_pct:+.2f}%"
            fc4.metric("Var (%)", fx_vals[k]["variacao"])

        # Extra FX fields for Mercado Cambial slide
        st.markdown("---")
        st.markdown("**Mercado Cambial — Volumes**")
        mc1, mc2 = st.columns(2)
        cambial_vol_usd   = mc1.number_input("Transações (USD Milhões)", value=0.0, format="%.2f",
                                              key="cambial_vol_usd")
        cambial_pos_kz    = mc2.number_input("Posição Cambial (mM Kz)", value=0.0, format="%.2f",
                                              key="cambial_pos_kz")

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: TRANSAÇÕES MERCADO CAMBIAL (T+0, T+1, T+2)
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("🏦 Transações do Mercado Cambial (T+0 / T+1 / T+2)"):
        st.caption("Montante USD, mínimo e máximo de cada liquidação")
        mercado_rows = []
        for liq in ["T+0", "T+1", "T+2"]:
            tc1, tc2, tc3, tc4 = st.columns([1, 2, 2, 2])
            tc1.markdown(f"**{liq}**")
            montante = tc2.number_input("Montante USD", value=0.0, format="%.2f",
                                        key=f"mc_{liq}_mt")
            minimo   = tc3.number_input("Mínimo",       value=0.0, format="%.2f",
                                        key=f"mc_{liq}_mn")
            maximo   = tc4.number_input("Máximo",       value=0.0, format="%.2f",
                                        key=f"mc_{liq}_mx")
            mercado_rows.append({
                "label": liq,
                "montante": f"{montante:,.2f}" if montante else "—",
                "min":      f"{minimo:,.2f}"   if minimo   else "—",
                "max":      f"{maximo:,.2f}"   if maximo   else "—",
            })

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: TRANSAÇÕES BDA (Mercado Cambial)
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("💹 Transações BDA — Mercado Cambial"):
        n_bda_tx = st.number_input("Nº de transações", min_value=0, max_value=20, value=0, step=1,
                                   key="n_bda_tx")
        transacoes_bda = []
        for i in range(int(n_bda_tx)):
            bc1, bc2, bc3, bc4, bc5 = st.columns(5)
            transacoes_bda.append({
                "cv":       bc1.selectbox("C/V", ["C", "V"], key=f"bda_cv_{i}"),
                "par":      bc2.text_input("Par de moeda", key=f"bda_par_{i}"),
                "montante": bc3.number_input("Montante Debt", value=0.0, format="%.2f",
                                             key=f"bda_mt_{i}"),
                "cambio":   bc4.number_input("Câmbio", value=0.0, format="%.4f",
                                             key=f"bda_cam_{i}"),
                "pl":       bc5.number_input("P/L AKZ", value=0.0, format="%.2f",
                                             key=f"bda_pl_{i}"),
            })

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: OPERAÇÕES VIVAS MN (DP / OMA)
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("📋 Operações Vivas MN (DP / OMA / REPO)", expanded=True):
        num_ops_mn = st.number_input("Número de operações MN", min_value=0, max_value=30,
                                     value=0, step=1, key="num_ops_mn")
        operations_mn: list[dict] = []
        for i in range(int(num_ops_mn)):
            with st.expander(f"Operação MN {i + 1}"):
                oc1, oc2, oc3, oc4, oc5, oc6, oc7 = st.columns(7)
                operations_mn.append({
                    "tipo":         oc1.selectbox("Tipo", ["DP", "OMA", "REPO"], key=f"mn_tipo_{i}"),
                    "contraparte":  oc2.text_input("Contraparte", key=f"mn_cp_{i}"),
                    "montante":     f"{oc3.number_input('Montante', key=f'mn_amt_{i}', value=0.0):,.2f}",
                    "taxa":         f"{oc4.number_input('Taxa %', key=f'mn_rate_{i}', value=0.0):.2f}",
                    "residual":     oc5.number_input("Dias Residuais", key=f"mn_res_{i}",
                                                     value=0, step=1),
                    "vencimento":   oc6.text_input("Vencimento", key=f"mn_vec_{i}"),
                    "juro_diario":  f"{oc7.number_input('Juro Diário', key=f'mn_jd_{i}', value=0.0):,.2f}",
                })

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: OPERAÇÕES VIVAS ME
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("💵 Operações Vivas ME"):
        num_ops_me = st.number_input("Número de operações ME", min_value=0, max_value=30,
                                     value=0, step=1, key="num_ops_me")
        operations_me: list[dict] = []
        for i in range(int(num_ops_me)):
            with st.expander(f"Operação ME {i + 1}"):
                mc1, mc2, mc3, mc4, mc5, mc6 = st.columns(6)
                operations_me.append({
                    "contraparte":  mc1.text_input("Contraparte", key=f"me_cp_{i}"),
                    "montante":     f"{mc2.number_input('Montante', key=f'me_amt_{i}', value=0.0):,.2f}",
                    "taxa":         f"{mc3.number_input('Taxa %', key=f'me_rate_{i}', value=0.0):.2f}",
                    "residual":     mc4.number_input("Dias Residuais", key=f"me_res_{i}", value=0, step=1),
                    "vencimento":   mc5.text_input("Vencimento", key=f"me_vec_{i}"),
                    "juro_diario":  f"{mc6.number_input('Juro Diário', key=f'me_jd_{i}', value=0.0):,.2f}",
                })

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: FLUXOS DE CAIXA MN
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("📈 Fluxos de Caixa MN"):
        cf_labels = [
            "Fluxos de Entradas (Cash in flow)",
            "Recebimentos de cupão de títulos",
            "Reembolsos de crédito (+)",
            "Reembolsos de OMA-O/N + Juros",
            "Transferencia a favor conta BNA",
            "Fluxos de Saídas (Cash out flow)",
            "(-) Juros, comissões e outros",
            "Custos Com Pessoal",
            "Fornecimentos e Serviços",
            "Desembolso de crédito (-)",
            "Impostos",
            "Aplicação em OMA",
            "GAP de Liquidez",
        ]
        fluxos_mn_rows = []
        for lbl in cf_labels:
            st.markdown(f"**{lbl}**")
            fc = st.columns(5)
            vals = [fc[j].number_input(day_labels[j], value=0.0, format="%.2f",
                                       key=f"cfmn_{lbl}_{j}",
                                       label_visibility="visible") for j in range(5)]
            fluxos_mn_rows.append({
                "label": lbl,
                "values": [f"{v:,.2f}" if v != 0 else "—" for v in vals],
            })

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: FLUXOS DE CAIXA ME
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("📈 Fluxos de Caixa ME"):
        cf_me_labels = [
            "Fluxos de entradas (Cash in flow)",
            "Outros recebimentos",
            "Reembolsos de DP + Juros",
            "Fluxos de Saídas (Cash out flow)",
            "Aplicação em DP ME",
            "GAP de Liquidez",
        ]
        fluxos_me_rows = []
        for lbl in cf_me_labels:
            st.markdown(f"**{lbl}**")
            fc = st.columns(5)
            vals = [fc[j].number_input(day_labels[j], value=0.0, format="%.2f",
                                       key=f"cfme_{lbl}_{j}",
                                       label_visibility="visible") for j in range(5)]
            fluxos_me_rows.append({
                "label": lbl,
                "values": [f"{v:,.2f}" if v != 0 else "—" for v in vals],
            })

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: P&L CONTROL
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("📊 P&L Control", expanded=True):
        pl_c1, pl_c2, pl_c3 = st.columns(3)
        reembolso_n   = pl_c1.number_input("Reembolso Crédito — Nº Ops", value=0, step=1)
        reembolso_val = pl_c2.number_input("Reembolso Crédito — Montante (Kz)", value=0.0, format="%.2f")
        desembolso_n  = pl_c3.number_input("Desembolso Crédito — Nº Ops", value=0, step=1)
        desembolso_val = pl_c3.number_input("Desembolso Crédito — Montante (Kz)", value=0.0, format="%.2f",
                                            key="desembolso_val")
        fornecedores_n   = pl_c1.number_input("Fornecedores — Nº Ops", value=0, step=1, key="forn_n")
        fornecedores_val = pl_c2.number_input("Fornecedores — Montante (Kz)", value=0.0, format="%.2f",
                                              key="forn_val")

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: KPIs DO SUMÁRIO EXECUTIVO (Slide 3)
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("🎯 KPIs — Sumário Executivo (Slide 3)", expanded=True):
        st.caption("Estes valores aparecem nos cartões KPI do Sumário Executivo. "
                   "A variação (%) é face ao dia anterior.")

        kpi_c1, kpi_c2, kpi_c3 = st.columns(3)

        juros_dp_central = kpi_c1.number_input("Juros de DP — valor central (M Kz)",
                                                value=0.0, format="%.2f", key="juros_dp")
        juros_dp_unit    = kpi_c1.selectbox("Unidade Juros DP", ["M Kz", "mM Kz"], key="juros_dp_unit")

        kpi_renta_mn     = kpi_c2.number_input("Rentabilidade MN (M Kz)", value=0.0, format="%.2f",
                                                key="renta_mn")
        kpi_renta_mn_var = kpi_c2.number_input("Rentabilidade MN — Var (%)", value=0.0, format="%.2f",
                                                key="renta_mn_var")

        kpi_renta_me     = kpi_c3.number_input("Rentabilidade ME (USD)", value=0.0, format="%.2f",
                                                key="renta_me")
        kpi_renta_me_var = kpi_c3.number_input("Rentabilidade ME — Var (%)", value=0.0, format="%.2f",
                                                key="renta_me_var")

        kpi_c4, kpi_c5, kpi_c6 = st.columns(3)
        kpi_renta_tit     = kpi_c4.number_input("Rentabilidade Títulos (M Kz)", value=0.0, format="%.2f",
                                                  key="renta_tit")
        kpi_renta_tit_var = kpi_c4.number_input("Rentabilidade Títulos — Var (%)", value=0.0,
                                                  format="%.2f", key="renta_tit_var")

        kpi_carteira      = kpi_c5.number_input("Carteira Títulos (mM Kz)", value=0.0, format="%.2f",
                                                  key="carteira_tit")
        kpi_carteira_var  = kpi_c5.number_input("Carteira Títulos — Var (%)", value=0.0, format="%.2f",
                                                  key="carteira_tit_var")

        kpi_pos_cambial     = kpi_c6.number_input("Posição Cambial (mM Kz)", value=0.0, format="%.2f",
                                                    key="pos_cambial")
        kpi_pos_cambial_var = kpi_c6.number_input("Posição Cambial — Var (%)", value=0.0, format="%.2f",
                                                    key="pos_cambial_var")

        st.markdown("---")
        enquadramento = st.text_area("Enquadramento (texto livre — aparece no Slide 3)",
                                     height=100, key="enquadramento")
        conclusao     = st.text_area("Breve Conclusão (texto livre — aparece no Slide 3)",
                                     height=80, key="conclusao")

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: BODIVA — MERCADO DE CAPITAIS
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("📈 BODIVA — Segmentos de Mercado (Slide 8)"):
        st.caption("Se não houver dados do scraper, introduza manualmente.")
        bodiva_seg_labels = [
            "Obrigações De Tesouro",
            "Bilhetes Do Tesouro",
            "Obrigações Privadas",
            "Unidades De Participações",
            "Acções",
            "Repos",
            "Total",
        ]
        bodiva_seg_rows = []
        for seg in bodiva_seg_labels:
            bs1, bs2, bs3, bs4 = st.columns([3, 2, 2, 2])
            bs1.markdown(f"**{seg}**")
            ant = bs2.number_input("Anterior", value=0.0, format="%.2f", key=f"bseg_{seg}_ant")
            atu = bs3.number_input("Actual",   value=0.0, format="%.2f", key=f"bseg_{seg}_atu")
            if ant:
                var_pct = (atu - ant) / ant * 100
                var_str = f"{var_pct:+.2f}%"
            else:
                var_str = "—"
            bs4.metric("Var (%)", var_str)
            bodiva_seg_rows.append({
                "segmento": seg,
                "anterior": f"{ant:,.2f}" if ant else "—",
                "atual":    f"{atu:,.2f}" if atu else "—",
                "variacao": var_str,
            })

        bodiva_total_tx = st.number_input("Total Transações BODIVA (mM Kz)", value=0.0,
                                          format="%.2f", key="bodiva_total_tx")

    with st.expander("📊 BODIVA — Acções (Slide 8)"):
        st.caption("Código de cada acção, volumes e preços. Usado na tabela de bolsa.")
        n_stocks = st.number_input("Nº de acções", min_value=0, max_value=20, value=0, step=1,
                                   key="n_stocks")
        bodiva_stocks_input = {}
        for i in range(int(n_stocks)):
            stk1, stk2, stk3, stk4, stk5, stk6 = st.columns(6)
            code    = stk1.text_input("Código", key=f"stk_code_{i}")
            volume  = stk2.number_input("Vol. Trans.", value=0, step=1, key=f"stk_vol_{i}")
            prev_p  = stk3.number_input("Preço Ant.", value=0.0, format="%.2f", key=f"stk_prev_{i}")
            curr_p  = stk4.number_input("Preço Actual", value=0.0, format="%.2f", key=f"stk_curr_{i}")
            cap_bol = stk5.number_input("Cap. Bolsista", value=0.0, format="%.2f", key=f"stk_cap_{i}")
            if code:
                chg = ((curr_p - prev_p) / prev_p * 100) if prev_p else 0.0
                stk6.metric("Var %", f"{chg:+.2f}%")
                bodiva_stocks_input[code] = {
                    "volume":   volume,
                    "previous": f"{prev_p:,.2f}",
                    "current":  f"{curr_p:,.2f}",
                    "change_pct": chg,
                    "cap_bolsista": f"{cap_bol:,.2f}" if cap_bol else "—",
                }

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION: OPERAÇÕES BDA — CARTEIRA DE TÍTULOS (Slide 9)
    # ─────────────────────────────────────────────────────────────────────────
    with st.expander("📂 Carteira de Títulos BDA (Slide 9)"):
        n_cart = st.number_input("Nº de linhas na carteira", min_value=0, max_value=30, value=0,
                                  step=1, key="n_cart")
        carteira_titulos = []
        for i in range(int(n_cart)):
            with st.expander(f"Título {i + 1}"):
                ct1, ct2, ct3, ct4 = st.columns(4)
                ct5, ct6, ct7, ct8, ct9 = st.columns(5)
                carteira_titulos.append({
                    "carteira":    ct1.text_input("Carteira", key=f"ct_cart_{i}"),
                    "cod":         ct2.text_input("Cód. Negociação", key=f"ct_cod_{i}"),
                    "qty_d1":      ct3.number_input("Qtd D-1", value=0, step=1, key=f"ct_q1_{i}"),
                    "qty_d":       ct4.number_input("Qtd D",   value=0, step=1, key=f"ct_qd_{i}"),
                    "nominal":     f"{ct5.number_input('Val. Nominal', value=0.0, format='%.2f', key=f'ct_nom_{i}'):,.2f}",
                    "taxa":        f"{ct6.number_input('Taxa (%)', value=0.0, format='%.4f', key=f'ct_taxa_{i}'):.4f}",
                    "montante":    f"{ct7.number_input('Montante D', value=0.0, format='%.2f', key=f'ct_mt_{i}'):,.2f}",
                    "juros_anual": f"{ct8.number_input('Juros Anual', value=0.0, format='%.2f', key=f'ct_ja_{i}'):,.2f}",
                    "juro_diario": f"{ct9.number_input('Juro Diário D', value=0.0, format='%.2f', key=f'ct_jd_{i}'):,.2f}",
                })
        bodiva_transacoes_valor = st.number_input("Transações Totais (mM Kz)", value=0.0,
                                                   format="%.2f", key="bodiva_tx_val")
        bodiva_juros_diario     = st.number_input("Juros Diário BDA (M Kz)", value=0.0,
                                                   format="%.2f", key="bodiva_jd")

    with st.expander("📑 Operações BODIVA — Transações BDA (Slide 9)"):
        n_bodiva_ops = st.number_input("Nº de operações BODIVA", min_value=0, max_value=30,
                                        value=0, step=1, key="n_bodiva_ops")
        bodiva_operacoes = []
        for i in range(int(n_bodiva_ops)):
            bo1, bo2, bo3, bo4, bo5, bo6 = st.columns(6)
            bodiva_operacoes.append({
                "tipo":       bo1.text_input("Tipo de Op.", key=f"bo_tipo_{i}"),
                "data":       bo2.text_input("Data Contrat.", key=f"bo_data_{i}"),
                "cv":         bo3.selectbox("C/V", ["C", "V"], key=f"bo_cv_{i}"),
                "preco":      f"{bo4.number_input('Preço', value=0.0, format='%.2f', key=f'bo_preco_{i}'):,.2f}",
                "quantidade": bo5.number_input("Quantidades", value=0, step=1, key=f"bo_qty_{i}"),
                "montante":   f"{bo6.number_input('Montante', value=0.0, format='%.2f', key=f'bo_mt_{i}'):,.2f}",
            })

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SAVE BUTTON
    # ─────────────────────────────────────────────────────────────────────────
    if st.button("💾 Guardar Dados Internos", type="primary"):

        # Compute day-indexed totals for MN
        total_mn_by_day = [
            lmn_vals["reservas"][d] + lmn_vals["do_com"][d] +
            lmn_vals["dp_com"][d]   + lmn_vals["omas"][d]
            for d in range(5)
        ]
        total_me_by_day = [
            lme_vals["saldo_do"][d] + lme_vals["dps_me"][d] + lme_vals["colateral"][d]
            for d in range(5)
        ]

        def _fmt(v):
            return f"{v:,.2f}" if v != 0 else "—"

        # Build liquidez_mn_rows (label + 5 day values)
        lmn_rows_out = []
        for key, label in lmn_fields:
            lmn_rows_out.append({
                "label": label,
                "values": [_fmt(lmn_vals[key][d]) for d in range(5)],
            })
        lmn_rows_out.append({
            "label": "LIQUIDEZ BDA",
            "values": [_fmt(total_mn_by_day[d]) for d in range(5)],
        })

        lme_rows_out = []
        for key, label in lme_fields:
            lme_rows_out.append({
                "label": label,
                "values": [_fmt(lme_vals[key][d]) for d in range(5)],
            })
        lme_rows_out.append({
            "label": "LIQUIDEZ BDA",
            "values": [_fmt(total_me_by_day[d]) for d in range(5)],
        })

        # Current day (D) totals
        total_mn = total_mn_by_day[4]
        total_me = total_me_by_day[4]

        # Build cambial_rows
        cambial_rows_out = []
        for pair in fx_pairs:
            k = pair.replace("/", "_")
            cambial_rows_out.append({
                "par":       pair,
                "anterior2": _fmt(fx_vals[k]["anterior2"]),
                "anterior":  _fmt(fx_vals[k]["anterior"]),
                "atual":     _fmt(fx_vals[k]["atual"]),
                "variacao":  fx_vals[k]["variacao"],
            })

        # Build full KPI list for Sumário Executivo
        def _var_str(v):
            if v == 0:
                return ""
            return f"{v:+.2f}%"

        kpis_out = [
            {"label": "Liquidez MN (mM Kz)",    "value": _fmt(total_mn / 1000),        "variation_str": ""},
            {"label": "Liquidez ME (M USD)",     "value": _fmt(total_me),               "variation_str": ""},
            {"label": "Posição Cambial (mM Kz)", "value": _fmt(kpi_pos_cambial),        "variation_str": _var_str(kpi_pos_cambial_var)},
            {"label": "Carteira Títulos (mM Kz)","value": _fmt(kpi_carteira),           "variation_str": _var_str(kpi_carteira_var)},
            {"label": "Rentabilidade MN (M Kz)", "value": _fmt(kpi_renta_mn),           "variation_str": _var_str(kpi_renta_mn_var)},
            {"label": "Rentabilidade ME (USD)",   "value": _fmt(kpi_renta_me),           "variation_str": _var_str(kpi_renta_me_var)},
            {"label": "Rentabilidade Títulos",    "value": _fmt(kpi_renta_tit),          "variation_str": _var_str(kpi_renta_tit_var)},
            {"label": "Reembolsos (M Kz)",        "value": _fmt(reembolso_val / 1e6 if reembolso_val > 1e4 else reembolso_val), "variation_str": ""},
        ]

        # Juros de DP — central KPI
        juros_dp_str = f"{juros_dp_central:,.2f} {juros_dp_unit}"

        st.session_state.internal_data = {
            # Liquidity
            "liquidez_mn_rows":  lmn_rows_out,
            "liquidez_me_rows":  lme_rows_out,
            "liquidez_mn_days":  day_labels,

            # FX
            "cambial": {
                "usd_akz":        _fmt(fx_vals["USD_AKZ"]["atual"]),
                "eur_akz":        _fmt(fx_vals["EUR_AKZ"]["atual"]),
                "eur_usd":        _fmt(fx_vals["EUR_USD"]["atual"]),
                "vol_total_usd":  f"{cambial_vol_usd:,.2f} M USD" if cambial_vol_usd else "—",
                "posicao_cambial": f"{cambial_pos_kz:,.2f} mM Kz" if cambial_pos_kz else "—",
            },
            "cambial_rows": cambial_rows_out,

            # Cambial transactions
            "mercado_rows":      mercado_rows,
            "transacoes_bda_rows": [
                {
                    "cv":       r["cv"],
                    "par":      r["par"],
                    "montante": _fmt(r["montante"]),
                    "cambio":   _fmt(r["cambio"]),
                    "pl":       _fmt(r["pl"]),
                }
                for r in transacoes_bda
            ],

            # Operations
            "operacoes_vivas":    operations_mn,
            "operacoes_vivas_me": operations_me,

            # Cash flows
            "fluxos_mn_rows": fluxos_mn_rows,
            "fluxos_me_rows": fluxos_me_rows,

            # P&L
            "pl_summary": [
                {"label": "Reembolso de Crédito", "n_ops": reembolso_n,  "montante": _fmt(reembolso_val)},
                {"label": "Fornecedores",          "n_ops": fornecedores_n, "montante": _fmt(fornecedores_val)},
                {"label": "Desembolso de Crédito", "n_ops": desembolso_n, "montante": _fmt(desembolso_val)},
            ],

            # KPIs
            "kpis":             kpis_out,
            "reembolso_credito": juros_dp_str,
            "enquadramento":    enquadramento,
            "conclusao":        conclusao,

            # Juros diários (slide 4/6 ovals)
            "juros_diario_mn": f"Kz {juros_diario_mn:,.2f}M",
            "juros_diario_me": f"USD {juros_diario_me:,.0f}",

            # BODIVA
            "bodiva_segment_rows":     bodiva_seg_rows,
            "bodiva_stocks":           bodiva_stocks_input,
            "bodiva_total_transacoes": f"{bodiva_total_tx:,.2f} mM Kz" if bodiva_total_tx else "—",

            # Operações BDA (slide 9)
            "carteira_titulos":         carteira_titulos,
            "bodiva_operacoes":         bodiva_operacoes,
            "bodiva_transacoes_valor":  f"{bodiva_transacoes_valor:,.2f} mM Kz" if bodiva_transacoes_valor else "0,00 mM Kz",
            "bodiva_juros_diario":      f"{bodiva_juros_diario:,.2f} M Kz" if bodiva_juros_diario else "—",
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
                    pd.DataFrame([
                        {"Descrição": r["label"], **{d: r["values"][i]
                         for i, d in enumerate(["D-4","D-3","D-2","D-1","D"])}}
                        for r in mn_rows
                    ]),
                    use_container_width=True,
                )
        with c_me:
            st.write("**Moeda Estrangeira (USD)**")
            me_rows = id_.get("liquidez_me_rows", [])
            if me_rows:
                st.dataframe(
                    pd.DataFrame([
                        {"Descrição": r["label"], **{d: r["values"][i]
                         for i, d in enumerate(["D-4","D-3","D-2","D-1","D"])}}
                        for r in me_rows
                    ]),
                    use_container_width=True,
                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — GENERATE & DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header("Gerar Relatório")

    report_date = st.date_input("Data do Relatório", datetime.now())
    provider_label = get_provider_name().upper()

    col_ai, col_hint = st.columns(2)
    with col_ai:
        use_ai = st.checkbox(
            f"Usar Resumos IA ({provider_label})",
            value=False,
            help="Requer configuração do provider LLM no ficheiro .env",
        )
    with col_hint:
        st.info(
            "💡 Para activar a IA: configure `LLM_PROVIDER` e a chave do provider "
            "(`OPENAI_API_KEY` ou `GEMINI_API_KEY`) no ficheiro `.env`."
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

                    # ── AI summaries ──────────────────────────────────────────
                    summaries: dict = {}
                    if use_ai and ed:
                        agent = DailyReportAgent()
                        if ed.get("markets") is not None:
                            summaries["cm_commentary"] = agent.summarize_markets(ed["markets"])
                        if ed.get("crypto") is not None:
                            summaries["crypto_commentary"] = agent.summarize_crypto(ed["crypto"])
                        if ed.get("commodities") is not None:
                            summaries["commodities_commentary"] = agent.summarize_commodities(ed["commodities"])

                    # ── Map external DataFrames → row dicts ───────────────────
                    def _first_str_col(df):
                        if df is None or df.empty:
                            return None
                        for col in df.columns:
                            if df[col].dtype == object:
                                return col
                        return df.columns[0]

                    def _df_to_rows(df, key_name):
                        if df is None or df.empty:
                            return []
                        nc = _first_str_col(df)
                        rows = []
                        for _, r in df.iterrows():
                            rows.append({
                                key_name:   str(r.get(nc, "—")),
                                "anterior": str(r.get("Anterior", "—")),
                                "atual":    str(r.get("Atual",    "—")),
                                "variacao": str(r.get("Var (%)",  "—")),
                            })
                        return rows

                    markets_df     = ed.get("markets")     if ed else None
                    commodities_df = ed.get("commodities") if ed else None
                    crypto_df      = ed.get("crypto")      if ed else None

                    market_info = {
                        "capital_markets":        _df_to_rows(markets_df, "indice"),
                        "crypto":                 _df_to_rows(crypto_df, "moeda"),
                        "commodities":            _df_to_rows(commodities_df, "nome"),
                        "minerais":               [],
                        "cm_commentary":          summaries.get("cm_commentary", ""),
                        "crypto_commentary":      summaries.get("crypto_commentary", ""),
                        "commodities_commentary": summaries.get("commodities_commentary", ""),
                        "minerais_commentary":    summaries.get("minerais_commentary", ""),
                    }

                    # ── LUIBOR ────────────────────────────────────────────────
                    luibor: dict = {}
                    luibor_var: dict = {}
                    luibor_df = ed.get("luibor") if ed else None
                    if luibor_df is not None and not luibor_df.empty:
                        for _, r in luibor_df.iterrows():
                            mat  = str(r.get("Maturidade", r.get(luibor_df.columns[0], "")))
                            rate = str(r.get("Taxa (%)", "—"))
                            luibor[mat]     = rate
                            luibor_var[mat] = str(r.get("Var (%)", "—"))

                    # ── Merge everything ──────────────────────────────────────
                    data = {
                        "report_date":  date_str,
                        "market_info":  market_info,
                        "luibor":       luibor,
                        "luibor_variation": luibor_var,
                        **(id_ or {}),
                    }

                    # ── Generate PPTX ─────────────────────────────────────────
                    os.makedirs("output", exist_ok=True)
                    date_tag  = report_date.strftime("%Y%m%d")
                    pptx_path = f"output/BDA_Report_{date_tag}.pptx"
                    BDAReportGenerator(data).build(pptx_path)
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
