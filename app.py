"""
Curva ABC — app standalone Tamaru Ferramentas
Execute: streamlit run app.py   (dentro da pasta curva-abc)
"""

import html

import streamlit as st
import pandas as pd

from abc_curva import render_pagina_curva_abc
from data_loader import load_vendas, ultima_atualizacao_vendas

st.set_page_config(
    page_title="Curva ABC | Tamaru",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
[data-testid="stSidebar"] { display: none; }
[data-testid="collapsedControl"] { display: none; }
header[data-testid="stHeader"] {
  background: rgba(255, 255, 255, 0.92);
  border-bottom: 1px solid #e5e7eb;
}
.block-container {
  padding-top: 2.5rem !important;
  padding-bottom: 2rem !important;
  padding-left: 1.75rem !important;
  padding-right: 1.75rem !important;
  max-width: 100% !important;
}
.abc-page-header {
  margin: 0 0 1.25rem 0;
}
.abc-page-header h1 {
  font-size: 1.75rem; font-weight: 700; margin: 0 0 4px 0; color: #111827;
}
.abc-page-header p {
  margin: 0; color: #6b7280; font-size: 0.95rem;
}
.abc-last-update {
  margin-top: 6px !important; font-size: 0.82rem !important; color: #94a3b8 !important;
}
div[data-testid="stVerticalBlockBorderWrapper"] {
  border-color: #e2e8f0 !important;
  border-radius: 10px !important;
  background: #fafbfc;
}
div[data-testid="stVerticalBlockBorderWrapper"] > div {
  padding-top: 0.75rem;
}
.panel-label {
  font-size: 0.72rem; font-weight: 700; letter-spacing: 0.06em;
  text-transform: uppercase; color: #64748b; margin: 0 0 0.35rem 0;
}
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=300)
def carregar_vendas():
    return load_vendas()


@st.cache_data(ttl=60)
def info_ultima_atualizacao():
    dt, nome = ultima_atualizacao_vendas()
    return dt.strftime("%d/%m/%Y %H:%M"), nome


def fmt_brl(v):
    if pd.isna(v):
        return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num(v, decimals=0):
    if pd.isna(v):
        return "—"
    return f"{v:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_mes(s: str) -> str:
    try:
        return pd.to_datetime(s).strftime("%b-%Y")
    except Exception:
        return str(s)


def caixa_insight(itens: list, titulo: str = "💡 Conclusões", tipo: str = "info"):
    cores = {
        "info": ("#1e40af", "#eff6ff", "#bfdbfe"),
        "ok": ("#14532d", "#f0fdf4", "#bbf7d0"),
    }
    cor_texto, cor_fundo, cor_borda = cores.get(tipo, cores["info"])
    bullets = "".join(
        f'<div style="display:flex;gap:8px;margin-bottom:4px;">'
        f'<span style="color:{cor_texto};">•</span>'
        f'<span style="font-size:0.84rem;">{it}</span></div>'
        for it in itens
    )
    st.markdown(
        f'<div style="background:{cor_fundo};border:1.5px solid {cor_borda};'
        f'border-radius:8px;padding:12px 16px;margin:8px 0;">'
        f'<div style="font-weight:700;color:{cor_texto};margin-bottom:6px;">{titulo}</div>{bullets}</div>',
        unsafe_allow_html=True,
    )


try:
    vendas = carregar_vendas()
    ultima_str, ultimo_arquivo = info_ultima_atualizacao()
except FileNotFoundError as err:
    st.error(str(err))
    st.stop()

st.markdown(
    '<div class="abc-page-header">'
    "<h1>📊 Curva ABC</h1>"
    "<p>Layout ERP · níveis hierárquicos · export Excel</p>"
    f'<p class="abc-last-update">Última atualização: <strong>{ultima_str}</strong>'
    f" · arquivo <strong>{html.escape(ultimo_arquivo)}</strong></p>"
    "</div>",
    unsafe_allow_html=True,
)

periodos = sorted(vendas["AnoMes_str"].dropna().unique())
_idx_de = max(0, len(periodos) - 13)

with st.container(border=True):
    st.markdown('<p class="panel-label">Período e dados</p>', unsafe_allow_html=True)
    p1, p2, p3, p4 = st.columns([2, 2, 2, 1])
    with p1:
        periodo_de = st.selectbox("De", periodos, index=_idx_de, key="abc_periodo_de")
    with p2:
        periodo_ate = st.selectbox("Até", periodos, index=len(periodos) - 1, key="abc_periodo_ate")
    with p3:
        vdf_preview = vendas[
            (vendas["AnoMes_str"] >= periodo_de) & (vendas["AnoMes_str"] <= periodo_ate)
        ]
        st.metric(
            "Registros no período",
            f"{len(vdf_preview):,}".replace(",", "."),
            help=f"{fmt_mes(periodo_de)} → {fmt_mes(periodo_ate)}",
        )
    with p4:
        st.markdown("<div style='height:1.6rem'></div>", unsafe_allow_html=True)
        if st.button("🔄 Recarregar", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

vdf = vendas[(vendas["AnoMes_str"] >= periodo_de) & (vendas["AnoMes_str"] <= periodo_ate)]

st.markdown("")

render_pagina_curva_abc(
    st,
    vendas=vendas,
    vdf=vdf,
    fmt_brl=fmt_brl,
    fmt_num=fmt_num,
    fmt_mes=fmt_mes,
    caixa_insight=caixa_insight,
)
