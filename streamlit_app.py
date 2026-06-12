# -*- coding: utf-8 -*-
"""
streamlit_app.py — MF Viagens e Hotéis
Execute localmente: streamlit run streamlit_app.py
"""
import json, uuid, base64, time
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date as _date, time as _time
from pathlib import Path

from core import (
    load_cfg, save_cfg, backup_auto,
    days, calc_totais, salvar_registro, gerar_canhoto,
    totais_painel, exportar_excel_bytes, enviar_email,
    hash_senha, verificar_senha, cifrar, decifrar, fmt_brl,
    TIPOS_PASSAGEM, TIPOS_TRANSPORTE,
)
from db import db_load, db_save, db_delete, db_migrar_json, db_disponivel, db_load_cfg, db_save_cfg, db_clear, db_diagnostico

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="MF Viagens e Hotéis",
    page_icon="✈",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND      = "#4A7A9B"
BRAND_DARK = "#3A6282"

# ── Splash screen ────────────────────────────────────────────────────────────
def _show_splash():
    logo_path = Path(__file__).parent / "LOGO 13.jpeg"
    logo_b64  = ""
    if logo_path.exists():
        with open(logo_path, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()

    logo_html = (
        f'<img src="data:image/jpeg;base64,{logo_b64}" '
        f'style="width:200px;height:200px;object-fit:contain;'
        f'border-radius:16px;margin-bottom:32px;'
        f'box-shadow:0 8px 32px rgba(0,0,0,0.3);">'
        if logo_b64 else
        '<div style="font-size:72px;margin-bottom:32px;">✈</div>'
    )

    st.markdown(f"""
    <style>
    #splash-overlay {{
        position: fixed; inset: 0; z-index: 9999;
        background: linear-gradient(135deg, #1a2f45 0%, #4A7A9B 50%, #1a2f45 100%);
        display: flex; flex-direction: column;
        align-items: center; justify-content: center;
        animation: fadeIn 0.8s ease;
    }}
    @keyframes fadeIn {{
        from {{ opacity: 0; transform: scale(0.95); }}
        to   {{ opacity: 1; transform: scale(1); }}
    }}
    @keyframes pulse {{
        0%, 100% {{ transform: scale(1); }}
        50%       {{ transform: scale(1.04); }}
    }}
    .splash-logo {{ animation: pulse 2.5s ease-in-out infinite; }}
    .splash-title {{
        color: white; font-size: 32px; font-weight: 800;
        letter-spacing: 3px; margin-bottom: 8px;
        text-shadow: 0 2px 12px rgba(0,0,0,0.4);
    }}
    .splash-sub {{
        color: rgba(255,255,255,0.7); font-size: 15px;
        letter-spacing: 5px; margin-bottom: 48px;
    }}
    .splash-btn {{
        background: white; color: #4A7A9B;
        border: none; padding: 14px 48px;
        font-size: 16px; font-weight: 700;
        border-radius: 50px; cursor: pointer;
        letter-spacing: 2px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.25);
        transition: all 0.2s;
    }}
    .splash-btn:hover {{
        transform: translateY(-2px);
        box-shadow: 0 8px 28px rgba(0,0,0,0.35);
    }}
    </style>

    <div id="splash-overlay">
        <div class="splash-logo">{logo_html}</div>
        <div class="splash-title">MF VIAGENS E HOTÉIS</div>
        <div class="splash-sub">SISTEMA DE REGISTRO</div>
    </div>
    """, unsafe_allow_html=True)

    time.sleep(2.5)
    st.session_state["splash_shown"] = True
    st.rerun()

# ── Inicialização do estado ───────────────────────────────────────────────────
def _carregar_cfg() -> dict:
    """Carrega config: Supabase tem prioridade sobre JSON local."""
    local = load_cfg()
    if db_disponivel():
        db_cfg = db_load_cfg()
        if db_cfg:
            local.update(db_cfg)
    return local

def _salvar_cfg(cfg: dict):
    """Salva config localmente e no Supabase se disponível."""
    from core import save_cfg
    save_cfg(cfg)           # JSON local (pode falhar em cloud, ok)
    db_save_cfg(cfg)        # Supabase (persistência real em cloud)
    st.session_state["cfg"] = cfg

def _init_state():
    defaults = {
        "passagens":    [],
        "transportes":  [],
        "materiais":    [],
        "cfg_senha":    "",
        "cfg_unlocked": False,
        "cfg":          _carregar_cfg(),
        "dark_mode":    True,
        "splash_shown": False,
        "confirm_clear":   False,
        "edit_id":         None,
        "edit_passagens":  [],
        "edit_transportes":[],
        "edit_materiais":  [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# Mostra splash apenas na primeira abertura da sessão
if not st.session_state["splash_shown"]:
    _show_splash()
cfg        = st.session_state["cfg"]
dark       = st.session_state["dark_mode"]
BG         = "#1E293B" if dark else "#F0F4F8"
SURFACE    = "#0F172A" if dark else "#FFFFFF"
TEXT       = "#F1F5F9" if dark else "#1A2530"
TEXT_DIM   = "#94A3B8" if dark else "#5A6978"
CHART_BG   = "#1E293B" if dark else "#F0F4F8"
CHART_TPL  = "plotly_dark" if dark else "plotly_white"
CHART_FONT = "#F1F5F9" if dark else "#1A2530"
CHART_GRID = "#334155" if dark else "#DDE4EA"
SIDEBAR_BG = "#0F172A" if dark else "#FFFFFF"
INPUT_BG   = "#0F172A" if dark else "#FFFFFF"
BORDER_C   = "#334155" if dark else "#D1D9E0"

backup_auto(cfg)

# ── CSS dinâmico (dark / light) ───────────────────────────────────────────────
st.markdown(f"""
<style>
    /* ── Fundo geral ── */
    .stApp, .main {{
        background-color: {BG} !important;
    }}
    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] > div {{
        background-color: {SIDEBAR_BG} !important;
    }}

    /* ── Textos globais ── */
    .stApp, .stApp p, .stApp label, .stApp span,
    .stApp div, .stApp h1, .stApp h2, .stApp h3,
    .stApp li, .stApp a, .stMarkdown {{
        color: {TEXT} !important;
    }}

    /* ── Inputs, selects, textareas ── */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div,
    .stMultiSelect > div > div {{
        background-color: {INPUT_BG} !important;
        color: {TEXT} !important;
        border: 1.5px solid #4A7A9B !important;
        border-radius: 6px !important;
        outline: none !important;
    }}
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus,
    .stNumberInput > div > div > input:focus {{
        border: 1.5px solid #7AAECB !important;
        box-shadow: 0 0 0 2px rgba(74,122,155,0.35) !important;
    }}

    /* ── Date e Time input ── */
    .stDateInput > div,
    .stDateInput > div > div,
    .stDateInput input,
    .stTimeInput > div,
    .stTimeInput > div > div,
    .stTimeInput input {{
        background-color: {INPUT_BG} !important;
        color: {TEXT} !important;
        border: 1.5px solid #4A7A9B !important;
        border-radius: 6px !important;
    }}
    .stDateInput input:focus,
    .stTimeInput input:focus {{
        border: 1.5px solid #7AAECB !important;
        box-shadow: 0 0 0 2px rgba(74,122,155,0.35) !important;
    }}

    /* ── Expanders ── */
    details, details > summary,
    [data-testid="stExpander"] {{
        background-color: {SURFACE} !important;
        border-color: {BORDER_C} !important;
    }}
    [data-testid="stExpander"] summary p {{
        color: {TEXT} !important;
    }}

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {{
        background-color: {SURFACE} !important;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {TEXT_DIM} !important;
    }}
    .stTabs [aria-selected="true"] {{
        color: {TEXT} !important;
        border-bottom-color: {BRAND} !important;
    }}

    /* ── Métricas ── */
    [data-testid="metric-container"] {{
        background-color: {SURFACE} !important;
        border-radius: 8px;
        padding: 8px 12px;
        border: 1px solid {BORDER_C};
    }}
    [data-testid="metric-container"] label,
    [data-testid="metric-container"] div {{
        color: {TEXT} !important;
    }}

    /* ── Dataframe / tabela ── */
    .stDataFrame, iframe {{
        background-color: {SURFACE} !important;
    }}

    /* ── Formulários ── */
    [data-testid="stForm"] {{
        background-color: {SURFACE} !important;
        border: 1px solid {BORDER_C} !important;
        border-radius: 8px;
        padding: 12px;
    }}

    /* ── Botões ── */
    .stButton > button {{
        border-color: {BORDER_C} !important;
        color: {TEXT} !important;
        background-color: {SURFACE} !important;
    }}
    .stButton > button[kind="primary"] {{
        background-color: {BRAND} !important;
        color: white !important;
        border-color: {BRAND} !important;
    }}

    /* ── Info / Warning / Success ── */
    [data-testid="stAlert"] {{
        background-color: {SURFACE} !important;
        border-color: {BORDER_C} !important;
    }}

    /* ── Scrollbar ── */
    ::-webkit-scrollbar {{ width: 6px; height: 6px; }}
    ::-webkit-scrollbar-track {{ background: {BG}; }}
    ::-webkit-scrollbar-thumb {{ background: #4A7A9B; border-radius: 4px; }}

    /* ── Header principal ── */
    .main-header {{
        background: linear-gradient(90deg, #1A3A5A 0%, {BRAND} 50%, #1A3A5A 100%);
        color: white;
        padding: 16px 24px;
        border-radius: 8px;
        margin-bottom: 20px;
        text-align: center;
        box-shadow: 0 2px 12px rgba(0,0,0,0.4);
    }}

    /* ── Cards do painel ── */
    .painel-card {{
        border-radius: 12px;
        padding: 18px 16px 14px 16px;
        color: white;
        margin-bottom: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.35);
    }}
    .painel-card .label {{
        font-size: 12px; opacity: 0.85;
        margin-bottom: 6px; text-transform: uppercase;
    }}
    .painel-card .valor {{
        font-size: 22px; font-weight: 700;
    }}

    div[data-testid="stSidebarNav"] {{ display: none; }}
</style>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
logo_path = Path(__file__).parent / "LOGO 13.jpeg"
if logo_path.exists():
    st.sidebar.image(str(logo_path), width=160)
else:
    st.sidebar.markdown("## ✈ MF Viagens")

st.sidebar.markdown("---")
pagina = st.sidebar.radio(
    "Navegação",
    ["📋 Novo Registro", "📜 Histórico", "📊 Painel", "⚙ Configurações"],
    label_visibility="collapsed",
)
st.sidebar.markdown("---")

dados_all = db_load()
total_acc = sum(d.get("val_total", d.get("valor",0)) for d in dados_all)
st.sidebar.markdown(f"**{len(dados_all)}** viagens registradas")
st.sidebar.markdown(f"**{fmt_brl(total_acc)}** total acumulado")

# Toggle dark/light
st.sidebar.markdown("---")
tema_label = "☀️ Modo Claro" if dark else "🌙 Modo Escuro"
if st.sidebar.button(tema_label, use_container_width=True):
    st.session_state["dark_mode"] = not dark
    st.rerun()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h2 style="margin:0; font-size:22px; color:white !important;">✈  MF VIAGENS E HOTÉIS</h2>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# PÁGINA: NOVO REGISTRO
# =============================================================================
if pagina == "📋 Novo Registro":
    st.subheader("Novo Registro de Viagem")

    # ── Seção A: Dados básicos ────────────────────────────────────────────────
    with st.expander("A  Dados da Viagem", expanded=True):
        c1, c2 = st.columns(2)
        nome    = c1.text_input("Nome do Viajante *", key="r_nome")
        destino = c2.text_input("Destino *", key="r_dest")

        c1, c2, c3, c4 = st.columns([2, 1, 2, 1])
        _d_ida  = c1.date_input("📅 Partida *", value=None, key="r_ida",
                                 format="DD/MM/YYYY", min_value=_date(2000,1,1))
        _t_ida  = c2.time_input("🕐 Hora", value=_time(8, 0), key="r_h_ida", step=300)
        _d_vta  = c3.date_input("📅 Retorno *", value=None, key="r_volta",
                                 format="DD/MM/YYYY", min_value=_date(2000,1,1))
        _t_vta  = c4.time_input("🕐 Hora", value=_time(18, 0), key="r_h_vta", step=300)

        ida   = _d_ida.strftime("%d/%m/%Y") if _d_ida else ""
        h_ida = _t_ida.strftime("%H:%M")    if _t_ida else ""
        volta = _d_vta.strftime("%d/%m/%Y") if _d_vta else ""
        h_vta = _t_vta.strftime("%H:%M")    if _t_vta else ""

        d = days(ida, volta)
        vd = cfg.get("valor_diaria", 150.0)
        if d:
            st.info(f"**{d} dia(s)** — Diárias: **{fmt_brl(d * vd)}**  (R$ {vd:.2f}/dia)")

    # ── Seção B: Hospedagem ───────────────────────────────────────────────────
    with st.expander("B  Hospedagem"):
        hotel_op = st.radio("Selecione", ["Sem hospedagem","Opção 1","Opção 2","Opção 3"],
                            horizontal=True, key="r_hotel_op",
                            label_visibility="collapsed")
        hotel_data = {}
        if hotel_op != "Sem hospedagem":
            c1, c2, c3 = st.columns([2,3,1])
            hotel_nome  = c1.text_input("Nome do hotel", key="r_hn")
            hotel_link  = c2.text_input("Link", key="r_hl")
            hotel_noite = c3.number_input("R$/noite", min_value=0.0, step=10.0, key="r_hv")
            hotel_data  = {"selecionado":"Sim","nome":hotel_nome,
                           "link":hotel_link,"valor_noite":hotel_noite}
        else:
            hotel_data = {"selecionado":"Nao","nome":"","link":"","valor_noite":0.0}

    # ── Secao C: Passagens
    # Injeta sugestão no session_state ANTES do form renderizar,
    # assim o Streamlit respeita o valor mesmo que a chave já exista.
    if _d_ida and st.session_state.get("pdt") != _d_ida:
        st.session_state["pdt"] = _d_ida
    if _t_ida and st.session_state.get("phr") != _t_ida:
        st.session_state["phr"] = _t_ida
    if _d_vta and st.session_state.get("pdt_v") != _d_vta:
        st.session_state["pdt_v"] = _d_vta
    if _t_vta and st.session_state.get("phr_v") != _t_vta:
        st.session_state["phr_v"] = _t_vta

    with st.expander(f"C  Passagens  ({len(st.session_state['passagens'])} adicionadas)"):

        ida_vta = st.checkbox("✈ Ida e Volta", key="p_ida_vta")

        if _d_ida or _d_vta:
            st.caption("💡 Datas sugeridas com base nos dados da viagem — ajuste se necessário.")

        with st.form("form_passagem", clear_on_submit=True):
            c1, c2 = st.columns([1, 3])
            pt = c1.selectbox("Tipo", TIPOS_PASSAGEM, key="pt")
            tr = c2.text_input("Trecho ida (ex: BSB > GRU)", key="ptrecho")

            st.markdown("**✈ Ida**")
            c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
            _pdt  = c1.date_input("📅 Data", key="pdt",
                                   format="DD/MM/YYYY", min_value=_date(2000,1,1))
            _phr  = c2.time_input("🕐 Hora", key="phr", step=300)
            vl    = c3.number_input("Valor R$", min_value=0.0, step=10.0, key="pvl")
            lc    = c4.text_input("Localizador", key="plc")
            dt    = _pdt.strftime("%d/%m/%Y") if _pdt else ""
            hr    = _phr.strftime("%H:%M")    if _phr else ""

            tr_v = dt_v = hr_v = lc_v = ""
            vl_v = 0.0
            if ida_vta:
                st.markdown("**↩ Volta**")
                c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
                tr_v   = c1.text_input("Trecho volta (ex: GRU > BSB)", key="ptrecho_v")
                _pdt_v = c2.date_input("📅 Data", key="pdt_v",
                                        format="DD/MM/YYYY", min_value=_date(2000,1,1))
                _phr_v = c3.time_input("🕐 Hora", key="phr_v", step=300)
                vl_v   = c4.number_input("Valor R$", min_value=0.0, step=10.0, key="pvl_v")
                lc_v   = st.text_input("Localizador volta", key="plc_v")
                dt_v   = _pdt_v.strftime("%d/%m/%Y") if _pdt_v else ""
                hr_v   = _phr_v.strftime("%H:%M")    if _phr_v else ""

            if st.form_submit_button("➕ Adicionar Passagem"):
                if tr and dt:
                    st.session_state["passagens"].append({
                        "tipo":pt, "trecho":tr, "data":dt, "hora":hr,
                        "valor":vl, "localizador":lc, "link":"", "sentido":"Ida"
                    })
                    if ida_vta and tr_v and dt_v:
                        st.session_state["passagens"].append({
                            "tipo":pt, "trecho":tr_v, "data":dt_v, "hora":hr_v,
                            "valor":vl_v, "localizador":lc_v, "link":"", "sentido":"Volta"
                        })
                    st.rerun()

        for i, p in enumerate(st.session_state["passagens"]):
            icon    = "✈" if p.get("sentido","") != "Volta" else "↩"
            sentido = p.get("sentido","Ida")
            cc1, cc2 = st.columns([9,1])
            cc1.markdown(
                f"{icon} **[{p['tipo']}]** `{sentido}` "
                f"{p['trecho']} — {p['data']} {p.get('hora','')} — "
                f"**{fmt_brl(p['valor'])}**"
                + (f" | Loc: {p['localizador']}" if p.get('localizador') else "")
            )
            if cc2.button("🗑", key=f"dpass{i}"):
                st.session_state["passagens"].pop(i); st.rerun()
                st.session_state["passagens"].pop(i); st.rerun()

    # ── Seção D: Transporte ───────────────────────────────────────────────────
    with st.expander(f"D  Transporte Local  ({len(st.session_state['transportes'])} adicionados)"):
        with st.form("form_transp", clear_on_submit=True):
            c1,c2,c3,c4 = st.columns([1,1,3,1])
            _tdt = c1.date_input("📅 Data", value=None, key="tdt",
                                  format="DD/MM/YYYY", min_value=_date(2000,1,1))
            ttp  = c2.selectbox("Tipo", TIPOS_TRANSPORTE, key="ttp")
            tdc  = c3.text_input("Descrição", key="tdc")
            tvl  = c4.number_input("Valor R$", min_value=0.0, step=1.0, key="tvl")
            tdt  = _tdt.strftime("%d/%m/%Y") if _tdt else ""
            if st.form_submit_button("➕ Adicionar Transporte"):
                st.session_state["transportes"].append({
                    "data":tdt,"tipo":ttp,"descricao":tdc,"valor":tvl,"anexo":""
                })
                st.rerun()

        for i, t in enumerate(st.session_state["transportes"]):
            cc1, cc2 = st.columns([9,1])
            cc1.markdown(f"🚗 {t['data']} **[{t['tipo']}]** {t['descricao']} — **{fmt_brl(t['valor'])}**")
            if cc2.button("🗑", key=f"dtransp{i}"):
                st.session_state["transportes"].pop(i); st.rerun()

    # ── Seção E: Materiais ────────────────────────────────────────────────────
    with st.expander(f"E  Materiais e Despesas  ({len(st.session_state['materiais'])} adicionados)"):
        with st.form("form_mat", clear_on_submit=True):
            c1,c2,c3 = st.columns([3,1,1])
            mdc = c1.text_input("Descrição", key="mdc")
            mqt = c2.number_input("Qtd", min_value=1, step=1, key="mqt")
            mvl = c3.number_input("Valor unit R$", min_value=0.0, step=1.0, key="mvl")
            if st.form_submit_button("➕ Adicionar Material"):
                st.session_state["materiais"].append({
                    "descricao":mdc,"qtd":mqt,"valor_unit":mvl,"total":mqt*mvl,"nf":""
                })
                st.rerun()

        for i, m in enumerate(st.session_state["materiais"]):
            cc1, cc2 = st.columns([9,1])
            cc1.markdown(f"📦 {m['descricao']} x{m['qtd']} — **{fmt_brl(m['total'])}**")
            if cc2.button("🗑", key=f"dmat{i}"):
                st.session_state["materiais"].pop(i); st.rerun()

    # ── Seção F: Resumo ───────────────────────────────────────────────────────
    with st.expander("F  Resumo de Custos", expanded=True):
        reg_tmp = {
            "ida":ida,"volta":volta,"hotel":hotel_data,
            "passagens":st.session_state["passagens"],
            "transportes":st.session_state["transportes"],
            "materiais":st.session_state["materiais"],
        }
        t = calc_totais(reg_tmp, cfg)
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Diárias",    fmt_brl(t["val_d"]))
        c2.metric("Passagens",  fmt_brl(t["val_p"]))
        c3.metric("Transporte", fmt_brl(t["val_t"]))
        c4.metric("Hospedagem", fmt_brl(t["val_h"]))
        c5.metric("Materiais",  fmt_brl(t["val_m"]))
        c6.metric("**TOTAL**",  fmt_brl(t["total"]))

    # ── PIX e Observações ─────────────────────────────────────────────────────
    c1, c2 = st.columns([1,2])
    pix = c1.text_input("Chave PIX *", key="r_pix")
    obs = c2.text_area("Observações", key="r_obs", height=80)

    # ── Botões ────────────────────────────────────────────────────────────────
    st.markdown("---")
    col1, col2, col3 = st.columns([2,1,1])
    gerar = col1.button("✅ Gerar Canhoto e Salvar", type="primary", use_container_width=True)
    limpar = col3.button("🧹 Limpar formulário", use_container_width=True)

    if limpar:
        for k in ["r_nome","r_dest","r_ida","r_h_ida","r_volta","r_h_vta",
                  "r_pix","r_obs","r_hotel_op"]:
            if k in st.session_state: del st.session_state[k]
        st.session_state["passagens"]   = []
        st.session_state["transportes"] = []
        st.session_state["materiais"]   = []
        st.rerun()

    if gerar:
        erros = []
        if not nome:           erros.append("Nome do viajante")
        if not destino:        erros.append("Destino")
        if not _d_ida:         erros.append("Data de partida")
        if not _d_vta:         erros.append("Data de retorno")
        if not pix:            erros.append("Chave PIX")
        if ida and volta and not days(ida, volta): erros.append("Data de retorno anterior à partida")
        if erros:
            st.error(f"Preencha: {', '.join(erros)}")
        else:
            reg = {
                "nome":nome,"destino":destino,
                "ida":ida,"hora_ida": h_ida if h_ida != "HH:MM" else "",
                "volta":volta,"hora_volta": h_vta if h_vta != "HH:MM" else "",
                "hotel":hotel_data,
                "passagens":  list(st.session_state["passagens"]),
                "transportes":list(st.session_state["transportes"]),
                "materiais":  list(st.session_state["materiais"]),
                "pix":pix,"obs":obs,
            }
            reg = salvar_registro(reg, cfg)   # calcula totais
            db_save(reg)                        # persiste no banco
            canhoto = gerar_canhoto(reg, cfg)
            st.success(f"✅ Viagem salva! Total: {fmt_brl(reg['val_total'])}")
            st.text_area("Canhoto gerado:", canhoto, height=400)
            st.download_button("📋 Baixar canhoto (.txt)", canhoto,
                               file_name=f"canhoto_{nome.replace(' ','_')}.txt")

            # ── Envio automático de e-mail ────────────────────────────────────
            if cfg.get("auto_email") and cfg.get("email_remetente") and cfg.get("email_destinatario"):
                try:
                    _salt = cfg.get("cfg_salt","")
                    _pwd_plain = st.session_state.get("cfg_senha","")
                    email_pwd = (
                        decifrar(cfg.get("email_senha",""), _pwd_plain, _salt)
                        if _salt and cfg.get("email_senha") else
                        cfg.get("email_senha","")
                    )
                    assunto = f"Canhoto de Viagem — {nome} → {destino}"
                    with st.spinner("Enviando e-mail..."):
                        ok_mail, msg_mail = enviar_email(cfg, assunto, canhoto, email_pwd)
                    if ok_mail:
                        st.success(f"📧 {msg_mail}")
                    else:
                        st.warning(f"⚠️ E-mail não enviado: {msg_mail}")
                except Exception as ex:
                    st.warning(f"⚠️ Erro ao enviar e-mail: {ex}")

# =============================================================================
# PÁGINA: HISTÓRICO
# =============================================================================
elif pagina == "📜 Histórico":
    st.subheader("Histórico de Viagens")

    dados = db_load()
    if not dados:
        st.info("Nenhuma viagem registrada ainda.")
    else:
        c1, c2, c3 = st.columns([2,2,1])
        busca_nome = c1.text_input("Buscar por nome", "")
        busca_dest = c2.text_input("Buscar por destino", "")
        anos = ["Todos"] + sorted({d.get("ida","")[-4:]
                for d in dados if len(d.get("ida",""))==10}, reverse=True)
        ano_sel = c3.selectbox("Ano", anos)

        res = dados
        if busca_nome: res = [d for d in res if busca_nome.lower() in d.get("nome","").lower()]
        if busca_dest: res = [d for d in res if busca_dest.lower() in d.get("destino","").lower()]
        if ano_sel != "Todos": res = [d for d in res if d.get("ida","")[-4:] == ano_sel]

        total = sum(d.get("val_total", d.get("valor",0)) for d in res)
        st.markdown(f"**{len(res)} registro(s)** — Total: **{fmt_brl(total)}**")

        # ── Tabela com botões de ação ─────────────────────────────────────────
        for d in res:
            h   = d.get("hotel",{})
            hn  = h.get("nome","") if isinstance(h,dict) and h.get("selecionado")=="Sim" else "—"
            vt  = d.get("val_total", d.get("valor",0))
            rid = d.get("id","")
            with st.container():
                c1,c2,c3,c4,c5,c6,c7 = st.columns([3,2,1,1,1,1,1])
                c1.markdown(f"**{d.get('nome','')}**  \n📍 {d.get('destino','')}")
                c2.markdown(f"📅 {d.get('ida','')} → {d.get('volta','')}")
                c3.markdown(f"🏨 {hn}")
                c4.markdown(f"{d.get('dias',0)}d")
                c5.markdown(f"**{fmt_brl(vt)}**")
                if c6.button("✏️", key=f"edit_{rid}", help="Editar viagem"):
                    st.session_state["edit_id"]          = rid
                    st.session_state["edit_passagens"]   = list(d.get("passagens",[]))
                    st.session_state["edit_transportes"] = list(d.get("transportes",[]))
                    st.session_state["edit_materiais"]   = list(d.get("materiais",[]))
                    st.rerun()
                if c7.button("🗑️", key=f"del_{rid}", help="Excluir viagem"):
                    db_delete(rid)
                    st.rerun()
            st.markdown("<hr style='margin:4px 0; border-color:#334155;'>", unsafe_allow_html=True)

        xls = exportar_excel_bytes(res)
        if xls:
            st.download_button("📥 Exportar Excel",
                               data=xls,
                               file_name=f"viagens_{datetime.now().strftime('%Y%m%d')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── Formulário de edição ──────────────────────────────────────────────────
    edit_id = st.session_state.get("edit_id")
    if edit_id:
        reg_orig = next((d for d in db_load() if d.get("id") == edit_id), None)
        if not reg_orig:
            st.session_state["edit_id"] = None
        else:
            st.markdown("---")
            st.subheader(f"✏️ Editando: {reg_orig.get('nome','')} → {reg_orig.get('destino','')}")

            def _parse_date(s):
                try:    return datetime.strptime(s, "%d/%m/%Y").date()
                except: return None
            def _parse_time(s):
                try:    return datetime.strptime(s, "%H:%M").time()
                except: return _time(8,0)

            with st.expander("A  Dados da Viagem", expanded=True):
                ec1, ec2 = st.columns(2)
                e_nome    = ec1.text_input("Nome do Viajante *", value=reg_orig.get("nome",""), key="e_nome")
                e_destino = ec2.text_input("Destino *", value=reg_orig.get("destino",""), key="e_dest")
                ec1,ec2,ec3,ec4 = st.columns([2,1,2,1])
                _ed_ida = ec1.date_input("📅 Partida *", value=_parse_date(reg_orig.get("ida","")),
                                          key="e_ida", format="DD/MM/YYYY", min_value=_date(2000,1,1))
                _et_ida = ec2.time_input("🕐 Hora", value=_parse_time(reg_orig.get("hora_ida","08:00")),
                                          key="e_h_ida", step=300)
                _ed_vta = ec3.date_input("📅 Retorno *", value=_parse_date(reg_orig.get("volta","")),
                                          key="e_vta", format="DD/MM/YYYY", min_value=_date(2000,1,1))
                _et_vta = ec4.time_input("🕐 Hora", value=_parse_time(reg_orig.get("hora_volta","18:00")),
                                          key="e_h_vta", step=300)
                e_ida   = _ed_ida.strftime("%d/%m/%Y") if _ed_ida else ""
                e_h_ida = _et_ida.strftime("%H:%M")    if _et_ida else ""
                e_volta = _ed_vta.strftime("%d/%m/%Y") if _ed_vta else ""
                e_h_vta = _et_vta.strftime("%H:%M")    if _et_vta else ""
                d_calc  = days(e_ida, e_volta)
                if d_calc:
                    st.info(f"**{d_calc} dia(s)** — Diárias: **{fmt_brl(d_calc * cfg.get('valor_diaria',150))}**")

            with st.expander("B  Hospedagem"):
                h_orig = reg_orig.get("hotel",{})
                h_sel  = h_orig.get("selecionado","Nao") if isinstance(h_orig,dict) else "Nao"
                opts   = ["Sem hospedagem","Opção 1","Opção 2","Opção 3"]
                idx    = 0 if h_sel != "Sim" else 1
                e_hotel_op = st.radio("Hotel", opts, index=idx, horizontal=True,
                                       key="e_hotel_op", label_visibility="collapsed")
                if e_hotel_op != "Sem hospedagem":
                    eh1,eh2,eh3 = st.columns([2,3,1])
                    e_hn  = eh1.text_input("Nome", value=h_orig.get("nome","") if isinstance(h_orig,dict) else "", key="e_hn")
                    e_hl  = eh2.text_input("Link", value=h_orig.get("link","") if isinstance(h_orig,dict) else "", key="e_hl")
                    e_hv  = eh3.number_input("R$/noite", value=float(h_orig.get("valor_noite",0) if isinstance(h_orig,dict) else 0),
                                              min_value=0.0, step=10.0, key="e_hv")
                    e_hotel = {"selecionado":"Sim","nome":e_hn,"link":e_hl,"valor_noite":e_hv}
                else:
                    e_hotel = {"selecionado":"Nao","nome":"","link":"","valor_noite":0.0}

            with st.expander(f"C  Passagens  ({len(st.session_state['edit_passagens'])} adicionadas)"):
                with st.form("form_edit_pass", clear_on_submit=True):
                    ep1,ep2 = st.columns([1,3])
                    ept = ep1.selectbox("Tipo", TIPOS_PASSAGEM, key="ept")
                    etr = ep2.text_input("Trecho", key="eptrecho")
                    ep1,ep2,ep3,ep4 = st.columns([2,1,1,1])
                    _epdt = ep1.date_input("📅 Data", value=_parse_date(reg_orig.get("ida","")),
                                            key="epdt", format="DD/MM/YYYY", min_value=_date(2000,1,1))
                    _ephr = ep2.time_input("🕐 Hora", value=_time(8,0), key="ephr", step=300)
                    epvl  = ep3.number_input("Valor R$", min_value=0.0, step=10.0, key="epvl")
                    eplc  = ep4.text_input("Localizador", key="eplc")
                    if st.form_submit_button("➕ Adicionar Passagem"):
                        if etr and _epdt:
                            st.session_state["edit_passagens"].append({
                                "tipo":ept,"trecho":etr,
                                "data":_epdt.strftime("%d/%m/%Y"),"hora":_ephr.strftime("%H:%M"),
                                "valor":epvl,"localizador":eplc,"link":"","sentido":"Ida"
                            })
                            st.rerun()
                for i,p in enumerate(st.session_state["edit_passagens"]):
                    pc1,pc2 = st.columns([9,1])
                    sentido = p.get("sentido","Ida")
                    pc1.markdown(f"✈ **[{p['tipo']}]** `{sentido}` {p['trecho']} — {p['data']} {p.get('hora','')} — **{fmt_brl(p['valor'])}**")
                    if pc2.button("🗑", key=f"edpass{i}"):
                        st.session_state["edit_passagens"].pop(i); st.rerun()

            with st.expander(f"D  Transporte  ({len(st.session_state['edit_transportes'])} adicionados)"):
                with st.form("form_edit_transp", clear_on_submit=True):
                    et1,et2,et3,et4 = st.columns([1,1,3,1])
                    _etdt = et1.date_input("📅 Data", value=None, key="etdt",
                                            format="DD/MM/YYYY", min_value=_date(2000,1,1))
                    ettp  = et2.selectbox("Tipo", TIPOS_TRANSPORTE, key="ettp")
                    etdc  = et3.text_input("Descrição", key="etdc")
                    etvl  = et4.number_input("Valor R$", min_value=0.0, step=1.0, key="etvl")
                    if st.form_submit_button("➕ Adicionar Transporte"):
                        st.session_state["edit_transportes"].append({
                            "data":_etdt.strftime("%d/%m/%Y") if _etdt else "",
                            "tipo":ettp,"descricao":etdc,"valor":etvl,"anexo":""
                        })
                        st.rerun()
                for i,t in enumerate(st.session_state["edit_transportes"]):
                    tc1,tc2 = st.columns([9,1])
                    tc1.markdown(f"🚗 {t['data']} **[{t['tipo']}]** {t['descricao']} — **{fmt_brl(t['valor'])}**")
                    if tc2.button("🗑", key=f"edtransp{i}"):
                        st.session_state["edit_transportes"].pop(i); st.rerun()

            with st.expander(f"E  Materiais  ({len(st.session_state['edit_materiais'])} adicionados)"):
                with st.form("form_edit_mat", clear_on_submit=True):
                    em1,em2,em3 = st.columns([3,1,1])
                    emdc = em1.text_input("Descrição", key="emdc")
                    emqt = em2.number_input("Qtd", min_value=1, step=1, key="emqt")
                    emvl = em3.number_input("Valor unit R$", min_value=0.0, step=1.0, key="emvl")
                    if st.form_submit_button("➕ Adicionar Material"):
                        st.session_state["edit_materiais"].append({
                            "descricao":emdc,"qtd":emqt,"valor_unit":emvl,"total":emqt*emvl,"nf":""
                        })
                        st.rerun()
                for i,m in enumerate(st.session_state["edit_materiais"]):
                    mc1,mc2 = st.columns([9,1])
                    mc1.markdown(f"📦 {m['descricao']} x{m['qtd']} — **{fmt_brl(m['total'])}**")
                    if mc2.button("🗑", key=f"edmat{i}"):
                        st.session_state["edit_materiais"].pop(i); st.rerun()

            ec1, ec2 = st.columns([1,2])
            e_pix = ec1.text_input("Chave PIX *", value=reg_orig.get("pix",""), key="e_pix")
            e_obs = ec2.text_area("Observações", value=reg_orig.get("obs",""), key="e_obs", height=80)

            # Resumo live
            reg_tmp = {"ida":e_ida,"volta":e_volta,"hotel":e_hotel,
                       "passagens":st.session_state["edit_passagens"],
                       "transportes":st.session_state["edit_transportes"],
                       "materiais":st.session_state["edit_materiais"]}
            t_ed = calc_totais(reg_tmp, cfg)
            et1,et2,et3,et4,et5,et6 = st.columns(6)
            et1.metric("Diárias",    fmt_brl(t_ed["val_d"]))
            et2.metric("Passagens",  fmt_brl(t_ed["val_p"]))
            et3.metric("Transporte", fmt_brl(t_ed["val_t"]))
            et4.metric("Hospedagem", fmt_brl(t_ed["val_h"]))
            et5.metric("Materiais",  fmt_brl(t_ed["val_m"]))
            et6.metric("**TOTAL**",  fmt_brl(t_ed["total"]))

            st.markdown("---")
            bc1, bc2 = st.columns([2,1])
            salvar_edicao  = bc1.button("💾 Salvar alterações", type="primary", use_container_width=True)
            cancelar_edicao= bc2.button("❌ Cancelar edição", use_container_width=True)

            if cancelar_edicao:
                st.session_state["edit_id"] = None
                st.rerun()

            if salvar_edicao:
                erros_ed = []
                if not e_nome:    erros_ed.append("Nome")
                if not e_destino: erros_ed.append("Destino")
                if not e_ida:     erros_ed.append("Data de partida")
                if not e_volta:   erros_ed.append("Data de retorno")
                if not e_pix:     erros_ed.append("Chave PIX")
                if erros_ed:
                    st.error(f"Preencha: {', '.join(erros_ed)}")
                else:
                    reg_ed = {
                        **reg_orig,
                        "nome":e_nome, "destino":e_destino,
                        "ida":e_ida, "hora_ida":e_h_ida,
                        "volta":e_volta, "hora_volta":e_h_vta,
                        "hotel":e_hotel,
                        "passagens":  list(st.session_state["edit_passagens"]),
                        "transportes":list(st.session_state["edit_transportes"]),
                        "materiais":  list(st.session_state["edit_materiais"]),
                        "pix":e_pix, "obs":e_obs,
                    }
                    reg_ed = salvar_registro(reg_ed, cfg)
                    db_save(reg_ed)
                    st.session_state["edit_id"] = None
                    st.success(f"✅ Viagem atualizada! Total: {fmt_brl(reg_ed['val_total'])}")
                    st.rerun()

# =============================================================================
# PÁGINA: PAINEL
# =============================================================================
elif pagina == "📊 Painel":

    # CSS dos cards coloridos
    st.markdown("""
    <style>
    .painel-card {
        border-radius: 12px;
        padding: 18px 16px 14px 16px;
        color: white;
        margin-bottom: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.18);
    }
    .painel-card .label {
        font-size: 12px;
        opacity: 0.85;
        margin-bottom: 6px;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }
    .painel-card .valor {
        font-size: 22px;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    </style>
    """, unsafe_allow_html=True)

    dados = db_load()
    anos  = ["Todos"] + sorted({d.get("ida","")[-4:]
             for d in dados if len(d.get("ida",""))==10}, reverse=True)

    col_titulo, col_filtro = st.columns([3,1])
    col_titulo.subheader("📊 Painel de Gastos")
    ano_sel = col_filtro.selectbox("Ano", anos, key="painel_ano", label_visibility="collapsed")
    if ano_sel != "Todos":
        dados = [d for d in dados if d.get("ida","")[-4:] == ano_sel]

    t = totais_painel(dados)

    # Cards coloridos via HTML
    cards = [
        ("✈  Diárias",     fmt_brl(t["diarias"]),    "#4A7A9B"),
        ("🎫  Passagens",   fmt_brl(t["passagens"]),  "#2E8B6E"),
        ("🚗  Transporte",  fmt_brl(t["transporte"]), "#C07A2E"),
        ("🏨  Hospedagem",  fmt_brl(t["hospedagem"]), "#7B3DAF"),
        ("📦  Materiais",   fmt_brl(t["materiais"]),  "#AF3D3D"),
        (f"💰  Total  ({t['viagens']} viagens)", fmt_brl(t["total"]), "#1A3A5A"),
    ]
    cols = st.columns(6)
    for col, (label, valor, cor) in zip(cols, cards):
        col.markdown(f"""
        <div class="painel-card" style="background:{cor};">
            <div class="label">{label}</div>
            <div class="valor">{valor}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if not dados:
        st.info("Sem dados para exibir.")
    else:
        # CHART_BG definido globalmente como CHART_BG

        c1, c2 = st.columns(2)

        # Gráfico 1 — Por categoria
        with c1:
            df_cat = pd.DataFrame([
                {"Categoria":"Diárias",    "Valor":t["diarias"],    "Cor":"#4A7A9B"},
                {"Categoria":"Passagens",  "Valor":t["passagens"],  "Cor":"#2E8B6E"},
                {"Categoria":"Transporte", "Valor":t["transporte"], "Cor":"#C07A2E"},
                {"Categoria":"Hospedagem", "Valor":t["hospedagem"], "Cor":"#7B3DAF"},
                {"Categoria":"Materiais",  "Valor":t["materiais"],  "Cor":"#AF3D3D"},
            ])
            fig = px.bar(
                df_cat, x="Categoria", y="Valor", color="Categoria",
                color_discrete_map={r["Categoria"]:r["Cor"] for _,r in df_cat.iterrows()},
                title="Gastos por Categoria",
                text_auto=".2s",
            )
            fig.update_layout(
                showlegend=False,
                plot_bgcolor=CHART_BG,
                paper_bgcolor=CHART_BG,
                template=CHART_TPL,
                height=360,
                margin=dict(t=40, b=20, l=20, r=20),
                title_font=dict(size=14, color=CHART_FONT),
                font=dict(color=CHART_FONT),
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            fig.update_yaxes(showgrid=True, gridcolor=CHART_GRID)
            fig.update_xaxes(showgrid=False)
            st.plotly_chart(fig, use_container_width=True)

        # Gráfico 2 — Pizza
        with c2:
            df_pie = df_cat[df_cat["Valor"] > 0]
            fig2 = px.pie(
                df_pie, names="Categoria", values="Valor",
                color="Categoria",
                color_discrete_map={r["Categoria"]:r["Cor"] for _,r in df_pie.iterrows()},
                title="Proporção por Categoria",
                hole=0.4,
            )
            fig2.update_layout(
                paper_bgcolor=CHART_BG,
                height=360,
                margin=dict(t=40, b=20, l=20, r=20),
                title_font=dict(size=14, color=CHART_FONT),
                font=dict(color=CHART_FONT),
                legend=dict(bgcolor=CHART_BG, font=dict(color=CHART_FONT)),
            )
            fig2.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig2, use_container_width=True)

        # Gráfico 3 — Total por viagem (largura total)
        rows_ev = []
        for d in dados:
            rows_ev.append({
                "Viajante": d.get("nome","")[:16],
                "Destino":  d.get("destino","")[:14],
                "Total":    d.get("val_total", d.get("valor",0)),
                "Data":     d.get("ida",""),
            })
        df_ev = pd.DataFrame(rows_ev).sort_values("Data")
        df_ev["Label"] = df_ev["Viajante"] + " / " + df_ev["Destino"]

        fig3 = px.bar(
            df_ev, x="Label", y="Total",
            title="Total por Viagem",
            color="Total",
            color_continuous_scale=["#D6E8F5", BRAND, "#1A3A5A"],
            text_auto=".2s",
        )
        fig3.update_layout(
            plot_bgcolor=CHART_BG,
                paper_bgcolor=CHART_BG,
                template=CHART_TPL,
            height=340,
            margin=dict(t=40, b=60, l=20, r=20),
            title_font=dict(size=14, color=CHART_FONT),
            font=dict(color=CHART_FONT),
            coloraxis_showscale=False,
            xaxis_tickangle=-30,
        )
        fig3.update_traces(textposition="outside", marker_line_width=0)
        fig3.update_yaxes(showgrid=True, gridcolor=CHART_GRID)
        fig3.update_xaxes(showgrid=False)
        st.plotly_chart(fig3, use_container_width=True)

# =============================================================================
# PÁGINA: CONFIGURAÇÕES
# =============================================================================
elif pagina == "⚙ Configurações":

    if not st.session_state["cfg_unlocked"]:
        st.subheader("🔐 Acesso às Configurações")
        has = bool(cfg.get("cfg_hash"))
        with st.form("login_form"):
            usuario = st.text_input("Usuário")
            senha   = st.text_input("Senha", type="password")
            entrar  = st.form_submit_button("Entrar" if has else "Criar acesso")

        if entrar:
            if not usuario or not senha:
                st.error("Preencha usuário e senha.")
            elif not has:
                h, s = hash_senha(usuario + senha)
                cfg["cfg_user"] = usuario
                cfg["cfg_hash"] = h
                cfg["cfg_salt"] = s
                save_cfg(cfg)
                st.session_state["cfg_senha"]    = senha
                st.session_state["cfg_unlocked"] = True
                st.success("Acesso criado!")
                st.rerun()
            else:
                salt = cfg.get("cfg_salt","")
                if salt:
                    ok = verificar_senha(usuario + senha, cfg["cfg_hash"], salt)
                else:
                    import hashlib as _hl
                    ok = _hl.sha256((usuario+senha).strip().encode()).hexdigest()==cfg["cfg_hash"]
                    if ok:
                        h, s = hash_senha(usuario + senha)
                        cfg["cfg_hash"]=h; cfg["cfg_salt"]=s; save_cfg(cfg)
                if ok:
                    st.session_state["cfg_senha"]    = senha
                    st.session_state["cfg_unlocked"] = True
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
    else:
        st.subheader(f"⚙ Configurações  —  usuário: {cfg.get('cfg_user','')}")
        if st.button("🔒 Bloquear"):
            st.session_state["cfg_unlocked"] = False
            st.rerun()

        tab1, tab2, tab3 = st.tabs(["Geral", "E-mail", "Sistema"])

        with tab1:
            nova_diaria = st.number_input("Valor da diária (R$)",
                                          value=float(cfg.get("valor_diaria",150)),
                                          step=10.0)
            bkp_auto = st.checkbox("Backup automático diário",
                                   value=cfg.get("backup_automatico",True))
            if st.button("💾 Salvar configurações gerais", type="primary"):
                cfg["valor_diaria"]      = nova_diaria
                cfg["backup_automatico"] = bkp_auto
                _salvar_cfg(cfg)
                st.success("✅ Salvo com sucesso!")

        with tab2:
            salt = cfg.get("cfg_salt","")
            pwd_atual = decifrar(cfg.get("email_senha",""),
                                 st.session_state.get("cfg_senha",""), salt)
            rem  = st.text_input("Remetente", cfg.get("email_remetente",""))
            pwd  = st.text_input("Senha / App Password", pwd_atual, type="password")
            smtp = st.text_input("SMTP",  cfg.get("email_smtp","smtp.gmail.com"))
            port = st.number_input("Porta", value=int(cfg.get("email_porta",587)), step=1)
            dest = st.text_input("Destinatário", cfg.get("email_destinatario",""))
            ae   = st.checkbox("Enviar e-mail automático ao gerar canhoto",
                               value=cfg.get("auto_email",False))
            st.caption("Gmail: use Senha de App em myaccount.google.com > Senhas de app")

            c1, c2 = st.columns(2)
            if c1.button("💾 Salvar e-mail", type="primary"):
                cfg["email_remetente"]    = rem
                cfg["email_destinatario"] = dest
                cfg["email_smtp"]         = smtp
                cfg["email_porta"]        = int(port)
                cfg["auto_email"]         = ae
                cfg["email_senha"] = (
                    cifrar(pwd, st.session_state.get("cfg_senha",""), salt)
                    if salt and pwd else pwd
                )
                _salvar_cfg(cfg)
                if db_disponivel():
                    st.success("✅ E-mail salvo com sucesso no Supabase!")
                else:
                    st.warning("⚠️ Salvo localmente. Para persistir no cloud, configure os Secrets do Streamlit.")
            if c2.button("📧 Enviar e-mail de teste"):
                ok, msg = enviar_email(cfg, "[TESTE] MF Viagens",
                                       "Configuração funcionando!", pwd)
                st.success(msg) if ok else st.error(msg)

        with tab3:
            from core import DATA_FILE, CONFIG_FILE, BACKUP_DIR
            from db import db_testar

            # ── Status do banco ───────────────────────────────────────────────
            st.markdown("**Banco de dados (Supabase):**")
            if db_disponivel():
                col_sb1, col_sb2 = st.columns([3, 1])
                with col_sb1:
                    if st.button("🔌 Testar conexão Supabase"):
                        with st.spinner("Testando..."):
                            ok_sb, msg_sb = db_testar()
                        if ok_sb:
                            st.success(f"✅ {msg_sb}")
                        else:
                            st.error(f"❌ {msg_sb}")
                    else:
                        st.success("✅ Supabase configurado")
                with col_sb2:
                    if st.button("📤 Migrar JSON → Supabase", type="primary"):
                        n = db_migrar_json()
                        if n:
                            st.success(f"{n} registro(s) migrado(s)!")
                        else:
                            st.warning("Nenhum dado local ou erro na migração.")
            else:
                st.warning("⚠️ Supabase não configurado — usando JSON local")
                with st.expander("🔍 Diagnóstico de conexão", expanded=True):
                    diag = db_diagnostico()
                    st.markdown(f"**httpx instalado:** {'✅' if diag.get('httpx') else '❌'}")
                    st.markdown(f"**Var. ambiente SUPABASE_URL:** {diag.get('env_url','❌')}")
                    st.markdown(f"**Var. ambiente SUPABASE_KEY:** {diag.get('env_key','❌')}")
                    sk = diag.get("secrets_keys", [])
                    st.markdown(f"**Chaves nos Secrets:** `{sk}`")
                    if diag.get("url"):
                        st.markdown(f"**URL lida:** `{diag['url']}`")
                    if diag.get("key_prefix"):
                        st.markdown(f"**Key lida (início):** `{diag['key_prefix']}`")
                    if diag.get("erro"):
                        st.error(f"Erro ao ler secrets: {diag['erro']}")
                    if sk and "SUPABASE_URL" not in sk and "SUPABASE_KEY" not in sk:
                        st.error("❌ As chaves **SUPABASE_URL** e **SUPABASE_KEY** não foram encontradas nos Secrets. Verifique os nomes exatos.")
                with st.expander("Como configurar o Supabase no Streamlit Cloud"):
                    st.markdown("""
1. Acesse **share.streamlit.io** → seu app → **Settings → Secrets**
2. Cole o conteúdo abaixo substituindo pelos seus valores:
```toml
SUPABASE_URL = "https://SEU-PROJETO.supabase.co"
SUPABASE_KEY = "sb_secret_SUA_CHAVE_SECRETA"

[email]
remetente    = "seuemail@gmail.com"
senha        = "SuaSenhaDeApp"
destinatario = "destino@empresa.com.br"
```
3. Clique **Save** e reinicie o app.
                    """)

            st.markdown("---")
            st.markdown("**Arquivos locais:**")
            st.code(f"Dados    : {DATA_FILE}\nConfig   : {CONFIG_FILE}\nBackups  : {BACKUP_DIR}")

            st.markdown("**Dependências:**")
            try:
                import openpyxl; st.success("openpyxl ✅")
            except: st.warning("openpyxl ❌  (pip install openpyxl)")
            try:
                import cryptography; st.success("cryptography ✅")
            except: st.warning("cryptography ❌  (pip install cryptography)")
            try:
                import httpx; st.success(f"httpx ✅  (REST Supabase direto — v{httpx.__version__})")
            except: st.error("httpx ❌  (pip install httpx)")

            # ── Recarregar cfg do banco ───────────────────────────────────────
            st.markdown("---")
            if st.button("🔄 Recarregar configurações do Supabase"):
                novo_cfg = _carregar_cfg()
                st.session_state["cfg"] = novo_cfg
                st.success("Configurações recarregadas!")

            # ── Limpar banco de dados ─────────────────────────────────────────
            st.markdown("---")
            st.markdown("**⚠️ Zona de perigo:**")
            if not st.session_state["confirm_clear"]:
                if st.button("🗑️ Limpar banco de dados", type="secondary"):
                    st.session_state["confirm_clear"] = True
                    st.rerun()
            else:
                st.error(
                    "**ATENÇÃO:** Esta ação apagará **TODOS** os registros de viagens permanentemente. "
                    "Para confirmar, digite sua senha de Administrador abaixo."
                )
                with st.form("form_clear_db"):
                    _pwd_clear = st.text_input("Senha do Administrador", type="password")
                    cc1, cc2 = st.columns(2)
                    confirmar = cc1.form_submit_button("✅ Confirmar exclusão", type="primary")
                    cancelar  = cc2.form_submit_button("❌ Cancelar")

                if cancelar:
                    st.session_state["confirm_clear"] = False
                    st.rerun()

                if confirmar:
                    _salt_c = cfg.get("cfg_salt","")
                    _ok_pwd = False
                    if _salt_c:
                        _ok_pwd = verificar_senha(
                            cfg.get("cfg_user","") + _pwd_clear,
                            cfg["cfg_hash"], _salt_c
                        )
                    else:
                        import hashlib as _hl
                        _ok_pwd = (
                            _hl.sha256((cfg.get("cfg_user","") + _pwd_clear)
                                       .strip().encode()).hexdigest() == cfg.get("cfg_hash","")
                        )
                    if not _ok_pwd:
                        st.error("❌ Senha incorreta. Operação cancelada.")
                    else:
                        with st.spinner("Apagando todos os registros..."):
                            n_apagados = db_clear()
                        if n_apagados >= 0:
                            st.session_state["confirm_clear"] = False
                            st.success(f"✅ Banco limpo! {n_apagados} registro(s) removido(s).")
                            st.rerun()
                        else:
                            st.error("❌ Erro ao limpar o banco. Verifique os logs.")



