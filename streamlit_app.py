# -*- coding: utf-8 -*-
"""
streamlit_app.py — MF Viagens e Hotéis
Execute localmente: streamlit run streamlit_app.py
"""
import json, uuid
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from pathlib import Path

from core import (
    load_data, save_data, load_cfg, save_cfg, backup_auto,
    days, calc_totais, salvar_registro, gerar_canhoto,
    totais_painel, exportar_excel_bytes, enviar_email,
    hash_senha, verificar_senha, cifrar, decifrar, fmt_brl,
    TIPOS_PASSAGEM, TIPOS_TRANSPORTE,
)

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="MF Viagens e Hotéis",
    page_icon="✈",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND = "#4A7A9B"

# ── CSS customizado ───────────────────────────────────────────────────────────
st.markdown(f"""
<style>
    .main-header {{
        background: {BRAND};
        color: white;
        padding: 16px 24px;
        border-radius: 8px;
        margin-bottom: 20px;
        text-align: center;
    }}
    .card {{
        background: white;
        border-radius: 10px;
        padding: 16px;
        border-left: 5px solid {BRAND};
        margin-bottom: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }}
    .stMetric label {{ font-size: 13px !important; }}
    div[data-testid="stSidebarNav"] {{ display: none; }}
</style>
""", unsafe_allow_html=True)

# ── Inicialização do estado ───────────────────────────────────────────────────
def _init_state():
    defaults = {
        "passagens":    [],
        "transportes":  [],
        "materiais":    [],
        "cfg_senha":    "",
        "cfg_unlocked": False,
        "cfg":          load_cfg(),
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()
cfg = st.session_state["cfg"]
backup_auto(cfg)

# ── Credenciais do Streamlit Secrets (opcional) ───────────────────────────────
def _get_from_secrets(key: str, fallback: str = "") -> str:
    try:
        return st.secrets.get(key, fallback)
    except Exception:
        return fallback

# ── Sidebar ───────────────────────────────────────────────────────────────────
logo_path = Path(__file__).parent / "LOGO 13.jpeg"
if logo_path.exists():
    st.sidebar.image(str(logo_path), width=160)
else:
    st.sidebar.markdown(f"## ✈ MF Viagens")

st.sidebar.markdown("---")
pagina = st.sidebar.radio(
    "Navegação",
    ["📋 Novo Registro", "📜 Histórico", "📊 Painel", "⚙ Configurações"],
    label_visibility="collapsed",
)
st.sidebar.markdown("---")

dados_all = load_data()
total_acc = sum(d.get("val_total", d.get("valor",0)) for d in dados_all)
st.sidebar.markdown(f"**{len(dados_all)}** viagens registradas")
st.sidebar.markdown(f"**{fmt_brl(total_acc)}** total acumulado")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h2 style="margin:0; font-size:22px;">✈  MF VIAGENS E HOTÉIS</h2>
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
        c1, c2, c3, c4 = st.columns([2,1,2,1])
        ida     = c1.text_input("Partida (DD/MM/AAAA) *", key="r_ida")
        h_ida   = c2.text_input("Hora", "HH:MM", key="r_h_ida")
        volta   = c3.text_input("Retorno (DD/MM/AAAA) *", key="r_volta")
        h_vta   = c4.text_input("Hora", "HH:MM", key="r_h_vta")

        d = days(ida, volta)
        vd = cfg.get("valor_diaria", 150.0)
        if d:
            st.info(f"**{d} dia(s)** — Diárias: **{fmt_brl(d * vd)}**  (R$ {vd:.2f}/dia)")

    # ── Seção B: Hospedagem ───────────────────────────────────────────────────
    with st.expander("B  Hospedagem"):
        hotel_op = st.radio("", ["Sem hospedagem","Opção 1","Opção 2","Opção 3"],
                            horizontal=True, key="r_hotel_op")
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

    # ── Seção C: Passagens ────────────────────────────────────────────────────
    with st.expander(f"C  Passagens  ({len(st.session_state['passagens'])} adicionadas)"):
        with st.form("form_passagem", clear_on_submit=True):
            c1,c2,c3,c4,c5 = st.columns([1,2,1,1,1])
            pt = c1.selectbox("Tipo", TIPOS_PASSAGEM, key="pt")
            tr = c2.text_input("Trecho (ex: BSB>GRU)", key="ptrecho")
            dt = c3.text_input("Data DD/MM/AAAA", key="pdt")
            hr = c4.text_input("Hora HH:MM", key="phr")
            vl = c5.number_input("Valor R$", min_value=0.0, step=10.0, key="pvl")
            lc = st.text_input("Localizador", key="plc")
            if st.form_submit_button("➕ Adicionar Passagem"):
                st.session_state["passagens"].append({
                    "tipo":pt,"trecho":tr,"data":dt,"hora":hr,
                    "valor":vl,"localizador":lc,"link":"","sentido":"Ida"
                })
                st.rerun()

        for i, p in enumerate(st.session_state["passagens"]):
            cc1, cc2 = st.columns([9,1])
            cc1.markdown(f"✈ **[{p['tipo']}]** {p['trecho']} — {p['data']} {p['hora']} — "
                         f"**{fmt_brl(p['valor'])}** | Loc: {p.get('localizador','')}")
            if cc2.button("🗑", key=f"dpass{i}"):
                st.session_state["passagens"].pop(i); st.rerun()

    # ── Seção D: Transporte ───────────────────────────────────────────────────
    with st.expander(f"D  Transporte Local  ({len(st.session_state['transportes'])} adicionados)"):
        with st.form("form_transp", clear_on_submit=True):
            c1,c2,c3,c4 = st.columns([1,1,3,1])
            tdt = c1.text_input("Data DD/MM/AAAA", key="tdt")
            ttp = c2.selectbox("Tipo", TIPOS_TRANSPORTE, key="ttp")
            tdc = c3.text_input("Descrição", key="tdc")
            tvl = c4.number_input("Valor R$", min_value=0.0, step=1.0, key="tvl")
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
        if not nome:    erros.append("Nome do viajante")
        if not destino: erros.append("Destino")
        if not ida:     erros.append("Data de partida")
        if not volta:   erros.append("Data de retorno")
        if not pix:     erros.append("Chave PIX")
        if not days(ida, volta): erros.append("Datas inválidas")
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
            reg = salvar_registro(reg, cfg)
            canhoto = gerar_canhoto(reg, cfg)
            st.success(f"✅ Viagem salva! Total: {fmt_brl(reg['val_total'])}")
            st.text_area("Canhoto gerado:", canhoto, height=400)
            st.download_button("📋 Baixar canhoto (.txt)", canhoto,
                               file_name=f"canhoto_{nome.replace(' ','_')}.txt")

# =============================================================================
# PÁGINA: HISTÓRICO
# =============================================================================
elif pagina == "📜 Histórico":
    st.subheader("Histórico de Viagens")

    dados = load_data()
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

        rows = []
        for d in res:
            h  = d.get("hotel",{})
            hn = h.get("nome","") if isinstance(h,dict) and h.get("selecionado")=="Sim" else "Não"
            vt = d.get("val_total", d.get("valor",0))
            rows.append({
                "Nome":     d.get("nome",""),
                "Destino":  d.get("destino",""),
                "Partida":  d.get("ida",""),
                "Retorno":  d.get("volta",""),
                "Hotel":    hn,
                "Dias":     d.get("dias",""),
                "Passagens":len(d.get("passagens",[])),
                "Total":    fmt_brl(vt),
            })

        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True)

        xls = exportar_excel_bytes(res)
        if xls:
            st.download_button("📥 Exportar Excel",
                               data=xls,
                               file_name=f"viagens_{datetime.now().strftime('%Y%m%d')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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

    dados = load_data()
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
        BG_CHART = "#F0F4F8"

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
                plot_bgcolor=BG_CHART,
                paper_bgcolor=BG_CHART,
                height=360,
                margin=dict(t=40, b=20, l=20, r=20),
                title_font=dict(size=14, color="#1A2530"),
                font=dict(color="#1A2530"),
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            fig.update_yaxes(showgrid=True, gridcolor="#DDE4EA")
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
                paper_bgcolor=BG_CHART,
                height=360,
                margin=dict(t=40, b=20, l=20, r=20),
                title_font=dict(size=14, color="#1A2530"),
                font=dict(color="#1A2530"),
                legend=dict(bgcolor=BG_CHART),
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
            plot_bgcolor=BG_CHART,
            paper_bgcolor=BG_CHART,
            height=340,
            margin=dict(t=40, b=60, l=20, r=20),
            title_font=dict(size=14, color="#1A2530"),
            font=dict(color="#1A2530"),
            coloraxis_showscale=False,
            xaxis_tickangle=-30,
        )
        fig3.update_traces(textposition="outside", marker_line_width=0)
        fig3.update_yaxes(showgrid=True, gridcolor="#DDE4EA")
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
                save_cfg(cfg)
                st.session_state["cfg"] = cfg
                st.success("Salvo!")

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
                cfg["email_senha"] = cifrar(pwd, st.session_state.get("cfg_senha",""), salt) if salt else pwd
                save_cfg(cfg)
                st.session_state["cfg"] = cfg
                st.success("Salvo!")
            if c2.button("📧 Enviar e-mail de teste"):
                ok, msg = enviar_email(cfg, "[TESTE] MF Viagens",
                                       "Configuração funcionando!", pwd)
                st.success(msg) if ok else st.error(msg)

        with tab3:
            from core import DATA_FILE, CONFIG_FILE, BACKUP_DIR
            st.code(f"Dados    : {DATA_FILE}\nConfig   : {CONFIG_FILE}\nBackups  : {BACKUP_DIR}")
            st.markdown("**Dependências opcionais:**")
            try:
                import openpyxl; st.success("openpyxl ✅")
            except: st.warning("openpyxl ❌  (pip install openpyxl)")
            try:
                import cryptography; st.success("cryptography ✅")
            except: st.warning("cryptography ❌  (pip install cryptography)")
