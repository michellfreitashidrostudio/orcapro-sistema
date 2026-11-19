import html
import streamlit as st
import pandas as pd
from datetime import date
from database.db import (
    init_db, SessionLocal, BaseReferencia, Composicao, PerfilBDI, Orcamento,
    OrcamentoItem, Insumo, Cadastro, Compra, EncargoSocial, MemoriaCalculo,
    CronogramaEtapa, DMTRegistro
)
from core.calculations import calcular_bdi_tcu, calcular_valores_item, gerar_curva_abc
from etl.importer import padronizar_planilha, salvar_base_no_banco
from exports.excel_exporter import (
    gerar_excel_orcamento, gerar_excel_bdi,
    gerar_excel_composicoes_proprias, gerar_excel_cronograma
)
from exports.pdf_exporter import gerar_pdf_orcamento

st.set_page_config(page_title="OrçaPro | Engenharia", page_icon="🏗️", layout="wide")

# ==============================================================================
# TELA DE LOGIN (PROTEÇÃO DE ACESSO DA EQUIPE)
# ==============================================================================
def verificar_login():
    if st.session_state.get("autenticado"):
        return True

    # Usuários padrão (para teste local). Na nuvem, você configura nos Secrets do Streamlit.
    try:
        usuarios_validos = dict(st.secrets["USUARIOS"])
    except Exception:
        usuarios_validos = {
            "hidrostudio": "Agsaoisfeu@#5s6",
            "engenharia": "orca2026"
        }

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("<h2 style='text-align:center; color:#0066b3;'>🏗️ OrçaPro</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align:center; color:#64748b;'>Acesso restrito à equipe</p>", unsafe_allow_html=True)
            with st.form("form_login"):
                usuario = st.text_input("Usuário")
                senha = st.text_input("Senha", type="password")
                entrar = st.form_submit_button("Entrar", type="primary", use_container_width=True)
                if entrar:
                    if usuario in usuarios_validos and str(usuarios_validos[usuario]) == senha:
                        st.session_state["autenticado"] = True
                        st.session_state["usuario_logado"] = usuario
                        st.rerun()
                    else:
                        st.error("Usuário ou senha incorretos.")
    return False

if not verificar_login():
    st.stop()

st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #1e242b; }
    [data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    .top-banner {
        background-color: #0066b3; color: white; padding: 12px 20px;
        border-radius: 6px; margin-bottom: 12px; display: flex;
        justify-content: space-between; align-items: center;
    }
    .top-banner h2 { margin: 0; font-size: 1.25rem; color: white; }
    .badge-deson { background: #dcfce7; color: #166534; padding: 4px 12px; border-radius: 14px; font-size: 0.8rem; font-weight: 700; }
    .badge-oner { background: #fef3c7; color: #92400e; padding: 4px 12px; border-radius: 14px; font-size: 0.8rem; font-weight: 700; }
    .kpi-card {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-left: 4px solid #0066b3; border-radius: 6px; padding: 10px 14px;
    }
    .kpi-label { font-size: 0.72rem; color: #64748b; font-weight: 600; text-transform: uppercase; }
    .kpi-value { font-size: 1.2rem; color: #0f172a; font-weight: 700; }
    .kpi-green { border-left-color: #16a34a; background: #f0fdf4; }
    .kpi-green .kpi-value { color: #15803d; }

    .of-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; margin-top: 8px; background: white; }
    .of-table th {
        background: #f1f5f9; color: #334155; font-weight: 700; font-size: 0.75rem;
        text-transform: uppercase; padding: 9px 8px; border-bottom: 2px solid #cbd5e1; text-align: left;
    }
    .of-table td { padding: 8px; border-bottom: 1px solid #e2e8f0; vertical-align: top; color: #1e293b; }
    .of-table .num { text-align: right; white-space: nowrap; }
    .of-table .ctr { text-align: center; white-space: nowrap; }
    .of-etapa td { background: #dbeafe; color: #1e40af; font-weight: 700; border-top: 1px solid #93c5fd; }
    .of-cod { color: #0066b3; font-weight: 700; display: block; }
    .of-fonte { color: #64748b; font-size: 0.72rem; display: block; }
    .of-total td { background: #0066b3; color: white; font-weight: 700; font-size: 0.92rem; padding: 10px 8px; }
</style>
""", unsafe_allow_html=True)

init_db()
db = SessionLocal()

def brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def num_br(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def obter_ou_criar_base_propria(session) -> BaseReferencia:
    base = session.query(BaseReferencia).filter_by(nome="PROPRIA").first()
    if not base:
        base = BaseReferencia(
            nome="PROPRIA", descricao="Composições Personalizadas da Equipe",
            uf="SP", data_base=date.today(), desonerado=False
        )
        session.add(base)
        session.commit()
    return base

def renderizar_tabela_orcafascio(itens_lista) -> str:
    """Gera tabela HTML 100% responsiva no padrão visual da OrçaFascio."""
    linhas_html = []
    df_tmp = pd.DataFrame([{
        "etapa": i.etapa, "item": i.item_eap, "fonte": i.fonte_base, "codigo": i.codigo,
        "desc": i.descricao, "unid": i.unidade, "qtd": i.quantidade,
        "unit": i.preco_sem_bdi, "bdi": i.bdi_aplicado, "unit_bdi": i.preco_com_bdi, "total": i.preco_total
    } for i in itens_lista])

    total_geral = float(df_tmp["total"].sum())

    for etapa, grupo in df_tmp.groupby("etapa", sort=False):
        sub_etapa = float(grupo["total"].sum())
        linhas_html.append(
            f'<tr class="of-etapa"><td colspan="7">{html.escape(str(etapa))}</td>'
            f'<td class="num">{num_br(sub_etapa)}</td></tr>'
        )
        for _, r in grupo.iterrows():
            linhas_html.append(
                f'<tr>'
                f'<td style="width:6%; font-weight:600; color:#475569;">{html.escape(str(r["item"]))}</td>'
                f'<td style="width:12%;"><span class="of-cod">{html.escape(str(r["codigo"]))}</span>'
                f'<span class="of-fonte">{html.escape(str(r["fonte"]))}</span></td>'
                f'<td style="width:40%; line-height:1.35;">{html.escape(str(r["desc"]))}</td>'
                f'<td class="ctr" style="width:6%;">{html.escape(str(r["unid"]))}</td>'
                f'<td class="num" style="width:8%;">{num_br(r["qtd"])}</td>'
                f'<td class="num" style="width:9%;">{num_br(r["unit"])}</td>'
                f'<td class="num" style="width:9%;">{num_br(r["unit_bdi"])}</td>'
                f'<td class="num" style="width:10%; font-weight:700;">{num_br(r["total"])}</td>'
                f'</tr>'
            )

    corpo = "".join(linhas_html)
    return (
        '<table class="of-table">'
        '<thead><tr>'
        '<th>Item</th><th>Código / Fonte</th><th>Descrição</th>'
        '<th class="ctr">Und</th><th class="num">Quant.</th>'
        '<th class="num">Val. Unit.</th><th class="num">Val. c/ BDI</th>'
        '<th class="num">Total (R$)</th>'
        '</tr></thead>'
        f'<tbody>{corpo}'
        f'<tr class="of-total"><td colspan="7" class="num">TOTAL GERAL DO ORÇAMENTO</td>'
        f'<td class="num">{brl(total_geral)}</td></tr>'
        '</tbody></table>'
    )

# ==============================================================================
# BARRA LATERAL
# ==============================================================================
with st.sidebar:
    st.markdown("### 🏗️ ORÇAPRO")
    st.caption("Sistema de Orçamentação")
    st.divider()
    menu = st.radio(
        "Menu",
        [
            "📋 Orçamentos",
            "🛠️ Composições Próprias",
            "🧱 Insumos",
            "🗄️ Gestão de Bases",
            "🛒 Compras",
            "👷 Encargos Sociais",
            "🧮 Perfis de BDI",
            "🗂️ Cadastros"
        ],
        label_visibility="collapsed"
    )
    st.divider()
    st.caption(f"👤 Usuário: **{st.session_state.get('usuario_logado', '')}**")
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

# ==============================================================================
# 1. ORÇAMENTOS
# ==============================================================================
if menu == "📋 Orçamentos":
    orcamentos = db.query(Orcamento).order_by(Orcamento.id.desc()).all()

    c_sel, c_btn = st.columns([3.5, 1])
    with c_sel:
        if orcamentos:
            opcoes_orc = {f"#{o.id} — {o.titulo} ({o.cliente or 'Sem cliente'})": o.id for o in orcamentos}
            orc_label = st.selectbox("Orçamento Ativo", list(opcoes_orc.keys()), label_visibility="collapsed")
            orc = db.query(Orcamento).get(opcoes_orc[orc_label])
        else:
            orc = None
    with c_btn:
        criar_novo = st.toggle("➕ Novo Orçamento", value=(len(orcamentos) == 0))

    if criar_novo:
        with st.container(border=True):
            st.subheader("Novo Orçamento")
            perfis = db.query(PerfilBDI).all()
            with st.form("form_novo_orc"):
                c1, c2, c3 = st.columns(3)
                tit = c1.text_input("Título do Projeto / Obra *")
                cli = c2.text_input("Cliente")
                loc = c3.text_input("Local da Obra", value="São Paulo - SP")
                c4, c5, c6 = st.columns(3)
                bdi_sel = c4.selectbox("BDI Padrão", perfis, format_func=lambda x: f"{x.nome} ({x.percentual_final:.2f}%)")
                reg_sel = c5.selectbox("Encargos Sociais", ["Onerado (Não Desonerado)", "Desonerado"])
                trunc = c6.checkbox("Truncar na 2ª casa (TCU)", value=True)
                if st.form_submit_button("Criar Orçamento", type="primary") and tit:
                    db.add(Orcamento(
                        titulo=tit, cliente=cli, local_obra=loc,
                        bdi_padrao_id=bdi_sel.id, truncar_valores=trunc,
                        desonerado=(reg_sel == "Desonerado")
                    ))
                    db.commit()
                    st.rerun()

    if not orc:
        st.info("Crie o primeiro orçamento acima para começar.")
        st.stop()

    itens = (
        db.query(OrcamentoItem)
        .filter_by(orcamento_id=orc.id)
        .order_by(OrcamentoItem.etapa, OrcamentoItem.item_eap, OrcamentoItem.id)
        .all()
    )

    total_sem_bdi = sum(i.preco_sem_bdi * i.quantidade for i in itens)
    total_com_bdi = sum(i.preco_total for i in itens)
    valor_bdi = total_com_bdi - total_sem_bdi
    bdi_padrao_val = orc.bdi_padrao.percentual_final if orc.bdi_padrao else 25.0
    is_deson = bool(orc.desonerado)

    badge_html = '<span class="badge-deson">🟢 DESONERADO</span>' if is_deson else '<span class="badge-oner">🟠 ONERADO (NÃO DESONERADO)</span>'
    st.markdown(
        f'<div class="top-banner">'
        f'<div><h2>📋 {html.escape(orc.titulo)}</h2>'
        f'<span style="font-size:0.85rem; opacity:0.9;">Cliente: {html.escape(orc.cliente or "Não informado")} | Local: {html.escape(orc.local_obra or "")}</span></div>'
        f'<div>{badge_html} &nbsp; <span class="badge-deson" style="background:#e0f2fe; color:#0369a1;">BDI: {bdi_padrao_val:.2f}%</span></div>'
        f'</div>',
        unsafe_allow_html=True
    )

    c_reg1, c_reg2, c_reg3 = st.columns([2, 2, 2])
    novo_regime = c_reg1.radio(
        "Regime de Encargos do Orçamento:",
        ["Onerado", "Desonerado"],
        index=1 if is_deson else 0,
        horizontal=True
    )
    if (novo_regime == "Desonerado") != is_deson:
        orc.desonerado = (novo_regime == "Desonerado")
        db.commit()
        st.rerun()

    with c_reg2:
        st.write("")
        if itens and st.button(f"🔄 Atualizar preços dos itens para {novo_regime}", help="Busca os mesmos códigos nas bases importadas com o regime selecionado"):
            atualizados = 0
            bases_regime = db.query(BaseReferencia).filter_by(desonerado=orc.desonerado).all()
            ids_bases_reg = [b.id for b in bases_regime]
            if ids_bases_reg:
                for it in itens:
                    comp_nova = (
                        db.query(Composicao)
                        .filter(Composicao.base_id.in_(ids_bases_reg), Composicao.codigo == it.codigo)
                        .first()
                    )
                    if comp_nova:
                        it.preco_sem_bdi = comp_nova.preco_unitario
                        pcb, ptot = calcular_valores_item(it.preco_sem_bdi, it.bdi_aplicado, it.quantidade, orc.truncar_valores)
                        it.preco_com_bdi, it.preco_total = pcb, ptot
                        atualizados += 1
                db.commit()
                st.success(f"{atualizados} item(ns) atualizado(s) para o regime {novo_regime}!")
                st.rerun()
            else:
                st.warning(f"Nenhuma base cadastrada como '{novo_regime}'. Importe a planilha correspondente em 'Gestão de Bases'.")

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div class="kpi-label">Total sem BDI</div><div class="kpi-value">{brl(total_sem_bdi)}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div class="kpi-label">Valor do BDI</div><div class="kpi-value">{brl(valor_bdi)}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card kpi-green"><div class="kpi-label">Total Geral (c/ BDI)</div><div class="kpi-value">{brl(total_com_bdi)}</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div class="kpi-label">Itens no Orçamento</div><div class="kpi-value">{len(itens)} serviços</div></div>', unsafe_allow_html=True)

    st.write("")

    t_plan, t_mem, t_crono, t_dmt, t_rel, t_cfg = st.tabs([
        "📋 Planilha",
        "📐 Memória de Cálculo",
        "📅 Cronograma",
        "🚛 DMT Transporte",
        "📊 Relatórios & Exportação",
        "⚙️ Config"
    ])

    # --- ABA 1: PLANILHA ---
    with t_plan:
        bases_disp = db.query(BaseReferencia).order_by(BaseReferencia.nome).all()
        with st.expander("➕ Inserir Composição / Serviço", expanded=(len(itens) == 0)):
            if not bases_disp:
                st.warning("Importe uma base em 'Gestão de Bases' ou cadastre em 'Composições Próprias'.")
            else:
                opcoes_bases = {
                    f"{b.nome} ({b.uf} - {b.data_base.strftime('%m/%Y')} - {'Deson.' if b.desonerado else 'Oner.'})": b.id
                    for b in bases_disp
                }
                padrao_bases = [lbl for lbl, bid in opcoes_bases.items() if any(b.id == bid and (b.desonerado == is_deson or b.nome == "PROPRIA") for b in bases_disp)]
                if not padrao_bases:
                    padrao_bases = list(opcoes_bases.keys())

                c_b, c_t = st.columns([1.2, 2])
                bases_sel = c_b.multiselect("Bancos de Dados", list(opcoes_bases.keys()), default=padrao_bases)
                termo = c_t.text_input("Buscar por código ou descrição", placeholder="Ex: escavação, concreto, CPU-001...")

                if termo and len(termo) >= 2 and bases_sel:
                    ids_b = [opcoes_bases[l] for l in bases_sel]
                    res = (
                        db.query(Composicao, BaseReferencia)
                        .join(BaseReferencia, Composicao.base_id == BaseReferencia.id)
                        .filter(Composicao.base_id.in_(ids_b))
                        .filter((Composicao.descricao.ilike(f"%{termo}%")) | (Composicao.codigo.ilike(f"%{termo}%")))
                        .limit(40).all()
                    )
                    if res:
                        op_c = {f"[{b.nome}] {c.codigo} — {c.descricao[:90]} ({c.unidade}) | {brl(c.preco_unitario)}": (c, b) for c, b in res}
                        esc = st.selectbox("Serviço encontrado", list(op_c.keys()))
                        c_sel, b_sel = op_c[esc]

                        etapas_ex = sorted(list(set(i.etapa for i in itens))) if itens else ["1. SERVIÇOS PRELIMINARES"]
                        c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 1])
                        et_in = c1.text_input("Etapa", value=etapas_ex[-1])
                        it_in = c2.text_input("Item", value=f"1.{len(itens)+1}")
                        qt_in = c3.number_input("Quant.", min_value=0.0, value=1.0, step=1.0, format="%.2f")
                        bd_in = c4.number_input("BDI (%)", value=float(bdi_padrao_val), step=0.5, format="%.2f")
                        c5.write("")
                        c5.write("")
                        if c5.button("Inserir", type="primary", use_container_width=True):
                            pcb, ptot = calcular_valores_item(c_sel.preco_unitario, bd_in, qt_in, orc.truncar_valores)
                            db.add(OrcamentoItem(
                                orcamento_id=orc.id, etapa=et_in.strip(), item_eap=it_in.strip(),
                                fonte_base=f"{b_sel.nome} {b_sel.data_base.strftime('%m/%Y')}",
                                codigo=c_sel.codigo, descricao=c_sel.descricao, unidade=c_sel.unidade,
                                quantidade=qt_in, preco_sem_bdi=c_sel.preco_unitario,
                                bdi_aplicado=bd_in, preco_com_bdi=pcb, preco_total=ptot
                            ))
                            db.commit()
                            st.rerun()
                    else:
                        st.warning("Nenhum serviço encontrado.")

        if itens:
            st.markdown(renderizar_tabela_orcafascio(itens), unsafe_allow_html=True)
            st.write("")

            with st.expander("✏️ Editar Quantidades, Etapas, BDI ou Excluir Itens"):
                ids_itens = [i.id for i in itens]
                df_ed = pd.DataFrame([{
                    "Excluir": False,
                    "Etapa": i.etapa,
                    "Item": i.item_eap,
                    "Descrição": i.descricao,
                    "Unid.": i.unidade,
                    "Quant.": float(i.quantidade),
                    "Unit. s/ BDI": float(i.preco_sem_bdi),
                    "BDI (%)": float(i.bdi_aplicado),
                } for i in itens])

                df_res = st.data_editor(
                    df_ed,
                    key=f"ed_orc_{orc.id}_{len(itens)}",
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Excluir": st.column_config.CheckboxColumn("🗑️", width="small"),
                        "Etapa": st.column_config.TextColumn("Etapa", width="medium"),
                        "Item": st.column_config.TextColumn("Item", width="small"),
                        "Descrição": st.column_config.TextColumn("Descrição", width="large"),
                        "Unid.": st.column_config.TextColumn("Unid.", width="small"),
                        "Quant.": st.column_config.NumberColumn("Quant.", min_value=0.0, format="%.2f", width="small"),
                        "Unit. s/ BDI": st.column_config.NumberColumn("Unit. s/ BDI", min_value=0.0, format="R$ %.2f", width="small"),
                        "BDI (%)": st.column_config.NumberColumn("BDI (%)", min_value=0.0, format="%.2f%%", width="small"),
                    }
                )

                if st.button("💾 Salvar Alterações na Planilha", type="primary"):
                    for idx, row in df_res.iterrows():
                        item_db = db.query(OrcamentoItem).get(ids_itens[idx])
                        if not item_db:
                            continue
                        if bool(row["Excluir"]):
                            db.delete(item_db)
                        else:
                            item_db.etapa = str(row["Etapa"]).strip()
                            item_db.item_eap = str(row["Item"]).strip()
                            item_db.descricao = str(row["Descrição"]).strip()
                            item_db.unidade = str(row["Unid."]).strip().upper()
                            item_db.quantidade = float(row["Quant."])
                            item_db.preco_sem_bdi = float(row["Unit. s/ BDI"])
                            item_db.bdi_aplicado = float(row["BDI (%)"])
                            pcb, ptot = calcular_valores_item(
                                item_db.preco_sem_bdi, item_db.bdi_aplicado, item_db.quantidade, orc.truncar_valores
                            )
                            item_db.preco_com_bdi = pcb
                            item_db.preco_total = ptot
                    db.commit()
                    st.success("Planilha atualizada!")
                    st.rerun()

    # --- ABA 2: MEMÓRIA DE CÁLCULO ---
    with t_mem:
        st.subheader("Memória de Cálculo Quantitativa (Linkada ao Orçamento)")
        if not itens:
            st.info("Insira serviços na Planilha primeiro.")
        else:
            op_it = {f"{i.item_eap} — {i.descricao[:75]} (Atual: {i.quantidade:.2f} {i.unidade})": i.id for i in itens}
            it_sel_lbl = st.selectbox("Selecione o serviço para detalhar a memória de cálculo", list(op_it.keys()))
            it_obj = db.query(OrcamentoItem).get(op_it[it_sel_lbl])

            with st.form("form_mem", clear_on_submit=True):
                c1, c2, c3, c4, c5 = st.columns([2.5, 1, 1, 1, 1])
                desc_t = c1.text_input("Trecho / Ambiente / Elemento *", placeholder="Ex: Vala Trecho PV1-PV2")
                fat = c2.number_input("Fator/Qtd", value=1.0, step=1.0)
                comp = c3.number_input("Comprimento", value=1.0, step=0.5)
                larg = c4.number_input("Largura", value=1.0, step=0.5)
                alt = c5.number_input("Altura", value=1.0, step=0.5)
                if st.form_submit_button("Adicionar Linha de Cálculo", type="primary") and desc_t:
                    sub = round(fat * comp * larg * alt, 4)
                    db.add(MemoriaCalculo(
                        orcamento_item_id=it_obj.id, descricao_trecho=desc_t,
                        fator=fat, comprimento=comp, largura=larg, altura=alt, subtotal=sub
                    ))
                    db.commit()
                    total_mem = sum(m.subtotal for m in db.query(MemoriaCalculo).filter_by(orcamento_item_id=it_obj.id).all())
                    it_obj.quantidade = total_mem
                    pcb, ptot = calcular_valores_item(it_obj.preco_sem_bdi, it_obj.bdi_aplicado, total_mem, orc.truncar_valores)
                    it_obj.preco_com_bdi, it_obj.preco_total = pcb, ptot
                    db.commit()
                    st.success(f"Quantidade atualizada para {total_mem:.2f} {it_obj.unidade}.")
                    st.rerun()

            mems = db.query(MemoriaCalculo).filter_by(orcamento_item_id=it_obj.id).all()
            if mems:
                ids_m = [m.id for m in mems]
                df_m = pd.DataFrame([{
                    "Excluir": False, "Trecho": m.descricao_trecho, "Fator": m.fator,
                    "Comp.": m.comprimento, "Larg.": m.largura, "Alt.": m.altura, "Subtotal": m.subtotal
                } for m in mems])
                df_m_ed = st.data_editor(df_m, hide_index=True, use_container_width=True, key=f"mem_{it_obj.id}_{len(mems)}")
                if st.button("🔄 Recalcular e Sincronizar com Orçamento", type="primary"):
                    for idx, r in df_m_ed.iterrows():
                        m_db = db.query(MemoriaCalculo).get(ids_m[idx])
                        if bool(r["Excluir"]): db.delete(m_db)
                        else:
                            m_db.descricao_trecho = r["Trecho"]
                            m_db.fator, m_db.comprimento, m_db.largura, m_db.altura = float(r["Fator"]), float(r["Comp."]), float(r["Larg."]), float(r["Alt."])
                            m_db.subtotal = round(m_db.fator * m_db.comprimento * m_db.largura * m_db.altura, 4)
                    db.commit()
                    total_mem = sum(m.subtotal for m in db.query(MemoriaCalculo).filter_by(orcamento_item_id=it_obj.id).all())
                    it_obj.quantidade = total_mem
                    pcb, ptot = calcular_valores_item(it_obj.preco_sem_bdi, it_obj.bdi_aplicado, total_mem, orc.truncar_valores)
                    it_obj.preco_com_bdi, it_obj.preco_total = pcb, ptot
                    db.commit()
                    st.rerun()

    # --- ABA 3: CRONOGRAMA ---
    with t_crono:
        st.subheader("Cronograma Físico-Financeiro (Linkado às Etapas)")
        if not itens:
            st.info("Insira serviços na Planilha para montar o cronograma.")
        else:
            df_it = pd.DataFrame([{"etapa": i.etapa, "total": i.preco_total} for i in itens])
            etapas_totais = df_it.groupby("etapa", sort=False)["total"].sum().to_dict()
            qtd_meses = st.number_input("Prazo da Obra (Meses)", min_value=1, max_value=24, value=6)
            reg_crono = {(c.etapa, c.mes): c.percentual for c in db.query(CronogramaEtapa).filter_by(orcamento_id=orc.id).all()}

            linhas_pct = []
            for et, val_et in etapas_totais.items():
                linha = {"Etapa": et, "Valor Etapa (R$)": round(val_et, 2)}
                for m in range(1, qtd_meses + 1):
                    linha[f"Mês {m} (%)"] = float(reg_crono.get((et, m), 0.0))
                linhas_pct.append(linha)

            df_crono_ed = st.data_editor(
                pd.DataFrame(linhas_pct), use_container_width=True, hide_index=True,
                column_config={"Etapa": st.column_config.TextColumn(disabled=True), "Valor Etapa (R$)": st.column_config.NumberColumn(disabled=True, format="R$ %.2f")}
            )
            if st.button("💾 Salvar Cronograma", type="primary"):
                for _, r in df_crono_ed.iterrows():
                    for m in range(1, qtd_meses + 1):
                        pct = float(r[f"Mês {m} (%)"])
                        ex = db.query(CronogramaEtapa).filter_by(orcamento_id=orc.id, etapa=r["Etapa"], mes=m).first()
                        if ex: ex.percentual = pct
                        else: db.add(CronogramaEtapa(orcamento_id=orc.id, etapa=r["Etapa"], mes=m, percentual=pct))
                db.commit()
                st.rerun()

            linhas_fin = []
            for _, r in df_crono_ed.iterrows():
                val_et = r["Valor Etapa (R$)"]
                soma_pct = sum(float(r[f"Mês {m} (%)"]) for m in range(1, qtd_meses + 1))
                l_fin = {"Etapa": r["Etapa"], "Total Etapa (R$)": val_et, "Alocado (%)": round(soma_pct, 1)}
                for m in range(1, qtd_meses + 1):
                    l_fin[f"Mês {m} (R$)"] = round(val_et * (float(r[f"Mês {m} (%)"]) / 100.0), 2)
                linhas_fin.append(l_fin)
            df_crono_fin = pd.DataFrame(linhas_fin)
            st.subheader("Desembolso Financeiro Previsto (R$)")
            st.dataframe(df_crono_fin, use_container_width=True, hide_index=True)
            st.download_button(
                "📥 Exportar Cronograma (.xlsx)",
                data=gerar_excel_cronograma(orc, df_crono_fin),
                file_name=f"Cronograma_{orc.id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # --- ABA 4: DMT ---
    with t_dmt:
        st.subheader("Cálculo de DMT (Distância Média de Transporte)")
        with st.form("form_dmt", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            mat = c1.text_input("Material / Solo *")
            orig = c2.text_input("Origem")
            dest = c3.text_input("Destino")
            c4, c5, c6, c7, c8 = st.columns(5)
            kp = c4.number_input("Km Pavimentado", min_value=0.0, value=10.0)
            knp = c5.number_input("Km Não Pavim.", min_value=0.0, value=2.0)
            cp = c6.number_input("R$/m³.km Pav.", min_value=0.0, value=1.80)
            cnp = c7.number_input("R$/m³.km Não Pav.", min_value=0.0, value=2.60)
            vol = c8.number_input("Volume (m³ ou t)", min_value=0.0, value=100.0)
            if st.form_submit_button("Registrar DMT", type="primary") and mat:
                db.add(DMTRegistro(
                    orcamento_id=orc.id, material=mat, origem=orig, destino=dest,
                    km_pavimentado=kp, km_nao_pavimentado=knp, custo_km_pav=cp, custo_km_nao_pav=cnp, volume_ton=vol
                ))
                db.commit()
                st.rerun()

        dmts = db.query(DMTRegistro).filter_by(orcamento_id=orc.id).all()
        if dmts:
            st.dataframe(pd.DataFrame([{
                "Material": d.material, "Origem": d.origem, "Destino": d.destino,
                "DMT Total (km)": round(d.km_pavimentado + d.km_nao_pavimentado, 2),
                "Volume": d.volume_ton,
                "Unit. Transporte (R$)": round(d.km_pavimentado * d.custo_km_pav + d.km_nao_pavimentado * d.custo_km_nao_pav, 2),
                "Total Transporte (R$)": round((d.km_pavimentado * d.custo_km_pav + d.km_nao_pavimentado * d.custo_km_nao_pav) * d.volume_ton, 2)
            } for d in dmts]), use_container_width=True, hide_index=True)

    # --- ABA 5: RELATÓRIOS ---
    with t_rel:
        st.subheader("Central de Relatórios e Exportações")
        if not itens:
            st.warning("Adicione itens ao orçamento primeiro.")
        else:
            df_itens_exp = pd.DataFrame([{
                "etapa": i.etapa, "item_eap": i.item_eap, "fonte_base": i.fonte_base,
                "codigo": i.codigo, "descricao": i.descricao, "unidade": i.unidade,
                "quantidade": i.quantidade, "preco_sem_bdi": i.preco_sem_bdi,
                "bdi_aplicado": i.bdi_aplicado, "preco_com_bdi": i.preco_com_bdi, "preco_total": i.preco_total
            } for i in itens])
            df_abc = gerar_curva_abc(df_itens_exp)

            r1, r2, r3, r4 = st.columns(4)
            r1.download_button(
                "📥 Planilha Sintética + ABC (.xlsx)",
                data=gerar_excel_orcamento(orc, df_itens_exp, df_abc),
                file_name=f"Orcamento_{orc.id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True
            )
            r2.download_button(
                "📄 Caderno Orçamentário (.pdf)",
                data=gerar_pdf_orcamento(orc, df_itens_exp, df_abc),
                file_name=f"Orcamento_{orc.id}.pdf",
                mime="application/pdf",
                type="primary", use_container_width=True
            )
            if orc.bdi_padrao:
                r3.download_button(
                    "🧮 Demonstrativo de BDI (.xlsx)",
                    data=gerar_excel_bdi(orc.bdi_padrao),
                    file_name=f"Demonstrativo_BDI_{orc.id}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            base_p = obter_ou_criar_base_propria(db)
            comps_p = db.query(Composicao).filter_by(base_id=base_p.id).all()
            if comps_p:
                r4.download_button(
                    "🛠️ Composições Próprias (.xlsx)",
                    data=gerar_excel_composicoes_proprias(comps_p),
                    file_name="Composicoes_Proprias.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.divider()
            st.subheader("Curva ABC de Serviços")
            st.dataframe(df_abc, use_container_width=True, hide_index=True)

    # --- ABA 6: CONFIG ---
    with t_cfg:
        perfis = db.query(PerfilBDI).all()
        idx_b = next((idx for idx, p in enumerate(perfis) if p.id == orc.bdi_padrao_id), 0)
        with st.form("f_cfg"):
            nt = st.text_input("Título", value=orc.titulo)
            nc = st.text_input("Cliente", value=orc.cliente or "")
            nl = st.text_input("Local", value=orc.local_obra or "")
            nb = st.selectbox("BDI Padrão", perfis, index=idx_b, format_func=lambda x: f"{x.nome} ({x.percentual_final:.2f}%)")
            ntr = st.checkbox("Truncar na 2ª casa (TCU)", value=orc.truncar_valores)
            if st.form_submit_button("Salvar", type="primary"):
                orc.titulo, orc.cliente, orc.local_obra, orc.bdi_padrao_id, orc.truncar_valores = nt, nc, nl, nb.id, ntr
                db.commit()
                st.rerun()
        st.divider()
        if st.button("🗑️ Excluir Orçamento Permanentemente"):
            db.delete(orc)
            db.commit()
            st.rerun()

# ==============================================================================
# 2. COMPOSIÇÕES PRÓPRIAS
# ==============================================================================
elif menu == "🛠️ Composições Próprias":
    st.header("Composições Personalizadas (Base Própria)")
    base_p = obter_ou_criar_base_propria(db)
    comps = db.query(Composicao).filter_by(base_id=base_p.id).order_by(Composicao.codigo).all()

    with st.form("f_cpu", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        cod = c1.text_input("Código *", value=f"CPU-{len(comps)+1:03d}")
        unid = c2.text_input("Unidade *", value="UN")
        prc = c3.number_input("Preço Unitário s/ BDI (R$) *", min_value=0.0, step=10.0)
        desc = st.text_area("Descrição do Serviço *")
        if st.form_submit_button("Cadastrar Composição", type="primary") and cod and desc:
            db.add(Composicao(base_id=base_p.id, codigo=cod.strip(), descricao=desc.strip(), unidade=unid.strip().upper(), preco_unitario=prc))
            db.commit()
            st.rerun()

    if comps:
        ids_c = [c.id for c in comps]
        df_c = pd.DataFrame([{"Excluir": False, "Código": c.codigo, "Descrição": c.descricao, "Unid.": c.unidade, "Preço s/ BDI": c.preco_unitario} for c in comps])
        df_ce = st.data_editor(df_c, use_container_width=True, hide_index=True, key=f"cpu_{len(comps)}")
        col_s, col_e = st.columns([1, 1])
        if col_s.button("💾 Salvar Alterações", type="primary"):
            for idx, r in df_ce.iterrows():
                obj = db.query(Composicao).get(ids_c[idx])
                if bool(r["Excluir"]): db.delete(obj)
                else:
                    obj.codigo, obj.descricao, obj.unidade, obj.preco_unitario = r["Código"], r["Descrição"], r["Unid."], float(r["Preço s/ BDI"])
            db.commit()
            st.rerun()
        col_e.download_button("📥 Exportar Composições Próprias (.xlsx)", data=gerar_excel_composicoes_proprias(comps), file_name="Composicoes_Proprias.xlsx")

# ==============================================================================
# 3. INSUMOS
# ==============================================================================
elif menu == "🧱 Insumos":
    st.header("Banco de Insumos (Materiais, Mão de Obra e Equipamentos)")
    with st.form("f_ins", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        cod = c1.text_input("Código *", value=f"INS-{db.query(Insumo).count()+1:04d}")
        tipo = c2.selectbox("Tipo", ["MATERIAL", "MAO_DE_OBRA", "EQUIPAMENTO", "SERVICO"])
        unid = c3.text_input("Unidade *", value="UN")
        prc = c4.number_input("Preço Unitário (R$)", min_value=0.0, step=1.0)
        desc = st.text_input("Descrição do Insumo *")
        if st.form_submit_button("Cadastrar Insumo", type="primary") and cod and desc:
            db.add(Insumo(codigo=cod.strip(), tipo=tipo, unidade=unid.strip().upper(), preco_unitario=prc, descricao=desc.strip()))
            db.commit()
            st.rerun()

    insumos = db.query(Insumo).order_by(Insumo.id.desc()).all()
    if insumos:
        ids_i = [i.id for i in insumos]
        df_i = pd.DataFrame([{"Excluir": False, "Código": i.codigo, "Tipo": i.tipo, "Descrição": i.descricao, "Unid.": i.unidade, "Preço (R$)": i.preco_unitario} for i in insumos])
        df_ie = st.data_editor(df_i, use_container_width=True, hide_index=True, key=f"ins_{len(insumos)}")
        if st.button("💾 Salvar Alterações nos Insumos", type="primary"):
            for idx, r in df_ie.iterrows():
                obj = db.query(Insumo).get(ids_i[idx])
                if bool(r["Excluir"]): db.delete(obj)
                else:
                    obj.descricao, obj.unidade, obj.preco_unitario = r["Descrição"], r["Unid."], float(r["Preço (R$)"])
            db.commit()
            st.rerun()

# ==============================================================================
# 4. GESTÃO DE BASES
# ==============================================================================
elif menu == "🗄️ Gestão de Bases":
    st.header("Gestão de Bancos de Dados (SINAPI, CDHU, SIURB, SICRO e Genéricas)")
    with st.expander("➕ Importar Nova Planilha", expanded=False):
        c1, c2 = st.columns(2)
        tipo_b = c1.selectbox("Tipo", ["SINAPI", "CDHU", "SIURB", "SICRO", "GENERICA"])
        nome_b = c1.text_input("Nome da Base", value=tipo_b if tipo_b != "GENERICA" else "TABELA PROPRIA")
        desc_b = c1.text_input("Descrição", value=f"{nome_b} - {date.today().strftime('%m/%Y')}")
        uf_b = c1.text_input("UF", value="SP", max_chars=2)
        dt_b = c1.date_input("Data-Base", value=date.today())
        des_b = c1.checkbox("Desonerado?", value=False)
        arq = c2.file_uploader("Arquivo Excel (.xlsx / .xls)", type=["xlsx", "xls"])
        mapa, l_hdr = None, 0
        if arq and tipo_b == "GENERICA":
            l_hdr = c2.number_input("Linha do cabeçalho (0 = primeira)", min_value=0, value=0)
            cols = list(pd.read_excel(arq, skiprows=l_hdr, nrows=3).columns)
            arq.seek(0)
            mapa = {
                "codigo": c2.selectbox("Coluna Código", cols),
                "descricao": c2.selectbox("Coluna Descrição", cols),
                "unidade": c2.selectbox("Coluna Unidade", cols),
                "preco": c2.selectbox("Coluna Preço", cols),
            }
        if c2.button("Importar Base", type="primary", disabled=not arq):
            df_n = padronizar_planilha(arq, tipo_b, mapa, l_hdr)
            qtd = salvar_base_no_banco(db, nome_b, desc_b, uf_b, dt_b, des_b, df_n)
            st.success(f"{qtd} composições importadas!")
            st.rerun()

    bases = db.query(BaseReferencia).order_by(BaseReferencia.id.desc()).all()
    if bases:
        st.dataframe(pd.DataFrame([{
            "ID": b.id, "Base": b.nome, "Descrição": b.descricao, "UF": b.uf,
            "Data-Base": b.data_base.strftime("%d/%m/%Y"), "Regime": "Desonerado" if b.desonerado else "Onerado",
            "Itens": db.query(Composicao).filter_by(base_id=b.id).count()
        } for b in bases]), use_container_width=True, hide_index=True)

        op_d = {f"ID {b.id} — {b.nome} ({b.descricao})": b.id for b in bases}
        c_d1, c_d2 = st.columns([3, 1])
        b_del = c_d1.selectbox("Excluir Banco de Dados", list(op_d.keys()))
        c_d2.write("")
        c_d2.write("")
        if c_d2.button("🗑️ Excluir", use_container_width=True):
            db.delete(db.query(BaseReferencia).get(op_d[b_del]))
            db.commit()
            st.rerun()

# ==============================================================================
# 5. COMPRAS
# ==============================================================================
elif menu == "🛒 Compras":
    st.header("Gestão de Compras e Cotações por Obra")
    orcs = db.query(Orcamento).all()
    if not orcs:
        st.info("Crie um orçamento primeiro.")
    else:
        op_o = {f"#{o.id} — {o.titulo}": o.id for o in orcs}
        o_id = op_o[st.selectbox("Obra / Orçamento", list(op_o.keys()))]
        with st.form("f_comp", clear_on_submit=True):
            c1, c2, c3, c4, c5 = st.columns([2, 1.5, 0.8, 1, 1])
            item_c = c1.text_input("Item / Material *")
            forn_c = c2.text_input("Fornecedor")
            un_c = c3.text_input("Unid.", value="UN")
            qt_c = c4.number_input("Quant.", min_value=0.0, value=1.0)
            vl_c = c5.number_input("Unit. (R$)", min_value=0.0, value=0.0)
            st_c = st.selectbox("Status", ["Em Cotação", "Aprovado", "Comprado", "Entregue"])
            if st.form_submit_button("Registrar Pedido / Cotação", type="primary") and item_c:
                db.add(Compra(orcamento_id=o_id, fornecedor=forn_c, item_descricao=item_c, unidade=un_c, quantidade=qt_c, valor_unitario=vl_c, status=st_c))
                db.commit()
                st.rerun()

        compras = db.query(Compra).filter_by(orcamento_id=o_id).all()
        if compras:
            st.dataframe(pd.DataFrame([{
                "Data": c.data_pedido.strftime("%d/%m/%Y"), "Item": c.item_descricao, "Fornecedor": c.fornecedor,
                "Quant.": c.quantidade, "Unid.": c.unidade, "Unit. (R$)": c.valor_unitario,
                "Total (R$)": round(c.quantidade * c.valor_unitario, 2), "Status": c.status
            } for c in compras]), use_container_width=True, hide_index=True)

# ==============================================================================
# 6. ENCARGOS SOCIAIS
# ==============================================================================
elif menu == "👷 Encargos Sociais":
    st.header("Tabelas de Encargos Sociais (Horista / Mensalista)")
    with st.form("f_enc", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        n_enc = c1.text_input("Nome da Tabela *")
        des_enc = c2.checkbox("Desonerado?", value=True)
        hor_enc = c3.number_input("Horista (%)", value=84.73)
        men_enc = c4.number_input("Mensalista (%)", value=47.76)
        obs_enc = st.text_input("Observações")
        if st.form_submit_button("Salvar Encargos", type="primary") and n_enc:
            db.add(EncargoSocial(nome=n_enc, desonerado=des_enc, horista_pct=hor_enc, mensalista_pct=men_enc, observacoes=obs_enc))
            db.commit()
            st.rerun()

    encs = db.query(EncargoSocial).all()
    st.dataframe(pd.DataFrame([{
        "Tabela": e.nome, "Regime": "Desonerado" if e.desonerado else "Onerado",
        "Horista (%)": e.horista_pct, "Mensalista (%)": e.mensalista_pct, "Obs.": e.observacoes
    } for e in encs]), use_container_width=True, hide_index=True)

# ==============================================================================
# 7. PERFIS DE BDI
# ==============================================================================
elif menu == "🧮 Perfis de BDI":
    st.header("Perfis de BDI (Acórdão 2622/2013 - TCU)")
    with st.form("f_bdi"):
        nome_bdi = st.text_input("Nome do Perfil")
        c1, c2, c3 = st.columns(3)
        ac = c1.number_input("Administração Central (%)", value=4.00) / 100
        sg = c2.number_input("Seguros e Garantias (%)", value=0.80) / 100
        r = c3.number_input("Riscos (%)", value=1.27) / 100
        df = c1.number_input("Despesas Financeiras (%)", value=1.23) / 100
        l = c2.number_input("Lucro (%)", value=7.40) / 100
        t = c3.number_input("Tributos Totais (%)", value=8.65) / 100
        bdi_calc = calcular_bdi_tcu(ac, sg, r, df, l, t)
        st.metric("BDI Calculado", f"{bdi_calc:.2f}%")
        if st.form_submit_button("Salvar Perfil", type="primary") and nome_bdi:
            db.add(PerfilBDI(nome=nome_bdi, ac=ac, sg=sg, r=r, df=df, l=l, t=t, percentual_final=bdi_calc))
            db.commit()
            st.rerun()

    for p in db.query(PerfilBDI).all():
        cp1, cp2 = st.columns([3, 1])
        cp1.write(f"**{p.nome}** — `{p.percentual_final:.2f}%`")
        cp2.download_button("📥 Exportar BDI (.xlsx)", data=gerar_excel_bdi(p), file_name=f"BDI_{p.id}.xlsx", key=f"dl_bdi_{p.id}")

# ==============================================================================
# 8. CADASTROS
# ==============================================================================
elif menu == "🗂️ Cadastros":
    st.header("Cadastros Gerais (Clientes, Fornecedores e Responsáveis Técnicos)")
    with st.form("f_cad", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([1, 2, 1.5, 1.5])
        tp = c1.selectbox("Tipo", ["CLIENTE", "FORNECEDOR", "RESPONSAVEL_TECNICO"])
        nm = c2.text_input("Nome / Razão Social *")
        doc = c3.text_input("CNPJ / CPF / CREA")
        cnt = c4.text_input("Telefone / E-mail")
        if st.form_submit_button("Salvar Cadastro", type="primary") and nm:
            db.add(Cadastro(tipo=tp, nome=nm, documento=doc, contato=cnt))
            db.commit()
            st.rerun()

    cads = db.query(Cadastro).all()
    if cads:
        st.dataframe(pd.DataFrame([{"Tipo": c.tipo, "Nome": c.nome, "Documento": c.documento, "Contato": c.contato} for c in cads]), use_container_width=True, hide_index=True)

db.close()
