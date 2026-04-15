import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import openpyxl

st.set_page_config(page_title="Calculadora de Recursos — Usinagem", layout="wide", page_icon="🏭")

st.markdown("""
<style>
.metric-card {
    background: #f0f4f8; border-radius: 10px; padding: 16px 20px; margin: 4px 0;
    border-left: 4px solid #1E3A5F;
}
.metric-card.destaque { border-left-color: #E8A020; background: #fff8ed; }
.stAlert { border-radius: 8px; }
.section-title { font-size: 18px; font-weight: 600; color: #1E3A5F; margin: 24px 0 8px; }
</style>
""", unsafe_allow_html=True)

MESES       = ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
               "Maio","Junho","Julho","Agosto","Setembro","Outubro"]
MESES_ABREV = ["NOV","DEZ","JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT"]

# ══════════════════════════════════════════
# FUNÇÕES DE LEITURA
# ══════════════════════════════════════════
@st.cache_data
def load_excel(file_bytes):
    xl = pd.ExcelFile(BytesIO(file_bytes))
    return xl.sheet_names

def read_pmp(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name='INPUT_PMP', header=None)
    dias = {}
    for i, m in enumerate(MESES, 1):
        v = df.iloc[0, i]
        dias[m] = int(v) if pd.notna(v) else 0
    rows = []
    for r in range(2, len(df)):
        modelo = df.iloc[r, 0]
        if pd.isna(modelo): continue
        for i, m in enumerate(MESES, 1):
            v = df.iloc[r, i]
            qtd = int(v) if pd.notna(v) else 0
            if qtd > 0:
                rows.append({"modelo": str(modelo), "mes": m, "qtd": qtd})
    return pd.DataFrame(rows), dias

def read_tempo(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name='IMPUTTEMPO', header=0)
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca",
                             df.columns[5]:"t_ciclo", df.columns[6]:"t_labor"})
    return df[["centro","peca","t_ciclo","t_labor"]].dropna(subset=["centro"])

def read_dist(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name='IMPUTDISTRIBUIÇÃO', header=0)
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca",
                             df.columns[7]:"div_carga", df.columns[9]:"div_volume",
                             df.columns[10]:"disponib"})
    return df[["centro","peca","div_carga","div_volume","disponib"]].dropna(subset=["centro"])

def read_aplic(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name='IMPUTAPLICAÇÃO', header=0)
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca"})
    modelo_cols = [c for c in df.columns if str(c).startswith("MODELO")]
    df2 = df[["centro","peca"] + modelo_cols].copy()
    melted = df2.melt(id_vars=["centro","peca"], var_name="modelo", value_name="ativo")
    return melted[melted["ativo"]==1][["centro","peca","modelo"]].reset_index(drop=True)

def read_turnos(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name='IMPUTTURNOS', header=None)
    vals = df.iloc[0, 1:4].tolist()
    hA = float(vals[0]) if pd.notna(vals[0]) else 7.5
    hB = float(vals[1]) - float(vals[0]) if pd.notna(vals[1]) else 6.75
    hC = float(vals[2]) - float(vals[1]) if pd.notna(vals[2]) else 5.25
    return {"A": hA, "B": hB, "C": hC}

# ══════════════════════════════════════════
# VALIDAÇÕES
# ══════════════════════════════════════════
def validar(pmp, tempo, dist, aplic):
    alertas = []
    erros   = []

    # Peças em tempo mas não em aplicação
    chaves_tempo  = set(zip(tempo.centro, tempo.peca))
    chaves_aplic  = set(zip(aplic.centro, aplic.peca))
    so_tempo = chaves_tempo - chaves_aplic
    if so_tempo:
        alertas.append(f"⚠️ {len(so_tempo)} combinações centro+peça estão em IMPUTTEMPO mas sem nenhum modelo em IMPUTAPLICAÇÃO — nunca gerarão carga.")

    # Modelos com volume mas sem aplicação
    modelos_pmp   = set(pmp.modelo.unique())
    modelos_aplic = set(aplic.modelo.unique())
    sem_aplic = modelos_pmp - modelos_aplic
    if sem_aplic:
        alertas.append(f"⚠️ {len(sem_aplic)} modelo(s) com demanda no INPUT_PMP mas sem nenhuma peça em IMPUTAPLICAÇÃO: {', '.join(list(sem_aplic)[:5])}")

    # Disponibilidade zero → divisão por zero
    zero_disp = dist[dist.disponib == 0]
    if len(zero_disp):
        erros.append(f"🔴 {len(zero_disp)} linha(s) com disponibilidade = 0 em IMPUTDISTRIBUIÇÃO — causará divisão por zero.")

    # Labor > ciclo (fisicamente improvável)
    merged_tc = tempo.merge(dist, on=["centro","peca"], how="left")
    labor_maior = merged_tc[merged_tc.t_labor > merged_tc.t_ciclo]
    if len(labor_maior):
        alertas.append(f"⚠️ {len(labor_maior)} linha(s) com tempo de labor maior que tempo de ciclo — verifique os dados.")

    # Dias zerados em meses com demanda
    meses_com_demanda = pmp.mes.unique()
    from collections import defaultdict
    dias_dict = st.session_state.get("dias", {})
    for m in meses_com_demanda:
        if dias_dict.get(m, 1) == 0:
            alertas.append(f"⚠️ Mês '{m}' tem demanda mas dias trabalhados = 0.")

    return erros, alertas

# ══════════════════════════════════════════
# CÁLCULO PRINCIPAL
# ══════════════════════════════════════════
def calcular(pmp, tempo, dist, aplic, dias, turnos_h, overrides=None):
    # Passo 4+5 — join completo
    df = (aplic
          .merge(pmp,  on="modelo")
          .merge(tempo, on=["centro","peca"])
          .merge(dist,  on=["centro","peca"]))

    # Passo 6+7 — índice e minutos
    df["indice_ciclo"] = (df.t_ciclo * df.div_carga * df.div_volume) / df.disponib
    df["min_ciclo"]    = df.indice_ciclo * df.qtd
    df["min_labor"]    = df.t_labor * df.div_carga * df.qtd

    # Passo 8 — agrupar por centro+mês
    agg = df.groupby(["centro","mes"])[["min_ciclo","min_labor"]].sum().reset_index()
    agg["horas_ciclo"] = agg.min_ciclo / 60
    agg["horas_labor"] = agg.min_labor / 60

    resultados = {}
    for mes in MESES:
        d = dias.get(mes, 0)
        hA = turnos_h["A"]
        hB = turnos_h["B"]
        hC = turnos_h["C"]
        minA = d * hA * 60
        minB = d * hB * 60
        minC = d * hC * 60

        sub = agg[agg.mes == mes].copy()
        if sub.empty or d == 0:
            resultados[mes] = None
            continue

        centros = []
        for _, row in sub.iterrows():
            cen   = row.centro
            mc    = row.min_ciclo
            ml    = row.min_labor
            hc    = row.horas_ciclo
            hl    = row.horas_labor

            pA = mc / minA if minA > 0 else 0
            pB = mc / minB if minB > 0 else 0
            pC = mc / minC if minC > 0 else 0

            # Passo 11 — thresholds reais
            aA = 1 if pA > 0.40 else 0
            aB = 1 if pA > 1.06 else 0
            aC = 1 if pB > 1.00 else 0

            # Override manual do simulador
            if overrides and mes in overrides and cen in overrides[mes]:
                ov = overrides[mes][cen]
                if "A" in ov: aA = ov["A"]
                if "B" in ov: aB = ov["B"]
                if "C" in ov: aC = ov["C"]

            centros.append({
                "centro": cen, "ocup_A": pA, "ocup_B": pB, "ocup_C": pC,
                "ativo_A": aA, "ativo_B": aB, "ativo_C": aC,
                "horas_ciclo": hc, "horas_labor": hl,
                "horas_disp_A": d * hA * aA,
                "horas_disp_B": d * hB * aB,
                "horas_disp_C": d * hC * aC,
            })

        df_c = pd.DataFrame(centros)
        op_A = int(df_c.ativo_A.sum())
        op_B = int(df_c.ativo_B.sum())
        op_C = int(df_c.ativo_C.sum())

        # Suporte
        sup_A = 1 + 1 + 2 + 1 + 1   # lavadora, gravação, preset×2, coringa, facilitador
        sup_B = 1 + 1 + 1 + 0 + 1
        sup_C = 0 + 0 + 1 + 0 + 0

        tot_A = op_A + (sup_A if op_A > 0 else 0)
        tot_B = op_B + (sup_B if op_B > 0 else 0)
        tot_C = op_C + (sup_C if op_C > 0 else 0)
        total = tot_A + tot_B + tot_C

        h_ciclo_total = float(df_c.horas_ciclo.sum())
        h_labor_total = float(df_c.horas_labor.sum())
        h_ativos = float((df_c.horas_disp_A + df_c.horas_disp_B + df_c.horas_disp_C).sum())
        h_todos  = tot_A * d * hA + tot_B * d * hB + tot_C * d * hC

        resultados[mes] = {
            "centros": df_c,
            "op_A": op_A, "op_B": op_B, "op_C": op_C,
            "tot_A": tot_A, "tot_B": tot_B, "tot_C": tot_C,
            "total": total,
            "h_ciclo": h_ciclo_total, "h_labor": h_labor_total,
            "h_ativos": h_ativos, "h_todos": h_todos,
            "prod_ciclo_op":  h_ciclo_total / h_ativos if h_ativos > 0 else 0,
            "prod_ciclo_tot": h_ciclo_total / h_todos  if h_todos  > 0 else 0,
            "prod_labor_op":  h_labor_total / h_ativos if h_ativos > 0 else 0,
            "prod_labor_tot": h_labor_total / h_todos  if h_todos  > 0 else 0,
            "dias": d,
        }
    return resultados

# ══════════════════════════════════════════
# GRÁFICO RESUMO
# ══════════════════════════════════════════
def grafico_resumo(resultados):
    meses_v, tA, tB, tC, tot, prod = [], [], [], [], [], []
    for m, abr in zip(MESES, MESES_ABREV):
        r = resultados.get(m)
        if r is None: continue
        meses_v.append(abr)
        tA.append(r["tot_A"]); tB.append(r["tot_B"]); tC.append(r["tot_C"])
        tot.append(r["total"])
        prod.append(r["prod_labor_tot"])

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(name="Turno A", x=meses_v, y=tA, marker_color="#2E7D32", text=tA, textposition="inside"), secondary_y=False)
    fig.add_trace(go.Bar(name="Turno B", x=meses_v, y=tB, marker_color="#F9A825", text=tB, textposition="inside"), secondary_y=False)
    fig.add_trace(go.Bar(name="Turno C", x=meses_v, y=tC, marker_color="#1565C0", text=tC, textposition="inside"), secondary_y=False)
    fig.add_trace(go.Bar(name="Total",   x=meses_v, y=tot, marker_color="#9E9E9E", text=tot, textposition="outside"), secondary_y=False)
    fig.add_trace(go.Scatter(
        name="Produtividade Labor Total", x=meses_v,
        y=[p*100 for p in prod],
        mode="lines+markers+text",
        marker=dict(color="#CC0000", size=10, symbol="circle"),
        line=dict(color="#CC0000", width=2),
        text=[f"{p*100:.0f}%" for p in prod],
        textposition="top center",
    ), secondary_y=True)

    fig.update_layout(
        barmode="group", title="MÃO-DE-OBRA POR TURNO E PRODUTIVIDADE",
        yaxis_title="Nº Funcionários",
        yaxis2_title="Produtividade (%)",
        yaxis2=dict(tickformat=".0f", ticksuffix="%", range=[0,100]),
        legend=dict(orientation="h", y=-0.2),
        height=420, plot_bgcolor="white",
    )
    return fig

# ══════════════════════════════════════════
# EXPORTAÇÃO
# ══════════════════════════════════════════
def exportar(resultados):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        rows = []
        for m in MESES:
            r = resultados.get(m)
            if r is None: continue
            rows.append({
                "Mês": m, "Dias": r["dias"],
                "Turno A": r["tot_A"], "Turno B": r["tot_B"], "Turno C": r["tot_C"],
                "Total Func.": r["total"],
                "Ciclo Operacional": f"{r['prod_ciclo_op']:.0%}",
                "Ciclo Total": f"{r['prod_ciclo_tot']:.0%}",
                "Labor Operacional": f"{r['prod_labor_op']:.0%}",
                "Labor Total ★": f"{r['prod_labor_tot']:.0%}",
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name="Resumo MO", index=False)
        for m in MESES:
            r = resultados.get(m)
            if r is None: continue
            r["centros"].to_excel(writer, sheet_name=m[:10], index=False)
    output.seek(0)
    return output

# ══════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════
st.title("🏭 Calculadora de Recursos — Usinagem")
st.caption("Faça upload da planilha para começar. Todas as abas de input serão lidas automaticamente.")

uploaded = st.file_uploader("Upload do arquivo .xlsm", type=["xlsm","xlsx"])

if not uploaded:
    st.info("👆 Faça upload do arquivo para começar.")
    st.stop()

file_bytes = uploaded.read()

with st.spinner("Lendo planilha..."):
    try:
        pmp,  dias   = read_pmp(file_bytes)
        tempo         = read_tempo(file_bytes)
        dist          = read_dist(file_bytes)
        aplic         = read_aplic(file_bytes)
        turnos_h      = read_turnos(file_bytes)
        st.session_state["dias"] = dias
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        st.stop()

st.success(f"✅ Arquivo lido — {len(aplic)} combinações centro/peça/modelo | {pmp.modelo.nunique()} modelos | {pmp.mes.nunique()} meses com demanda")

# ── Validações ──
erros, alertas = validar(pmp, tempo, dist, aplic)
if erros:
    for e in erros:
        st.error(e)
    st.stop()
if alertas:
    with st.expander(f"⚠️ {len(alertas)} aviso(s) encontrado(s) — clique para ver", expanded=False):
        for a in alertas:
            st.warning(a)

# ── Tabs ──
tab1, tab2, tab3, tab4 = st.tabs(["📊 Resultados", "🔧 Simulador", "🔍 Detalhes por etapa", "📥 Exportar"])

# ══════════════════════════════════════════
# TAB 1 — RESULTADOS
# ══════════════════════════════════════════
with tab1:
    overrides = st.session_state.get("overrides", {})
    resultados = calcular(pmp, tempo, dist, aplic, dias, turnos_h, overrides)

    st.plotly_chart(grafico_resumo(resultados), use_container_width=True)

    st.markdown('<div class="section-title">Resumo por mês</div>', unsafe_allow_html=True)
    rows = []
    for m, abr in zip(MESES, MESES_ABREV):
        r = resultados.get(m)
        if r is None:
            rows.append({"Mês": abr, "Turno A": "-", "Turno B": "-", "Turno C": "-",
                         "Total": "-", "Ciclo Op.": "-", "Labor Total ★": "-"})
        else:
            rows.append({
                "Mês": abr,
                "Turno A": r["tot_A"], "Turno B": r["tot_B"], "Turno C": r["tot_C"],
                "Total": r["total"],
                "Ciclo Op.": f"{r['prod_ciclo_op']:.0%}",
                "Labor Total ★": f"{r['prod_labor_tot']:.0%}",
            })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Indicadores do mês selecionado</div>', unsafe_allow_html=True)
    mes_sel = st.selectbox("Selecione o mês", [m for m in MESES if resultados.get(m)], key="mes_res")
    if mes_sel:
        r = resultados[mes_sel]
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Ciclo Operacional",  f"{r['prod_ciclo_op']:.0%}")
        c2.metric("Ciclo Total",        f"{r['prod_ciclo_tot']:.0%}")
        c3.metric("Labor Operacional",  f"{r['prod_labor_op']:.0%}")
        c4.metric("Labor Total ★",      f"{r['prod_labor_tot']:.0%}")

        st.markdown("**Ocupação por centro**")
        df_show = r["centros"][["centro","ocup_A","ocup_B","ocup_C","ativo_A","ativo_B","ativo_C"]].copy()
        df_show["ocup_A"] = df_show.ocup_A.map(lambda x: f"{x:.0%}")
        df_show["ocup_B"] = df_show.ocup_B.map(lambda x: f"{x:.0%}")
        df_show["ocup_C"] = df_show.ocup_C.map(lambda x: f"{x:.0%}")
        st.dataframe(df_show, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════
# TAB 2 — SIMULADOR
# ══════════════════════════════════════════
with tab2:
    st.markdown("### Simulador de headcount")
    st.caption("Force a ativação ou desativação de turnos por centro e veja o impacto nos indicadores em tempo real.")

    mes_sim = st.selectbox("Mês para simular", MESES, key="mes_sim")
    r_base  = calcular(pmp, tempo, dist, aplic, dias, turnos_h)
    r_orig  = r_base.get(mes_sim)

    if r_orig is None:
        st.warning("Mês sem dados disponíveis.")
    else:
        centros_list = sorted(r_orig["centros"].centro.tolist())
        st.markdown("**Ative ou desative turnos por centro:**")

        if "overrides" not in st.session_state:
            st.session_state.overrides = {}
        if mes_sim not in st.session_state.overrides:
            st.session_state.overrides[mes_sim] = {}

        cols = st.columns([2,1,1,1])
        cols[0].markdown("**Centro**")
        cols[1].markdown("**Turno A**")
        cols[2].markdown("**Turno B**")
        cols[3].markdown("**Turno C**")

        for cen in centros_list:
            row_cen = r_orig["centros"][r_orig["centros"].centro == cen].iloc[0]
            ov = st.session_state.overrides[mes_sim].get(cen, {})
            c0,c1,c2,c3 = st.columns([2,1,1,1])
            c0.markdown(f"`{cen}`  {row_cen.ocup_A:.0%} / {row_cen.ocup_B:.0%} / {row_cen.ocup_C:.0%}")
            vA = c1.checkbox("", value=bool(ov.get("A", row_cen.ativo_A)), key=f"{mes_sim}_{cen}_A")
            vB = c2.checkbox("", value=bool(ov.get("B", row_cen.ativo_B)), key=f"{mes_sim}_{cen}_B")
            vC = c3.checkbox("", value=bool(ov.get("C", row_cen.ativo_C)), key=f"{mes_sim}_{cen}_C")
            st.session_state.overrides[mes_sim][cen] = {
                "A": int(vA), "B": int(vB), "C": int(vC)
            }

        r_sim = calcular(pmp, tempo, dist, aplic, dias, turnos_h, st.session_state.overrides)
        r_new = r_sim.get(mes_sim)

        if r_new:
            st.markdown("---")
            st.markdown("**Impacto da simulação**")
            c1,c2,c3,c4 = st.columns(4)
            delta_tot = r_new["total"] - r_orig["total"]
            delta_lab = r_new["prod_labor_tot"] - r_orig["prod_labor_tot"]
            c1.metric("Total funcionários", r_new["total"], delta=f"{delta_tot:+d}")
            c2.metric("Turno A", r_new["tot_A"], delta=f"{r_new['tot_A']-r_orig['tot_A']:+d}")
            c3.metric("Turno B", r_new["tot_B"], delta=f"{r_new['tot_B']-r_orig['tot_B']:+d}")
            c4.metric("Turno C", r_new["tot_C"], delta=f"{r_new['tot_C']-r_orig['tot_C']:+d}")
            c1b,c2b,c3b,c4b = st.columns(4)
            c1b.metric("Labor Total ★", f"{r_new['prod_labor_tot']:.0%}", delta=f"{delta_lab:+.1%}")
            c2b.metric("Labor Operacional", f"{r_new['prod_labor_op']:.0%}")
            c3b.metric("Ciclo Total", f"{r_new['prod_ciclo_tot']:.0%}")
            c4b.metric("Ciclo Operacional", f"{r_new['prod_ciclo_op']:.0%}")

# ══════════════════════════════════════════
# TAB 3 — DETALHES POR ETAPA
# ══════════════════════════════════════════
with tab3:
    st.markdown("### Detalhes por etapa do cálculo")
    etapa = st.radio("Escolha a etapa", [
        "Passo 2 — INPUT_PMP (normalizado)",
        "Passo 3 — IMPUTAPLICAÇÃO (normalizado)",
        "Passo 4 — JOIN aplicacao × pmp",
        "Passo 7 — Minutos calculados por linha",
        "Passo 8 — Totais por centro + mês",
        "Passo 10 — % ocupação por centro",
    ], horizontal=False)

    mes_det = st.selectbox("Mês", MESES, key="mes_det")

    if etapa == "Passo 2 — INPUT_PMP (normalizado)":
        st.dataframe(pmp[pmp.mes == mes_det].reset_index(drop=True), use_container_width=True, hide_index=True)

    elif etapa == "Passo 3 — IMPUTAPLICAÇÃO (normalizado)":
        st.dataframe(aplic.head(200), use_container_width=True, hide_index=True)
        st.caption(f"Total: {len(aplic)} combinações ativas (somente onde flag = 1)")

    elif etapa == "Passo 4 — JOIN aplicacao × pmp":
        pmp_mes = pmp[pmp.mes == mes_det]
        p4 = aplic.merge(pmp_mes, on="modelo")
        st.dataframe(p4.head(200), use_container_width=True, hide_index=True)
        st.caption(f"{len(p4)} linhas para {mes_det}")

    elif etapa == "Passo 7 — Minutos calculados por linha":
        pmp_mes = pmp[pmp.mes == mes_det]
        p7 = (aplic.merge(pmp_mes, on="modelo")
                   .merge(tempo, on=["centro","peca"])
                   .merge(dist, on=["centro","peca"]))
        p7["indice_ciclo"] = (p7.t_ciclo * p7.div_carga * p7.div_volume) / p7.disponib
        p7["min_ciclo"]    = (p7.indice_ciclo * p7.qtd).round(1)
        p7["min_labor"]    = (p7.t_labor * p7.div_carga * p7.qtd).round(1)
        st.dataframe(p7.head(300), use_container_width=True, hide_index=True)
        st.caption(f"{len(p7)} linhas para {mes_det}")

    elif etapa == "Passo 8 — Totais por centro + mês":
        pmp_mes = pmp[pmp.mes == mes_det]
        p7 = (aplic.merge(pmp_mes, on="modelo")
                   .merge(tempo, on=["centro","peca"])
                   .merge(dist, on=["centro","peca"]))
        p7["indice_ciclo"] = (p7.t_ciclo * p7.div_carga * p7.div_volume) / p7.disponib
        p7["min_ciclo"]    = p7.indice_ciclo * p7.qtd
        p7["min_labor"]    = p7.t_labor * p7.div_carga * p7.qtd
        p8 = p7.groupby("centro")[["min_ciclo","min_labor"]].sum().reset_index()
        p8["horas_ciclo"] = (p8.min_ciclo/60).round(1)
        p8["horas_labor"] = (p8.min_labor/60).round(1)
        st.dataframe(p8, use_container_width=True, hide_index=True)

    elif etapa == "Passo 10 — % ocupação por centro":
        r = calcular(pmp, tempo, dist, aplic, dias, turnos_h).get(mes_det)
        if r:
            df_oc = r["centros"][["centro","ocup_A","ocup_B","ocup_C"]].copy()
            def cor(v):
                if v > 1.0: return "🔴"
                if v >= 0.85: return "🟡"
                return "🟢"
            df_oc["status A"] = df_oc.ocup_A.map(cor)
            df_oc["status B"] = df_oc.ocup_B.map(cor)
            df_oc["status C"] = df_oc.ocup_C.map(cor)
            df_oc["ocup_A"] = df_oc.ocup_A.map(lambda x: f"{x:.0%}")
            df_oc["ocup_B"] = df_oc.ocup_B.map(lambda x: f"{x:.0%}")
            df_oc["ocup_C"] = df_oc.ocup_C.map(lambda x: f"{x:.0%}")
            st.dataframe(df_oc, use_container_width=True, hide_index=True)
            st.caption("🟢 < 85%   🟡 85–100%   🔴 > 100%")

# ══════════════════════════════════════════
# TAB 4 — EXPORTAR
# ══════════════════════════════════════════
with tab4:
    st.markdown("### Exportar resultados")
    overrides = st.session_state.get("overrides", {})
    resultados_exp = calcular(pmp, tempo, dist, aplic, dias, turnos_h, overrides)
    xlsx = exportar(resultados_exp)
    st.download_button(
        "📥 Baixar Excel com resultados",
        data=xlsx,
        file_name="resultado_usinagem.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.caption("O arquivo inclui uma aba de resumo + uma aba de detalhes por centro para cada mês.")
