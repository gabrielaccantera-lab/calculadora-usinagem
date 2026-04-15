import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import copy

st.set_page_config(
    page_title="Calculadora de Recursos — Usinagem",
    layout="wide", page_icon="🏭"
)

st.markdown("""
<style>
/* Cores principais */
:root {
    --azul: #1E3A5F;
    --azul-medio: #2D6A9F;
    --verde: #2E7D32;
    --amarelo: #F9A825;
    --vermelho: #C62828;
    --cinza: #F0F4F8;
}
/* Cabeçalho */
.main-header {
    background: linear-gradient(90deg, #1E3A5F 0%, #2D6A9F 100%);
    color: white; padding: 18px 24px; border-radius: 10px;
    margin-bottom: 20px;
}
.main-header h1 { color: white; margin: 0; font-size: 22px; }
.main-header p  { color: #B5D4F4; margin: 4px 0 0; font-size: 13px; }

/* Cards de métricas */
.metric-row { display: flex; gap: 12px; margin: 12px 0; flex-wrap: wrap; }
.metric-card {
    flex: 1; min-width: 140px;
    background: #F0F4F8; border-radius: 10px;
    padding: 14px 16px; border-left: 4px solid #1E3A5F;
}
.metric-card.destaque { border-left-color: #E8A020; background: #FFF8ED; }
.metric-card.verde    { border-left-color: #2E7D32; background: #F1F8E9; }
.metric-card.vermelho { border-left-color: #C62828; background: #FFEBEE; }
.metric-card .label   { font-size: 11px; color: #666; margin-bottom: 4px; }
.metric-card .valor   { font-size: 22px; font-weight: 700; color: #1E3A5F; }
.metric-card .delta   { font-size: 11px; margin-top: 2px; }
.delta-pos { color: #2E7D32; }
.delta-neg { color: #C62828; }

/* Tabela de resultados estilo Excel */
.tabela-excel {
    width: 100%; border-collapse: collapse; font-size: 12px;
    font-family: Arial, sans-serif;
}
.tabela-excel th {
    background: #1E3A5F; color: white; padding: 7px 10px;
    text-align: center; border: 1px solid #0D2137; font-weight: 600;
}
.tabela-excel td {
    padding: 5px 10px; border: 1px solid #D0D7E0;
    text-align: center;
}
.tabela-excel tr:nth-child(even) td { background: #EAF3FB; }
.tabela-excel tr:nth-child(odd)  td { background: #FFFFFF; }
.tabela-excel .centro-cell { text-align: left; font-weight: 600; color: #1E3A5F; }
.tabela-excel .pct-verde   { background: #C8E6C9 !important; color: #1B5E20; font-weight: 600; }
.tabela-excel .pct-amarelo { background: #FFF9C4 !important; color: #F57F17; font-weight: 600; }
.tabela-excel .pct-vermelho{ background: #FFCDD2 !important; color: #B71C1C; font-weight: 600; }
.tabela-excel .flag-ativo  { background: #B3E5FC !important; color: #01579B; font-weight: 700; }
.tabela-excel .flag-inativo{ background: #FFF9C4 !important; color: #F57F17; font-weight: 700; }
.tabela-excel .total-row td{ background: #FF8A80 !important; color: #000; font-weight: 700; border: 1px solid #CC0000; }
.tabela-excel .suporte-row td{ background: #FFFFFF; }
.tabela-excel .horas-cell  { background: #B3E5FC !important; color: #01579B; }
.tabela-excel .horas-zero  { background: #FFF9C4 !important; color: #888; }
.tabela-excel .prod-row td { background: #FFFDE7; }
.tabela-excel .prod-destaque td { background: #FFF9C4 !important; font-weight: 700; color: #E65100; }

/* Aviso cards */
.aviso-erro    { background: #FFEBEE; border-left: 4px solid #C62828; border-radius: 8px; padding: 10px 14px; margin: 6px 0; font-size: 13px; color: #B71C1C; }
.aviso-alerta  { background: #FFF8E1; border-left: 4px solid #F9A825; border-radius: 8px; padding: 10px 14px; margin: 6px 0; font-size: 13px; color: #E65100; }
.aviso-ok      { background: #E8F5E9; border-left: 4px solid #2E7D32; border-radius: 8px; padding: 10px 14px; margin: 6px 0; font-size: 13px; color: #1B5E20; }

/* Seção título */
.section-title {
    font-size: 16px; font-weight: 700; color: #1E3A5F;
    border-bottom: 2px solid #1E3A5F; padding-bottom: 6px;
    margin: 20px 0 12px;
}
.subsection { font-size: 13px; font-weight: 600; color: #2D6A9F; margin: 14px 0 6px; }

/* Input arquivo */
.upload-box {
    border: 2px dashed #2D6A9F; border-radius: 10px;
    padding: 24px; text-align: center; background: #F0F4F8;
    margin: 12px 0;
}

/* Simulador */
.sim-centro {
    display: flex; align-items: center; gap: 12px;
    padding: 8px 12px; border-radius: 8px; margin: 3px 0;
    background: #F8FAFC; border: 1px solid #E0E7EF;
}
.sim-badge {
    background: #1E3A5F; color: white; font-size: 11px;
    font-weight: 700; padding: 3px 8px; border-radius: 6px;
    min-width: 64px; text-align: center;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
MESES       = ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
               "Maio","Junho","Julho","Agosto","Setembro","Outubro"]
MESES_ABREV = ["NOV","DEZ","JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT"]

HORAS_TURNO = {"A": 8.80, "B": 8.23, "C": 7.68}  # horas de duração de cada turno

# ─────────────────────────────────────────
# LEITURA
# ─────────────────────────────────────────
def read_pmp(fb):
    df = pd.read_excel(BytesIO(fb), sheet_name='INPUT_PMP', header=None)
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
                rows.append({"modelo": str(modelo).strip(), "mes": m, "qtd": qtd})
    return pd.DataFrame(rows), dias

def read_tempo(fb):
    df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTTEMPO', header=0)
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca",
                             df.columns[5]:"t_ciclo", df.columns[6]:"t_labor"})
    return df[["centro","peca","t_ciclo","t_labor"]].dropna(subset=["centro"]).copy()

def read_dist(fb):
    df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTDISTRIBUIÇÃO', header=0)
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca",
                             df.columns[7]:"div_carga", df.columns[9]:"div_volume",
                             df.columns[10]:"disponib"})
    return df[["centro","peca","div_carga","div_volume","disponib"]].dropna(subset=["centro"]).copy()

def read_aplic(fb):
    df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTAPLICAÇÃO', header=0)
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca"})
    mcols = [c for c in df.columns if str(c).startswith("MODELO")]
    melted = df[["centro","peca"]+mcols].melt(
        id_vars=["centro","peca"], var_name="modelo", value_name="ativo")
    return melted[melted["ativo"]==1][["centro","peca","modelo"]].reset_index(drop=True)

# ─────────────────────────────────────────
# VALIDAÇÕES
# ─────────────────────────────────────────
def validar(pmp, tempo, dist, aplic, dias):
    erros, alertas, oks = [], [], []

    chaves_tempo = set(zip(tempo.centro, tempo.peca))
    chaves_dist  = set(zip(dist.centro,  dist.peca))
    chaves_aplic = set(zip(aplic.centro, aplic.peca))

    # Erros críticos
    zero_disp = dist[dist.disponib == 0]
    if len(zero_disp):
        linhas = ", ".join([f"{r.centro}/{r.peca}" for _, r in zero_disp.iterrows()][:5])
        erros.append(f"Disponibilidade = 0 em {len(zero_disp)} linha(s) — causará divisão por zero: {linhas}")

    so_tempo_sem_dist = chaves_tempo - chaves_dist
    if so_tempo_sem_dist:
        erros.append(f"{len(so_tempo_sem_dist)} combinação(ões) centro+peça estão em IMPUTTEMPO mas faltam em IMPUTDISTRIBUIÇÃO.")

    # Avisos
    so_tempo_sem_aplic = chaves_tempo - chaves_aplic
    if so_tempo_sem_aplic:
        exemplos = list(so_tempo_sem_aplic)[:3]
        alertas.append(f"{len(so_tempo_sem_aplic)} combinação(ões) centro+peça em IMPUTTEMPO sem nenhum modelo em IMPUTAPLICAÇÃO — nunca gerarão carga. Exemplos: {exemplos}")

    modelos_pmp   = set(pmp.modelo.unique())
    modelos_aplic = set(aplic.modelo.unique())
    sem_aplic = modelos_pmp - modelos_aplic
    if sem_aplic:
        alertas.append(f"{len(sem_aplic)} modelo(s) com demanda no INPUT_PMP mas sem aplicação: {', '.join(list(sem_aplic)[:5])}")

    merged = tempo.merge(dist, on=["centro","peca"], how="inner")
    labor_maior = merged[merged.t_labor > merged.t_ciclo]
    if len(labor_maior):
        ex = [(r.centro, r.peca) for _, r in labor_maior.iterrows()][:3]
        alertas.append(f"{len(labor_maior)} linha(s) com tempo de labor maior que ciclo (improvável fisicamente): {ex}")

    for m in MESES:
        qtd_mes = pmp[pmp.mes==m].qtd.sum() if len(pmp[pmp.mes==m]) else 0
        if qtd_mes > 0 and dias.get(m, 0) == 0:
            alertas.append(f"Mês '{m}' tem {qtd_mes} peças de demanda mas dias trabalhados = 0 — será ignorado no cálculo.")

    # OK
    if not erros and not alertas:
        oks.append("Nenhuma inconsistência encontrada nos inputs.")
    if len(chaves_aplic - chaves_tempo) == 0:
        oks.append("Todos os centros+peças da IMPUTAPLICAÇÃO têm tempos cadastrados.")

    return erros, alertas, oks

# ─────────────────────────────────────────
# CÁLCULO
# ─────────────────────────────────────────
def calcular(pmp, tempo, dist, aplic, dias, overrides=None):
    df = (aplic
          .merge(pmp,   on="modelo")
          .merge(tempo, on=["centro","peca"])
          .merge(dist,  on=["centro","peca"]))

    df["indice_ciclo"] = (df.t_ciclo * df.div_carga * df.div_volume) / df.disponib
    df["min_ciclo"]    = df.indice_ciclo * df.qtd
    df["min_labor"]    = df.t_labor * df.div_carga * df.qtd

    agg = df.groupby(["centro","mes"])[["min_ciclo","min_labor"]].sum().reset_index()

    resultados = {}
    for mes in MESES:
        d = dias.get(mes, 0)
        if d == 0:
            resultados[mes] = None
            continue

        hA = HORAS_TURNO["A"]
        hB = HORAS_TURNO["B"]
        hC = HORAS_TURNO["C"]
        minA = d * hA * 60
        minB = d * hB * 60
        minC = d * hC * 60

        sub = agg[agg.mes == mes].copy()
        if sub.empty:
            resultados[mes] = None
            continue

        centros = []
        for _, row in sub.iterrows():
            cen = row.centro
            mc  = row.min_ciclo
            ml  = row.min_labor

            pA = mc / minA if minA > 0 else 0
            pB = mc / minB if minB > 0 else 0
            pC = mc / minC if minC > 0 else 0

            # Thresholds reais do Excel
            aA = 1 if pA > 0.40 else 0
            aB = 1 if pA > 1.06 else 0
            aC = 1 if pB > 1.00 else 0

            # Override do simulador
            if overrides and mes in overrides and cen in overrides[mes]:
                ov = overrides[mes][cen]
                if "A" in ov: aA = ov["A"]
                if "B" in ov: aB = ov["B"]
                if "C" in ov: aC = ov["C"]

            centros.append({
                "centro":    cen,
                "ocup_A":    pA, "ocup_B": pB, "ocup_C": pC,
                "ativo_A":   aA, "ativo_B": aB, "ativo_C": aC,
                "horas_ciclo": mc/60, "horas_labor": ml/60,
                "horas_disp_A": d * hA * aA,
                "horas_disp_B": d * hB * aB,
                "horas_disp_C": d * hC * aC,
            })

        df_c = pd.DataFrame(centros)
        op_A = int(df_c.ativo_A.sum())
        op_B = int(df_c.ativo_B.sum())
        op_C = int(df_c.ativo_C.sum())

        # Suporte (regras fixas)
        lav_A = 1 if op_A > 0 else 0
        lav_B = 1 if op_B > 0 else 0
        lav_C = 0
        gra_A = 1 if op_A > 0 else 0
        gra_B = 1 if op_B > 0 else 0
        gra_C = 0
        pre_A = 2; pre_B = 1; pre_C = 1 if op_C > 0 else 0
        cor_A = 1 if op_A > 0 else 0
        cor_B = 0; cor_C = 0
        fac_A = 1 if op_A > 0 else 0
        fac_B = 1 if op_B > 0 else 0
        fac_C = 0

        tot_A = op_A + lav_A + gra_A + pre_A + cor_A + fac_A
        tot_B = op_B + lav_B + gra_B + pre_B + cor_B + fac_B
        tot_C = op_C + lav_C + gra_C + pre_C + cor_C + fac_C
        total = tot_A + tot_B + tot_C

        h_ciclo  = float(df_c.horas_ciclo.sum())
        h_labor  = float(df_c.horas_labor.sum())
        h_ativos = float((df_c.horas_disp_A + df_c.horas_disp_B + df_c.horas_disp_C).sum())
        h_todos  = tot_A * d * hA + tot_B * d * hB + tot_C * d * hC

        resultados[mes] = {
            "centros": df_c,
            "op_A": op_A, "op_B": op_B, "op_C": op_C,
            "tot_A": tot_A, "tot_B": tot_B, "tot_C": tot_C,
            "total": total,
            "suporte": {
                "lavadora":  {"A": lav_A, "B": lav_B, "C": lav_C},
                "gravacao":  {"A": gra_A, "B": gra_B, "C": gra_C},
                "preset":    {"A": pre_A, "B": pre_B, "C": pre_C},
                "coringa":   {"A": cor_A, "B": cor_B, "C": cor_C},
                "facilitador":{"A": fac_A,"B": fac_B, "C": fac_C},
            },
            "h_ciclo": h_ciclo, "h_labor": h_labor,
            "h_ativos": h_ativos, "h_todos": h_todos,
            "prod_ciclo_op":  h_ciclo / h_ativos if h_ativos > 0 else 0,
            "prod_ciclo_tot": h_ciclo / h_todos  if h_todos  > 0 else 0,
            "prod_labor_op":  h_labor / h_ativos if h_ativos > 0 else 0,
            "prod_labor_tot": h_labor / h_todos  if h_todos  > 0 else 0,
            "dias": d,
            "horas_turno_A": d * HORAS_TURNO["A"],
            "horas_turno_B": d * HORAS_TURNO["B"],
            "horas_turno_C": d * HORAS_TURNO["C"],
        }
    return resultados

# ─────────────────────────────────────────
# TABELA ESTILO EXCEL
# ─────────────────────────────────────────
def cor_pct(v):
    if v > 1.0:   return "pct-vermelho"
    if v >= 0.85: return "pct-amarelo"
    return "pct-verde"

def tabela_html(r, mes):
    hA = HORAS_TURNO["A"] * r["dias"]
    hB = HORAS_TURNO["B"] * r["dias"]
    hC = HORAS_TURNO["C"] * r["dias"]

    html = f"""
    <div style="overflow-x:auto">
    <table class="tabela-excel">
    <thead>
      <tr>
        <th colspan="4" style="background:#1565C0">DADOS AUTOMÁTICOS</th>
        <th colspan="3" style="background:#1565C0"></th>
        <th colspan="3" style="background:#1565C0">HORAS POR TURNO DE TRABALHO</th>
      </tr>
      <tr>
        <th colspan="4" style="background:#1E3A5F">PERÍODO: {mes.upper()}</th>
        <th colspan="3" style="background:#1E3A5F"></th>
        <th style="background:#1E3A5F">{HORAS_TURNO['A']}</th>
        <th style="background:#1E3A5F">{HORAS_TURNO['B']}</th>
        <th style="background:#1E3A5F">{HORAS_TURNO['C']}</th>
      </tr>
      <tr>
        <th></th>
        <th style="background:#4CAF50">TURNO A</th>
        <th style="background:#FDD835;color:#000">TURNO B</th>
        <th style="background:#90CAF9;color:#000">TURNO C</th>
        <th style="background:#4CAF50">TURNO A</th>
        <th style="background:#FDD835;color:#000">TURNO B</th>
        <th style="background:#90CAF9;color:#000">TURNO C</th>
        <th style="background:#4CAF50">TURNO A</th>
        <th style="background:#FDD835;color:#000">TURNO B</th>
        <th style="background:#90CAF9;color:#000">TURNO C</th>
      </tr>
    </thead>
    <tbody>
    """

    for _, row in r["centros"].iterrows():
        cA = cor_pct(row.ocup_A)
        cB = cor_pct(row.ocup_B)
        cC = cor_pct(row.ocup_C)
        fA = "flag-ativo" if row.ativo_A else "flag-inativo"
        fB = "flag-ativo" if row.ativo_B else "flag-inativo"
        fC = "flag-ativo" if row.ativo_C else "flag-inativo"
        hcA = f"{row.horas_disp_A:.2f}" if row.ativo_A else "0"
        hcB = f"{row.horas_disp_B:.2f}" if row.ativo_B else "0"
        hcC = f"{row.horas_disp_C:.2f}" if row.ativo_C else "0"
        hclA = "horas-cell" if row.ativo_A else "horas-zero"
        hclB = "horas-cell" if row.ativo_B else "horas-zero"
        hclC = "horas-cell" if row.ativo_C else "horas-zero"
        html += f"""
        <tr>
          <td class="centro-cell">{row.centro}</td>
          <td class="{cA}">{row.ocup_A:.0%}</td>
          <td class="{cB}">{row.ocup_B:.0%}</td>
          <td class="{cC}">{row.ocup_C:.0%}</td>
          <td class="{fA}">{row.ativo_A}</td>
          <td class="{fB}">{row.ativo_B}</td>
          <td class="{fC}">{row.ativo_C}</td>
          <td class="{hclA}">{hcA}</td>
          <td class="{hclB}">{hcB}</td>
          <td class="{hclC}">{hcC}</td>
        </tr>"""

    sup = r["suporte"]
    tot_h = r["tot_A"]*hA + r["tot_B"]*hB + r["tot_C"]*hC

    def sup_row(nome, s):
        fA = "flag-ativo" if s["A"] else "flag-inativo"
        fB = "flag-ativo" if s["B"] else "flag-inativo"
        fC = "flag-ativo" if s["C"] else "flag-inativo"
        hA2 = f"{s['A']*HORAS_TURNO['A']*r['dias']:.2f}" if s["A"] else "0"
        hB2 = f"{s['B']*HORAS_TURNO['B']*r['dias']:.2f}" if s["B"] else "0"
        hC2 = f"{s['C']*HORAS_TURNO['C']*r['dias']:.2f}" if s["C"] else "0"
        hclA2 = "horas-cell" if s["A"] else "horas-zero"
        hclB2 = "horas-cell" if s["B"] else "horas-zero"
        hclC2 = "horas-cell" if s["C"] else "horas-zero"
        return f"""
        <tr class="suporte-row">
          <td class="centro-cell" style="font-weight:400;color:#444">{nome}</td>
          <td></td><td></td><td></td>
          <td class="{fA}">{s['A']}</td>
          <td class="{fB}">{s['B']}</td>
          <td class="{fC}">{s['C']}</td>
          <td class="{hclA2}">{hA2}</td>
          <td class="{hclB2}">{hB2}</td>
          <td class="{hclC2}">{hC2}</td>
        </tr>"""

    # Total operadores CEN
    totop_hA = r["op_A"]*hA; totop_hB = r["op_B"]*hB; totop_hC = r["op_C"]*hC
    html += f"""
    <tr class="total-row">
      <td>TOTAL DE OPERADORES</td>
      <td></td><td></td><td></td>
      <td>{r['op_A']}</td><td>{r['op_B']}</td><td>{r['op_C']}</td>
      <td>{totop_hA:.2f}</td><td>{totop_hB:.2f}</td><td>{totop_hC:.2f}</td>
    </tr>"""

    for nome, key in [("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),
                      ("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador")]:
        html += sup_row(nome, sup[key])

    html += f"""
    <tr class="total-row">
      <td>TOTAL POR TURNO</td>
      <td></td><td></td><td></td>
      <td>{r['tot_A']}</td><td>{r['tot_B']}</td><td>{r['tot_C']}</td>
      <td>{r['tot_A']*hA:.2f}</td><td>{r['tot_B']*hB:.2f}</td><td>{r['tot_C']*hC:.2f}</td>
    </tr>
    <tr class="total-row">
      <td colspan="3">TOTAL FUNCIONÁRIOS</td>
      <td style="font-size:15px">{r['total']}</td>
      <td colspan="3"></td>
      <td colspan="3" style="font-size:13px">{tot_h:.2f}</td>
    </tr>
    </tbody>
    <tfoot>
      <tr class="prod-row">
        <td colspan="9" style="text-align:right;color:#555">PRODUTIVIDADE POR TEMPO DE CICLO OPERACIONAL</td>
        <td><b>{r['prod_ciclo_op']:.0%}</b></td>
      </tr>
      <tr class="prod-row">
        <td colspan="9" style="text-align:right;color:#555">PRODUTIVIDADE POR TEMPO DE CICLO TOTAL</td>
        <td><b>{r['prod_ciclo_tot']:.0%}</b></td>
      </tr>
      <tr class="prod-row">
        <td colspan="9" style="text-align:right;color:#555">PRODUTIVIDADE POR TEMPO DE LABOR OPERACIONAL</td>
        <td><b>{r['prod_labor_op']:.0%}</b></td>
      </tr>
      <tr class="prod-destaque">
        <td colspan="9" style="text-align:right">PRODUTIVIDADE POR TEMPO DE LABOR TOTAL</td>
        <td>{r['prod_labor_tot']:.0%}</td>
      </tr>
    </tfoot>
    </table></div>
    """
    return html

# ─────────────────────────────────────────
# GRÁFICO
# ─────────────────────────────────────────
def grafico_resumo(res1, res2=None, label1="Base", label2="Simulado"):
    meses_v, tA1, tB1, tC1, prod1 = [], [], [], [], []
    tA2, tB2, tC2, prod2 = [], [], [], []
    for m, abr in zip(MESES, MESES_ABREV):
        r = res1.get(m)
        if not r: continue
        meses_v.append(abr)
        tA1.append(r["tot_A"]); tB1.append(r["tot_B"]); tC1.append(r["tot_C"])
        prod1.append(r["prod_labor_tot"]*100)
        if res2:
            r2 = res2.get(m)
            if r2:
                tA2.append(r2["tot_A"]); tB2.append(r2["tot_B"]); tC2.append(r2["tot_C"])
                prod2.append(r2["prod_labor_tot"]*100)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    opacity = 0.5 if res2 else 1.0
    fig.add_trace(go.Bar(name=f"Turno A ({label1})", x=meses_v, y=tA1,
        marker_color="#2E7D32", opacity=opacity,
        text=tA1, textposition="inside", textfont=dict(color="white", size=10)), secondary_y=False)
    fig.add_trace(go.Bar(name=f"Turno B ({label1})", x=meses_v, y=tB1,
        marker_color="#F9A825", opacity=opacity,
        text=tB1, textposition="inside", textfont=dict(size=10)), secondary_y=False)
    fig.add_trace(go.Bar(name=f"Turno C ({label1})", x=meses_v, y=tC1,
        marker_color="#1565C0", opacity=opacity,
        text=tC1, textposition="inside", textfont=dict(color="white", size=10)), secondary_y=False)

    if res2 and tA2:
        fig.add_trace(go.Bar(name=f"Turno A ({label2})", x=meses_v, y=tA2,
            marker_color="#66BB6A", marker_line=dict(color="#2E7D32",width=2),
            text=tA2, textposition="inside", textfont=dict(color="white",size=10)), secondary_y=False)
        fig.add_trace(go.Bar(name=f"Turno B ({label2})", x=meses_v, y=tB2,
            marker_color="#FFD54F", marker_line=dict(color="#F9A825",width=2),
            text=tB2, textposition="inside", textfont=dict(size=10)), secondary_y=False)
        fig.add_trace(go.Bar(name=f"Turno C ({label2})", x=meses_v, y=tC2,
            marker_color="#64B5F6", marker_line=dict(color="#1565C0",width=2),
            text=tC2, textposition="inside", textfont=dict(color="white",size=10)), secondary_y=False)

    fig.add_trace(go.Scatter(
        name=f"Labor Total ({label1})", x=meses_v, y=prod1,
        mode="lines+markers+text",
        marker=dict(color="#CC0000", size=10),
        line=dict(color="#CC0000", width=2),
        text=[f"{p:.0f}%" for p in prod1], textposition="top center",
    ), secondary_y=True)

    if res2 and prod2:
        fig.add_trace(go.Scatter(
            name=f"Labor Total ({label2})", x=meses_v, y=prod2,
            mode="lines+markers+text",
            marker=dict(color="#FF6D00", size=10, symbol="diamond"),
            line=dict(color="#FF6D00", width=2, dash="dot"),
            text=[f"{p:.0f}%" for p in prod2], textposition="bottom center",
        ), secondary_y=True)

    fig.update_layout(
        barmode="group",
        title=dict(text="MÃO-DE-OBRA POR TURNO", font=dict(size=16, color="#1E3A5F")),
        yaxis_title="Nº Funcionários",
        yaxis2=dict(title="Produtividade (%)", tickformat=".0f", ticksuffix="%", range=[0,100]),
        legend=dict(orientation="h", y=-0.25, font=dict(size=11)),
        height=440, plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor="#E0E7EF"),
    )
    return fig

# ─────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────
def exportar(resultados):
    output = BytesIO()
    wb = openpyxl.Workbook()

    borda = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    def estilo(cell, bg="FFFFFF", fg="000000", bold=False, fmt=None, center=True):
        cell.font = Font(name="Arial", bold=bold, color=fg, size=9)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
        cell.border = borda
        if fmt: cell.number_format = fmt

    # Aba resumo
    ws = wb.active
    ws.title = "RESUMO MO"
    ws.row_dimensions[1].height = 20
    headers_res = ["Mês","Dias","Turno A","Turno B","Turno C","Total","Ciclo Op.","Ciclo Total","Labor Op.","Labor Total ★"]
    for i, h in enumerate(headers_res, 1):
        c = ws.cell(row=1, column=i, value=h)
        estilo(c, "1E3A5F", "FFFFFF", True)
    ws.column_dimensions["A"].width = 13

    for row_i, (m, abr) in enumerate(zip(MESES, MESES_ABREV), 2):
        r = resultados.get(m)
        bg = "EAF3FB" if row_i % 2 == 0 else "FFFFFF"
        if r is None:
            vals = [abr, 0,"-","-","-","-","-","-","-","-"]
        else:
            vals = [abr, r["dias"], r["tot_A"], r["tot_B"], r["tot_C"], r["total"],
                    r["prod_ciclo_op"], r["prod_ciclo_tot"],
                    r["prod_labor_op"], r["prod_labor_tot"]]
        for col_i, v in enumerate(vals, 1):
            c = ws.cell(row=row_i, column=col_i, value=v)
            fmt = "0%" if col_i >= 7 and isinstance(v, float) else None
            estilo(c, bg if col_i < 9 else ("FFF9C4" if col_i == 10 else bg), fmt=fmt)
        ws.row_dimensions[row_i].height = 16

    # Aba por mês (estilo planilha original)
    for mes in MESES:
        r = resultados.get(mes)
        if r is None: continue
        ws_m = wb.create_sheet(mes[:10])
        hA = HORAS_TURNO["A"] * r["dias"]
        hB = HORAS_TURNO["B"] * r["dias"]
        hC = HORAS_TURNO["C"] * r["dias"]

        # Cabeçalho
        for col, txt in [(1,""), (2,"TURNO A"),(3,"TURNO B"),(4,"TURNO C"),
                         (5,"TURNO A"),(6,"TURNO B"),(7,"TURNO C"),
                         (8,"TURNO A"),(9,"TURNO B"),(10,"TURNO C")]:
            c = ws_m.cell(row=1, column=col, value=txt)
            estilo(c, "1E3A5F","FFFFFF", True)
        for col, txt in [(2,f"{HORAS_TURNO['A']}h"),(3,f"{HORAS_TURNO['B']}h"),(4,f"{HORAS_TURNO['C']}h")]:
            c = ws_m.cell(row=2, column=col, value=txt)
            estilo(c, "2D6A9F","FFFFFF", True)
        ws_m.merge_cells("A1:A2")
        ws_m.cell(row=1, column=1).value = mes.upper()
        estilo(ws_m.cell(row=1, column=1), "1E3A5F","FFFFFF", True)

        row_i = 3
        def cor_bg(v):
            if v > 1.0: return "FFCDD2"
            if v >= 0.85: return "FFF9C4"
            return "C8E6C9"

        for _, row in r["centros"].iterrows():
            cells_data = [
                (row.centro, "FFFFFF", False),
                (f"{row.ocup_A:.0%}", cor_bg(row.ocup_A), True),
                (f"{row.ocup_B:.0%}", cor_bg(row.ocup_B), True),
                (f"{row.ocup_C:.0%}", cor_bg(row.ocup_C), True),
                (row.ativo_A, "B3E5FC" if row.ativo_A else "FFF9C4", True),
                (row.ativo_B, "B3E5FC" if row.ativo_B else "FFF9C4", True),
                (row.ativo_C, "B3E5FC" if row.ativo_C else "FFF9C4", True),
                (f"{row.horas_disp_A:.2f}" if row.ativo_A else "0", "B3E5FC" if row.ativo_A else "FFF9C4", True),
                (f"{row.horas_disp_B:.2f}" if row.ativo_B else "0", "B3E5FC" if row.ativo_B else "FFF9C4", True),
                (f"{row.horas_disp_C:.2f}" if row.ativo_C else "0", "B3E5FC" if row.ativo_C else "FFF9C4", True),
            ]
            for col_i, (val, bg, ctr) in enumerate(cells_data, 1):
                c = ws_m.cell(row=row_i, column=col_i, value=val)
                estilo(c, bg, center=ctr)
            row_i += 1
            ws_m.row_dimensions[row_i-1].height = 15

        # Suporte
        sup = r["suporte"]
        for nome, key in [("TOTAL DE OPERADORES", None),
                          ("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),
                          ("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador"),
                          ("TOTAL POR TURNO", None),("TOTAL FUNCIONÁRIOS", None)]:
            bg_row = "FF8A80" if "TOTAL" in nome else "FFFFFF"
            bold = "TOTAL" in nome
            c = ws_m.cell(row=row_i, column=1, value=nome)
            estilo(c, bg_row, bold=bold, center=False)
            if key:
                s = sup[key]
                for col_idx, (turno, val) in enumerate([(5,s["A"]),(6,s["B"]),(7,s["C"])], 5):
                    cell = ws_m.cell(row=row_i, column=col_idx, value=val)
                    estilo(cell, "B3E5FC" if val else "FFF9C4", bold=bold)
                    hval = val * [HORAS_TURNO["A"],HORAS_TURNO["B"],HORAS_TURNO["C"]][col_idx-5] * r["dias"]
                    hcell = ws_m.cell(row=row_i, column=col_idx+3, value=f"{hval:.2f}" if val else "0")
                    estilo(hcell, "B3E5FC" if val else "FFF9C4", bold=bold)
            elif nome == "TOTAL DE OPERADORES":
                for col_idx, val in [(5,r["op_A"]),(6,r["op_B"]),(7,r["op_C"])]:
                    cell = ws_m.cell(row=row_i, column=col_idx, value=val)
                    estilo(cell, "FF8A80", bold=True)
            elif nome == "TOTAL POR TURNO":
                for col_idx, val in [(5,r["tot_A"]),(6,r["tot_B"]),(7,r["tot_C"])]:
                    cell = ws_m.cell(row=row_i, column=col_idx, value=val)
                    estilo(cell, "FF8A80", bold=True)
                for col_idx, val in [(8,r["tot_A"]*hA),(9,r["tot_B"]*hB),(10,r["tot_C"]*hC)]:
                    cell = ws_m.cell(row=row_i, column=col_idx, value=f"{val:.2f}")
                    estilo(cell, "FF8A80", bold=True)
            elif nome == "TOTAL FUNCIONÁRIOS":
                cell = ws_m.cell(row=row_i, column=4, value=r["total"])
                estilo(cell, "FF8A80", bold=True)
                tot_h = r["tot_A"]*hA + r["tot_B"]*hB + r["tot_C"]*hC
                cell2 = ws_m.cell(row=row_i, column=8, value=f"{tot_h:.2f}")
                estilo(cell2, "FF8A80", bold=True)
            row_i += 1
            ws_m.row_dimensions[row_i-1].height = 15

        # Produtividades
        row_i += 1
        for nome, val, destaque in [
            ("PRODUTIVIDADE POR TEMPO DE CICLO OPERACIONAL", r["prod_ciclo_op"], False),
            ("PRODUTIVIDADE POR TEMPO DE CICLO TOTAL",       r["prod_ciclo_tot"], False),
            ("PRODUTIVIDADE POR TEMPO DE LABOR OPERACIONAL", r["prod_labor_op"],  False),
            ("PRODUTIVIDADE POR TEMPO DE LABOR TOTAL",       r["prod_labor_tot"], True),
        ]:
            bg = "FFF9C4" if destaque else "FFFFFF"
            fg = "E65100" if destaque else "000000"
            c1 = ws_m.cell(row=row_i, column=8, value=nome)
            estilo(c1, bg, fg, destaque, center=False)
            ws_m.merge_cells(f"H{row_i}:I{row_i}")
            c2 = ws_m.cell(row=row_i, column=10, value=val)
            estilo(c2, bg, fg, destaque, "0%")
            row_i += 1

        for col, w in enumerate([14,8,8,8,8,8,8,10,10,10], 1):
            ws_m.column_dimensions[get_column_letter(col)].width = w

    output.seek(0)
    wb.save(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────
# INTERFACE PRINCIPAL
# ─────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🏭 Calculadora de Recursos — Usinagem</h1>
  <p>Faça upload da planilha de inputs para calcular headcount, ocupação por turno e indicadores de produtividade.</p>
</div>
""", unsafe_allow_html=True)

# ── Instrução de upload ──
with st.expander("📋 Como preparar o arquivo para upload", expanded=False):
    st.markdown("""
**O app lê automaticamente 5 abas do seu arquivo `.xlsm` ou `.xlsx`:**

| Aba | O que contém | Obrigatório? |
|---|---|---|
| `INPUT_PMP` | Demanda mensal por modelo + dias trabalhados por mês | ✅ Sim |
| `IMPUTTEMPO` | Tempo de ciclo e tempo de labor por centro/peça | ✅ Sim |
| `IMPUTDISTRIBUIÇÃO` | Divisão de carga, divisão de volume e disponibilidade por centro/peça | ✅ Sim |
| `IMPUTAPLICAÇÃO` | Matriz indicando quais modelos passam em cada centro/peça (0 ou 1) | ✅ Sim |
| `IMPUTTURNOS` | Horas acumuladas por turno desde o início do dia | ✅ Sim |

> **Não é necessário ter as abas mensais** (NovFY26, DezFY26, etc.) — o app recalcula tudo a partir dos inputs acima.
> As horas de duração de cada turno usadas são: **Turno A = 8,80h | Turno B = 8,23h | Turno C = 7,68h**
> Se o seu arquivo tiver valores diferentes, avise e ajustamos.
    """)

uploaded = st.file_uploader(
    "Selecione o arquivo Excel (.xlsm ou .xlsx)",
    type=["xlsm","xlsx"],
    help="O arquivo deve conter as 5 abas de input descritas acima."
)

if not uploaded:
    st.info("👆 Faça upload do arquivo para começar.")
    st.stop()

file_bytes = uploaded.read()

with st.spinner("Lendo e processando planilha..."):
    try:
        pmp,  dias = read_pmp(file_bytes)
        tempo       = read_tempo(file_bytes)
        dist        = read_dist(file_bytes)
        aplic       = read_aplic(file_bytes)
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        st.stop()

st.success(f"✅ Arquivo lido com sucesso — {len(aplic)} combinações centro/peça/modelo | {pmp.modelo.nunique()} modelos | {pmp.mes.nunique()} meses com demanda")

# ── Validações ──
erros, alertas, oks = validar(pmp, tempo, dist, aplic, dias)

with st.expander(
    f"{'🔴 ' + str(len(erros)) + ' erro(s) crítico(s)' if erros else ''}"
    f"{'  ⚠️ ' + str(len(alertas)) + ' aviso(s)' if alertas else ''}"
    f"{'✅ Tudo certo nos inputs' if not erros and not alertas else ''}",
    expanded=bool(erros)
):
    for e in erros:
        st.markdown(f'<div class="aviso-erro">🔴 <b>ERRO CRÍTICO:</b> {e}</div>', unsafe_allow_html=True)
    for a in alertas:
        st.markdown(f'<div class="aviso-alerta">⚠️ {a}</div>', unsafe_allow_html=True)
    for o in oks:
        st.markdown(f'<div class="aviso-ok">✅ {o}</div>', unsafe_allow_html=True)

if erros:
    st.error("Corrija os erros acima antes de continuar.")
    st.stop()

# ── Tabs ──
tab1, tab2, tab3, tab4 = st.tabs(["📊 Resultados", "🔧 Simulador", "🔍 Detalhes por etapa", "📥 Exportar"])

# ══════════════════════════════════════════
# TAB 1 — RESULTADOS
# ══════════════════════════════════════════
with tab1:
    res_base = calcular(pmp, tempo, dist, aplic, dias)

    st.plotly_chart(grafico_resumo(res_base), use_container_width=True)

    st.markdown('<div class="section-title">Resultado detalhado por mês</div>', unsafe_allow_html=True)
    mes_sel = st.selectbox("Selecione o mês", [m for m in MESES if res_base.get(m)], key="mes_res")

    if mes_sel and res_base.get(mes_sel):
        r = res_base[mes_sel]
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Total funcionários", r["total"])
        c2.metric("Turno A", r["tot_A"])
        c3.metric("Turno B", r["tot_B"])
        c4.metric("Turno C", r["tot_C"])
        st.markdown(tabela_html(r, mes_sel), unsafe_allow_html=True)

# ══════════════════════════════════════════
# TAB 2 — SIMULADOR
# ══════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Simulador de headcount</div>', unsafe_allow_html=True)
    st.caption("Ajuste o número de operadores por centro e compare cenários.")

    col_cfg1, col_cfg2 = st.columns([2,1])
    with col_cfg1:
        mes_sim = st.selectbox("Mês para simular", MESES, key="mes_sim")
    with col_cfg2:
        nome_cenario = st.text_input("Nome do cenário", value="Cenário simulado")

    res_base_sim = calcular(pmp, tempo, dist, aplic, dias)
    r_orig = res_base_sim.get(mes_sim)

    if r_orig is None:
        st.warning("Mês sem dados disponíveis.")
    else:
        if "overrides" not in st.session_state:
            st.session_state.overrides = {}
        if mes_sim not in st.session_state.overrides:
            st.session_state.overrides[mes_sim] = {}

        centros_list = sorted(r_orig["centros"].centro.tolist())

        st.markdown('<div class="subsection">Ajuste operadores por centro e turno</div>', unsafe_allow_html=True)
        st.caption("Escolha quantos operadores alocar em cada turno (0, 1 ou mais). O padrão calculado aparece abaixo do campo.")

        cols_h = st.columns([3,1,1,1])
        cols_h[0].markdown("**Centro — ocupação A / B / C**")
        cols_h[1].markdown("**Turno A**")
        cols_h[2].markdown("**Turno B**")
        cols_h[3].markdown("**Turno C**")

        for cen in centros_list:
            row_cen = r_orig["centros"][r_orig["centros"].centro == cen].iloc[0]
            ov = st.session_state.overrides[mes_sim].get(cen, {})
            def_A = int(ov.get("A", row_cen.ativo_A))
            def_B = int(ov.get("B", row_cen.ativo_B))
            def_C = int(ov.get("C", row_cen.ativo_C))

            c0, c1, c2, c3 = st.columns([3,1,1,1])
            with c0:
                cA_label = "🔴" if row_cen.ocup_A>1 else ("🟡" if row_cen.ocup_A>=0.85 else "🟢")
                cB_label = "🔴" if row_cen.ocup_B>1 else ("🟡" if row_cen.ocup_B>=0.85 else "🟢")
                cC_label = "🔴" if row_cen.ocup_C>1 else ("🟡" if row_cen.ocup_C>=0.85 else "🟢")
                st.markdown(f"`{cen}` &nbsp; {cA_label}{row_cen.ocup_A:.0%} / {cB_label}{row_cen.ocup_B:.0%} / {cC_label}{row_cen.ocup_C:.0%}")
            with c1:
                vA = st.number_input("", min_value=0, max_value=5, value=def_A,
                    key=f"sim_{mes_sim}_{cen}_A", label_visibility="collapsed",
                    help=f"Calculado: {row_cen.ativo_A}")
            with c2:
                vB = st.number_input("", min_value=0, max_value=5, value=def_B,
                    key=f"sim_{mes_sim}_{cen}_B", label_visibility="collapsed",
                    help=f"Calculado: {row_cen.ativo_B}")
            with c3:
                vC = st.number_input("", min_value=0, max_value=5, value=def_C,
                    key=f"sim_{mes_sim}_{cen}_C", label_visibility="collapsed",
                    help=f"Calculado: {row_cen.ativo_C}")

            st.session_state.overrides[mes_sim][cen] = {"A": vA, "B": vB, "C": vC}

        # Calcular simulado
        res_sim = calcular(pmp, tempo, dist, aplic, dias, st.session_state.overrides)
        r_new   = res_sim.get(mes_sim)

        if r_new:
            st.markdown("---")
            st.markdown('<div class="subsection">Impacto da simulação</div>', unsafe_allow_html=True)
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Total funcionários", r_new["total"], delta=f"{r_new['total']-r_orig['total']:+d}")
            c2.metric("Turno A", r_new["tot_A"], delta=f"{r_new['tot_A']-r_orig['tot_A']:+d}")
            c3.metric("Turno B", r_new["tot_B"], delta=f"{r_new['tot_B']-r_orig['tot_B']:+d}")
            c4.metric("Turno C", r_new["tot_C"], delta=f"{r_new['tot_C']-r_orig['tot_C']:+d}")
            c1b,c2b,c3b,c4b = st.columns(4)
            c1b.metric("Labor Total ★", f"{r_new['prod_labor_tot']:.0%}",
                       delta=f"{r_new['prod_labor_tot']-r_orig['prod_labor_tot']:+.1%}")
            c2b.metric("Labor Operacional", f"{r_new['prod_labor_op']:.0%}")
            c3b.metric("Ciclo Total", f"{r_new['prod_ciclo_tot']:.0%}")
            c4b.metric("Ciclo Operacional", f"{r_new['prod_ciclo_op']:.0%}")

            st.markdown('<div class="subsection">Tabela detalhada — simulado</div>', unsafe_allow_html=True)
            st.markdown(tabela_html(r_new, f"{mes_sim} — {nome_cenario}"), unsafe_allow_html=True)

            # Comparação gráfica
            st.markdown('<div class="subsection">Comparativo anual — base vs simulado</div>', unsafe_allow_html=True)
            st.plotly_chart(
                grafico_resumo(res_base_sim, res_sim, "Base", nome_cenario),
                use_container_width=True
            )

            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                xlsx_sim = exportar(res_sim)
                st.download_button("📥 Baixar Excel do cenário simulado",
                    data=xlsx_sim, file_name=f"simulado_{mes_sim}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_exp2:
                if st.button("🔄 Resetar simulação"):
                    st.session_state.overrides = {}
                    st.rerun()

# ══════════════════════════════════════════
# TAB 3 — DETALHES
# ══════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">Detalhes por etapa do cálculo</div>', unsafe_allow_html=True)
    etapa = st.radio("Etapa", [
        "Passo 2 — INPUT_PMP normalizado",
        "Passo 3 — IMPUTAPLICAÇÃO normalizado",
        "Passo 4 — JOIN aplicacao × pmp",
        "Passo 7 — Minutos calculados por linha",
        "Passo 8 — Totais por centro + mês",
        "Passo 10 — % ocupação por centro",
    ], horizontal=True)
    mes_det = st.selectbox("Mês", MESES, key="mes_det")

    if etapa == "Passo 2 — INPUT_PMP normalizado":
        st.dataframe(pmp[pmp.mes==mes_det].reset_index(drop=True), use_container_width=True, hide_index=True)
    elif etapa == "Passo 3 — IMPUTAPLICAÇÃO normalizado":
        st.dataframe(aplic.head(300), use_container_width=True, hide_index=True)
        st.caption(f"Total: {len(aplic)} combinações ativas (somente onde flag = 1)")
    elif etapa == "Passo 4 — JOIN aplicacao × pmp":
        p4 = aplic.merge(pmp[pmp.mes==mes_det], on="modelo")
        st.dataframe(p4.head(300), use_container_width=True, hide_index=True)
        st.caption(f"{len(p4)} linhas para {mes_det}")
    elif etapa == "Passo 7 — Minutos calculados por linha":
        p7 = (aplic.merge(pmp[pmp.mes==mes_det], on="modelo")
                   .merge(tempo, on=["centro","peca"])
                   .merge(dist,  on=["centro","peca"]))
        p7["indice_ciclo"] = (p7.t_ciclo * p7.div_carga * p7.div_volume) / p7.disponib
        p7["min_ciclo"]    = (p7.indice_ciclo * p7.qtd).round(1)
        p7["min_labor"]    = (p7.t_labor * p7.div_carga * p7.qtd).round(1)
        st.dataframe(p7.head(300), use_container_width=True, hide_index=True)
        st.caption(f"{len(p7)} linhas para {mes_det}")
    elif etapa == "Passo 8 — Totais por centro + mês":
        p7 = (aplic.merge(pmp[pmp.mes==mes_det], on="modelo")
                   .merge(tempo, on=["centro","peca"])
                   .merge(dist,  on=["centro","peca"]))
        p7["min_ciclo"] = (p7.t_ciclo * p7.div_carga * p7.div_volume / p7.disponib) * p7.qtd
        p7["min_labor"] = p7.t_labor * p7.div_carga * p7.qtd
        p8 = p7.groupby("centro")[["min_ciclo","min_labor"]].sum().reset_index()
        p8["horas_ciclo"] = (p8.min_ciclo/60).round(1)
        p8["horas_labor"] = (p8.min_labor/60).round(1)
        st.dataframe(p8, use_container_width=True, hide_index=True)
    elif etapa == "Passo 10 — % ocupação por centro":
        res_det = calcular(pmp, tempo, dist, aplic, dias)
        r = res_det.get(mes_det)
        if r:
            df_oc = r["centros"][["centro","ocup_A","ocup_B","ocup_C","ativo_A","ativo_B","ativo_C"]].copy()
            def fmt_pct(v):
                emoji = "🔴" if v>1 else ("🟡" if v>=0.85 else "🟢")
                return f"{emoji} {v:.0%}"
            df_oc["ocup_A"] = df_oc.ocup_A.map(fmt_pct)
            df_oc["ocup_B"] = df_oc.ocup_B.map(fmt_pct)
            df_oc["ocup_C"] = df_oc.ocup_C.map(fmt_pct)
            st.dataframe(df_oc, use_container_width=True, hide_index=True)
            st.caption("🟢 < 85%   🟡 85–100%   🔴 > 100%   |   Threshold ativação: A>40% · B se A>106% · C se B>100%")

# ══════════════════════════════════════════
# TAB 4 — EXPORTAR
# ══════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">Exportar resultados</div>', unsafe_allow_html=True)
    st.markdown("""
O arquivo exportado contém:
- **Aba RESUMO MO** — visão anual consolidada com todos os meses, turnos e 4 indicadores
- **Uma aba por mês** — tabela completa no mesmo formato da planilha original, com % de ocupação colorida, flags de turno, horas disponíveis, funções de suporte e indicadores de produtividade
    """)
    res_exp = calcular(pmp, tempo, dist, aplic, dias)
    xlsx_bytes = exportar(res_exp)
    st.download_button(
        "📥 Baixar Excel completo",
        data=xlsx_bytes,
        file_name="resultado_usinagem.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
