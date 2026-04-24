import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import math

st.set_page_config(page_title="Calculadora de Recursos — Usinagem", layout="wide", page_icon="🏭")

JD_VERDE       = "#367C2B"
JD_VERDE_ESC   = "#1F4D19"
JD_AMARELO     = "#FFDE00"
JD_AMARELO_ESC = "#C9A800"
JD_TEXTO       = "#1A1A1A"

st.markdown("""
<style>
.jd-header{background:#1F4D19;padding:16px 24px;border-radius:10px;border-left:6px solid #FFDE00;margin-bottom:20px;}
.jd-header h1{color:#FFFFFF;margin:0;font-size:21px;font-weight:700;}
.jd-header p{color:#b8d4b4;margin:4px 0 0;font-size:12px;}
.jd-section{font-size:15px;font-weight:700;color:#FFDE00;border-left:4px solid #FFDE00;padding-left:10px;margin:22px 0 10px;}
.jd-sub{font-size:12px;font-weight:600;color:#7BC67A;margin:12px 0 4px;text-transform:uppercase;letter-spacing:.04em;}
.aviso-erro{background:#3D0000;border-left:4px solid #FF5252;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#FF8A80;}
.aviso-warn{background:#3D2D00;border-left:4px solid #FFDE00;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#FFE57F;}
.aviso-ok{background:#003D10;border-left:4px solid #69F0AE;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#B9F6CA;}
.formula-box{background:#0D1117;color:#A8E6A3;font-family:monospace;padding:10px 14px;border-radius:6px;font-size:12px;line-height:1.8;border-left:3px solid #FFDE00;}
.mem-step{background:#1A1A2E;border:1px solid #444;border-radius:8px;padding:12px 16px;margin:6px 0;color:#FAFAFA;}
.mem-step b{color:#FFDE00;}
.mem-step .step-num{background:#FFDE00;color:#1F4D19;border-radius:50%;width:24px;height:24px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;margin-right:8px;}
.log-line{font-family:monospace;font-size:11px;color:#AAAAAA;padding:1px 0;}
.log-ok{color:#69F0AE;} .log-warn{color:#FFDE00;} .log-err{color:#FF5252;}
.cenario-card{background:#1A1A2E;border:1.5px solid #333;border-radius:10px;padding:14px 16px;margin:6px 0;color:#FAFAFA;}
.cenario-card b{color:#FFDE00;}
/* KPI CARDS */
.kpi-card{background:linear-gradient(135deg,#1A2E1A 0%,#0D1F0D 100%);border:1px solid #2A4A2A;border-radius:12px;padding:16px 18px;margin:4px 0;position:relative;overflow:hidden;}
.kpi-card::before{content:'';position:absolute;top:0;left:0;width:4px;height:100%;background:#FFDE00;}
.kpi-card .kpi-icon{font-size:22px;margin-bottom:4px;}
.kpi-card .kpi-label{font-size:10px;color:#7BC67A;text-transform:uppercase;letter-spacing:.06em;font-weight:600;}
.kpi-card .kpi-value{font-size:26px;font-weight:800;color:#FFFFFF;line-height:1.1;}
.kpi-card .kpi-sub{font-size:11px;color:#AAAAAA;margin-top:2px;}
.kpi-card.destaque{border-color:#FFDE00;}
.kpi-card.destaque::before{background:#367C2B;width:100%;height:3px;top:0;left:0;}
/* MES BARRAS */
.mes-row{display:flex;align-items:center;gap:10px;padding:7px 12px;border-radius:8px;margin:3px 0;background:#131313;border:1px solid #222;}
.mes-nome{font-size:12px;font-weight:700;color:#FFDE00;min-width:36px;}
.mes-bar{flex:1;height:8px;background:#1A2E1A;border-radius:4px;overflow:hidden;}
.mes-bar-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,#367C2B,#FFDE00);}
.mes-num{font-size:12px;color:#FFFFFF;font-weight:700;min-width:28px;text-align:right;}
.mes-labor{font-size:11px;min-width:44px;text-align:right;}
/* TURNO PILLS */
.tp{display:inline-block;border-radius:20px;padding:1px 8px;font-size:10px;font-weight:700;margin:0 1px;}
.tpA{background:#1F4D19;color:#92D050;} .tpB{background:#3D2D00;color:#FFDE00;} .tpC{background:#0D2040;color:#00B0F0;}
/* GAUGE */
.gauge-wrap{text-align:center;padding:6px 0;}
</style>
""", unsafe_allow_html=True)

MESES       = ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
               "Maio","Junho","Julho","Agosto","Setembro","Outubro"]
MESES_ABREV = ["NOV","DEZ","JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT"]

COL_MAP = {
    "IMPUTTEMPO": {
        "centro":  ["Máquina","MÁQUINA","maquina","Centro","CENTRO"],
        "peca":    ["PEÇA","Peça","peca","PECA"],
        "t_ciclo": ["Tempo\nCiclo\n(min)","Tempo Ciclo (min)","T.CICLO","t_ciclo","CICLO"],
        "t_labor": ["Tempo\nLabor\n(min)","Tempo Labor (min)","T.LABOR","t_labor","LABOR"],
    },
    "IMPUTDISTRIBUIÇÃO": {
        "centro":     ["Máquina","MÁQUINA","maquina","Centro","CENTRO"],
        "peca":       ["PEÇA","Peça","peca","PECA"],
        "div_carga":  ["Divisão\nCarga\nENTRE\nMÁQUINAS","Div Carga","DIV_CARGA","div_carga"],
        "vol_int":    ["Vol.\nInterna","Vol. Interna","VOL_INT","vol_int","VOL. INTERNA","Volume Interna"],
        "div_volume": ["Divisão \nde\nVolume\nENTRE\nPEÇAS","Div Volume","DIV_VOLUME","div_volume"],
        "disponib":   ["Disponi-\nbilidade","Disponibilidade","DISPONIB","disponib"],
    },
}

ABA_FORMATOS = {
    "INPUT_PMP": "**INPUT_PMP** — Linha 1: dias trabalhados (colunas B→M = Nov→Out). Linhas 3+: modelos, colunas B→M = qtd peças.",
    "IMPUTTEMPO": "**IMPUTTEMPO** — Cabeçalho linha 1. Colunas: `Máquina`, `PEÇA`, `Tempo Ciclo (min)`, `Tempo Labor (min)`.",
    "IMPUTDISTRIBUIÇÃO": "**IMPUTDISTRIBUIÇÃO** — Cabeçalho linha 1. Colunas: `Máquina`, `PEÇA`, `Divisão Carga`, `Vol. Interna`, `Divisão de Volume`, `Disponibilidade`.",
    "IMPUTAPLICAÇÃO": "**IMPUTAPLICAÇÃO** — Cabeçalho linha 1. Col A=Centro, Col B=PEÇA, depois colunas por modelo (MODELO01…).",
    "IMPUTTURNOS": "**IMPUTTURNOS** — Linha 1: horas acumuladas. B1=Turno A, C1=Turno B, D1=Turno C.",
}

def find_col(df, candidates, aba, campo):
    for c in candidates:
        if c in df.columns:
            return c
    idx_fallback = {"centro":0,"peca":1,"t_ciclo":5,"t_labor":6,"div_carga":7,"vol_int":8,"div_volume":9,"disponib":10}
    if campo in idx_fallback:
        idx = idx_fallback[campo]
        if idx < len(df.columns):
            st.session_state.setdefault("log_leitura",[]).append(
                f"⚠️ [{aba}] Campo '{campo}' não encontrado — usando coluna {idx} ({df.columns[idx]}) como fallback")
            return df.columns[idx]
    raise ValueError(f"[{aba}] Campo '{campo}' não encontrado. Colunas: {list(df.columns)}")

def verificar_abas(fb):
    try:
        wb = openpyxl.load_workbook(BytesIO(fb), read_only=True, data_only=True)
        abas = set(wb.sheetnames); wb.close()
    except: abas = set()
    return {a: a in abas for a in ["INPUT_PMP","IMPUTTEMPO","IMPUTDISTRIBUIÇÃO","IMPUTAPLICAÇÃO","IMPUTTURNOS"]}

def read_pmp(fb, log):
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name='INPUT_PMP', header=None)
    except Exception as e:
        raise ValueError(f"Não foi possível ler INPUT_PMP: {e}\n\n{ABA_FORMATOS['INPUT_PMP']}")
    log.append(f"✅ INPUT_PMP lido: {df.shape[0]}L × {df.shape[1]}C")
    dias = {}
    for i, m in enumerate(MESES, 1):
        v = df.iloc[0, i] if i < df.shape[1] else None
        try: dias[m] = int(v) if pd.notna(v) else 0
        except: dias[m] = 0
    log.append(f"   Dias: { {m:d for m,d in dias.items() if d>0} }")
    rows = []
    for r in range(2, len(df)):
        modelo = df.iloc[r, 0]
        if pd.isna(modelo): continue
        for i, m in enumerate(MESES, 1):
            v = df.iloc[r, i] if i < df.shape[1] else None
            try: qtd = int(v) if pd.notna(v) else 0
            except: qtd = 0
            if qtd > 0:
                rows.append({"modelo": str(modelo).strip(), "mes": m, "qtd": qtd})
    log.append(f"   {len(rows)} registros com qtd>0")
    return pd.DataFrame(rows), dias

def read_turnos(fb):
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTTURNOS', header=None)
        hA = float(df.iloc[0,1]) if pd.notna(df.iloc[0,1]) else 7.5
        hB = float(df.iloc[0,2]) if pd.notna(df.iloc[0,2]) else 14.25
        hC = float(df.iloc[0,3]) if pd.notna(df.iloc[0,3]) else 19.5
        return {"A": hA, "B": hB, "C": hC}, True
    except:
        return {"A": 7.5, "B": 14.25, "C": 19.5}, False

def read_tempo(fb, log):
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTTEMPO', header=0)
    except Exception as e:
        raise ValueError(f"Não foi possível ler IMPUTTEMPO: {e}\n\n{ABA_FORMATOS['IMPUTTEMPO']}")
    log.append(f"✅ IMPUTTEMPO lido: {df.shape[0]}L")
    mp = COL_MAP["IMPUTTEMPO"]
    c = {k: find_col(df, v, "IMPUTTEMPO", k) for k,v in mp.items()}
    out = df[[c["centro"],c["peca"],c["t_ciclo"],c["t_labor"]]].copy()
    out.columns = ["centro","peca","t_ciclo","t_labor"]
    out = out.dropna(subset=["centro"])
    log.append(f"   {len(out)} combinações centro+peça")
    return out.copy()

def read_dist(fb, log):
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTDISTRIBUIÇÃO', header=0)
    except Exception as e:
        raise ValueError(f"Não foi possível ler IMPUTDISTRIBUIÇÃO: {e}\n\n{ABA_FORMATOS['IMPUTDISTRIBUIÇÃO']}")
    log.append(f"✅ IMPUTDISTRIBUIÇÃO lido: {df.shape[0]}L")
    mp = COL_MAP["IMPUTDISTRIBUIÇÃO"]
    c = {k: find_col(df, v, "IMPUTDISTRIBUIÇÃO", k) for k,v in mp.items()}
    out = df[[c["centro"],c["peca"],c["div_carga"],c["vol_int"],c["div_volume"],c["disponib"]]].copy()
    out.columns = ["centro","peca","div_carga","vol_int","div_volume","disponib"]
    out["vol_int"] = pd.to_numeric(out["vol_int"], errors="coerce").fillna(1.0)
    out = out.dropna(subset=["centro"])
    log.append(f"   {len(out)} combinações")
    return out.copy()

def read_aplic(fb, log):
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTAPLICAÇÃO', header=0)
    except Exception as e:
        raise ValueError(f"Não foi possível ler IMPUTAPLICAÇÃO: {e}\n\n{ABA_FORMATOS['IMPUTAPLICAÇÃO']}")
    log.append(f"✅ IMPUTAPLICAÇÃO lido: {df.shape[0]}L")
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca"})
    mcols = [c for c in df.columns if str(c).startswith("MODELO")]
    if not mcols:
        raise ValueError(f"IMPUTAPLICAÇÃO: nenhuma coluna 'MODELO...' encontrada. Colunas: {list(df.columns)}")
    log.append(f"   {len(mcols)} modelos")
    melted = df[["centro","peca"]+mcols].melt(id_vars=["centro","peca"], var_name="modelo", value_name="ativo")
    out = melted[melted["ativo"]==1][["centro","peca","modelo"]].reset_index(drop=True)
    log.append(f"   {len(out)} combinações ativas")
    return out

def validar(pmp, tempo, dist, aplic, dias):
    erros, alertas, oks = [], [], []
    chaves_tempo = set(zip(tempo.centro, tempo.peca))
    chaves_dist  = set(zip(dist.centro,  dist.peca))
    chaves_aplic = set(zip(aplic.centro, aplic.peca))
    zero_disp = dist[dist.disponib == 0]
    if len(zero_disp):
        erros.append(f"Disponibilidade=0 em {len(zero_disp)} linha(s)")
    diff_td = chaves_tempo - chaves_dist
    if diff_td:
        erros.append(f"{len(diff_td)} combinações em IMPUTTEMPO sem IMPUTDISTRIBUIÇÃO")
    sem_aplic = chaves_tempo - chaves_aplic
    if sem_aplic:
        alertas.append(f"{len(sem_aplic)} centro+peça sem modelo em IMPUTAPLICAÇÃO")
    modelos_sem = set(pmp.modelo.unique()) - set(aplic.modelo.unique())
    if modelos_sem:
        alertas.append(f"{len(modelos_sem)} modelo(s) com demanda mas sem aplicação")
    merged = tempo.merge(dist, on=["centro","peca"], how="inner")
    labor_maior = merged[merged.t_labor > merged.t_ciclo]
    if len(labor_maior):
        alertas.append(f"{len(labor_maior)} linha(s) com t_labor > t_ciclo")
    for m in MESES:
        qtd_m = pmp[pmp.mes==m].qtd.sum() if len(pmp[pmp.mes==m]) else 0
        if qtd_m > 0 and dias.get(m,0) == 0:
            alertas.append(f"Mês '{m}' tem {int(qtd_m)} peças mas dias=0")
    if not erros and not alertas:
        oks.append("Todos os inputs validados sem inconsistências.")
    return erros, alertas, oks

def calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg,
             horas_efetivas=None, overrides=None, retornar_intermediarios=False):
    if horas_efetivas is None:
        horas_efetivas = horas_turno
    df = (aplic.merge(pmp, on="modelo")
               .merge(tempo, on=["centro","peca"])
               .merge(dist,  on=["centro","peca"]))
    if "vol_int" not in df.columns: df["vol_int"] = 1.0
    df["vol_int"] = pd.to_numeric(df["vol_int"], errors="coerce").fillna(1.0)
    df["indice_ciclo"] = (df.t_ciclo * df.div_carga * df.div_volume * df.vol_int) / df.disponib
    df["min_ciclo"]    = df.indice_ciclo * df.qtd
    df["min_labor"]    = df.t_labor * df.div_carga * df.qtd
    agg = df.groupby(["centro","mes"])[["min_ciclo","min_labor"]].sum().reset_index()
    thr_A = thresholds["A"] / 100
    thr_B = thresholds["B"] / 100
    thr_C = thresholds["C"] / 100
    hA, hB, hC = horas_turno["A"], horas_turno["B"], horas_turno["C"]
    heA, heB, heC = horas_efetivas["A"], horas_efetivas["B"], horas_efetivas["C"]
    resultados = {}
    for mes in MESES:
        d = dias.get(mes, 0)
        if d == 0: resultados[mes] = None; continue
        sub = agg[agg.mes == mes].copy()
        if sub.empty: resultados[mes] = None; continue
        minA = d*hA*60; minB = d*hB*60; minC = d*hC*60
        centros = []
        for _, row in sub.iterrows():
            cen = row.centro; mc = row.min_ciclo; ml = row.min_labor
            pA = mc/minA if minA>0 else 0
            pB = mc/minB if minB>0 else 0
            pC = mc/minC if minC>0 else 0
            aA = 1 if pA > thr_A else 0
            aB = 1 if pA > thr_B else 0
            aC = 1 if pB > thr_C else 0
            if overrides and mes in overrides and cen in overrides[mes]:
                ov = overrides[mes][cen]
                if "A" in ov: aA = ov["A"]
                if "B" in ov: aB = ov["B"]
                if "C" in ov: aC = ov["C"]
            centros.append({
                "centro":cen,"ocup_A":pA,"ocup_B":pB,"ocup_C":pC,
                "ativo_A":aA,"ativo_B":aB,"ativo_C":aC,
                "min_ciclo_total":mc,"min_labor_total":ml,
                "min_disp_A":minA,"min_disp_B":minB,"min_disp_C":minC,
                "horas_ciclo":mc/60,"horas_labor":ml/60,
                "horas_disp_A":d*heA*aA,"horas_disp_B":d*heB*aB,"horas_disp_C":d*heC*aC,
            })
        df_c = pd.DataFrame(centros)
        op_A = int(df_c.ativo_A.sum()); op_B = int(df_c.ativo_B.sum()); op_C = int(df_c.ativo_C.sum())
        def get_sup(key, t, op_count):
            cfg = suporte_cfg[key]
            if op_count == 0: return 0
            if cfg["modo"] == "auto":
                defaults = {"lavadora":{"A":1,"B":1,"C":0},"gravacao":{"A":1,"B":1,"C":0},
                            "preset":{"A":2,"B":1,"C":1},"coringa":{"A":1,"B":0,"C":0},
                            "facilitador":{"A":1,"B":1,"C":0}}
                return defaults[key][t]
            return cfg[t]
        lav={t:get_sup("lavadora",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        gra={t:get_sup("gravacao",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        pre={t:get_sup("preset",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        cor={t:get_sup("coringa",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        fac={t:get_sup("facilitador",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        tot_A = op_A+lav["A"]+gra["A"]+pre["A"]+cor["A"]+fac["A"]
        tot_B = op_B+lav["B"]+gra["B"]+pre["B"]+cor["B"]+fac["B"]
        tot_C = op_C+lav["C"]+gra["C"]+pre["C"]+cor["C"]+fac["C"]
        total = tot_A+tot_B+tot_C
        h_ciclo = float(df_c.horas_ciclo.sum()); h_labor = float(df_c.horas_labor.sum())
        h_ativos = float((df_c.horas_disp_A+df_c.horas_disp_B+df_c.horas_disp_C).sum())
        h_todos  = tot_A*d*heA+tot_B*d*heB+tot_C*d*heC
        resultados[mes] = {
            "centros":df_c,
            "op_A":op_A,"op_B":op_B,"op_C":op_C,
            "tot_A":tot_A,"tot_B":tot_B,"tot_C":tot_C,"total":total,
            "suporte":{"lavadora":lav,"gravacao":gra,"preset":pre,"coringa":cor,"facilitador":fac},
            "h_ciclo":h_ciclo,"h_labor":h_labor,"h_ativos":h_ativos,"h_todos":h_todos,
            "prod_ciclo_op":  h_ciclo/h_ativos if h_ativos>0 else 0,
            "prod_ciclo_tot": h_ciclo/h_todos  if h_todos>0 else 0,
            "prod_labor_op":  h_labor/h_ativos if h_ativos>0 else 0,
            "prod_labor_tot": h_labor/h_todos  if h_todos>0 else 0,
            "dias":d,"hA":hA,"hB":hB,"hC":hC,"heA":heA,"heB":heB,"heC":heC,
            "thr_A":thr_A,"thr_B":thr_B,"thr_C":thr_C,
            "minA":d*hA*60,"minB":d*hB*60,"minC":d*hC*60,
        }
    if retornar_intermediarios:
        return resultados, df, agg
    return resultados

def agregar_ano(res_dict, meses_lista):
    rr = [res_dict.get(m) for m in meses_lista if res_dict.get(m)]
    if not rr: return None
    n = len(rr)
    sh_ciclo = sum(r["h_ciclo"]  for r in rr)
    sh_labor = sum(r["h_labor"]  for r in rr)
    sh_ativos= sum(r["h_ativos"] for r in rr)
    sh_todos = sum(r["h_todos"]  for r in rr)
    return {
        "tot_A":  round(sum(r["tot_A"] for r in rr)/n, 1),
        "tot_B":  round(sum(r["tot_B"] for r in rr)/n, 1),
        "tot_C":  round(sum(r["tot_C"] for r in rr)/n, 1),
        "total":  round(sum(r["total"]  for r in rr)/n, 1),
        "prod_labor_tot": sh_labor/sh_todos  if sh_todos>0  else 0,
        "prod_ciclo_tot": sh_ciclo/sh_todos  if sh_todos>0  else 0,
        "prod_labor_op":  sh_labor/sh_ativos if sh_ativos>0 else 0,
        "prod_ciclo_op":  sh_ciclo/sh_ativos if sh_ativos>0 else 0,
    }

def show_tabela(r):
    dias=r["dias"]; hA,hB,hC=r["hA"],r["hB"],r["hC"]
    heA,heB,heC=r.get("heA",hA),r.get("heB",hB),r.get("heC",hC)
    rows=[]
    for _,row in r["centros"].iterrows():
        rows.append({"Centro":row.centro,
            "Ocup. A":row.ocup_A,"Ocup. B":row.ocup_B,"Ocup. C":row.ocup_C,
            "Ativo A":int(row.ativo_A),"Ativo B":int(row.ativo_B),"Ativo C":int(row.ativo_C),
            "Horas A":round(row.horas_disp_A,2),"Horas B":round(row.horas_disp_B,2),"Horas C":round(row.horas_disp_C,2)})
    df=pd.DataFrame(rows)
    def sr(row):
        s=[""]*len(row)
        for i,col in enumerate(df.columns):
            v=row.iloc[i]
            if col in("Ocup. A","Ocup. B","Ocup. C"):
                if v>1.0: s[i]="background-color:#FFCDD2;color:#B71C1C;font-weight:600"
                elif v>=0.85: s[i]="background-color:#FFFDE7;color:#7B5800;font-weight:600"
                else: s[i]=f"background-color:#E8F5E9;color:{JD_VERDE_ESC};font-weight:600"
            elif col in("Ativo A","Ativo B","Ativo C"):
                s[i]="background-color:#B3E5FC;color:#01579B;font-weight:700" if v else "background-color:#FFFDE7;color:#888"
            elif col in("Horas A","Horas B","Horas C"):
                s[i]="background-color:#B3E5FC;color:#01579B" if v>0 else "background-color:#F5F5F5;color:#AAA"
        return s
    st.dataframe(df.style.apply(sr,axis=1).format({
        "Ocup. A":"{:.0%}","Ocup. B":"{:.0%}","Ocup. C":"{:.0%}",
        "Horas A":"{:.1f}","Horas B":"{:.1f}","Horas C":"{:.1f}"}),
        use_container_width=True,hide_index=True)
    sup=r["suporte"]
    srows=[]
    for nm,k in [("Lavadora e Inspeção","lavadora"),("Gravação e Estanqueidade","gravacao"),
                 ("Preset","preset"),("Coringa","coringa"),("Facilitador","facilitador")]:
        s=sup[k]
        srows.append({"Função":nm,"Qtd A":s["A"],"Qtd B":s["B"],"Qtd C":s["C"],
            "Horas A":round(s["A"]*heA*dias,1),"Horas B":round(s["B"]*heB*dias,1),"Horas C":round(s["C"]*heC*dias,1)})
    srows.append({"Função":"▶ TOTAL POR TURNO",
        "Qtd A":r["tot_A"],"Qtd B":r["tot_B"],"Qtd C":r["tot_C"],
        "Horas A":round(r["tot_A"]*heA*dias,1),"Horas B":round(r["tot_B"]*heB*dias,1),"Horas C":round(r["tot_C"]*heC*dias,1)})
    df_s=pd.DataFrame(srows)
    def ss(row):
        is_t="TOTAL" in str(row["Função"])
        return [f"background-color:{JD_AMARELO};color:{JD_VERDE_ESC};font-weight:700" if is_t else ""]*len(row)
    st.dataframe(df_s.style.apply(ss,axis=1).format({"Horas A":"{:.1f}","Horas B":"{:.1f}","Horas C":"{:.1f}"}),
        use_container_width=True,hide_index=True)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total funcionários",r["total"])
    c2.metric("Ciclo operacional",f"{r['prod_ciclo_op']:.0%}")
    c3.metric("Labor operacional",f"{r['prod_labor_op']:.0%}")
    c4.metric("⭐ Labor total",f"{r['prod_labor_tot']:.0%}")

def grafico_cenarios(cenarios_dict):
    fig=make_subplots(specs=[[{"secondary_y":True}]])
    cores_A=[JD_VERDE,"#66BB6A","#A5D6A7","#1B5E20"]
    cores_B=[JD_AMARELO_ESC,"#FFD54F","#FFE082","#FF6F00"]
    cores_C=["#1565C0","#64B5F6","#BBDEFB","#0D47A1"]
    cores_p=["#C62828","#FF6D00","#7B1FA2","#00695C"]
    for i,(nome,res) in enumerate(cenarios_dict.items()):
        mv,tA,tB,tC,prod=[],[],[],[],[]
        for m,abr in zip(MESES,MESES_ABREV):
            r=res.get(m)
            if not r: continue
            mv.append(abr); tA.append(r["tot_A"]); tB.append(r["tot_B"])
            tC.append(r["tot_C"]); prod.append(r["prod_labor_tot"]*100)
        op=0.9 if i==0 else 0.65
        fig.add_trace(go.Bar(name=f"A—{nome}",x=mv,y=tA,marker_color=cores_A[i%4],opacity=op,
            offsetgroup=i,legendgroup=nome,text=tA,textposition="inside",textfont=dict(color="white",size=9)),secondary_y=False)
        fig.add_trace(go.Bar(name=f"B—{nome}",x=mv,y=tB,marker_color=cores_B[i%4],opacity=op,
            offsetgroup=i,legendgroup=nome,base=tA,text=tB,textposition="inside",textfont=dict(size=9)),secondary_y=False)
        fig.add_trace(go.Bar(name=f"C—{nome}",x=mv,y=tC,marker_color=cores_C[i%4],opacity=op,
            offsetgroup=i,legendgroup=nome,base=[a+b for a,b in zip(tA,tB)],
            text=tC,textposition="inside",textfont=dict(color="white",size=9)),secondary_y=False)
        fig.add_trace(go.Scatter(name=f"Labor—{nome}",x=mv,y=prod,mode="lines+markers+text",
            marker=dict(color=cores_p[i%4],size=10,symbol="circle" if i==0 else "diamond"),
            line=dict(color=cores_p[i%4],width=2,dash="solid" if i==0 else "dot"),
            text=[f"{p:.0f}%" for p in prod],textposition="top center",
            textfont=dict(color=cores_p[i%4],size=11)),secondary_y=True)
    fig.update_layout(
        barmode="stack",
        title=dict(text="MÃO-DE-OBRA POR TURNO",font=dict(size=14,color=JD_VERDE_ESC)),
        legend=dict(orientation="h",y=-0.32,x=0,font=dict(size=10,color="#000000"),
                    bgcolor="rgba(255,255,255,0.95)",bordercolor="#AAAAAA",borderwidth=1),
        height=480,plot_bgcolor="white",paper_bgcolor="white",
        xaxis=dict(showgrid=False,tickfont=dict(size=11,color="#1A1A1A")),
        yaxis=dict(showgrid=True,gridcolor="#E8E8E8",tickfont=dict(size=11,color="#1A1A1A"),
                   title="Nº Funcionários",title_font=dict(size=12,color="#1A1A1A")),
        yaxis2=dict(title="Labor Total (%)",tickformat=".0f",ticksuffix="%",range=[0,100],
                    tickfont=dict(size=11,color="#1A1A1A"),title_font=dict(size=12,color="#1A1A1A")))
    return fig

def safe_int(v):
    try: return int(float(v)) if v is not None else 0
    except: return 0

def safe_float(v):
    try: return float(v) if v is not None else 0.0
    except: return 0.0

_BRD_DIAG = Border(left=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'),
                   top=Side(style='thin',color='CCCCCC'),bottom=Side(style='thin',color='CCCCCC'))
VERDE_HEADER    = PatternFill("solid", fgColor="1F4D19")
VERMELHO_HEADER = PatternFill("solid", fgColor="C62828")
CINZA_HEADER    = PatternFill("solid", fgColor="555555")
VERMELHO_CLARO  = PatternFill("solid", fgColor="FFCDD2")
VERMELHO_ESCURO = PatternFill("solid", fgColor="C62828")
AMARELO         = PatternFill("solid", fgColor="FFDE00")
VERDE           = PatternFill("solid", fgColor="E8F5E9")
F_HDR_VERD  = PatternFill("solid", fgColor="1F4D19")
F_HDR_AMAR  = PatternFill("solid", fgColor="FFDE00")
F_HDR_CINZA = PatternFill("solid", fgColor="555555")
F_CINZA_CLR = PatternFill("solid", fgColor="F4F4F4")
F_BRANCO    = PatternFill("solid", fgColor="FFFFFF")
F_VERM_CLR  = PatternFill("solid", fgColor="FFCDD2")
F_AMAR      = PatternFill("solid", fgColor="FFDE00")
F_VERDE     = PatternFill("solid", fgColor="E8F5E9")
F_VERDE_MED = PatternFill("solid", fgColor="C8E6C9")
F_AZUL_CLR  = PatternFill("solid", fgColor="E3F2FD")
C_VERDE     = PatternFill("solid", fgColor="92D050")
C_AMAR      = PatternFill("solid", fgColor="FFFF00")
C_AZUL      = PatternFill("solid", fgColor="00B0F0")
C_PRETO     = PatternFill("solid", fgColor="000000")
C_CINZA     = PatternFill("solid", fgColor="D9D9D9")
C_BRANCO    = PatternFill("solid", fgColor="FFFFFF")
C_ROSA      = PatternFill("solid", fgColor="FFB6C1")
C_VERM_DIV  = PatternFill("solid", fgColor="FF0000")

def cor_ocup(pct):
    try:
        v = float(pct)
        if v >= 1.06: return PatternFill("solid", fgColor="FF0000")
        if v >= 1.00: return PatternFill("solid", fgColor="FFFF00")
        if v >= 0.40: return PatternFill("solid", fgColor="92D050")
        return PatternFill("solid", fgColor="FFFFFF")
    except: return PatternFill("solid", fgColor="FFFFFF")

def ec(ws, r, c, val, fill=None, bold=False, color="000000", size=9, center=True, italic=False, comment_text=None):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name="Arial", bold=bold, color=color, size=size, italic=italic)
    cell.fill = fill or F_BRANCO
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center", wrap_text=True)
    cell.border = _BRD_DIAG
    return cell

def cell_style(ws, r, c, val, fill=None, bold=False, color="000000", font_size=9, center=True, italic=False, comment_text=None):
    return ec(ws, r, c, val, fill, bold, color, font_size, center, italic, comment_text)

def build_cp_data_anual(resultados, tempo, dist, aplic, pmp):
    meses_c = [m for m in MESES if resultados.get(m)]
    if not meses_c: return None
    try:
        df = (aplic.merge(pmp, on="modelo").merge(tempo, on=["centro","peca"]).merge(dist, on=["centro","peca"]))
        if "vol_int" not in df.columns: df["vol_int"] = 1.0
        df["vol_int"] = pd.to_numeric(df["vol_int"], errors="coerce").fillna(1.0)
        df["indice_ciclo"] = (df.t_ciclo*df.div_carga*df.div_volume*df.vol_int)/df.disponib
        df["min_ciclo"] = df.indice_ciclo * df.qtd
        df["min_labor"] = df.t_labor * df.div_carga * df.qtd
        df_ano = df[df.mes.isin(meses_c)]
        agg = df_ano.groupby(["centro","peca"])[["min_ciclo","min_labor","qtd"]].sum().reset_index()
        attrs = df_ano.drop_duplicates(["centro","peca"])[
            ["centro","peca","t_ciclo","t_labor","div_carga","vol_int","div_volume","disponib","indice_ciclo"]
        ].set_index(["centro","peca"])
        result = []
        for _, row in agg.iterrows():
            cen, peca = row.centro, row.peca
            try:
                at = attrs.loc[(cen, peca)]
                tc=float(at.t_ciclo); tl=float(at.t_labor)
                dc=float(at.div_carga); vi=float(at.vol_int)
                dv=float(at.div_volume); di=float(at.disponib)
                idx=float(at.indice_ciclo)
            except: tc=tl=dc=vi=dv=di=idx=0.0
            result.append((cen, peca, tc, tl, dc, vi, dv, di, idx,
                           float(row.min_ciclo), float(row.min_labor), int(row.qtd)))
        return result
    except: return None

def gerar_aba_anual(wb, resultados, label="ANO", cp_data=None):
    meses_com_dados = [(m, resultados[m]) for m in MESES if resultados.get(m)]
    if not meses_com_dados: return
    brd = Border(left=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'),
                 top=Side(style='thin',color='CCCCCC'),bottom=Side(style='thin',color='CCCCCC'))
    F_VERDE_a  = PatternFill("solid",fgColor="92D050"); F_AMAR_a  = PatternFill("solid",fgColor="FFDE00")
    F_AZUL_a   = PatternFill("solid",fgColor="00B0F0"); F_PRETO_a = PatternFill("solid",fgColor="000000")
    F_CINZA_a  = PatternFill("solid",fgColor="D9D9D9"); F_CINZA2_a= PatternFill("solid",fgColor="BFBFBF")
    F_BRANCO_a = PatternFill("solid",fgColor="FFFFFF"); F_VD_JD_a = PatternFill("solid",fgColor="1F4D19")
    F_AM_JD_a  = PatternFill("solid",fgColor="FFDE00"); F_VERM_a  = PatternFill("solid",fgColor="FF0000")
    F_CINZA_H_a= PatternFill("solid",fgColor="D9D9D9")
    def _e(ws,r,c,val,fill=None,bold=False,color="000000",size=9,center=True,wrap=False):
        cell=ws.cell(row=r,column=c,value=val)
        cell.font=Font(name="Arial",bold=bold,color=color,size=size)
        cell.fill=fill or F_BRANCO_a
        cell.alignment=Alignment(horizontal="center" if center else "left",vertical="center",wrap_text=wrap)
        cell.border=brd; return cell
    def _cor_pct_a(v):
        try:
            f=float(str(v).strip('%'))/100 if '%' in str(v) else float(v)
            if f>=1.06: return F_VERM_a
            if f>=1.00: return PatternFill("solid",fgColor="FFFF00")
            if f>=0.40: return F_VERDE_a
            return F_BRANCO_a
        except: return F_BRANCO_a
    n_meses=len(meses_com_dados); dias_ano=sum(r["dias"] for _,r in meses_com_dados)
    hA=meses_com_dados[0][1]["hA"]; hB=meses_com_dados[0][1]["hB"]; hC=meses_com_dados[0][1]["hC"]
    heA=meses_com_dados[0][1].get("heA",hA); heB=meses_com_dados[0][1].get("heB",hB); heC=meses_com_dados[0][1].get("heC",hC)
    minA_ano=dias_ano*hA*60; minB_ano=dias_ano*hB*60; minC_ano=dias_ano*hC*60
    from collections import defaultdict
    cen_mc=defaultdict(float); cen_ml=defaultdict(float)
    sup_somas={k:{"A":0,"B":0,"C":0} for k in ["lavadora","gravacao","preset","coringa","facilitador"]}
    tot_A_ano=0; tot_B_ano=0; tot_C_ano=0; tot_func_ano=0
    sum_hciclo=0; sum_hlabor=0; sum_hativos=0; sum_htodos=0
    for mes,r in meses_com_dados:
        for _,row in r["centros"].iterrows():
            cen_mc[row.centro]+=row.min_ciclo_total; cen_ml[row.centro]+=row.min_labor_total
        sum_hciclo+=r["h_ciclo"]; sum_hlabor+=r["h_labor"]
        sum_hativos+=r["h_ativos"]; sum_htodos+=r["h_todos"]
        tot_A_ano+=r["tot_A"]; tot_B_ano+=r["tot_B"]; tot_C_ano+=r["tot_C"]; tot_func_ano+=r["total"]
        for k in sup_somas:
            sup_somas[k]["A"]+=r["suporte"][k]["A"]; sup_somas[k]["B"]+=r["suporte"][k]["B"]; sup_somas[k]["C"]+=r["suporte"][k]["C"]
    prod_lt=sum_hlabor/sum_htodos if sum_htodos>0 else 0
    prod_ct=sum_hciclo/sum_htodos if sum_htodos>0 else 0
    prod_lo=sum_hlabor/sum_hativos if sum_hativos>0 else 0
    prod_co=sum_hciclo/sum_hativos if sum_hativos>0 else 0
    centros_ord=list(cen_mc.keys())
    thr_A=meses_com_dados[0][1]["thr_A"]; thr_B=meses_com_dados[0][1]["thr_B"]; thr_C=meses_com_dados[0][1]["thr_C"]
    def ocup_ano(cen,t):
        mc=cen_mc[cen]
        if t=="A": return mc/minA_ano if minA_ano>0 else 0
        if t=="B": return mc/minB_ano if minB_ano>0 else 0
        return mc/minC_ano if minC_ano>0 else 0
    def ativo_ano_A(cen): return 1 if ocup_ano(cen,"A")>thr_A else 0
    def ativo_ano_B(cen): return 1 if ocup_ano(cen,"A")>thr_B else 0
    def ativo_ano_C(cen): return 1 if ocup_ano(cen,"B")>thr_C else 0
    def hdA(cen): return dias_ano*heA*ativo_ano_A(cen)
    def hdB(cen): return dias_ano*heB*ativo_ano_B(cen)
    def hdC(cen): return dias_ano*heC*ativo_ano_C(cen)
    op_A_ano=sum(ativo_ano_A(c) for c in centros_ord)
    op_B_ano=sum(ativo_ano_B(c) for c in centros_ord)
    op_C_ano=sum(ativo_ano_C(c) for c in centros_ord)
    def sup_ano(key,t): return round(sup_somas[key][t]/n_meses) if n_meses else 0
    tot_suporte_A=sum(sup_ano(k,"A") for k in sup_somas)
    tot_suporte_B=sum(sup_ano(k,"B") for k in sup_somas)
    tot_suporte_C=sum(sup_ano(k,"C") for k in sup_somas)
    tot_A_calc=op_A_ano+tot_suporte_A; tot_B_calc=op_B_ano+tot_suporte_B; tot_C_calc=op_C_ano+tot_suporte_C
    tot_func_calc=tot_A_calc+tot_B_calc+tot_C_calc
    ws=wb.create_sheet(label); ws.freeze_panes="F7"
    ws.merge_cells("A1:O1"); _e(ws,1,1,"TOTAIS — ANO",F_VD_JD_a,True,"FFFFFF",9,True)
    for ci,txt,f in [(16,"TURNO A",F_VERDE_a),(17,"TURNO B",F_AMAR_a),(18,"TURNO C",F_AZUL_a)]:
        _e(ws,1,ci,txt,f,True,"000000",8)
    ws.row_dimensions[1].height=14
    ws.merge_cells("A2:O2"); _e(ws,2,1,"TOTAL DE MINUTOS",F_CINZA_H_a,True,"000000",8,False)
    _e(ws,2,16,round(minA_ano,1),F_VERDE_a,True,"000000",8); _e(ws,2,17,round(minB_ano,1),F_AMAR_a,True,"000000",8); _e(ws,2,18,round(minC_ano,1),F_AZUL_a,True,"000000",8)
    ws.row_dimensions[2].height=13
    ws.merge_cells("A3:O3"); _e(ws,3,1,"TOTAL DE HORAS",F_CINZA_H_a,True,"000000",8,False)
    _e(ws,3,16,round(minA_ano/60,2),F_VERDE_a,True,"000000",8); _e(ws,3,17,round(minB_ano/60,2),F_AMAR_a,True,"000000",8); _e(ws,3,18,round(minC_ano/60,2),F_AZUL_a,True,"000000",8)
    ws.row_dimensions[3].height=13
    ws.merge_cells("A4:O4"); _e(ws,4,1,"Nº DIAS TRABALHADOS",F_CINZA_H_a,True,"000000",8,False)
    _e(ws,4,16,dias_ano,F_VERDE_a,True,"FF0000",9); _e(ws,4,17,dias_ano,F_AMAR_a,True,"FF0000",9); _e(ws,4,18,dias_ano,F_AZUL_a,True,"FF0000",9)
    ws.row_dimensions[4].height=13
    ws.merge_cells("A5:O5"); _e(ws,5,1,f"RESUMO DA CARGA — ANO ({dias_ano} dias / {n_meses} meses)",F_VD_JD_a,True,"FFFFFF",10,True)
    ws.row_dimensions[5].height=18
    hdrs_f=[("Máquina",F_CINZA2_a,"000000"),("PEÇA",F_CINZA2_a,"000000"),("DESCRIÇÃO",F_CINZA2_a,"000000"),("PÇ/TRAT",F_CINZA2_a,"000000"),("UM",F_CINZA2_a,"000000"),
            ("Tempo Ciclo (min)",F_PRETO_a,"FFFFFF"),("Tempo Labor (min)",F_PRETO_a,"FFFFFF"),
            ("Div. Carga",PatternFill("solid",fgColor="FF0000"),"FFFF00"),("Vol. Interna",F_CINZA2_a,"000000"),
            ("Div. Volume",PatternFill("solid",fgColor="FF0000"),"FFFF00"),("Disponib.",F_CINZA2_a,"000000"),("Indice Ciclo",F_CINZA2_a,"000000"),
            ("JA.A",F_VERDE_a,"000000"),("JA.B",F_AMAR_a,"000000"),("JA.C",F_AZUL_a,"000000"),
            ("TOTAL CICLOS (MIN)",F_CINZA_a,"000000"),("TOTAL LABOR (MIN)",F_CINZA_a,"000000"),("TOTAL PECAS",F_CINZA_a,"000000")]
    largs=[9,8,16,6,5,9,9,8,8,8,8,9,8,8,8,12,12,8]
    for ci,(h,f,cor) in enumerate(hdrs_f,1):
        _e(ws,6,ci,h,f,True,cor,8,True,True); ws.column_dimensions[get_column_letter(ci)].width=largs[ci-1]
    ws.row_dimensions[6].height=42
    ri=7
    if cp_data:
        for (cen,peca,tc,tl,dc,vi,dv,di,idx_c,mc,ml,_qtd) in cp_data:
            pA=mc/minA_ano if minA_ano>0 else 0; pB=mc/minB_ano if minB_ano>0 else 0; pC=mc/minC_ano if minC_ano>0 else 0
            _e(ws,ri,1,cen,F_BRANCO_a,False,"000000",8,False); _e(ws,ri,2,peca,F_BRANCO_a,False,"000000",8,False)
            _e(ws,ri,3,"ANO",F_BRANCO_a,False,"000000",8,False); _e(ws,ri,4,1,F_BRANCO_a,False,"000000",8); _e(ws,ri,5,"PC",F_BRANCO_a,False,"000000",8)
            _e(ws,ri,6,round(tc,4) if tc else "",F_PRETO_a,False,"FFFFFF",8); _e(ws,ri,7,round(tl,4) if tl else "",F_PRETO_a,False,"FFFFFF",8)
            _e(ws,ri,8,round(dc,4) if dc else "",PatternFill("solid",fgColor="FF0000"),False,"FFFF00",8)
            _e(ws,ri,9,round(vi,4) if vi else "",F_BRANCO_a,False,"000000",8)
            _e(ws,ri,10,round(dv,4) if dv else "",PatternFill("solid",fgColor="FF0000"),False,"FFFF00",8)
            _e(ws,ri,11,round(di,4) if di else "",F_CINZA2_a,False,"000000",8); _e(ws,ri,12,round(idx_c,4),F_BRANCO_a,False,"000000",8)
            _e(ws,ri,13,f"{pA:.1%}",_cor_pct_a(pA),False,"000000",8); _e(ws,ri,14,f"{pB:.1%}",_cor_pct_a(pB),False,"000000",8); _e(ws,ri,15,f"{pC:.1%}",_cor_pct_a(pC),False,"000000",8)
            _e(ws,ri,16,round(mc,1),F_BRANCO_a,False,"000000",8); _e(ws,ri,17,round(ml,1),F_BRANCO_a,False,"000000",8); _e(ws,ri,18,_qtd,F_BRANCO_a,False,"000000",8)
            ws.row_dimensions[ri].height=13; ri+=1
    else:
        for cen in centros_ord:
            mc=cen_mc[cen]; ml=cen_ml[cen]
            pA=mc/minA_ano if minA_ano>0 else 0; pB=mc/minB_ano if minB_ano>0 else 0; pC=mc/minC_ano if minC_ano>0 else 0
            _e(ws,ri,1,cen,F_BRANCO_a,False,"000000",8,False); _e(ws,ri,2,"—",F_BRANCO_a,False,"000000",8)
            for ci_z in range(3,13): _e(ws,ri,ci_z,"",F_BRANCO_a,False,"000000",8)
            _e(ws,ri,13,f"{pA:.1%}",_cor_pct_a(pA),False,"000000",8); _e(ws,ri,14,f"{pB:.1%}",_cor_pct_a(pB),False,"000000",8); _e(ws,ri,15,f"{pC:.1%}",_cor_pct_a(pC),False,"000000",8)
            _e(ws,ri,16,round(mc,1),F_BRANCO_a,False,"000000",8); _e(ws,ri,17,round(ml,1),F_BRANCO_a,False,"000000",8)
            ws.row_dimensions[ri].height=13; ri+=1
    _CF=6; _RS=66
    ws.merge_cells(start_row=_RS-1,start_column=_CF,end_row=_RS-1,end_column=_CF+9)
    _e(ws,_RS-1,_CF,"RESUMO DA CARGA MÁQUINA X QUADRO DE LOTAÇÃO",F_CINZA_H_a,True,"000000",9,True); ws.row_dimensions[_RS-1].height=14
    _e(ws,_RS,_CF,"PERÍODO:",F_CINZA_H_a,True,"000000",8,False); _e(ws,_RS,_CF+1,"ANO",F_BRANCO_a,True,"FF0000",9,False)
    _e(ws,_RS,_CF+3,"DATA DE REVISÃO:",F_CINZA_H_a,True,"000000",8,False)
    _e(ws,_RS,_CF+4,datetime.now().strftime("%d/%m/%Y"),F_BRANCO_a,True,"FF0000",9)
    _e(ws,_RS,_CF+6,"HORAS POR TURNO DE TRABALHO",F_CINZA_H_a,True,"000000",8,True); ws.row_dimensions[_RS].height=14
    _e(ws,_RS+1,_CF+6,heA,F_CINZA_H_a,True,"000000",8); _e(ws,_RS+1,_CF+7,heB,F_CINZA_H_a,True,"000000",8); _e(ws,_RS+1,_CF+8,heC,F_CINZA_H_a,True,"000000",8); ws.row_dimensions[_RS+1].height=14
    for _ci,_txt,_f in [(_CF+1,"TURNO A",F_VERDE_a),(_CF+2,"TURNO B",F_AMAR_a),(_CF+3,"TURNO C",F_AZUL_a),(_CF+4,"TURNO A",F_VERDE_a),(_CF+5,"TURNO B",F_AMAR_a),(_CF+6,"TURNO C",F_AZUL_a),(_CF+7,"TURNO A",F_VERDE_a),(_CF+8,"TURNO B",F_AMAR_a),(_CF+9,"TURNO C",F_AZUL_a)]:
        _e(ws,_RS+2,_ci,_txt,_f,True,"000000",8); ws.row_dimensions[_RS+2].height=14
    _e(ws,_RS+3,_CF,"Centro",F_PRETO_a,True,"FFFFFF",8)
    for _ci,_txt in [(_CF+1,"% Ocup"),(_CF+2,"% Ocup"),(_CF+3,"% Ocup"),(_CF+4,"Ativo"),(_CF+5,"Ativo"),(_CF+6,"Ativo"),(_CF+7,"Horas"),(_CF+8,"Horas"),(_CF+9,"Horas")]:
        _e(ws,_RS+3,_ci,_txt,F_PRETO_a,True,"FFFFFF",8); ws.row_dimensions[_RS+3].height=14
    _ri=_RS+4
    for cen in centros_ord:
        oA=ocup_ano(cen,"A"); oB=ocup_ano(cen,"B"); oC=ocup_ano(cen,"C")
        aA=ativo_ano_A(cen); aB=ativo_ano_B(cen); aC=ativo_ano_C(cen)
        def _cbg_a(v):
            if v>=1.06: return F_VERM_a
            if v>=1.00: return PatternFill("solid",fgColor="FFFF00")
            if v>=0.40: return F_VERDE_a
            return F_BRANCO_a
        _e(ws,_ri,_CF,cen,F_BRANCO_a,False,"000000",8,False)
        _e(ws,_ri,_CF+1,f"{oA:.1%}",_cbg_a(oA),False,"000000",8); _e(ws,_ri,_CF+2,f"{oB:.1%}",_cbg_a(oB),False,"000000",8); _e(ws,_ri,_CF+3,f"{oC:.1%}",_cbg_a(oC),False,"000000",8)
        _e(ws,_ri,_CF+4,aA,F_VERDE_a if aA else F_AMAR_a,True,"000000",8); _e(ws,_ri,_CF+5,aB,F_VERDE_a if aB else F_AMAR_a,True,"000000",8); _e(ws,_ri,_CF+6,aC,F_AZUL_a if aC else F_CINZA_a,True,"000000",8)
        _e(ws,_ri,_CF+7,round(dias_ano*heA*aA,2) if aA else 0,F_VERDE_a if aA else F_BRANCO_a,True,"000000",8)
        _e(ws,_ri,_CF+8,round(dias_ano*heB*aB,2) if aB else 0,F_AMAR_a if aB else F_BRANCO_a,True,"000000",8)
        _e(ws,_ri,_CF+9,round(dias_ano*heC*aC,2) if aC else 0,F_AZUL_a if aC else F_BRANCO_a,True,"000000",8)
        ws.row_dimensions[_ri].height=13; _ri+=1
    for _snm,_sk in [("TOTAL DE OPERADORES",None),("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador"),("TOTAL POR TURNO",None),("TOTAL FUNCIONÁRIOS",None)]:
        _bold="TOTAL" in _snm; _bg=F_AM_JD_a if _bold else F_BRANCO_a; _fg="1F4D19" if _bold else "000000"
        _e(ws,_ri,_CF,_snm,_bg,_bold,_fg,8,False)
        if _sk:
            sA=sup_ano(_sk,"A"); sB=sup_ano(_sk,"B"); sC=sup_ano(_sk,"C")
            for _ci,_v in [(_CF+4,sA),(_CF+5,sB),(_CF+6,sC)]: _e(ws,_ri,_ci,_v,F_VERDE_a if _ci==_CF+4 else (F_AMAR_a if _ci==_CF+5 else F_AZUL_a),True,"000000",8)
            for _ci,_v,_he in [(_CF+7,sA,heA),(_CF+8,sB,heB),(_CF+9,sC,heC)]: _e(ws,_ri,_ci,round(_v*_he*dias_ano,2) if _v else 0,F_VERDE_a if _ci==_CF+7 else (F_AMAR_a if _ci==_CF+8 else F_AZUL_a),True,"000000",8)
        elif "TOTAL DE OPERADORES" in _snm:
            for _ci,_v in [(_CF+4,op_A_ano),(_CF+5,op_B_ano),(_CF+6,op_C_ano)]: _e(ws,_ri,_ci,_v,F_AM_JD_a,True,"1F4D19",8)
            for _ci,_v,_he in [(_CF+7,op_A_ano,heA),(_CF+8,op_B_ano,heB),(_CF+9,op_C_ano,heC)]: _e(ws,_ri,_ci,round(_v*_he*dias_ano,2),F_AM_JD_a,True,"1F4D19",8)
        elif "TOTAL POR TURNO" in _snm:
            for _ci,_v in [(_CF+4,tot_A_calc),(_CF+5,tot_B_calc),(_CF+6,tot_C_calc)]: _e(ws,_ri,_ci,_v,F_AM_JD_a,True,"1F4D19",8)
            for _ci,_v,_he in [(_CF+7,tot_A_calc,heA),(_CF+8,tot_B_calc,heB),(_CF+9,tot_C_calc,heC)]: _e(ws,_ri,_ci,round(_v*_he*dias_ano,2),F_AM_JD_a,True,"1F4D19",8)
        elif "FUNCIONÁRIOS" in _snm:
            _e(ws,_ri,_CF+4,tot_func_calc,F_AM_JD_a,True,"1F4D19",9)
            _e(ws,_ri,_CF+7,round(tot_A_calc*heA*dias_ano+tot_B_calc*heB*dias_ano+tot_C_calc*heC*dias_ano,2),F_AM_JD_a,True,"1F4D19",9)
        ws.row_dimensions[_ri].height=13; _ri+=1
    _ri+=1
    for _pnm,_pv,_dest in [("PRODUTIVIDADE POR TEMPO DE CICLO OPERACIONAL",prod_co,False),("PRODUTIVIDADE POR TEMPO DE CICLO TOTAL",prod_ct,False),("PRODUTIVIDADE POR TEMPO DE LABOR OPERACIONAL",prod_lo,False),("PRODUTIVIDADE POR TEMPO DE LABOR TOTAL ★",prod_lt,True)]:
        ws.merge_cells(start_row=_ri,start_column=_CF,end_row=_ri,end_column=_CF+8)
        _e(ws,_ri,_CF,_pnm,F_AM_JD_a if _dest else F_BRANCO_a,_dest,"1F4D19" if _dest else "000000",8,False)
        _e(ws,_ri,_CF+9,f"{_pv:.1%}",F_AM_JD_a if _dest else F_BRANCO_a,_dest,"1F4D19" if _dest else "000000",8)
        ws.row_dimensions[_ri].height=14; _ri+=1
    for ci,w in enumerate([9,8,16,6,5,9,9,8,8,8,8,9,8,8,8,12,12,8],1): ws.column_dimensions[get_column_letter(ci)].width=w
    for _ci,_ww in [(_CF,14),(_CF+1,8),(_CF+2,8),(_CF+3,8),(_CF+4,8),(_CF+5,8),(_CF+6,8),(_CF+7,10),(_CF+8,10),(_CF+9,10)]:
        ws.column_dimensions[get_column_letter(_ci)].width=_ww

def exportar(resultados, _tempo=None, _dist=None, _aplic=None, _pmp=None):
    out=BytesIO(); wb=openpyxl.Workbook()
    brd=Border(left=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'),top=Side(style='thin',color='CCCCCC'),bottom=Side(style='thin',color='CCCCCC'))
    def ec_l(c,bg="FFFFFF",fg="000000",bold=False,fmt=None,center=True):
        c.font=Font(name="Arial",bold=bold,color=fg,size=9); c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center" if center else "left",vertical="center"); c.border=brd
        if fmt:
            try: c.number_format=fmt
            except: pass
    ws=wb.active; ws.title="RESUMO MO"
    JD_V=JD_VERDE_ESC.replace("#",""); JD_Y=JD_AMARELO.replace("#","")
    for i,h in enumerate(["Mês","Dias","Turno A","Turno B","Turno C","Total","Ciclo Op.","Ciclo Total","Labor Op.","Labor Total ★"],1):
        ec_l(ws.cell(1,i,h),JD_V,"FFFFFF",True)
    for ri,(m,abr) in enumerate(zip(MESES,MESES_ABREV),2):
        r=resultados.get(m); bg="EAF3FB" if ri%2==0 else "FFFFFF"
        vals=[abr,0,"-","-","-","-","-","-","-","-"] if not r else [abr,r["dias"],r["tot_A"],r["tot_B"],r["tot_C"],r["total"],r["prod_ciclo_op"],r["prod_ciclo_tot"],r["prod_labor_op"],r["prod_labor_tot"]]
        for ci,v in enumerate(vals,1):
            v_cell=f"{v:.1%}" if ci>=7 and isinstance(v,float) else v
            c_obj=ws.cell(ri,ci,v_cell)
            ec_l(c_obj,JD_Y if ci==10 and isinstance(v,float) else bg,JD_V if ci==10 and isinstance(v,float) else "000000",ci==10 and isinstance(v,float))
        ws.row_dimensions[ri].height=15
    for mes in MESES:
        r=resultados.get(mes)
        if not r: continue
        wsm=wb.create_sheet(mes[:10]); hA,hB,hC,dias=r["hA"],r["hB"],r["hC"],r["dias"]
        heA,heB,heC=r.get("heA",hA),r.get("heB",hB),r.get("heC",hC)
        for ci,txt in [(1,""),(2,"TURNO A"),(3,"TURNO B"),(4,"TURNO C"),(5,"TURNO A"),(6,"TURNO B"),(7,"TURNO C"),(8,"TURNO A"),(9,"TURNO B"),(10,"TURNO C")]:
            ec_l(wsm.cell(1,ci,txt),JD_V,"FFFFFF",True)
        wsm.cell(1,1,mes.upper()); ec_l(wsm.cell(1,1),JD_V,"FFFFFF",True)
        JD_V2=JD_VERDE_ESC.replace("#",""); JD_Y2=JD_AMARELO.replace("#","")
        wsm.merge_cells("B2:D2")
        c2=wsm.cell(2,1,"CENTRO"); c2.font=Font(name="Arial",bold=True,color="FFFFFF",size=9); c2.fill=PatternFill("solid",fgColor=JD_V2); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        c2=wsm.cell(2,2,"% OCUPAÇÃO"); c2.font=Font(name="Arial",bold=True,color="FFFFFF",size=9); c2.fill=PatternFill("solid",fgColor=JD_V2); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        wsm.merge_cells("E2:G2")
        c2=wsm.cell(2,5,"TURNO ATIVO (0=inativo 1=ativo)"); c2.font=Font(name="Arial",bold=True,color=JD_V2,size=9); c2.fill=PatternFill("solid",fgColor=JD_Y2); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        wsm.merge_cells("H2:J2")
        c2=wsm.cell(2,8,"HORAS DISPONIVEIS NO MES"); c2.font=Font(name="Arial",bold=True,color="FFFFFF",size=9); c2.fill=PatternFill("solid",fgColor="1565C0"); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        wsm.row_dimensions[2].height=16
        def cbg(v):
            if v>1.0: return "FFCDD2"
            if v>=0.85: return "FFFDE7"
            return "E8F5E9"
        ri2=3
        for _,row in r["centros"].iterrows():
            for ci,(val,bg,ctr) in enumerate([(row.centro,"FFFFFF",False),(f"{row.ocup_A:.1%}",cbg(row.ocup_A),True),(f"{row.ocup_B:.1%}",cbg(row.ocup_B),True),(f"{row.ocup_C:.1%}",cbg(row.ocup_C),True),(row.ativo_A,"B3E5FC" if row.ativo_A else "FFFDE7",True),(row.ativo_B,"B3E5FC" if row.ativo_B else "FFFDE7",True),(row.ativo_C,"B3E5FC" if row.ativo_C else "FFFDE7",True),(f"{row.horas_disp_A:.2f}" if row.ativo_A else "0","B3E5FC" if row.ativo_A else "F5F5F5",True),(f"{row.horas_disp_B:.2f}" if row.ativo_B else "0","B3E5FC" if row.ativo_B else "F5F5F5",True),(f"{row.horas_disp_C:.2f}" if row.ativo_C else "0","B3E5FC" if row.ativo_C else "F5F5F5",True)],1):
                ec_l(wsm.cell(ri2,ci,val),bg,center=ctr)
            ri2+=1
        sup=r["suporte"]
        for nome,key in [("TOTAL DE OPERADORES",None),("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador"),("TOTAL POR TURNO",None),("TOTAL FUNCIONÁRIOS",None)]:
            bold="TOTAL" in nome; bg_r=JD_Y if bold else "FFFFFF"; fg_r=JD_V if bold else "000000"
            ec_l(wsm.cell(ri2,1,nome),bg_r,fg_r,bold,center=False)
            if key:
                s=sup[key]
                for ci,t in [(5,"A"),(6,"B"),(7,"C")]: ec_l(wsm.cell(ri2,ci,s[t]),"B3E5FC" if s[t] else "FFFDE7",bold=bold)
                for ci,t,h in [(8,"A",heA),(9,"B",heB),(10,"C",heC)]:
                    v=s[t]*h*dias; ec_l(wsm.cell(ri2,ci,f"{v:.2f}" if v else "0"),"B3E5FC" if v else "F5F5F5",bold=bold)
            elif "TOTAL DE OPERADORES" in nome:
                for ci,v in [(5,r["op_A"]),(6,r["op_B"]),(7,r["op_C"])]: ec_l(wsm.cell(ri2,ci,v),JD_Y,JD_V,True)
                for ci,v,h in [(8,r["op_A"],heA),(9,r["op_B"],heB),(10,r["op_C"],heC)]: ec_l(wsm.cell(ri2,ci,f"{v*h*dias:.2f}"),JD_Y,JD_V,True)
            elif "TOTAL POR TURNO" in nome:
                for ci,v in [(5,r["tot_A"]),(6,r["tot_B"]),(7,r["tot_C"])]: ec_l(wsm.cell(ri2,ci,v),JD_Y,JD_V,True)
                for ci,v,h in [(8,r["tot_A"],heA),(9,r["tot_B"],heB),(10,r["tot_C"],heC)]: ec_l(wsm.cell(ri2,ci,f"{v*h*dias:.2f}"),JD_Y,JD_V,True)
            elif "FUNCIONÁRIOS" in nome:
                ec_l(wsm.cell(ri2,4,r["total"]),JD_Y,JD_V,True)
                ec_l(wsm.cell(ri2,8,f"{r['tot_A']*heA*dias+r['tot_B']*heB*dias+r['tot_C']*heC*dias:.2f}"),JD_Y,JD_V,True)
            ri2+=1
        ri2+=1
        for nm,v,dest in [("PROD. CICLO OPERACIONAL",r["prod_ciclo_op"],False),("PROD. CICLO TOTAL",r["prod_ciclo_tot"],False),("PROD. LABOR OPERACIONAL",r["prod_labor_op"],False),("PROD. LABOR TOTAL ★",r["prod_labor_tot"],True)]:
            wsm.merge_cells(f"H{ri2}:I{ri2}")
            ec_l(wsm.cell(ri2,8,nm),JD_Y if dest else "FFFFFF",JD_V if dest else "000000",dest,center=False)
            ec_l(wsm.cell(ri2,10,f"{v:.1%}" if isinstance(v,float) else v),JD_Y if dest else "FFFFFF",JD_V if dest else "000000",dest)
            ri2+=1
        for ci,w in enumerate([14,8,8,8,8,8,8,24,10,10],1): wsm.column_dimensions[get_column_letter(ci)].width=w
    if _tempo is not None and _dist is not None and _aplic is not None and _pmp is not None:
        _cp_ano=build_cp_data_anual(resultados,_tempo,_dist,_aplic,_pmp)
    else:
        _cp_ano=None
    gerar_aba_anual(wb,resultados,cp_data=_cp_ano)
    wb.save(out); out.seek(0); return out

def exportar_cenario_vs_base(res_base, res_cenario, mes, nome_cenario):
    out=BytesIO(); wb=openpyxl.Workbook()
    brd=Border(left=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'),top=Side(style='thin',color='CCCCCC'),bottom=Side(style='thin',color='CCCCCC'))
    JD_V=JD_VERDE_ESC.replace("#",""); JD_Y=JD_AMARELO.replace("#","")
    def ec_c(c,bg="FFFFFF",fg="000000",bold=False,center=True):
        c.font=Font(name="Arial",bold=bold,color=fg,size=9); c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center" if center else "left",vertical="center"); c.border=brd
    def escrever_mes(ws,r,titulo):
        hA,hB,hC,dias=r["hA"],r["hB"],r["hC"],r["dias"]
        heA,heB,heC=r.get("heA",hA),r.get("heB",hB),r.get("heC",hC)
        for ci,txt in [(1,titulo),(2,"TURNO A"),(3,"TURNO B"),(4,"TURNO C"),(5,"TURNO A"),(6,"TURNO B"),(7,"TURNO C"),(8,"TURNO A"),(9,"TURNO B"),(10,"TURNO C")]:
            ec_c(ws.cell(1,ci,txt),JD_V,"FFFFFF",True)
        ws.row_dimensions[1].height=16
        ec_c(ws.cell(2,1,"CENTRO"),JD_V,"FFFFFF",True); ec_c(ws.cell(2,2,"% OCUPAÇÃO"),JD_V,"FFFFFF",True)
        ws.merge_cells("E2:G2"); ec_c(ws.cell(2,5,"TURNO ATIVO (0=inativo 1=ativo)"),JD_Y,JD_V,True)
        ws.merge_cells("H2:J2"); ec_c(ws.cell(2,8,"HORAS DISPONÍVEIS NO MÊS"),"1565C0","FFFFFF",True)
        ws.row_dimensions[2].height=16
        def cbg(v):
            if v>1.0: return "FFCDD2"
            if v>=0.85: return "FFFDE7"
            return "E8F5E9"
        ri=3
        for _,row in r["centros"].iterrows():
            for ci,(val,bg,ctr) in enumerate([(row.centro,"FFFFFF",False),(f"{row.ocup_A:.1%}",cbg(row.ocup_A),True),(f"{row.ocup_B:.1%}",cbg(row.ocup_B),True),(f"{row.ocup_C:.1%}",cbg(row.ocup_C),True),(row.ativo_A,"B3E5FC" if row.ativo_A else "FFFDE7",True),(row.ativo_B,"B3E5FC" if row.ativo_B else "FFFDE7",True),(row.ativo_C,"B3E5FC" if row.ativo_C else "FFFDE7",True),(f"{row.horas_disp_A:.2f}" if row.ativo_A else "0","B3E5FC" if row.ativo_A else "F5F5F5",True),(f"{row.horas_disp_B:.2f}" if row.ativo_B else "0","B3E5FC" if row.ativo_B else "F5F5F5",True),(f"{row.horas_disp_C:.2f}" if row.ativo_C else "0","B3E5FC" if row.ativo_C else "F5F5F5",True)],1):
                ec_c(ws.cell(ri,ci,val),bg,center=ctr)
            ri+=1
        sup=r["suporte"]
        for nome,key in [("TOTAL DE OPERADORES",None),("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador"),("TOTAL POR TURNO",None),("TOTAL FUNCIONÁRIOS",None)]:
            bold="TOTAL" in nome; bg_r=JD_Y if bold else "FFFFFF"; fg_r=JD_V if bold else "000000"
            ec_c(ws.cell(ri,1,nome),bg_r,fg_r,bold,False)
            if key:
                s=sup[key]
                for ci,t in [(5,"A"),(6,"B"),(7,"C")]: ec_c(ws.cell(ri,ci,s[t]),"B3E5FC" if s[t] else "FFFDE7",bold=bold)
                for ci,t,h in [(8,"A",heA),(9,"B",heB),(10,"C",heC)]:
                    v2=s[t]*h*dias; ec_c(ws.cell(ri,ci,f"{v2:.2f}" if v2 else "0"),"B3E5FC" if v2 else "F5F5F5",bold=bold)
            elif "TOTAL DE OPERADORES" in nome:
                for ci,v2 in [(5,r["op_A"]),(6,r["op_B"]),(7,r["op_C"])]: ec_c(ws.cell(ri,ci,v2),JD_Y,JD_V,True)
                for ci,v2,h in [(8,r["op_A"],heA),(9,r["op_B"],heB),(10,r["op_C"],heC)]: ec_c(ws.cell(ri,ci,f"{v2*h*dias:.2f}"),JD_Y,JD_V,True)
            elif "TOTAL POR TURNO" in nome:
                for ci,v2 in [(5,r["tot_A"]),(6,r["tot_B"]),(7,r["tot_C"])]: ec_c(ws.cell(ri,ci,v2),JD_Y,JD_V,True)
                for ci,v2,h in [(8,r["tot_A"],heA),(9,r["tot_B"],heB),(10,r["tot_C"],heC)]: ec_c(ws.cell(ri,ci,f"{v2*h*dias:.2f}"),JD_Y,JD_V,True)
            elif "FUNCIONÁRIOS" in nome:
                ec_c(ws.cell(ri,4,r["total"]),JD_Y,JD_V,True)
                ec_c(ws.cell(ri,8,f"{r['tot_A']*heA*dias+r['tot_B']*heB*dias+r['tot_C']*heC*dias:.2f}"),JD_Y,JD_V,True)
            ri+=1
        ri+=1
        for nm2,v2,dest in [("PROD. CICLO OPERACIONAL",r["prod_ciclo_op"],False),("PROD. CICLO TOTAL",r["prod_ciclo_tot"],False),("PROD. LABOR OPERACIONAL",r["prod_labor_op"],False),("PROD. LABOR TOTAL ★",r["prod_labor_tot"],True)]:
            ws.merge_cells(f"H{ri}:I{ri}")
            ec_c(ws.cell(ri,8,nm2),JD_Y if dest else "FFFFFF",JD_V if dest else "000000",dest,False)
            ec_c(ws.cell(ri,10,f"{v2:.1%}"),JD_Y if dest else "FFFFFF",JD_V if dest else "000000",dest)
            ri+=1
        for ci,w in enumerate([14,8,8,8,8,8,8,24,10,10],1): ws.column_dimensions[get_column_letter(ci)].width=w
    r_b=res_base.get(mes); r_c=res_cenario.get(mes)
    ws1=wb.active; ws1.title=f"{mes[:8]} Base"[:31]
    if r_b: escrever_mes(ws1,r_b,f"{mes.upper()} — BASE")
    else: ws1.cell(1,1,"Sem dados para este mês no cenário base")
    ws2=wb.create_sheet(f"{mes[:6]} {nome_cenario[:8]}"[:31])
    if r_c: escrever_mes(ws2,r_c,f"{mes.upper()} — {nome_cenario.upper()}")
    else: ws2.cell(1,1,"Sem dados para este mês no cenário")
    ws3=wb.create_sheet("Comparação")
    ws3.merge_cells("A1:N1")
    ct=ws3.cell(1,1,f"COMPARAÇÃO — {mes.upper()} | Base vs {nome_cenario}")
    ct.font=Font(name="Arial",bold=True,color="FFFFFF",size=11); ct.fill=PatternFill("solid",fgColor=JD_V); ct.alignment=Alignment(horizontal="center",vertical="center"); ct.border=brd
    ws3.row_dimensions[1].height=24
    for i,h in enumerate(["","Base","Cenário","Δ"],1): ec_c(ws3.cell(2,i,h),JD_V,"FFFFFF",True)
    ws3.row_dimensions[2].height=15
    metricas=[]
    if r_b and r_c:
        metricas=[("Turno A (total)",r_b["tot_A"],r_c["tot_A"]),("Turno B (total)",r_b["tot_B"],r_c["tot_B"]),("Turno C (total)",r_b["tot_C"],r_c["tot_C"]),("TOTAL FUNCIONÁRIOS",r_b["total"],r_c["total"]),("Operadores CEN A",r_b["op_A"],r_c["op_A"]),("Operadores CEN B",r_b["op_B"],r_c["op_B"]),("Operadores CEN C",r_b["op_C"],r_c["op_C"])]
    for ri3,(nome3,vb3,vc3) in enumerate(metricas,3):
        is_total="TOTAL" in nome3; delta3=vc3-vb3
        ec_c(ws3.cell(ri3,1,nome3),JD_Y if is_total else "F5F5F5",JD_V if is_total else "000000",is_total,False)
        ec_c(ws3.cell(ri3,2,vb3),"EAF3FB","000000",is_total); ec_c(ws3.cell(ri3,3,vc3),"EAF3FB","000000",is_total)
        cor_d="003D10" if delta3<0 else ("3D0000" if delta3>0 else "555555")
        bg_d="B9F6CA" if delta3<0 else ("FFCDD2" if delta3>0 else "F5F5F5")
        ec_c(ws3.cell(ri3,4,f"{delta3:+d}"),bg_d,cor_d,is_total); ws3.row_dimensions[ri3].height=14
    for ci,w in enumerate([16,9,9,9],1): ws3.column_dimensions[get_column_letter(ci)].width=w
    wb.save(out); out.seek(0); return out

def comparar_com_excel(res_app, file_bytes, tempo, dist, aplic, pmp, dias, horas_turno, thresholds, suporte_cfg):
    MAPA={"Novembro":"NovFY26","Dezembro":"DezFY26","Janeiro":"JanFY26","Fevereiro":"FevFY26","Março":"MarFY26","Abril":"AbrFY26","Maio":"MaiFY26","Junho":"JunFY26","Julho":"JulFY26","Agosto":"AgoFY26","Setembro":"SetFY26","Outubro":"OutFY26"}
    try:
        wb=openpyxl.load_workbook(BytesIO(file_bytes),read_only=True,data_only=True); abas=wb.sheetnames
    except Exception as e:
        return None,None,f"Erro ao abrir: {e}"
    thr_A=thresholds["A"]/100; thr_B=thresholds["B"]/100; thr_C=thresholds["C"]/100
    hA=horas_turno["A"]; hB=horas_turno["B"]
    try:
        df_all=(aplic.merge(pmp,on="modelo").merge(tempo,on=["centro","peca"]).merge(dist,on=["centro","peca"]))
        if "vol_int" not in df_all.columns: df_all["vol_int"]=1.0
        df_all["vol_int"]=pd.to_numeric(df_all["vol_int"],errors="coerce").fillna(1.0)
        df_all["indice_ciclo"]=(df_all.t_ciclo*df_all.div_carga*df_all.div_volume*df_all.vol_int)/df_all.disponib
        df_all["min_ciclo"]=df_all.indice_ciclo*df_all.qtd
        agg_all=df_all.groupby(["centro","mes"]).agg(min_ciclo=("min_ciclo","sum"),qtd_total=("qtd","sum"),indice_medio=("indice_ciclo","mean")).reset_index()
    except Exception as e:
        wb.close(); return None,None,f"Erro no cálculo: {e}"
    resumo_rows=[]; detalhe_rows=[]
    for mes,aba in MAPA.items():
        r_app=res_app.get(mes)
        if not r_app: continue
        if aba not in abas:
            resumo_rows.append({"Mês":mes,"Status":"⚠️ aba ausente","Observação":f"Aba {aba} não encontrada"})
            continue
        ws_r=wb[aba]; d=dias.get(mes,0)
        if d==0: continue
        minA=d*hA*60; minB=d*hB*60
        xl_opA=safe_int(ws_r.cell(89,27).value); xl_opB=safe_int(ws_r.cell(89,28).value)
        xl_opC=safe_int(ws_r.cell(89,29).value); xl_tot=safe_int(ws_r.cell(96,27).value)
        xl_labor=safe_float(ws_r.cell(101,30).value)
        dA=r_app["op_A"]-xl_opA; dB=r_app["op_B"]-xl_opB; dC=r_app["op_C"]-xl_opC; dT=r_app["total"]-xl_tot
        if dT==0 and dA==0 and dB==0 and dC==0: status="✅ Igual"
        elif abs(dT)<=2: status="🟡 Pequena diferença"
        else: status="🔴 Divergência"
        resumo_rows.append({"Mês":mes,"Status":status,"CEN A App":r_app["op_A"],"CEN A Excel":xl_opA,"Δ A":f"{dA:+d}","CEN B App":r_app["op_B"],"CEN B Excel":xl_opB,"Δ B":f"{dB:+d}","CEN C App":r_app["op_C"],"CEN C Excel":xl_opC,"Δ C":f"{dC:+d}","Total App":r_app["total"],"Total Excel":xl_tot,"Δ Total":f"{dT:+d}","Labor App":f"{r_app['prod_labor_tot']:.1%}","Labor Excel":f"{xl_labor:.1%}" if xl_labor else "—"})
        if status=="✅ Igual": continue
        agg_mes=agg_all[agg_all.mes==mes].copy()
        centros_xl={}
        for r_row in range(69,89):
            cen_val=ws_r.cell(r_row,23).value
            if not cen_val: continue
            centros_xl[str(cen_val).strip()]={"ocup_A":safe_float(ws_r.cell(r_row,24).value),"ocup_B":safe_float(ws_r.cell(r_row,25).value),"ativo_A":safe_int(ws_r.cell(r_row,27).value),"ativo_B":safe_int(ws_r.cell(r_row,28).value),"ativo_C":safe_int(ws_r.cell(r_row,29).value)}
        for _,row in agg_mes.iterrows():
            try:
                cen=row.centro; mc=row.min_ciclo; qtd_app=row.qtd_total; idx_medio=row.indice_medio
                oA_app=mc/minA if minA>0 else 0; oB_app=mc/minB if minB>0 else 0
                aA_app=1 if oA_app>thr_A else 0; aB_app=1 if oA_app>thr_B else 0; aC_app=1 if oB_app>thr_C else 0
                xl=centros_xl.get(cen,{}); aA_xl=xl.get("ativo_A",0); aB_xl=xl.get("ativo_B",0); aC_xl=xl.get("ativo_C",0)
                oA_xl=xl.get("ocup_A",0.0); oB_xl=xl.get("ocup_B",0.0)
                for turno,a_app,a_xl,ocup_app,ocup_xl in [("A",aA_app,aA_xl,oA_app,oA_xl),("B",aB_app,aB_xl,oA_app,oA_xl),("C",aC_app,aC_xl,oB_app,oB_xl)]:
                    if a_app==a_xl: continue
                    delta_ocup=ocup_app-float(ocup_xl); abs_delta=abs(delta_ocup)
                    mc_xl_esp=float(ocup_xl)*(minA if turno in ("A","B") else minB)
                    vol_xl_estim=mc_xl_esp/idx_medio if idx_medio>0 else 0
                    idx_esperado=mc_xl_esp/qtd_app if qtd_app>0 else 0
                    if abs_delta>0.15:
                        if qtd_app<vol_xl_estim*0.7: causa="Volume menor que esperado"; origem=f"IMPUTAPLICAÇÃO — verifique modelos do {cen}"; expl=f"App: {qtd_app:.0f} peças vs Excel: ~{vol_xl_estim:.0f}"
                        elif qtd_app>vol_xl_estim*1.3: causa="Volume maior que esperado"; origem=f"IMPUTAPLICAÇÃO — verifique modelos extras do {cen}"; expl=f"App: {qtd_app:.0f} peças vs Excel: ~{vol_xl_estim:.0f}"
                        else: causa="Índice de ciclo diferente"; origem=f"IMPUTDISTRIBUIÇÃO — div_carga/div_volume/disponib do {cen}"; expl=f"Índice app={idx_medio:.2f} vs esperado={idx_esperado:.2f}"
                    else:
                        thr_u=thr_A if turno=="A" else (thr_B if turno=="B" else thr_C)
                        causa=f"Ocupação próxima do threshold ({thr_u:.0%})"; origem=f"INPUT_PMP — volumes do {cen}"; expl=f"Ocup app={ocup_app:.1%} vs Excel={ocup_xl:.1%}"
                    detalhe_rows.append({"Mês":mes,"Centro":cen,"Turno":turno,"App Ativo":"✅ Sim" if a_app else "❌ Não","Excel Ativo":"✅ Sim" if a_xl else "❌ Não","Ocup. App":f"{ocup_app:.1%}","Ocup. Excel":f"{float(ocup_xl):.1%}","Δ Ocupação":f"{delta_ocup:+.1%}","Causa":causa,"Onde investigar":origem,"Explicação":expl})
            except: continue
    wb.close()
    return pd.DataFrame(resumo_rows),pd.DataFrame(detalhe_rows) if detalhe_rows else pd.DataFrame(),None

def show_memoria(r, mes, df_intermediario, agg, horas_turno, thresholds):
    st.markdown(f'<div class="jd-section">Memória de cálculo — {mes}</div>', unsafe_allow_html=True)
    sup=r["suporte"]; d=r["dias"]; hA,hB,hC=r["hA"],r["hB"],r["hC"]
    heA,heB,heC=r.get("heA",hA),r.get("heB",hB),r.get("heC",hC)
    st.markdown('<div class="mem-step"><span class="step-num">1</span> <b>Inputs utilizados</b></div>', unsafe_allow_html=True)
    c1,c2,c3=st.columns(3)
    c1.metric("Turno A",f"{r['minA']:.0f} min",f"{d}×{hA}×60"); c2.metric("Turno B",f"{r['minB']:.0f} min",f"{d}×{hB}×60"); c3.metric("Turno C",f"{r['minC']:.0f} min")
    df_inp=df_intermediario[df_intermediario.mes==mes][["centro","peca","modelo","t_ciclo","t_labor","div_carga","div_volume","vol_int","disponib","qtd"]].head(8).copy()
    df_inp.columns=["Centro","Peça","Modelo","T.Ciclo","T.Labor","Div.Carga","Div.Volume","Vol.Int","Disponib","Qtd"]
    st.dataframe(df_inp.reset_index(drop=True),use_container_width=True,hide_index=True)
    st.markdown('<div class="mem-step"><span class="step-num">2</span> <b>Índice de ciclo</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">indice_ciclo = (t_ciclo × div_carga × div_volume × vol_interna) ÷ disponibilidade</div>', unsafe_allow_html=True)
    st.markdown('<div class="mem-step"><span class="step-num">3</span> <b>Minutos por linha</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">min_ciclo = indice_ciclo × qtd_pecas<br>min_labor = t_labor × div_carga × qtd_pecas</div>', unsafe_allow_html=True)
    st.markdown('<div class="mem-step"><span class="step-num">4</span> <b>Agrupamento e ocupação por centro</b></div>', unsafe_allow_html=True)
    df_p4=r["centros"][["centro","min_ciclo_total","ocup_A","ocup_B","ocup_C"]].copy()
    df_p4.columns=["Centro","Σ Min.Ciclo","Ocup. A","Ocup. B","Ocup. C"]
    st.dataframe(df_p4.reset_index(drop=True).style.format({"Ocup. A":"{:.1%}","Ocup. B":"{:.1%}","Ocup. C":"{:.1%}","Σ Min.Ciclo":"{:.1f}"}),use_container_width=True,hide_index=True)
    st.markdown('<div class="mem-step"><span class="step-num">5</span> <b>Ativação de turno</b></div>', unsafe_allow_html=True)
    st.markdown(f"- Turno A abre se ocup_A > **{thresholds['A']}%**\n- Turno B abre se ocup_A > **{thresholds['B']}%**\n- Turno C abre se ocup_B > **{thresholds['C']}%**")
    st.markdown('<div class="mem-step"><span class="step-num">6</span> <b>Total por turno</b></div>', unsafe_allow_html=True)
    tot_data={"Função":["Operadores CEN","Lavadora","Gravação","Preset","Coringa","Facilitador","TOTAL ✓"],"Turno A":[r["op_A"],sup["lavadora"]["A"],sup["gravacao"]["A"],sup["preset"]["A"],sup["coringa"]["A"],sup["facilitador"]["A"],r["tot_A"]],"Turno B":[r["op_B"],sup["lavadora"]["B"],sup["gravacao"]["B"],sup["preset"]["B"],sup["coringa"]["B"],sup["facilitador"]["B"],r["tot_B"]],"Turno C":[r["op_C"],sup["lavadora"]["C"],sup["gravacao"]["C"],sup["preset"]["C"],sup["coringa"]["C"],sup["facilitador"]["C"],r["tot_C"]]}
    st.dataframe(pd.DataFrame(tot_data),use_container_width=True,hide_index=True)
    st.markdown('<div class="mem-step"><span class="step-num">7</span> <b>Produtividades</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">Labor Total ★ = horas_labor ÷ horas_todos</div>', unsafe_allow_html=True)
    prod_data={"Indicador":["Ciclo Operacional","Ciclo Total","Labor Operacional","⭐ Labor Total"],"Resultado":[f"{r['prod_ciclo_op']:.1%}",f"{r['prod_ciclo_tot']:.1%}",f"{r['prod_labor_op']:.1%}",f"{r['prod_labor_tot']:.1%}"]}
    st.dataframe(pd.DataFrame(prod_data),use_container_width=True,hide_index=True)
    buf=BytesIO(); df_intermediario[df_intermediario.mes==mes].to_excel(buf,index=False); buf.seek(0)
    st.download_button("📥 Baixar base completa pós-JOIN",data=buf,file_name=f"base_{mes}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ════════════════════════════════════════
# INTERFACE PRINCIPAL
# ════════════════════════════════════════
st.markdown("""
<div class="jd-header">
  <h1>🏭 Calculadora de Recursos — Usinagem</h1>
  <p>Ferramenta de planejamento de headcount por turno · John Deere Manufatura</p>
</div>
""", unsafe_allow_html=True)

with st.expander("👋 Primeira vez aqui? Veja como usar em 3 passos", expanded=False):
    col1,col2,col3=st.columns(3)
    with col1:
        st.markdown('<div class="mem-step"><span class="step-num">1</span> <b>Suba seu arquivo Excel</b><br><br>O app lê: INPUT_PMP, IMPUTTEMPO, IMPUTDISTRIBUIÇÃO, IMPUTAPLICAÇÃO, IMPUTTURNOS</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="mem-step"><span class="step-num">2</span> <b>Confira os resultados</b><br><br>Veja headcount por turno. Compare com seu Excel atual na aba Comparar.</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="mem-step"><span class="step-num">3</span> <b>Crie cenários</b><br><br>Simule alterações por mês OU para o ANO inteiro de uma vez.</div>', unsafe_allow_html=True)

uploaded=st.file_uploader("Upload do arquivo de inputs (.xlsm ou .xlsx)",type=["xlsm","xlsx"])
if not uploaded:
    st.info("👆 Faça upload do arquivo para começar.")
    st.stop()

file_bytes=uploaded.read()
_file_id=hash(file_bytes)
if st.session_state.get("_file_id")!=_file_id:
    for _k in ["diag_mensal_buf","diag_inp_buf","layout_buf","tabelona_buf","tab_pura_buf"]:
        st.session_state.pop(_k,None)
    st.session_state["_file_id"]=_file_id
if "log_leitura" not in st.session_state: st.session_state.log_leitura=[]
st.session_state.log_leitura=[]

_abas_status=verificar_abas(file_bytes)
_abas_falt=[a for a in ["INPUT_PMP","IMPUTTEMPO","IMPUTDISTRIBUIÇÃO","IMPUTAPLICAÇÃO"] if not _abas_status.get(a)]
if _abas_falt:
    st.error(f"🔴 Abas obrigatórias não encontradas: {', '.join(_abas_falt)}")
    st.stop()

with st.spinner("Lendo planilha..."):
    try:
        log=st.session_state.log_leitura
        pmp,dias=read_pmp(file_bytes,log)
        tempo=read_tempo(file_bytes,log)
        dist=read_dist(file_bytes,log)
        aplic=read_aplic(file_bytes,log)
        turnos_arq,_turnos_ok=read_turnos(file_bytes)
        st.session_state["turnos_arq"]=turnos_arq
        log.append(f"✅ Leitura concluída em {datetime.now().strftime('%H:%M:%S')}")
    except ValueError as e:
        st.error(f"🔴 Erro de formato: {e}"); st.stop()
    except Exception as e:
        st.error(f"🔴 Erro inesperado: {e}"); st.stop()

st.success(f"✅ {len(aplic)} combinações · {pmp.modelo.nunique()} modelos · {pmp.mes.nunique()} meses")
erros,alertas,oks=validar(pmp,tempo,dist,aplic,dias)
n_prob=len(erros)+len(alertas)
label_exp=(f"🔴 {len(erros)} erro(s)  " if erros else "")+(f"⚠️ {len(alertas)} aviso(s)" if alertas else "")+("✅ Inputs validados sem problemas" if not n_prob else "")
with st.expander(label_exp,expanded=bool(erros)):
    for e in erros: st.markdown(f'<div class="aviso-erro">🔴 <b>ERRO:</b> {e}</div>',unsafe_allow_html=True)
    for a in alertas: st.markdown(f'<div class="aviso-warn">⚠️ {a}</div>',unsafe_allow_html=True)
    for o in oks: st.markdown(f'<div class="aviso-ok">✅ {o}</div>',unsafe_allow_html=True)
if erros: st.error("Corrija os erros antes de continuar."); st.stop()

# ── SIDEBAR
with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    _def=st.session_state.get("turnos_arq",{"A":7.5,"B":14.25,"C":19.5})
    st.markdown("**Horas acumuladas por turno (IMPUTTURNOS)**")
    hA=st.number_input("Turno A",value=_def["A"],step=0.01,format="%.2f")
    hB=st.number_input("Turno B",value=_def["B"],step=0.01,format="%.2f")
    hC=st.number_input("Turno C",value=_def["C"],step=0.01,format="%.2f")
    horas_turno={"A":hA,"B":hB,"C":hC}
    st.markdown("---")
    st.markdown("**Horas efetivas por turno**")
    heA=st.number_input("A (efetivas)",value=8.80,step=0.01,format="%.2f",key="input_heA")
    heB=st.number_input("B (efetivas)",value=8.23,step=0.01,format="%.2f",key="input_heB")
    heC=st.number_input("C (efetivas)",value=7.68,step=0.01,format="%.2f",key="input_heC")
    horas_efetivas={"A":heA,"B":heB,"C":heC}
    st.markdown("---")
    st.markdown("**Thresholds de ativação (%)**")
    thr_A=st.number_input("A abre quando ocup.A >",value=40,min_value=0,max_value=200,step=1)
    thr_B=st.number_input("B abre quando ocup.A >",value=106,min_value=0,max_value=200,step=1)
    thr_C=st.number_input("C abre quando ocup.B >",value=100,min_value=0,max_value=200,step=1)
    thresholds={"A":thr_A,"B":thr_B,"C":thr_C}
    st.markdown("---")
    st.markdown("**Funções de suporte**")
    suporte_cfg={}
    for key,label,defs in [("lavadora","Lavadora e Inspeção",{"A":1,"B":1,"C":0}),("gravacao","Gravação e Estanqueidade",{"A":1,"B":1,"C":0}),("preset","Preset",{"A":2,"B":1,"C":1}),("coringa","Coringa",{"A":1,"B":0,"C":0}),("facilitador","Facilitador",{"A":1,"B":1,"C":0})]:
        with st.expander(f"🔧 {label}"):
            modo=st.radio("Modo",["Automático","Manual"],key=f"m_{key}",horizontal=True,label_visibility="collapsed")
            if modo=="Automático":
                st.caption(f"Padrão: A={defs['A']} · B={defs['B']} · C={defs['C']}")
                suporte_cfg[key]={"modo":"auto",**defs}
            else:
                c1,c2,c3=st.columns(3)
                vA=c1.number_input("A",0,10,defs["A"],key=f"s_{key}_A"); vB=c2.number_input("B",0,10,defs["B"],key=f"s_{key}_B"); vC=c3.number_input("C",0,10,defs["C"],key=f"s_{key}_C")
                suporte_cfg[key]={"modo":"manual","A":vA,"B":vB,"C":vC}

tab_vis,tab_inp,tab_mem,tab_res,tab_cen,tab_cmp,tab_exp=st.tabs(["🏠 Visão Geral","📂 Dados de Input","🔬 Como foi Calculado","📊 Resultado por Mês","🎯 Cenários","🔄 Comparar com Excel","📥 Exportar"])

@st.cache_data(show_spinner=False)
def calcular_cached(pmp_hash,_pmp,_tempo,_dist,_aplic,dias_hash,dias,hA,hB,hC,heA,heB,heC,tA,tB,tC,_sup):
    return calcular(_pmp,_tempo,_dist,_aplic,dias,{"A":hA,"B":hB,"C":hC},{"A":tA,"B":tB,"C":tC},_sup,horas_efetivas={"A":heA,"B":heB,"C":heC},retornar_intermediarios=True)

pmp_hash=hash(pmp.to_json()); dias_hash=hash(str(dias))
res_base,df_interm,agg_interm=calcular_cached(pmp_hash,pmp,tempo,dist,aplic,dias_hash,dias,horas_turno["A"],horas_turno["B"],horas_turno["C"],horas_efetivas["A"],horas_efetivas["B"],horas_efetivas["C"],thresholds["A"],thresholds["B"],thresholds["C"],suporte_cfg)

# ── TAB 1 VISÃO GERAL
with tab_vis:
    st.plotly_chart(grafico_cenarios({"Base":res_base}),use_container_width=True)
    meses_ok=[m for m in MESES if res_base.get(m)]
    if meses_ok:
        media_labor=np.mean([res_base[m]["prod_labor_tot"] for m in meses_ok])
        max_total=max(res_base[m]["total"] for m in meses_ok)
        min_total=min(res_base[m]["total"] for m in meses_ok)
        mes_pico=max(meses_ok,key=lambda m:res_base[m]["total"]); mes_vale=min(meses_ok,key=lambda m:res_base[m]["total"])
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Meses calculados",len(meses_ok)); c2.metric("⭐ Labor Total médio",f"{media_labor:.0%}")
        c3.metric("Pico de headcount",f"{max_total} func.",delta=f"em {mes_pico[:3].upper()}")
        c4.metric("Variação anual",f"{max_total-min_total} func.")

# ── TAB 2 INPUTS
with tab_inp:
    st.markdown('<div class="jd-section">Dados carregados</div>',unsafe_allow_html=True)
    aba_inp=st.radio("Qual dado conferir?",["INPUT_PMP","IMPUTTEMPO","IMPUTDISTRIBUIÇÃO","IMPUTAPLICAÇÃO"],horizontal=True)
    if aba_inp=="INPUT_PMP": st.dataframe(pmp.head(100),use_container_width=True,hide_index=True)
    elif aba_inp=="IMPUTTEMPO": st.dataframe(tempo.head(100),use_container_width=True,hide_index=True)
    elif aba_inp=="IMPUTDISTRIBUIÇÃO": st.dataframe(dist.head(100),use_container_width=True,hide_index=True)
    elif aba_inp=="IMPUTAPLICAÇÃO": st.dataframe(aplic.head(200),use_container_width=True,hide_index=True)
    log_html="".join([f'<div class="log-line {"log-ok" if "✅" in l else "log-warn" if "⚠️" in l else ""}">{l}</div>' for l in st.session_state.get("log_leitura",[])])
    st.markdown(f'<div style="background:#1A1A1A;padding:12px;border-radius:8px;max-height:180px;overflow-y:auto">{log_html}</div>',unsafe_allow_html=True)
    def to_xlsx(df): b=BytesIO(); df.to_excel(b,index=False); b.seek(0); return b
    c1,c2,c3=st.columns(3)
    c1.download_button("📥 IMPUTTEMPO",data=to_xlsx(tempo),file_name="tempo.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("📥 IMPUTDIST.",data=to_xlsx(dist),file_name="dist.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c3.download_button("📥 IMPUTAPLIC.",data=to_xlsx(aplic),file_name="aplic.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── TAB 3 MEMÓRIA
with tab_mem:
    _opcoes_mem=[m for m in MESES if res_base.get(m)]
    mes_mem=st.selectbox("Mês",_opcoes_mem,key="mes_mem")
    if mes_mem and res_base.get(mes_mem):
        show_memoria(res_base[mes_mem],mes_mem,df_interm,agg_interm,horas_turno,thresholds)

# ── TAB 4 RESULTADOS
with tab_res:
    st.markdown('<div class="jd-section">Resultado por mês</div>',unsafe_allow_html=True)
    _opcoes_res=[m for m in MESES if res_base.get(m)]+["📅 ANO (resumo anual)"]
    mes_r=st.selectbox("Selecione o mês",_opcoes_res,key="mes_r")

    if mes_r=="📅 ANO (resumo anual)":
        _meses_ano=[(m,res_base[m]) for m in MESES if res_base.get(m)]
        if _meses_ano:
            _n=len(_meses_ano); _d=sum(r["dias"] for _,r in _meses_ano)
            _shc=sum(r["h_ciclo"] for _,r in _meses_ano); _shl=sum(r["h_labor"] for _,r in _meses_ano)
            _sha=sum(r["h_ativos"] for _,r in _meses_ano); _sht=sum(r["h_todos"] for _,r in _meses_ano)
            _pl=_shl/_sht if _sht>0 else 0; _pc=_shc/_sht if _sht>0 else 0
            _plo=_shl/_sha if _sha>0 else 0
            _tA=round(sum(r["tot_A"] for _,r in _meses_ano)/_n,1)
            _tB=round(sum(r["tot_B"] for _,r in _meses_ano)/_n,1)
            _tC=round(sum(r["tot_C"] for _,r in _meses_ano)/_n,1)
            _tf=round(sum(r["total"] for _,r in _meses_ano)/_n,1)
            _pico=max(_meses_ano,key=lambda x:x[1]["total"]); _vale=min(_meses_ano,key=lambda x:x[1]["total"])
            _max_tot=max(r["total"] for _,r in _meses_ano)

            # HEADER
            st.markdown(f'''
<div style="background:linear-gradient(135deg,#1F4D19,#0D2A0D);border-radius:12px;padding:16px 22px;margin-bottom:18px;border-left:5px solid #FFDE00;">
  <div style="font-size:12px;color:#7BC67A;font-weight:700;text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px;">📅 Visão Anual</div>
  <div style="font-size:20px;font-weight:800;color:#FFFFFF;">{_n} meses calculados · {_d} dias trabalhados</div>
  <div style="font-size:12px;color:#AAAAAA;margin-top:4px;">
    Pico: <b style="color:#FFDE00">{_pico[0][:3].upper()} ({_pico[1]["total"]} func.)</b> &nbsp;·&nbsp;
    Vale: <b style="color:#7BC67A">{_vale[0][:3].upper()} ({_vale[1]["total"]} func.)</b>
  </div>
</div>''', unsafe_allow_html=True)

            # KPI CARDS
            c1,c2,c3,c4=st.columns(4)
            with c1:
                st.markdown(f'''<div class="kpi-card destaque">
  <div class="kpi-icon">⭐</div>
  <div class="kpi-label">Labor Total Anual</div>
  <div class="kpi-value">{_pl:.1%}</div>
  <div class="kpi-sub">Produtividade real de toda a equipe</div>
</div>''', unsafe_allow_html=True)
            with c2:
                st.markdown(f'''<div class="kpi-card">
  <div class="kpi-icon">👷</div>
  <div class="kpi-label">Média Funcionários / Mês</div>
  <div class="kpi-value">{_tf:.0f}</div>
  <div class="kpi-sub"><span class="tp tpA">A {_tA:.0f}</span><span class="tp tpB">B {_tB:.0f}</span><span class="tp tpC">C {_tC:.0f}</span></div>
</div>''', unsafe_allow_html=True)
            with c3:
                st.markdown(f'''<div class="kpi-card">
  <div class="kpi-icon">🔄</div>
  <div class="kpi-label">Ciclo Total Anual</div>
  <div class="kpi-value">{_pc:.1%}</div>
  <div class="kpi-sub">Labor Operacional: {_plo:.1%}</div>
</div>''', unsafe_allow_html=True)
            with c4:
                _var=_max_tot-min(r["total"] for _,r in _meses_ano)
                st.markdown(f'''<div class="kpi-card">
  <div class="kpi-icon">📈</div>
  <div class="kpi-label">Variação Anual</div>
  <div class="kpi-value">{_var}</div>
  <div class="kpi-sub">func. entre pico e vale</div>
</div>''', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # GAUGE + BARRAS
            col_g,col_b=st.columns([1,2])
            with col_g:
                st.markdown('<div class="jd-sub">Labor Total Anual</div>',unsafe_allow_html=True)
                pct_int=int(_pl*100)
                cor_g="#69F0AE" if _pl>=0.45 else ("#FFDE00" if _pl>=0.30 else "#FF5252")
                ang=min(180,int(_pl*180)); rad=math.radians(ang)
                ex=80+60*math.cos(math.pi-rad); ey=80-60*math.sin(rad)
                lg="1" if ang>90 else "0"
                st.markdown(f'''<div class="gauge-wrap">
  <svg viewBox="0 0 160 100" width="180">
    <path d="M 20 80 A 60 60 0 0 1 140 80" fill="none" stroke="#2A2A2A" stroke-width="12" stroke-linecap="round"/>
    <path d="M 20 80 A 60 60 0 {lg} 1 {ex:.1f} {ey:.1f}" fill="none" stroke="{cor_g}" stroke-width="12" stroke-linecap="round"/>
    <text x="80" y="76" text-anchor="middle" font-size="20" font-weight="900" fill="{cor_g}">{pct_int}%</text>
    <text x="80" y="92" text-anchor="middle" font-size="9" fill="#888">Labor Total</text>
  </svg>
  <div style="font-size:11px;color:#7BC67A;">Meta: acima de 40%</div>
</div>''', unsafe_allow_html=True)
                pct2=int(_pc*100); ang2=min(180,int(_pc*180)); rad2=math.radians(ang2)
                ex2=80+60*math.cos(math.pi-rad2); ey2=80-60*math.sin(rad2); lg2="1" if ang2>90 else "0"
                st.markdown(f'''<div class="gauge-wrap" style="margin-top:8px;">
  <svg viewBox="0 0 160 100" width="140">
    <path d="M 20 80 A 60 60 0 0 1 140 80" fill="none" stroke="#2A2A2A" stroke-width="10" stroke-linecap="round"/>
    <path d="M 20 80 A 60 60 0 {lg2} 1 {ex2:.1f} {ey2:.1f}" fill="none" stroke="#7BC67A" stroke-width="10" stroke-linecap="round"/>
    <text x="80" y="76" text-anchor="middle" font-size="18" font-weight="800" fill="#7BC67A">{pct2}%</text>
    <text x="80" y="92" text-anchor="middle" font-size="9" fill="#888">Ciclo Total</text>
  </svg>
</div>''', unsafe_allow_html=True)

            with col_b:
                st.markdown('<div class="jd-sub">Funcionários por Mês</div>',unsafe_allow_html=True)
                _mb=max(r["total"] for _,r in _meses_ano) or 1
                bars="".join([f'''<div class="mes-row">
  <div class="mes-nome">{m[:3].upper()}</div>
  <div style="flex:1;display:flex;align-items:center;gap:6px;">
    <div class="mes-bar"><div class="mes-bar-fill" style="width:{r["total"]/_mb*100:.0f}%"></div></div>
    <div style="display:flex;gap:3px;min-width:90px;"><span class="tp tpA">{r["tot_A"]}</span><span class="tp tpB">{r["tot_B"]}</span><span class="tp tpC">{r["tot_C"]}</span></div>
  </div>
  <div class="mes-num">{r["total"]}</div>
  <div class="mes-labor" style="color:{"#69F0AE" if r["prod_labor_tot"]>=0.45 else ("#FFDE00" if r["prod_labor_tot"]>=0.30 else "#FF5252")}">{r["prod_labor_tot"]:.0%}</div>
</div>''' for m,r in _meses_ano])
                st.markdown(bars,unsafe_allow_html=True)

            st.markdown("<br>",unsafe_allow_html=True)
            st.markdown('<div class="jd-sub">Tabela detalhada</div>',unsafe_allow_html=True)
            _rows=[]
            for _m,_r in _meses_ano:
                _rows.append({"Mês":_m,"Dias":_r["dias"],"Turno A":_r["tot_A"],"Turno B":_r["tot_B"],"Turno C":_r["tot_C"],"Total":_r["total"],"Labor Total":f'{_r["prod_labor_tot"]:.1%}',"Labor Op.":f'{_r["prod_labor_op"]:.1%}',"Ciclo Total":f'{_r["prod_ciclo_tot"]:.1%}'})
            _rows.append({"Mês":"📅 MÉDIA ANO","Dias":_d,"Turno A":_tA,"Turno B":_tB,"Turno C":_tC,"Total":_tf,"Labor Total":f'{_pl:.1%}',"Labor Op.":f'{_plo:.1%}',"Ciclo Total":f'{_pc:.1%}'})
            def _sty_ano(row):
                if "ANO" in str(row["Mês"]): return [f"background-color:{JD_AMARELO};color:{JD_VERDE_ESC};font-weight:700"]*len(row)
                styles=[""]*len(row)
                try:
                    lv=float(str(row["Labor Total"]).strip("%"))/100
                    cor="#003D10" if lv>=0.45 else ("#3D2D00" if lv>=0.30 else "#3D0000")
                    txt="#B9F6CA" if lv>=0.45 else ("#FFE57F" if lv>=0.30 else "#FF8A80")
                    for i,col in enumerate(pd.DataFrame(_rows).columns):
                        if col=="Labor Total": styles[i]=f"background-color:{cor};color:{txt};font-weight:600"
                except: pass
                return styles
            st.dataframe(pd.DataFrame(_rows).style.apply(_sty_ano,axis=1),use_container_width=True,hide_index=True)

    elif mes_r and res_base.get(mes_r):
        r=res_base[mes_r]
        st.markdown(f'<div class="aviso-ok">📋 <b>{mes_r}</b> — {r["dias"]} dias &nbsp;|&nbsp; A: <b>{r["tot_A"]}</b> &nbsp;|&nbsp; B: <b>{r["tot_B"]}</b> &nbsp;|&nbsp; C: <b>{r["tot_C"]}</b> &nbsp;|&nbsp; <b>Total: {r["total"]} func.</b></div>',unsafe_allow_html=True)
        show_tabela(r)

# ── TAB 5 CENÁRIOS
with tab_cen:
    if "cenarios" not in st.session_state: st.session_state.cenarios={}
    st.markdown('<div class="jd-section">Simulador de cenários</div>',unsafe_allow_html=True)
    st.caption("Crie variações alterando turnos ativos por centro — por mês ou para o ANO inteiro.")

    with st.expander("➕ Criar novo cenário",expanded=len(st.session_state.cenarios)==0):
        ca,cb=st.columns([2,1])
        novo_nome=ca.text_input("Nome do cenário",placeholder="Ex: Redução turno B novembro")
        opcoes_cen=[m for m in MESES if res_base.get(m)]+["📅 ANO (todos os meses)"]
        mes_novo=cb.selectbox("Mês base",opcoes_cen,key="mes_novo")
        eh_ano_novo=mes_novo=="📅 ANO (todos os meses)"

        if eh_ano_novo:
            _meses_ativos=[m for m in MESES if res_base.get(m)]
            _centros_set=set()
            for _m in _meses_ativos: _centros_set.update(res_base[_m]["centros"].centro.tolist())
            centros_list=sorted(_centros_set)
            ocup_ref={}
            for cen in centros_list:
                vA,vB,vC,aA,aB,aC=[],[],[],[],[],[]
                for _m in _meses_ativos:
                    rc_=res_base[_m]["centros"]; row_=rc_[rc_.centro==cen]
                    if not row_.empty:
                        r_=row_.iloc[0]
                        vA.append(r_.ocup_A); vB.append(r_.ocup_B); vC.append(r_.ocup_C)
                        aA.append(int(r_.ativo_A)); aB.append(int(r_.ativo_B)); aC.append(int(r_.ativo_C))
                ocup_ref[cen]={"oA":np.mean(vA) if vA else 0,"oB":np.mean(vB) if vB else 0,"oC":np.mean(vC) if vC else 0,"aA":round(np.mean(aA)) if aA else 0,"aB":round(np.mean(aB)) if aB else 0,"aC":round(np.mean(aC)) if aC else 0}
            st.markdown(f'<div class="aviso-warn">📅 <b>Modo ANO</b> — override aplicado em todos os {len(_meses_ativos)} meses. Ocupação exibida = média anual.</div>',unsafe_allow_html=True)
        else:
            if mes_novo and res_base.get(mes_novo):
                centros_list=sorted(res_base[mes_novo]["centros"].centro.tolist())
                ocup_ref={}
                for cen in centros_list:
                    rc_=res_base[mes_novo]["centros"]; row_=rc_[rc_.centro==cen]
                    if not row_.empty:
                        r_=row_.iloc[0]
                        ocup_ref[cen]={"oA":r_.ocup_A,"oB":r_.ocup_B,"oC":r_.ocup_C,"aA":int(r_.ativo_A),"aB":int(r_.ativo_B),"aC":int(r_.ativo_C)}
            else:
                centros_list=[]; ocup_ref={}

        if novo_nome and centros_list:
            cols_h=st.columns([3,1,1,1])
            cols_h[0].markdown("**Centro — ocup. média anual**" if eh_ano_novo else "**Centro — ocup. A/B/C**")
            cols_h[1].markdown("**A**"); cols_h[2].markdown("**B**"); cols_h[3].markdown("**C**")
            novo_ov={}
            for cen in centros_list:
                ref=ocup_ref.get(cen,{"oA":0,"oB":0,"oC":0,"aA":0,"aB":0,"aC":0})
                eA="🔴" if ref["oA"]>1 else ("🟡" if ref["oA"]>=0.85 else "🟢")
                eB="🔴" if ref["oB"]>1 else ("🟡" if ref["oB"]>=0.85 else "🟢")
                eC="🔴" if ref["oC"]>1 else ("🟡" if ref["oC"]>=0.85 else "🟢")
                c0,c1,c2,c3=st.columns([3,1,1,1])
                c0.markdown(f"`{cen}` {eA}{ref['oA']:.0%}/{eB}{ref['oB']:.0%}/{eC}{ref['oC']:.0%}")
                vA=c1.number_input("A",0,5,ref["aA"],key=f"n_{novo_nome}_{cen}_A",label_visibility="collapsed")
                vB=c2.number_input("B",0,5,ref["aB"],key=f"n_{novo_nome}_{cen}_B",label_visibility="collapsed")
                vC=c3.number_input("C",0,5,ref["aC"],key=f"n_{novo_nome}_{cen}_C",label_visibility="collapsed")
                novo_ov[cen]={"A":vA,"B":vB,"C":vC}

            if st.button("💾 Salvar cenário",type="primary",key="btn_salvar_cen"):
                if eh_ano_novo:
                    _meses_ativos=[m for m in MESES if res_base.get(m)]
                    ov_c={m:novo_ov for m in _meses_ativos}
                else:
                    ov_c={mes_novo:novo_ov}
                res_cen=calcular(pmp,tempo,dist,aplic,dias,horas_turno,thresholds,suporte_cfg,horas_efetivas=horas_efetivas,overrides=ov_c)
                st.session_state.cenarios[novo_nome]={"resultados":res_cen,"mes":mes_novo,"overrides":ov_c,"eh_ano":eh_ano_novo}
                st.success(f"✅ '{novo_nome}' salvo!"); st.rerun()

    if st.session_state.cenarios:
        todos={"📌 Base":res_base}
        todos.update({k:v["resultados"] for k,v in st.session_state.cenarios.items()})
        st.plotly_chart(grafico_cenarios(todos),use_container_width=True)

        st.markdown('<div class="jd-sub">📊 Comparação detalhada</div>',unsafe_allow_html=True)
        opcoes_cmp=[m for m in MESES if res_base.get(m)]+["📅 ANO (todos os meses)"]
        mes_cmp=st.selectbox("Mês para comparar",opcoes_cmp,key="mes_cmp_r")
        eh_ano_cmp=mes_cmp=="📅 ANO (todos os meses)"
        meses_cmp_lista=[m for m in MESES if res_base.get(m)] if eh_ano_cmp else ([mes_cmp] if res_base.get(mes_cmp) else [])

        if meses_cmp_lista:
            r_base_agg=agregar_ano(res_base,meses_cmp_lista)
            sufixo=" (méd.)" if eh_ano_cmp else ""
            cmp_rows=[]
            for nm,res in todos.items():
                r_agg=agregar_ano(res,meses_cmp_lista)
                if not r_agg or not r_base_agg: continue
                is_base="Base" in nm
                dA=round(r_agg["tot_A"]-r_base_agg["tot_A"],1) if not is_base else "—"
                dB=round(r_agg["tot_B"]-r_base_agg["tot_B"],1) if not is_base else "—"
                dC=round(r_agg["tot_C"]-r_base_agg["tot_C"],1) if not is_base else "—"
                dT=round(r_agg["total"]-r_base_agg["total"],1) if not is_base else "—"
                dL=f'{r_agg["prod_labor_tot"]-r_base_agg["prod_labor_tot"]:+.1%}' if not is_base else "—"
                cmp_rows.append({"Cenário":nm,f"Turno A{sufixo}":r_agg["tot_A"],f"Turno B{sufixo}":r_agg["tot_B"],f"Turno C{sufixo}":r_agg["tot_C"],f"Total{sufixo}":r_agg["total"],"Labor Tot.":f'{r_agg["prod_labor_tot"]:.1%}',"Ciclo Tot.":f'{r_agg["prod_ciclo_tot"]:.1%}',"ΔA":f"{dA:+.1f}" if isinstance(dA,float) else dA,"ΔB":f"{dB:+.1f}" if isinstance(dB,float) else dB,"ΔC":f"{dC:+.1f}" if isinstance(dC,float) else dC,"Δ Total":f"{dT:+.1f}" if isinstance(dT,float) else dT,"Δ Labor":dL})
            df_cmp=pd.DataFrame(cmp_rows)
            def _sty_cmp(row):
                is_base="Base" in str(row["Cenário"])
                if is_base: return [f"background-color:{JD_VERDE_ESC};color:#FFFFFF;font-weight:700"]*len(row)
                styles=[""]*len(row)
                try:
                    d=float(str(row["Δ Total"]).replace("+",""))
                    cd="#003D10" if d<0 else ("#3D0000" if d>0 else ""); td="#B9F6CA" if d<0 else ("#FF8A80" if d>0 else "")
                    for i,col in enumerate(df_cmp.columns):
                        if col in("ΔA","ΔB","ΔC","Δ Total","Δ Labor"): styles[i]=f"background-color:{cd};color:{td};font-weight:600"
                except: pass
                return styles
            st.dataframe(df_cmp.style.apply(_sty_cmp,axis=1),use_container_width=True,hide_index=True)

            for nome_cen,dados_cen in st.session_state.cenarios.items():
                r_cen_res=dados_cen["resultados"]
                with st.expander("🔍 Detalhamento — " + nome_cen + " vs Base"):
                    _m_ref=meses_cmp_lista[0] if meses_cmp_lista else None
                    _meses_prod=meses_cmp_lista if eh_ano_cmp else ([_m_ref] if _m_ref else [])

                    def _calc_prod(res_d, meses_l):
                        rr=[res_d.get(m) for m in meses_l if res_d.get(m)]
                        if not rr: return None
                        shc=sum(r["h_ciclo"] for r in rr); shl=sum(r["h_labor"] for r in rr)
                        sha=sum(r["h_ativos"] for r in rr); sht=sum(r["h_todos"] for r in rr)
                        return {"ciclo_op":shc/sha if sha>0 else 0,"ciclo_tot":shc/sht if sht>0 else 0,
                                "labor_op":shl/sha if sha>0 else 0,"labor_tot":shl/sht if sht>0 else 0}

                    prod_b=_calc_prod(res_base,_meses_prod)
                    prod_c=_calc_prod(r_cen_res,_meses_prod)

                    if prod_b and prod_c:
                        st.markdown('<div class="jd-sub">Produtividades — Base vs Cenário</div>',unsafe_allow_html=True)
                        _items=[
                            ("Ciclo Operacional",prod_b["ciclo_op"],prod_c["ciclo_op"],False),
                            ("Ciclo Total",prod_b["ciclo_tot"],prod_c["ciclo_tot"],False),
                            ("Labor Operacional",prod_b["labor_op"],prod_c["labor_op"],False),
                            ("⭐ Labor Total",prod_b["labor_tot"],prod_c["labor_tot"],True),
                        ]
                        # 4 cards lado a lado
                        parts=[]
                        for lbl,vb,vc,dest in _items:
                            delta=vc-vb
                            arrow="↑" if delta>0 else ("↓" if delta<0 else "→")
                            cor_d="#69F0AE" if delta>0 else ("#FF5252" if delta<0 else "#888888")
                            bg="linear-gradient(135deg,#1F4D19,#0D2A0D)" if dest else "linear-gradient(135deg,#151525,#0D0D1A)"
                            brd="#FFDE00" if dest else "#2A3A4A"
                            parts.append(
                                '<div style="background:' + bg + ';border:1.5px solid ' + brd + ';border-radius:10px;padding:12px 14px;">'
                                + '<div style="font-size:9px;color:#7BC67A;text-transform:uppercase;letter-spacing:.05em;font-weight:600;margin-bottom:8px;">' + lbl + '</div>'
                                + '<div style="display:flex;justify-content:space-between;align-items:flex-end;">'
                                + '<div><div style="font-size:9px;color:#888;margin-bottom:1px;">Base</div>'
                                + '<div style="font-size:19px;font-weight:800;color:#AAAAAA;">' + f"{vb:.1%}" + '</div></div>'
                                + '<div style="font-size:16px;color:#444;padding-bottom:3px;">→</div>'
                                + '<div style="text-align:right"><div style="font-size:9px;color:#FFDE00;margin-bottom:1px;">Cenário</div>'
                                + '<div style="font-size:19px;font-weight:800;color:#FFFFFF;">' + f"{vc:.1%}" + '</div></div>'
                                + '</div>'
                                + '<div style="margin-top:6px;padding-top:6px;border-top:1px solid #333;display:flex;align-items:center;gap:5px;">'
                                + '<span style="font-size:13px;">' + arrow + '</span>'
                                + '<span style="font-size:13px;font-weight:700;color:' + cor_d + ';">' + f"{delta:+.1%}" + '</span>'
                                + '<span style="font-size:10px;color:#666;">vs base</span>'
                                + '</div>'
                                + '</div>'
                            )
                        cards_html = '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:14px;">' + "".join(parts) + '</div>'
                        st.markdown(cards_html, unsafe_allow_html=True)

                        # tabela compacta
                        _prod_rows=[]
                        for lbl,vb,vc,dest in _items:
                            delta=vc-vb
                            _prod_rows.append({"Indicador":lbl,"Base":f"{vb:.1%}","Cenário":f"{vc:.1%}","Δ":f"{delta:+.1%}","":("✅" if delta>0 else ("⚠️" if delta<0 else "➖"))})
                        _df_pr=pd.DataFrame(_prod_rows)
                        def _sty_pr(row):
                            dest2="Labor Total" in str(row["Indicador"])
                            bg_d=f"background-color:{JD_AMARELO};color:{JD_VERDE_ESC};font-weight:700" if dest2 else ""
                            try:
                                dv=float(str(row["Δ"]).replace("+","").replace("%",""))/100
                                cd2="background-color:#003D10;color:#B9F6CA;font-weight:700" if dv>0 else ("background-color:#3D0000;color:#FF8A80;font-weight:700" if dv<0 else "background-color:#222;color:#888")
                            except: cd2=""
                            return [bg_d if dest2 else "" for col in _df_pr.columns[:-2]] + [cd2, bg_d if dest2 else ""]
                        st.dataframe(_df_pr.style.apply(_sty_pr,axis=1),use_container_width=True,hide_index=True)

                    st.markdown('<div class="jd-sub">Ativação por centro</div>',unsafe_allow_html=True)
                    det_rows=[]
                    if _m_ref and res_base.get(_m_ref) and r_cen_res.get(_m_ref):
                        centros_set2=sorted(set(res_base[_m_ref]["centros"].centro.tolist()+r_cen_res[_m_ref]["centros"].centro.tolist()))
                        for cen in centros_set2:
                            rb_c=res_base[_m_ref]["centros"]; rc_c=r_cen_res[_m_ref]["centros"]
                            rb_r=rb_c[rb_c.centro==cen]; rc_r=rc_c[rc_c.centro==cen]
                            if rb_r.empty or rc_r.empty: continue
                            rb=rb_r.iloc[0]; rc=rc_r.iloc[0]
                            mA=int(rb.ativo_A)!=int(rc.ativo_A)
                            mB=int(rb.ativo_B)!=int(rc.ativo_B)
                            mC=int(rb.ativo_C)!=int(rc.ativo_C)
                            det_rows.append({
                                "Centro":cen,
                                "Ocup.A":f"{rb.ocup_A:.0%}","Base A":int(rb.ativo_A),"Cen A":int(rc.ativo_A),
                                "Ocup.B":f"{rb.ocup_B:.0%}","Base B":int(rb.ativo_B),"Cen B":int(rc.ativo_B),
                                "Ocup.C":f"{rb.ocup_C:.0%}","Base C":int(rb.ativo_C),"Cen C":int(rc.ativo_C),
                                "Mudou":"✅ Igual" if not(mA or mB or mC) else
                                    ("A " if mA else "")+("B " if mB else "")+("C" if mC else "")+"alterado(s)"
                            })
                    if det_rows:
                        df_det=pd.DataFrame(det_rows)
                        def _sty_det(row):
                            if "alterado" in str(row["Mudou"]):
                                return ["background-color:#3D2D00;color:#FFE57F"]*len(row)
                            return [""]*len(row)
                        st.dataframe(df_det.style.apply(_sty_det,axis=1),use_container_width=True,hide_index=True)
                    if dados_cen.get("eh_ano"):
                        st.markdown('<div class="aviso-ok">📅 Cenário anual — override em todos os meses. Detalhamento acima = ' + str(_m_ref) + '</div>',unsafe_allow_html=True)
        st.markdown("---")
        col_exp,col_del=st.columns([3,1])
        with col_exp:
            for nm,v in st.session_state.cenarios.items():
                if v.get("eh_ano"):
                    _m_exp=next((m for m in MESES if res_base.get(m)),None)
                else:
                    _m_exp=v.get("mes",MESES[0])
                if _m_exp and res_base.get(_m_exp):
                    label_dl=f"📥 {nm} ({'ANO' if v.get('eh_ano') else _m_exp})"
                    fname_dl=f"cenario_{nm.replace(' ','_')}_{'ANO' if v.get('eh_ano') else _m_exp}.xlsx"
                    st.download_button(label_dl,data=exportar_cenario_vs_base(res_base,v["resultados"],_m_exp,nm),file_name=fname_dl,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key=f"dl_cen_{nm}")
        with col_del:
            if st.session_state.cenarios:
                dn=st.selectbox("Remover",list(st.session_state.cenarios.keys()),key="del_c")
                if st.button("🗑️ Remover",type="secondary",key="btn_del_cen"):
                    del st.session_state.cenarios[dn]; st.rerun()
    else:
        st.info("Nenhum cenário criado ainda.")

# ── TAB 6 COMPARAÇÃO
with tab_cmp:
    st.markdown('<div class="jd-section">Comparação com o Excel atual</div>',unsafe_allow_html=True)
    cache_key=f"cmp_{hash(str(dias))}_{hash(str(thresholds))}_{hash(str(horas_turno))}"
    if st.session_state.get("cmp_cache_key")!=cache_key:
        with st.spinner("Comparando..."):
            _r,_d,_e=comparar_com_excel(res_base,file_bytes,tempo,dist,aplic,pmp,dias,horas_turno,thresholds,suporte_cfg)
        st.session_state["cmp_cache_key"]=cache_key; st.session_state["cmp_cache_resumo"]=_r; st.session_state["cmp_cache_detalhe"]=_d; st.session_state["cmp_cache_err"]=_e
    df_resumo=st.session_state["cmp_cache_resumo"]; err=st.session_state["cmp_cache_err"]
    if err: st.error(err)
    elif df_resumo is not None and len(df_resumo)>0:
        n_ok=(df_resumo["Status"].str.startswith("✅")).sum() if "Status" in df_resumo else 0
        n_warn=(df_resumo["Status"].str.startswith("🟡")).sum() if "Status" in df_resumo else 0
        n_err=(df_resumo["Status"].str.startswith("🔴")).sum() if "Status" in df_resumo else 0
        c1,c2,c3=st.columns(3)
        c1.metric("✅ Meses iguais",n_ok); c2.metric("🟡 Pequena diferença",n_warn); c3.metric("🔴 Com divergência",n_err)
        def build_vis(df):
            rows=[]
            for _,r in df.iterrows():
                def cell(app,excel,delta):
                    try:
                        d_n=int(str(delta).replace("+",""))
                        icon="✅" if d_n==0 else ("🟡" if abs(d_n)<=2 else "🔴")
                        return f"{icon} App={app} | Excel={excel} ({delta})" if d_n!=0 else f"✅ {app}"
                    except: return f"App={app} / Excel={excel}"
                rows.append({"Mês":r["Mês"],"Status":r["Status"],"Turno A":cell(r.get("CEN A App","?"),r.get("CEN A Excel","?"),r.get("Δ A","?")),"Turno B":cell(r.get("CEN B App","?"),r.get("CEN B Excel","?"),r.get("Δ B","?")),"Turno C":cell(r.get("CEN C App","?"),r.get("CEN C Excel","?"),r.get("Δ C","?")),"Total":cell(r.get("Total App","?"),r.get("Total Excel","?"),r.get("Δ Total","?")),"Labor":f"App={r.get('Labor App','?')} / Excel={r.get('Labor Excel','?')}"})
            return pd.DataFrame(rows)
        df_vis=build_vis(df_resumo)
        def sty_res(row):
            st_v=str(row.get("Status",""))
            if "✅" in st_v: base="background-color:#003D10;color:#B9F6CA"
            elif "🟡" in st_v: base="background-color:#3D2D00;color:#FFE57F"
            elif "🔴" in st_v: base="background-color:#3D0000;color:#FF8A80"
            else: base="background-color:#2D1A00;color:#FFD54F"
            styles=[]
            for col in row.index:
                if col=="Status": styles.append(base)
                elif col in("Turno A","Turno B","Turno C","Total"):
                    val=str(row[col])
                    if "🔴" in val: styles.append("background-color:#3D0000;color:#FF8A80")
                    elif "🟡" in val: styles.append("background-color:#3D2D00;color:#FFE57F")
                    elif "✅" in val: styles.append("background-color:#003D10;color:#B9F6CA")
                    else: styles.append("")
                else: styles.append("")
            return styles
        st.dataframe(df_vis.astype(str).style.apply(sty_res,axis=1),use_container_width=True,hide_index=True)
    else:
        st.warning("Nenhuma aba mensal encontrada (NovFY26, DezFY26…)")

# ── TAB 7 EXPORTAR
with tab_exp:
    st.markdown('<div class="jd-section">Exportação</div>',unsafe_allow_html=True)
    c1,c2=st.columns(2)
    with c1:
        st.markdown("**Resultado completo (todas as abas)**")
        st.download_button("📥 Baixar resultado base",data=exportar(res_base,tempo,dist,aplic,pmp),file_name="resultado_usinagem.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_res_base")
    with c2:
        st.markdown("**Base tratada (pós-JOIN)**")
        def _to_xlsx_full():
            b=BytesIO(); df_interm.to_excel(b,index=False); b.seek(0); return b
        st.download_button("📥 Baixar base tratada",data=_to_xlsx_full(),file_name="base_tratada.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_base")
    if st.session_state.get("cenarios"):
        st.markdown('<div class="jd-sub">Cenários salvos</div>',unsafe_allow_html=True)
        for nm,v in st.session_state.cenarios.items():
            st.download_button(f"📥 Cenário: {nm}",data=exportar(v["resultados"],tempo,dist,aplic,pmp),file_name=f"cenario_{nm.replace(' ','_')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key=f"exp_{nm}")
