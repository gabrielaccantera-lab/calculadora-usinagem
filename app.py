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

st.set_page_config(page_title="Calculadora de Recursos — Usinagem", layout="wide", page_icon="🏭")

# ─── PALETA JOHN DEERE ───────────────────
JD_VERDE       = "#367C2B"
JD_VERDE_ESC   = "#1F4D19"
JD_AMARELO     = "#FFDE00"
JD_AMARELO_ESC = "#C9A800"
JD_CINZA_CLR   = "#F4F4F4"
JD_CINZA_BD    = "#D0D0D0"
JD_TEXTO       = "#1A1A1A"

st.markdown("""
<style>

.jd-header {
    background: #1F4D19;
    padding: 16px 24px; border-radius: 10px;
    border-left: 6px solid #FFDE00;
    margin-bottom: 20px;
}
.jd-header h1 { color: white; margin: 0; font-size: 21px; font-weight: 700; }
.jd-header p  { color: #b8d4b4; margin: 4px 0 0; font-size: 12px; }
.jd-section {
    font-size: 15px; font-weight: 700; color: #1F4D19;
    border-left: 4px solid #FFDE00; padding-left: 10px;
    margin: 22px 0 10px;
}
.jd-sub {
    font-size: 12px; font-weight: 600; color: #367C2B;
    margin: 12px 0 4px; text-transform: uppercase; letter-spacing: .04em;
}
.badge-ok    { background:#367C2B;color:white;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600; }
.badge-warn  { background:#C9A800;color:white;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600; }
.badge-erro  { background:#C62828;color:white;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600; }
.aviso-erro  { background:#FFEBEE;border-left:4px solid #C62828;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#B71C1C; }
.aviso-warn  { background:#FFFDE7;border-left:4px solid #C9A800;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#7B5800; }
.aviso-ok    { background:#E8F5E9;border-left:4px solid #367C2B;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#1F4D19; }
.formula-box {
    background:#1A1A1A;color:#A8E6A3;font-family:monospace;
    padding:10px 14px;border-radius:6px;font-size:12px;line-height:1.8;
    border-left:3px solid #FFDE00;
}
.mem-step {
    background:white;border:1px solid #D0D0D0;border-radius:8px;
    padding:12px 16px;margin:6px 0;
}
.mem-step .step-num {
    background:#1F4D19;color:white;border-radius:50%;
    width:22px;height:22px;display:inline-flex;align-items:center;
    justify-content:center;font-size:11px;font-weight:700;margin-right:8px;
}
.diff-igual   { background:#E8F5E9;color:#1F4D19;font-weight:600;padding:2px 6px;border-radius:4px; }
.diff-pequeno { background:#FFFDE7;color:#7B5800;font-weight:600;padding:2px 6px;border-radius:4px; }
.diff-grande  { background:#FFEBEE;color:#B71C1C;font-weight:600;padding:2px 6px;border-radius:4px; }
.log-line { font-family:monospace;font-size:11px;color:#444;padding:1px 0; }
.log-ok   { color:#367C2B; }
.log-warn { color:#C9A800; }
.log-err  { color:#C62828; }
.cenario-card{background:#F8FAFC;border:1.5px solid #B5D4F4;border-radius:10px;padding:14px 16px;margin:6px 0}
.cenario-card.ativo{border-color:#1F4D19;background:#E8F5E9}
</style>
""", unsafe_allow_html=True)

MESES       = ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
               "Maio","Junho","Julho","Agosto","Setembro","Outubro"]
MESES_ABREV = ["NOV","DEZ","JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT"]

# ─────────────────────────────────────────
# MAPEAMENTO DE COLUNAS (por nome, não índice)
# ─────────────────────────────────────────
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
        "div_volume": ["Divisão \nde\nVolume\nENTRE\nPEÇAS","Div Volume","DIV_VOLUME","div_volume"],
        "disponib":   ["Disponi-\nbilidade","Disponibilidade","DISPONIB","disponib"],
    },
}

def find_col(df, candidates, aba, campo):
    """Busca coluna por lista de nomes candidatos. Fallback por índice com aviso."""
    for c in candidates:
        if c in df.columns:
            return c
    # fallback posicional com aviso
    idx_fallback = {"centro":0,"peca":1,"t_ciclo":5,"t_labor":6,
                    "div_carga":7,"div_volume":9,"disponib":10}
    if campo in idx_fallback:
        idx = idx_fallback[campo]
        if idx < len(df.columns):
            st.session_state.setdefault("log_leitura",[]).append(
                f"⚠️ [{aba}] Campo '{campo}' não encontrado por nome — usando coluna {idx} ({df.columns[idx]}) como fallback")
            return df.columns[idx]
    raise ValueError(f"[{aba}] Campo '{campo}' não encontrado. Colunas disponíveis: {list(df.columns)}")

# ─────────────────────────────────────────
# LEITURA COM LOG
# ─────────────────────────────────────────
def read_pmp(fb, log):
    df = pd.read_excel(BytesIO(fb), sheet_name='INPUT_PMP', header=None)
    log.append(f"✅ INPUT_PMP lido: {df.shape[0]} linhas × {df.shape[1]} colunas")
    dias = {}
    for i, m in enumerate(MESES, 1):
        v = df.iloc[0, i] if i < df.shape[1] else None
        dias[m] = int(v) if pd.notna(v) else 0
    log.append(f"   Dias trabalhados: { {m:d for m,d in dias.items() if d>0} }")

    rows = []
    modelos_lidos = 0
    for r in range(2, len(df)):
        modelo = df.iloc[r, 0]
        if pd.isna(modelo): continue
        modelos_lidos += 1
        for i, m in enumerate(MESES, 1):
            v = df.iloc[r, i] if i < df.shape[1] else None
            qtd = int(v) if pd.notna(v) else 0
            if qtd > 0:
                rows.append({"modelo": str(modelo).strip(), "mes": m, "qtd": qtd})
    log.append(f"   {modelos_lidos} modelos · {len(rows)} registros com qtd > 0")
    return pd.DataFrame(rows), dias

def read_tempo(fb, log):
    df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTTEMPO', header=0)
    log.append(f"✅ IMPUTTEMPO lido: {df.shape[0]} linhas · colunas: {list(df.columns[:4])}")
    mp = COL_MAP["IMPUTTEMPO"]
    c = {k: find_col(df, v, "IMPUTTEMPO", k) for k,v in mp.items()}
    out = df[[c["centro"],c["peca"],c["t_ciclo"],c["t_labor"]]].copy()
    out.columns = ["centro","peca","t_ciclo","t_labor"]
    out = out.dropna(subset=["centro"])
    nulos_ciclo = out["t_ciclo"].isna().sum()
    nulos_labor = out["t_labor"].isna().sum()
    if nulos_ciclo: log.append(f"⚠️ IMPUTTEMPO: {nulos_ciclo} linhas com t_ciclo nulo")
    if nulos_labor: log.append(f"⚠️ IMPUTTEMPO: {nulos_labor} linhas com t_labor nulo")
    log.append(f"   {len(out)} combinações centro+peça carregadas")
    return out.copy()

def read_dist(fb, log):
    df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTDISTRIBUIÇÃO', header=0)
    log.append(f"✅ IMPUTDISTRIBUIÇÃO lido: {df.shape[0]} linhas")
    mp = COL_MAP["IMPUTDISTRIBUIÇÃO"]
    c = {k: find_col(df, v, "IMPUTDISTRIBUIÇÃO", k) for k,v in mp.items()}
    out = df[[c["centro"],c["peca"],c["div_carga"],c["div_volume"],c["disponib"]]].copy()
    out.columns = ["centro","peca","div_carga","div_volume","disponib"]
    out = out.dropna(subset=["centro"])
    zero_d = (out["disponib"] == 0).sum()
    if zero_d: log.append(f"⚠️ IMPUTDISTRIBUIÇÃO: {zero_d} linhas com disponib=0")
    log.append(f"   {len(out)} combinações carregadas")
    return out.copy()

def read_aplic(fb, log):
    df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTAPLICAÇÃO', header=0)
    log.append(f"✅ IMPUTAPLICAÇÃO lido: {df.shape[0]} linhas")
    df = df.rename(columns={df.columns[0]:"centro", df.columns[1]:"peca"})
    mcols = [c for c in df.columns if str(c).startswith("MODELO")]
    log.append(f"   {len(mcols)} modelos encontrados na matriz")
    melted = df[["centro","peca"]+mcols].melt(id_vars=["centro","peca"], var_name="modelo", value_name="ativo")
    out = melted[melted["ativo"]==1][["centro","peca","modelo"]].reset_index(drop=True)
    log.append(f"   {len(out)} combinações centro+peça+modelo ativas (flag=1)")
    return out

# ─────────────────────────────────────────
# VALIDAÇÕES
# ─────────────────────────────────────────
def validar(pmp, tempo, dist, aplic, dias):
    erros, alertas, oks = [], [], []
    chaves_tempo = set(zip(tempo.centro, tempo.peca))
    chaves_dist  = set(zip(dist.centro,  dist.peca))
    chaves_aplic = set(zip(aplic.centro, aplic.peca))

    zero_disp = dist[dist.disponib == 0]
    if len(zero_disp):
        ex = ", ".join([f"{r.centro}/{r.peca}" for _,r in zero_disp.iterrows()][:5])
        erros.append(f"Disponibilidade = 0 em {len(zero_disp)} linha(s) — divisão por zero: {ex}")
    diff_td = chaves_tempo - chaves_dist
    if diff_td:
        erros.append(f"{len(diff_td)} combinação(ões) centro+peça em IMPUTTEMPO sem IMPUTDISTRIBUIÇÃO: {list(diff_td)[:3]}")
    sem_aplic = chaves_tempo - chaves_aplic
    if sem_aplic:
        alertas.append(f"{len(sem_aplic)} centro+peça sem nenhum modelo em IMPUTAPLICAÇÃO (não gerarão carga): {list(sem_aplic)[:3]}")
    modelos_sem = set(pmp.modelo.unique()) - set(aplic.modelo.unique())
    if modelos_sem:
        alertas.append(f"{len(modelos_sem)} modelo(s) com demanda mas sem aplicação: {', '.join(list(modelos_sem)[:5])}")
    merged = tempo.merge(dist, on=["centro","peca"], how="inner")
    labor_maior = merged[merged.t_labor > merged.t_ciclo]
    if len(labor_maior):
        alertas.append(f"{len(labor_maior)} linha(s) com t_labor > t_ciclo (fisicamente improvável): {[(r.centro,r.peca) for _,r in labor_maior.iterrows()][:3]}")
    for m in MESES:
        qtd_m = pmp[pmp.mes==m].qtd.sum() if len(pmp[pmp.mes==m]) else 0
        if qtd_m > 0 and dias.get(m,0) == 0:
            alertas.append(f"Mês '{m}' tem {int(qtd_m)} peças de demanda mas dias trabalhados = 0.")
    nulos_t = tempo[["t_ciclo","t_labor"]].isna().sum()
    if nulos_t.sum() > 0:
        alertas.append(f"Valores nulos em IMPUTTEMPO: t_ciclo={nulos_t['t_ciclo']}, t_labor={nulos_t['t_labor']}")
    nulos_d = dist[["div_carga","div_volume","disponib"]].isna().sum()
    if nulos_d.sum() > 0:
        alertas.append(f"Valores nulos em IMPUTDISTRIBUIÇÃO: {dict(nulos_d[nulos_d>0])}")
    if not erros and not alertas:
        oks.append("Todos os inputs foram validados sem inconsistências.")
    return erros, alertas, oks

# ─────────────────────────────────────────
# CÁLCULO COM RASTREABILIDADE
# ─────────────────────────────────────────
def calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg,
             overrides=None, retornar_intermediarios=False):

    # JOIN completo — rastreável
    df = (aplic.merge(pmp, on="modelo")
               .merge(tempo, on=["centro","peca"])
               .merge(dist,  on=["centro","peca"]))
    df["indice_ciclo"] = (df.t_ciclo * df.div_carga * df.div_volume) / df.disponib
    df["min_ciclo"]    = df.indice_ciclo * df.qtd
    df["min_labor"]    = df.t_labor * df.div_carga * df.qtd
    agg = df.groupby(["centro","mes"])[["min_ciclo","min_labor"]].sum().reset_index()

    thr_A = thresholds["A"] / 100
    thr_B = thresholds["B"] / 100
    thr_C = thresholds["C"] / 100
    hA, hB, hC = horas_turno["A"], horas_turno["B"], horas_turno["C"]

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
                "horas_disp_A":d*hA*aA,"horas_disp_B":d*hB*aB,"horas_disp_C":d*hC*aC,
            })

        df_c = pd.DataFrame(centros)
        op_A = int(df_c.ativo_A.sum()); op_B = int(df_c.ativo_B.sum()); op_C = int(df_c.ativo_C.sum())

        def get_sup(key, t, op_count):
            cfg = suporte_cfg[key]
            if cfg["modo"] == "auto":
                defaults = {"lavadora":{"A":1,"B":1,"C":0},"gravacao":{"A":1,"B":1,"C":0},
                            "preset":{"A":2,"B":1,"C":1},"coringa":{"A":1,"B":0,"C":0},
                            "facilitador":{"A":1,"B":1,"C":0}}
                return defaults[key][t] if op_count > 0 else 0
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
        h_todos  = tot_A*d*hA+tot_B*d*hB+tot_C*d*hC

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
            "dias":d,"hA":hA,"hB":hB,"hC":hC,
            "thr_A":thr_A,"thr_B":thr_B,"thr_C":thr_C,
            "minA":d*hA*60,"minB":d*hB*60,"minC":d*hC*60,
        }

    if retornar_intermediarios:
        return resultados, df, agg
    return resultados

# ─────────────────────────────────────────
# TABELA RESULTADO
# ─────────────────────────────────────────
def show_tabela(r):
    dias=r["dias"]; hA,hB,hC=r["hA"],r["hB"],r["hC"]
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
            "Horas A":round(s["A"]*hA*dias,1),"Horas B":round(s["B"]*hB*dias,1),"Horas C":round(s["C"]*hC*dias,1)})
    srows.append({"Função":"▶ TOTAL POR TURNO",
        "Qtd A":r["tot_A"],"Qtd B":r["tot_B"],"Qtd C":r["tot_C"],
        "Horas A":round(r["tot_A"]*hA*dias,1),"Horas B":round(r["tot_B"]*hB*dias,1),"Horas C":round(r["tot_C"]*hC*dias,1)})
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

# ─────────────────────────────────────────
# GRÁFICO
# ─────────────────────────────────────────
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
        yaxis_title="Nº Funcionários",
        yaxis2=dict(title="Labor Total (%)",tickformat=".0f",ticksuffix="%",range=[0,100]),
        legend=dict(orientation="h",y=-0.3,font=dict(size=10)),
        height=440,plot_bgcolor="white",paper_bgcolor="white",
        xaxis=dict(showgrid=False),yaxis=dict(showgrid=True,gridcolor="#E8E8E8"))
    return fig

# ─────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────
def exportar(resultados):
    out=BytesIO(); wb=openpyxl.Workbook()
    brd=Border(left=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'),
               top=Side(style='thin',color='CCCCCC'),bottom=Side(style='thin',color='CCCCCC'))
    def ec(c,bg="FFFFFF",fg="000000",bold=False,fmt=None,center=True):
        c.font=Font(name="Arial",bold=bold,color=fg,size=9)
        c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center" if center else "left",vertical="center")
        c.border=brd
        if fmt: c.number_format=fmt
    ws=wb.active; ws.title="RESUMO MO"
    JD_V=JD_VERDE_ESC.replace("#",""); JD_Y=JD_AMARELO.replace("#","")
    for i,h in enumerate(["Mês","Dias","Turno A","Turno B","Turno C","Total",
                           "Ciclo Op.","Ciclo Total","Labor Op.","Labor Total ★"],1):
        ec(ws.cell(1,i,h),JD_V,"FFFFFF",True)
    for ri,(m,abr) in enumerate(zip(MESES,MESES_ABREV),2):
        r=resultados.get(m); bg="EAF3FB" if ri%2==0 else "FFFFFF"
        vals=[abr,0,"-","-","-","-","-","-","-","-"] if not r else [
            abr,r["dias"],r["tot_A"],r["tot_B"],r["tot_C"],r["total"],
            r["prod_ciclo_op"],r["prod_ciclo_tot"],r["prod_labor_op"],r["prod_labor_tot"]]
        for ci,v in enumerate(vals,1):
            c=ws.cell(ri,ci,v); fmt="0%" if ci>=7 and isinstance(v,float) else None
            ec(c,JD_Y if ci==10 and isinstance(v,float) else bg,
               JD_V if ci==10 and isinstance(v,float) else "000000",
               ci==10 and isinstance(v,float),fmt)
        ws.row_dimensions[ri].height=15
    for mes in MESES:
        r=resultados.get(mes)
        if not r: continue
        wsm=wb.create_sheet(mes[:10]); hA,hB,hC,dias=r["hA"],r["hB"],r["hC"],r["dias"]
        for ci,txt in [(1,""),(2,"TURNO A"),(3,"TURNO B"),(4,"TURNO C"),
                       (5,"TURNO A"),(6,"TURNO B"),(7,"TURNO C"),
                       (8,"TURNO A"),(9,"TURNO B"),(10,"TURNO C")]:
            ec(wsm.cell(1,ci,txt),JD_V,"FFFFFF",True)
        wsm.cell(1,1,mes.upper()); ec(wsm.cell(1,1),JD_V,"FFFFFF",True)
        def cbg(v):
            if v>1.0: return "FFCDD2"
            if v>=0.85: return "FFFDE7"
            return "E8F5E9"
        ri=3
        for _,row in r["centros"].iterrows():
            for ci,(val,bg,ctr) in enumerate([
                (row.centro,"FFFFFF",False),
                (f"{row.ocup_A:.0%}",cbg(row.ocup_A),True),(f"{row.ocup_B:.0%}",cbg(row.ocup_B),True),(f"{row.ocup_C:.0%}",cbg(row.ocup_C),True),
                (row.ativo_A,"B3E5FC" if row.ativo_A else "FFFDE7",True),(row.ativo_B,"B3E5FC" if row.ativo_B else "FFFDE7",True),(row.ativo_C,"B3E5FC" if row.ativo_C else "FFFDE7",True),
                (f"{row.horas_disp_A:.2f}" if row.ativo_A else "0","B3E5FC" if row.ativo_A else "F5F5F5",True),
                (f"{row.horas_disp_B:.2f}" if row.ativo_B else "0","B3E5FC" if row.ativo_B else "F5F5F5",True),
                (f"{row.horas_disp_C:.2f}" if row.ativo_C else "0","B3E5FC" if row.ativo_C else "F5F5F5",True)],1):
                ec(wsm.cell(ri,ci,val),bg,center=ctr)
            ri+=1
        sup=r["suporte"]
        for nome,key in [("TOTAL DE OPERADORES",None),("LAVADORA E INSPEÇÃO","lavadora"),
                         ("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),("PRESET","preset"),
                         ("CORINGA","coringa"),("FACILITADOR","facilitador"),
                         ("TOTAL POR TURNO",None),("TOTAL FUNCIONÁRIOS",None)]:
            bold="TOTAL" in nome; bg_r=JD_Y if bold else "FFFFFF"; fg_r=JD_V if bold else "000000"
            ec(wsm.cell(ri,1,nome),bg_r,fg_r,bold,center=False)
            if key:
                s=sup[key]
                for ci,t in [(5,"A"),(6,"B"),(7,"C")]:
                    ec(wsm.cell(ri,ci,s[t]),"B3E5FC" if s[t] else "FFFDE7",bold=bold)
                for ci,t,h in [(8,"A",hA),(9,"B",hB),(10,"C",hC)]:
                    v=s[t]*h*dias; ec(wsm.cell(ri,ci,f"{v:.2f}" if v else "0"),"B3E5FC" if v else "F5F5F5",bold=bold)
            elif "TOTAL DE OPERADORES" in nome:
                for ci,v in [(5,r["op_A"]),(6,r["op_B"]),(7,r["op_C"])]:
                    ec(wsm.cell(ri,ci,v),JD_Y,JD_V,True)
                for ci,v,h in [(8,r["op_A"],hA),(9,r["op_B"],hB),(10,r["op_C"],hC)]:
                    ec(wsm.cell(ri,ci,f"{v*h*dias:.2f}"),JD_Y,JD_V,True)
            elif "TOTAL POR TURNO" in nome:
                for ci,v in [(5,r["tot_A"]),(6,r["tot_B"]),(7,r["tot_C"])]:
                    ec(wsm.cell(ri,ci,v),JD_Y,JD_V,True)
                for ci,v,h in [(8,r["tot_A"],hA),(9,r["tot_B"],hB),(10,r["tot_C"],hC)]:
                    ec(wsm.cell(ri,ci,f"{v*h*dias:.2f}"),JD_Y,JD_V,True)
            elif "FUNCIONÁRIOS" in nome:
                ec(wsm.cell(ri,4,r["total"]),JD_Y,JD_V,True)
                tot_h=r["tot_A"]*hA*dias+r["tot_B"]*hB*dias+r["tot_C"]*hC*dias
                ec(wsm.cell(ri,8,f"{tot_h:.2f}"),JD_Y,JD_V,True)
            ri+=1
        ri+=1
        for nm,v,dest in [("PROD. CICLO OPERACIONAL",r["prod_ciclo_op"],False),
                          ("PROD. CICLO TOTAL",r["prod_ciclo_tot"],False),
                          ("PROD. LABOR OPERACIONAL",r["prod_labor_op"],False),
                          ("PROD. LABOR TOTAL ★",r["prod_labor_tot"],True)]:
            wsm.merge_cells(f"H{ri}:I{ri}")
            ec(wsm.cell(ri,8,nm),JD_Y if dest else "FFFFFF",JD_V if dest else "000000",dest,center=False)
            ec(wsm.cell(ri,10,v),JD_Y if dest else "FFFFFF",JD_V if dest else "000000",dest,"0%")
            ri+=1
        for ci,w in enumerate([14,8,8,8,8,8,8,24,10,10],1):
            wsm.column_dimensions[get_column_letter(ci)].width=w
    wb.save(out); out.seek(0)
    return out

# ─────────────────────────────────────────
# COMPARAÇÃO COM EXCEL REFERÊNCIA
# ─────────────────────────────────────────
def comparar_com_excel(res_app, fb_ref):
    """Lê resultado do Excel de referência e compara com o app."""
    try:
        xl = pd.ExcelFile(BytesIO(fb_ref))
        sheets = xl.sheet_names
    except Exception as e:
        return None, f"Erro ao ler arquivo de referência: {e}"

    # Tentar encontrar aba de resumo
    resumo_sheet = None
    for s in sheets:
        if "RESUMO" in s.upper() or "MO" in s.upper():
            resumo_sheet = s; break
    if not resumo_sheet and sheets:
        resumo_sheet = sheets[0]

    try:
        df_ref = pd.read_excel(BytesIO(fb_ref), sheet_name=resumo_sheet, header=None)
    except Exception as e:
        return None, f"Erro ao ler aba '{resumo_sheet}': {e}"

    comparacoes = []
    for mes, abr in zip(MESES, MESES_ABREV):
        r_app = res_app.get(mes)
        if not r_app: continue

        # Procurar linha do mês no Excel de referência
        linha_ref = None
        for idx, row in df_ref.iterrows():
            for val in row:
                if str(val).upper().strip() in [mes.upper(), abr]:
                    linha_ref = row
                    break
            if linha_ref is not None: break

        if linha_ref is None:
            comparacoes.append({"Mês": abr, "Status": "⚠️ não encontrado", "Ref Total": "-",
                                 "App Total": r_app["total"], "Δ": "-", "Δ%": "-", "Causa": "mês não localizado no Excel"})
            continue

        # Extrair valores numéricos da linha
        nums = [v for v in linha_ref if pd.notna(v) and isinstance(v, (int, float))]
        ref_total = None
        for n in nums:
            if 10 <= n <= 200:  # range razoável para total funcionários
                ref_total = int(n); break

        if ref_total is None:
            comparacoes.append({"Mês": abr, "Status": "⚠️ valor não extraído", "Ref Total": "-",
                                 "App Total": r_app["total"], "Δ": "-", "Δ%": "-", "Causa": "não foi possível extrair total de funcionários"})
            continue

        delta = r_app["total"] - ref_total
        delta_pct = (delta / ref_total * 100) if ref_total != 0 else 0
        abs_delta = abs(delta)

        if abs_delta == 0: status = "✅ Igual"
        elif abs_delta <= 2: status = "🟡 Diferença pequena"
        else: status = "🔴 Divergência"

        causa = ""
        if abs_delta > 0:
            if abs_delta <= 1: causa = "Possível diferença de arredondamento ou 1 suporte manual"
            elif abs_delta <= 3: causa = "Provável diferença em threshold de ativação ou suporte configurado manualmente"
            else: causa = "Divergência significativa — verificar demanda, thresholds ou mapeamento de colunas"

        comparacoes.append({
            "Mês": abr, "Status": status,
            "Ref Total": ref_total, "App Total": r_app["total"],
            "Δ": f"{delta:+d}", "Δ%": f"{delta_pct:+.1f}%", "Causa provável": causa
        })

    return pd.DataFrame(comparacoes), None

# ─────────────────────────────────────────
# MEMÓRIA DE CÁLCULO
# ─────────────────────────────────────────
def show_memoria(r, mes, df_intermediario, agg, horas_turno, thresholds):
    st.markdown(f'<div class="jd-section">Memória de cálculo — {mes}</div>', unsafe_allow_html=True)

    steps = [
        ("1", "Inputs utilizados",
         f"Dias trabalhados: **{r['dias']}**  |  Turno A: **{r['hA']}h**  |  Turno B: **{r['hB']}h**  |  Turno C: **{r['hC']}h**\n\n"
         f"Min. disponíveis: A = **{r['minA']:.0f} min** ({r['dias']}×{r['hA']}×60)  |  B = **{r['minB']:.0f} min**  |  C = **{r['minC']:.0f} min**"),
        ("2", "Fórmula do índice de ciclo",
         None),
        ("3", "Fórmula dos minutos necessários por linha",
         None),
        ("4", "Agrupamento por centro",
         None),
        ("5", "Regras de ativação de turno",
         f"Turno A ativo se ocup_A > **{thresholds['A']}%**\n\n"
         f"Turno B ativo se ocup_A > **{thresholds['B']}%** (A sobrecarregado)\n\n"
         f"Turno C ativo se ocup_B > **{thresholds['C']}%** (B sobrecarregado)"),
        ("6", "Output final",
         f"Operadores CEN: A={r['op_A']} · B={r['op_B']} · C={r['op_C']}\n\n"
         f"Total funcionários: **{r['total']}** (CEN + suportes)\n\n"
         f"Labor Total: **{r['prod_labor_tot']:.1%}**  ({r['h_labor']:.1f}h labor ÷ {r['h_todos']:.1f}h totais)"),
    ]

    for num, titulo, conteudo in steps:
        st.markdown(f'<div class="mem-step"><span class="step-num">{num}</span> <b>{titulo}</b></div>', unsafe_allow_html=True)
        if num == "2":
            st.markdown('<div class="formula-box">indice_ciclo = (t_ciclo × div_carga × div_volume) ÷ disponib</div>', unsafe_allow_html=True)
        elif num == "3":
            st.markdown('<div class="formula-box">min_ciclo = indice_ciclo × qtd_pecas_mes<br>min_labor = t_labor × div_carga × qtd_pecas_mes</div>', unsafe_allow_html=True)
        elif num == "4":
            st.markdown('<div class="formula-box">GROUP BY centro + mes → SUM(min_ciclo), SUM(min_labor)<br>ocup_A = total_min_ciclo ÷ min_disp_A</div>', unsafe_allow_html=True)
        elif conteudo:
            st.markdown(conteudo)

    # Tabela intermediária por centro
    st.markdown('<div class="jd-sub">Tabela intermediária — centro × ocupação × ativação</div>', unsafe_allow_html=True)
    df_show = r["centros"][["centro","min_ciclo_total","min_labor_total",
                              "min_disp_A","ocup_A","ocup_B","ocup_C",
                              "ativo_A","ativo_B","ativo_C"]].copy()
    df_show["ocup_A"] = df_show.ocup_A.map(lambda x: f"{x:.1%}")
    df_show["ocup_B"] = df_show.ocup_B.map(lambda x: f"{x:.1%}")
    df_show["ocup_C"] = df_show.ocup_C.map(lambda x: f"{x:.1%}")
    df_show["min_ciclo_total"] = df_show.min_ciclo_total.round(1)
    df_show["min_labor_total"] = df_show.min_labor_total.round(1)
    st.dataframe(df_show, use_container_width=True, hide_index=True)

    # Download da base intermediária
    buf = BytesIO()
    df_intermediario[df_intermediario.mes==mes].to_excel(buf, index=False)
    buf.seek(0)
    st.download_button("📥 Baixar base tratada (pós-JOIN, pré-agrupamento)",
        data=buf, file_name=f"base_tratada_{mes}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ─────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────
st.markdown(f"""
<div class="jd-header">
  <h1>🏭 Calculadora de Recursos — Usinagem</h1>
  <p>Planejamento de headcount por turno · Auditável · Comparável com Excel · John Deere</p>
</div>
""", unsafe_allow_html=True)

with st.expander("📋 Como preparar o arquivo de upload", expanded=False):
    st.markdown("""
**O app lê 5 abas do seu `.xlsm` ou `.xlsx`. Abas mensais (NovFY26 etc.) não são necessárias.**

| Aba | Conteúdo obrigatório |
|---|---|
| `INPUT_PMP` | Linha 1 = dias por mês · Linhas seguintes = volume por modelo |
| `IMPUTTEMPO` | Colunas: Máquina, Peça, Tempo Ciclo (min), Tempo Labor (min) |
| `IMPUTDISTRIBUIÇÃO` | Colunas: Máquina, Peça, Div. Carga, Div. Volume, Disponibilidade |
| `IMPUTAPLICAÇÃO` | Matriz: Máquina × Peça × MODELO (0 ou 1) |
| `IMPUTTURNOS` | Horas acumuladas por turno (referência) |

> Horas de duração por turno e thresholds de ativação são configuráveis na barra lateral.
    """)

uploaded = st.file_uploader("Upload do arquivo de inputs (.xlsm ou .xlsx)", type=["xlsm","xlsx"])
if not uploaded:
    st.info("👆 Faça upload do arquivo para começar.")
    st.stop()

file_bytes = uploaded.read()

if "log_leitura" not in st.session_state:
    st.session_state.log_leitura = []
st.session_state.log_leitura = []

with st.spinner("Lendo planilha..."):
    try:
        log = st.session_state.log_leitura
        pmp, dias = read_pmp(file_bytes, log)
        tempo      = read_tempo(file_bytes, log)
        dist       = read_dist(file_bytes, log)
        aplic      = read_aplic(file_bytes, log)
        log.append(f"✅ Leitura concluída em {datetime.now().strftime('%H:%M:%S')}")
    except Exception as e:
        st.error(f"Erro ao ler: {e}"); st.stop()

st.success(f"✅ {len(aplic)} combinações · {pmp.modelo.nunique()} modelos · {pmp.mes.nunique()} meses")

erros, alertas, oks = validar(pmp, tempo, dist, aplic, dias)
n_prob = len(erros) + len(alertas)
label_exp = (f"🔴 {len(erros)} erro(s) crítico(s)  " if erros else "") + \
            (f"⚠️ {len(alertas)} aviso(s)" if alertas else "") + \
            ("✅ Inputs validados sem problemas" if not n_prob else "")
with st.expander(label_exp, expanded=bool(erros)):
    for e in erros:
        st.markdown(f'<div class="aviso-erro">🔴 <b>ERRO:</b> {e}</div>', unsafe_allow_html=True)
    for a in alertas:
        st.markdown(f'<div class="aviso-warn">⚠️ {a}</div>', unsafe_allow_html=True)
    for o in oks:
        st.markdown(f'<div class="aviso-ok">✅ {o}</div>', unsafe_allow_html=True)
if erros:
    st.error("Corrija os erros antes de continuar."); st.stop()

# ─── SIDEBAR ─────────────────────────────
with st.sidebar:
    st.markdown(f"## ⚙️ Configurações")
    st.markdown(f"**Duração dos turnos (h)**")
    hA = st.number_input("Turno A", value=8.80, step=0.01, format="%.2f")
    hB = st.number_input("Turno B", value=8.23, step=0.01, format="%.2f")
    hC = st.number_input("Turno C", value=7.68, step=0.01, format="%.2f")
    horas_turno = {"A":hA,"B":hB,"C":hC}

    st.markdown("---")
    st.markdown("**Thresholds de ativação (%)**")
    st.caption("Turno abre quando ocupação ultrapassa esse valor.")
    thr_A = st.number_input("A abre quando ocup.A >", value=40, min_value=0, max_value=200, step=1)
    thr_B = st.number_input("B abre quando ocup.A >", value=106, min_value=0, max_value=200, step=1)
    thr_C = st.number_input("C abre quando ocup.B >", value=100, min_value=0, max_value=200, step=1)
    thresholds = {"A":thr_A,"B":thr_B,"C":thr_C}

    st.markdown("---")
    st.markdown("**Funções de suporte**")
    suporte_cfg = {}
    for key,label,defs in [
        ("lavadora","Lavadora e Inspeção",{"A":1,"B":1,"C":0}),
        ("gravacao","Gravação e Estanqueidade",{"A":1,"B":1,"C":0}),
        ("preset","Preset",{"A":2,"B":1,"C":1}),
        ("coringa","Coringa",{"A":1,"B":0,"C":0}),
        ("facilitador","Facilitador",{"A":1,"B":1,"C":0}),
    ]:
        with st.expander(f"🔧 {label}"):
            modo = st.radio("",["Automático","Manual"],key=f"m_{key}",horizontal=True)
            if modo=="Automático":
                st.caption(f"A={defs['A']} · B={defs['B']} · C={defs['C']}")
                suporte_cfg[key]={"modo":"auto",**defs}
            else:
                c1,c2,c3=st.columns(3)
                vA=c1.number_input("A",0,10,defs["A"],key=f"s_{key}_A")
                vB=c2.number_input("B",0,10,defs["B"],key=f"s_{key}_B")
                vC=c3.number_input("C",0,10,defs["C"],key=f"s_{key}_C")
                suporte_cfg[key]={"modo":"manual","A":vA,"B":vB,"C":vC}
    st.markdown("---")
    st.caption("Alterações afetam todos os cálculos em tempo real.")

# ─── TABS ────────────────────────────────
tab_vis, tab_inp, tab_mem, tab_res, tab_cmp, tab_diag, tab_exp = st.tabs([
    "🏠 Visão Geral", "📂 Inputs", "🔬 Memória de Cálculo",
    "📊 Resultados", "🔄 Comparação", "🩺 Diagnóstico", "📥 Exportação"
])

# Cache do resultado base
@st.cache_data(show_spinner=False)
def calcular_cached(pmp_hash, _pmp, _tempo, _dist, _aplic, dias_hash, dias, hA, hB, hC, tA, tB, tC, _sup):
    return calcular(_pmp, _tempo, _dist, _aplic, dias,
                    {"A":hA,"B":hB,"C":hC}, {"A":tA,"B":tB,"C":tC}, _sup,
                    retornar_intermediarios=True)

pmp_hash = hash(pmp.to_json())
dias_hash = hash(str(dias))
res_base, df_interm, agg_interm = calcular(
    pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg,
    retornar_intermediarios=True)

# ══════════════════════════════════════════
# TAB 1 — VISÃO GERAL
# ══════════════════════════════════════════
with tab_vis:
    st.plotly_chart(grafico_cenarios({"Base": res_base}), use_container_width=True)
    meses_ok = [m for m in MESES if res_base.get(m)]
    if meses_ok:
        media_labor = np.mean([res_base[m]["prod_labor_tot"] for m in meses_ok])
        max_total   = max(res_base[m]["total"] for m in meses_ok)
        min_total   = min(res_base[m]["total"] for m in meses_ok)
        mes_pico    = max(meses_ok, key=lambda m: res_base[m]["total"])
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Meses calculados", len(meses_ok))
        c2.metric("Labor Total médio", f"{media_labor:.0%}")
        c3.metric("Pico de headcount", f"{max_total} func. ({mes_pico[:3].upper()})")
        c4.metric("Variação anual", f"{max_total - min_total} func.")

# ══════════════════════════════════════════
# TAB 2 — INPUTS
# ══════════════════════════════════════════
with tab_inp:
    st.markdown('<div class="jd-section">Preview das bases carregadas</div>', unsafe_allow_html=True)
    aba_inp = st.radio("", ["INPUT_PMP","IMPUTTEMPO","IMPUTDISTRIBUIÇÃO","IMPUTAPLICAÇÃO"], horizontal=True)
    if aba_inp == "INPUT_PMP":
        st.dataframe(pmp.head(100), use_container_width=True, hide_index=True)
        st.caption(f"{len(pmp)} registros com qtd > 0 · {pmp.modelo.nunique()} modelos")
    elif aba_inp == "IMPUTTEMPO":
        st.dataframe(tempo.head(100), use_container_width=True, hide_index=True)
        st.caption(f"{len(tempo)} combinações centro+peça")
    elif aba_inp == "IMPUTDISTRIBUIÇÃO":
        st.dataframe(dist.head(100), use_container_width=True, hide_index=True)
        st.caption(f"{len(dist)} combinações centro+peça")
    elif aba_inp == "IMPUTAPLICAÇÃO":
        st.dataframe(aplic.head(200), use_container_width=True, hide_index=True)
        st.caption(f"{len(aplic)} combinações ativas (flag=1)")

    st.markdown('<div class="jd-section">Log de leitura</div>', unsafe_allow_html=True)
    log_html = "".join([
        f'<div class="log-line {"log-ok" if "✅" in l else "log-warn" if "⚠️" in l else "log-err" if "🔴" in l else ""}">{l}</div>'
        for l in st.session_state.get("log_leitura", [])
    ])
    st.markdown(f'<div style="background:#1A1A1A;padding:12px 16px;border-radius:8px;max-height:220px;overflow-y:auto">{log_html}</div>', unsafe_allow_html=True)

    # Download bases tratadas
    st.markdown('<div class="jd-section">Download das bases tratadas</div>', unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    def to_xlsx(df):
        b=BytesIO(); df.to_excel(b,index=False); b.seek(0); return b
    c1.download_button("📥 IMPUTTEMPO tratado",  data=to_xlsx(tempo),  file_name="tempo_tratado.xlsx",  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("📥 IMPUTDIST. tratado",  data=to_xlsx(dist),   file_name="dist_tratada.xlsx",   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c3.download_button("📥 IMPUTAPLIC. tratado", data=to_xlsx(aplic),  file_name="aplic_tratada.xlsx",  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════
# TAB 3 — MEMÓRIA DE CÁLCULO
# ══════════════════════════════════════════
with tab_mem:
    mes_mem = st.selectbox("Mês", [m for m in MESES if res_base.get(m)], key="mes_mem")
    if mes_mem and res_base.get(mes_mem):
        show_memoria(res_base[mes_mem], mes_mem, df_interm, agg_interm, horas_turno, thresholds)

# ══════════════════════════════════════════
# TAB 4 — RESULTADOS
# ══════════════════════════════════════════
with tab_res:
    if "cenarios" not in st.session_state:
        st.session_state.cenarios = {}

    mes_r = st.selectbox("Mês", [m for m in MESES if res_base.get(m)], key="mes_r")
    if mes_r: show_tabela(res_base[mes_r])

    st.markdown('<div class="jd-section">Simulador de cenários</div>', unsafe_allow_html=True)
    with st.expander("➕ Criar novo cenário", expanded=len(st.session_state.cenarios)==0):
        ca,cb = st.columns([2,1])
        novo_nome = ca.text_input("Nome", placeholder="Ex: Redução turno B novembro")
        mes_novo  = cb.selectbox("Mês base", MESES, key="mes_novo")
        if novo_nome and mes_novo and res_base.get(mes_novo):
            r_orig = res_base[mes_novo]
            centros_list = sorted(r_orig["centros"].centro.tolist())
            cols_h = st.columns([3,1,1,1])
            cols_h[0].markdown("**Centro — ocup. A/B/C**"); cols_h[1].markdown("**A**"); cols_h[2].markdown("**B**"); cols_h[3].markdown("**C**")
            novo_ov = {}
            for cen in centros_list:
                rc = r_orig["centros"][r_orig["centros"].centro==cen].iloc[0]
                eA="🔴" if rc.ocup_A>1 else ("🟡" if rc.ocup_A>=0.85 else "🟢")
                eB="🔴" if rc.ocup_B>1 else ("🟡" if rc.ocup_B>=0.85 else "🟢")
                eC="🔴" if rc.ocup_C>1 else ("🟡" if rc.ocup_C>=0.85 else "🟢")
                c0,c1,c2,c3 = st.columns([3,1,1,1])
                c0.markdown(f"`{cen}` {eA}{rc.ocup_A:.0%}/{eB}{rc.ocup_B:.0%}/{eC}{rc.ocup_C:.0%}")
                vA=c1.number_input("",0,5,int(rc.ativo_A),key=f"n_{novo_nome}_{cen}_A",label_visibility="collapsed",help=f"Base:{rc.ativo_A}")
                vB=c2.number_input("",0,5,int(rc.ativo_B),key=f"n_{novo_nome}_{cen}_B",label_visibility="collapsed",help=f"Base:{rc.ativo_B}")
                vC=c3.number_input("",0,5,int(rc.ativo_C),key=f"n_{novo_nome}_{cen}_C",label_visibility="collapsed",help=f"Base:{rc.ativo_C}")
                novo_ov[cen]={"A":vA,"B":vB,"C":vC}
            if st.button("💾 Salvar cenário", type="primary"):
                ov_c={mes_novo:novo_ov}
                res_cen=calcular(pmp,tempo,dist,aplic,dias,horas_turno,thresholds,suporte_cfg,ov_c)
                st.session_state.cenarios[novo_nome]={"resultados":res_cen,"mes":mes_novo,"overrides":ov_c}
                st.success(f"✅ '{novo_nome}' salvo!"); st.rerun()

    if st.session_state.cenarios:
        todos={"📌 Base":res_base}; todos.update({k:v["resultados"] for k,v in st.session_state.cenarios.items()})
        st.plotly_chart(grafico_cenarios(todos), use_container_width=True)
        mes_cmp=st.selectbox("Mês para comparar",MESES,key="mes_cmp_r")
        cmp_rows=[]
        for nm,res in todos.items():
            r=res.get(mes_cmp)
            if r: cmp_rows.append({"Cenário":nm,"Turno A":r["tot_A"],"Turno B":r["tot_B"],"Turno C":r["tot_C"],"Total":r["total"],"Labor Total":f"{r['prod_labor_tot']:.0%}"})
        if cmp_rows: st.dataframe(pd.DataFrame(cmp_rows),use_container_width=True,hide_index=True)
        cd,ce = st.columns(2)
        with cd:
            dn=st.selectbox("Remover",list(st.session_state.cenarios.keys()),key="del_c")
            if st.button("🗑️ Remover",type="secondary"): del st.session_state.cenarios[dn]; st.rerun()
        with ce:
            for nm,v in st.session_state.cenarios.items():
                st.download_button(f"📥 {nm}",data=exportar(v["resultados"]),
                    file_name=f"cenario_{nm.replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key=f"dl_{nm}")

# ══════════════════════════════════════════
# TAB 5 — COMPARAÇÃO
# ══════════════════════════════════════════
with tab_cmp:
    st.markdown('<div class="jd-section">Comparação com resultado do Excel</div>', unsafe_allow_html=True)
    st.markdown("Suba o arquivo Excel com o resultado esperado. O app tentará extrair os totais por mês e comparar com o resultado calculado.")
    uploaded_ref = st.file_uploader("Arquivo de referência (.xlsx/.xlsm)", type=["xlsx","xlsm"], key="ref_upload")
    if uploaded_ref:
        fb_ref = uploaded_ref.read()
        df_cmp, err = comparar_com_excel(res_base, fb_ref)
        if err:
            st.error(err)
        elif df_cmp is not None:
            st.markdown('<div class="jd-sub">Resultado da comparação</div>', unsafe_allow_html=True)
            n_ok   = (df_cmp["Status"].str.startswith("✅")).sum()
            n_warn = (df_cmp["Status"].str.startswith("🟡")).sum()
            n_err  = (df_cmp["Status"].str.startswith("🔴")).sum()
            c1,c2,c3 = st.columns(3)
            c1.metric("✅ Iguais",    n_ok)
            c2.metric("🟡 Pequena diferença", n_warn)
            c3.metric("🔴 Divergências", n_err)
            def style_cmp(row):
                s=["background-color:#E8F5E9" if "✅" in row["Status"]
                   else "background-color:#FFFDE7" if "🟡" in row["Status"]
                   else "background-color:#FFEBEE"]*len(row)
                return s
            st.dataframe(df_cmp.style.apply(style_cmp,axis=1),use_container_width=True,hide_index=True)
            buf=BytesIO(); df_cmp.to_excel(buf,index=False); buf.seek(0)
            st.download_button("📥 Exportar comparação",data=buf,file_name="comparacao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Faça upload do arquivo de referência para iniciar a comparação.")

# ══════════════════════════════════════════
# TAB 6 — DIAGNÓSTICO
# ══════════════════════════════════════════
with tab_diag:
    st.markdown('<div class="jd-section">Diagnóstico de divergências</div>', unsafe_allow_html=True)
    st.markdown("Quando o resultado do app divergir do Excel, use esta seção para identificar a causa.")

    mes_diag = st.selectbox("Mês para diagnosticar", [m for m in MESES if res_base.get(m)], key="mes_diag")
    ref_total_input = st.number_input("Total de funcionários esperado (do Excel)", min_value=0, max_value=500, value=0, step=1)

    if mes_diag and ref_total_input > 0 and res_base.get(mes_diag):
        r = res_base[mes_diag]
        delta = r["total"] - ref_total_input
        st.markdown("---")
        if delta == 0:
            st.markdown('<div class="aviso-ok">✅ Resultado idêntico ao esperado. Nenhuma divergência.</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="aviso-{"warn" if abs(delta)<=2 else "erro"}">{"⚠️" if abs(delta)<=2 else "🔴"} Diferença de <b>{delta:+d} funcionários</b> ({r["total"]} app vs {ref_total_input} esperado)</div>', unsafe_allow_html=True)
            st.markdown('<div class="jd-sub">Causas mais prováveis — em ordem de investigação</div>', unsafe_allow_html=True)

            causas = []
            if abs(delta) == 1:
                causas.append(("🔢 Arredondamento", "Uma função de suporte pode estar sendo computada diferente (ex: Preset=2 no app vs 1 no Excel).", "Verificar configuração de suporte na sidebar"))
            if abs(delta) <= 3:
                causas.append(("⚙️ Threshold de ativação", f"O turno pode estar sendo ativado com threshold diferente. App usa: A>{thr_A}% / B>{thr_B}% / C>{thr_C}%.", "Conferir thresholds na sidebar e comparar com lógica SE() do Excel"))
            causas.append(("📅 Dias trabalhados", f"App usa {r['dias']} dias para {mes_diag}. Se o Excel usa valor diferente, todos os denominadores de ocupação mudam.", "Conferir célula de dias na aba INPUT_PMP"))
            causas.append(("🗂️ Mapeamento de colunas", "Se colunas foram reordenadas no Excel, o app pode ter lido campo errado (verificar log de leitura na aba Inputs).", "Abrir aba Inputs → Log de leitura e checar avisos de fallback posicional"))
            causas.append(("📊 Demanda divergente", "Volume de peças diferente entre arquivos.", "Comparar soma de qtd por modelo na aba Inputs"))

            for titulo, descricao, acao in causas:
                with st.expander(titulo):
                    st.markdown(f"**O que pode ter acontecido:** {descricao}")
                    st.markdown(f"**👉 Como investigar:** {acao}")

            st.markdown('<div class="jd-sub">Rastreabilidade — onde olhar primeiro</div>', unsafe_allow_html=True)
            df_rastr = r["centros"][["centro","ocup_A","ocup_B","ocup_C","ativo_A","ativo_B","ativo_C","min_ciclo_total"]].copy()
            df_rastr["ocup_A"]=df_rastr.ocup_A.map(lambda x:f"{x:.1%}")
            df_rastr["ocup_B"]=df_rastr.ocup_B.map(lambda x:f"{x:.1%}")
            df_rastr["ocup_C"]=df_rastr.ocup_C.map(lambda x:f"{x:.1%}")
            df_rastr["min_ciclo_total"]=df_rastr.min_ciclo_total.round(0)
            st.dataframe(df_rastr, use_container_width=True, hide_index=True)
            st.caption("Verifique centros onde a ocupação está próxima dos thresholds — são os mais sensíveis a pequenas variações.")
    else:
        st.info("Selecione o mês e informe o total esperado do Excel para iniciar o diagnóstico.")

# ══════════════════════════════════════════
# TAB 7 — EXPORTAÇÃO
# ══════════════════════════════════════════
with tab_exp:
    st.markdown('<div class="jd-section">Exportação</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Resultado completo (todas as abas)**")
        st.caption("RESUMO MO + uma aba por mês com tabela no formato original")
        st.download_button("📥 Baixar resultado base",
            data=exportar(res_base), file_name="resultado_usinagem.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        st.markdown("**Base tratada (pós-JOIN completo)**")
        st.caption("Todos os meses, todas as linhas de centro+peça+modelo+minutos")
        buf=BytesIO(); df_interm.to_excel(buf,index=False); buf.seek(0)
        st.download_button("📥 Baixar base tratada completa",
            data=buf, file_name="base_tratada_completa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if st.session_state.get("cenarios"):
        st.markdown('<div class="jd-sub">Cenários salvos</div>', unsafe_allow_html=True)
        for nm,v in st.session_state.cenarios.items():
            st.download_button(f"📥 Cenário: {nm}", data=exportar(v["resultados"]),
                file_name=f"cenario_{nm.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"exp_{nm}")
