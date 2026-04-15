import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Calculadora de Recursos — Usinagem", layout="wide", page_icon="🏭")
st.markdown("""
<style>
.main-header{background:linear-gradient(90deg,#1E3A5F,#2D6A9F);color:white;padding:18px 24px;border-radius:10px;margin-bottom:20px}
.main-header h1{color:white;margin:0;font-size:22px}
.main-header p{color:#B5D4F4;margin:4px 0 0;font-size:13px}
.section-title{font-size:16px;font-weight:700;color:#1E3A5F;border-bottom:2px solid #1E3A5F;padding-bottom:6px;margin:20px 0 12px}
.subsection{font-size:13px;font-weight:600;color:#2D6A9F;margin:14px 0 6px}
.aviso-erro{background:#FFEBEE;border-left:4px solid #C62828;border-radius:8px;padding:10px 14px;margin:6px 0;font-size:13px;color:#B71C1C}
.aviso-alerta{background:#FFF8E1;border-left:4px solid #F9A825;border-radius:8px;padding:10px 14px;margin:6px 0;font-size:13px;color:#E65100}
.aviso-ok{background:#E8F5E9;border-left:4px solid #2E7D32;border-radius:8px;padding:10px 14px;margin:6px 0;font-size:13px;color:#1B5E20}
.cenario-card{background:#F8FAFC;border:1.5px solid #B5D4F4;border-radius:10px;padding:14px 16px;margin:6px 0}
.cenario-card.ativo{border-color:#1E3A5F;background:#EAF3FB}
</style>
""", unsafe_allow_html=True)

MESES       = ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
               "Maio","Junho","Julho","Agosto","Setembro","Outubro"]
MESES_ABREV = ["NOV","DEZ","JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT"]

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
    melted = df[["centro","peca"]+mcols].melt(id_vars=["centro","peca"], var_name="modelo", value_name="ativo")
    return melted[melted["ativo"]==1][["centro","peca","modelo"]].reset_index(drop=True)

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
        ex = ", ".join([f"{r.centro}/{r.peca}" for _, r in zero_disp.iterrows()][:5])
        erros.append(f"Disponibilidade = 0 em {len(zero_disp)} linha(s) — causará divisão por zero: {ex}")
    diff = chaves_tempo - chaves_dist
    if diff:
        erros.append(f"{len(diff)} combinação(ões) centro+peça em IMPUTTEMPO sem correspondência em IMPUTDISTRIBUIÇÃO.")
    sem_aplic = chaves_tempo - chaves_aplic
    if sem_aplic:
        alertas.append(f"{len(sem_aplic)} combinação(ões) centro+peça sem nenhum modelo em IMPUTAPLICAÇÃO (nunca gerarão carga): {list(sem_aplic)[:3]}")
    modelos_sem = set(pmp.modelo.unique()) - set(aplic.modelo.unique())
    if modelos_sem:
        alertas.append(f"{len(modelos_sem)} modelo(s) com demanda mas sem aplicação: {', '.join(list(modelos_sem)[:5])}")
    merged = tempo.merge(dist, on=["centro","peca"], how="inner")
    labor_maior = merged[merged.t_labor > merged.t_ciclo]
    if len(labor_maior):
        alertas.append(f"{len(labor_maior)} linha(s) com tempo de labor > ciclo: {[(r.centro,r.peca) for _,r in labor_maior.iterrows()][:3]}")
    for m in MESES:
        qtd_m = pmp[pmp.mes==m].qtd.sum() if len(pmp[pmp.mes==m]) else 0
        if qtd_m > 0 and dias.get(m, 0) == 0:
            alertas.append(f"Mês '{m}' tem {qtd_m} peças de demanda mas dias trabalhados = 0.")
    if not erros and not alertas:
        oks.append("Nenhuma inconsistência encontrada nos inputs.")
    return erros, alertas, oks

# ─────────────────────────────────────────
# CÁLCULO
# ─────────────────────────────────────────
def calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg, overrides=None):
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

        minA = d * hA * 60; minB = d * hB * 60; minC = d * hC * 60
        centros = []
        for _, row in sub.iterrows():
            cen = row.centro
            mc, ml = row.min_ciclo, row.min_labor
            pA = mc/minA if minA > 0 else 0
            pB = mc/minB if minB > 0 else 0
            pC = mc/minC if minC > 0 else 0
            aA = 1 if pA > thr_A else 0
            aB = 1 if pA > thr_B else 0
            aC = 1 if pB > thr_C else 0
            if overrides and mes in overrides and cen in overrides[mes]:
                ov = overrides[mes][cen]
                if "A" in ov: aA = ov["A"]
                if "B" in ov: aB = ov["B"]
                if "C" in ov: aC = ov["C"]
            centros.append({
                "centro": cen,
                "ocup_A": pA, "ocup_B": pB, "ocup_C": pC,
                "ativo_A": aA, "ativo_B": aB, "ativo_C": aC,
                "horas_ciclo": mc/60, "horas_labor": ml/60,
                "horas_disp_A": d*hA*aA, "horas_disp_B": d*hB*aB, "horas_disp_C": d*hC*aC,
            })

        df_c = pd.DataFrame(centros)
        op_A = int(df_c.ativo_A.sum())
        op_B = int(df_c.ativo_B.sum())
        op_C = int(df_c.ativo_C.sum())

        # Suporte configurável
        def s(key, turno, op):
            cfg = suporte_cfg[key]
            if cfg["modo"] == "auto":
                defaults = {"lavadora":{"A":1,"B":1,"C":0},"gravacao":{"A":1,"B":1,"C":0},
                            "preset":{"A":2,"B":1,"C":1},"coringa":{"A":1,"B":0,"C":0},
                            "facilitador":{"A":1,"B":1,"C":0}}
                base = defaults[key][turno]
                if turno in ("A","B","C"):
                    if turno=="A" and op==0: return 0
                    if turno=="B" and op_B==0: return 0
                    if turno=="C" and op_C==0: return 0
                return base
            else:
                return cfg[turno]

        lav = {t: s("lavadora",t,op_A if t=="A" else (op_B if t=="B" else op_C)) for t in "ABC"}
        gra = {t: s("gravacao",t,op_A if t=="A" else (op_B if t=="B" else op_C)) for t in "ABC"}
        pre = {t: s("preset",t,op_A if t=="A" else (op_B if t=="B" else op_C)) for t in "ABC"}
        cor = {t: s("coringa",t,op_A if t=="A" else (op_B if t=="B" else op_C)) for t in "ABC"}
        fac = {t: s("facilitador",t,op_A if t=="A" else (op_B if t=="B" else op_C)) for t in "ABC"}

        tot_A = op_A + lav["A"] + gra["A"] + pre["A"] + cor["A"] + fac["A"]
        tot_B = op_B + lav["B"] + gra["B"] + pre["B"] + cor["B"] + fac["B"]
        tot_C = op_C + lav["C"] + gra["C"] + pre["C"] + cor["C"] + fac["C"]
        total = tot_A + tot_B + tot_C

        h_ciclo  = float(df_c.horas_ciclo.sum())
        h_labor  = float(df_c.horas_labor.sum())
        h_ativos = float((df_c.horas_disp_A+df_c.horas_disp_B+df_c.horas_disp_C).sum())
        h_todos  = tot_A*d*hA + tot_B*d*hB + tot_C*d*hC

        resultados[mes] = {
            "centros": df_c,
            "op_A": op_A, "op_B": op_B, "op_C": op_C,
            "tot_A": tot_A, "tot_B": tot_B, "tot_C": tot_C, "total": total,
            "suporte": {"lavadora":lav,"gravacao":gra,"preset":pre,"coringa":cor,"facilitador":fac},
            "h_ciclo": h_ciclo, "h_labor": h_labor,
            "h_ativos": h_ativos, "h_todos": h_todos,
            "prod_ciclo_op":  h_ciclo/h_ativos if h_ativos>0 else 0,
            "prod_ciclo_tot": h_ciclo/h_todos  if h_todos>0 else 0,
            "prod_labor_op":  h_labor/h_ativos if h_ativos>0 else 0,
            "prod_labor_tot": h_labor/h_todos  if h_todos>0 else 0,
            "dias": d, "hA": hA, "hB": hB, "hC": hC,
        }
    return resultados

# ─────────────────────────────────────────
# TABELA
# ─────────────────────────────────────────
def show_tabela(r):
    dias = r["dias"]; hA,hB,hC = r["hA"],r["hB"],r["hC"]

    # Centros
    rows = []
    for _, row in r["centros"].iterrows():
        rows.append({
            "Centro": row.centro,
            "Ocup. A": row.ocup_A, "Ocup. B": row.ocup_B, "Ocup. C": row.ocup_C,
            "Ativo A": int(row.ativo_A), "Ativo B": int(row.ativo_B), "Ativo C": int(row.ativo_C),
            "Horas A": round(row.horas_disp_A,2), "Horas B": round(row.horas_disp_B,2), "Horas C": round(row.horas_disp_C,2),
        })
    df = pd.DataFrame(rows)

    def style_row(row):
        s = [""] * len(row)
        def co(v):
            if v>1.0: return "background-color:#FFCDD2;color:#B71C1C;font-weight:600"
            if v>=0.85: return "background-color:#FFF9C4;color:#F57F17;font-weight:600"
            return "background-color:#C8E6C9;color:#1B5E20;font-weight:600"
        for i,col in enumerate(df.columns):
            if col in ("Ocup. A","Ocup. B","Ocup. C"): s[i]=co(row.iloc[i])
            elif col in ("Ativo A","Ativo B","Ativo C"):
                s[i]="background-color:#B3E5FC;color:#01579B;font-weight:700" if row.iloc[i] else "background-color:#FFF9C4;color:#888"
            elif col in ("Horas A","Horas B","Horas C"):
                s[i]="background-color:#B3E5FC;color:#01579B" if row.iloc[i]>0 else "background-color:#F5F5F5;color:#AAA"
        return s

    st.dataframe(
        df.style.apply(style_row, axis=1)
          .format({"Ocup. A":"{:.0%}","Ocup. B":"{:.0%}","Ocup. C":"{:.0%}",
                   "Horas A":"{:.1f}","Horas B":"{:.1f}","Horas C":"{:.1f}"}),
        use_container_width=True, hide_index=True
    )

    # Suporte + totais
    sup = r["suporte"]
    sup_rows = []
    for nome, key in [("Lavadora e Inspeção","lavadora"),("Gravação e Estanqueidade","gravacao"),
                      ("Preset","preset"),("Coringa","coringa"),("Facilitador","facilitador")]:
        s = sup[key]
        sup_rows.append({
            "Função": nome,
            "Qtd A": s["A"], "Qtd B": s["B"], "Qtd C": s["C"],
            "Horas A": round(s["A"]*hA*dias,1),
            "Horas B": round(s["B"]*hB*dias,1),
            "Horas C": round(s["C"]*hC*dias,1),
        })
    sup_rows.append({
        "Função": "▶ TOTAL POR TURNO",
        "Qtd A": r["tot_A"], "Qtd B": r["tot_B"], "Qtd C": r["tot_C"],
        "Horas A": round(r["tot_A"]*hA*dias,1),
        "Horas B": round(r["tot_B"]*hB*dias,1),
        "Horas C": round(r["tot_C"]*hC*dias,1),
    })
    df_sup = pd.DataFrame(sup_rows)

    def style_sup(row):
        is_tot = "TOTAL" in str(row["Função"])
        bg = "background-color:#FF8A80;font-weight:700" if is_tot else ""
        return [bg]*len(row)

    st.dataframe(
        df_sup.style.apply(style_sup, axis=1).format(
            {"Horas A":"{:.1f}","Horas B":"{:.1f}","Horas C":"{:.1f}"}),
        use_container_width=True, hide_index=True
    )

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total funcionários", r["total"])
    c2.metric("Ciclo operacional",  f"{r['prod_ciclo_op']:.0%}")
    c3.metric("Labor operacional",  f"{r['prod_labor_op']:.0%}")
    c4.metric("⭐ Labor total",     f"{r['prod_labor_tot']:.0%}")

# ─────────────────────────────────────────
# GRÁFICO
# ─────────────────────────────────────────
def grafico_cenarios(cenarios_dict):
    cores_A   = ["#2E7D32","#66BB6A","#A5D6A7","#1B5E20"]
    cores_B   = ["#F9A825","#FFD54F","#FFE082","#FF6F00"]
    cores_C   = ["#1565C0","#64B5F6","#BBDEFB","#0D47A1"]
    cores_prod= ["#CC0000","#FF6D00","#7B1FA2","#00695C"]
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    for i,(nome,res) in enumerate(cenarios_dict.items()):
        mv,tA,tB,tC,tot,prod=[],[],[],[],[],[]
        for m,abr in zip(MESES,MESES_ABREV):
            r=res.get(m)
            if not r: continue
            mv.append(abr); tA.append(r["tot_A"]); tB.append(r["tot_B"])
            tC.append(r["tot_C"]); tot.append(r["total"])
            prod.append(r["prod_labor_tot"]*100)
        op=0.9 if i==0 else 0.6
        fig.add_trace(go.Bar(name=f"A—{nome}",x=mv,y=tA,marker_color=cores_A[i%4],opacity=op,
            offsetgroup=i,legendgroup=nome,text=tA,textposition="inside",
            textfont=dict(color="white",size=9)),secondary_y=False)
        fig.add_trace(go.Bar(name=f"B—{nome}",x=mv,y=tB,marker_color=cores_B[i%4],opacity=op,
            offsetgroup=i,legendgroup=nome,base=tA,text=tB,textposition="inside",
            textfont=dict(size=9)),secondary_y=False)
        fig.add_trace(go.Bar(name=f"C—{nome}",x=mv,y=tC,marker_color=cores_C[i%4],opacity=op,
            offsetgroup=i,legendgroup=nome,base=[a+b for a,b in zip(tA,tB)],
            text=tC,textposition="inside",textfont=dict(color="white",size=9)),secondary_y=False)
        fig.add_trace(go.Scatter(
            name=f"Labor—{nome}",x=mv,y=prod,mode="lines+markers+text",
            marker=dict(color=cores_prod[i%4],size=10,symbol="circle" if i==0 else "diamond"),
            line=dict(color=cores_prod[i%4],width=2,dash="solid" if i==0 else "dot"),
            text=[f"{p:.0f}%" for p in prod],textposition="top center",
            textfont=dict(color=cores_prod[i%4],size=11)),secondary_y=True)
    fig.update_layout(
        barmode="stack",
        title=dict(text="MÃO-DE-OBRA POR TURNO",font=dict(size=15,color="#1E3A5F")),
        yaxis_title="Nº Funcionários",
        yaxis2=dict(title="Labor Total (%)",tickformat=".0f",ticksuffix="%",range=[0,100]),
        legend=dict(orientation="h",y=-0.3,font=dict(size=10)),
        height=460,plot_bgcolor="white",
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True,gridcolor="#E0E7EF"),
    )
    return fig

# ─────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────
def exportar(resultados):
    output = BytesIO()
    wb = openpyxl.Workbook()
    brd = Border(*[Side(style='thin',color='CCCCCC')]*4)

    def ec(cell,bg="FFFFFF",fg="000000",bold=False,fmt=None,center=True):
        cell.font=Font(name="Arial",bold=bold,color=fg,size=9)
        cell.fill=PatternFill("solid",fgColor=bg)
        cell.alignment=Alignment(horizontal="center" if center else "left",vertical="center")
        cell.border=brd
        if fmt: cell.number_format=fmt

    ws=wb.active; ws.title="RESUMO MO"
    for i,h in enumerate(["Mês","Dias","Turno A","Turno B","Turno C","Total",
                           "Ciclo Op.","Ciclo Total","Labor Op.","Labor Total ★"],1):
        ec(ws.cell(1,i,h),"1E3A5F","FFFFFF",True)

    for ri,(m,abr) in enumerate(zip(MESES,MESES_ABREV),2):
        r=resultados.get(m)
        bg="EAF3FB" if ri%2==0 else "FFFFFF"
        vals=[abr,0,"-","-","-","-","-","-","-","-"] if not r else [
            abr,r["dias"],r["tot_A"],r["tot_B"],r["tot_C"],r["total"],
            r["prod_ciclo_op"],r["prod_ciclo_tot"],r["prod_labor_op"],r["prod_labor_tot"]]
        for ci,v in enumerate(vals,1):
            c=ws.cell(ri,ci,v)
            fmt="0%" if ci>=7 and isinstance(v,float) else None
            ec(c,"FFF9C4" if ci==10 and isinstance(v,float) else bg,fmt=fmt)
        ws.row_dimensions[ri].height=15

    for mes in MESES:
        r=resultados.get(mes)
        if not r: continue
        wsm=wb.create_sheet(mes[:10])
        hA,hB,hC,dias=r["hA"],r["hB"],r["hC"],r["dias"]
        for ci,txt in [(1,""),(2,"TURNO A"),(3,"TURNO B"),(4,"TURNO C"),
                       (5,"TURNO A"),(6,"TURNO B"),(7,"TURNO C"),
                       (8,"TURNO A"),(9,"TURNO B"),(10,"TURNO C")]:
            ec(wsm.cell(1,ci,txt),"1E3A5F","FFFFFF",True)
        for ci,txt in [(2,f"{hA}h"),(3,f"{hB}h"),(4,f"{hC}h")]:
            ec(wsm.cell(2,ci,txt),"2D6A9F","FFFFFF",True)
        wsm.cell(1,1,mes.upper()); ec(wsm.cell(1,1),"1E3A5F","FFFFFF",True)
        def cbg(v):
            if v>1.0: return "FFCDD2"
            if v>=0.85: return "FFF9C4"
            return "C8E6C9"
        ri=3
        for _,row in r["centros"].iterrows():
            data=[(row.centro,"FFFFFF",False),
                  (f"{row.ocup_A:.0%}",cbg(row.ocup_A),True),(f"{row.ocup_B:.0%}",cbg(row.ocup_B),True),(f"{row.ocup_C:.0%}",cbg(row.ocup_C),True),
                  (row.ativo_A,"B3E5FC" if row.ativo_A else "FFF9C4",True),(row.ativo_B,"B3E5FC" if row.ativo_B else "FFF9C4",True),(row.ativo_C,"B3E5FC" if row.ativo_C else "FFF9C4",True),
                  (f"{row.horas_disp_A:.2f}" if row.ativo_A else "0","B3E5FC" if row.ativo_A else "F5F5F5",True),
                  (f"{row.horas_disp_B:.2f}" if row.ativo_B else "0","B3E5FC" if row.ativo_B else "F5F5F5",True),
                  (f"{row.horas_disp_C:.2f}" if row.ativo_C else "0","B3E5FC" if row.ativo_C else "F5F5F5",True)]
            for ci,(val,bg,ctr) in enumerate(data,1):
                ec(wsm.cell(ri,ci,val),bg,center=ctr)
            ri+=1
        sup=r["suporte"]
        for nome,key in [("TOTAL DE OPERADORES",None),("LAVADORA E INSPEÇÃO","lavadora"),
                         ("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),("PRESET","preset"),
                         ("CORINGA","coringa"),("FACILITADOR","facilitador"),
                         ("TOTAL POR TURNO",None),("TOTAL FUNCIONÁRIOS",None)]:
            bold="TOTAL" in nome; bg_r="FF8A80" if bold else "FFFFFF"
            ec(wsm.cell(ri,1,nome),bg_r,bold=bold,center=False)
            if key:
                s=sup[key]
                for ci,t in [(5,"A"),(6,"B"),(7,"C")]:
                    ec(wsm.cell(ri,ci,s[t]),"B3E5FC" if s[t] else "FFF9C4",bold=bold)
                for ci,t,h in [(8,"A",hA),(9,"B",hB),(10,"C",hC)]:
                    v=s[t]*h*dias
                    ec(wsm.cell(ri,ci,f"{v:.2f}" if v else "0"),"B3E5FC" if v else "F5F5F5",bold=bold)
            elif "TOTAL DE OPERADORES" in nome:
                for ci,v in [(5,r["op_A"]),(6,r["op_B"]),(7,r["op_C"])]:
                    ec(wsm.cell(ri,ci,v),"FF8A80",bold=True)
                for ci,v,h in [(8,r["op_A"],hA),(9,r["op_B"],hB),(10,r["op_C"],hC)]:
                    ec(wsm.cell(ri,ci,f"{v*h*dias:.2f}"),"FF8A80",bold=True)
            elif "TOTAL POR TURNO" in nome:
                for ci,v in [(5,r["tot_A"]),(6,r["tot_B"]),(7,r["tot_C"])]:
                    ec(wsm.cell(ri,ci,v),"FF8A80",bold=True)
                for ci,v,h in [(8,r["tot_A"],hA),(9,r["tot_B"],hB),(10,r["tot_C"],hC)]:
                    ec(wsm.cell(ri,ci,f"{v*h*dias:.2f}"),"FF8A80",bold=True)
            elif "FUNCIONÁRIOS" in nome:
                ec(wsm.cell(ri,4,r["total"]),"FF8A80",bold=True)
                tot_h=r["tot_A"]*hA*dias+r["tot_B"]*hB*dias+r["tot_C"]*hC*dias
                ec(wsm.cell(ri,8,f"{tot_h:.2f}"),"FF8A80",bold=True)
            ri+=1
        ri+=1
        wsm.merge_cells(f"H{ri}:I{ri}")
        for nm,v,dest in [("PROD. CICLO OPERACIONAL",r["prod_ciclo_op"],False),
                          ("PROD. CICLO TOTAL",r["prod_ciclo_tot"],False),
                          ("PROD. LABOR OPERACIONAL",r["prod_labor_op"],False),
                          ("PROD. LABOR TOTAL ★",r["prod_labor_tot"],True)]:
            bg="FFF9C4" if dest else "FFFFFF"; fg="E65100" if dest else "000000"
            wsm.merge_cells(f"H{ri}:I{ri}")
            ec(wsm.cell(ri,8,nm),bg,fg,dest,center=False)
            ec(wsm.cell(ri,10,v),bg,fg,dest,"0%")
            ri+=1
        for ci,w in enumerate([14,8,8,8,8,8,8,24,10,10],1):
            wsm.column_dimensions[get_column_letter(ci)].width=w
    wb.save(output); output.seek(0)
    return output

# ─────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🏭 Calculadora de Recursos — Usinagem</h1>
  <p>Upload da planilha de inputs → cálculo automático de headcount, ocupação e produtividade</p>
</div>
""", unsafe_allow_html=True)

with st.expander("📋 Como preparar o arquivo para upload", expanded=False):
    st.markdown("""
**O app lê 5 abas do seu `.xlsm` ou `.xlsx`. Abas mensais (NovFY26 etc.) não são necessárias.**

| Aba | Conteúdo |
|---|---|
| `INPUT_PMP` | Demanda mensal por modelo + dias trabalhados |
| `IMPUTTEMPO` | Tempo de ciclo e labor por centro/peça |
| `IMPUTDISTRIBUIÇÃO` | Divisão de carga, volume e disponibilidade |
| `IMPUTAPLICAÇÃO` | Matriz 0/1 — quais modelos passam em cada centro |
| `IMPUTTURNOS` | Horas acumuladas por turno (referência) |

> As horas de **duração** de cada turno e os thresholds de ativação são configuráveis na barra lateral.
    """)

uploaded = st.file_uploader("Selecione o arquivo Excel (.xlsm ou .xlsx)", type=["xlsm","xlsx"])
if not uploaded:
    st.info("👆 Faça upload do arquivo para começar.")
    st.stop()

file_bytes = uploaded.read()
with st.spinner("Lendo planilha..."):
    try:
        pmp, dias = read_pmp(file_bytes)
        tempo      = read_tempo(file_bytes)
        dist       = read_dist(file_bytes)
        aplic      = read_aplic(file_bytes)
    except Exception as e:
        st.error(f"Erro ao ler: {e}"); st.stop()

st.success(f"✅ {len(aplic)} combinações centro/peça/modelo · {pmp.modelo.nunique()} modelos · {pmp.mes.nunique()} meses")

erros, alertas, oks = validar(pmp, tempo, dist, aplic, dias)
label_exp = ("🔴 "+str(len(erros))+" erro(s) crítico(s)  " if erros else "") + \
            ("⚠️ "+str(len(alertas))+" aviso(s)" if alertas else "") + \
            ("✅ Inputs sem inconsistências" if not erros and not alertas else "")
with st.expander(label_exp, expanded=bool(erros)):
    for e in erros:
        st.markdown(f'<div class="aviso-erro">🔴 <b>ERRO CRÍTICO:</b> {e}</div>', unsafe_allow_html=True)
    for a in alertas:
        st.markdown(f'<div class="aviso-alerta">⚠️ {a}</div>', unsafe_allow_html=True)
    for o in oks:
        st.markdown(f'<div class="aviso-ok">✅ {o}</div>', unsafe_allow_html=True)
if erros:
    st.error("Corrija os erros acima antes de continuar."); st.stop()

# ─────────────────────────────────────────
# SIDEBAR — configurações
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configurações")
    st.markdown("**Duração dos turnos (horas)**")
    hA = st.number_input("Turno A (h)", value=8.80, step=0.01, format="%.2f")
    hB = st.number_input("Turno B (h)", value=8.23, step=0.01, format="%.2f")
    hC = st.number_input("Turno C (h)", value=7.68, step=0.01, format="%.2f")
    horas_turno = {"A": hA, "B": hB, "C": hC}

    st.markdown("---")
    st.markdown("**Thresholds de ativação (%)**")
    st.caption("Turno abre quando a ocupação ultrapassa esse valor.")
    thr_A = st.number_input("Turno A abre quando ocup. A >", value=40, min_value=0, max_value=200, step=1)
    thr_B = st.number_input("Turno B abre quando ocup. A >", value=106, min_value=0, max_value=200, step=1)
    thr_C = st.number_input("Turno C abre quando ocup. B >", value=100, min_value=0, max_value=200, step=1)
    thresholds = {"A": thr_A, "B": thr_B, "C": thr_C}

    st.markdown("---")
    st.markdown("**Funções de suporte**")
    st.caption("Modo automático usa as regras padrão. Manual permite definir o número exato por turno.")

    suporte_cfg = {}
    for key, label, defaults in [
        ("lavadora",   "Lavadora e Inspeção",         {"A":1,"B":1,"C":0}),
        ("gravacao",   "Gravação e Estanqueidade",     {"A":1,"B":1,"C":0}),
        ("preset",     "Preset",                       {"A":2,"B":1,"C":1}),
        ("coringa",    "Coringa",                      {"A":1,"B":0,"C":0}),
        ("facilitador","Facilitador",                  {"A":1,"B":1,"C":0}),
    ]:
        with st.expander(f"🔧 {label}"):
            modo = st.radio("Modo", ["Automático","Manual"], key=f"modo_{key}", horizontal=True)
            if modo == "Automático":
                st.caption(f"Padrão: A={defaults['A']} · B={defaults['B']} · C={defaults['C']}")
                suporte_cfg[key] = {"modo":"auto","A":defaults["A"],"B":defaults["B"],"C":defaults["C"]}
            else:
                c1,c2,c3 = st.columns(3)
                vA = c1.number_input("A", 0, 10, defaults["A"], key=f"sup_{key}_A")
                vB = c2.number_input("B", 0, 10, defaults["B"], key=f"sup_{key}_B")
                vC = c3.number_input("C", 0, 10, defaults["C"], key=f"sup_{key}_C")
                suporte_cfg[key] = {"modo":"manual","A":vA,"B":vB,"C":vC}

    st.markdown("---")
    st.caption("Todas as alterações afetam os cálculos em tempo real.")

# ─────────────────────────────────────────
# TABS
# ─────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📊 Resultados", "🔧 Simulador de Cenários", "🔍 Detalhes por etapa", "📥 Exportar"])

# ══════════════════════════════════════════
# TAB 1
# ══════════════════════════════════════════
with tab1:
    res_base = calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg)
    st.plotly_chart(grafico_cenarios({"Base": res_base}), use_container_width=True)

    st.markdown('<div class="section-title">Detalhe por mês</div>', unsafe_allow_html=True)
    mes_sel = st.selectbox("Mês", [m for m in MESES if res_base.get(m)], key="mes_r")
    if mes_sel and res_base.get(mes_sel):
        show_tabela(res_base[mes_sel])

# ══════════════════════════════════════════
# TAB 2 — CENÁRIOS
# ══════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Simulador de cenários</div>', unsafe_allow_html=True)
    st.caption("Crie cenários ajustando operadores por centro e compare todos lado a lado.")

    if "cenarios" not in st.session_state:
        st.session_state.cenarios = {}

    # Criar cenário
    with st.expander("➕ Criar novo cenário", expanded=len(st.session_state.cenarios)==0):
        col_a, col_b = st.columns([2,1])
        novo_nome = col_a.text_input("Nome do cenário", placeholder="Ex: Redução turno B, Novembro otimizado...")
        mes_novo  = col_b.selectbox("Mês base", MESES, key="mes_novo")

        if novo_nome and mes_novo:
            res_b  = calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg)
            r_orig = res_b.get(mes_novo)
            if r_orig:
                centros_list = sorted(r_orig["centros"].centro.tolist())
                st.markdown(f"**Ajuste operadores para '{novo_nome}' em {mes_novo}:**")
                st.caption("🟢 <85% · 🟡 85–100% · 🔴 >100%  |  Número = operadores naquele turno")

                cols_h = st.columns([3,1,1,1])
                cols_h[0].markdown("**Centro — ocup. A / B / C**")
                cols_h[1].markdown("**Turno A**")
                cols_h[2].markdown("**Turno B**")
                cols_h[3].markdown("**Turno C**")

                novo_ov = {}
                for cen in centros_list:
                    rc = r_orig["centros"][r_orig["centros"].centro==cen].iloc[0]
                    eA = "🔴" if rc.ocup_A>1 else ("🟡" if rc.ocup_A>=0.85 else "🟢")
                    eB = "🔴" if rc.ocup_B>1 else ("🟡" if rc.ocup_B>=0.85 else "🟢")
                    eC = "🔴" if rc.ocup_C>1 else ("🟡" if rc.ocup_C>=0.85 else "🟢")
                    c0,c1,c2,c3 = st.columns([3,1,1,1])
                    c0.markdown(f"`{cen}` {eA}{rc.ocup_A:.0%} / {eB}{rc.ocup_B:.0%} / {eC}{rc.ocup_C:.0%}")
                    vA = c1.number_input("",0,5,int(rc.ativo_A),key=f"n_{novo_nome}_{cen}_A",label_visibility="collapsed",help=f"Calculado: {rc.ativo_A}")
                    vB = c2.number_input("",0,5,int(rc.ativo_B),key=f"n_{novo_nome}_{cen}_B",label_visibility="collapsed",help=f"Calculado: {rc.ativo_B}")
                    vC = c3.number_input("",0,5,int(rc.ativo_C),key=f"n_{novo_nome}_{cen}_C",label_visibility="collapsed",help=f"Calculado: {rc.ativo_C}")
                    novo_ov[cen] = {"A":vA,"B":vB,"C":vC}

                if st.button("💾 Salvar cenário", type="primary"):
                    ov_completo = {mes_novo: novo_ov}
                    res_cen = calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg, ov_completo)
                    st.session_state.cenarios[novo_nome] = {
                        "resultados": res_cen, "mes": mes_novo, "overrides": ov_completo
                    }
                    st.success(f"✅ Cenário '{novo_nome}' salvo!")
                    st.rerun()

    # Cenários salvos
    if st.session_state.cenarios:
        res_base_cmp = calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg)
        todos = {"📌 Base": res_base_cmp}
        todos.update({k: v["resultados"] for k,v in st.session_state.cenarios.items()})

        st.markdown('<div class="section-title">Comparativo de cenários</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_cenarios(todos), use_container_width=True)

        # Cards resumo
        cols_c = st.columns(min(len(todos), 4))
        for i,(nome,res) in enumerate(todos.items()):
            meses_ok = [m for m in MESES if res.get(m)]
            media_labor = np.mean([res[m]["prod_labor_tot"] for m in meses_ok]) if meses_ok else 0
            total_anual = sum(res[m]["total"] for m in meses_ok)
            with cols_c[i % len(cols_c)]:
                st.markdown(f"""
<div class="cenario-card {'ativo' if i==0 else ''}">
  <b style="font-size:13px">{nome}</b><br>
  <span style="font-size:12px;color:#555">
    Labor Total médio: <b>{media_labor:.0%}</b><br>
    Func. total anual: <b>{total_anual}</b>
  </span>
</div>
                """, unsafe_allow_html=True)

        # Tabela comparativa
        st.markdown('<div class="section-title">Tabela comparativa por mês</div>', unsafe_allow_html=True)
        mes_cmp = st.selectbox("Mês", MESES, key="mes_cmp")
        cmp_rows = []
        for nome,res in todos.items():
            r=res.get(mes_cmp)
            if r:
                cmp_rows.append({
                    "Cenário":nome,"Turno A":r["tot_A"],"Turno B":r["tot_B"],"Turno C":r["tot_C"],
                    "Total":r["total"],"Labor Total ★":f"{r['prod_labor_tot']:.0%}",
                    "Labor Op.":f"{r['prod_labor_op']:.0%}","Ciclo Total":f"{r['prod_ciclo_tot']:.0%}",
                })
        if cmp_rows:
            st.dataframe(pd.DataFrame(cmp_rows), use_container_width=True, hide_index=True)

        # Detalhe individual
        st.markdown('<div class="section-title">Tabela detalhada por cenário</div>', unsafe_allow_html=True)
        sel_cen = st.selectbox("Cenário", list(todos.keys()), key="sel_cen")
        sel_mes = st.selectbox("Mês", [m for m in MESES if todos[sel_cen].get(m)], key="sel_mes_cen")
        if sel_mes and todos[sel_cen].get(sel_mes):
            show_tabela(todos[sel_cen][sel_mes])

        # Ações
        st.markdown("---")
        col_del, col_exp_base, col_exp_cen = st.columns(3)
        with col_del:
            del_nome = st.selectbox("Remover cenário", list(st.session_state.cenarios.keys()), key="del_cen")
            if st.button("🗑️ Remover", type="secondary"):
                del st.session_state.cenarios[del_nome]; st.rerun()
        with col_exp_base:
            st.download_button("📥 Exportar Base",
                data=exportar(res_base_cmp),
                file_name="resultado_base.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_exp_cen:
            for nome,v in st.session_state.cenarios.items():
                st.download_button(f"📥 Exportar — {nome}",
                    data=exportar(v["resultados"]),
                    file_name=f"cenario_{nome.replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{nome}")
    else:
        st.info("Nenhum cenário criado ainda. Use o formulário acima.")

# ══════════════════════════════════════════
# TAB 3 — DETALHES
# ══════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">Detalhes por etapa do cálculo</div>', unsafe_allow_html=True)
    etapa = st.radio("Etapa",[
        "Passo 2 — INPUT_PMP",
        "Passo 3 — IMPUTAPLICAÇÃO",
        "Passo 4 — JOIN aplicacao × pmp",
        "Passo 7 — Minutos por linha",
        "Passo 8 — Totais por centro",
        "Passo 10 — % ocupação",
    ], horizontal=True)
    mes_det = st.selectbox("Mês", MESES, key="mes_det")
    res_det = calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg)

    if etapa=="Passo 2 — INPUT_PMP":
        st.dataframe(pmp[pmp.mes==mes_det].reset_index(drop=True),use_container_width=True,hide_index=True)
    elif etapa=="Passo 3 — IMPUTAPLICAÇÃO":
        st.dataframe(aplic.head(300),use_container_width=True,hide_index=True)
        st.caption(f"Total: {len(aplic)} combinações ativas")
    elif etapa=="Passo 4 — JOIN aplicacao × pmp":
        p4=aplic.merge(pmp[pmp.mes==mes_det],on="modelo")
        st.dataframe(p4.head(300),use_container_width=True,hide_index=True)
        st.caption(f"{len(p4)} linhas")
    elif etapa=="Passo 7 — Minutos por linha":
        p7=(aplic.merge(pmp[pmp.mes==mes_det],on="modelo")
               .merge(tempo,on=["centro","peca"]).merge(dist,on=["centro","peca"]))
        p7["indice_ciclo"]=(p7.t_ciclo*p7.div_carga*p7.div_volume)/p7.disponib
        p7["min_ciclo"]=(p7.indice_ciclo*p7.qtd).round(1)
        p7["min_labor"]=(p7.t_labor*p7.div_carga*p7.qtd).round(1)
        st.dataframe(p7.head(300),use_container_width=True,hide_index=True)
        st.caption(f"{len(p7)} linhas")
    elif etapa=="Passo 8 — Totais por centro":
        p7=(aplic.merge(pmp[pmp.mes==mes_det],on="modelo")
               .merge(tempo,on=["centro","peca"]).merge(dist,on=["centro","peca"]))
        p7["min_ciclo"]=(p7.t_ciclo*p7.div_carga*p7.div_volume/p7.disponib)*p7.qtd
        p7["min_labor"]=p7.t_labor*p7.div_carga*p7.qtd
        p8=p7.groupby("centro")[["min_ciclo","min_labor"]].sum().reset_index()
        p8["horas_ciclo"]=(p8.min_ciclo/60).round(1)
        p8["horas_labor"]=(p8.min_labor/60).round(1)
        st.dataframe(p8,use_container_width=True,hide_index=True)
    elif etapa=="Passo 10 — % ocupação":
        r=res_det.get(mes_det)
        if r:
            df_oc=r["centros"][["centro","ocup_A","ocup_B","ocup_C","ativo_A","ativo_B","ativo_C"]].copy()
            def fmt(v):
                e="🔴" if v>1 else ("🟡" if v>=0.85 else "🟢")
                return f"{e} {v:.0%}"
            df_oc["ocup_A"]=df_oc.ocup_A.map(fmt)
            df_oc["ocup_B"]=df_oc.ocup_B.map(fmt)
            df_oc["ocup_C"]=df_oc.ocup_C.map(fmt)
            st.dataframe(df_oc,use_container_width=True,hide_index=True)
            st.caption(f"Thresholds: A>{thr_A}% · B quando A>{thr_B}% · C quando B>{thr_C}%")

# ══════════════════════════════════════════
# TAB 4 — EXPORTAR
# ══════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">Exportar resultados</div>', unsafe_allow_html=True)
    st.markdown("O Excel exportado contém a aba **RESUMO MO** + uma aba por mês com a tabela completa no formato original.")
    res_exp = calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg)
    st.download_button(
        "📥 Baixar Excel — Resultado Base",
        data=exportar(res_exp),
        file_name="resultado_usinagem.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
