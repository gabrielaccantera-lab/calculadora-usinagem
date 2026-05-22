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
.kpi-card{background:linear-gradient(135deg,#1A2E1A 0%,#0D1F0D 100%);border:1px solid #2A4A2A;border-radius:12px;padding:16px 18px;margin:4px 0;position:relative;overflow:hidden;}
.kpi-card::before{content:'';position:absolute;top:0;left:0;width:4px;height:100%;background:#FFDE00;}
.kpi-card .kpi-icon{font-size:22px;margin-bottom:4px;}
.kpi-card .kpi-label{font-size:10px;color:#7BC67A;text-transform:uppercase;letter-spacing:.06em;font-weight:600;}
.kpi-card .kpi-value{font-size:26px;font-weight:800;color:#FFFFFF;line-height:1.1;}
.kpi-card .kpi-sub{font-size:11px;color:#AAAAAA;margin-top:2px;}
.kpi-card.destaque{border-color:#FFDE00;}
.kpi-card.destaque::before{background:#367C2B;width:100%;height:3px;top:0;left:0;}
.mes-row{display:flex;align-items:center;gap:10px;padding:7px 12px;border-radius:8px;margin:3px 0;background:#131313;border:1px solid #222;}
.mes-nome{font-size:12px;font-weight:700;color:#FFDE00;min-width:36px;}
.mes-bar{flex:1;height:8px;background:#1A2E1A;border-radius:4px;overflow:hidden;}
.mes-bar-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,#367C2B,#FFDE00);}
.mes-num{font-size:12px;color:#FFFFFF;font-weight:700;min-width:28px;text-align:right;}
.mes-labor{font-size:11px;min-width:44px;text-align:right;}
.tp{display:inline-block;border-radius:20px;padding:1px 8px;font-size:10px;font-weight:700;margin:0 1px;}
.tpA{background:#1F4D19;color:#92D050;} .tpB{background:#3D2D00;color:#FFDE00;} .tpC{background:#0D2040;color:#00B0F0;}
.gauge-wrap{text-align:center;padding:6px 0;}
</style>
""", unsafe_allow_html=True)

MESES       = ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
               "Maio","Junho","Julho","Agosto","Setembro","Outubro"]
MESES_ABREV = ["NOV","DEZ","JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT"]

COL_MAP = {
    "IMPUTTEMPO": {
        "centro":  ["Máquina","MÁQUINA","maquina","Centro","CENTRO"],
        "peca":    ["PEÇA","Peça","peca","PECA","REFERÊNCIA","Referência","referencia","REFERENCIA","REF"],
        "t_ciclo": ["Tempo\nCiclo\n(min)","Tempo Ciclo (min)","T.CICLO","t_ciclo","CICLO"],
        "t_labor": ["Tempo\nLabor\n(min)","Tempo Labor (min)","T.LABOR","t_labor","LABOR"],
    },
    "IMPUTDISTRIBUIÇÃO": {
        "centro":     ["Máquina","MÁQUINA","maquina","Centro","CENTRO"],
        "peca":       ["PEÇA","Peça","peca","PECA","REFERÊNCIA","Referência","referencia","REFERENCIA","REF"],
        "div_carga":  ["Divisão\nCarga\nENTRE\nMÁQUINAS","Div Carga","DIV_CARGA","div_carga"],
        "vol_int":    ["Vol.\nInterna","Vol. Interna","VOL_INT","vol_int","VOL. INTERNA","Volume Interna","Volume de \nProdução\nInterna","Volume de\nProdução\nInterna","Volume Produção Interna","Vol Produção Interna"],
        "div_volume": ["Divisão \nde\nVolume\nENTRE\nPEÇAS","Divisão\nde\nVolume\nENTRE\nPEÇAS","Div Volume","DIV_VOLUME","div_volume","Divisão de Volume"],
        "disponib":   ["Disponi-\nbilidade","Disponibilidade","DISPONIB","disponib"],
        "perf_op":    ["Performance\nOperador X\nMáquina","Performance Operador X Máquina",
                       "Performance\nOperador X\nMaquina","Performance Operador X Maquina",
                       "Performance\nOperador\nMáquina","Performance Operador Máquina",
                       "Performance Operador","PERF_OP","perf_op","PERFORMANCE"],
    },
}

ABA_FORMATOS = {
    "INPUT_PMP": "**INPUT_PMP** — Linha 1: dias trabalhados (colunas B→M = Nov→Out). Linhas 3+: modelos, colunas B→M = qtd peças.",
    "IMPUTTEMPO": "**IMPUTTEMPO** — Cabeçalho linha 1. Colunas: `Máquina`, `REFERÊNCIA` (ou `PEÇA`), `Tempo Ciclo (min)`, `Tempo Labor (min)`.",
    "IMPUTDISTRIBUIÇÃO": "**IMPUTDISTRIBUIÇÃO** — Cabeçalho linha 1. Colunas: `Máquina`, `REFERÊNCIA` (ou `PEÇA`), `Divisão Carga`, `Vol. Interna`, `Divisão de Volume`, `Disponibilidade`, `Performance Operador X Máquina`.",
    "IMPUTAPLICAÇÃO": "**IMPUTAPLICAÇÃO** — Cabeçalho linha 1. Col A=Centro, Col B=REFERÊNCIA (ou PEÇA), depois colunas por modelo (qualquer nome).",
    "IMPUTTURNOS": "**IMPUTTURNOS** — Linha 1: horas acumuladas. B1=Turno A, C1=Turno B, D1=Turno C.",
}

def _norm(s):
    return str(s).lower().replace("\n"," ").replace("\r"," ").strip()

def find_col(df, candidates, aba, campo):
    # 1) busca exata
    for c in candidates:
        if c in df.columns:
            return c
    # 2) busca normalizada
    norm_map = {_norm(col): col for col in df.columns}
    for c in candidates:
        nc = _norm(c)
        if nc in norm_map:
            st.session_state.setdefault("log_leitura",[]).append(
                f"ℹ️ [{aba}] Campo '{campo}' encontrado via normalização: '{norm_map[nc]}'")
            return norm_map[nc]
    # 3) fallback por índice posicional
    idx_fallback = {
        "centro":0,"peca":1,"t_ciclo":5,"t_labor":6,
        "div_carga":7,"vol_int":8,"div_volume":9,"disponib":10,"perf_op":11
    }
    if campo in idx_fallback:
        idx = idx_fallback[campo]
        if idx < len(df.columns):
            st.session_state.setdefault("log_leitura",[]).append(
                f"⚠️ [{aba}] Campo '{campo}' não encontrado pelo nome — usando coluna {idx+1} ({df.columns[idx]}) como fallback")
            return df.columns[idx]
    # 4) Erro detalhado com sugestão
    nomes_esperados = " ou ".join(f'"{c}"' for c in candidates[:3])
    colunas_disponiveis = ", ".join(f'"{c}"' for c in df.columns)
    raise ValueError(
        f"\n🔴 [{aba}] Campo obrigatório '{campo}' não encontrado!\n\n"
        f"   O app esperava uma coluna chamada {nomes_esperados} (entre outras variações).\n\n"
        f"   Colunas encontradas na aba '{aba}':\n"
        f"   {colunas_disponiveis}\n\n"
        f"   👉 Renomeie a coluna correspondente para um dos nomes esperados:\n"
        f"   {', '.join(f'{c}' for c in candidates)}"
    )

def verificar_prefixo_centro(df, aba, log):
    """Avisa se algum centro não começa com 'CEN', mas deixa passar."""
    if "centro" not in df.columns:
        return
    centros_fora = df[~df["centro"].astype(str).str.upper().str.startswith("CEN")]["centro"].dropna().unique()
    if len(centros_fora) > 0:
        exemplos = ", ".join(f'"{c}"' for c in centros_fora[:5])
        log.append(
            f"⚠️ [{aba}] {len(centros_fora)} centro(s) não iniciam com 'CEN' — "
            f"verifique se estão corretos: {exemplos}"
            f"{'...' if len(centros_fora) > 5 else ''}"
        )

def find_aba(sheetnames, candidatos):
    """Busca aba por nome exato, depois por similaridade (ignora underline, espaço, maiúsculas, acento)."""
    import unicodedata
    def _norm_aba(s):
        s = str(s).upper().strip()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        for ch in [' ', '_', '-', '.']: s = s.replace(ch, '')
        return s
    # 1) busca exata
    for c in candidatos:
        if c in sheetnames:
            return c
    # 2) busca normalizada
    norm_map = {_norm_aba(a): a for a in sheetnames}
    for c in candidatos:
        nc = _norm_aba(c)
        if nc in norm_map:
            return norm_map[nc]
    return None

def verificar_abas(fb):
    try:
        wb = openpyxl.load_workbook(BytesIO(fb), read_only=True, data_only=True)
        sheetnames = wb.sheetnames; wb.close()
    except: sheetnames = []

    CANDIDATOS = {
        "INPUT_PMP":         ["INPUT_PMP","INPUTPMP","INPUT PMP","PMP"],
        "IMPUTTEMPO":        ["IMPUTTEMPO","INPUT_TEMPO","INPUTTEMPO","INPUT TEMPO","TEMPO"],
        "IMPUTDISTRIBUIÇÃO": ["IMPUTDISTRIBUIÇÃO","IMPUTDISTRIBUICAO","INPUT_DISTRIBUIÇÃO",
                              "INPUTDISTRIBUIÇÃO","INPUTDISTRIBUICAO","INPUT DISTRIBUIÇÃO",
                              "INPUT_DISTRIBUICAO","DISTRIBUIÇÃO","DISTRIBUICAO"],
        "IMPUTAPLICAÇÃO":    ["IMPUTAPLICAÇÃO","IMPUTAPLICACAO","INPUT_APLICAÇÃO",
                              "INPUTAPLICAÇÃO","INPUTAPLICACAO","INPUT APLICAÇÃO",
                              "INPUT_APLICACAO","APLICAÇÃO","APLICACAO"],
        "IMPUTTURNOS":       ["IMPUTTURNOS","INPUT_TURNOS","INPUTTURNOS","INPUT TURNOS","TURNOS"],
    }

    resultado = {}
    for aba_padrao, candidatos in CANDIDATOS.items():
        encontrada = find_aba(sheetnames, candidatos)
        resultado[aba_padrao] = encontrada  # None se não encontrou, nome real se encontrou
    return resultado

def _get_aba(chave):
    """Retorna o nome real da aba encontrada no arquivo, ou o nome padrão como fallback."""
    return st.session_state.get("_abas_map", {}).get(chave) or chave

def read_pmp(fb, log):
    aba = _get_aba("INPUT_PMP")
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name=aba, header=None)
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{aba}': {e}\n\n{ABA_FORMATOS['INPUT_PMP']}")
    log.append(f"✅ INPUT_PMP lido (aba: '{aba}'): {df.shape[0]}L × {df.shape[1]}C")
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
    aba = _get_aba("IMPUTTURNOS")
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name=aba, header=None)
        hA = float(df.iloc[0,1]) if pd.notna(df.iloc[0,1]) else 7.5
        hB = float(df.iloc[0,2]) if pd.notna(df.iloc[0,2]) else 14.25
        hC = float(df.iloc[0,3]) if pd.notna(df.iloc[0,3]) else 19.5
        return {"A": hA, "B": hB, "C": hC}, True
    except:
        return {"A": 7.5, "B": 14.25, "C": 19.5}, False

def read_tempo(fb, log):
    aba = _get_aba("IMPUTTEMPO")
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name=aba, header=0)
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{aba}': {e}\n\n{ABA_FORMATOS['IMPUTTEMPO']}")
    log.append(f"✅ IMPUTTEMPO lido (aba: '{aba}'): {df.shape[0]}L")
    mp = COL_MAP["IMPUTTEMPO"]
    c = {k: find_col(df, v, aba, k) for k,v in mp.items()}
    out = df[[c["centro"],c["peca"],c["t_ciclo"],c["t_labor"]]].copy()
    out.columns = ["centro","peca","t_ciclo","t_labor"]
    out = out.dropna(subset=["centro"])
    verificar_prefixo_centro(out, aba, log)
    log.append(f"   {len(out)} combinações centro+peça")
    return out.copy()

def read_dist(fb, log):
    aba = _get_aba("IMPUTDISTRIBUIÇÃO")
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name=aba, header=0)
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{aba}': {e}\n\n{ABA_FORMATOS['IMPUTDISTRIBUIÇÃO']}")
    log.append(f"✅ IMPUTDISTRIBUIÇÃO lido (aba: '{aba}'): {df.shape[0]}L")
    log.append(f"   Colunas brutas: {list(df.columns)}")
    mp = COL_MAP["IMPUTDISTRIBUIÇÃO"]
    c = {k: find_col(df, v, aba, k) for k,v in mp.items()}
    log.append(f"   Mapeamento: { {k: v for k,v in c.items()} }")
    out = df[[c["centro"],c["peca"],c["div_carga"],c["vol_int"],c["div_volume"],c["disponib"],c["perf_op"]]].copy()
    out.columns = ["centro","peca","div_carga","vol_int","div_volume","disponib","perf_op"]
    out["vol_int"]    = pd.to_numeric(out["vol_int"],    errors="coerce").fillna(1.0)
    out["div_carga"]  = pd.to_numeric(out["div_carga"],  errors="coerce").fillna(0.0)
    out["div_volume"] = pd.to_numeric(out["div_volume"], errors="coerce").fillna(0.0)
    out["disponib"]   = pd.to_numeric(out["disponib"],   errors="coerce").fillna(1.0)
    out["perf_op"]    = pd.to_numeric(out["perf_op"],    errors="coerce").fillna(1.0)
    out = out.dropna(subset=["centro"])
    verificar_prefixo_centro(out, aba, log)
    amostra = out[["centro","peca","div_carga","vol_int","div_volume","disponib","perf_op"]].head(3)
    for _, r in amostra.iterrows():
        log.append(f"   ✔ {r.centro}/{r.peca}: div_carga={r.div_carga}, vol_int={r.vol_int}, div_volume={r.div_volume}, disponib={r.disponib}, perf_op={r.perf_op}")
    log.append(f"   {len(out)} combinações")
    return out.copy()

def read_aplic(fb, log):
    aba = _get_aba("IMPUTAPLICAÇÃO")
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name=aba, header=0)
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{aba}': {e}\n\n{ABA_FORMATOS['IMPUTAPLICAÇÃO']}")
    log.append(f"✅ IMPUTAPLICAÇÃO lido (aba: '{aba}'): {df.shape[0]}L")
    # Sempre usa índice posicional: coluna 0 = centro, coluna 1 = peça/referência
    df = df.rename(columns={df.columns[0]: "centro", df.columns[1]: "peca"})

    pc_trat_candidates = ["PEÇA\nTRATOR","PÇ/TRAT","PC/TRAT","PCTRAT","Peça Trator","pc_trat",
                          "ERA PEÇA\nTRATOR","ERA PEÇA TRATOR","ERA PECA TRATOR","ERA PECA\nTRATOR",
                          "QUANTIDADE\nE\nPOR\nVEÍCULO","QUANTIDADE E POR VEÍCULO",
                          "QUANTIDADE E POR VEICULO","QUANTIDADE\nE\nPOR\nVEICULO",
                          "QTD POR VEÍCULO","QTD POR VEICULO","QTD/VEÍCULO","QTD/VEICULO"]
    pc_trat_col = next((c for c in pc_trat_candidates if c in df.columns), None)
    if pc_trat_col is None and len(df.columns) > 3:
        pc_trat_col = df.columns[3]

    colunas_ignorar = {"centro", "peca"}
    if pc_trat_col:
        colunas_ignorar.add(pc_trat_col)

    # Aceita qualquer coluna a partir do índice 4 como modelo (qualquer nome)
    mcols = [
        c for i, c in enumerate(df.columns)
        if i >= 4
        and c not in colunas_ignorar
        and not str(c).startswith("Unnamed")
        and str(c).strip() not in ("", "nan", "None")
    ]

    if not mcols:
        raise ValueError(
            f"\n🔴 [IMPUTAPLICAÇÃO] Nenhuma coluna de modelo encontrada a partir da coluna 5!\n\n"
            f"   Colunas encontradas: {', '.join(str(c) for c in df.columns)}\n\n"
            f"   👉 Verifique se as colunas de modelo estão a partir da 5ª coluna da aba IMPUTAPLICAÇÃO.\n"
            f"   Os modelos podem ter qualquer nome — não precisam seguir um padrão específico."
        )

    log.append(f"   {len(mcols)} modelos: {mcols[:3]}{'...' if len(mcols)>3 else ''}")

    id_vars = ["centro", "peca", "pc_trat"] if pc_trat_col else ["centro", "peca"]
    if pc_trat_col:
        df["pc_trat"] = pd.to_numeric(df[pc_trat_col], errors="coerce").fillna(1.0).clip(lower=1.0)
        log.append(f"   PÇ/TRAT lido de '{pc_trat_col}'")
    else:
        df["pc_trat"] = 1.0

    melted = df[id_vars + mcols].melt(id_vars=id_vars, var_name="modelo", value_name="ativo")
    out = melted[melted["ativo"] == 1][id_vars + ["modelo"]].reset_index(drop=True)
    verificar_prefixo_centro(out, "IMPUTAPLICAÇÃO", log)
    log.append(f"   {len(out)} combinações ativas")
    return out

def validar(pmp, tempo, dist, aplic, dias):
    erros, alertas, oks = [], [], []
    chaves_tempo = set(zip(tempo.centro, tempo.peca))
    chaves_dist  = set(zip(dist.centro,  dist.peca))
    chaves_aplic = set(zip(aplic.centro, aplic.peca))

    zero_disp = dist[dist.disponib == 0]
    if len(zero_disp):
        exemplos = zero_disp[["centro","peca"]].head(3).apply(lambda r: f"{r.centro}/{r.peca}", axis=1).tolist()
        erros.append(f"Disponibilidade=0 em {len(zero_disp)} linha(s) — causa divisão por zero no índice de ciclo. Ex: {', '.join(exemplos)}")

    diff_td = chaves_tempo - chaves_dist
    if diff_td:
        exemplos = list(diff_td)[:3]
        erros.append(f"{len(diff_td)} combinações em IMPUTTEMPO sem IMPUTDISTRIBUIÇÃO — não terão carga calculada. Ex: {exemplos}")

    t_invalidos = tempo[(tempo.t_ciclo <= 0) | (tempo.t_labor < 0)]
    if len(t_invalidos):
        exemplos = t_invalidos[["centro","peca","t_ciclo","t_labor"]].head(3).to_dict("records")
        erros.append(f"Tempo de ciclo ≤0 ou labor <0 em {len(t_invalidos)} linha(s) — verifique IMPUTTEMPO. Ex: {exemplos[0]}")

    dist_num = dist.copy()
    dist_num["div_carga"]  = pd.to_numeric(dist_num["div_carga"],  errors="coerce").fillna(0)
    dist_num["div_volume"] = pd.to_numeric(dist_num["div_volume"], errors="coerce").fillna(0)
    zero_carga  = dist_num[dist_num["div_carga"]  == 0]
    zero_volume = dist_num[dist_num["div_volume"] == 0]
    if len(zero_carga):
        erros.append(f"div_carga=0 em {len(zero_carga)} linha(s) — zera completamente a carga daquele centro/peça")
    if len(zero_volume):
        erros.append(f"div_volume=0 em {len(zero_volume)} linha(s) — zera completamente a carga daquele centro/peça")

    sem_aplic = chaves_tempo - chaves_aplic
    if sem_aplic:
        exemplos = list(sem_aplic)[:3]
        alertas.append(f"{len(sem_aplic)} centro+peça sem modelo em IMPUTAPLICAÇÃO — não entrarão no cálculo de carga. Ex: {exemplos}")

    modelos_sem = set(pmp.modelo.unique()) - set(aplic.modelo.unique())
    if modelos_sem:
        exemplos = list(modelos_sem)[:5]
        alertas.append(f"{len(modelos_sem)} modelo(s) com demanda no PMP mas sem aplicação em nenhuma máquina: {exemplos}")

    merged = tempo.merge(dist, on=["centro","peca"], how="inner")
    labor_maior = merged[merged.t_labor > merged.t_ciclo]
    if len(labor_maior):
        exemplos = labor_maior[["centro","peca"]].head(3).apply(lambda r: f"{r.centro}/{r.peca}", axis=1).tolist()
        alertas.append(f"{len(labor_maior)} linha(s) com t_labor > t_ciclo — operador mais ocupado que a máquina, verifique se é intencional. Ex: {', '.join(exemplos)}")

    for m in MESES:
        qtd_m = pmp[pmp.mes==m].qtd.sum() if len(pmp[pmp.mes==m]) else 0
        if qtd_m > 0 and dias.get(m,0) == 0:
            alertas.append(f"Mês '{m}' tem {int(qtd_m)} peças no PMP mas dias trabalhados=0 — mês será ignorado no cálculo")

    for m in MESES:
        qtd_m = pmp[pmp.mes==m].qtd.sum() if len(pmp[pmp.mes==m]) else 0
        if qtd_m == 0 and dias.get(m,0) > 0:
            alertas.append(f"Mês '{m}' tem {dias[m]} dias configurados mas nenhuma demanda no PMP — headcount será zero")

    dist_num2 = dist.copy()
    dist_num2["div_carga"] = pd.to_numeric(dist_num2["div_carga"], errors="coerce").fillna(0)
    soma_carga = dist_num2.groupby("centro")["div_carga"].sum()
    sobrecarga = soma_carga[soma_carga > 1.001]
    if len(sobrecarga):
        detalhes = [f"{c}={v:.2f}" for c,v in sobrecarga.items()][:4]
        alertas.append(f"div_carga soma >1 em {len(sobrecarga)} centro(s) — a carga está sendo multiplicada, verifique a distribuição. Ex: {', '.join(detalhes)}")

    disp_alta = dist[pd.to_numeric(dist.disponib, errors="coerce").fillna(0) > 1]
    if len(disp_alta):
        exemplos = disp_alta[["centro","peca","disponib"]].head(3).apply(lambda r: f"{r.centro} disp={r.disponib}", axis=1).tolist()
        alertas.append(f"Disponibilidade >1 (>100%) em {len(disp_alta)} linha(s) — verifique se está em decimal (ex: 0.85 = 85%). Ex: {', '.join(exemplos)}")

    pecas_com_demanda = set(pmp.merge(aplic, on="modelo")["peca"].unique()) if len(pmp) > 0 else set()
    centros_sem_demanda = set(tempo.centro.unique()) - set(
        tempo[tempo.peca.isin(pecas_com_demanda)].centro.unique()
    )
    if centros_sem_demanda:
        alertas.append(f"{len(centros_sem_demanda)} centro(s) sem nenhuma demanda ativa — aparecem em IMPUTTEMPO mas nenhuma peça deles tem produção no PMP: {sorted(centros_sem_demanda)[:5]}")

    dist_vi = dist.copy()
    dist_vi["vol_int"] = pd.to_numeric(dist_vi["vol_int"], errors="coerce")
    vi_alto = dist_vi[dist_vi["vol_int"] > 5]
    if len(vi_alto):
        alertas.append(f"vol_interna > 5 em {len(vi_alto)} linha(s) — valor incomum, verifique se não é um erro de digitação")

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
    df["vol_int"]    = pd.to_numeric(df["vol_int"],    errors="coerce").fillna(1.0)
    df["div_carga"]  = pd.to_numeric(df["div_carga"],  errors="coerce").fillna(0.0)
    df["div_volume"] = pd.to_numeric(df["div_volume"], errors="coerce").fillna(0.0)
    df["disponib"]   = pd.to_numeric(df["disponib"],   errors="coerce").fillna(1.0)
    df["perf_op"]    = pd.to_numeric(df["perf_op"],    errors="coerce").fillna(1.0) if "perf_op" in df.columns else 1.0
    df["indice_ciclo"] = (df.t_ciclo * df.div_carga * df.div_volume * df.vol_int) / (df.disponib * df.perf_op)
    df["min_ciclo"]    = df.indice_ciclo * df.qtd
    df["min_labor"]    = df.t_labor * df.div_carga * df.qtd * df.pc_trat.fillna(1.0)
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
            nA=aA; nB=aB; nC=aC
            if overrides and mes in overrides and cen in overrides[mes]:
                ov = overrides[mes][cen]
                if "A" in ov: nA=int(ov["A"]); aA = 1 if nA > 0 else 0
                if "B" in ov: nB=int(ov["B"]); aB = 1 if nB > 0 else 0
                if "C" in ov: nC=int(ov["C"]); aC = 1 if nC > 0 else 0
            centros.append({
                "centro":cen,"ocup_A":pA,"ocup_B":pB,"ocup_C":pC,
                "ativo_A":aA,"ativo_B":aB,"ativo_C":aC,
                "num_A":nA,"num_B":nB,"num_C":nC,
                "min_ciclo_total":mc,"min_labor_total":ml,
                "min_disp_A":minA,"min_disp_B":minB,"min_disp_C":minC,
                "horas_ciclo":mc/60,"horas_labor":ml/60,
                "horas_disp_A":d*heA*nA,"horas_disp_B":d*heB*nB,"horas_disp_C":d*heC*nC,
            })
        df_c = pd.DataFrame(centros)
        op_A = int(df_c.num_A.sum()); op_B = int(df_c.num_B.sum()); op_C = int(df_c.num_C.sum())
        def get_sup(key, t, op_count):
            cfg = suporte_cfg[key]
            if op_count == 0: return 0
            if cfg["modo"] == "auto":
                defaults = {"lavadora":{"A":1,"B":1,"C":0},"gravacao":{"A":1,"B":1,"C":0},
                            "preset":{"A":2,"B":1,"C":1},"coringa":{"A":1,"B":0,"C":0},
                            "facilitador":{"A":1,"B":1,"C":0}}
                return defaults[key][t]
            return cfg[t] if op_count > 0 else 0
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
            "Ativo A":int(row.num_A) if hasattr(row,"num_A") else int(row.ativo_A),
            "Ativo B":int(row.num_B) if hasattr(row,"num_B") else int(row.ativo_B),
            "Ativo C":int(row.num_C) if hasattr(row,"num_C") else int(row.ativo_C),
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

# ─────────────────────────────────────────
# TABELONA PURA — sem comparação com referência
# ─────────────────────────────────────────
def gerar_tabelona_pura(resultados, tempo, dist, aplic, pmp, dias, horas_turno, horas_efetivas, thresholds):
    import openpyxl as _opx
    from openpyxl.styles import PatternFill as _PF, Font as _Ft, Alignment as _Al, Border as _Bd, Side as _Sd
    _F_VERDE=_PF("solid",fgColor="92D050"); _F_AMAR=_PF("solid",fgColor="FFDE00")
    _F_AZUL=_PF("solid",fgColor="00B0F0"); _F_PRETO=_PF("solid",fgColor="000000")
    _F_CINZA=_PF("solid",fgColor="D9D9D9"); _F_CINZA2=_PF("solid",fgColor="BFBFBF")
    _F_BRANCO=_PF("solid",fgColor="FFFFFF"); _F_VERDE_JD=_PF("solid",fgColor="1F4D19"); _F_VERM=_PF("solid",fgColor="FF0000"); _F_VERM_S=_PF("solid",fgColor="FF9999")
    _F_AMAR_JD=_PF("solid",fgColor="FFDE00")
    _BRD=_Bd(left=_Sd(style="thin",color="AAAAAA"),right=_Sd(style="thin",color="AAAAAA"),
             top=_Sd(style="thin",color="AAAAAA"),bottom=_Sd(style="thin",color="AAAAAA"))
    def _ec(ws,r,c,val,fill=None,bold=False,color="000000",size=8,center=True,wrap=False):
        cell=ws.cell(row=r,column=c,value=val)
        cell.font=_Ft(name="Arial",bold=bold,color=color,size=size)
        cell.fill=fill or _F_BRANCO
        cell.alignment=_Al(horizontal="center" if center else "left",vertical="center",wrap_text=wrap)
        cell.border=_BRD
        return cell
    def _ec_pct(ws,r,c,val,fill=None,bold=False,color="000000",size=8):
        cell=_ec(ws,r,c,val,fill,bold,color,size)
        cell.number_format="0.0000000000%"
        return cell
    def _cor_pct(v):
        try:
            f=float(v)
            if f>=1.06: return _PF("solid",fgColor="FF0000")
            if f>=1.00: return _PF("solid",fgColor="FFFF00")
            if f>=0.40: return _PF("solid",fgColor="92D050")
            return _F_BRANCO
        except: return _F_BRANCO
    hA_t=horas_turno["A"]; hB_t=horas_turno["B"]; hC_t=horas_turno["C"]
    heA_t=horas_efetivas["A"]; heB_t=horas_efetivas["B"]; heC_t=horas_efetivas["C"]
    try:
        df_all_t=(aplic.merge(pmp,on="modelo").merge(tempo,on=["centro","peca"]).merge(dist,on=["centro","peca"]))
        if "vol_int" not in df_all_t.columns: df_all_t["vol_int"] = 1.0
        df_all_t["vol_int"]    = pd.to_numeric(df_all_t["vol_int"],    errors="coerce").fillna(1.0)
        df_all_t["div_carga"]  = pd.to_numeric(df_all_t["div_carga"],  errors="coerce").fillna(0.0)
        df_all_t["div_volume"] = pd.to_numeric(df_all_t["div_volume"], errors="coerce").fillna(0.0)
        df_all_t["disponib"]   = pd.to_numeric(df_all_t["disponib"],   errors="coerce").fillna(1.0)
        df_all_t["perf_op"]    = pd.to_numeric(df_all_t["perf_op"],    errors="coerce").fillna(1.0) if "perf_op" in df_all_t.columns else 1.0
        df_all_t["indice_ciclo"]=(df_all_t.t_ciclo*df_all_t.div_carga*df_all_t.div_volume*df_all_t.vol_int)/(df_all_t.disponib*df_all_t.perf_op)
        df_all_t["min_ciclo"]=df_all_t.indice_ciclo*df_all_t.qtd
        df_all_t["min_labor"]=df_all_t.t_labor*df_all_t.div_carga*df_all_t.qtd*df_all_t.pc_trat.fillna(1.0)
        agg_cp_t=df_all_t.groupby(["centro","peca","mes"])[["min_ciclo","min_labor"]].sum()
    except: agg_cp_t=pd.DataFrame()
    pares_cp = list(dist[["centro","peca"]].drop_duplicates().itertuples(index=False, name=None))
    modelos_lista = sorted(pmp["modelo"].unique().tolist())

    _dist_idx = {(r.centro, r.peca): r for r in dist.itertuples()}
    _tempo_idx = {(r.centro, r.peca): r for r in tempo.itertuples()}
    _aplic_set = set(zip(aplic.centro, aplic.peca, aplic.modelo))
    _pmp_pivot = pmp.pivot_table(index=["modelo","mes"], values="qtd", aggfunc="sum").to_dict()
    _aplic_pc_idx = {(r.centro, r.peca): float(r.pc_trat) if hasattr(r, "pc_trat") else 1.0 for r in aplic.drop_duplicates(["centro","peca"]).itertuples()}

    wb_out = _opx.Workbook(); primeira_t = True
    for mes_t in MESES:
        d_t = dias.get(mes_t, 0)
        if d_t == 0: continue
        r_auto = resultados.get(mes_t)
        if not r_auto: continue
        minA_t=d_t*hA_t*60; minB_t=d_t*hB_t*60; minC_t=d_t*hC_t*60
        if primeira_t: ws_out=wb_out.active; ws_out.title=mes_t[:10]; primeira_t=False
        else: ws_out=wb_out.create_sheet(mes_t[:10])
        ws_out.freeze_panes="F7"
        _F_CINZA_H=_PF("solid",fgColor="D9D9D9"); _F_VD_H=_PF("solid",fgColor="1F4D19")
        ws_out.merge_cells("A1:O1")
        _ec(ws_out,1,1,f"TOTAIS — {mes_t.upper()}",_F_VD_H,True,"FFFFFF",9,True)
        for ci_h,txt_h,f_h in [(16,"TURNO A",_F_VERDE),(17,"TURNO B",_F_AMAR),(18,"TURNO C",_F_AZUL)]:
            _ec(ws_out,1,ci_h,txt_h,f_h,True,"000000",8,True)
        ws_out.row_dimensions[1].height=14
        ws_out.merge_cells("A2:O2"); _ec(ws_out,2,1,"TOTAL DE MINUTOS",_F_CINZA_H,True,"000000",8,False)
        _ec(ws_out,2,16,minA_t,_F_VERDE,True,"000000",8); _ec(ws_out,2,17,minB_t,_F_AMAR,True,"000000",8); _ec(ws_out,2,18,minC_t,_F_AZUL,True,"000000",8)
        ws_out.row_dimensions[2].height=13
        ws_out.merge_cells("A3:O3"); _ec(ws_out,3,1,"TOTAL DE HORAS",_F_CINZA_H,True,"000000",8,False)
        _ec(ws_out,3,16,minA_t/60,_F_VERDE,True,"000000",8); _ec(ws_out,3,17,minB_t/60,_F_AMAR,True,"000000",8); _ec(ws_out,3,18,minC_t/60,_F_AZUL,True,"000000",8)
        ws_out.row_dimensions[3].height=13
        ws_out.merge_cells("A4:O4"); _ec(ws_out,4,1,"Nº DIAS TRABALHADOS",_F_CINZA_H,True,"000000",8,False)
        _ec(ws_out,4,16,d_t,_F_VERDE,True,"FF0000",9); _ec(ws_out,4,17,d_t,_F_AMAR,True,"FF0000",9); _ec(ws_out,4,18,d_t,_F_AZUL,True,"FF0000",9)
        ws_out.row_dimensions[4].height=13
        n_mod=len(modelos_lista)
        ws_out.merge_cells(f"A5:{get_column_letter(19+n_mod)}5")
        _ec(ws_out,5,1,f"RESUMO DA CARGA — {mes_t.upper()} ({d_t} dias)",_F_VERDE_JD,True,"FFFFFF",10,True)
        ws_out.row_dimensions[5].height=18
        hdrs_f=[("Máquina",_F_CINZA2,"000000"),("PEÇA",_F_CINZA2,"000000"),("DESCRIÇÃO",_F_CINZA2,"000000"),("PÇ/TRAT",_F_CINZA2,"000000"),("UM",_F_CINZA2,"000000"),("Tempo Ciclo (min)",_F_PRETO,"FFFFFF"),("Tempo Labor (min)",_F_PRETO,"FFFFFF"),("Div. Carga",_PF("solid",fgColor="FF0000"),"FFFF00"),("Vol. Interna",_F_CINZA2,"000000"),("Div. Volume",_PF("solid",fgColor="FF0000"),"FFFF00"),("Disponib.",_F_CINZA2,"000000"),("Perf. Op.",_F_CINZA2,"000000"),("Indice Ciclo",_F_CINZA2,"000000"),("JA.A",_F_VERDE,"000000"),("JA.B",_F_AMAR,"000000"),("JA.C",_F_AZUL,"000000"),("TOTAL CICLOS (MIN)",_F_CINZA,"000000"),("TOTAL LABOR (MIN)",_F_CINZA,"000000"),("TOTAL PECAS",_F_CINZA,"000000")]
        largs_t=[9,8,16,6,5,9,9,8,8,8,8,8,9,8,8,8,12,12,8]
        for ci_t,(h_t,f_t,cor_t) in enumerate(hdrs_f,1):
            _ec(ws_out,6,ci_t,h_t,f_t,True,cor_t,8,True,True); ws_out.column_dimensions[get_column_letter(ci_t)].width=largs_t[ci_t-1]
        for mi_t,mod_t in enumerate(modelos_lista):
            ci_t=20+mi_t; _ec(ws_out,6,ci_t,mod_t,_F_CINZA,True,"000000",7,True,True); ws_out.column_dimensions[get_column_letter(ci_t)].width=7
        ws_out.row_dimensions[6].height=42
        _qtd_mes = {mod: int(_pmp_pivot.get("qtd", {}).get((mod, mes_t), 0)) for mod in modelos_lista}
        ri_t=7
        for cen_t,peca_t in pares_cp:
            _dk = (cen_t, peca_t); _dr = _dist_idx.get(_dk); _tr = _tempo_idx.get(_dk)
            if _dr is None or _tr is None: continue
            tc=float(_tr.t_ciclo); tl=float(_tr.t_labor)
            dc=float(_dr.div_carga); vi=float(_dr.vol_int)
            dv=float(_dr.div_volume); di=float(_dr.disponib)
            po=float(_dr.perf_op) if hasattr(_dr, "perf_op") else 1.0
            idx_c=(tc*dc*dv*vi)/(di*po) if (di>0 and po>0) else 0
            try: mc_t=float(agg_cp_t.loc[(cen_t,peca_t,mes_t),"min_ciclo"])
            except: mc_t=0.0
            try: ml_t=float(agg_cp_t.loc[(cen_t,peca_t,mes_t),"min_labor"])
            except: ml_t=0.0
            pA_t=mc_t/minA_t if minA_t>0 else 0; pB_t=mc_t/minB_t if minB_t>0 else 0; pC_t=mc_t/minC_t if minC_t>0 else 0
            app_mod_v={}; tot_pecas=0
            for mod_t2 in modelos_lista:
                qtd_t = _qtd_mes.get(mod_t2, 0)
                flag_t = 1 if (cen_t, peca_t, mod_t2) in _aplic_set else 0
                app_mod_v[mod_t2]=qtd_t*flag_t; tot_pecas+=qtd_t*flag_t
            pc_t = _aplic_pc_idx.get((cen_t, peca_t), 1.0)
            _ec(ws_out,ri_t,1,cen_t,_F_BRANCO,False,"000000",8,False); _ec(ws_out,ri_t,2,peca_t,_F_BRANCO,False,"000000",8,False)
            _ec(ws_out,ri_t,3,"",_F_BRANCO,False,"000000",8,False); _ec(ws_out,ri_t,4,int(pc_t),_F_BRANCO,False,"000000",8); _ec(ws_out,ri_t,5,"PC",_F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,6,tc,_F_PRETO,False,"FFFFFF",8); _ec(ws_out,ri_t,7,tl,_F_PRETO,False,"FFFFFF",8)
            _fill_dc_p = _F_VERM if abs(dc-1.0)>0.001 else _F_BRANCO
            _fill_dv_p = _F_VERM if abs(dv-1.0)>0.001 else _F_BRANCO
            _ec(ws_out,ri_t,8,dc,_fill_dc_p,False,"000000",8); _ec(ws_out,ri_t,9,vi,_F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,10,dv,_fill_dv_p,False,"000000",8); _ec(ws_out,ri_t,11,di,_F_BRANCO,False,"000000",8)
            _fill_po_p = _F_VERM if abs(po-1.0)>0.001 else _F_BRANCO
            _ec(ws_out,ri_t,12,po,_fill_po_p,False,"000000",8)
            _ec(ws_out,ri_t,13,idx_c,_F_BRANCO,False,"000000",8)
            _ec_pct(ws_out,ri_t,14,pA_t,_cor_pct(pA_t)); _ec_pct(ws_out,ri_t,15,pB_t,_cor_pct(pB_t)); _ec_pct(ws_out,ri_t,16,pC_t,_cor_pct(pC_t))
            _ec(ws_out,ri_t,17,mc_t,_F_BRANCO,False,"000000",8); _ec(ws_out,ri_t,18,ml_t,_F_BRANCO,False,"000000",8); _ec(ws_out,ri_t,19,tot_pecas,_F_BRANCO,False,"000000",8)
            for mi_t2,mod_t2 in enumerate(modelos_lista):
                ci_t2=20+mi_t2; v_app_t=app_mod_v.get(mod_t2,0)
                _ec(ws_out,ri_t,ci_t2,v_app_t if v_app_t else None,_F_CINZA if v_app_t else _F_BRANCO,False,"000000",7)
            ws_out.row_dimensions[ri_t].height=13; ri_t+=1
        _COL_F=6; _ROW_START=66; _F_CINZA_D=_PF("solid",fgColor="D9D9D9")
        ws_out.merge_cells(start_row=_ROW_START,start_column=_COL_F,end_row=_ROW_START,end_column=_COL_F+9)
        _ec(ws_out,_ROW_START,_COL_F,"DADOS AUTOMÁTICOS",_F_CINZA_D,True,"000000",9,True); ws_out.row_dimensions[_ROW_START].height=14
        _r67=_ROW_START+1
        _ec(ws_out,_r67,_COL_F,"PERÍODO:",_F_CINZA_D,True,"000000",8,False); _ec(ws_out,_r67,_COL_F+1,mes_t,_F_BRANCO,True,"FF0000",9,False)
        _ec(ws_out,_r67,_COL_F+3,"DATA DE REVISÃO:",_F_CINZA_D,True,"000000",8,False); _ec(ws_out,_r67,_COL_F+4,datetime.now().strftime("%d/%m/%Y"),_F_BRANCO,True,"FF0000",9)
        _ec(ws_out,_r67,_COL_F+6,"HORAS POR TURNO DE TRABALHO",_F_CINZA_D,True,"000000",8,True); ws_out.row_dimensions[_r67].height=14
        _r68=_ROW_START+2
        _ec(ws_out,_r68,_COL_F+6,heA_t,_F_CINZA_D,True,"000000",8); _ec(ws_out,_r68,_COL_F+7,heB_t,_F_CINZA_D,True,"000000",8); _ec(ws_out,_r68,_COL_F+8,heC_t,_F_CINZA_D,True,"000000",8); ws_out.row_dimensions[_r68].height=14
        _r69=_ROW_START+3
        for _ci_g,_txt_g,_fg in [(_COL_F+1,"TURNO A",_F_VERDE),(_COL_F+2,"TURNO B",_F_AMAR),(_COL_F+3,"TURNO C",_F_AZUL),(_COL_F+4,"TURNO A",_F_VERDE),(_COL_F+5,"TURNO B",_F_AMAR),(_COL_F+6,"TURNO C",_F_AZUL),(_COL_F+7,"TURNO A",_F_VERDE),(_COL_F+8,"TURNO B",_F_AMAR),(_COL_F+9,"TURNO C",_F_AZUL)]:
            _ec(ws_out,_r69,_ci_g,_txt_g,_fg,True,"000000",8,True)
        ws_out.row_dimensions[_r69].height=14
        _r70=_ROW_START+4; _ec(ws_out,_r70,_COL_F,"Centro",_F_PRETO,True,"FFFFFF",8)
        for _ci_s,_txt_s in [(_COL_F+1,"% Ocup"),(_COL_F+2,"% Ocup"),(_COL_F+3,"% Ocup"),(_COL_F+4,"Nº Func"),(_COL_F+5,"Nº Func"),(_COL_F+6,"Nº Func"),(_COL_F+7,"Horas"),(_COL_F+8,"Horas"),(_COL_F+9,"Horas")]:
            _ec(ws_out,_r70,_ci_s,_txt_s,_F_PRETO,True,"FFFFFF",8)
        ws_out.row_dimensions[_r70].height=14
        _ri_c=_ROW_START+5
        if r_auto:
            for _,_crow in r_auto["centros"].iterrows():
                def _cbg_t(v):
                    if v>=1.06: return _PF("solid",fgColor="FF0000")
                    if v>=1.00: return _PF("solid",fgColor="FFFF00")
                    if v>=0.40: return _PF("solid",fgColor="92D050")
                    return _F_BRANCO
                _ec(ws_out,_ri_c,_COL_F,_crow.centro,_F_BRANCO,False,"000000",8,False)
                _ec_pct(ws_out,_ri_c,_COL_F+1,_crow.ocup_A,_cbg_t(_crow.ocup_A)); _ec_pct(ws_out,_ri_c,_COL_F+2,_crow.ocup_B,_cbg_t(_crow.ocup_B)); _ec_pct(ws_out,_ri_c,_COL_F+3,_crow.ocup_C,_cbg_t(_crow.ocup_C))
                _nA_t=int(_crow.num_A) if hasattr(_crow,"num_A") else int(_crow.ativo_A); _nB_t=int(_crow.num_B) if hasattr(_crow,"num_B") else int(_crow.ativo_B); _nC_t=int(_crow.num_C) if hasattr(_crow,"num_C") else int(_crow.ativo_C)
                _ec(ws_out,_ri_c,_COL_F+4,_nA_t,_F_VERDE if _crow.ativo_A else _F_AMAR,True,"000000",8); _ec(ws_out,_ri_c,_COL_F+5,_nB_t,_F_VERDE if _crow.ativo_B else _F_AMAR,True,"000000",8); _ec(ws_out,_ri_c,_COL_F+6,_nC_t,_F_AZUL if _crow.ativo_C else _F_CINZA,True,"000000",8)
                _ec(ws_out,_ri_c,_COL_F+7,_crow.horas_disp_A if _crow.ativo_A else 0,_F_VERDE if _crow.ativo_A else _F_BRANCO,True,"000000",8)
                _ec(ws_out,_ri_c,_COL_F+8,_crow.horas_disp_B if _crow.ativo_B else 0,_F_AMAR if _crow.ativo_B else _F_BRANCO,True,"000000",8)
                _ec(ws_out,_ri_c,_COL_F+9,_crow.horas_disp_C if _crow.ativo_C else 0,_F_AZUL if _crow.ativo_C else _F_BRANCO,True,"000000",8)
                ws_out.row_dimensions[_ri_c].height=13; _ri_c+=1
            _sup_a=r_auto["suporte"]
            for _snm,_skey in [("TOTAL DE OPERADORES",None),("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador"),("TOTAL POR TURNO",None),("TOTAL FUNCIONÁRIOS",None)]:
                _bold_s="TOTAL" in _snm; _bg_s=_F_AMAR_JD if _bold_s else _F_BRANCO; _fg_s="1F4D19" if _bold_s else "000000"
                _ec(ws_out,_ri_c,_COL_F,_snm,_bg_s,_bold_s,_fg_s,8,False)
                if _skey:
                    _sv=_sup_a[_skey]
                    for _ci_sv,_tk in [(_COL_F+4,"A"),(_COL_F+5,"B"),(_COL_F+6,"C")]: _ec(ws_out,_ri_c,_ci_sv,_sv[_tk],_F_VERDE if _tk=="A" else (_F_AMAR if _tk=="B" else _F_AZUL),True,"000000",8)
                    for _ci_hv,_tk,_hef in [(_COL_F+7,"A",heA_t),(_COL_F+8,"B",heB_t),(_COL_F+9,"C",heC_t)]:
                        _hv=_sv[_tk]*_hef*d_t; _ec(ws_out,_ri_c,_ci_hv,_hv if _hv else 0,_F_VERDE if _tk=="A" else (_F_AMAR if _tk=="B" else _F_AZUL),True,"000000",8)
                elif "TOTAL DE OPERADORES" in _snm:
                    for _ci_sv,_vv in [(_COL_F+4,r_auto["op_A"]),(_COL_F+5,r_auto["op_B"]),(_COL_F+6,r_auto["op_C"])]: _ec(ws_out,_ri_c,_ci_sv,_vv,_F_AMAR_JD,True,"1F4D19",8)
                    for _ci_hv,_vv,_hef in [(_COL_F+7,r_auto["op_A"],heA_t),(_COL_F+8,r_auto["op_B"],heB_t),(_COL_F+9,r_auto["op_C"],heC_t)]: _ec(ws_out,_ri_c,_ci_hv,_vv*_hef*d_t,_F_AMAR_JD,True,"1F4D19",8)
                elif "TOTAL POR TURNO" in _snm:
                    for _ci_sv,_vv in [(_COL_F+4,r_auto["tot_A"]),(_COL_F+5,r_auto["tot_B"]),(_COL_F+6,r_auto["tot_C"])]: _ec(ws_out,_ri_c,_ci_sv,_vv,_F_AMAR_JD,True,"1F4D19",8)
                    for _ci_hv,_vv,_hef in [(_COL_F+7,r_auto["tot_A"],heA_t),(_COL_F+8,r_auto["tot_B"],heB_t),(_COL_F+9,r_auto["tot_C"],heC_t)]: _ec(ws_out,_ri_c,_ci_hv,_vv*_hef*d_t,_F_AMAR_JD,True,"1F4D19",8)
                elif "FUNCIONÁRIOS" in _snm:
                    _ec(ws_out,_ri_c,_COL_F+4,r_auto["total"],_F_AMAR_JD,True,"1F4D19",9)
                    _th=r_auto["tot_A"]*heA_t*d_t+r_auto["tot_B"]*heB_t*d_t+r_auto["tot_C"]*heC_t*d_t
                    _ec(ws_out,_ri_c,_COL_F+7,_th,_F_AMAR_JD,True,"1F4D19",9)
                ws_out.row_dimensions[_ri_c].height=13; _ri_c+=1
            _ri_c+=1
            for _pnm,_pv,_dest in [("PROD. CICLO OPERACIONAL",r_auto["prod_ciclo_op"],False),("PROD. CICLO TOTAL",r_auto["prod_ciclo_tot"],False),("PROD. LABOR OPERACIONAL",r_auto["prod_labor_op"],False),("PROD. LABOR TOTAL ★",r_auto["prod_labor_tot"],True)]:
                ws_out.merge_cells(start_row=_ri_c,start_column=_COL_F,end_row=_ri_c,end_column=_COL_F+8)
                _ec(ws_out,_ri_c,_COL_F,_pnm,_F_AMAR_JD if _dest else _F_BRANCO,_dest,"1F4D19" if _dest else "000000",8,False)
                _ec_pct(ws_out,_ri_c,_COL_F+9,_pv,_F_AMAR_JD if _dest else _F_BRANCO)
                ws_out.row_dimensions[_ri_c].height=14; _ri_c+=1
        for _ci_w,_ww in [(_COL_F,14),(_COL_F+1,8),(_COL_F+2,8),(_COL_F+3,8),(_COL_F+4,8),(_COL_F+5,8),(_COL_F+6,8),(_COL_F+7,10),(_COL_F+8,10),(_COL_F+9,10)]:
            ws_out.column_dimensions[get_column_letter(_ci_w)].width=_ww
    _fb_ano = st.session_state.get("_fb_anual")
    _cp_data_ano = build_cp_data_anual(resultados, tempo, dist, aplic, pmp, file_bytes=_fb_ano)
    _horas_ano = read_horas_anual(_fb_ano)
    gerar_aba_anual(wb_out, resultados, label="ANO", cp_data=_cp_data_ano, horas_anual=_horas_ano)
    buf_out = BytesIO(); wb_out.save(buf_out); buf_out.seek(0)
    return buf_out


def build_cp_data_anual(resultados, tempo, dist, aplic, pmp, file_bytes=None):
    if file_bytes is not None:
        try:
            wb_ref = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
            if "AnoFY26" in wb_ref.sheetnames:
                ws_ref = wb_ref["AnoFY26"]
                rows_ref = list(ws_ref.rows)
                result = []
                for row in rows_ref[6:]:
                    vals = [c.value for c in row[:18]]
                    if not vals[0] or not str(vals[0]).startswith("CEN"): continue
                    try:
                        cen  = str(vals[0]); peca = str(vals[1])
                        pc_t = float(vals[3] or 1)
                        tc   = float(vals[5] or 0); tl = float(vals[6] or 0)
                        dc   = float(vals[7] or 0); vi = float(vals[8] or 1)
                        dv   = float(vals[9] or 0); di = float(vals[10] or 1)
                        idx  = float(vals[11] or 0)
                        mc   = float(vals[15] or 0); ml = float(vals[16] or 0)
                        qt   = int(vals[17] or 0)
                        result.append((cen, peca, pc_t, tc, tl, dc, vi, dv, di, idx, mc, ml, qt))
                    except: pass
                wb_ref.close()
                if result: return result
            wb_ref.close()
        except: pass
    if pmp is None or tempo is None or dist is None or aplic is None: return None
    meses_com_resultado = [m for m in MESES if resultados.get(m)]
    meses_com_pmp = [m for m in MESES if not pmp[pmp.mes==m].empty and pmp[pmp.mes==m]["qtd"].sum()>0]
    meses_c = list(dict.fromkeys([m for m in MESES if m in meses_com_resultado or m in meses_com_pmp]))
    if not meses_c: return None
    try:
        df = (aplic.merge(pmp, on="modelo").merge(tempo, on=["centro","peca"]).merge(dist, on=["centro","peca"]))
        if "vol_int" not in df.columns: df["vol_int"] = 1.0
        df["vol_int"]    = pd.to_numeric(df["vol_int"],    errors="coerce").fillna(1.0)
        df["div_carga"]  = pd.to_numeric(df["div_carga"],  errors="coerce").fillna(0.0)
        df["div_volume"] = pd.to_numeric(df["div_volume"], errors="coerce").fillna(0.0)
        df["disponib"]   = pd.to_numeric(df["disponib"],   errors="coerce").fillna(1.0)
        df["perf_op"]    = pd.to_numeric(df["perf_op"],    errors="coerce").fillna(1.0) if "perf_op" in df.columns else 1.0
        df["indice_ciclo"] = (df.t_ciclo*df.div_carga*df.div_volume*df.vol_int)/(df.disponib*df.perf_op)
        df["min_ciclo"] = df.indice_ciclo * df.qtd
        df["min_labor"] = df.t_labor * df.div_carga * df.qtd
        df_ano = df[df.mes.isin(meses_c)]
        agg = df_ano.groupby(["centro","peca"])[["min_ciclo","min_labor","qtd"]].sum().reset_index()
        attrs = df_ano.drop_duplicates(["centro","peca"])[
            ["centro","peca","t_ciclo","t_labor","div_carga","vol_int","div_volume","disponib","perf_op","indice_ciclo"]
        ].set_index(["centro","peca"])
        result = []
        for _, row in agg.iterrows():
            cen, peca = row.centro, row.peca
            try:
                at = attrs.loc[(cen, peca)]
                tc=float(at.t_ciclo); tl=float(at.t_labor)
                dc=float(at.div_carga); vi=float(at.vol_int)
                dv=float(at.div_volume); di=float(at.disponib)
                po=float(at.perf_op) if hasattr(at,"perf_op") else 1.0
                idx=float(at.indice_ciclo)
            except: tc=tl=dc=vi=dv=di=po=0.0; idx=0.0
            result.append((cen, peca, 1.0, tc, tl, dc, vi, dv, di, po, idx,
                           float(row.min_ciclo), float(row.min_labor), int(row.qtd)))
        return result
    except: return None

def read_horas_anual(file_bytes):
    if file_bytes is None: return None
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
        if "AnoFY26" not in wb.sheetnames:
            wb.close(); return None
        ws = wb["AnoFY26"]
        rows = list(ws.rows)
        l4 = [c.value for c in rows[3]]
        h_ciclo = float(l4[15]) if len(l4) > 15 and l4[15] else 0
        h_labor = float(l4[16]) if len(l4) > 16 and l4[16] else 0
        h_ativos = 0; h_todos = 0
        for row in rows:
            vals = [c.value for c in row]
            txt = str(vals[11]) if len(vals) > 11 and vals[11] else ""
            if "TOTAL DE OPERADORES" in txt:
                hA = float(vals[29]) if len(vals) > 29 and vals[29] else 0
                hB = float(vals[30]) if len(vals) > 30 and vals[30] else 0
                hC = float(vals[31]) if len(vals) > 31 and vals[31] else 0
                h_ativos = hA + hB + hC
            if "TOTAL FUNCION" in txt:
                h_todos = float(vals[29]) if len(vals) > 29 and vals[29] else 0
        wb.close()
        if h_ciclo > 0 and h_todos > 0:
            return {"h_ciclo": h_ciclo, "h_labor": h_labor,
                    "h_ativos": h_ativos, "h_todos": h_todos}
        return None
    except: return None

def calcular_ano_fy26(file_bytes, overrides_ano, horas_efetivas, suporte_cfg, horas_turno):
    if file_bytes is None: return None
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
        if "AnoFY26" not in wb.sheetnames:
            wb.close(); return None
        ws = wb["AnoFY26"]
        rows = list(ws.rows)
        l2  = [c.value for c in rows[1]]
        l5  = [c.value for c in rows[4]]
        l4d = [c.value for c in rows[3]]
        minA_ano = float(l2[12]) if l2[12] else 96300.0
        minB_ano = float(l2[13]) if l2[13] else 182970.0
        minC_ano = float(l2[14]) if l2[14] else 250380.0
        hA = float(l5[12]) if l5[12] else 7.5
        hB = float(l5[13]) if l5[13] else 14.25
        hC = float(l5[14]) if l5[14] else 19.5
        dias_ano = int(l4d[11]) if l4d[11] else 214
        heA = horas_efetivas.get("A", hA)
        heB = horas_efetivas.get("B", hB)
        heC = horas_efetivas.get("C", hC)
        thr_A = 0.0; thr_B = 0.0; thr_C = 0.0
        from collections import defaultdict
        cen_mc = defaultdict(float); cen_ml = defaultdict(float)
        for row in rows[6:]:
            vals = [c.value for c in row[:18]]
            if not vals[0] or not str(vals[0]).startswith("CEN"): continue
            cen = str(vals[0])
            cen_mc[cen] += float(vals[15] or 0)
            cen_ml[cen] += float(vals[16] or 0)
        cen_base = {}
        wb.close()
        centros = []
        for cen in sorted(cen_mc.keys()):
            mc = cen_mc[cen]; ml = cen_ml[cen]
            base = cen_base.get(cen, {})
            oA = mc / minA_ano if minA_ano > 0 else 0
            oB = mc / minB_ano if minB_ano > 0 else 0
            oC = mc / minC_ano if minC_ano > 0 else 0
            aA = base.get("aA", 1 if oA > 0.40 else 0)
            aB = base.get("aB", 1 if oA > 0.75 else 0)
            aC = base.get("aC", 1 if oB > 0.75 else 0)
            nA = aA; nB = aB; nC = aC
            if overrides_ano and cen in overrides_ano:
                ov = overrides_ano[cen]
                if "A" in ov: nA = int(ov["A"]); aA = 1 if nA > 0 else 0
                if "B" in ov: nB = int(ov["B"]); aB = 1 if nB > 0 else 0
                if "C" in ov: nC = int(ov["C"]); aC = 1 if nC > 0 else 0
            centros.append({
                "centro": cen, "ocup_A": oA, "ocup_B": oB, "ocup_C": oC,
                "ativo_A": aA, "ativo_B": aB, "ativo_C": aC,
                "num_A": nA, "num_B": nB, "num_C": nC,
                "min_ciclo_total": mc, "min_labor_total": ml,
                "min_disp_A": minA_ano, "min_disp_B": minB_ano, "min_disp_C": minC_ano,
                "horas_ciclo": mc / 60, "horas_labor": ml / 60,
                "horas_disp_A": dias_ano * heA * nA,
                "horas_disp_B": dias_ano * heB * nB,
                "horas_disp_C": dias_ano * heC * nC,
            })
        df_c = pd.DataFrame(centros)
        op_A = int(df_c.num_A.sum()); op_B = int(df_c.num_B.sum()); op_C = int(df_c.num_C.sum())
        def _sup(key, t, op_count):
            cfg = suporte_cfg[key]
            if op_count == 0: return 0
            if cfg["modo"] == "auto":
                defaults = {"lavadora":{"A":1,"B":1,"C":0},"gravacao":{"A":1,"B":1,"C":0},
                            "preset":{"A":2,"B":1,"C":1},"coringa":{"A":1,"B":0,"C":0},
                            "facilitador":{"A":1,"B":1,"C":0}}
                return defaults[key][t]
            return cfg[t] if op_count > 0 else 0
        lav={t:_sup("lavadora",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        gra={t:_sup("gravacao",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        pre={t:_sup("preset",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        cor={t:_sup("coringa",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        fac={t:_sup("facilitador",t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
        tot_A = op_A+lav["A"]+gra["A"]+pre["A"]+cor["A"]+fac["A"]
        tot_B = op_B+lav["B"]+gra["B"]+pre["B"]+cor["B"]+fac["B"]
        tot_C = op_C+lav["C"]+gra["C"]+pre["C"]+cor["C"]+fac["C"]
        h_ciclo  = float(df_c.horas_ciclo.sum())
        h_labor  = float(df_c.horas_labor.sum())
        h_ativos = float((df_c.horas_disp_A + df_c.horas_disp_B + df_c.horas_disp_C).sum())
        h_todos  = tot_A * dias_ano * heA + tot_B * dias_ano * heB + tot_C * dias_ano * heC
        return {
            "centros": df_c,
            "op_A": op_A, "op_B": op_B, "op_C": op_C,
            "tot_A": tot_A, "tot_B": tot_B, "tot_C": tot_C, "total": tot_A + tot_B + tot_C,
            "suporte": {"lavadora":lav,"gravacao":gra,"preset":pre,"coringa":cor,"facilitador":fac},
            "h_ciclo": h_ciclo, "h_labor": h_labor, "h_ativos": h_ativos, "h_todos": h_todos,
            "prod_ciclo_op":  h_ciclo / h_ativos if h_ativos > 0 else 0,
            "prod_ciclo_tot": h_ciclo / h_todos  if h_todos  > 0 else 0,
            "prod_labor_op":  h_labor / h_ativos if h_ativos > 0 else 0,
            "prod_labor_tot": h_labor / h_todos  if h_todos  > 0 else 0,
            "dias": dias_ano, "hA": hA, "hB": hB, "hC": hC,
            "heA": heA, "heB": heB, "heC": heC,
            "thr_A": thr_A, "thr_B": thr_B, "thr_C": thr_C,
            "minA": minA_ano, "minB": minB_ano, "minC": minC_ano,
        }
    except Exception as _e:
        return None

def build_cp_data_from_meses(resultados, tempo, dist, aplic, pmp, dias_por_mes, horas_turno, horas_efetivas, overrides_ano=None, suporte_cfg=None):
    from collections import defaultdict
    meses_com_dados = [m for m in MESES if resultados.get(m)]
    if not meses_com_dados: return None, None
    try:
        df = (aplic.merge(pmp, on="modelo").merge(tempo, on=["centro","peca"]).merge(dist, on=["centro","peca"]))
        df["vol_int"]    = pd.to_numeric(df.get("vol_int",1),    errors="coerce").fillna(1.0)
        df["div_carga"]  = pd.to_numeric(df["div_carga"],  errors="coerce").fillna(0.0)
        df["div_volume"] = pd.to_numeric(df["div_volume"], errors="coerce").fillna(0.0)
        df["disponib"]   = pd.to_numeric(df["disponib"],   errors="coerce").fillna(1.0)
        df["perf_op"]    = pd.to_numeric(df["perf_op"],    errors="coerce").fillna(1.0) if "perf_op" in df.columns else 1.0
        df["indice_ciclo"] = (df.t_ciclo*df.div_carga*df.div_volume*df.vol_int)/(df.disponib*df.perf_op)
        df_ano = df[df.mes.isin(meses_com_dados)]
        agg = df_ano.groupby(["centro","peca"])[["qtd"]].sum().reset_index()
        attrs = df_ano.drop_duplicates(["centro","peca"])[
            ["centro","peca","t_ciclo","t_labor","div_carga","vol_int","div_volume","disponib","perf_op","indice_ciclo","pc_trat"]
        ].set_index(["centro","peca"])
        cp_data = []
        for _, row in agg.iterrows():
            cen, peca = row.centro, row.peca
            try:
                at = attrs.loc[(cen,peca)]
                tc=float(at.t_ciclo); tl=float(at.t_labor)
                dc=float(at.div_carga); vi=float(at.vol_int)
                dv=float(at.div_volume); di=float(at.disponib)
                po=float(at.perf_op) if hasattr(at,"perf_op") else 1.0
                idx=float(at.indice_ciclo)
                pct=float(at.pc_trat) if hasattr(at,'pc_trat') else 1.0
                qt=int(row.qtd)
                mc=idx*qt; ml=tl*dc*qt
            except: continue
            cp_data.append((cen, peca, pct, tc, tl, dc, vi, dv, di, po, idx, mc, ml, qt))
        if not cp_data: return None, None
        dias_total = sum(dias_por_mes.get(m, 0) for m in meses_com_dados)
        hA = horas_turno.get("A", 7.5); hB = horas_turno.get("B", 14.25); hC = horas_turno.get("C", 19.5)
        heA = horas_efetivas.get("A", hA); heB = horas_efetivas.get("B", hB); heC = horas_efetivas.get("C", hC)
        minA = dias_total * hA * 60; minB = dias_total * hB * 60; minC = dias_total * hC * 60
        cen_mc = defaultdict(float); cen_ml = defaultdict(float)
        for (_cen,_peca,_pct,_tc,_tl,_dc,_vi,_dv,_di,_idx,_mc,_ml,_qt) in cp_data:
            cen_mc[_cen]+=_mc; cen_ml[_cen]+=_ml
        cen_ativos = defaultdict(lambda: {"aA":0,"aB":0,"aC":0})
        for m in meses_com_dados:
            r = resultados[m]
            for _, row in r["centros"].iterrows():
                cen = row.centro
                cen_ativos[cen]["aA"] = max(cen_ativos[cen]["aA"], 1 if int(row.ativo_A) > 0 else 0)
                cen_ativos[cen]["aB"] = max(cen_ativos[cen]["aB"], 1 if int(row.ativo_B) > 0 else 0)
                cen_ativos[cen]["aC"] = max(cen_ativos[cen]["aC"], 1 if int(row.ativo_C) > 0 else 0)
        centros = []
        for cen in sorted(cen_mc.keys()):
            mc=cen_mc[cen]; ml=cen_ml[cen]
            oA=mc/minA if minA>0 else 0; oB=mc/minB if minB>0 else 0; oC=mc/minC if minC>0 else 0
            ca = cen_ativos.get(cen,{"aA":0,"aB":0,"aC":0})
            aA=ca["aA"]; aB=ca["aB"]; aC=ca["aC"]
            nA = aA; nB = aB; nC = aC
            if overrides_ano and cen in overrides_ano:
                ov = overrides_ano[cen]
                if "A" in ov: nA = int(ov["A"]); aA = 1 if nA > 0 else 0
                if "B" in ov: nB = int(ov["B"]); aB = 1 if nB > 0 else 0
                if "C" in ov: nC = int(ov["C"]); aC = 1 if nC > 0 else 0
            centros.append({"centro":cen,"ocup_A":oA,"ocup_B":oB,"ocup_C":oC,
                            "ativo_A":aA,"ativo_B":aB,"ativo_C":aC,
                            "num_A":nA,"num_B":nB,"num_C":nC,
                            "min_ciclo_total":mc,"min_labor_total":ml,
                            "min_disp_A":minA,"min_disp_B":minB,"min_disp_C":minC,
                            "horas_ciclo":mc/60,"horas_labor":ml/60,
                            "horas_disp_A":dias_total*heA*nA,
                            "horas_disp_B":dias_total*heB*nB,
                            "horas_disp_C":dias_total*heC*nC})
        df_c = pd.DataFrame(centros)
        op_A=int(df_c.num_A.sum()); op_B=int(df_c.num_B.sum()); op_C=int(df_c.num_C.sum())
        h_ciclo=float(df_c.horas_ciclo.sum()); h_labor=float(df_c.horas_labor.sum())
        h_ativos=float((df_c.horas_disp_A+df_c.horas_disp_B+df_c.horas_disp_C).sum())
        if suporte_cfg:
            def _sup2(key,t,op_count):
                cfg=suporte_cfg[key]
                if op_count==0: return 0
                if cfg["modo"]=="auto":
                    defaults={"lavadora":{"A":1,"B":1,"C":0},"gravacao":{"A":1,"B":1,"C":0},
                              "preset":{"A":2,"B":1,"C":1},"coringa":{"A":1,"B":0,"C":0},
                              "facilitador":{"A":1,"B":1,"C":0}}
                    return defaults[key][t]
                return cfg[t]
            sup_d={k:{t:_sup2(k,t,[op_A,op_B,op_C]["ABC".index(t)]) for t in "ABC"}
                   for k in ["lavadora","gravacao","preset","coringa","facilitador"]}
            sup_tot_A=sum(sup_d[k]["A"] for k in sup_d); sup_tot_B=sum(sup_d[k]["B"] for k in sup_d); sup_tot_C=sum(sup_d[k]["C"] for k in sup_d)
        else:
            sup_d={"lavadora":{"A":1,"B":1,"C":0},"gravacao":{"A":1,"B":1,"C":0},
                   "preset":{"A":2,"B":1,"C":1},"coringa":{"A":1,"B":0,"C":0},"facilitador":{"A":1,"B":1,"C":0}}
            sup_tot_A=1+1+2+1+1; sup_tot_B=1+1+1+0+1; sup_tot_C=0+0+1+0+0
        tot_A=op_A+sup_tot_A; tot_B=op_B+sup_tot_B; tot_C=op_C+sup_tot_C
        h_todos=tot_A*dias_total*heA+tot_B*dias_total*heB+tot_C*dias_total*heC
        res_ano = {"centros":df_c,"op_A":op_A,"op_B":op_B,"op_C":op_C,
                   "tot_A":tot_A,"tot_B":tot_B,"tot_C":tot_C,"total":tot_A+tot_B+tot_C,
                   "suporte":sup_d,
                   "h_ciclo":h_ciclo,"h_labor":h_labor,"h_ativos":h_ativos,"h_todos":h_todos,
                   "prod_ciclo_op":h_ciclo/h_ativos if h_ativos>0 else 0,
                   "prod_ciclo_tot":h_ciclo/h_todos if h_todos>0 else 0,
                   "prod_labor_op":h_labor/h_ativos if h_ativos>0 else 0,
                   "prod_labor_tot":h_labor/h_todos if h_todos>0 else 0,
                   "dias":dias_total,"hA":hA,"hB":hB,"hC":hC,
                   "heA":heA,"heB":heB,"heC":heC,
                   "thr_A":0.0,"thr_B":0.0,"thr_C":0.0,
                   "minA":minA,"minB":minB,"minC":minC}
        return cp_data, res_ano
    except: return None, None


def show_memoria(r, mes, df_intermediario, agg, horas_turno, thresholds):
    st.markdown(f'<div class="jd-section">Memória de cálculo — {mes}</div>', unsafe_allow_html=True)
    sup=r["suporte"]; d=r["dias"]; hA,hB,hC=r["hA"],r["hB"],r["hC"]
    heA,heB,heC=r.get("heA",hA),r.get("heB",hB),r.get("heC",hC)
    st.markdown('<div class="mem-step"><span class="step-num">1</span> <b>Inputs utilizados</b></div>', unsafe_allow_html=True)
    c1,c2,c3=st.columns(3)
    c1.metric("Turno A",f"{r['minA']:.0f} min",f"{d}×{hA}×60"); c2.metric("Turno B",f"{r['minB']:.0f} min",f"{d}×{hB}×60"); c3.metric("Turno C",f"{r['minC']:.0f} min")
    df_inp=df_intermediario[df_intermediario.mes==mes][["centro","peca","modelo","t_ciclo","t_labor","div_carga","div_volume","vol_int","disponib","perf_op","qtd"]].head(8).copy()
    df_inp.columns=["Centro","Peça","Modelo","T.Ciclo","T.Labor","Div.Carga","Div.Volume","Vol.Int","Disponib","Perf.Op","Qtd"]
    st.dataframe(df_inp.reset_index(drop=True),use_container_width=True,hide_index=True)
    st.markdown('<div class="mem-step"><span class="step-num">2</span> <b>Índice de ciclo</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">indice_ciclo = (t_ciclo × div_carga × div_volume × vol_interna) ÷ (disponibilidade × performance_operador)</div>', unsafe_allow_html=True)
    st.markdown('<div class="mem-step"><span class="step-num">3</span> <b>Minutos por linha</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">min_ciclo = indice_ciclo × qtd_pecas<br>min_labor = t_labor × div_carga × pç_trat × qtd_pecas</div>', unsafe_allow_html=True)
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
    _buf_mem_key = f"mem_base_{mes}"
    if st.session_state.get(_buf_mem_key) is None:
        _buf_mem=BytesIO(); df_intermediario[df_intermediario.mes==mes].to_excel(_buf_mem,index=False); _buf_mem.seek(0)
        st.session_state[_buf_mem_key]=_buf_mem.read()
    st.download_button("📥 Baixar base completa pós-JOIN",data=st.session_state[_buf_mem_key],file_name=f"base_{mes}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key=f"dl_mem_{mes}")

def show_memoria_ano(res_base, df_intermediario, agg, horas_turno, thresholds):
    meses_ativos = [(m, res_base[m]) for m in MESES if res_base.get(m)]
    if not meses_ativos:
        st.warning("Nenhum mês calculado."); return
    st.markdown('<div class="jd-section">Memória de cálculo — 📅 ANO COMPLETO</div>', unsafe_allow_html=True)
    st.markdown('<div class="mem-step"><span class="step-num">1</span> <b>Inputs consolidados do ano</b></div>', unsafe_allow_html=True)
    n_meses = len(meses_ativos)
    dias_ano = sum(r["dias"] for _, r in meses_ativos)
    hA = meses_ativos[0][1]["hA"]; hB = meses_ativos[0][1]["hB"]; hC = meses_ativos[0][1]["hC"]
    heA = meses_ativos[0][1].get("heA", hA); heB = meses_ativos[0][1].get("heB", hB); heC = meses_ativos[0][1].get("heC", hC)
    minA_ano = dias_ano * hA * 60; minB_ano = dias_ano * hB * 60; minC_ano = dias_ano * hC * 60
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Meses calculados", n_meses)
    c2.metric("Dias trabalhados (total)", dias_ano)
    c3.metric("Min. disponíveis Turno A", f"{minA_ano:,.0f}")
    c4.metric("Min. disponíveis Turno B/C", f"{minB_ano:,.0f}")
    st.markdown('<div class="mem-step"><span class="step-num">2</span> <b>Índice de ciclo (igual ao mensal)</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">indice_ciclo = (t_ciclo × div_carga × div_volume × vol_interna) ÷ (disponibilidade × performance_operador)<br>min_ciclo_ano = Σ (indice_ciclo × pç_trat × qtd) em todos os meses</div>', unsafe_allow_html=True)
    st.markdown('<div class="mem-step"><span class="step-num">3</span> <b>Minutos acumulados por centro (todos os meses)</b></div>', unsafe_allow_html=True)
    from collections import defaultdict
    cen_mc = defaultdict(float); cen_ml = defaultdict(float)
    for _, r in meses_ativos:
        for _, row in r["centros"].iterrows():
            cen_mc[row.centro] += row.min_ciclo_total
            cen_ml[row.centro] += row.min_labor_total
    rows_cen = []
    for cen in sorted(cen_mc.keys()):
        mc = cen_mc[cen]; ml = cen_ml[cen]
        pA = mc / minA_ano if minA_ano > 0 else 0
        pB = mc / minB_ano if minB_ano > 0 else 0
        pC = mc / minC_ano if minC_ano > 0 else 0
        thr_A = thresholds["A"] / 100; thr_B = thresholds["B"] / 100; thr_C = thresholds["C"] / 100
        aA = 1 if pA > thr_A else 0; aB = 1 if pA > thr_B else 0; aC = 1 if pB > thr_C else 0
        rows_cen.append({"Centro": cen, "Σ Min.Ciclo (ano)": round(mc, 1), "Σ Min.Labor (ano)": round(ml, 1),
                         "Ocup. A (ano)": pA, "Ocup. B (ano)": pB, "Ocup. C (ano)": pC,
                         "Ativo A": aA, "Ativo B": aB, "Ativo C": aC})
    df_ano_cen = pd.DataFrame(rows_cen)
    st.dataframe(df_ano_cen.style.format({"Ocup. A (ano)": "{:.1%}", "Ocup. B (ano)": "{:.1%}", "Ocup. C (ano)": "{:.1%}", "Σ Min.Ciclo (ano)": "{:,.1f}", "Σ Min.Labor (ano)": "{:,.1f}"}),
                 use_container_width=True, hide_index=True)
    st.markdown('<div class="mem-step"><span class="step-num">4</span> <b>Ativação de turno — lógica anual</b></div>', unsafe_allow_html=True)
    st.markdown(f"- Turno A abre se ocup_A (ano) > **{thresholds['A']}%**\n- Turno B abre se ocup_A (ano) > **{thresholds['B']}%**\n- Turno C abre se ocup_B (ano) > **{thresholds['C']}%**")
    st.markdown('<div class="mem-step"><span class="step-num">5</span> <b>Evolução mês a mês</b></div>', unsafe_allow_html=True)
    rows_mes = []
    for m, r in meses_ativos:
        rows_mes.append({"Mês": m, "Dias": r["dias"],
                         "Op. A": r["op_A"], "Op. B": r["op_B"], "Op. C": r["op_C"],
                         "Total func.": r["total"],
                         "Labor Total": r["prod_labor_tot"], "Ciclo Total": r["prod_ciclo_tot"]})
    df_mes = pd.DataFrame(rows_mes)
    def _sty_mes(row):
        styles = [""] * len(row)
        try:
            lv = float(row["Labor Total"])
            cor = "#003D10" if lv >= 0.45 else ("#3D2D00" if lv >= 0.30 else "#3D0000")
            txt = "#B9F6CA" if lv >= 0.45 else ("#FFE57F" if lv >= 0.30 else "#FF8A80")
            for i, col in enumerate(df_mes.columns):
                if col == "Labor Total": styles[i] = f"background-color:{cor};color:{txt};font-weight:600"
        except: pass
        return styles
    st.dataframe(df_mes.style.apply(_sty_mes, axis=1).format({"Labor Total": "{:.1%}", "Ciclo Total": "{:.1%}"}),
                 use_container_width=True, hide_index=True)
    st.markdown('<div class="mem-step"><span class="step-num">6</span> <b>Totais médios por turno (ano)</b></div>', unsafe_allow_html=True)
    sh_ciclo = sum(r["h_ciclo"] for _, r in meses_ativos)
    sh_labor = sum(r["h_labor"] for _, r in meses_ativos)
    sh_ativos = sum(r["h_ativos"] for _, r in meses_ativos)
    sh_todos = sum(r["h_todos"] for _, r in meses_ativos)
    prod_lt = sh_labor / sh_todos if sh_todos > 0 else 0
    prod_ct = sh_ciclo / sh_todos if sh_todos > 0 else 0
    prod_lo = sh_labor / sh_ativos if sh_ativos > 0 else 0
    prod_co = sh_ciclo / sh_ativos if sh_ativos > 0 else 0
    media_A = sum(r["tot_A"] for _, r in meses_ativos) / n_meses
    media_B = sum(r["tot_B"] for _, r in meses_ativos) / n_meses
    media_C = sum(r["tot_C"] for _, r in meses_ativos) / n_meses
    media_tot = sum(r["total"] for _, r in meses_ativos) / n_meses
    prod_data = {"Indicador": ["Ciclo Operacional", "Ciclo Total", "Labor Operacional", "⭐ Labor Total"],
                 "Resultado": [f"{prod_co:.1%}", f"{prod_ct:.1%}", f"{prod_lo:.1%}", f"{prod_lt:.1%}"]}
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Média Turno A", f"{media_A:.1f}"); c2.metric("Média Turno B", f"{media_B:.1f}")
    c3.metric("Média Turno C", f"{media_C:.1f}"); c4.metric("Média Total", f"{media_tot:.1f}")
    st.markdown('<div class="formula-box">Labor Total ★ (ano) = Σ horas_labor ÷ Σ horas_todos (todos os meses)</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame(prod_data), use_container_width=True, hide_index=True)
    if st.session_state.get("mem_base_ano") is None:
        _buf_ano=BytesIO(); df_intermediario.to_excel(_buf_ano,index=False); _buf_ano.seek(0)
        st.session_state["mem_base_ano"]=_buf_ano.read()
    st.download_button("📥 Baixar base completa pós-JOIN (ano todo)", data=st.session_state["mem_base_ano"],
                       file_name="base_ano_completo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="dl_mem_ano")


def gerar_aba_anual(wb, resultados, label="ANO", cp_data=None, horas_anual=None, eh_cenario=False, ws_existente=None, res_ano_override=None):
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
    if cp_data:
        _cmc=defaultdict(float); _cml=defaultdict(float)
        for (_cen,_peca,_pct,_tc,_tl,_dc,_vi,_dv,_di,_po,_idx,_mc,_ml,_qt) in cp_data:
            _cmc[_cen]+=_mc; _cml[_cen]+=_ml
        cen_mc=_cmc; cen_ml=_cml
    if res_ano_override:
        _ro = res_ano_override
        tot_A_ano=_ro.get("tot_A",tot_A_ano); tot_B_ano=_ro.get("tot_B",tot_B_ano)
        tot_C_ano=_ro.get("tot_C",tot_C_ano); tot_func_ano=_ro.get("total",tot_func_ano)
        sum_hciclo=_ro.get("h_ciclo",sum_hciclo); sum_hlabor=_ro.get("h_labor",sum_hlabor)
        sum_hativos=_ro.get("h_ativos",sum_hativos); sum_htodos=_ro.get("h_todos",sum_htodos)
        sup_somas={k:{"A":_ro["suporte"][k]["A"]*n_meses,"B":_ro["suporte"][k]["B"]*n_meses,"C":_ro["suporte"][k]["C"]*n_meses}
                   for k in sup_somas} if _ro.get("suporte") else sup_somas
    if horas_anual and not eh_cenario:
        _hc=horas_anual["h_ciclo"]; _hl=horas_anual["h_labor"]
        _ha=horas_anual["h_ativos"]; _ht=horas_anual["h_todos"]
        prod_co=_hc/_ha if _ha>0 else 0; prod_ct=_hc/_ht if _ht>0 else 0
        prod_lo=_hl/_ha if _ha>0 else 0; prod_lt=_hl/_ht if _ht>0 else 0
    else:
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
    if res_ano_override and "centros" in res_ano_override and not res_ano_override["centros"].empty:
        _df_ro = res_ano_override["centros"]
        op_A_ano = int(_df_ro.num_A.sum()) if "num_A" in _df_ro.columns else int(_df_ro.ativo_A.sum())
        op_B_ano = int(_df_ro.num_B.sum()) if "num_B" in _df_ro.columns else int(_df_ro.ativo_B.sum())
        op_C_ano = int(_df_ro.num_C.sum()) if "num_C" in _df_ro.columns else int(_df_ro.ativo_C.sum())
        if res_ano_override.get("suporte"):
            sup_somas = {k: {"A": res_ano_override["suporte"][k]["A"]*n_meses,
                             "B": res_ano_override["suporte"][k]["B"]*n_meses,
                             "C": res_ano_override["suporte"][k]["C"]*n_meses}
                         for k in sup_somas}
    else:
        op_A_ano=sum(ativo_ano_A(c) for c in centros_ord)
        op_B_ano=sum(ativo_ano_B(c) for c in centros_ord)
        op_C_ano=sum(ativo_ano_C(c) for c in centros_ord)
    def sup_ano(key,t): return round(sup_somas[key][t]/n_meses) if n_meses else 0
    tot_suporte_A=sum(sup_ano(k,"A") for k in sup_somas)
    tot_suporte_B=sum(sup_ano(k,"B") for k in sup_somas)
    tot_suporte_C=sum(sup_ano(k,"C") for k in sup_somas)
    tot_A_calc=op_A_ano+tot_suporte_A; tot_B_calc=op_B_ano+tot_suporte_B; tot_C_calc=op_C_ano+tot_suporte_C
    tot_func_calc=tot_A_calc+tot_B_calc+tot_C_calc
    if cp_data and not (horas_anual and not eh_cenario) and not eh_cenario:
        _hc_cp = sum(mc for (_,_,_,_,_,_,_,_,_,_,_,mc,_,_) in cp_data) / 60
        _hl_cp = sum(ml for (_,_,_,_,_,_,_,_,_,_,_,_,ml,_) in cp_data) / 60
        _ha_cp = sum(hdA(c)+hdB(c)+hdC(c) for c in centros_ord)
        _ht_cp = tot_A_calc*dias_ano*heA + tot_B_calc*dias_ano*heB + tot_C_calc*dias_ano*heC
        if _ha_cp > 0: prod_co=_hc_cp/_ha_cp; prod_lo=_hl_cp/_ha_cp
        if _ht_cp > 0: prod_ct=_hc_cp/_ht_cp; prod_lt=_hl_cp/_ht_cp
    if ws_existente is not None:
        ws=ws_existente; ws.title=label; ws.freeze_panes="F7"
    else:
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
    # LINHA 5 — agora 19 colunas fixas (adicionou Perf. Op.)
    ws.merge_cells("A5:S5"); _e(ws,5,1,f"RESUMO DA CARGA — ANO ({dias_ano} dias / {n_meses} meses)",F_VD_JD_a,True,"FFFFFF",10,True)
    ws.row_dimensions[5].height=18
    # CABEÇALHO — Perf. Op. entre Disponib. e Indice Ciclo
    hdrs_f=[("Máquina",F_CINZA2_a,"000000"),("PEÇA",F_CINZA2_a,"000000"),("DESCRIÇÃO",F_CINZA2_a,"000000"),("PÇ/TRAT",F_CINZA2_a,"000000"),("UM",F_CINZA2_a,"000000"),
            ("Tempo Ciclo (min)",F_PRETO_a,"FFFFFF"),("Tempo Labor (min)",F_PRETO_a,"FFFFFF"),
            ("Div. Carga",PatternFill("solid",fgColor="FF0000"),"FFFF00"),("Vol. Interna",F_CINZA2_a,"000000"),
            ("Div. Volume",PatternFill("solid",fgColor="FF0000"),"FFFF00"),("Disponib.",F_CINZA2_a,"000000"),("Perf. Op.",F_CINZA2_a,"000000"),("Indice Ciclo",F_CINZA2_a,"000000"),
            ("JA.A",F_VERDE_a,"000000"),("JA.B",F_AMAR_a,"000000"),("JA.C",F_AZUL_a,"000000"),
            ("TOTAL CICLOS (MIN)",F_CINZA_a,"000000"),("TOTAL LABOR (MIN)",F_CINZA_a,"000000"),("TOTAL PECAS",F_CINZA_a,"000000")]
    largs=[9,8,16,6,5,9,9,8,8,8,8,8,9,8,8,8,12,12,8]
    for ci,(h,f,cor) in enumerate(hdrs_f,1):
        _e(ws,6,ci,h,f,True,cor,8,True,True); ws.column_dimensions[get_column_letter(ci)].width=largs[ci-1]
    ws.row_dimensions[6].height=42
    ri=7
    if cp_data:
        # cp_data agora tem 14 elementos: cen,peca,_pc_t,tc,tl,dc,vi,dv,di,po,idx_c,mc,ml,_qtd
        for (cen,peca,_pc_t,tc,tl,dc,vi,dv,di,po,idx_c,mc,ml,_qtd) in cp_data:
            pA=mc/minA_ano if minA_ano>0 else 0; pB=mc/minB_ano if minB_ano>0 else 0; pC=mc/minC_ano if minC_ano>0 else 0
            _e(ws,ri,1,cen,F_BRANCO_a,False,"000000",8,False); _e(ws,ri,2,peca,F_BRANCO_a,False,"000000",8,False)
            _e(ws,ri,3,"ANO",F_BRANCO_a,False,"000000",8,False); _e(ws,ri,4,int(_pc_t),F_BRANCO_a,False,"000000",8); _e(ws,ri,5,"PC",F_BRANCO_a,False,"000000",8)
            _e(ws,ri,6,round(tc,4) if tc else "",F_PRETO_a,False,"FFFFFF",8); _e(ws,ri,7,round(tl,4) if tl else "",F_PRETO_a,False,"FFFFFF",8)
            _e(ws,ri,8,round(dc,4) if dc else "",PatternFill("solid",fgColor="FF0000"),False,"FFFF00",8)
            _e(ws,ri,9,round(vi,4) if vi else "",F_BRANCO_a,False,"000000",8)
            _e(ws,ri,10,round(dv,4) if dv else "",PatternFill("solid",fgColor="FF0000"),False,"FFFF00",8)
            _e(ws,ri,11,round(di,4) if di else "",F_CINZA2_a,False,"000000",8)
            # col 12 = Perf. Op. (NOVO)
            _e(ws,ri,12,round(po,4) if po else "",F_CINZA2_a,False,"000000",8)
            # col 13 = Indice Ciclo
            _e(ws,ri,13,round(idx_c,4),F_BRANCO_a,False,"000000",8)
            # cols 14/15/16 = JA.A / JA.B / JA.C
            _e(ws,ri,14,f"{pA:.1%}",_cor_pct_a(pA),False,"000000",8)
            _e(ws,ri,15,f"{pB:.1%}",_cor_pct_a(pB),False,"000000",8)
            _e(ws,ri,16,f"{pC:.1%}",_cor_pct_a(pC),False,"000000",8)
            # cols 17/18/19 = TOTAL CICLOS / LABOR / PECAS
            _e(ws,ri,17,round(mc,4),F_BRANCO_a,False,"000000",8)
            _e(ws,ri,18,round(ml,4),F_BRANCO_a,False,"000000",8)
            _e(ws,ri,19,_qtd,F_BRANCO_a,False,"000000",8)
            ws.row_dimensions[ri].height=13; ri+=1
    else:
        for cen in centros_ord:
            mc=cen_mc[cen]; ml=cen_ml[cen]
            pA=mc/minA_ano if minA_ano>0 else 0; pB=mc/minB_ano if minB_ano>0 else 0; pC=mc/minC_ano if minC_ano>0 else 0
            _e(ws,ri,1,cen,F_BRANCO_a,False,"000000",8,False); _e(ws,ri,2,"—",F_BRANCO_a,False,"000000",8)
            for ci_z in range(3,14): _e(ws,ri,ci_z,"",F_BRANCO_a,False,"000000",8)
            _e(ws,ri,14,f"{pA:.1%}",_cor_pct_a(pA),False,"000000",8)
            _e(ws,ri,15,f"{pB:.1%}",_cor_pct_a(pB),False,"000000",8)
            _e(ws,ri,16,f"{pC:.1%}",_cor_pct_a(pC),False,"000000",8)
            _e(ws,ri,17,round(mc,4),F_BRANCO_a,False,"000000",8); _e(ws,ri,18,round(ml,4),F_BRANCO_a,False,"000000",8)
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
    for _ci,_txt in [(_CF+1,"% Ocup"),(_CF+2,"% Ocup"),(_CF+3,"% Ocup"),(_CF+4,"Nº Func"),(_CF+5,"Nº Func"),(_CF+6,"Nº Func"),(_CF+7,"Horas"),(_CF+8,"Horas"),(_CF+9,"Horas")]:
        _e(ws,_RS+3,_ci,_txt,F_PRETO_a,True,"FFFFFF",8); ws.row_dimensions[_RS+3].height=14
    _ri=_RS+4
    def _cbg_a(v):
        if v>=1.06: return F_VERM_a
        if v>=1.00: return PatternFill("solid",fgColor="FFFF00")
        if v>=0.40: return F_VERDE_a
        return F_BRANCO_a
    _ro_df = res_ano_override["centros"] if (res_ano_override and "centros" in res_ano_override and not res_ano_override["centros"].empty) else None
    _ro_map = {row.centro: row for _, row in _ro_df.iterrows()} if _ro_df is not None else {}
    for cen in centros_ord:
        oA=ocup_ano(cen,"A"); oB=ocup_ano(cen,"B"); oC=ocup_ano(cen,"C")
        if cen in _ro_map:
            _row_ro = _ro_map[cen]
            nA=int(_row_ro.num_A) if hasattr(_row_ro,"num_A") else int(_row_ro.ativo_A)
            nB=int(_row_ro.num_B) if hasattr(_row_ro,"num_B") else int(_row_ro.ativo_B)
            nC=int(_row_ro.num_C) if hasattr(_row_ro,"num_C") else int(_row_ro.ativo_C)
            aA=1 if nA>0 else 0; aB=1 if nB>0 else 0; aC=1 if nC>0 else 0
        else:
            aA=ativo_ano_A(cen); aB=ativo_ano_B(cen); aC=ativo_ano_C(cen)
            nA=aA; nB=aB; nC=aC
        _e(ws,_ri,_CF,cen,F_BRANCO_a,False,"000000",8,False)
        for _pci,_pv,_pbg in [(_CF+1,oA,_cbg_a(oA)),(_CF+2,oB,_cbg_a(oB)),(_CF+3,oC,_cbg_a(oC))]:
            _pc=_e(ws,_ri,_pci,_pv,_pbg,False,"000000",8)
            _pc.number_format="0.0000000000%"
        _e(ws,_ri,_CF+4,nA,F_VERDE_a if aA else F_AMAR_a,True,"000000",8); _e(ws,_ri,_CF+5,nB,F_VERDE_a if aB else F_AMAR_a,True,"000000",8); _e(ws,_ri,_CF+6,nC,F_AZUL_a if aC else F_CINZA_a,True,"000000",8)
        _e(ws,_ri,_CF+7,round(dias_ano*heA*nA,2) if aA else 0,F_VERDE_a if aA else F_BRANCO_a,True,"000000",8)
        _e(ws,_ri,_CF+8,round(dias_ano*heB*nB,2) if aB else 0,F_AMAR_a if aB else F_BRANCO_a,True,"000000",8)
        _e(ws,_ri,_CF+9,round(dias_ano*heC*nC,2) if aC else 0,F_AZUL_a if aC else F_BRANCO_a,True,"000000",8)
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
        _pv_cell=_e(ws,_ri,_CF+9,_pv,F_AM_JD_a if _dest else F_BRANCO_a,_dest,"1F4D19" if _dest else "000000",8)
        _pv_cell.number_format="0.0000000000%"
        ws.row_dimensions[_ri].height=14; _ri+=1
    for ci,w in enumerate([9,8,16,6,5,9,9,8,8,8,8,8,9,8,8,8,12,12,8],1): ws.column_dimensions[get_column_letter(ci)].width=w
    for _ci,_ww in [(_CF,14),(_CF+1,8),(_CF+2,8),(_CF+3,8),(_CF+4,8),(_CF+5,8),(_CF+6,8),(_CF+7,10),(_CF+8,10),(_CF+9,10)]:
        ws.column_dimensions[get_column_letter(_ci)].width=_ww
    _nota_row = _ri + 2
    F_NOTA = PatternFill("solid", fgColor="FFF2CC")
    F_NOTA_H = PatternFill("solid", fgColor="FFD966")
    _brd_n = Border(left=Side(style="thin",color="C9A700"),right=Side(style="thin",color="C9A700"),
                    top=Side(style="thin",color="C9A700"),bottom=Side(style="thin",color="C9A700"))
    ws.merge_cells(start_row=_nota_row, start_column=1, end_row=_nota_row, end_column=19)
    _nh = ws.cell(row=_nota_row, column=1, value="⚠️  NOTA SOBRE O TOTAL LABOR E TOTAL CICLOS — ANO")
    _nh.font = Font(name="Arial", bold=True, size=9, color="7D4E00")
    _nh.fill = F_NOTA_H
    _nh.border = _brd_n
    _nh.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
    ws.row_dimensions[_nota_row].height = 16
    _linhas_nota = [
        ("Por que alguns valores do TOTAL LABOR (MIN) ou TOTAL CICLOS (MIN) podem diferir do Excel de referência (AnoFY26)?", True),
        ("O App calcula os totais anuais a partir do INPUT_PMP × IMPUTAPLICAÇÃO × IMPUTTEMPO × IMPUTDISTRIBUIÇÃO.", False),
        ("O Excel de referência calcula via fórmulas próprias em cada aba mensal, que podem incorporar ajustes manuais,", False),
        ("   células editadas diretamente ou lógicas de arredondamento diferentes das fórmulas do App.", False),
        ("", False),
        ("✅  Os valores calculados pelo App são fiéis ao INPUT_PMP e às regras definidas nos inputs.", True),
        ("   Se desejar que o anual bata com o Excel de referência, alinhe o INPUT_PMP com as quantidades reais de cada aba mensal.", False),
    ]
    for i, (txt, bold) in enumerate(_linhas_nota):
        r = _nota_row + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=19)
        cell = ws.cell(row=r, column=1, value=txt)
        cell.font = Font(name="Arial", bold=bold, size=8, color="7D4E00")
        cell.fill = F_NOTA
        cell.border = _brd_n
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 14 if txt else 6
