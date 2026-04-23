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
import json as _json
import hashlib

# ─── Helpers globais ─────────────────────────────────────────────────────────
def safe_int(v):
    try: return int(float(v)) if v is not None else 0
    except: return 0

def safe_float(v):
    try: return float(v) if v is not None else 0.0
    except: return 0.0



# ─── Cache de objetos openpyxl (evita recriar Font/PatternFill em loops pesados) ──
_FONT_CACHE  = {}
_ALIGN_CACHE = {}
_FILL_CACHE  = {}
_BRD_CACHE   = [None]

def _get_font(bold=False, color="000000", size=9, italic=False):
    k = (bold, color, size, italic)
    if k not in _FONT_CACHE:
        _FONT_CACHE[k] = Font(name="Arial", bold=bold, color=color, size=size, italic=italic)
    return _FONT_CACHE[k]

def _get_align(center=True):
    if center not in _ALIGN_CACHE:
        _ALIGN_CACHE[center] = Alignment(
            horizontal="center" if center else "left", vertical="center", wrap_text=True)
    return _ALIGN_CACHE[center]

def _get_fill(hex_color):
    if hex_color not in _FILL_CACHE:
        _FILL_CACHE[hex_color] = PatternFill("solid", fgColor=hex_color)
    return _FILL_CACHE[hex_color]

def _get_brd():
    if _BRD_CACHE[0] is None:
        s = Side(style='thin', color='CCCCCC')
        _BRD_CACHE[0] = Border(left=s, right=s, top=s, bottom=s)
    return _BRD_CACHE[0]

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
.jd-header h1 { color: #FFFFFF; margin: 0; font-size: 21px; font-weight: 700; }
.jd-header p  { color: #b8d4b4; margin: 4px 0 0; font-size: 12px; }
.jd-section {
    font-size: 15px; font-weight: 700; color: #FFDE00;
    border-left: 4px solid #FFDE00; padding-left: 10px;
    margin: 22px 0 10px;
}
.jd-sub {
    font-size: 12px; font-weight: 600; color: #7BC67A;
    margin: 12px 0 4px; text-transform: uppercase; letter-spacing: .04em;
}
.aviso-erro  { background:#3D0000;border-left:4px solid #FF5252;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#FF8A80; }
.aviso-warn  { background:#3D2D00;border-left:4px solid #FFDE00;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#FFE57F; }
.aviso-ok    { background:#003D10;border-left:4px solid #69F0AE;border-radius:6px;padding:9px 13px;margin:5px 0;font-size:12px;color:#B9F6CA; }
.formula-box {
    background:#0D1117;color:#A8E6A3;font-family:monospace;
    padding:10px 14px;border-radius:6px;font-size:12px;line-height:1.8;
    border-left:3px solid #FFDE00;
}
.mem-step {
    background:#1A1A2E;border:1px solid #444;border-radius:8px;
    padding:12px 16px;margin:6px 0;color:#FAFAFA;
}
.mem-step b { color:#FFDE00; }
.mem-step .step-num {
    background:#FFDE00;color:#1F4D19;border-radius:50%;
    width:24px;height:24px;display:inline-flex;align-items:center;
    justify-content:center;font-size:12px;font-weight:700;margin-right:8px;
}
.diff-igual   { background:#003D10;color:#B9F6CA;font-weight:600;padding:2px 6px;border-radius:4px; }
.diff-pequeno { background:#3D2D00;color:#FFE57F;font-weight:600;padding:2px 6px;border-radius:4px; }
.diff-grande  { background:#3D0000;color:#FF8A80;font-weight:600;padding:2px 6px;border-radius:4px; }
.log-line { font-family:monospace;font-size:11px;color:#AAAAAA;padding:1px 0; }
.log-ok   { color:#69F0AE; }
.log-warn { color:#FFDE00; }
.log-err  { color:#FF5252; }
.cenario-card{background:#1A1A2E;border:1.5px solid #333;border-radius:10px;padding:14px 16px;margin:6px 0;color:#FAFAFA;}
.cenario-card b{color:#FFDE00;}
.cenario-card.ativo{border-color:#FFDE00;background:#1F2D1A;}
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
        "vol_int":    ["Vol.\nInterna","Vol. Interna","VOL_INT","vol_int","VOL. INTERNA","Volume Interna"],
        "div_volume": ["Divisão \nde\nVolume\nENTRE\nPEÇAS","Div Volume","DIV_VOLUME","div_volume"],
        "disponib":   ["Disponi-\nbilidade","Disponibilidade","DISPONIB","disponib"],
    },
}

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
        legend=dict(orientation="h",y=-0.32,x=0,
                    font=dict(size=10,color="#000000"),
                    bgcolor="rgba(255,255,255,0.95)",
                    bordercolor="#AAAAAA",borderwidth=1),
        height=480,plot_bgcolor="white",paper_bgcolor="white",
        xaxis=dict(showgrid=False,showticklabels=True,tickfont=dict(size=11,color="#1A1A1A")),
        yaxis=dict(showgrid=True,gridcolor="#E8E8E8",showticklabels=True,
                   tickfont=dict(size=11,color="#1A1A1A"),
                   title="Nº Funcionários",title_font=dict(size=12,color="#1A1A1A")),
        yaxis2=dict(title="Labor Total (%)",tickformat=".0f",ticksuffix="%",range=[0,100],
                    showticklabels=True,tickfont=dict(size=11,color="#1A1A1A"),
                    title_font=dict(size=12,color="#1A1A1A")))
    return fig

# ─────────────────────────────────────────
# EXPORTAÇÃO
# ─────────────────────────────────────────
# Fills globais para diagnóstico — usa cache para não recriar objetos
def _make_brd(): return _get_brd()
def _pf(color): return _get_fill(color)

# Lazy globals — só instanciados na primeira chamada de função que precisar deles
class _Fills:
    """Instancia PatternFill só quando acessado (lazy), evitando custo no startup."""
    _cache = {}
    def __getattr__(self, name):
        _map = {
            "_BRD_DIAG":      _make_brd,
            "VERDE_HEADER":   lambda: _pf("1F4D19"),
            "VERMELHO_HEADER":lambda: _pf("C62828"),
            "CINZA_HEADER":   lambda: _pf("555555"),
            "VERMELHO_CLARO": lambda: _pf("FFCDD2"),
            "VERMELHO_ESCURO":lambda: _pf("C62828"),
            "AMARELO":        lambda: _pf("FFDE00"),
            "VERDE":          lambda: _pf("E8F5E9"),
            "F_HDR_VERD":     lambda: _pf("1F4D19"),
            "F_HDR_AMAR":     lambda: _pf("FFDE00"),
            "F_HDR_CINZA":    lambda: _pf("555555"),
            "F_CINZA_CLR":    lambda: _pf("F4F4F4"),
            "F_BRANCO":       lambda: _pf("FFFFFF"),
            "F_VERM_CLR":     lambda: _pf("FFCDD2"),
            "F_AMAR":         lambda: _pf("FFDE00"),
            "F_VERDE":        lambda: _pf("E8F5E9"),
            "F_VERDE_MED":    lambda: _pf("C8E6C9"),
            "F_AZUL_CLR":     lambda: _pf("E3F2FD"),
            "C_VERDE":        lambda: _pf("92D050"),
            "C_AMAR":         lambda: _pf("FFFF00"),
            "C_AZUL":         lambda: _pf("00B0F0"),
            "C_PRETO":        lambda: _pf("000000"),
            "C_CINZA":        lambda: _pf("D9D9D9"),
            "C_BRANCO":       lambda: _pf("FFFFFF"),
            "C_ROSA":         lambda: _pf("FFB6C1"),
            "C_VERM_DIV":     lambda: _pf("FF0000"),
        }
        if name not in self._cache and name in _map:
            self._cache[name] = _map[name]()
        if name in self._cache:
            return self._cache[name]
        raise AttributeError(name)

_F = _Fills()

# Aliases diretos para compatibilidade com o código existente
_BRD_DIAG      = property(lambda s: _F._BRD_DIAG)
VERDE_HEADER    = _F.VERDE_HEADER
VERMELHO_HEADER = _F.VERMELHO_HEADER
CINZA_HEADER    = _F.CINZA_HEADER
VERMELHO_CLARO  = _F.VERMELHO_CLARO
VERMELHO_ESCURO = _F.VERMELHO_ESCURO
AMARELO         = _F.AMARELO
VERDE           = _F.VERDE
F_HDR_VERD      = _F.F_HDR_VERD
F_HDR_AMAR      = _F.F_HDR_AMAR
F_HDR_CINZA     = _F.F_HDR_CINZA
F_CINZA_CLR     = _F.F_CINZA_CLR
F_BRANCO        = _F.F_BRANCO
F_VERM_CLR      = _F.F_VERM_CLR
F_AMAR          = _F.F_AMAR
F_VERDE         = _F.F_VERDE
F_VERDE_MED     = _F.F_VERDE_MED
F_AZUL_CLR      = _F.F_AZUL_CLR
C_VERDE         = _F.C_VERDE
C_AMAR          = _F.C_AMAR
C_AZUL          = _F.C_AZUL
C_PRETO         = _F.C_PRETO
C_CINZA         = _F.C_CINZA
C_BRANCO        = _F.C_BRANCO
C_ROSA          = _F.C_ROSA
C_VERM_DIV      = _F.C_VERM_DIV
_BRD_DIAG       = _make_brd()

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
    cell.font      = _get_font(bold, color, size, italic)
    cell.fill      = fill or F_BRANCO
    cell.alignment = _get_align(center)
    cell.border    = _get_brd()
    return cell

def cell_style(ws, r, c, val, fill=None, bold=False, color="000000", font_size=9, center=True, italic=False, comment_text=None):
    return ec(ws, r, c, val, fill, bold, color, font_size, center, italic, comment_text)

# ─────────────────────────────────────────
# DIAGNÓSTICO DE DIVERGÊNCIAS — EXCEL DE SAÍDA
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=3600)
def gerar_excel_diagnostico(file_bytes):
    """Gera Excel com diagnóstico completo de divergências entre inputs e Excel de referência."""
    MESES = ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
             "Maio","Junho","Julho","Agosto","Setembro","Outubro"]

    # ── Ler inputs ──────────────────────────────────────────────────────────
    tempo_raw = pd.read_excel(BytesIO(file_bytes), sheet_name='IMPUTTEMPO', header=0)
    dist_raw  = pd.read_excel(BytesIO(file_bytes), sheet_name='IMPUTDISTRIBUIÇÃO', header=0)
    aplic_raw = pd.read_excel(BytesIO(file_bytes), sheet_name='IMPUTAPLICAÇÃO', header=0)
    pmp_raw   = pd.read_excel(BytesIO(file_bytes), sheet_name='INPUT_PMP', header=None)

    tempo_raw = tempo_raw.rename(columns={
        tempo_raw.columns[0]:"centro", tempo_raw.columns[1]:"peca",
        tempo_raw.columns[5]:"t_ciclo", tempo_raw.columns[6]:"t_labor"})
    dist_raw = dist_raw.rename(columns={
        dist_raw.columns[0]:"centro", dist_raw.columns[1]:"peca",
        dist_raw.columns[7]:"div_carga", dist_raw.columns[9]:"div_volume",
        dist_raw.columns[10]:"disponib"})

    # ── Ler base do Excel de referência (NovFY26) ──────────────────────────
    wb_ref = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    abas = wb_ref.sheetnames
    aba_ref = next((a for a in ["NovFY26","DezFY26","JanFY26"] if a in abas), None)

    excel_base = []
    if aba_ref:
        ws_ref = wb_ref[aba_ref]
        for r in range(7, 64):
            centro = ws_ref.cell(r,1).value; peca = ws_ref.cell(r,2).value
            if not centro or not peca: continue
            excel_base.append({
                "centro":     str(centro).strip(), "peca": str(peca).strip(),
                "t_ciclo":    ws_ref.cell(r,6).value,
                "t_labor":    ws_ref.cell(r,7).value,
                "div_carga":  ws_ref.cell(r,8).value,
                "div_volume": ws_ref.cell(r,10).value,
                "disponib":   ws_ref.cell(r,11).value,
                "indice_xl":  ws_ref.cell(r,12).value,
            })
    wb_ref.close()
    df_xl = pd.DataFrame(excel_base) if excel_base else pd.DataFrame()

    # ── Construir workbook de diagnóstico ──────────────────────────────────
    wb = openpyxl.Workbook()

    # ════════════════════════════════════════════════════════════════
    # ABA 1 — RESUMO GERAL
    # ════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "📋 RESUMO"

    ws1.merge_cells("A1:F1")
    c = ws1.cell(1, 1, "DIAGNÓSTICO DE DIVERGÊNCIAS — CALCULADORA DE USINAGEM")
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    c.fill = VERDE_HEADER
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:F2")
    c = ws1.cell(2, 1, "Este arquivo mostra onde os dados de input diferem do que o Excel de referência usa. Células em VERMELHO = divergência. Veja as abas seguintes para detalhes.")
    c.font = Font(name="Arial", size=9, color="333333", italic=True)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws1.row_dimensions[2].height = 24

    # Legenda
    ws1.cell(4,1, "LEGENDA:").font = Font(bold=True, size=9)
    for r, (cor, fill, texto) in enumerate([
        (5, VERMELHO_CLARO, "🔴 VERMELHO — Valor diferente do Excel de referência → precisa verificar"),
        (6, AMARELO,        "🟡 AMARELO — Valor igual, mas índice calculado difere (efeito de outro campo)"),
        (7, VERDE,          "🟢 VERDE — Sem divergência detectada"),
    ], 5):
        ws1.cell(r,2, texto).fill = fill
        ws1.cell(r,2).font = Font(size=9)
        ws1.cell(r,2).alignment = Alignment(horizontal="left")
    ws1.column_dimensions["B"].width = 70

    # Contadores
    df_dist_clean = dist_raw[["centro","peca","div_carga","div_volume","disponib"]].dropna(subset=["centro"])
    df_cmp_res = df_xl.merge(df_dist_clean, on=["centro","peca"], suffixes=("_xl","_inp"), how="outer") if not df_xl.empty else pd.DataFrame()

    n_div = 0
    if not df_cmp_res.empty:
        for _, row in df_cmp_res.iterrows():
            for xl_col, inp_col in [("div_carga_xl","div_carga_inp"),("div_volume_xl","div_volume_inp"),("disponib_xl","disponib_inp")]:
                xl_v = row.get(xl_col); inp_v = row.get(inp_col)
                if xl_v is not None and inp_v is not None:
                    try:
                        if abs(float(xl_v) - float(inp_v)) > 0.001: n_div += 1
                    except: pass

    ws1.cell(9,1, f"Total de divergências encontradas no IMPUTDISTRIBUIÇÃO: {n_div}").font = Font(bold=True, size=10)
    ws1.cell(10,1, "→ Veja a aba 'IMPUTDISTRIBUIÇÃO' para ver linha a linha o que está diferente").font = Font(size=9, italic=True)
    ws1.cell(12,1, "→ Veja a aba 'IMPUTTEMPO' para conferir tempos de ciclo e labor").font = Font(size=9, italic=True)
    ws1.cell(13,1, "→ Veja a aba 'IMPUTAPLICAÇÃO' para conferir quais modelos passam por cada centro").font = Font(size=9, italic=True)
    ws1.column_dimensions["A"].width = 70

    # ════════════════════════════════════════════════════════════════
    # ABA 2 — IMPUTDISTRIBUIÇÃO COM DESTAQUE
    # ════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("IMPUTDISTRIBUIÇÃO")

    headers = ["Centro", "Peça", "Div. Carga\n(Input)", "Div. Carga\n(Excel Ref.)", "Div. Volume\n(Input)",
               "Div. Volume\n(Excel Ref.)", "Disponib.\n(Input)", "Disponib.\n(Excel Ref.)",
               "Índice Ciclo\n(App calcula)", "Índice Ciclo\n(Excel usa)", "Δ Índice", "Situação"]
    col_widths = [12, 12, 14, 14, 14, 14, 12, 12, 14, 14, 10, 30]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell_style(ws2, 1, ci, h, CINZA_HEADER, True, "FFFFFF", font_size=9)
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.row_dimensions[1].height = 36

    ri = 2
    if not df_xl.empty:
        df_tempo_clean = tempo_raw[["centro","peca","t_ciclo","t_labor"]].dropna(subset=["centro"])
        df_full = df_xl.merge(df_dist_clean, on=["centro","peca"], suffixes=("_xl","_inp"), how="outer")
        df_full = df_full.merge(df_tempo_clean, on=["centro","peca"], how="left")

        for _, row in df_full.iterrows():
            bg = VERDE if ri % 2 == 0 else PatternFill("solid", fgColor="F1F8F1")

            dc_xl  = row.get("div_carga_xl");  dc_inp = row.get("div_carga_inp")
            dv_xl  = row.get("div_volume_xl"); dv_inp = row.get("div_volume_inp")
            di_xl  = row.get("disponib_xl");   di_inp = row.get("disponib_inp")
            tc     = row.get("t_ciclo") or 0
            xl_idx = row.get("indice_xl")

            def safe(v): return round(float(v),4) if v is not None else "—"

            # Calcular índice app
            dc_v = float(dc_inp) if dc_inp is not None else 1
            dv_v = float(dv_inp) if dv_inp is not None else 1
            di_v = float(di_inp) if di_inp is not None else 0.9
            app_idx = round((float(tc) * dc_v * dv_v) / di_v, 4) if di_v > 0 else 0
            xl_idx_v = round(float(xl_idx), 4) if xl_idx is not None else 0
            delta_idx = round(xl_idx_v - app_idx, 4)

            def diverge(xl, inp):
                if xl is None or inp is None: return False
                try: return abs(float(xl) - float(inp)) > 0.001
                except: return False

            div_dc = diverge(dc_xl, dc_inp)
            div_dv = diverge(dv_xl, dv_inp)
            div_di = diverge(di_xl, di_inp)
            any_div = div_dc or div_dv or div_di

            comentarios = []
            if div_dc: comentarios.append(f"Div. Carga: Input={safe(dc_inp)} vs Excel={safe(dc_xl)}")
            if div_dv: comentarios.append(f"Div. Volume: Input={safe(dv_inp)} vs Excel={safe(dv_xl)}")
            if div_di: comentarios.append(f"Disponib.: Input={safe(di_inp)} vs Excel={safe(di_xl)}")
            if abs(delta_idx) > 0.5: comentarios.append(f"Índice diverge em {delta_idx:+.4f} min/peça → afeta ocupação dos turnos")

            situacao = "✅ OK" if not any_div else f"🔴 {len(comentarios)} campo(s) diferente(s)"
            comentario_txt = "\n".join(comentarios) if comentarios else None

            cell_style(ws2, ri, 1,  row.get("centro",""), bg, center=False)
            cell_style(ws2, ri, 2,  row.get("peca",""), bg, center=False)
            cell_style(ws2, ri, 3,  safe(dc_inp), VERMELHO_CLARO if div_dc else bg,
                       comment_text=f"IMPUTDISTRIBUIÇÃO usa: {safe(dc_inp)}\nExcel de referência usa: {safe(dc_xl)}\nDiferença: {round(float(dc_xl or 0)-float(dc_inp or 0),4)}" if div_dc else None)
            cell_style(ws2, ri, 4,  safe(dc_xl), VERMELHO_CLARO if div_dc else bg)
            cell_style(ws2, ri, 5,  safe(dv_inp), VERMELHO_CLARO if div_dv else bg,
                       comment_text=f"IMPUTDISTRIBUIÇÃO usa: {safe(dv_inp)}\nExcel de referência usa: {safe(dv_xl)}" if div_dv else None)
            cell_style(ws2, ri, 6,  safe(dv_xl), VERMELHO_CLARO if div_dv else bg)
            cell_style(ws2, ri, 7,  safe(di_inp), VERMELHO_CLARO if div_di else bg,
                       comment_text=f"IMPUTDISTRIBUIÇÃO usa: {safe(di_inp)}\nExcel de referência usa: {safe(di_xl)}" if div_di else None)
            cell_style(ws2, ri, 8,  safe(di_xl), VERMELHO_CLARO if div_di else bg)
            cell_style(ws2, ri, 9,  app_idx, AMARELO if abs(delta_idx) > 0.5 else bg)
            cell_style(ws2, ri, 10, xl_idx_v, AMARELO if abs(delta_idx) > 0.5 else bg)
            cell_style(ws2, ri, 11, f"{delta_idx:+.4f}" if delta_idx != 0 else "0",
                       VERMELHO_CLARO if abs(delta_idx) > 0.5 else bg)
            cell_style(ws2, ri, 12, situacao, VERMELHO_CLARO if any_div else VERDE,
                       comment_text=comentario_txt, bold=any_div)
            ws2.row_dimensions[ri].height = 16
            ri += 1

    # ════════════════════════════════════════════════════════════════
    # ABA 3 — IMPUTTEMPO
    # ════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("IMPUTTEMPO")
    headers3 = ["Centro","Peça","T. Ciclo\n(Input)","T. Ciclo\n(Excel Ref.)","T. Labor\n(Input)","T. Labor\n(Excel Ref.)","Situação"]
    for ci, h in enumerate(headers3, 1):
        cell_style(ws3, 1, ci, h, CINZA_HEADER, True, "FFFFFF", font_size=9)
        ws3.column_dimensions[get_column_letter(ci)].width = 16
    ws3.column_dimensions["G"].width = 30
    ws3.row_dimensions[1].height = 36

    df_tempo_cmp = df_xl.merge(tempo_raw[["centro","peca","t_ciclo","t_labor"]].dropna(subset=["centro"]),
                               on=["centro","peca"], suffixes=("_xl","_inp"), how="outer") if not df_xl.empty else pd.DataFrame()

    ri3 = 2
    for _, row in df_tempo_cmp.iterrows():
        bg = PatternFill("solid", fgColor="F8F8F8") if ri3 % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        tc_xl  = row.get("t_ciclo_xl");  tc_inp = row.get("t_ciclo_inp")
        tl_xl  = row.get("t_labor_xl");  tl_inp = row.get("t_labor_inp")

        def div(xl, inp):
            if xl is None or inp is None: return False
            try: return abs(float(xl) - float(inp)) > 0.01
            except: return False

        div_tc = div(tc_xl, tc_inp); div_tl = div(tl_xl, tl_inp)
        sit = "✅ OK" if not div_tc and not div_tl else f"🔴 {'Ciclo ' if div_tc else ''}{'Labor' if div_tl else ''} diferente"

        cell_style(ws3, ri3, 1, row.get("centro",""), bg, center=False)
        cell_style(ws3, ri3, 2, row.get("peca",""), bg, center=False)
        cell_style(ws3, ri3, 3, round(float(tc_inp),2) if tc_inp else "—",
                   VERMELHO_CLARO if div_tc else bg,
                   comment_text=f"Input: {tc_inp}\nExcel: {tc_xl}" if div_tc else None)
        cell_style(ws3, ri3, 4, round(float(tc_xl),2) if tc_xl else "—", VERMELHO_CLARO if div_tc else bg)
        cell_style(ws3, ri3, 5, round(float(tl_inp),2) if tl_inp else "—",
                   VERMELHO_CLARO if div_tl else bg,
                   comment_text=f"Input: {tl_inp}\nExcel: {tl_xl}" if div_tl else None)
        cell_style(ws3, ri3, 6, round(float(tl_xl),2) if tl_xl else "—", VERMELHO_CLARO if div_tl else bg)
        cell_style(ws3, ri3, 7, sit, VERMELHO_CLARO if (div_tc or div_tl) else VERDE,
                   bold=(div_tc or div_tl))
        ws3.row_dimensions[ri3].height = 15
        ri3 += 1

    # ════════════════════════════════════════════════════════════════
    # ABA 4 — SÓ DIVERGÊNCIAS (resumo executivo)
    # ════════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("🔴 SÓ DIVERGÊNCIAS")
    ws4.merge_cells("A1:G1")
    c = ws4.cell(1,1,"APENAS AS DIVERGÊNCIAS — corrija esses dados nos seus arquivos de input")
    c.font = Font(bold=True, color="FFFFFF", size=11)
    c.fill = VERMELHO_HEADER
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 24

    hdrs4 = ["Onde corrigir","Centro","Peça","Campo","Valor no Input","Valor no Excel","O que fazer"]
    for ci,h in enumerate(hdrs4,1):
        cell_style(ws4, 2, ci, h, CINZA_HEADER, True, "FFFFFF")
    ws4.column_dimensions["A"].width = 20; ws4.column_dimensions["B"].width = 12
    ws4.column_dimensions["C"].width = 12; ws4.column_dimensions["D"].width = 28
    ws4.column_dimensions["E"].width = 16; ws4.column_dimensions["F"].width = 16
    ws4.column_dimensions["G"].width = 50
    ws4.row_dimensions[2].height = 16

    ri4 = 3
    if not df_xl.empty:
        df_d = df_xl.merge(dist_raw[["centro","peca","div_carga","div_volume","disponib"]].dropna(subset=["centro"]),
                           on=["centro","peca"], suffixes=("_xl","_inp"), how="outer")
        df_d = df_d.merge(tempo_raw[["centro","peca","t_ciclo","t_labor"]].dropna(subset=["centro"]),
                          on=["centro","peca"], suffixes=("","_tempo"), how="outer")

        for _, row in df_d.iterrows():
            cen = row.get("centro",""); peca = row.get("peca","")
            for campo, xl_col, inp_col, aba_txt, acao in [
                ("Divisão de Carga",  "div_carga_xl",  "div_carga_inp",  "IMPUTDISTRIBUIÇÃO",
                 "Corrija o valor na coluna 'Divisão Carga Entre Máquinas' da aba IMPUTDISTRIBUIÇÃO"),
                ("Divisão de Volume", "div_volume_xl", "div_volume_inp", "IMPUTDISTRIBUIÇÃO",
                 "Corrija o valor na coluna 'Divisão de Volume Entre Peças' da aba IMPUTDISTRIBUIÇÃO"),
                ("Disponibilidade",   "disponib_xl",   "disponib_inp",   "IMPUTDISTRIBUIÇÃO",
                 "Corrija o valor na coluna 'Disponibilidade' da aba IMPUTDISTRIBUIÇÃO"),
            ]:
                xl_v = row.get(xl_col); inp_v = row.get(inp_col)
                if xl_v is None or inp_v is None: continue
                try:
                    if abs(float(xl_v) - float(inp_v)) > 0.001:
                        bg4 = VERMELHO_CLARO if ri4 % 2 == 0 else PatternFill("solid", fgColor="FFEBEE")
                        cell_style(ws4, ri4, 1, aba_txt, bg4, bold=True)
                        cell_style(ws4, ri4, 2, cen, bg4)
                        cell_style(ws4, ri4, 3, peca, bg4)
                        cell_style(ws4, ri4, 4, campo, bg4)
                        cell_style(ws4, ri4, 5, round(float(inp_v),4), bg4)
                        cell_style(ws4, ri4, 6, round(float(xl_v),4), VERMELHO_ESCURO, False, "FFFFFF")
                        cell_style(ws4, ri4, 7, acao, bg4, center=False)
                        ws4.row_dimensions[ri4].height = 16
                        ri4 += 1
                except: pass

    if ri4 == 3:
        ws4.cell(3,1,"✅ Nenhuma divergência encontrada nos dados de input!").font = Font(bold=True, color="1F4D19")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@st.cache_data(show_spinner=False, ttl=3600)
def gerar_diagnostico_mensal(file_bytes, res_app, tempo, dist, aplic, pmp, dias, horas_turno, thresholds):
    """Gera Excel com uma aba por mês mostrando App vs Excel célula a célula."""
    MAPA = {
        "Novembro":"NovFY26","Dezembro":"DezFY26","Janeiro":"JanFY26",
        "Fevereiro":"FevFY26","Março":"MarFY26","Abril":"AbrFY26",
        "Maio":"MaiFY26","Junho":"JunFY26","Julho":"JulFY26",
        "Agosto":"AgoFY26","Setembro":"SetFY26","Outubro":"OutFY26",
    }
    hA = horas_turno["A"]; hB = horas_turno["B"]
    thr_A = thresholds["A"]/100; thr_B = thresholds["B"]/100; thr_C = thresholds["C"]/100

    # ── Calcular ocupação por centro via inputs (uma vez para todos os meses) ──
    df_all = (aplic.merge(pmp, on="modelo")
                   .merge(tempo, on=["centro","peca"])
                   .merge(dist,  on=["centro","peca"]))
    if "vol_int" not in df_all.columns: df_all["vol_int"] = 1.0
    df_all["vol_int"] = pd.to_numeric(df_all["vol_int"], errors="coerce").fillna(1.0)
    df_all["indice_ciclo"] = (df_all.t_ciclo * df_all.div_carga * df_all.div_volume * df_all.vol_int) / df_all.disponib
    df_all["min_ciclo"]    = df_all.indice_ciclo * df_all.qtd
    agg_all = df_all.groupby(["centro","mes"])["min_ciclo"].sum().reset_index()

    # ── Ler dados do Excel de referência (iter_rows — rápido) ──
    wb_ref = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    abas_ref = wb_ref.sheetnames

    dados_xl = {}
    for mes, aba in MAPA.items():
        if aba not in abas_ref: continue
        ws_ref = wb_ref[aba]
        centros_xl = {}
        for row in ws_ref.iter_rows(min_row=69, max_row=88, min_col=23, max_col=29, values_only=True):
            cen = row[0]
            if not cen: continue
            centros_xl[str(cen).strip()] = {
                "ocup_A": safe_float(row[1]), "ocup_B": safe_float(row[2]),
                "ativo_A": safe_int(row[4]), "ativo_B": safe_int(row[5]), "ativo_C": safe_int(row[6])
            }
        dados_xl[mes] = {
            "centros": centros_xl,
            "tot_A": safe_int(ws_ref.cell(95,27).value),
            "tot_B": safe_int(ws_ref.cell(95,28).value),
            "tot_C": safe_int(ws_ref.cell(95,29).value),
            "total": safe_int(ws_ref.cell(96,27).value),
            "labor": safe_float(ws_ref.cell(101,30).value),
        }
    wb_ref.close()

    # ── Montar workbook de diagnóstico ──────────────────────────────────────
    wb = openpyxl.Workbook()

    # ── ABA CAPA ─────────────────────────────────────────────────────────────
    ws_capa = wb.active
    ws_capa.title = "CAPA"
    ws_capa.column_dimensions["A"].width = 5
    ws_capa.column_dimensions["B"].width = 45
    ws_capa.column_dimensions["C"].width = 20
    ws_capa.column_dimensions["D"].width = 20
    ws_capa.column_dimensions["E"].width = 12

    ws_capa.merge_cells("B1:E1")
    ec(ws_capa,1,2,"DIAGNÓSTICO DE DIVERGÊNCIAS — USINAGEM",F_HDR_VERD,True,"FFFFFF",14,True)
    ws_capa.row_dimensions[1].height = 32

    ws_capa.merge_cells("B2:E2")
    ec(ws_capa,2,2,"Comparação entre os dados de input e o Excel de referência, mês a mês",
       F_CINZA_CLR,False,"555555",9,True,True)
    ws_capa.row_dimensions[2].height = 18

    # Legenda
    ec(ws_capa,4,2,"LEGENDA:",None,True,"000000",10,False)
    for r,(f,txt) in enumerate([
        (F_VERM_CLR, "🔴 VERMELHO — App e Excel diferem → verifique o input"),
        (F_AMAR,     "🟡 AMARELO — Diferença pequena (ocupação próxima do threshold)"),
        (F_VERDE,    "🟢 VERDE — Sem divergência"),
        (F_AZUL_CLR, "🔵 AZUL — Informativo (sem divergência de ativação)"),
    ],5):
        ws_capa.merge_cells(f"B{r}:E{r}")
        ec(ws_capa,r,2,txt,f,False,"000000",9,False)
        ws_capa.row_dimensions[r].height = 16

    # Índice de meses
    ec(ws_capa,10,2,"MÊS",F_HDR_CINZA,True,"FFFFFF",9); ec(ws_capa,10,3,"STATUS",F_HDR_CINZA,True,"FFFFFF",9)
    ec(ws_capa,10,4,"Δ TOTAL (App−Excel)",F_HDR_CINZA,True,"FFFFFF",9); ec(ws_capa,10,5,"Nº DIV.",F_HDR_CINZA,True,"FFFFFF",9)
    ws_capa.row_dimensions[10].height = 16

    resumo_capa = []

    # ── UMA ABA POR MÊS ─────────────────────────────────────────────────────
    for mes in MESES:
        r_app = res_app.get(mes)
        d     = dias.get(mes, 0)
        if not r_app or d == 0: continue

        xl = dados_xl.get(mes, {})
        centros_xl = xl.get("centros", {})

        minA = d*hA*60; minB = d*hB*60

        # Centros do app
        agg_mes = agg_all[agg_all.mes == mes]
        app_centros = {}
        for _, row in agg_mes.iterrows():
            mc = row.min_ciclo
            oA = mc/minA if minA>0 else 0
            oB = mc/minB if minB>0 else 0
            app_centros[row.centro] = {
                "ocup_A": oA, "ocup_B": oB,
                "ativo_A": 1 if oA>thr_A else 0,
                "ativo_B": 1 if oA>thr_B else 0,
                "ativo_C": 1 if oB>thr_C else 0,
            }

        todos_centros = sorted(set(list(app_centros.keys()) + list(centros_xl.keys())))

        # Criar aba do mês
        nome_aba = mes[:10]
        ws_mes = wb.create_sheet(nome_aba)
        ws_mes.freeze_panes = "C3"

        # Larguras
        for ci,w in [(1,12),(2,12),(3,13),(4,13),(5,9),(6,13),(7,13),(8,9),(9,13),(10,13),(11,9),(12,28)]:
            ws_mes.column_dimensions[get_column_letter(ci)].width = w

        # Cabeçalho linha 1 — grupos
        ws_mes.merge_cells("A1:B1"); ec(ws_mes,1,1,mes.upper(),F_HDR_VERD,True,"FFFFFF",11,True)
        ws_mes.merge_cells("C1:E1"); ec(ws_mes,1,3,"TURNO A",F_HDR_VERD,True,"FFFFFF",10,True)
        ws_mes.merge_cells("F1:H1"); ec(ws_mes,1,6,"TURNO B",F_HDR_AMAR,True,"1F4D19",10,True)
        ws_mes.merge_cells("I1:K1"); ec(ws_mes,1,9,"TURNO C",F_HDR_CINZA,True,"FFFFFF",10,True)
        ec(ws_mes,1,12,"",F_HDR_CINZA,True,"FFFFFF",10,True)
        ws_mes.row_dimensions[1].height = 22

        # Cabeçalho linha 2
        for ci,txt,fill in [
            (1,"Centro",F_HDR_CINZA),(2,"",F_HDR_CINZA),
            (3,"Ocup. App",F_HDR_VERD),(4,"Ocup. Excel",F_HDR_VERD),(5,"Ativo?",F_HDR_VERD),
            (6,"Ocup. App",F_HDR_AMAR),(7,"Ocup. Excel",F_HDR_AMAR),(8,"Ativo?",F_HDR_AMAR),
            (9,"Ocup. App",F_HDR_CINZA),(10,"Ocup. Excel",F_HDR_CINZA),(11,"Ativo?",F_HDR_CINZA),
            (12,"Diagnóstico",F_HDR_CINZA),
        ]:
            col_txt = "FFFFFF" if fill != F_HDR_AMAR else "1F4D19"
            ec(ws_mes,2,ci,txt,fill,True,col_txt,8,True)
        ws_mes.row_dimensions[2].height = 16

        n_div_mes = 0
        for ri, cen in enumerate(todos_centros, 3):
            ap = app_centros.get(cen, {})
            xl_c = centros_xl.get(cen, {})

            oA_ap = ap.get("ocup_A",0); oA_xl = xl_c.get("ocup_A",0)
            oB_ap = ap.get("ocup_B",0); oB_xl = xl_c.get("ocup_B",0)
            aA_ap = ap.get("ativo_A",0); aA_xl = xl_c.get("ativo_A",0)
            aB_ap = ap.get("ativo_B",0); aB_xl = xl_c.get("ativo_B",0)
            aC_ap = ap.get("ativo_C",0); aC_xl = xl_c.get("ativo_C",0)

            div_A = aA_ap != aA_xl
            div_B = aB_ap != aB_xl
            div_C = aC_ap != aC_xl
            any_div = div_A or div_B or div_C
            if any_div: n_div_mes += 1

            bg_row = F_CINZA_CLR if ri%2==0 else F_BRANCO

            # Diagnóstico textual
            diags = []
            if div_A: diags.append(f"Turno A: App={aA_ap} vs Excel={aA_xl} (ocup App={oA_ap:.0%} vs Excel={oA_xl:.0%})")
            if div_B: diags.append(f"Turno B: App={aB_ap} vs Excel={aB_xl} (ocup App={oA_ap:.0%} vs Excel={oA_xl:.0%})")
            if div_C: diags.append(f"Turno C: App={aC_ap} vs Excel={aC_xl} (ocup B App={oB_ap:.0%} vs Excel={oB_xl:.0%})")

            delta_A = abs(oA_ap - oA_xl)
            if any_div:
                if delta_A > 0.15: diags.append("→ Causa provável: volume ou fator de distribuição diferente no IMPUTDISTRIBUIÇÃO")
                else:              diags.append(f"→ Causa provável: ocupação próxima do threshold (thr_A>{thr_A:.0%}/thr_B>{thr_B:.0%})")

            diag_txt = " | ".join(diags) if diags else "✅ Sem divergência"
            comment_txt = "\n".join(diags) if diags else None

            # Cor por célula
            def cor_ocup(ap_v, xl_v, div):
                if div: return F_VERM_CLR
                if ap_v > 1.0 or xl_v > 1.0: return F_AMAR
                return bg_row

            def cor_ativo(a_ap, a_xl, div):
                if div: return F_VERM_CLR
                if a_ap: return F_VERDE
                return bg_row

            ec(ws_mes, ri, 1, cen, F_VERM_CLR if any_div else bg_row, any_div, "000000")
            ec(ws_mes, ri, 2, "← VER" if any_div else "", F_VERM_CLR if any_div else bg_row, True, "C62828" if any_div else "000000", 8)

            # Turno A
            ec(ws_mes,ri,3,f"{oA_ap:.1%}", cor_ocup(oA_ap,oA_xl,div_A))
            ec(ws_mes,ri,4,f"{oA_xl:.1%}", cor_ocup(oA_ap,oA_xl,div_A))
            ec(ws_mes,ri,5,f"App={aA_ap}/XL={aA_xl}", cor_ativo(aA_ap,aA_xl,div_A), div_A)

            # Turno B
            ec(ws_mes,ri,6,f"{oA_ap:.1%}", cor_ocup(oA_ap,oA_xl,div_B))
            ec(ws_mes,ri,7,f"{oA_xl:.1%}", cor_ocup(oA_ap,oA_xl,div_B))
            ec(ws_mes,ri,8,f"App={aB_ap}/XL={aB_xl}", cor_ativo(aB_ap,aB_xl,div_B), div_B)

            # Turno C
            ec(ws_mes,ri,9,f"{oB_ap:.1%}", cor_ocup(oB_ap,oB_xl,div_C))
            ec(ws_mes,ri,10,f"{oB_xl:.1%}", cor_ocup(oB_ap,oB_xl,div_C))
            ec(ws_mes,ri,11,f"App={aC_ap}/XL={aC_xl}", cor_ativo(aC_ap,aC_xl,div_C), div_C)

            # Diagnóstico
            ec(ws_mes,ri,12,diag_txt,
               F_VERM_CLR if any_div else bg_row,
               any_div,"000000",8,False,True,comment_txt)
            ws_mes.row_dimensions[ri].height = 15

        # Linha de totais
        ri_tot = len(todos_centros) + 4
        xl_tot = dados_xl.get(mes,{})
        app_tot = r_app["total"]; xl_tot_v = xl_tot.get("total",0)
        delta_tot = app_tot - xl_tot_v

        ws_mes.merge_cells(f"A{ri_tot}:B{ri_tot}")
        ec(ws_mes,ri_tot,1,"TOTAL POR TURNO",F_HDR_VERD,True,"FFFFFF",9)
        for ci,label,app_v,xl_v in [
            (3,f"App={r_app['tot_A']}", r_app['tot_A'], xl_tot.get('tot_A',0)),
            (4,f"Excel={xl_tot.get('tot_A',0)}", r_app['tot_A'], xl_tot.get('tot_A',0)),
            (6,f"App={r_app['tot_B']}", r_app['tot_B'], xl_tot.get('tot_B',0)),
            (7,f"Excel={xl_tot.get('tot_B',0)}", r_app['tot_B'], xl_tot.get('tot_B',0)),
            (9,f"App={r_app['tot_C']}", r_app['tot_C'], xl_tot.get('tot_C',0)),
            (10,f"Excel={xl_tot.get('tot_C',0)}", r_app['tot_C'], xl_tot.get('tot_C',0)),
        ]:
            div = abs(int(app_v or 0) - int(xl_v or 0)) > 0
            ec(ws_mes,ri_tot,ci,label,F_VERM_CLR if div else F_VERDE_MED,True,"000000",9)

        ws_mes.merge_cells(f"A{ri_tot+1}:B{ri_tot+1}")
        ec(ws_mes,ri_tot+1,1,"TOTAL FUNCIONÁRIOS",F_HDR_VERD,True,"FFFFFF",9)
        div_tot = delta_tot != 0
        ec(ws_mes,ri_tot+1,3,f"App={app_tot}", F_VERM_CLR if div_tot else F_VERDE_MED,True,"000000",9)
        ec(ws_mes,ri_tot+1,4,f"Excel={xl_tot_v}", F_VERM_CLR if div_tot else F_VERDE_MED,True,"000000",9)
        ec(ws_mes,ri_tot+1,5,f"Δ={delta_tot:+d}", F_VERM_CLR if div_tot else F_VERDE_MED,True,"000000",9)

        lab_app = r_app["prod_labor_tot"]; lab_xl = xl_tot.get("labor",0)
        ec(ws_mes,ri_tot+2,3,f"Labor App={lab_app:.1%}",F_AZUL_CLR,True,"000000",9)
        ec(ws_mes,ri_tot+2,4,f"Labor Excel={lab_xl:.1%}",F_AZUL_CLR,True,"000000",9)
        delta_lab = lab_app - lab_xl
        ec(ws_mes,ri_tot+2,5,f"Δ={delta_lab:+.1%}",F_AMAR if abs(delta_lab)>0.02 else F_AZUL_CLR,True,"000000",9)
        ws_mes.row_dimensions[ri_tot].height=16; ws_mes.row_dimensions[ri_tot+1].height=16; ws_mes.row_dimensions[ri_tot+2].height=16

        # Resumo para capa
        status = "✅ Igual" if n_div_mes==0 else ("🟡 Pequena" if n_div_mes<=2 else "🔴 Divergente")
        resumo_capa.append((mes, status, f"{delta_tot:+d}", n_div_mes))

    # Preencher capa com índice
    for ri, (mes, status, delta, n_div) in enumerate(resumo_capa, 11):
        fill = F_VERDE if "✅" in status else (F_AMAR if "🟡" in status else F_VERM_CLR)
        ec(ws_capa,ri,2,mes,fill,False,"000000",9,False)
        ec(ws_capa,ri,3,status,fill,True,"000000",9)
        ec(ws_capa,ri,4,delta,fill,True,"000000",9)
        ec(ws_capa,ri,5,n_div,fill,True,"000000",9)
        ws_capa.row_dimensions[ri].height=15

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def gerar_output_layout(file_bytes, res_app, tempo, dist, aplic, pmp, dias, horas_turno, thresholds, suporte_cfg):
    """Gera Excel no mesmo layout do original, com células em vermelho onde diverge."""
    MAPA = {
        "Novembro":"NovFY26","Dezembro":"DezFY26","Janeiro":"JanFY26",
        "Fevereiro":"FevFY26","Março":"MarFY26","Abril":"AbrFY26",
        "Maio":"MaiFY26","Junho":"JunFY26","Julho":"JulFY26",
        "Agosto":"AgoFY26","Setembro":"SetFY26","Outubro":"OutFY26",
    }
    hA=horas_turno["A"]; hB=horas_turno["B"]; hC=horas_turno["C"]
    thr_A=thresholds["A"]/100; thr_B=thresholds["B"]/100; thr_C=thresholds["C"]/100

    # Calcular min_ciclo por centro×mes
    df_all = (aplic.merge(pmp, on="modelo")
                   .merge(tempo, on=["centro","peca"])
                   .merge(dist,  on=["centro","peca"]))
    if "vol_int" not in df_all.columns: df_all["vol_int"] = 1.0
    df_all["vol_int"] = pd.to_numeric(df_all["vol_int"], errors="coerce").fillna(1.0)
    df_all["indice_ciclo"] = (df_all.t_ciclo * df_all.div_carga * df_all.div_volume * df_all.vol_int) / df_all.disponib
    df_all["min_ciclo"]    = df_all.indice_ciclo * df_all.qtd
    df_all["min_labor"]    = df_all.t_labor * df_all.div_carga * df_all.qtd
    agg_all = df_all.groupby(["centro","mes"])[["min_ciclo","min_labor"]].sum().reset_index()

    # Ler Excel de referência (rápido — só totais e ocupação por centro)
    wb_ref = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    abas_ref = wb_ref.sheetnames
    dados_xl = {}
    for mes, aba in MAPA.items():
        if aba not in abas_ref: continue
        ws_ref = wb_ref[aba]
        centros_xl = {}
        for row in ws_ref.iter_rows(min_row=69, max_row=88, min_col=23, max_col=29, values_only=True):
            cen = row[0]
            if not cen: continue
            centros_xl[str(cen).strip()] = {
                "ocup_A": safe_float(row[1]), "ocup_B": safe_float(row[2]),
                "ativo_A": safe_int(row[4]), "ativo_B": safe_int(row[5]), "ativo_C": safe_int(row[6])
            }
        dados_xl[mes] = {
            "centros": centros_xl,
            "tot_A": safe_int(ws_ref.cell(95,27).value),
            "tot_B": safe_int(ws_ref.cell(95,28).value),
            "tot_C": safe_int(ws_ref.cell(95,29).value),
            "total": safe_int(ws_ref.cell(96,27).value),
            "labor_tot": safe_float(ws_ref.cell(101,30).value),
            "ciclo_op":  safe_float(ws_ref.cell(98,30).value),
            "ciclo_tot": safe_float(ws_ref.cell(99,30).value),
            "labor_op":  safe_float(ws_ref.cell(100,30).value),
        }
    wb_ref.close()

    wb = openpyxl.Workbook()
    primeira = True

    for mes in MESES:
        r_app = res_app.get(mes)
        d     = dias.get(mes, 0)
        if not r_app or d == 0: continue

        xl = dados_xl.get(mes, {})
        centros_xl = xl.get("centros", {})

        minA=d*hA*60; minB=d*hB*60; minC=d*hC*60
        agg_mes = agg_all[agg_all.mes==mes].copy()

        # Calcular por centro
        app_centros = {}
        for _, row in agg_mes.iterrows():
            mc=row.min_ciclo; ml=row.min_labor
            oA=mc/minA if minA>0 else 0
            oB=mc/minB if minB>0 else 0
            oC=mc/minC if minC>0 else 0
            aA=1 if oA>thr_A else 0
            aB=1 if oA>thr_B else 0
            aC=1 if oB>thr_C else 0
            app_centros[row.centro] = {
                "ocup_A":oA,"ocup_B":oB,"ocup_C":oC,
                "ativo_A":aA,"ativo_B":aB,"ativo_C":aC,
                "hA":d*hA*aA,"hB":d*hB*aB,"hC":d*hC*aC,
                "min_ciclo":mc,"min_labor":ml
            }

        # Ordenar centros igual ao Excel (pela ordem que aparecem no agg)
        ordem_centros = list(agg_mes.centro)

        # ── Criar aba ──────────────────────────────────────────────
        if primeira:
            ws = wb.active; ws.title = mes[:10]; primeira = False
        else:
            ws = wb.create_sheet(mes[:10])

        # Larguras das colunas (igual ao Excel original)
        col_widths = {1:5,2:9,3:16,4:5,5:5,6:5,
                      7:9,8:9,9:9,10:9,11:9,12:9,
                      13:9,14:9,15:9,16:5}
        for c,w in col_widths.items():
            ws.column_dimensions[get_column_letter(c)].width = w

        # ── LINHA 1 — Cabeçalho "DADOS AUTOMÁTICOS" ───────────────
        ws.merge_cells("A1:P1")
        ec(ws,1,1,"DADOS AUTOMÁTICOS", C_CINZA, True, "000000", 10, True)
        ws.row_dimensions[1].height = 16

        # ── LINHA 2 — Período e horas ─────────────────────────────
        ec(ws,2,1,"PERÍODO:", C_CINZA, True, "000000", 8, False)
        ws.merge_cells("B2:C2")
        ec(ws,2,2,mes, C_BRANCO, True, "FF0000", 9, False)
        ec(ws,2,5,"DATA DE REVISÃO:", C_CINZA, True, "000000", 8, False)
        ws.merge_cells("G2:H2")
        ec(ws,2,7,datetime.now().strftime("%d/%m/%Y"), C_BRANCO, True, "FF0000", 9)
        ec(ws,2,10,"HORAS POR TURNO DE TRABALHO", C_CINZA, True, "000000", 8, True)
        ws.row_dimensions[2].height = 14

        # ── LINHA 3 — Valores das horas ───────────────────────────
        for ci,h in [(10,hA),(13,hB),(16,hC)]:
            ec(ws,3,ci,h,C_CINZA,True,"000000",8)
        ws.row_dimensions[3].height = 14

        # ── LINHA 4 — Cabeçalho dos grupos ───────────────────────
        for ci, txt, fill in [
            (1,"",C_CINZA),(2,"",C_CINZA),(3,"",C_CINZA),
            (4,"TURNO A",C_VERDE),(5,"TURNO B",C_AMAR),(6,"TURNO C",C_AZUL),
            (7,"TURNO A",C_VERDE),(8,"TURNO B",C_AMAR),(9,"TURNO C",C_AZUL),
            (10,"TURNO A",C_VERDE),(11,"TURNO B",C_AMAR),(12,"TURNO C",C_AZUL),
        ]:
            ec(ws,4,ci,txt,fill,True,"000000",8)
        ws.row_dimensions[4].height = 14

        # ── LINHA 5 — Sub-cabeçalho ───────────────────────────────
        ec(ws,5,1,"Centro",C_PRETO,True,"FFFFFF",8)
        ec(ws,5,2,"",C_PRETO,True,"FFFFFF",8)
        ec(ws,5,3,"",C_PRETO,True,"FFFFFF",8)
        for ci,txt in [(4,"% Ocup"),(5,"% Ocup"),(6,"% Ocup"),
                        (7,"Ativo"),(8,"Ativo"),(9,"Ativo"),
                        (10,"Horas"),(11,"Horas"),(12,"Horas")]:
            ec(ws,5,ci,txt,C_PRETO,True,"FFFFFF",8)
        ws.row_dimensions[5].height = 14

        # ── DADOS por centro ───────────────────────────────────────
        ri = 6
        for cen in ordem_centros:
            ap = app_centros.get(cen, {})
            xl_c = centros_xl.get(cen, {})

            oA_ap=ap.get("ocup_A",0); oA_xl=xl_c.get("ocup_A",0)
            oB_ap=ap.get("ocup_B",0); oB_xl=xl_c.get("ocup_B",0)
            oC_ap=ap.get("ocup_C",0)
            aA_ap=ap.get("ativo_A",0); aA_xl=xl_c.get("ativo_A",0)
            aB_ap=ap.get("ativo_B",0); aB_xl=xl_c.get("ativo_B",0)
            aC_ap=ap.get("ativo_C",0); aC_xl=xl_c.get("ativo_C",0)
            hA_ap=ap.get("hA",0); hB_ap=ap.get("hB",0); hC_ap=ap.get("hC",0)

            div_A = aA_ap != aA_xl
            div_B = aB_ap != aB_xl
            div_C = aC_ap != aC_xl

            # Cor do nome do centro
            cen_fill = C_VERM_DIV if (div_A or div_B or div_C) else C_BRANCO
            cen_color = "FFFFFF" if (div_A or div_B or div_C) else "000000"

            ec(ws,ri,1,cen,cen_fill,True,cen_color,8,False)
            ec(ws,ri,2,"",C_BRANCO); ec(ws,ri,3,"",C_BRANCO)

            # % Ocupação — cor igual ao Excel + vermelho se diverge
            def fill_ocup(pct, div):
                if div: return C_VERM_DIV
                return cor_ocup(pct)

            ec(ws,ri,4,f"{oA_ap:.0%}", fill_ocup(oA_ap, div_A and abs(oA_ap-oA_xl)>0.01))
            ec(ws,ri,5,f"{oB_ap:.0%}", fill_ocup(oB_ap, div_B and abs(oB_ap-oB_xl)>0.01))
            ec(ws,ri,6,f"{oC_ap:.0%}", cor_ocup(oC_ap))

            # Ativo — vermelho se diverge, verde/amarelo/azul se igual
            ec(ws,ri,7,aA_ap, C_VERM_DIV if div_A else (C_VERDE if aA_ap else C_BRANCO), True)
            ec(ws,ri,8,aB_ap, C_VERM_DIV if div_B else (C_AMAR  if aB_ap else C_BRANCO), True)
            ec(ws,ri,9,aC_ap, C_VERM_DIV if div_C else (C_AZUL  if aC_ap else C_BRANCO), True)

            # Horas
            ec(ws,ri,10,round(hA_ap,2) if hA_ap else 0, C_VERDE if hA_ap else C_BRANCO)
            ec(ws,ri,11,round(hB_ap,2) if hB_ap else 0, C_AMAR  if hB_ap else C_BRANCO)
            ec(ws,ri,12,round(hC_ap,2) if hC_ap else 0, C_AZUL  if hC_ap else C_BRANCO)

            ws.row_dimensions[ri].height = 13
            ri += 1

        # ── TOTAIS DE OPERADORES ───────────────────────────────────
        ri += 1  # linha em branco
        xl_tot = dados_xl.get(mes, {})

        def cel_tot(row, col, app_v, xl_v, fill_base):
            div = int(app_v) != safe_int(xl_v)
            f = C_VERM_DIV if div else fill_base
            ec(ws, row, col, app_v, f, True, "000000", 8)
            if div:
                # Adicionar valor do Excel ao lado
                ec(ws, row, col+3, f"(Excel:{safe_int(xl_v)})", C_VERM_DIV, False, "FFFFFF", 7)

        ws.merge_cells(f"A{ri}:C{ri}")
        ec(ws,ri,1,"TOTAL DE OPERADORES",C_ROSA,True,"000000",8,False)
        cel_tot(ri,7, r_app["op_A"], xl_tot.get("tot_A",0)-6, C_VERDE)  # descontando suporte
        cel_tot(ri,8, r_app["op_B"], xl_tot.get("tot_B",0)-4, C_AMAR)
        cel_tot(ri,9, r_app["op_C"], xl_tot.get("tot_C",0)-1, C_AZUL)
        ws.row_dimensions[ri].height = 14; ri += 1

        # Suportes
        sup = r_app["suporte"]
        for nome, key in [("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUIEDADE","gravacao"),
                          ("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador")]:
            ws.merge_cells(f"A{ri}:C{ri}")
            ec(ws,ri,1,nome,C_BRANCO,False,"000000",8,False)
            s=sup[key]
            ec(ws,ri,7,s["A"],C_VERDE if s["A"] else C_BRANCO,True)
            ec(ws,ri,8,s["B"],C_AMAR  if s["B"] else C_BRANCO,True)
            ec(ws,ri,9,s["C"],C_AZUL  if s["C"] else C_BRANCO,True)
            ec(ws,ri,10,round(s["A"]*d*hA,2) if s["A"] else 0, C_VERDE if s["A"] else C_BRANCO)
            ec(ws,ri,11,round(s["B"]*d*hB,2) if s["B"] else 0, C_AMAR  if s["B"] else C_BRANCO)
            ec(ws,ri,12,round(s["C"]*d*hC,2) if s["C"] else 0, C_AZUL  if s["C"] else C_BRANCO)
            ws.row_dimensions[ri].height = 13; ri += 1

        # TOTAL POR TURNO
        ws.merge_cells(f"A{ri}:C{ri}")
        ec(ws,ri,1,"TOTAL POR TURNO",C_ROSA,True,"000000",8,False)
        for col, app_v, xl_v, fill in [
            (7, r_app["tot_A"], xl_tot.get("tot_A",0), C_VERDE),
            (8, r_app["tot_B"], xl_tot.get("tot_B",0), C_AMAR),
            (9, r_app["tot_C"], xl_tot.get("tot_C",0), C_AZUL),
            (10, round(r_app["tot_A"]*d*hA,2), 0, C_VERDE),
            (11, round(r_app["tot_B"]*d*hB,2), 0, C_AMAR),
            (12, round(r_app["tot_C"]*d*hC,2), 0, C_AZUL),
        ]:
            div = xl_v and (int(app_v) != int(xl_v))
            ec(ws,ri,col,app_v, C_VERM_DIV if div else fill, True, "FFFFFF" if div else "000000", 8)
        ws.row_dimensions[ri].height = 14; ri += 1

        # TOTAL FUNCIONÁRIOS
        ws.merge_cells(f"A{ri}:C{ri}")
        ec(ws,ri,1,"TOTAL FUNCIONÁRIOS",C_ROSA,True,"000000",8,False)
        div_tot = r_app["total"] != xl_tot.get("total",0)
        ec(ws,ri,7,r_app["total"], C_VERM_DIV if div_tot else C_ROSA, True,
           "FFFFFF" if div_tot else "000000", 9)
        if div_tot:
            ec(ws,ri,8,f"Excel: {xl_tot.get('total',0)}", C_VERM_DIV, True, "FFFFFF", 8)
        ws.row_dimensions[ri].height = 16; ri += 2

        # ── PRODUTIVIDADES ─────────────────────────────────────────
        prods = [
            ("PRODUTIVIDADE POR TEMPO DE CICLO OPERACIONAL", r_app["prod_ciclo_op"],  xl_tot.get("ciclo_op",0),  False),
            ("PRODUTIVIDADE POR TEMPO DE CICLO TOTAL",       r_app["prod_ciclo_tot"], xl_tot.get("ciclo_tot",0), False),
            ("PRODUTIVIDADE POR TEMPO DE LABOR OPERACIONAL", r_app["prod_labor_op"],  xl_tot.get("labor_op",0),  False),
            ("PRODUTIVIDADE POR TEMPO DE LABOR TOTAL ★",     r_app["prod_labor_tot"], xl_tot.get("labor_tot",0), True),
        ]
        for nome, app_v, xl_v, destaque in prods:
            ws.merge_cells(f"A{ri}:K{ri}")
            fill_p = C_AMAR if destaque else C_BRANCO
            ec(ws,ri,1,nome, fill_p, destaque, "000000", 8, False)
            div_p = xl_v and abs(app_v - xl_v) > 0.005
            ec(ws,ri,12,f"{app_v:.1%}", C_VERM_DIV if div_p else fill_p, True,
               "FFFFFF" if div_p else "000000", 8)
            if div_p:
                ec(ws,ri,13,f"(Excel:{xl_v:.1%})", C_VERM_DIV, True, "FFFFFF", 7)
            ws.row_dimensions[ri].height = 14; ri += 1

        # Nota de divergências
        n_div = sum(1 for cen in ordem_centros
                    for a,b in [(app_centros.get(cen,{}).get("ativo_A",0), centros_xl.get(cen,{}).get("ativo_A",0)),
                                (app_centros.get(cen,{}).get("ativo_B",0), centros_xl.get(cen,{}).get("ativo_B",0)),
                                (app_centros.get(cen,{}).get("ativo_C",0), centros_xl.get(cen,{}).get("ativo_C",0))]
                    if a != b)
        if n_div > 0:
            ws.merge_cells(f"A{ri}:P{ri}")
            ec(ws,ri,1, f"⚠️ {n_div} divergência(s) em vermelho — verifique IMPUTDISTRIBUIÇÃO ou IMPUTAPLICAÇÃO",
               C_VERM_DIV, True, "FFFFFF", 8, False)
            ws.row_dimensions[ri].height = 14

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────
# TABELONA COMPLETA (extraído para função — chamado só no botão)
# ─────────────────────────────────────────
def _gerar_tabelona(file_bytes, res_base, tempo, dist, aplic, pmp, dias, horas_turno, thresholds, suporte_cfg):
    """Gera a tabelona completa: layout IMPUTDISTRIBUIÇÃO + DADOS AUTOMÁTICOS por mês."""
    import openpyxl as _opx
    from openpyxl.styles import PatternFill as _PF, Font as _Ft, Alignment as _Al, Border as _Bd, Side as _Sd

    MAPA_T = {"Novembro":"NovFY26","Dezembro":"DezFY26","Janeiro":"JanFY26",
              "Fevereiro":"FevFY26","Março":"MarFY26","Abril":"AbrFY26",
              "Maio":"MaiFY26","Junho":"JunFY26","Julho":"JulFY26",
              "Agosto":"AgoFY26","Setembro":"SetFY26","Outubro":"OutFY26"}

    _F_VERDE  = _PF("solid",fgColor="66FF66"); _F_AMAR   = _PF("solid",fgColor="FFFF00")
    _F_AZUL   = _PF("solid",fgColor="00B0F0"); _F_PRETO  = _PF("solid",fgColor="000000")
    _F_CINZA  = _PF("solid",fgColor="D9D9D9"); _F_CINZA2 = _PF("solid",fgColor="BFBFBF")
    _F_BRANCO = _PF("solid",fgColor="FFFFFF"); _F_VERM   = _PF("solid",fgColor="FF0000")
    _F_VERM_S = _PF("solid",fgColor="FFCCCC"); _F_VERDE_JD = _PF("solid",fgColor="1F4D19")
    _BRD = _Bd(left=_Sd(style="thin",color="AAAAAA"),right=_Sd(style="thin",color="AAAAAA"),
               top=_Sd(style="thin",color="AAAAAA"),bottom=_Sd(style="thin",color="AAAAAA"))

    def _ec(ws,r,c,val,fill=None,bold=False,color="000000",size=8,center=True,wrap=False):
        cell=ws.cell(row=r,column=c,value=val)
        cell.font=_Ft(name="Arial",bold=bold,color=color,size=size)
        cell.fill=fill or _F_BRANCO
        cell.alignment=_Al(horizontal="center" if center else "left",vertical="center",wrap_text=wrap)
        cell.border=_BRD
        return cell

    def _cor_pct(v):
        try:
            f=float(v)
            if f>=1.06: return _PF("solid",fgColor="FF0000")
            if f>=1.00: return _F_AMAR
            if f>=0.40: return _PF("solid",fgColor="92D050")
            return _F_BRANCO
        except: return _F_BRANCO

    hA_t=horas_turno["A"]; hB_t=horas_turno["B"]; hC_t=horas_turno["C"]
    thr_A_t=thresholds["A"]/100; thr_B_t=thresholds["B"]/100; thr_C_t=thresholds["C"]/100

    # JOIN único para todos os meses
    df_all_t = (aplic.merge(pmp,on="modelo").merge(tempo,on=["centro","peca"]).merge(dist,on=["centro","peca"]))
    if "vol_int" not in df_all_t.columns: df_all_t["vol_int"]=1.0
    df_all_t["vol_int"] = pd.to_numeric(df_all_t["vol_int"],errors="coerce").fillna(1.0)
    df_all_t["indice_ciclo"]=(df_all_t.t_ciclo*df_all_t.div_carga*df_all_t.div_volume*df_all_t.vol_int)/df_all_t.disponib
    df_all_t["min_ciclo"]=df_all_t.indice_ciclo*df_all_t.qtd
    df_all_t["min_labor"]=df_all_t.t_labor*df_all_t.div_carga*df_all_t.qtd
    agg_cp_t = df_all_t.groupby(["centro","peca","mes"])[["min_ciclo","min_labor"]].sum()

    # Ler Excel de referência UMA ÚNICA VEZ
    wb_r = _opx.load_workbook(BytesIO(file_bytes),read_only=True,data_only=True)
    # Ler aba base (NovFY26 ou primeira disponível)
    aba_base = next((a for a in ["NovFY26","DezFY26","JanFY26"] if a in wb_r.sheetnames),None)
    if not aba_base:
        wb_r.close()
        return BytesIO()
    ws_base = wb_r[aba_base]
    base_rows_t = list(ws_base.iter_rows(min_row=7,max_row=63,min_col=1,max_col=87,values_only=True))
    base_rows_t = [r for r in base_rows_t if r[0] and r[1]]
    modelos_xl_t = [str(ws_base.cell(6,c).value) for c in range(19,88)
                    if ws_base.cell(6,c).value and str(ws_base.cell(6,c).value).startswith("MODELO")]
    modelo_col_idx = {str(ws_base.cell(6,c).value):(c-19) for c in range(19,88)
                      if ws_base.cell(6,c).value and str(ws_base.cell(6,c).value).startswith("MODELO")}

    # Ler dados mensais de uma vez
    dados_mes_t = {}
    for mes_t, aba_t in MAPA_T.items():
        if aba_t not in wb_r.sheetnames: continue
        ws_m_t = wb_r[aba_t]
        dados_mes_t[mes_t] = {
            "main": list(ws_m_t.iter_rows(min_row=7,max_row=63,min_col=1,max_col=18,values_only=True)),
            "vols": list(ws_m_t.iter_rows(min_row=7,max_row=63,min_col=19,max_col=87,values_only=True))
        }
    wb_r.close()

    aplic_orig = pd.read_excel(BytesIO(file_bytes),sheet_name="IMPUTAPLICAÇÃO",header=0)
    aplic_orig = aplic_orig.rename(columns={aplic_orig.columns[0]:"centro",aplic_orig.columns[1]:"peca"})

    wb_out = _opx.Workbook(); primeira_t = True

    from datetime import date as _date

    for mes_t in MESES:
        d_t = dias.get(mes_t,0)
        if d_t == 0: continue
        minA_t=d_t*hA_t*60; minB_t=d_t*hB_t*60; minC_t=d_t*hC_t*60
        dm_t = dados_mes_t.get(mes_t,{}); pmp_mes_t = pmp[pmp.mes==mes_t]

        if primeira_t: ws_out=wb_out.active; ws_out.title=mes_t[:10]; primeira_t=False
        else: ws_out=wb_out.create_sheet(mes_t[:10])
        ws_out.freeze_panes="F7"

        # Larguras das colunas fixas
        col_widths={1:5,2:9,3:16,4:6,5:5,6:9,7:9,8:8,9:8,10:8,11:8,12:9,13:9,14:9,15:9,16:12,17:12,18:8}
        for c,w in col_widths.items(): ws_out.column_dimensions[get_column_letter(c)].width=w

        # Linha 5 — título
        ws_out.merge_cells(f"A5:{get_column_letter(18+len(modelos_xl_t))}5")
        _ec(ws_out,5,1,f"RESUMO DA CARGA — {mes_t.upper()} ({d_t} dias)",_F_VERDE_JD,True,"FFFFFF",10,True)
        ws_out.row_dimensions[5].height=18

        # Linha 6 — cabeçalhos
        hdrs_f=[("Máquina",_F_CINZA2,"000000"),("PEÇA",_F_CINZA2,"000000"),("DESCRIÇÃO",_F_CINZA2,"000000"),
                ("PÇ/TRAT",_F_CINZA2,"000000"),("UM",_F_CINZA2,"000000"),
                ("Tempo Ciclo (min)",_F_PRETO,"FFFFFF"),("Tempo Labor (min)",_F_PRETO,"FFFFFF"),
                ("Div. Carga",_PF("solid",fgColor="FF0000"),"FFFF00"),("Vol. Interna",_F_CINZA2,"000000"),
                ("Div. Volume",_PF("solid",fgColor="FF0000"),"FFFF00"),("Disponib.",_F_CINZA2,"000000"),
                ("Indice Ciclo",_F_CINZA2,"000000"),
                ("JA.A",_F_VERDE,"000000"),("JA.B",_F_AMAR,"000000"),("JA.C",_F_AZUL,"000000"),
                ("TOTAL CICLOS (MIN)",_F_CINZA,"000000"),("TOTAL LABOR (MIN)",_F_CINZA,"000000"),
                ("TOTAL PECAS",_F_CINZA,"000000")]
        for ci_t,(h_t,f_t,cor_t) in enumerate(hdrs_f,1):
            _ec(ws_out,6,ci_t,h_t,f_t,True,cor_t,8,True,True)
        for mi_t,mod_t in enumerate(modelos_xl_t):
            ci_t=19+mi_t
            _ec(ws_out,6,ci_t,mod_t,_F_CINZA,True,"000000",7,True,True)
            ws_out.column_dimensions[get_column_letter(ci_t)].width=7
        ws_out.row_dimensions[6].height=42

        # Linhas de dados
        main_data_t=dm_t.get("main",[]); vols_data_t=dm_t.get("vols",[])
        for ri_t_idx,base_row_t in enumerate(base_rows_t):
            cen_t=str(base_row_t[0]).strip(); peca_t=str(base_row_t[1]).strip()
            ri_t=7+ri_t_idx
            tc_xl_t=base_row_t[5]; tl_xl_t=base_row_t[6]
            dc_xl_t=base_row_t[7]; vi_xl_t=base_row_t[8]; dv_xl_t=base_row_t[9]
            di_xl_t=base_row_t[10]; idx_xl_t=base_row_t[11]
            mrow_t=main_data_t[ri_t_idx] if ri_t_idx<len(main_data_t) else [None]*18
            xl_pA_t=mrow_t[12] if len(mrow_t)>12 else None
            xl_pB_t=mrow_t[13] if len(mrow_t)>13 else None
            xl_ciclo_t=mrow_t[15] if len(mrow_t)>15 else None
            xl_pecas_t=mrow_t[17] if len(mrow_t)>17 else None
            vrow_t=vols_data_t[ri_t_idx] if vols_data_t and ri_t_idx<len(vols_data_t) else []
            try: mc_t=float(agg_cp_t.loc[(cen_t,peca_t,mes_t),"min_ciclo"])
            except: mc_t=0.0
            try: ml_t=float(agg_cp_t.loc[(cen_t,peca_t,mes_t),"min_labor"])
            except: ml_t=0.0
            pA_t=mc_t/minA_t if minA_t>0 else 0
            pB_t=mc_t/minB_t if minB_t>0 else 0
            pC_t=mc_t/minC_t if minC_t>0 else 0
            app_mod_v={}
            for mod_t2 in modelos_xl_t:
                qtd_t=int(pmp_mes_t[pmp_mes_t.modelo==mod_t2]["qtd"].sum()) if mod_t2 in pmp_mes_t.modelo.values else 0
                fr_t=aplic_orig[(aplic_orig.centro==cen_t)&(aplic_orig.peca==peca_t)]
                flag_t=int(fr_t[mod_t2].values[0]) if len(fr_t)>0 and mod_t2 in fr_t.columns and not pd.isna(fr_t[mod_t2].values[0] if len(fr_t)>0 else 0) else 0
                app_mod_v[mod_t2]=qtd_t*flag_t
            app_tot_t=sum(app_mod_v.values())
            def _df(a,b,tol=0.02):
                if b is None: return False
                try: return abs(float(a or 0)-float(b))>tol
                except: return False
            div_A_t=_df(pA_t,xl_pA_t,0.02); div_B_t=_df(pB_t,xl_pB_t,0.02)
            div_c_t=_df(mc_t,xl_ciclo_t,1); div_p_t=_df(app_tot_t,xl_pecas_t,0.5)
            dc_i=dist[(dist.centro==cen_t)&(dist.peca==peca_t)]["div_carga"].values
            vi_i=dist[(dist.centro==cen_t)&(dist.peca==peca_t)]["vol_int"].values
            dv_i=dist[(dist.centro==cen_t)&(dist.peca==peca_t)]["div_volume"].values
            di_i=dist[(dist.centro==cen_t)&(dist.peca==peca_t)]["disponib"].values
            vi_val=float(vi_i[0]) if len(vi_i) else 1.0
            idx_app_t=(float(tc_xl_t or 0)*dc_i[0]*dv_i[0]*vi_val)/di_i[0] if len(dc_i) and len(di_i) and di_i[0] else float(idx_xl_t or 0)
            div_idx_t=abs(float(idx_xl_t or 0)-float(idx_app_t or 0))>0.5
            _ec(ws_out,ri_t,1,cen_t,_F_BRANCO,False,"000000",8,False)
            _ec(ws_out,ri_t,2,peca_t,_F_BRANCO,False,"000000",8,False)
            _ec(ws_out,ri_t,3,base_row_t[2],_F_BRANCO,False,"000000",8,False)
            _ec(ws_out,ri_t,4,base_row_t[3],_F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,5,base_row_t[4],_F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,6,tc_xl_t,_F_PRETO,False,"FFFFFF",8)
            _ec(ws_out,ri_t,7,tl_xl_t,_F_PRETO,False,"FFFFFF",8)
            _ec(ws_out,ri_t,8,dc_xl_t,_PF("solid",fgColor="FF0000"),False,"FFFF00",8)
            _ec(ws_out,ri_t,9,vi_xl_t,_F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,10,dv_xl_t,_PF("solid",fgColor="FF0000"),False,"FFFF00",8)
            _ec(ws_out,ri_t,11,di_xl_t,_F_CINZA2,False,"000000",8)
            _ec(ws_out,ri_t,12,round(float(idx_app_t),4),_F_VERM_S if div_idx_t else _F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,13,f"{pA_t:.1%}",_F_VERM if div_A_t else _cor_pct(pA_t),False,"000000",8)
            _ec(ws_out,ri_t,14,f"{pB_t:.1%}",_F_VERM if div_B_t else _cor_pct(pB_t),False,"000000",8)
            _ec(ws_out,ri_t,15,f"{pC_t:.1%}",_cor_pct(pC_t),False,"000000",8)
            _ec(ws_out,ri_t,16,round(mc_t,1),_F_VERM_S if div_c_t else _F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,17,round(ml_t,1),_F_BRANCO,False,"000000",8)
            _ec(ws_out,ri_t,18,app_tot_t,_F_VERM_S if div_p_t else _F_BRANCO,False,"000000",8)
            for mi_t2,mod_t2 in enumerate(modelos_xl_t):
                ci_t2=19+mi_t2
                v_app_t=app_mod_v.get(mod_t2,0)
                col_idx=modelo_col_idx.get(mod_t2,mi_t2)
                v_xl_t=vrow_t[col_idx] if vrow_t and col_idx<len(vrow_t) else None
                if v_xl_t is not None:
                    try: div_vm=abs(float(v_app_t)-float(v_xl_t or 0))>0.5
                    except: div_vm=False
                    fill_vm=_F_VERM if div_vm else (_F_CINZA if v_app_t else _F_BRANCO)
                else: fill_vm=_F_CINZA if v_app_t else _F_BRANCO
                _ec(ws_out,ri_t,ci_t2,v_app_t if v_app_t else None,fill_vm,False,"000000",7)
            ws_out.row_dimensions[ri_t].height=13

        # ── Linha de nota ─────────────────────────────────────────────────────
        nota_rt=7+len(base_rows_t)+1
        ws_out.merge_cells(f"A{nota_rt}:{get_column_letter(18+len(modelos_xl_t))}{nota_rt}")
        nt=ws_out.cell(nota_rt,1,"🔴 JA.A/JA.B vermelho = % ocupação difere do Excel  |  🔴 Rosa = total ciclos ou peças difere  |  Cinza = presente no App")
        nt.font=_Ft(name="Arial",bold=True,size=8,color="CC0000")
        nt.fill=_PF("solid",fgColor="FFEEEE")
        nt.alignment=_Al(horizontal="left",vertical="center")
        ws_out.row_dimensions[nota_rt].height=14

        # ════════════════════════════════════════════════════════════════
        # BLOCO DADOS AUTOMÁTICOS — coluna L (12) em diante
        # ════════════════════════════════════════════════════════════════
        res_t=res_base.get(mes_t)
        if res_t:
            d_t2=res_t["dias"]; hA_v=res_t["hA"]; hB_v=res_t["hB"]; hC_v=res_t["hC"]
            sup_t=res_t["suporte"]
            _F_HDR_VERD=_PF("solid",fgColor="1F4D19"); _F_HDR_AMAR=_PF("solid",fgColor="FFDE00")
            _F_HDR_GRAY=_PF("solid",fgColor="D9D9D9"); _F_VERM_L=_PF("solid",fgColor="FFCDD2")
            _F_AMAR_L=_PF("solid",fgColor="FFFF00"); _F_VERDE_L=_PF("solid",fgColor="92D050")
            _F_AZUL_L=_PF("solid",fgColor="B3E5FC"); _F_CINZA_L=_PF("solid",fgColor="F4F4F4")
            _F_AMAR_TOT=_PF("solid",fgColor="FFDE00")
            COL_OFF=12; da_row=nota_rt+2

            def _da(r,c,val,fill=None,bold=False,color="000000",size=9,center=True,merge_end=None):
                cell=ws_out.cell(row=r,column=c,value=val)
                cell.font=_Ft(name="Arial",bold=bold,color=color,size=size)
                cell.fill=fill or _F_BRANCO
                cell.alignment=_Al(horizontal="center" if center else "left",vertical="center",wrap_text=True)
                cell.border=_BRD
                if merge_end: ws_out.merge_cells(start_row=r,start_column=c,end_row=r,end_column=merge_end)
                return cell

            # TOTAL DE MINUTOS / HORAS / DIAS / HORAS POR TURNO
            min_A=d_t2*hA_v*60; min_B=d_t2*hB_v*60; min_C=d_t2*hC_v*60
            for ri_hm,(label,vA,vB,vC) in enumerate([
                ("TOTAL DE MINUTOS",round(min_A,1),round(min_B,1),round(min_C,1)),
                ("TOTAL DE HORAS",  round(d_t2*hA_v,2),round(d_t2*hB_v,2),round(d_t2*hC_v,2)),
                ("DIAS TRABALHADOS",d_t2,d_t2,d_t2),
                ("HORAS POR TURNO", hA_v,hB_v,hC_v),
            ]):
                r_hm=da_row+ri_hm
                fill_l=_F_VERM_L if "DIAS" in label else _F_HDR_GRAY
                _da(r_hm,COL_OFF,label,fill_l,bold=True,center=False)
                _da(r_hm,COL_OFF+1,vA,fill_l,bold=True)
                _da(r_hm,COL_OFF+2,vB,fill_l,bold=True)
                _da(r_hm,COL_OFF+3,vC,fill_l,bold=True)
                ws_out.row_dimensions[r_hm].height=14

            # Cabeçalho RESUMO DA CARGA
            r_titulo=da_row+5
            _da(r_titulo,COL_OFF,"RESUMO DA CARGA MÁQUINA X QUADRO DE LOTAÇÃO",
                _F_HDR_GRAY,bold=True,center=True,merge_end=COL_OFF+9)
            ws_out.row_dimensions[r_titulo].height=16

            # PERÍODO / DATA / HORAS POR TURNO
            r_per=r_titulo+1
            _da(r_per,COL_OFF,"PERÍODO:",_F_HDR_GRAY,bold=True,center=False)
            _da(r_per,COL_OFF+1,mes_t,_PF("solid",fgColor="FF0000"),bold=True,color="FFFFFF")
            _da(r_per,COL_OFF+3,"DATA DE REVISÃO:",_F_HDR_GRAY,bold=True,center=False)
            _da(r_per,COL_OFF+5,_date.today().strftime("%d/%m/%Y"),_PF("solid",fgColor="FF0000"),bold=True,color="FFFFFF")
            _da(r_per,COL_OFF+7,"HORAS POR TURNO DE TRABALHO",_F_HDR_GRAY,bold=True,merge_end=COL_OFF+9)
            ws_out.row_dimensions[r_per].height=14

            # Horas por turno
            r_h=r_per+1
            _da(r_h,COL_OFF+7,hA_v,_F_HDR_GRAY,bold=True)
            _da(r_h,COL_OFF+8,hB_v,_F_HDR_GRAY,bold=True)
            _da(r_h,COL_OFF+9,hC_v,_F_HDR_GRAY,bold=True)
            ws_out.row_dimensions[r_h].height=14

            # Cabeçalho Turnos
            r_hdr=r_h+1
            for ci_da,txt_da,fill_da in [
                (COL_OFF,"",_F_HDR_VERD),(COL_OFF+1,"TURNO A",_F_VERDE),(COL_OFF+2,"TURNO B",_F_AMAR),
                (COL_OFF+3,"TURNO C",_F_AZUL),(COL_OFF+4,"TURNO A",_F_VERDE),(COL_OFF+5,"TURNO B",_F_AMAR),
                (COL_OFF+6,"TURNO C",_F_AZUL),(COL_OFF+7,"TURNO A",_F_VERDE),(COL_OFF+8,"TURNO B",_F_AMAR),
                (COL_OFF+9,"TURNO C",_F_AZUL)
            ]:
                _da(r_hdr,ci_da,txt_da,fill_da,bold=True,color="FFFFFF" if fill_da==_F_HDR_VERD else "000000")
            ws_out.row_dimensions[r_hdr].height=14

            # Sub-cabeçalho
            r_sub=r_hdr+1
            for ci_da,txt_da,fill_da in [
                (COL_OFF,"CENTRO",_F_HDR_VERD),(COL_OFF+1,"% Ocup",_F_VERDE),(COL_OFF+2,"% Ocup",_F_AMAR),
                (COL_OFF+3,"% Ocup",_F_AZUL),(COL_OFF+4,"Ativo",_F_VERDE),(COL_OFF+5,"Ativo",_F_AMAR),
                (COL_OFF+6,"Ativo",_F_AZUL),(COL_OFF+7,"Horas",_F_VERDE),(COL_OFF+8,"Horas",_F_AMAR),
                (COL_OFF+9,"Horas",_F_AZUL)
            ]:
                _da(r_sub,ci_da,txt_da,fill_da,bold=True,
                    color="FFFFFF" if fill_da==_F_HDR_VERD else "000000",
                    center=fill_da!=_F_HDR_VERD)
            ws_out.row_dimensions[r_sub].height=14

            # Linhas de centros
            def _ocup_fill(v):
                try:
                    f=float(v)
                    if f>1.06: return _PF("solid",fgColor="FF0000")
                    if f>=1.00: return _F_AMAR_L
                    if f>=0.40: return _F_VERDE_L
                    return _F_BRANCO
                except: return _F_BRANCO

            r_cen=r_sub+1
            for _,crow in res_t["centros"].iterrows():
                oA,oB,oC=crow.ocup_A,crow.ocup_B,crow.ocup_C
                aA,aB,aC=int(crow.ativo_A),int(crow.ativo_B),int(crow.ativo_C)
                hA_d,hB_d,hC_d=crow.horas_disp_A,crow.horas_disp_B,crow.horas_disp_C
                _da(r_cen,COL_OFF,crow.centro,_F_BRANCO,center=False)
                _da(r_cen,COL_OFF+1,f"{oA:.0%}",_ocup_fill(oA))
                _da(r_cen,COL_OFF+2,f"{oB:.0%}",_ocup_fill(oB))
                _da(r_cen,COL_OFF+3,f"{oC:.0%}",_ocup_fill(oC))
                _da(r_cen,COL_OFF+4,aA,_F_VERDE if aA else _F_BRANCO,bold=True)
                _da(r_cen,COL_OFF+5,aB,_F_AMAR  if aB else _F_BRANCO,bold=True)
                _da(r_cen,COL_OFF+6,aC,_F_AZUL  if aC else _F_BRANCO,bold=True)
                _da(r_cen,COL_OFF+7,round(hA_d,2) if hA_d else 0,_F_VERDE if hA_d else _F_BRANCO)
                _da(r_cen,COL_OFF+8,round(hB_d,2) if hB_d else 0,_F_AMAR  if hB_d else _F_BRANCO)
                _da(r_cen,COL_OFF+9,round(hC_d,2) if hC_d else 0,_F_AZUL  if hC_d else _F_BRANCO)
                ws_out.row_dimensions[r_cen].height=13; r_cen+=1

            def _da_sup(ri,nome,qA,qB,qC,is_total=False):
                ft=_F_AMAR_TOT if is_total else _F_CINZA_L
                fq=_F_AMAR_TOT if is_total else _F_VERDE_L
                fh=_F_AMAR_TOT if is_total else _F_VERDE
                ct="1F4D19" if is_total else "000000"
                _da(ri,COL_OFF,nome,ft,bold=is_total,center=False,color=ct)
                _da(ri,COL_OFF+4,qA,fq if qA else _F_AMAR_L,bold=is_total,color=ct)
                _da(ri,COL_OFF+5,qB,fq if qB else _F_AMAR_L,bold=is_total,color=ct)
                _da(ri,COL_OFF+6,qC,fq if qC else _F_AMAR_L,bold=is_total,color=ct)
                _da(ri,COL_OFF+7,round(qA*hA_v*d_t2,2) if qA else 0,fh if qA else _F_BRANCO,bold=is_total,color=ct)
                _da(ri,COL_OFF+8,round(qB*hB_v*d_t2,2) if qB else 0,fh if qB else _F_BRANCO,bold=is_total,color=ct)
                _da(ri,COL_OFF+9,round(qC*hC_v*d_t2,2) if qC else 0,fh if qC else _F_BRANCO,bold=is_total,color=ct)
                ws_out.row_dimensions[ri].height=14

            _da_sup(r_cen,"TOTAL DE OPERADORES",res_t["op_A"],res_t["op_B"],res_t["op_C"],True); r_cen+=1
            for nm_s,ks in [("LAVADORA E INSPEÇÃO","lavadora"),("GRAVAÇÃO E ESTANQUEIDADE","gravacao"),
                             ("PRESET","preset"),("CORINGA","coringa"),("FACILITADOR","facilitador")]:
                s=sup_t[ks]; _da_sup(r_cen,nm_s,s["A"],s["B"],s["C"],False); r_cen+=1
            _da_sup(r_cen,"TOTAL POR TURNO",res_t["tot_A"],res_t["tot_B"],res_t["tot_C"],True); r_cen+=1

            _da(r_cen,COL_OFF,"TOTAL FUNCIONÁRIOS",_F_AMAR_TOT,bold=True,center=False,color="1F4D19")
            _da(r_cen,COL_OFF+4,res_t["total"],_F_AMAR_TOT,bold=True,color="1F4D19")
            tot_h_da=res_t["tot_A"]*hA_v*d_t2+res_t["tot_B"]*hB_v*d_t2+res_t["tot_C"]*hC_v*d_t2
            _da(r_cen,COL_OFF+7,round(tot_h_da,2),_F_AMAR_TOT,bold=True,color="1F4D19")
            ws_out.row_dimensions[r_cen].height=15; r_cen+=2

            for nm_p,vl_p,dest_p in [
                ("PRODUTIVIDADE POR TEMPO DE CICLO OPERACIONAL",res_t["prod_ciclo_op"],False),
                ("PRODUTIVIDADE POR TEMPO DE CICLO TOTAL",      res_t["prod_ciclo_tot"],False),
                ("PRODUTIVIDADE POR TEMPO DE LABOR OPERACIONAL",res_t["prod_labor_op"],False),
                ("PRODUTIVIDADE POR TEMPO DE LABOR TOTAL",      res_t["prod_labor_tot"],True),
            ]:
                fp=_F_AMAR_TOT if dest_p else _F_CINZA_L; cp="1F4D19" if dest_p else "000000"
                _da(r_cen,COL_OFF,nm_p,fp,bold=dest_p,center=False,color=cp,merge_end=COL_OFF+8)
                _da(r_cen,COL_OFF+9,f"{vl_p:.0%}" if isinstance(vl_p,float) else vl_p,fp,bold=dest_p,color=cp)
                ws_out.row_dimensions[r_cen].height=14; r_cen+=1

            # Larguras bloco DADOS AUTOMÁTICOS
            for ci_adj,w_adj in [(COL_OFF,16),(COL_OFF+1,8),(COL_OFF+2,8),(COL_OFF+3,8),
                                  (COL_OFF+4,7),(COL_OFF+5,7),(COL_OFF+6,7),
                                  (COL_OFF+7,10),(COL_OFF+8,10),(COL_OFF+9,10)]:
                ws_out.column_dimensions[get_column_letter(ci_adj)].width=w_adj

    tabelona_buf=BytesIO(); wb_out.save(tabelona_buf); tabelona_buf.seek(0)
    return tabelona_buf


# ─────────────────────────────────────────
# COMPARAÇÃO COM EXCEL REFERÊNCIA
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=3600)
def comparar_com_excel(res_app, file_bytes, tempo, dist, aplic, pmp, dias, horas_turno, thresholds, suporte_cfg):
    """Compara app vs Excel. Otimizado: JOIN e agrupamento feitos uma vez para todos os meses."""
    MAPA = {
        "Novembro":"NovFY26","Dezembro":"DezFY26","Janeiro":"JanFY26",
        "Fevereiro":"FevFY26","Março":"MarFY26","Abril":"AbrFY26",
        "Maio":"MaiFY26","Junho":"JunFY26","Julho":"JulFY26",
        "Agosto":"AgoFY26","Setembro":"SetFY26","Outubro":"OutFY26",
    }
    def safe_int(v):
        try: return int(float(v)) if v is not None else 0
        except: return 0
    def safe_float(v):
        try: return float(v) if v is not None else 0.0
        except: return 0.0

    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
        abas = wb.sheetnames
    except Exception as e:
        return None, None, f"Erro ao abrir arquivo: {e}"

    thr_A = thresholds["A"] / 100
    thr_B = thresholds["B"] / 100
    thr_C = thresholds["C"] / 100
    hA    = horas_turno["A"]
    hB    = horas_turno["B"]

    # ── PRÉ-CALCULAR: JOIN completo uma única vez para todos os meses ──────────
    try:
        df_all = (aplic.merge(pmp, on="modelo")
                       .merge(tempo, on=["centro","peca"])
                       .merge(dist,  on=["centro","peca"]))
        if "vol_int" not in df_all.columns: df_all["vol_int"] = 1.0
        df_all["vol_int"] = pd.to_numeric(df_all["vol_int"], errors="coerce").fillna(1.0)
        df_all["indice_ciclo"] = (df_all.t_ciclo * df_all.div_carga * df_all.div_volume * df_all.vol_int) / df_all.disponib
        df_all["min_ciclo"]    = df_all.indice_ciclo * df_all.qtd
        # Agrupar por centro+mes de uma vez
        agg_all = df_all.groupby(["centro","mes"]).agg(
            min_ciclo=("min_ciclo","sum"),
            qtd_total=("qtd","sum"),
            indice_medio=("indice_ciclo","mean")
        ).reset_index()
    except Exception as e:
        wb.close()
        return None, None, f"Erro no cálculo: {e}"

    resumo_rows  = []
    detalhe_rows = []

    for mes, aba in MAPA.items():
        r_app = res_app.get(mes)
        if not r_app: continue
        if aba not in abas:
            resumo_rows.append({"Mês":mes,"Status":"⚠️ aba ausente","Observação":f"Aba {aba} não encontrada"})
            continue

        ws  = wb[aba]
        d   = dias.get(mes, 0)
        if d == 0: continue
        minA = d * hA * 60
        minB = d * hB * 60

        xl_opA = safe_int(ws.cell(89,27).value)
        xl_opB = safe_int(ws.cell(89,28).value)
        xl_opC = safe_int(ws.cell(89,29).value)
        xl_tot = safe_int(ws.cell(96,27).value)
        xl_labor = safe_float(ws.cell(101,30).value)

        dA = r_app["op_A"] - xl_opA
        dB = r_app["op_B"] - xl_opB
        dC = r_app["op_C"] - xl_opC
        dT = r_app["total"] - xl_tot

        if dT==0 and dA==0 and dB==0 and dC==0:
            status = "✅ Igual"
        elif abs(dT) <= 2:
            status = "🟡 Pequena diferença"
        else:
            status = "🔴 Divergência"

        resumo_rows.append({
            "Mês":           mes,  "Status": status,
            "CEN A App":     r_app["op_A"],  "CEN A Excel": xl_opA,  "Δ A": f"{dA:+d}",
            "CEN B App":     r_app["op_B"],  "CEN B Excel": xl_opB,  "Δ B": f"{dB:+d}",
            "CEN C App":     r_app["op_C"],  "CEN C Excel": xl_opC,  "Δ C": f"{dC:+d}",
            "Total App":     r_app["total"], "Total Excel": xl_tot,  "Δ Total": f"{dT:+d}",
            "Labor App":     f"{r_app['prod_labor_tot']:.1%}",
            "Labor Excel":   f"{xl_labor:.1%}" if xl_labor else "—",
        })

        if status == "✅ Igual":
            continue

        # Usar slice do agg pré-calculado — sem recalcular JOIN
        agg_mes = agg_all[agg_all.mes == mes].copy()

        # Ler ocupação por centro do Excel (só leitura, rápido)
        centros_xl = {}
        for r in range(69, 89):
            cen_val = ws.cell(r,23).value
            if not cen_val: continue
            centros_xl[str(cen_val).strip()] = {
                "ocup_A": safe_float(ws.cell(r,24).value),
                "ocup_B": safe_float(ws.cell(r,25).value),
                "ativo_A": safe_int(ws.cell(r,27).value),
                "ativo_B": safe_int(ws.cell(r,28).value),
                "ativo_C": safe_int(ws.cell(r,29).value),
            }

        for _, row in agg_mes.iterrows():
            try:
                cen         = row.centro
                mc          = row.min_ciclo
                qtd_app     = row.qtd_total
                idx_medio   = row.indice_medio
                oA_app      = mc / minA if minA > 0 else 0
                oB_app      = mc / minB if minB > 0 else 0
                aA_app = 1 if oA_app > thr_A else 0
                aB_app = 1 if oA_app > thr_B else 0
                aC_app = 1 if oB_app > thr_C else 0

                xl      = centros_xl.get(cen, {})
                aA_xl   = xl.get("ativo_A", 0)
                aB_xl   = xl.get("ativo_B", 0)
                aC_xl   = xl.get("ativo_C", 0)
                oA_xl   = xl.get("ocup_A", 0.0)
                oB_xl   = xl.get("ocup_B", 0.0)

                for turno, a_app, a_xl, ocup_app, ocup_xl in [
                    ("A", aA_app, aA_xl, oA_app, oA_xl),
                    ("B", aB_app, aB_xl, oA_app, oA_xl),
                    ("C", aC_app, aC_xl, oB_app, oB_xl),
                ]:
                    if a_app == a_xl:
                        continue

                    delta_ocup    = ocup_app - float(ocup_xl)
                    abs_delta     = abs(delta_ocup)
                    mc_xl_esp     = float(ocup_xl) * (minA if turno in ("A","B") else minB)
                    vol_xl_estim  = mc_xl_esp / idx_medio if idx_medio > 0 else 0
                    idx_esperado  = mc_xl_esp / qtd_app   if qtd_app  > 0 else 0

                    if abs_delta > 0.15:
                        if qtd_app < vol_xl_estim * 0.7:
                            causa  = "Volume de peças menor que o esperado"
                            origem = f"IMPUTAPLICAÇÃO — verifique se todos os modelos do {cen} estão marcados com 1"
                            expl   = (f"App encontrou {qtd_app:.0f} peças no {cen} em {mes}, "
                                      f"mas o Excel indica carga equivalente a ~{vol_xl_estim:.0f} peças. "
                                      f"Algum modelo pode estar faltando na IMPUTAPLICAÇÃO ou o volume no INPUT_PMP está baixo.")
                        elif qtd_app > vol_xl_estim * 1.3:
                            causa  = "Volume de peças maior que o esperado"
                            origem = f"IMPUTAPLICAÇÃO — verifique se há modelos a mais para o {cen}"
                            expl   = (f"App encontrou {qtd_app:.0f} peças no {cen} em {mes}, "
                                      f"mas o Excel indica carga equivalente a ~{vol_xl_estim:.0f} peças. "
                                      f"Algum modelo pode estar marcado com 1 indevidamente na IMPUTAPLICAÇÃO.")
                        else:
                            causa  = "Índice de ciclo diferente"
                            origem = f"IMPUTDISTRIBUIÇÃO — verifique div_carga, div_volume e disponibilidade do {cen}"
                            expl   = (f"Volume compatível ({qtd_app:.0f} app vs ~{vol_xl_estim:.0f} Excel), "
                                      f"mas índice de ciclo app={idx_medio:.2f} vs esperado={idx_esperado:.2f} min/peça. "
                                      f"Fórmula: (t_ciclo × div_carga × div_volume) ÷ disponibilidade.")
                    else:
                        thr_u  = thr_A if turno=="A" else (thr_B if turno=="B" else thr_C)
                        causa  = f"Ocupação próxima do threshold ({thr_u:.0%})"
                        origem = f"INPUT_PMP — verifique volumes dos modelos do {cen}"
                        expl   = (f"Ocupação app={ocup_app:.1%} vs Excel={ocup_xl:.1%}. "
                                  f"Threshold={thr_u:.0%}. Pequena diferença de dado cruza o limite.")

                    detalhe_rows.append({
                        "Mês": mes, "Centro": cen, "Turno": turno,
                        "App — Ativo":   "✅ Sim" if a_app else "❌ Não",
                        "Excel — Ativo": "✅ Sim" if a_xl  else "❌ Não",
                        "Ocup. App":     f"{ocup_app:.1%}",
                        "Ocup. Excel":   f"{float(ocup_xl):.1%}",
                        "Δ Ocupação":    f"{delta_ocup:+.1%}",
                        "Causa":         causa,
                        "Onde investigar": origem,
                        "Explicação":    expl,
                    })
            except Exception:
                continue

    wb.close()
    return (pd.DataFrame(resumo_rows),
            pd.DataFrame(detalhe_rows) if detalhe_rows else pd.DataFrame(),
            None)


def find_col(df, candidates, aba, campo):
    """Busca coluna por lista de nomes candidatos. Fallback por índice com aviso."""
    for c in candidates:
        if c in df.columns:
            return c
    # fallback posicional com aviso
    idx_fallback = {"centro":0,"peca":1,"t_ciclo":5,"t_labor":6,
                    "div_carga":7,"vol_int":8,"div_volume":9,"disponib":10}
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

def read_indices(fb, log):
    """Lê os índices de ciclo e labor diretamente da primeira aba mensal disponível.
    Os índices (col L) são constantes entre meses — calculados a partir de H,I,J,K
    que podem diferir do IMPUTDISTRIBUIÇÃO em casos de máquinas compartilhadas."""
    ABAS_MENSAIS = ["NovFY26","DezFY26","JanFY26","FevFY26","MarFY26","AbrFY26",
                    "MaiFY26","JunFY26","JulFY26","AgoFY26","SetFY26","OutFY26"]
    try:
        import openpyxl
        wb = openpyxl.load_workbook(BytesIO(fb), read_only=True, data_only=True)
        abas = wb.sheetnames
        aba_ref = next((a for a in ABAS_MENSAIS if a in abas), None)
        if not aba_ref:
            log.append("⚠️ Nenhuma aba mensal encontrada — usando índices calculados do IMPUTDISTRIBUIÇÃO")
            wb.close()
            return None
        ws = wb[aba_ref]
        rows = []
        for r in range(7, 64):
            centro = ws.cell(r, 1).value
            peca   = ws.cell(r, 2).value
            l_val  = ws.cell(r, 12).value   # índice_ciclo = (t_ciclo*H*I*J)/K
            g_val  = ws.cell(r, 7).value    # t_labor
            h_val  = ws.cell(r, 8).value    # div_carga
            d_val  = ws.cell(r, 4).value    # peças/trator
            if centro and peca and l_val:
                rows.append({
                    "centro": str(centro).strip(),
                    "peca":   str(peca).strip(),
                    "indice_ciclo": float(l_val),
                    "t_labor":      float(g_val) if g_val else 0,
                    "div_carga":    float(h_val) if h_val else 1,
                    "pecas_trator": float(d_val) if d_val else 1,
                })
        wb.close()
        df = pd.DataFrame(rows)
        log.append(f"✅ Índices lidos da aba {aba_ref}: {len(df)} linhas")
        return df
    except Exception as e:
        log.append(f"⚠️ Erro ao ler índices da aba mensal: {e} — usando cálculo pelo IMPUTDISTRIBUIÇÃO")
        return None

def read_turnos(fb):
    """Lê IMPUTTURNOS — retorna horas acumuladas por turno (7.5, 14.25, 19.5)."""
    try:
        df = pd.read_excel(BytesIO(fb), sheet_name='IMPUTTURNOS', header=None)
        hA = float(df.iloc[0,1]) if pd.notna(df.iloc[0,1]) else 7.5
        hB = float(df.iloc[0,2]) if pd.notna(df.iloc[0,2]) else 14.25
        hC = float(df.iloc[0,3]) if pd.notna(df.iloc[0,3]) else 19.5
        return {"A": hA, "B": hB, "C": hC}
    except:
        return {"A": 7.5, "B": 14.25, "C": 19.5}

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
    out = df[[c["centro"],c["peca"],c["div_carga"],c["vol_int"],c["div_volume"],c["disponib"]]].copy()
    out.columns = ["centro","peca","div_carga","vol_int","div_volume","disponib"]
    out["vol_int"] = pd.to_numeric(out["vol_int"], errors="coerce").fillna(1.0)
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
        ex = ", ".join([f"{c}/{p}" for c,p in zip(zero_disp.centro[:5], zero_disp.peca[:5])])
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
        exemplos = list(zip(labor_maior.centro.head(3), labor_maior.peca.head(3)))
        alertas.append(f"{len(labor_maior)} linha(s) com t_labor > t_ciclo (fisicamente improvável): {exemplos}")
    _qtd_mes = pmp.groupby("mes")["qtd"].sum()
    for m in MESES:
        qtd_m = _qtd_mes.get(m, 0)
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
    # Cálculo 100% baseado nos inputs — IMPUTDISTRIBUIÇÃO é a fonte de verdade
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

    resultados = {}
    for mes in MESES:
        d = dias.get(mes, 0)
        if d == 0: resultados[mes] = None; continue
        sub = agg[agg.mes == mes].copy()
        if sub.empty: resultados[mes] = None; continue

        minA = d*hA*60; minB = d*hB*60; minC = d*hC*60
        # Vetorização — elimina loop Python por centro
        df_c = sub[["centro","min_ciclo","min_labor"]].copy()
        df_c["ocup_A"] = df_c.min_ciclo / minA if minA > 0 else 0.0
        df_c["ocup_B"] = df_c.min_ciclo / minB if minB > 0 else 0.0
        df_c["ocup_C"] = df_c.min_ciclo / minC if minC > 0 else 0.0
        df_c["ativo_A"] = (df_c.ocup_A > thr_A).astype(int)
        df_c["ativo_B"] = (df_c.ocup_A > thr_B).astype(int)
        df_c["ativo_C"] = (df_c.ocup_B > thr_C).astype(int)
        if overrides and mes in overrides:
            for cen, ov in overrides[mes].items():
                msk = df_c.centro == cen
                if "A" in ov: df_c.loc[msk, "ativo_A"] = ov["A"]
                if "B" in ov: df_c.loc[msk, "ativo_B"] = ov["B"]
                if "C" in ov: df_c.loc[msk, "ativo_C"] = ov["C"]
        df_c = df_c.rename(columns={"min_ciclo":"min_ciclo_total","min_labor":"min_labor_total"})
        df_c["min_disp_A"]   = minA; df_c["min_disp_B"] = minB; df_c["min_disp_C"] = minC
        df_c["horas_ciclo"]  = df_c.min_ciclo_total / 60
        df_c["horas_labor"]  = df_c.min_labor_total / 60
        df_c["horas_disp_A"] = d * hA * df_c.ativo_A
        df_c["horas_disp_B"] = d * hB * df_c.ativo_B
        df_c["horas_disp_C"] = d * hC * df_c.ativo_C
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
# CACHE DE CÁLCULO (nível de módulo — obrigatório para st.cache_data)
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=3600)
def _calcular_cached(pmp_hash, _pmp, _tempo, _dist, _aplic,
                     dias_hash, _dias, hA, hB, hC, tA, tB, tC, sup_str):
    """Wrapper cacheável: suporte_cfg passa como string para ser hashável."""
    import json as _json
    _sup = _json.loads(sup_str)
    return calcular(_pmp, _tempo, _dist, _aplic, _dias,
                    {"A":hA,"B":hB,"C":hC}, {"A":tA,"B":tB,"C":tC}, _sup,
                    retornar_intermediarios=True)


# ─────────────────────────────────────────
# TABELA RESULTADO (versão principal — com horas/minutos + comparação real)
# ─────────────────────────────────────────
def show_tabela(r, real_mensal=None, mes=None):
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

    # ── TABELA DE HORAS E MINUTOS ─────────────────────────────────────────
    st.markdown('<div class="jd-sub">📊 Horas e minutos disponíveis no mês</div>', unsafe_allow_html=True)
    show_horas_minutos_table(r, None, mes)

    # ── COMPARAÇÃO COM EXCEL DADOS AUTOMÁTICOS ─────────────────────────────
    if real_mensal is not None and mes:
        st.markdown('<div class="jd-sub">🔄 App vs Excel — DADOS AUTOMÁTICOS (comparação com o real)</div>', unsafe_allow_html=True)
        show_real_vs_app(r, real_mensal, mes)

def exportar(resultados):
    out=BytesIO(); wb=openpyxl.Workbook()
    brd=Border(left=Side(style='thin',color='CCCCCC'),right=Side(style='thin',color='CCCCCC'),
               top=Side(style='thin',color='CCCCCC'),bottom=Side(style='thin',color='CCCCCC'))
    def ec(c,bg="FFFFFF",fg="000000",bold=False,fmt=None,center=True):
        c.font=Font(name="Arial",bold=bold,color=fg,size=9)
        c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="center" if center else "left",vertical="center")
        c.border=brd
        if fmt and isinstance(fmt, str) and len(fmt) > 0:
            try: c.number_format=str(fmt)
            except: pass
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
            # Format percentages as strings to avoid openpyxl number_format issues
            v_cell = f"{v:.1%}" if ci>=7 and isinstance(v,float) else v
            c=ws.cell(ri,ci,v_cell)
            ec(c,JD_Y if ci==10 and isinstance(v,float) else bg,
               JD_V if ci==10 and isinstance(v,float) else "000000",
               bool(ci==10 and isinstance(v,float)))
        ws.row_dimensions[ri].height=15
    for mes in MESES:
        r=resultados.get(mes)
        if not r: continue
        wsm=wb.create_sheet(mes[:10]); hA,hB,hC,dias=r["hA"],r["hB"],r["hC"],r["dias"]
        # Linha 1 — mês + grupos
        for ci,txt in [(1,""),(2,"TURNO A"),(3,"TURNO B"),(4,"TURNO C"),
                       (5,"TURNO A"),(6,"TURNO B"),(7,"TURNO C"),
                       (8,"TURNO A"),(9,"TURNO B"),(10,"TURNO C")]:
            ec(wsm.cell(1,ci,txt),JD_V,"FFFFFF",True)
        wsm.cell(1,1,mes.upper()); ec(wsm.cell(1,1),JD_V,"FFFFFF",True)

        # Linha 2 — descrição de cada bloco de colunas
        JD_V2 = JD_VERDE_ESC.replace("#","")
        JD_Y2 = JD_AMARELO.replace("#","")
        JD_V3 = JD_VERDE_ESC.replace("#","")
        wsm.merge_cells("B2:D2")
        c2=wsm.cell(2,1,"CENTRO"); c2.font=Font(name="Arial",bold=True,color="FFFFFF",size=9); c2.fill=PatternFill("solid",fgColor=JD_V2); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        c2=wsm.cell(2,2,"% OCUPAÇÃO"); c2.font=Font(name="Arial",bold=True,color="FFFFFF",size=9); c2.fill=PatternFill("solid",fgColor=JD_V2); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        wsm.merge_cells("E2:G2")
        c2=wsm.cell(2,5,"TURNO ATIVO  (0=inativo  1=ativo)"); c2.font=Font(name="Arial",bold=True,color=JD_V3,size=9); c2.fill=PatternFill("solid",fgColor=JD_Y2); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        wsm.merge_cells("H2:J2")
        c2=wsm.cell(2,8,"HORAS DISPONIVEIS NO MES"); c2.font=Font(name="Arial",bold=True,color="FFFFFF",size=9); c2.fill=PatternFill("solid",fgColor="1565C0"); c2.alignment=Alignment(horizontal="center",vertical="center"); c2.border=brd
        wsm.row_dimensions[2].height = 16

        def cbg(v):
            if v>1.0: return "FFCDD2"
            if v>=0.85: return "FFFDE7"
            return "E8F5E9"
        ri=3
        for _,row in r["centros"].iterrows():
            for ci,(val,bg,ctr) in enumerate([
                (row.centro,"FFFFFF",False),
                (f"{row.ocup_A:.1%}",cbg(row.ocup_A),True),(f"{row.ocup_B:.1%}",cbg(row.ocup_B),True),(f"{row.ocup_C:.1%}",cbg(row.ocup_C),True),
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
            ec(wsm.cell(ri,10,f"{v:.1%}" if isinstance(v,float) else v),JD_Y if dest else "FFFFFF",JD_V if dest else "000000",dest)
            ri+=1
        for ci,w in enumerate([14,8,8,8,8,8,8,24,10,10],1):
            wsm.column_dimensions[get_column_letter(ci)].width=w
    wb.save(out); out.seek(0)
    return out

# ─────────────────────────────────────────
# COMPARAÇÃO COM EXCEL REFERÊNCIA
# ─────────────────────────────────────────
# ─────────────────────────────────────────
# LEITURA DE DADOS REAIS DO EXCEL (DADOS AUTOMÁTICOS)
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=3600)
def read_real_data(file_bytes):
    """Lê os dados reais calculados pelo Excel:
    - RESUMO_MO: totais por turno e produtividade por mês
    - Abas mensais: DADOS AUTOMÁTICOS (cols W-AF, linhas 68-96) + horas por turno (linha 67)
    """
    MAPA = {
        "Novembro":"NovFY26","Dezembro":"DezFY26","Janeiro":"JanFY26",
        "Fevereiro":"FevFY26","Março":"MarFY26","Abril":"AbrFY26",
        "Maio":"MaiFY26","Junho":"JunFY26","Julho":"JulFY26",
        "Agosto":"AgoFY26","Setembro":"SetFY26","Outubro":"OutFY26",
    }
    MESES_ABREV_MAP = dict(zip(
        ["Novembro","Dezembro","Janeiro","Fevereiro","Março","Abril",
         "Maio","Junho","Julho","Agosto","Setembro","Outubro"],
        ["NOV","DEZ","JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT"]
    ))
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception as e:
        return {}, {}

    abas = wb.sheetnames

    # ── RESUMO_MO: totais por turno e produtividade ──────────────────────────
    resumo_real = {}
    if "RESUMO_MO" in abas:
        ws_r = wb["RESUMO_MO"]
        # R3=TURNO A, R4=B, R5=C, R6=TOTAL, R7=PROD
        # cols: 3=NOV, 4=DEZ, ..., 14=OUT, 15=ANO
        meses_list = list(MESES_ABREV_MAP.keys())
        for idx, mes in enumerate(meses_list, 3):
            tA = safe_int(ws_r.cell(3, idx).value)
            tB = safe_int(ws_r.cell(4, idx).value)
            tC = safe_int(ws_r.cell(5, idx).value)
            tot = safe_int(ws_r.cell(6, idx).value)
            prod_raw = ws_r.cell(7, idx).value
            try:
                prod = float(prod_raw) if prod_raw not in (None, "#DIV/0!") else None
            except:
                prod = None
            resumo_real[mes] = {"tot_A": tA, "tot_B": tB, "tot_C": tC,
                                 "total": tot, "prod_esperada": prod}
        # ANO (col 15)
        resumo_real["_ANO"] = {
            "tot_A": safe_int(ws_r.cell(3, 15).value),
            "tot_B": safe_int(ws_r.cell(4, 15).value),
            "tot_C": safe_int(ws_r.cell(5, 15).value),
            "total": safe_int(ws_r.cell(6, 15).value),
        }

    # ── Abas mensais: DADOS AUTOMÁTICOS (cols W=23 a AF=32, linhas 68-96) ──
    mensal_real = {}
    for mes, aba in MAPA.items():
        if aba not in abas:
            continue
        ws_m = wb[aba]
        # Linha 67 col 30=AF: horas por turno (W=col23=hA, X=24=hB, Y=25 n/a, Z=26 n/a)
        # Row 67: cols 30(AF)=hA, 31=hB, 32=hC  based on image: 8.8, 8.23, 7.68
        # Actually from our inspection: R67 col 30=hA=8.8, 31=hB=8.23, 32=hC=7.68
        hA_r = safe_float(ws_m.cell(67, 30).value)
        hB_r = safe_float(ws_m.cell(67, 31).value)
        hC_r = safe_float(ws_m.cell(67, 32).value)
        # Row 66: col 34=data revisao; col 28=horas por turno label
        # Days from header row 66 col 29 area? Actually dias comes from INPUT_PMP
        # But we have it in res_app already. Let's just read totais.
        # Rows 69-88: centros (col 23=centro, 24=ocup_A, 25=ocup_B, 26=ocup_C, 27=ativo_A, 28=ativo_B, 29=ativo_C, 30=horas_A, 31=horas_B, 32=horas_C)
        centros_r = {}
        for row_i in range(69, 89):
            cen_v = ws_m.cell(row_i, 23).value
            if not cen_v:
                continue
            centros_r[str(cen_v).strip()] = {
                "ocup_A": safe_float(ws_m.cell(row_i, 24).value),
                "ocup_B": safe_float(ws_m.cell(row_i, 25).value),
                "ocup_C": safe_float(ws_m.cell(row_i, 26).value),
                "ativo_A": safe_int(ws_m.cell(row_i, 27).value),
                "ativo_B": safe_int(ws_m.cell(row_i, 28).value),
                "ativo_C": safe_int(ws_m.cell(row_i, 29).value),
                "horas_A": safe_float(ws_m.cell(row_i, 30).value),
                "horas_B": safe_float(ws_m.cell(row_i, 31).value),
                "horas_C": safe_float(ws_m.cell(row_i, 32).value),
            }
        # Row 89: TOTAL DE OPERADORES
        op_A_r = safe_int(ws_m.cell(89, 27).value)
        op_B_r = safe_int(ws_m.cell(89, 28).value)
        op_C_r = safe_int(ws_m.cell(89, 29).value)
        # Row 95: TOTAL POR TURNO
        tot_A_r = safe_int(ws_m.cell(95, 27).value)
        tot_B_r = safe_int(ws_m.cell(95, 28).value)
        tot_C_r = safe_int(ws_m.cell(95, 29).value)
        h_tot_A_r = safe_float(ws_m.cell(95, 30).value)
        h_tot_B_r = safe_float(ws_m.cell(95, 31).value)
        h_tot_C_r = safe_float(ws_m.cell(95, 32).value)
        # Row 96: TOTAL FUNCIONARIOS
        total_r = safe_int(ws_m.cell(96, 27).value)
        h_total_r = safe_float(ws_m.cell(96, 30).value)
        # Rows 98-101: produtividades (col 30)
        ciclo_op_r  = safe_float(ws_m.cell(98, 30).value)
        ciclo_tot_r = safe_float(ws_m.cell(99, 30).value)
        labor_op_r  = safe_float(ws_m.cell(100, 30).value)
        labor_tot_r = safe_float(ws_m.cell(101, 30).value)

        mensal_real[mes] = {
            "hA": hA_r, "hB": hB_r, "hC": hC_r,
            "centros": centros_r,
            "op_A": op_A_r, "op_B": op_B_r, "op_C": op_C_r,
            "tot_A": tot_A_r, "tot_B": tot_B_r, "tot_C": tot_C_r,
            "total": total_r,
            "h_tot_A": h_tot_A_r, "h_tot_B": h_tot_B_r, "h_tot_C": h_tot_C_r,
            "h_total": h_total_r,
            "prod_ciclo_op": ciclo_op_r, "prod_ciclo_tot": ciclo_tot_r,
            "prod_labor_op": labor_op_r, "prod_labor_tot": labor_tot_r,
        }

    wb.close()
    return resumo_real, mensal_real


def show_horas_minutos_table(r, dias_dict, mes):
    """Exibe a tabela de TOTAL DE MINUTOS / TOTAL DE HORAS / DIAS TRABALHADOS / HORAS POR TURNO."""
    d = r["dias"]; hA = r["hA"]; hB = r["hB"]; hC = r["hC"]
    # Horas por turno = acumuladas (hA, hB, hC são acumuladas)
    # Minutos por turno = dias * horas_acumuladas * 60
    min_A = d * hA * 60
    min_B = d * hB * 60
    min_C = d * hC * 60
    h_A = d * hA
    h_B = d * hB
    h_C = d * hC

    df_hm = pd.DataFrame([
        {"": "TOTAL DE MINUTOS", "Turno A": round(min_A, 1), "Turno B": round(min_B, 1), "Turno C": round(min_C, 1)},
        {"": "TOTAL DE HORAS",   "Turno A": round(h_A, 2),   "Turno B": round(h_B, 2),   "Turno C": round(h_C, 2)},
        {"": "DIAS TRABALHADOS", "Turno A": d,                "Turno B": d,                "Turno C": d},
        {"": "HORAS POR TURNO",  "Turno A": hA,               "Turno B": hB,               "Turno C": hC},
    ])

    def style_hm(row):
        label = row[""]
        if "MINUTOS" in label:
            return ["background-color:#1F4D19;color:#FFDE00;font-weight:700"]*4
        if "HORAS" in label and "TURNO" not in label:
            return ["background-color:#1F4D19;color:#FFFFFF;font-weight:700"]*4
        if "DIAS" in label:
            return ["background-color:#2E7D32;color:#FFFFFF;font-weight:600"]*4
        return ["background-color:#F4F4F4;color:#1A1A1A"]*4

    st.dataframe(df_hm.style.apply(style_hm, axis=1), use_container_width=True, hide_index=True)


def show_real_vs_app(r_app, real_mensal, mes):
    """Compara App calculado vs Excel DADOS AUTOMÁTICOS para o mês."""
    if not real_mensal:
        st.info("Dados reais do Excel não encontrados. Suba o arquivo com as abas mensais calculadas.")
        return
    real = real_mensal.get(mes)
    if not real:
        st.info(f"Aba {mes} não encontrada no Excel de referência.")
        return

    # Construir tabela de comparação de totais
    rows_cmp = []
    for label, app_v, xl_v in [
        ("Turno A — Operadores CEN", r_app["op_A"],  real["op_A"]),
        ("Turno B — Operadores CEN", r_app["op_B"],  real["op_B"]),
        ("Turno C — Operadores CEN", r_app["op_C"],  real["op_C"]),
        ("TOTAL POR TURNO — A",      r_app["tot_A"], real["tot_A"]),
        ("TOTAL POR TURNO — B",      r_app["tot_B"], real["tot_B"]),
        ("TOTAL POR TURNO — C",      r_app["tot_C"], real["tot_C"]),
        ("TOTAL FUNCIONÁRIOS",       r_app["total"], real["total"]),
    ]:
        delta = int(app_v) - int(xl_v)
        icon = "✅" if delta == 0 else ("🟡" if abs(delta) <= 2 else "🔴")
        rows_cmp.append({"Indicador": label, "App": app_v, "Excel (Real)": xl_v,
                          "Δ": f"{delta:+d}", "Status": icon})

    # Produtividades
    for label, app_v, xl_v in [
        ("Ciclo Operacional",  r_app["prod_ciclo_op"],  real["prod_ciclo_op"]),
        ("Ciclo Total",        r_app["prod_ciclo_tot"], real["prod_ciclo_tot"]),
        ("Labor Operacional",  r_app["prod_labor_op"],  real["prod_labor_op"]),
        ("⭐ Labor Total",      r_app["prod_labor_tot"], real["prod_labor_tot"]),
    ]:
        delta_p = app_v - xl_v if xl_v else 0
        icon_p = "✅" if abs(delta_p) < 0.005 else ("🟡" if abs(delta_p) < 0.02 else "🔴")
        rows_cmp.append({"Indicador": label,
                          "App": f"{app_v:.1%}", "Excel (Real)": f"{xl_v:.1%}" if xl_v else "—",
                          "Δ": f"{delta_p:+.1%}", "Status": icon_p})

    # Separar linhas de inteiros e percentuais para evitar coluna de tipo misto (ArrowInvalid)
    rows_int = [r for r in rows_cmp if isinstance(r["App"], (int, float)) and not isinstance(r["App"], bool)]
    rows_pct = [r for r in rows_cmp if isinstance(r["App"], str)]

    def _style_cmp(row, cols):
        st_val = str(row.get("Status", ""))
        if "✅" in st_val:   base = "background-color:#003D10;color:#B9F6CA"
        elif "🟡" in st_val: base = "background-color:#3D2D00;color:#FFE57F"
        else:                base = "background-color:#3D0000;color:#FF8A80"
        return ["" if c != "Status" else base for c in cols]

    if rows_int:
        df_int = pd.DataFrame(rows_int)
        df_int["App"] = df_int["App"].astype(int)
        df_int["Excel (Real)"] = pd.to_numeric(df_int["Excel (Real)"], errors="coerce").fillna(0).astype(int)
        cols_int = list(df_int.columns)
        st.dataframe(df_int.style.apply(lambda r: _style_cmp(r, cols_int), axis=1),
                     use_container_width=True, hide_index=True)
    if rows_pct:
        df_pct = pd.DataFrame(rows_pct)
        cols_pct = list(df_pct.columns)
        st.dataframe(df_pct.style.apply(lambda r: _style_cmp(r, cols_pct), axis=1),
                     use_container_width=True, hide_index=True)


def show_fy26_anual(res_base, real_resumo, dias, horas_turno):
    """Exibe resumo FY26 anual: app calculado vs real do Excel, e totais acumulados."""
    st.markdown('<div class="jd-section">📅 Resumo FY26 — Visão Anual</div>', unsafe_allow_html=True)

    rows_ano = []
    totais_app = {"tot_A": 0, "tot_B": 0, "tot_C": 0, "h_ciclo": 0, "h_labor": 0, "h_ativos": 0, "h_todos": 0}
    totais_xl  = {"tot_A": 0, "tot_B": 0, "tot_C": 0}

    for mes, abr in zip(MESES, MESES_ABREV):
        r = res_base.get(mes)
        d = dias.get(mes, 0)
        real = real_resumo.get(mes, {}) if real_resumo else {}
        if not r:
            continue

        # Acumular para totais anuais
        totais_app["tot_A"]   += r["tot_A"]
        totais_app["tot_B"]   += r["tot_B"]
        totais_app["tot_C"]   += r["tot_C"]
        totais_app["h_ciclo"] += r["h_ciclo"]
        totais_app["h_labor"] += r["h_labor"]
        totais_app["h_ativos"]+= r["h_ativos"]
        totais_app["h_todos"] += r["h_todos"]
        if real:
            totais_xl["tot_A"] += real.get("tot_A", 0)
            totais_xl["tot_B"] += real.get("tot_B", 0)
            totais_xl["tot_C"] += real.get("tot_C", 0)

        hA = horas_turno["A"]; hB = horas_turno["B"]; hC = horas_turno["C"]
        min_A = d * hA * 60; min_B = d * hB * 60; min_C = d * hC * 60

        xl_tot = real.get("total", None)
        delta_tot = r["total"] - xl_tot if xl_tot else None
        icon = "—" if xl_tot is None else ("✅" if delta_tot == 0 else ("🟡" if abs(delta_tot) <= 2 else "🔴"))

        xl_labor = real.get("prod_esperada", None)
        row = {
            "Mês":          abr,
            "Dias":         d,
            "Min. Turno A": round(min_A, 0),
            "Min. Turno B": round(min_B, 0),
            "Min. Turno C": round(min_C, 0),
            "H. Turno A":   round(d * hA, 1),
            "H. Turno B":   round(d * hB, 1),
            "H. Turno C":   round(d * hC, 1),
            "App — A":      r["tot_A"],
            "App — B":      r["tot_B"],
            "App — C":      r["tot_C"],
            "App — Total":  r["total"],
            "Excel — Total": int(xl_tot) if xl_tot else None,
            "Δ Total":      f"{delta_tot:+d}" if delta_tot is not None else "—",
            "Labor App":    f"{r['prod_labor_tot']:.1%}",
            "Labor Excel":  f"{xl_labor:.1%}" if xl_labor else "—",
            "Status":       icon,
        }
        rows_ano.append(row)

    df_ano = pd.DataFrame(rows_ano)

    def style_ano(row):
        st_val = str(row.get("Status", ""))
        styles = []
        for col in df_ano.columns:
            if col == "Status":
                if "✅" in st_val:   styles.append("background-color:#003D10;color:#B9F6CA;font-weight:700")
                elif "🟡" in st_val: styles.append("background-color:#3D2D00;color:#FFE57F;font-weight:700")
                elif "🔴" in st_val: styles.append("background-color:#3D0000;color:#FF8A80;font-weight:700")
                else:                styles.append("")
            elif col in ("Min. Turno A", "Min. Turno B", "Min. Turno C"):
                styles.append("background-color:#1F4D19;color:#FFDE00;font-weight:600")
            elif col in ("H. Turno A", "H. Turno B", "H. Turno C"):
                styles.append("background-color:#2E7D32;color:#FFFFFF")
            elif col in ("App — A", "App — B", "App — C", "App — Total"):
                styles.append("background-color:#E3F2FD;color:#01579B")
            elif col == "Δ Total":
                v = str(row.get("Δ Total", "0"))
                try:
                    n = int(v.replace("+",""))
                    if n == 0:   styles.append("background-color:#003D10;color:#B9F6CA;font-weight:700")
                    elif abs(n) <= 2: styles.append("background-color:#3D2D00;color:#FFE57F;font-weight:700")
                    else:        styles.append("background-color:#3D0000;color:#FF8A80;font-weight:700")
                except:          styles.append("")
            else:
                styles.append("")
        return styles

    st.dataframe(df_ano.style.apply(style_ano, axis=1), use_container_width=True, hide_index=True)

    # Métricas resumo anuais
    st.markdown('<div class="jd-sub">Totais acumulados FY26</div>', unsafe_allow_html=True)
    prod_labor_tot_ano = totais_app["h_labor"] / totais_app["h_todos"] if totais_app["h_todos"] > 0 else 0
    prod_ciclo_tot_ano = totais_app["h_ciclo"] / totais_app["h_todos"] if totais_app["h_todos"] > 0 else 0
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total h. ciclo acum.", f"{totais_app['h_ciclo']:.0f}h")
    c2.metric("Total h. labor acum.", f"{totais_app['h_labor']:.0f}h")
    c3.metric("Total h. disponíveis", f"{totais_app['h_todos']:.0f}h")
    c4.metric("⭐ Labor Total FY26",   f"{prod_labor_tot_ano:.1%}")
    c5.metric("Ciclo Total FY26",     f"{prod_ciclo_tot_ano:.1%}")

    # Comparação anual com Excel RESUMO_MO
    if real_resumo and real_resumo.get("_ANO"):
        ano_xl = real_resumo["_ANO"]
        st.markdown('<div class="jd-sub">App vs Excel — Totais por turno FY26</div>', unsafe_allow_html=True)
        rows_diff = []
        for turno, app_v, xl_v in [
            ("Turno A", totais_app["tot_A"], ano_xl["tot_A"]),
            ("Turno B", totais_app["tot_B"], ano_xl["tot_B"]),
            ("Turno C", totais_app["tot_C"], ano_xl["tot_C"]),
        ]:
            delta = app_v - xl_v
            icon = "✅" if delta == 0 else ("🟡" if abs(delta) <= 5 else "🔴")
            rows_diff.append({"Turno": turno, "App (acum.)": app_v,
                               "Excel RESUMO_MO (acum.)": xl_v,
                               "Δ": f"{delta:+d}", "Status": icon})
        st.dataframe(pd.DataFrame(rows_diff), use_container_width=True, hide_index=True)


def show_memoria(r, mes, df_intermediario, agg, horas_turno, thresholds):
    st.markdown(f'<div class="jd-section">Memória de cálculo — {mes}</div>', unsafe_allow_html=True)
    st.caption("Cada passo mostra a fórmula usada e os dados reais calculados para esse mês.")

    sup = r["suporte"]
    d   = r["dias"]
    hA, hB, hC = r["hA"], r["hB"], r["hC"]

    # ── PASSO 1: Inputs ─────────────────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">1</span> <b>Inputs utilizados</b></div>', unsafe_allow_html=True)
    st.markdown(
        f"**Dias trabalhados:** {d} &nbsp;|&nbsp; Turno A: **{hA}h** &nbsp;|&nbsp; Turno B: **{hB}h** &nbsp;|&nbsp; Turno C: **{hC}h**\n\n"
        f"Minutos disponíveis por turno (dias × horas × 60):")
    c1,c2,c3 = st.columns(3)
    c1.metric("Turno A", f"{r['minA']:.0f} min", f"{d}×{hA}×60")
    c2.metric("Turno B", f"{r['minB']:.0f} min", f"{d}×{hB}×60")
    c3.metric("Turno C", f"{r['minC']:.0f} min", f"{d}×{hC}×60")

    # Amostra dos inputs lidos
    df_inp = df_intermediario[df_intermediario.mes==mes][
        ["centro","peca","modelo","t_ciclo","t_labor","div_carga","div_volume","vol_int","disponib","qtd"]
    ].head(8).copy()
    df_inp.columns = ["Centro","Peça","Modelo","T.Ciclo","T.Labor","Div.Carga","Div.Volume","Vol.Int","Disponib","Qtd"]
    st.caption(f"Amostra dos dados de input após JOIN (primeiras 8 linhas de {len(df_intermediario[df_intermediario.mes==mes])} combinações):")
    st.dataframe(df_inp.reset_index(drop=True), use_container_width=True, hide_index=True)

    # ── PASSO 2: Índice de ciclo ─────────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">2</span> <b>Cálculo do índice de ciclo por linha</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">indice_ciclo = (t_ciclo × div_carga × div_volume × vol_interna) ÷ disponibilidade<br><br>→ Representa quantos minutos de máquina são necessários para produzir 1 peça neste centro.</div>', unsafe_allow_html=True)
    df_p2 = df_intermediario[df_intermediario.mes==mes][
        ["centro","peca","modelo","t_ciclo","div_carga","div_volume","vol_int","disponib","indice_ciclo"]
    ].head(8).copy().round(4)
    df_p2.columns = ["Centro","Peça","Modelo","T.Ciclo","Div.Carga","Div.Volume","Vol.Int","Disponib","Índice Ciclo ✓"]
    st.caption("Output do passo 2 — índice calculado por linha:")
    st.dataframe(df_p2.reset_index(drop=True), use_container_width=True, hide_index=True)

    # ── PASSO 3: Minutos por linha ───────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">3</span> <b>Minutos necessários por linha</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">min_ciclo = indice_ciclo × qtd_pecas_no_mes<br>min_labor = t_labor × div_carga × qtd_pecas_no_mes</div>', unsafe_allow_html=True)
    df_p3 = df_intermediario[df_intermediario.mes==mes][
        ["centro","peca","modelo","indice_ciclo","qtd","min_ciclo","min_labor"]
    ].head(8).copy().round(2)
    df_p3.columns = ["Centro","Peça","Modelo","Índice Ciclo","Qtd Peças","Min. Ciclo ✓","Min. Labor ✓"]
    st.caption("Output do passo 3 — minutos por combinação modelo+centro+peça:")
    st.dataframe(df_p3.reset_index(drop=True), use_container_width=True, hide_index=True)

    # ── PASSO 4: Agrupamento por centro ─────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">4</span> <b>Agrupamento por centro</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">SOMA de min_ciclo e min_labor de todas as combinações do mesmo centro<br>ocup_A = Σmin_ciclo ÷ min_disp_A &nbsp;|&nbsp; ocup_B = Σmin_ciclo ÷ min_disp_B &nbsp;|&nbsp; ocup_C = Σmin_ciclo ÷ min_disp_C</div>', unsafe_allow_html=True)
    df_p4 = r["centros"][["centro","min_ciclo_total","min_labor_total","min_disp_A","ocup_A","ocup_B","ocup_C"]].copy()
    df_p4["ocup_A"] = df_p4.ocup_A.map(lambda x: f"{x:.1%}")
    df_p4["ocup_B"] = df_p4.ocup_B.map(lambda x: f"{x:.1%}")
    df_p4["ocup_C"] = df_p4.ocup_C.map(lambda x: f"{x:.1%}")
    df_p4 = df_p4.rename(columns={"centro":"Centro","min_ciclo_total":"Σ Min.Ciclo",
        "min_labor_total":"Σ Min.Labor","min_disp_A":"Min.Disp.A",
        "ocup_A":"Ocup. A ✓","ocup_B":"Ocup. B ✓","ocup_C":"Ocup. C ✓"})
    st.caption("Output do passo 4 — ocupação calculada por centro:")
    st.dataframe(df_p4.reset_index(drop=True), use_container_width=True, hide_index=True)

    # ── PASSO 5: Ativação de turno ───────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">5</span> <b>Regras de ativação de turno</b></div>', unsafe_allow_html=True)
    st.markdown(
        f"- **Turno A** abre se ocup_A > **{thresholds['A']}%**\n"
        f"- **Turno B** abre se ocup_A > **{thresholds['B']}%** (A não aguenta sozinho)\n"
        f"- **Turno C** abre se ocup_B > **{thresholds['C']}%** (A+B não suficientes)")
    df_p5 = r["centros"][["centro","ocup_A","ocup_B","ocup_C","ativo_A","ativo_B","ativo_C"]].copy()
    df_p5["ocup_A"] = df_p5.ocup_A.map(lambda x: f"{x:.1%}")
    df_p5["ocup_B"] = df_p5.ocup_B.map(lambda x: f"{x:.1%}")
    df_p5["ocup_C"] = df_p5.ocup_C.map(lambda x: f"{x:.1%}")
    df_p5 = df_p5.rename(columns={"centro":"Centro","ocup_A":"Ocup. A","ocup_B":"Ocup. B","ocup_C":"Ocup. C",
        "ativo_A":"Ativo A ✓","ativo_B":"Ativo B ✓","ativo_C":"Ativo C ✓"})
    st.caption(f"Output do passo 5 — ativação por centro ({r['op_A']} centros no A, {r['op_B']} no B, {r['op_C']} no C):")
    st.dataframe(df_p5.reset_index(drop=True), use_container_width=True, hide_index=True)

    # ── PASSO 6: Suporte ─────────────────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">6</span> <b>Funções de suporte</b></div>', unsafe_allow_html=True)
    st.markdown("Cada função só é adicionada se o turno tiver **pelo menos 1 CEN ativo**. Se não tiver CEN no turno → suporte = 0.")
    sup_data = []
    for nome, key in [("Lavadora e Inspeção","lavadora"),("Gravação e Estanqueidade","gravacao"),
                      ("Preset","preset"),("Coringa","coringa"),("Facilitador","facilitador")]:
        s = sup[key]
        sup_data.append({"Função": nome, "Turno A ✓": s["A"], "Turno B ✓": s["B"], "Turno C ✓": s["C"]})
    st.caption("Output do passo 6 — pessoas de suporte por turno:")
    st.dataframe(pd.DataFrame(sup_data), use_container_width=True, hide_index=True)

    # ── PASSO 7: Totais por turno ────────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">7</span> <b>Total por turno</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">Total Turno X = Operadores CEN + Lavadora + Gravação + Preset + Coringa + Facilitador</div>', unsafe_allow_html=True)
    tot_data = {
        "Função": ["Operadores CEN","Lavadora","Gravação","Preset","Coringa","Facilitador","TOTAL ✓"],
        "Turno A": [r["op_A"], sup["lavadora"]["A"], sup["gravacao"]["A"], sup["preset"]["A"],
                    sup["coringa"]["A"], sup["facilitador"]["A"], r["tot_A"]],
        "Turno B": [r["op_B"], sup["lavadora"]["B"], sup["gravacao"]["B"], sup["preset"]["B"],
                    sup["coringa"]["B"], sup["facilitador"]["B"], r["tot_B"]],
        "Turno C": [r["op_C"], sup["lavadora"]["C"], sup["gravacao"]["C"], sup["preset"]["C"],
                    sup["coringa"]["C"], sup["facilitador"]["C"], r["tot_C"]],
    }
    st.dataframe(pd.DataFrame(tot_data), use_container_width=True, hide_index=True)
    st.markdown(f"➡️ **Total geral: {r['total']} funcionários** (A: {r['tot_A']} + B: {r['tot_B']} + C: {r['tot_C']})")

    # ── PASSO 8: Horas totais ────────────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">8</span> <b>Horas totais disponíveis</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">Horas CEN ativos = Σ(centros_ativos_A × dias × hA) + Σ(centros_ativos_B × dias × hB) + Σ(centros_ativos_C × dias × hC)<br>Horas todos = Σ(total_A × dias × hA) + Σ(total_B × dias × hB) + Σ(total_C × dias × hC)</div>', unsafe_allow_html=True)
    h_todos = r["h_todos"]; h_ativos = r["h_ativos"]
    c1,c2 = st.columns(2)
    c1.metric("Horas CEN ativos (denominador operacional)", f"{h_ativos:.1f}h")
    c2.metric("Horas todos os funcionários (denominador total)", f"{h_todos:.1f}h")

    # ── PASSO 9: Produtividades ──────────────────────────────────────────────
    st.markdown('<div class="mem-step"><span class="step-num">9</span> <b>Produtividades</b></div>', unsafe_allow_html=True)
    st.markdown('<div class="formula-box">Ciclo Operacional = horas_ciclo ÷ horas_CEN_ativos<br>Ciclo Total       = horas_ciclo ÷ horas_todos<br>Labor Operacional = horas_labor ÷ horas_CEN_ativos<br>Labor Total ★     = horas_labor ÷ horas_todos</div>', unsafe_allow_html=True)
    h_ciclo = r["h_ciclo"]; h_labor = r["h_labor"]
    prod_data = {
        "Indicador": ["Ciclo Operacional","Ciclo Total","Labor Operacional","⭐ Labor Total"],
        "Numerador (h)": [f"{h_ciclo:.1f}", f"{h_ciclo:.1f}", f"{h_labor:.1f}", f"{h_labor:.1f}"],
        "Denominador (h)": [f"{h_ativos:.1f} (CEN ativos)", f"{h_todos:.1f} (todos)", f"{h_ativos:.1f} (CEN ativos)", f"{h_todos:.1f} (todos)"],
        "Resultado ✓": [f"{r['prod_ciclo_op']:.1%}", f"{r['prod_ciclo_tot']:.1%}",
                        f"{r['prod_labor_op']:.1%}", f"{r['prod_labor_tot']:.1%}"],
    }
    st.caption("Output do passo 9 — produtividades calculadas:")
    st.dataframe(pd.DataFrame(prod_data), use_container_width=True, hide_index=True)

    st.markdown("---")
    buf = BytesIO()
    df_intermediario[df_intermediario.mes==mes].to_excel(buf, index=False)
    buf.seek(0)
    st.download_button("📥 Baixar base completa pós-JOIN (todas as etapas)",
        data=buf, file_name=f"base_tratada_{mes}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ─────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────
st.markdown("""
<div class="jd-header">
  <h1>🏭 Calculadora de Recursos — Usinagem</h1>
  <p>Ferramenta de planejamento de headcount por turno · John Deere Manufatura</p>
</div>
""", unsafe_allow_html=True)

# ── GUIA DE PRIMEIROS PASSOS ──────────────────────────────────────────────────
with st.expander("👋 Primeira vez aqui? Veja como usar em 3 passos", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
<div class="mem-step">
  <span class="step-num">1</span> <b>Suba seu arquivo Excel</b><br><br>
  O mesmo arquivo que você já usa — com as abas de input preenchidas.<br><br>
  <b>O app vai ler automaticamente:</b><br>
  • <code>INPUT_PMP</code> — demanda mensal por modelo<br>
  • <code>IMPUTTEMPO</code> — tempo de ciclo e labor<br>
  • <code>IMPUTDISTRIBUIÇÃO</code> — fatores de carga<br>
  • <code>IMPUTAPLICAÇÃO</code> — quais modelos passam por cada centro<br>
  • <code>IMPUTTURNOS</code> — horas dos turnos
</div>
""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
<div class="mem-step">
  <span class="step-num">2</span> <b>Confira os resultados</b><br><br>
  A aba <b>📊 Resultados</b> mostra o headcount calculado por turno.<br><br>
  Vá para <b>🔄 Comparação</b> para ver se o resultado bate com o seu Excel atual — 
  o app lê as abas mensais (NovFY26, DezFY26…) e compara automaticamente.<br><br>
  <b>Verde ✅</b> = igual ao Excel<br>
  <b>Amarelo 🟡</b> = diferença pequena (até 2 pessoas)<br>
  <b>Vermelho 🔴</b> = divergência maior
</div>
""", unsafe_allow_html=True)
    with col3:
        st.markdown("""
<div class="mem-step">
  <span class="step-num">3</span> <b>Investigue divergências</b><br><br>
  Se algo divergir, clique no mês na aba <b>🔄 Comparação</b> para ver:<br><br>
  • Qual centro está diferente<br>
  • Em qual turno<br>
  • Se é diferença de volume, de fator ou de threshold<br>
  • <b>Onde exatamente corrigir</b> no seu Excel de input<br><br>
  Use a aba <b>🔬 Memória de Cálculo</b> para ver a fórmula passo a passo.
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="aviso-warn">
💡 <b>Dica rápida:</b> O app usa os dados das abas de input como fonte de verdade. 
Se o resultado divergir do Excel de referência, significa que algum dado de input pode estar diferente 
do que foi usado para gerar aquele Excel. A comparação te mostra exatamente onde.
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("Upload do arquivo de inputs (.xlsm ou .xlsx)", type=["xlsm","xlsx"])
if not uploaded:
    st.info("👆 Faça upload do arquivo para começar.")
    st.stop()

file_bytes = uploaded.read()

# Detectar arquivo novo por hash (evita re-leitura desnecessária)
_fhash = hashlib.md5(file_bytes).hexdigest()
if st.session_state.get("_fhash") != _fhash:
    st.session_state["_fhash"] = _fhash
    st.session_state["log_leitura"] = []
    for _k in ("real_resumo", "real_mensal", "turnos_arq"): st.session_state.pop(_k, None)

if "log_leitura" not in st.session_state:
    st.session_state.log_leitura = []

# Leitura cacheada: todas as abas em uma só chamada ──────────────────────────
@st.cache_data(show_spinner=False, ttl=3600)
def _ler_todas_abas(fb):
    _log = []
    _pmp, _dias = read_pmp(fb, _log)
    _tempo      = read_tempo(fb, _log)
    _dist       = read_dist(fb, _log)
    _aplic      = read_aplic(fb, _log)
    _turnos     = read_turnos(fb)
    _log.append(f"✅ IMPUTTURNOS lido: A={_turnos['A']}h · B={_turnos['B']}h · C={_turnos['C']}h (acumulados)")
    _log.append("✅ Leitura concluída")
    return _pmp, _dias, _tempo, _dist, _aplic, _turnos, _log

with st.spinner("Lendo planilha..."):
    try:
        pmp, dias, tempo, dist, aplic, turnos_arq, _log_c = _ler_todas_abas(file_bytes)
        st.session_state["turnos_arq"] = turnos_arq
        if not st.session_state.log_leitura:
            st.session_state.log_leitura = list(_log_c)
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
    st.caption("Horas acumuladas desde o início do dia — lidas do IMPUTTURNOS.")
    _def = st.session_state.get("turnos_arq", {"A":7.5,"B":14.25,"C":19.5})
    hA = st.number_input("Turno A (h acumulado)", value=_def["A"], step=0.01, format="%.2f")
    hB = st.number_input("Turno B (h acumulado)", value=_def["B"], step=0.01, format="%.2f")
    hC = st.number_input("Turno C (h acumulado)", value=_def["C"], step=0.01, format="%.2f")
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
    st.caption("⚠️ Se não houver operador CEN ativo num turno, o suporte desse turno é automaticamente zero.")
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
                st.caption(f"Padrão: A={defs['A']} · B={defs['B']} · C={defs['C']}")
                suporte_cfg[key]={"modo":"auto",**defs}
            else:
                st.caption("Define quantos por turno **quando o turno estiver ativo**. Se não houver CEN no turno, fica 0.")
                c1,c2,c3=st.columns(3)
                vA=c1.number_input("A",0,10,defs["A"],key=f"s_{key}_A")
                vB=c2.number_input("B",0,10,defs["B"],key=f"s_{key}_B")
                vC=c3.number_input("C",0,10,defs["C"],key=f"s_{key}_C")
                suporte_cfg[key]={"modo":"manual","A":vA,"B":vB,"C":vC}
    st.markdown("---")
    st.caption("Alterações afetam todos os cálculos em tempo real.")

# ─── TABS ────────────────────────────────
tab_vis, tab_inp, tab_mem, tab_res, tab_cmp, tab_diag, tab_exp = st.tabs([
    "🏠 Visão Geral", "📂 Dados de Input", "🔬 Como foi Calculado",
    "📊 Resultado por Mês", "🔄 Comparar com Excel", "🩺 Diagnóstico", "📥 Exportar"
])

# Hash rápido — pmp.to_json() é lento para DataFrames grandes
pmp_hash  = hashlib.md5(pd.util.hash_pandas_object(pmp, index=True).values.tobytes()).hexdigest()
dias_hash = hashlib.md5(str(dias).encode()).hexdigest()
import json as _json
_sup_str = _json.dumps(suporte_cfg, sort_keys=True)
res_base, df_interm, agg_interm = _calcular_cached(
    pmp_hash, pmp, tempo, dist, aplic, dias_hash, dias,
    horas_turno["A"], horas_turno["B"], horas_turno["C"],
    thresholds["A"], thresholds["B"], thresholds["C"],
    _sup_str)

# ── Dados reais carregados sob demanda (primeira vez que a aba precisar) ──────
# (não bloqueia o startup)
def _load_real_data_once():
    if "real_resumo" not in st.session_state:
        _real_resumo, _real_mensal = read_real_data(file_bytes)
        st.session_state["real_resumo"] = _real_resumo
        st.session_state["real_mensal"] = _real_mensal

# ══════════════════════════════════════════
# TAB 1 — VISÃO GERAL
# ══════════════════════════════════════════
with tab_vis:
    st.plotly_chart(grafico_cenarios({"Base": res_base}), use_container_width=True)

    with st.expander("ℹ️ Como ler este gráfico"):
        st.markdown("""
**Barras empilhadas** = total de funcionários por turno em cada mês.
- 🟢 Verde = Turno A · 🟡 Amarelo = Turno B · 🔵 Azul = Turno C
- Inclui operadores CEN + suporte (Lavadora, Gravação, Preset, Coringa, Facilitador)

**Linha vermelha (%)** = Produtividade Labor Total — tempo produtivo ÷ tempo total disponível.
Quanto maior, melhor. Valores baixos indicam turno ocioso ou suporte superdimensionado.
        """)

    meses_ok = [m for m in MESES if res_base.get(m)]
    if meses_ok:
        media_labor = np.mean([res_base[m]["prod_labor_tot"] for m in meses_ok])
        max_total   = max(res_base[m]["total"] for m in meses_ok)
        min_total   = min(res_base[m]["total"] for m in meses_ok)
        mes_pico    = max(meses_ok, key=lambda m: res_base[m]["total"])
        mes_vale    = min(meses_ok, key=lambda m: res_base[m]["total"])
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Meses calculados", len(meses_ok))
        c2.metric("⭐ Labor Total médio", f"{media_labor:.0%}",
                  help="Produtividade média anual. Meta: quanto maior, melhor.")
        c3.metric("Pico de headcount", f"{max_total} func.",
                  delta=f"em {mes_pico[:3].upper()}",
                  help="Mês com maior necessidade de pessoal.")
        c4.metric("Variação anual", f"{max_total - min_total} func.",
                  help=f"Diferença entre o mês de pico ({mes_pico[:3].upper()}) e o menor ({mes_vale[:3].upper()}).")

        st.markdown('<div class="jd-section">Alertas automáticos</div>', unsafe_allow_html=True)
        alertas = []
        for m in meses_ok:
            r = res_base[m]
            if r["op_C"] > 5:
                alertas.append(f"⚠️ **{m}** — Turno C com {r['op_C']} centros ativos. Verifique fatores de distribuição no IMPUTDISTRIBUIÇÃO.")
            if r["prod_labor_tot"] < 0.30:
                alertas.append(f"⚠️ **{m}** — Labor Total baixo ({r['prod_labor_tot']:.0%}). Pode indicar turno superdimensionado.")
            if r["op_B"] == 0 and r["tot_A"] > 0:
                alertas.append(f"ℹ️ **{m}** — Nenhum centro ativo no Turno B. Apenas Turno A operando.")
        if alertas:
            for a in alertas:
                st.markdown(f'<div class="aviso-warn">{a}</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="aviso-ok">✅ Nenhum alerta identificado nos dados calculados.</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════
# TAB 2 — INPUTS
# ══════════════════════════════════════════
with tab_inp:
    st.markdown('<div class="jd-section">Dados carregados do seu arquivo</div>', unsafe_allow_html=True)
    st.caption("Verifique aqui se os dados foram lidos corretamente. Se algo parecer errado, o problema está na planilha de input.")
    aba_inp = st.radio("Qual dado você quer conferir?", [
        "INPUT_PMP — Demanda por modelo",
        "IMPUTTEMPO — Tempo de ciclo e labor",
        "IMPUTDISTRIBUIÇÃO — Fatores de carga",
        "IMPUTAPLICAÇÃO — Quais modelos por centro"
    ], horizontal=True)
    aba_inp = aba_inp.split(" — ")[0]  # pegar só o nome da aba
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
    _load_real_data_once()  # carrega sob demanda, sem bloquear startup
    if "cenarios" not in st.session_state:
        st.session_state.cenarios = {}

    st.markdown('<div class="jd-section">Resultado por mês</div>', unsafe_allow_html=True)
    mes_r = st.selectbox("Selecione o mês", [m for m in MESES if res_base.get(m)], key="mes_r")

    if mes_r and res_base.get(mes_r):
        r = res_base[mes_r]
        # Resumo rápido antes da tabela detalhada
        st.markdown(f"""
<div class="aviso-ok">
📋 <b>{mes_r}</b> — {r['dias']} dias trabalhados &nbsp;|&nbsp;
Turno A: <b>{r['tot_A']} pessoas</b> &nbsp;|&nbsp;
Turno B: <b>{r['tot_B']} pessoas</b> &nbsp;|&nbsp;
Turno C: <b>{r['tot_C']} pessoas</b> &nbsp;|&nbsp;
<b>Total: {r['total']} funcionários</b>
</div>
        """, unsafe_allow_html=True)

        with st.expander("ℹ️ Como ler a tabela abaixo"):
            st.markdown("""
**Tabela de cima — Por centro de usinagem (CEN):**
- **% Ocupação A/B/C** = quanto da capacidade do turno está sendo usada por este centro
  - 🟢 Verde = abaixo de 85% (confortável)
  - 🟡 Amarelo = entre 85% e 100% (atenção)
  - 🔴 Vermelho = acima de 100% (sobrecarregado)
- **Ativo 0/1** = se o turno está aberto para este centro (1=aberto, 0=fechado)
- **Horas** = horas disponíveis no mês para aquele centro naquele turno

**Tabela de baixo — Funções de suporte:**
- Mostra quantas pessoas de cada função estão alocadas por turno
- A linha **TOTAL POR TURNO** já inclui operadores CEN + suporte
            """)

        show_tabela(r, real_mensal=st.session_state.get("real_mensal"), mes=mes_r)

    # ── FY26 ANUAL ──────────────────────────────────────────────────────
    st.markdown("---")
    show_fy26_anual(res_base, st.session_state.get("real_resumo", {}), dias, horas_turno)

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
                vA=c1.number_input("Turno A",0,5,int(rc.ativo_A),key=f"n_{novo_nome}_{cen}_A",label_visibility="collapsed",help=f"Base:{rc.ativo_A}")
                vB=c2.number_input("Turno B",0,5,int(rc.ativo_B),key=f"n_{novo_nome}_{cen}_B",label_visibility="collapsed",help=f"Base:{rc.ativo_B}")
                vC=c3.number_input("Turno C",0,5,int(rc.ativo_C),key=f"n_{novo_nome}_{cen}_C",label_visibility="collapsed",help=f"Base:{rc.ativo_C}")
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
    st.markdown('<div class="jd-section">Comparação com o seu Excel atual</div>', unsafe_allow_html=True)
    st.markdown("""
O app lê as abas mensais do arquivo que você subiu (NovFY26, DezFY26…) e compara o resultado 
calculado pelos inputs com o que está nessas abas.

**Se divergir**, não significa que o app está errado — significa que algum dado de input 
está diferente do que foi usado para gerar aquele Excel. O diagnóstico abaixo te mostra exatamente onde.
    """)

    # Comparação: roda só quando o usuário clicar
    _cmp_key = f"cmp_{hash(str(dias))}_{hash(str(thresholds))}_{hash(str(horas_turno))}"
    if st.button("🔄 Comparar App com Excel agora", key="btn_comparar", type="primary"):
        with st.spinner("Comparando com o Excel..."):
            _r, _d, _e = comparar_com_excel(res_base, file_bytes, tempo, dist, aplic, pmp, dias, horas_turno, thresholds, suporte_cfg)
        st.session_state["cmp_cache_key"]     = _cmp_key
        st.session_state["cmp_cache_resumo"]  = _r
        st.session_state["cmp_cache_detalhe"] = _d
        st.session_state["cmp_cache_err"]     = _e
    elif st.session_state.get("cmp_cache_key") != _cmp_key:
        # parâmetros mudaram — avisar sem rodar automaticamente
        st.info("⚠️ Parâmetros mudaram. Clique em **Comparar** para atualizar.")

    df_resumo  = st.session_state.get("cmp_cache_resumo")
    df_detalhe = st.session_state.get("cmp_cache_detalhe")
    err        = st.session_state.get("cmp_cache_err")

    if err:
        st.error(err)
    elif df_resumo is not None and not df_resumo.empty:
        n_ok   = (df_resumo["Status"].str.startswith("✅")).sum() if "Status" in df_resumo else 0
        n_warn = (df_resumo["Status"].str.startswith("🟡")).sum() if "Status" in df_resumo else 0
        n_err  = (df_resumo["Status"].str.startswith("🔴")).sum() if "Status" in df_resumo else 0
        n_aus  = (df_resumo["Status"].str.startswith("⚠️")).sum() if "Status" in df_resumo else 0

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("✅ Meses iguais",         n_ok)
        c2.metric("🟡 Pequena diferença",    n_warn)
        c3.metric("🔴 Com divergência",      n_err)
        c4.metric("⚠️ Abas não encontradas", n_aus)

        # Tabela resumo
        st.markdown('<div class="jd-sub">Resumo por mês</div>', unsafe_allow_html=True)

        # Construir tabela visual por turno
        def cor_delta(d):
            try:
                n = int(str(d).replace("+",""))
                if n == 0:  return "✅"
                if abs(n) <= 2: return "🟡"
                return "🔴"
            except: return ""

        def build_resumo_visual(df):
            rows = []
            for _, r in df.iterrows():
                def cell(app, excel, delta):
                    icon = cor_delta(delta)
                    try:
                        d_num = int(str(delta).replace("+",""))
                        if d_num == 0: return f"✅ {app}"
                        return f"{icon} App={app}  |  Excel={excel}  ({delta})"
                    except:
                        return f"{icon} App={app} / Excel={excel} ({delta})"
                rows.append({
                    "Mês":     r["Mês"],
                    "Status":  r["Status"],
                    "Turno A": cell(r.get("CEN A App","?"), r.get("CEN A Excel","?"), r.get("Δ A","?")),
                    "Turno B": cell(r.get("CEN B App","?"), r.get("CEN B Excel","?"), r.get("Δ B","?")),
                    "Turno C": cell(r.get("CEN C App","?"), r.get("CEN C Excel","?"), r.get("Δ C","?")),
                    "Total":   cell(r.get("Total App","?"), r.get("Total Excel","?"), r.get("Δ Total","?")),
                    "Labor":   f"App={r.get('Labor App','?')} / Excel={r.get('Labor Excel','?')}",
                })
            return pd.DataFrame(rows)

        df_vis = build_resumo_visual(df_resumo)

        def style_resumo(row):
            st_val = str(row.get("Status",""))
            if "✅" in st_val:   base = "background-color:#003D10;color:#B9F6CA"
            elif "🟡" in st_val: base = "background-color:#3D2D00;color:#FFE57F"
            elif "🔴" in st_val: base = "background-color:#3D0000;color:#FF8A80"
            else:                base = "background-color:#2D1A00;color:#FFD54F"

            styles = []
            for col in row.index:
                if col == "Status":
                    styles.append(base)
                elif col in ("Turno A","Turno B","Turno C","Total"):
                    val = str(row[col])
                    if "🔴" in val:   styles.append("background-color:#3D0000;color:#FF8A80")
                    elif "🟡" in val: styles.append("background-color:#3D2D00;color:#FFE57F")
                    elif "✅" in val: styles.append("background-color:#003D10;color:#B9F6CA")
                    else:             styles.append("")
                else:
                    styles.append("")
            return styles

        # Ensure all columns are string type to avoid Arrow type inference issues
        for _col in ["Turno A","Turno B","Turno C","Total","Labor"]:
            if _col in df_vis.columns:
                df_vis[_col] = df_vis[_col].astype(str)
        st.dataframe(
            df_vis.style.apply(style_resumo, axis=1),
            use_container_width=True, hide_index=True
        )

        # Detalhe das divergências
        if df_detalhe is not None and len(df_detalhe) > 0:
            st.markdown('<div class="jd-sub">Detalhamento das divergências por centro</div>', unsafe_allow_html=True)
            st.caption(f"{len(df_detalhe)} divergência(s) encontrada(s) em {len(df_detalhe['Mês'].unique())} mês(es). Selecione o mês para ver a análise detalhada.")

            meses_div = df_detalhe["Mês"].unique().tolist()
            mes_sel_div = st.selectbox("Ver detalhes do mês", meses_div, key="mes_div")

            df_mes = df_detalhe[df_detalhe["Mês"] == mes_sel_div].reset_index(drop=True)

            # Tabela resumida por centro
            cols_tab = ["Centro","Turno","App — Ativo","Excel — Ativo",
                        "Ocup. App","Ocup. Excel","Δ Ocupação","Causa"]
            cols_tab = [c for c in cols_tab if c in df_mes.columns]

            def style_det(row):
                causa = str(row.get("Causa","")).lower()
                if "menor" in causa or "maior" in causa:
                    return ["background-color:#3D0000;color:#FF8A80"]*len(row)
                if "índice" in causa or "indice" in causa:
                    return ["background-color:#3D1A00;color:#FFAB40"]*len(row)
                return ["background-color:#3D2D00;color:#FFE57F"]*len(row)

            st.dataframe(
                df_mes[cols_tab].style.apply(style_det, axis=1),
                use_container_width=True, hide_index=True
            )

            st.markdown('<div class="jd-sub">Análise detalhada — clique para expandir cada divergência</div>', unsafe_allow_html=True)
            for _, row in df_mes.iterrows():
                causa = str(row.get("Causa",""))
                icon  = "🔴" if "volume" in causa.lower() else ("🟠" if "índice" in causa.lower() or "indice" in causa.lower() else "🟡")
                label = f"{icon} {row.get('Centro','')} — Turno {row.get('Turno','')}: {causa}"
                with st.expander(label):
                    c1,c2,c3,c4 = st.columns(4)
                    c1.metric("App — Ativo",    row.get("App — Ativo",""))
                    c2.metric("Excel — Ativo",  row.get("Excel — Ativo",""))
                    c3.metric("Ocupação App",   row.get("Ocup. App",""))
                    c4.metric("Ocupação Excel", row.get("Ocup. Excel",""))
                    onde = row.get("Onde investigar","")
                    expl = row.get("Explicação","")
                    if onde:
                        st.markdown(f'<div class="aviso-warn">📍 <b>Onde investigar:</b> {onde}</div>', unsafe_allow_html=True)
                    if expl:
                        st.markdown(f'<div class="aviso-warn">🔍 <b>O que aconteceu:</b> {expl}</div>', unsafe_allow_html=True)

        # Export
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
            if df_detalhe is not None and len(df_detalhe) > 0:
                df_detalhe.to_excel(writer, sheet_name="Divergências", index=False)
        buf.seek(0)
        st.download_button("📥 Exportar comparação completa",
            data=buf, file_name="comparacao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Nenhuma aba mensal encontrada no arquivo (NovFY26, DezFY26 etc.). O arquivo precisa ter as abas mensais calculadas para a comparação funcionar.")



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

    sub_cmp, sub_res = st.tabs(["🔍 Comparação e Diagnóstico", "📊 Resultados"])

    # ══════════════════════════════════════════
    # ══════════════════════════════════════════
    # SUB-ABA 1 — COMPARAÇÃO E DIAGNÓSTICO
    # ══════════════════════════════════════════
    with sub_cmp:
        # ── Cache key: regenera só quando parâmetros mudam ──────────────────
        _exp_key = f"exp_{hash(str(dias))}_{hash(str(thresholds))}_{hash(str(horas_turno))}"

        st.markdown('<div class="jd-sub">🔴 Diagnóstico de divergências</div>', unsafe_allow_html=True)
        st.markdown("Gera um Excel com **uma aba por mês** mostrando, centro a centro, onde o App diverge do Excel de referência.")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            if st.button("🔴 Gerar diagnóstico por mês", key="btn_diag_mensal", type="primary"):
                with st.spinner("Gerando (~6s)..."):
                    st.session_state["_diag_mensal"] = gerar_diagnostico_mensal(
                        file_bytes, res_base, tempo, dist, aplic, pmp, dias, horas_turno, thresholds)
                    st.session_state["_diag_mensal_key"] = _exp_key
            if st.session_state.get("_diag_mensal"):
                st.download_button("📥 Baixar diagnóstico por mês",
                    data=st.session_state["_diag_mensal"],
                    file_name="diagnostico_por_mes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_diag_mensal")
        with col_d2:
            if st.button("🔍 Gerar diagnóstico de inputs", key="btn_diag_inp"):
                with st.spinner("Gerando..."):
                    st.session_state["_diag_inp"] = gerar_excel_diagnostico(file_bytes)
            if st.session_state.get("_diag_inp"):
                st.download_button("📥 Baixar diagnóstico de inputs",
                    data=st.session_state["_diag_inp"],
                    file_name="diagnostico_inputs.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_diag_inp")

        st.markdown("---")
        st.markdown('<div class="jd-sub">📊 Resultado no layout do Excel original</div>', unsafe_allow_html=True)
        st.markdown("Gera um Excel **idêntico ao layout do seu Excel de referência** com células em **vermelho** onde diverge.")

        if st.button("📊 Gerar resultado no layout do Excel", key="btn_layout", type="primary"):
            with st.spinner("Gerando (~8s)..."):
                st.session_state["_layout_data"] = gerar_output_layout(
                    file_bytes, res_base, tempo, dist, aplic, pmp, dias, horas_turno, thresholds, suporte_cfg)
                st.session_state["_layout_key"] = _exp_key
        if st.session_state.get("_layout_data"):
            st.download_button("📥 Baixar resultado no layout do Excel",
                data=st.session_state["_layout_data"],
                file_name="resultado_layout_excel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_layout")

        st.markdown("---")
        st.markdown('<div class="jd-sub">📋 Tabelona completa — layout IMPUTDISTRIBUIÇÃO</div>', unsafe_allow_html=True)
        st.markdown("Gera a **tabelona completa** com as colunas de % ocupação + bloco DADOS AUTOMÁTICOS. Células em **vermelho** = divergência.")

        if st.button("📋 Gerar tabelona completa", key="btn_tabelona", type="primary"):
            with st.spinner("Gerando (~10s)..."):
                st.session_state["_tabelona"] = _gerar_tabelona(
                    file_bytes, res_base, tempo, dist, aplic, pmp, dias, horas_turno, thresholds, suporte_cfg)
                st.session_state["_tabelona_key"] = _exp_key
        if st.session_state.get("_tabelona"):
            st.download_button("📥 Baixar tabelona",
                data=st.session_state["_tabelona"],
                file_name="tabelona_por_mes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_tabelona")

    # ══════════════════════════════════════════
    # SUB-ABA 2 — RESULTADOS REAIS
    # ══════════════════════════════════════════
    with sub_res:
        st.markdown('<div class="jd-sub">📋 Resultado completo e base tratada</div>', unsafe_allow_html=True)
        st.markdown("Exportações dos **resultados reais** calculados pelo App, sem comparação com o Excel de referência.")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Resultado completo (todas as abas)**")
            st.caption("RESUMO MO + uma aba por mês com tabela no formato original")
            if st.button("📥 Gerar resultado base", key="btn_res_base"):
                with st.spinner("Gerando..."):
                    st.session_state["_res_base_xl"] = exportar(res_base)
            if st.session_state.get("_res_base_xl"):
                st.download_button("📥 Baixar resultado base",
                    data=st.session_state["_res_base_xl"], file_name="resultado_usinagem.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_res_base")
        with c2:
            st.markdown("**Base tratada (pós-JOIN completo)**")
            st.caption("Todos os meses, todas as linhas de centro+peça+modelo+minutos")
            if st.button("📥 Gerar base tratada", key="btn_base_tratada"):
                with st.spinner("Gerando..."):
                    _buf = BytesIO(); df_interm.to_excel(_buf, index=False); _buf.seek(0)
                    st.session_state["_base_tratada"] = _buf
            if st.session_state.get("_base_tratada"):
                st.download_button("📥 Baixar base tratada completa",
                    data=st.session_state["_base_tratada"], file_name="base_tratada_completa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_base_tratada")
        if st.session_state.get("cenarios"):
            st.markdown('<div class="jd-sub">Cenários salvos</div>', unsafe_allow_html=True)
            for nm, v in st.session_state.cenarios.items():
                if st.button(f"📥 Gerar cenário: {nm}", key=f"btn_exp_{nm}"):
                    with st.spinner("Gerando..."):
                        st.session_state[f"_cen_{nm}"] = exportar(v["resultados"])
                if st.session_state.get(f"_cen_{nm}"):
                    st.download_button(f"📥 Baixar cenário: {nm}",
                        data=st.session_state[f"_cen_{nm}"],
                        file_name=f"cenario_{nm.replace(' ','_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"exp_{nm}")
