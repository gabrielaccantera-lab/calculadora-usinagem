"""
MELHORIAS — Cole cada bloco no lugar correto do app principal.
Há 3 substituições a fazer:

1. CSS extra — adicione dentro do st.markdown(css) existente, antes do </style>
2. TAB RESULTADOS — substitua o bloco `elif mes_r == "📅 ANO (resumo anual)":` 
3. TAB CENÁRIOS — substitua TODA a seção `with tab_cen:`

Cada bloco está identificado abaixo.
"""

# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 1 — CSS EXTRA (adicionar no final do CSS existente, antes de </style>)
# ══════════════════════════════════════════════════════════════════════════════
CSS_EXTRA = """
.kpi-card {
    background: linear-gradient(135deg, #1A2E1A 0%, #0D1F0D 100%);
    border: 1px solid #2A4A2A;
    border-radius: 12px;
    padding: 16px 18px;
    margin: 4px 0;
    position: relative;
    overflow: hidden;
}
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 4px; height: 100%;
    background: #FFDE00;
}
.kpi-card .kpi-icon { font-size: 22px; margin-bottom: 4px; }
.kpi-card .kpi-label { font-size: 10px; color: #7BC67A; text-transform: uppercase; letter-spacing: .06em; font-weight: 600; }
.kpi-card .kpi-value { font-size: 26px; font-weight: 800; color: #FFFFFF; line-height: 1.1; }
.kpi-card .kpi-sub   { font-size: 11px; color: #AAAAAA; margin-top: 2px; }
.kpi-card.destaque { border-color: #FFDE00; }
.kpi-card.destaque::before { background: #367C2B; width: 100%; height: 3px; top: 0; left: 0; }

.gauge-wrap { text-align: center; padding: 8px 0; }
.gauge-label { font-size: 11px; color: #7BC67A; text-transform: uppercase; letter-spacing:.05em; font-weight:600; margin-bottom:4px; }
.gauge-pct   { font-size: 32px; font-weight: 900; color: #FFDE00; }
.gauge-sub   { font-size: 10px; color: #888; }

.mes-row { 
    display: flex; align-items: center; gap: 10px;
    padding: 7px 12px; border-radius: 8px; margin: 3px 0;
    background: #131313; border: 1px solid #222;
    transition: border-color .15s;
}
.mes-row:hover { border-color: #FFDE00; }
.mes-row .mes-nome { font-size: 12px; font-weight: 700; color: #FFDE00; min-width: 80px; }
.mes-row .mes-bar  { flex: 1; height: 8px; background: #1F4D19; border-radius: 4px; position: relative; }
.mes-row .mes-bar-fill { height: 100%; border-radius: 4px; background: linear-gradient(90deg, #367C2B, #FFDE00); }
.mes-row .mes-num  { font-size: 12px; color: #FFFFFF; font-weight: 700; min-width: 28px; text-align: right; }
.mes-row .mes-labor{ font-size: 11px; color: #7BC67A; min-width: 42px; text-align: right; }

.turno-pill {
    display: inline-block; border-radius: 20px;
    padding: 2px 10px; font-size: 11px; font-weight: 700; margin: 0 2px;
}
.turno-A { background:#1F4D19; color:#92D050; }
.turno-B { background:#3D2D00; color:#FFDE00; }
.turno-C { background:#0D2040; color:#00B0F0; }

.cenario-badge {
    display:inline-block; padding:2px 8px; border-radius:4px;
    background:#FFDE00; color:#1F4D19; font-size:10px; font-weight:700;
    margin-left:6px; vertical-align:middle;
}
"""

# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 2 — RESULTADOS ANO (substitua o elif mes_r == "📅 ANO..." INTEIRO)
# ══════════════════════════════════════════════════════════════════════════════
RESULTADOS_ANO = """
    elif mes_r and mes_r == "📅 ANO (resumo anual)":
        _meses_ano_r = [(m, res_base[m]) for m in MESES if res_base.get(m)]
        if _meses_ano_r:
            _n_r  = len(_meses_ano_r)
            _d_r  = sum(r["dias"] for _, r in _meses_ano_r)
            _shc_r = sum(r["h_ciclo"]  for _, r in _meses_ano_r)
            _shl_r = sum(r["h_labor"]  for _, r in _meses_ano_r)
            _sha_r = sum(r["h_ativos"] for _, r in _meses_ano_r)
            _sht_r = sum(r["h_todos"]  for _, r in _meses_ano_r)
            _pl_r  = _shl_r / _sht_r  if _sht_r  > 0 else 0
            _pc_r  = _shc_r / _sht_r  if _sht_r  > 0 else 0
            _plo_r = _shl_r / _sha_r  if _sha_r  > 0 else 0
            _tA_r  = round(sum(r["tot_A"] for _, r in _meses_ano_r) / _n_r, 1)
            _tB_r  = round(sum(r["tot_B"] for _, r in _meses_ano_r) / _n_r, 1)
            _tC_r  = round(sum(r["tot_C"] for _, r in _meses_ano_r) / _n_r, 1)
            _tf_r  = round(sum(r["total"] for _, r in _meses_ano_r) / _n_r, 1)
            _pico_r = max(_meses_ano_r, key=lambda x: x[1]["total"])
            _vale_r = min(_meses_ano_r, key=lambda x: x[1]["total"])
            _max_tot = max(r["total"] for _, r in _meses_ano_r)

            # ── HEADER ────────────────────────────────────────────────────────
            st.markdown(f'''
<div style="background:linear-gradient(135deg,#1F4D19,#0D2A0D);border-radius:12px;
            padding:16px 22px;margin-bottom:18px;border-left:5px solid #FFDE00;">
  <div style="font-size:13px;color:#7BC67A;font-weight:700;text-transform:uppercase;
              letter-spacing:.08em;margin-bottom:4px;">📅 Visão Anual</div>
  <div style="font-size:20px;font-weight:800;color:#FFFFFF;">
    {_n_r} meses calculados · {_d_r} dias trabalhados
  </div>
  <div style="font-size:12px;color:#AAAAAA;margin-top:4px;">
    Pico: <b style="color:#FFDE00">{_pico_r[0][:3].upper()} ({_pico_r[1]["total"]} func.)</b> &nbsp;·&nbsp;
    Vale: <b style="color:#7BC67A">{_vale_r[0][:3].upper()} ({_vale_r[1]["total"]} func.)</b>
  </div>
</div>
''', unsafe_allow_html=True)

            # ── KPI CARDS ──────────────────────────────────────────────────────
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown(f'''
<div class="kpi-card destaque">
  <div class="kpi-icon">⭐</div>
  <div class="kpi-label">Labor Total Anual</div>
  <div class="kpi-value">{_pl_r:.1%}</div>
  <div class="kpi-sub">Produtividade real de toda a equipe</div>
</div>''', unsafe_allow_html=True)
            with c2:
                st.markdown(f'''
<div class="kpi-card">
  <div class="kpi-icon">👷</div>
  <div class="kpi-label">Média de Funcionários / Mês</div>
  <div class="kpi-value">{_tf_r:.0f}</div>
  <div class="kpi-sub">
    <span class="turno-pill turno-A">A {_tA_r:.0f}</span>
    <span class="turno-pill turno-B">B {_tB_r:.0f}</span>
    <span class="turno-pill turno-C">C {_tC_r:.0f}</span>
  </div>
</div>''', unsafe_allow_html=True)
            with c3:
                st.markdown(f'''
<div class="kpi-card">
  <div class="kpi-icon">🔄</div>
  <div class="kpi-label">Ciclo Total Anual</div>
  <div class="kpi-value">{_pc_r:.1%}</div>
  <div class="kpi-sub">Labor Operacional: {_plo_r:.1%}</div>
</div>''', unsafe_allow_html=True)
            with c4:
                variacao = _max_tot - min(r["total"] for _, r in _meses_ano_r)
                st.markdown(f'''
<div class="kpi-card">
  <div class="kpi-icon">📈</div>
  <div class="kpi-label">Variação Anual</div>
  <div class="kpi-value">{variacao}</div>
  <div class="kpi-sub">func. entre pico e vale</div>
</div>''', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ── GAUGE LABOR + BARRAS POR MÊS lado a lado ─────────────────────
            col_gauge, col_bars = st.columns([1, 2])

            with col_gauge:
                st.markdown('<div class="jd-sub">Labor Total Anual</div>', unsafe_allow_html=True)
                pct_int = int(_pl_r * 100)
                cor_gauge = "#69F0AE" if _pl_r >= 0.45 else ("#FFDE00" if _pl_r >= 0.30 else "#FF5252")
                # SVG gauge semicircle
                ang = min(180, int(_pl_r * 180))
                import math
                rad = math.radians(ang)
                cx, cy, r_g = 80, 80, 60
                ex = cx + r_g * math.cos(math.pi - rad)
                ey = cy - r_g * math.sin(rad)
                st.markdown(f'''
<div class="gauge-wrap">
  <svg viewBox="0 0 160 100" width="180">
    <path d="M 20 80 A 60 60 0 0 1 140 80" fill="none" stroke="#2A2A2A" stroke-width="12" stroke-linecap="round"/>
    <path d="M 20 80 A 60 60 0 {"1" if ang > 90 else "0"} 1 {ex:.1f} {ey:.1f}"
          fill="none" stroke="{cor_gauge}" stroke-width="12" stroke-linecap="round"/>
    <text x="80" y="76" text-anchor="middle" font-size="20" font-weight="900" fill="{cor_gauge}">{pct_int}%</text>
    <text x="80" y="92" text-anchor="middle" font-size="9" fill="#888">Labor Total</text>
  </svg>
  <div style="font-size:11px;color:#7BC67A;margin-top:4px;">
    Meta: manter acima de 40%
  </div>
</div>''', unsafe_allow_html=True)

                # Mini gauge ciclo
                pct_ciclo = int(_pc_r * 100)
                ang2 = min(180, int(_pc_r * 180))
                rad2 = math.radians(ang2)
                ex2 = cx + r_g * math.cos(math.pi - rad2)
                ey2 = cy - r_g * math.sin(rad2)
                cor2 = "#7BC67A"
                st.markdown(f'''
<div class="gauge-wrap" style="margin-top:8px;">
  <svg viewBox="0 0 160 100" width="140">
    <path d="M 20 80 A 60 60 0 0 1 140 80" fill="none" stroke="#2A2A2A" stroke-width="10" stroke-linecap="round"/>
    <path d="M 20 80 A 60 60 0 {"1" if ang2 > 90 else "0"} 1 {ex2:.1f} {ey2:.1f}"
          fill="none" stroke="{cor2}" stroke-width="10" stroke-linecap="round"/>
    <text x="80" y="76" text-anchor="middle" font-size="18" font-weight="800" fill="{cor2}">{pct_ciclo}%</text>
    <text x="80" y="92" text-anchor="middle" font-size="9" fill="#888">Ciclo Total</text>
  </svg>
</div>''', unsafe_allow_html=True)

            with col_bars:
                st.markdown('<div class="jd-sub">Funcionários por Mês</div>', unsafe_allow_html=True)
                _max_bar = max(r["total"] for _, r in _meses_ano_r) or 1
                bars_html = ""
                for _m, _r in _meses_ano_r:
                    _pct_bar = _r["total"] / _max_bar * 100
                    _lab = f'{_r["prod_labor_tot"]:.0%}'
                    _lab_cor = "#69F0AE" if _r["prod_labor_tot"] >= 0.45 else ("#FFDE00" if _r["prod_labor_tot"] >= 0.30 else "#FF5252")
                    bars_html += f'''
<div class="mes-row">
  <div class="mes-nome">{_m[:3].upper()}</div>
  <div style="flex:1;display:flex;align-items:center;gap:6px;">
    <div class="mes-bar"><div class="mes-bar-fill" style="width:{_pct_bar:.0f}%"></div></div>
    <div style="display:flex;gap:4px;min-width:90px;font-size:10px;">
      <span class="turno-pill turno-A">{_r["tot_A"]}</span>
      <span class="turno-pill turno-B">{_r["tot_B"]}</span>
      <span class="turno-pill turno-C">{_r["tot_C"]}</span>
    </div>
  </div>
  <div class="mes-num">{_r["total"]}</div>
  <div class="mes-labor" style="color:{_lab_cor}">{_lab}</div>
</div>'''
                st.markdown(bars_html, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ── TABELA RESUMO ANUAL ────────────────────────────────────────────
            st.markdown('<div class="jd-sub">Tabela detalhada</div>', unsafe_allow_html=True)
            _rows_r = []
            for _m, _r in _meses_ano_r:
                _rows_r.append({
                    "Mês": _m, "Dias": _r["dias"],
                    "Turno A": _r["tot_A"], "Turno B": _r["tot_B"], "Turno C": _r["tot_C"],
                    "Total": _r["total"],
                    "Labor Total": f'{_r["prod_labor_tot"]:.1%}',
                    "Labor Op.":   f'{_r["prod_labor_op"]:.1%}',
                    "Ciclo Total": f'{_r["prod_ciclo_tot"]:.1%}'
                })
            _rows_r.append({
                "Mês": "📅 MÉDIA ANO", "Dias": _d_r,
                "Turno A": _tA_r, "Turno B": _tB_r, "Turno C": _tC_r, "Total": _tf_r,
                "Labor Total": f'{_pl_r:.1%}', "Labor Op.": f'{_plo_r:.1%}',
                "Ciclo Total": f'{_pc_r:.1%}'
            })
            def _sty_r(row):
                if "ANO" in str(row["Mês"]):
                    return [f"background-color:{JD_AMARELO};color:{JD_VERDE_ESC};font-weight:700"] * len(row)
                # Colorir Labor Total
                styles = [""] * len(row)
                try:
                    lab_v = float(str(row["Labor Total"]).strip("%")) / 100
                    cor = "#003D10" if lab_v >= 0.45 else ("#3D2D00" if lab_v >= 0.30 else "#3D0000")
                    txt = "#B9F6CA" if lab_v >= 0.45 else ("#FFE57F" if lab_v >= 0.30 else "#FF8A80")
                    for i, col in enumerate(pd.DataFrame(_rows_r).columns):
                        if col == "Labor Total":
                            styles[i] = f"background-color:{cor};color:{txt};font-weight:600"
                except: pass
                return styles
            st.dataframe(
                pd.DataFrame(_rows_r).style.apply(_sty_r, axis=1),
                use_container_width=True, hide_index=True
            )
"""

# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 3 — TAB CENÁRIOS COMPLETA (substitua TODA a seção `with tab_cen:`)
# ══════════════════════════════════════════════════════════════════════════════
CENARIOS_TAB = """
with tab_cen:
    if "cenarios" not in st.session_state:
        st.session_state.cenarios = {}

    st.markdown('<div class="jd-section">Simulador de cenários</div>', unsafe_allow_html=True)
    st.caption("Crie variações do resultado base alterando quais turnos ficam ativos por centro. Compare lado a lado com o cenário atual.")

    with st.expander("➕ Criar novo cenário", expanded=len(st.session_state.cenarios)==0):
        ca, cb = st.columns([2, 1])
        novo_nome = ca.text_input("Nome do cenário", placeholder="Ex: Redução turno B novembro")
        opcoes_cenario = [m for m in MESES if res_base.get(m)] + ["📅 ANO (todos os meses)"]
        mes_novo = cb.selectbox("Mês base", opcoes_cenario, key="mes_novo")

        eh_ano = mes_novo == "📅 ANO (todos os meses)"

        if eh_ano:
            meses_ativos_cen = [m for m in MESES if res_base.get(m)]
            centros_set_cen = set()
            for _m in meses_ativos_cen:
                centros_set_cen.update(res_base[_m]["centros"].centro.tolist())
            centros_list = sorted(centros_set_cen)
            ocup_ref = {}
            for cen in centros_list:
                vals_A, vals_B, vals_C, ats_A, ats_B, ats_C = [], [], [], [], [], []
                for _m in meses_ativos_cen:
                    rc_ = res_base[_m]["centros"]
                    row_ = rc_[rc_.centro == cen]
                    if not row_.empty:
                        r_ = row_.iloc[0]
                        vals_A.append(r_.ocup_A); vals_B.append(r_.ocup_B); vals_C.append(r_.ocup_C)
                        ats_A.append(int(r_.ativo_A)); ats_B.append(int(r_.ativo_B)); ats_C.append(int(r_.ativo_C))
                ocup_ref[cen] = {
                    "oA": np.mean(vals_A) if vals_A else 0,
                    "oB": np.mean(vals_B) if vals_B else 0,
                    "oC": np.mean(vals_C) if vals_C else 0,
                    "aA": round(np.mean(ats_A)) if ats_A else 0,
                    "aB": round(np.mean(ats_B)) if ats_B else 0,
                    "aC": round(np.mean(ats_C)) if ats_C else 0,
                }
            st.markdown(f'''
<div class="aviso-warn">
📅 <b>Modo ANO</b> — override aplicado em <b>todos os {len(meses_ativos_cen)} meses com dados</b>.
Ocupação exibida = <b>média anual</b> por centro. Defina como cada turno deve se comportar.
</div>''', unsafe_allow_html=True)
        else:
            if not (mes_novo and res_base.get(mes_novo)):
                st.info("Selecione um mês com dados."); st.stop()
            centros_list = sorted(res_base[mes_novo]["centros"].centro.tolist())
            ocup_ref = {}
            for cen in centros_list:
                rc_ = res_base[mes_novo]["centros"]
                row_ = rc_[rc_.centro == cen]
                if not row_.empty:
                    r_ = row_.iloc[0]
                    ocup_ref[cen] = {
                        "oA": r_.ocup_A, "oB": r_.ocup_B, "oC": r_.ocup_C,
                        "aA": int(r_.ativo_A), "aB": int(r_.ativo_B), "aC": int(r_.ativo_C)
                    }

        if novo_nome and centros_list:
            cols_h = st.columns([3, 1, 1, 1])
            cols_h[0].markdown(
                "**Centro — ocup. média anual**" if eh_ano else "**Centro — ocup. A/B/C**"
            )
            cols_h[1].markdown("**A**"); cols_h[2].markdown("**B**"); cols_h[3].markdown("**C**")
            novo_ov = {}
            for cen in centros_list:
                ref = ocup_ref.get(cen, {"oA":0,"oB":0,"oC":0,"aA":0,"aB":0,"aC":0})
                oA, oB, oC = ref["oA"], ref["oB"], ref["oC"]
                aA_d, aB_d, aC_d = ref["aA"], ref["aB"], ref["aC"]
                eA = "🔴" if oA > 1 else ("🟡" if oA >= 0.85 else "🟢")
                eB = "🔴" if oB > 1 else ("🟡" if oB >= 0.85 else "🟢")
                eC = "🔴" if oC > 1 else ("🟡" if oC >= 0.85 else "🟢")
                c0, c1, c2, c3 = st.columns([3, 1, 1, 1])
                c0.markdown(f"`{cen}` {eA}{oA:.0%}/{eB}{oB:.0%}/{eC}{oC:.0%}")
                vA = c1.number_input("A", 0, 5, aA_d, key=f"n_{novo_nome}_{cen}_A", label_visibility="collapsed", help=f"Base:{aA_d}")
                vB = c2.number_input("B", 0, 5, aB_d, key=f"n_{novo_nome}_{cen}_B", label_visibility="collapsed", help=f"Base:{aB_d}")
                vC = c3.number_input("C", 0, 5, aC_d, key=f"n_{novo_nome}_{cen}_C", label_visibility="collapsed", help=f"Base:{aC_d}")
                novo_ov[cen] = {"A": vA, "B": vB, "C": vC}

            if st.button("💾 Salvar cenário", type="primary", key="btn_salvar_cen"):
                if eh_ano:
                    meses_ativos_cen = [m for m in MESES if res_base.get(m)]
                    ov_c = {m: novo_ov for m in meses_ativos_cen}
                else:
                    ov_c = {mes_novo: novo_ov}
                res_cen = calcular(pmp, tempo, dist, aplic, dias, horas_turno, thresholds, suporte_cfg,
                                   horas_efetivas=horas_efetivas, overrides=ov_c)
                st.session_state.cenarios[novo_nome] = {
                    "resultados": res_cen, "mes": mes_novo,
                    "overrides": ov_c, "eh_ano": eh_ano
                }
                st.success(f"✅ '{novo_nome}' salvo!"); st.rerun()

    if st.session_state.cenarios:
        todos = {"📌 Base": res_base}
        todos.update({k: v["resultados"] for k, v in st.session_state.cenarios.items()})
        st.plotly_chart(grafico_cenarios(todos), use_container_width=True)

        st.markdown('<div class="jd-sub">📊 Comparação detalhada — Base vs Cenários</div>', unsafe_allow_html=True)
        opcoes_cmp = [m for m in MESES if res_base.get(m)] + ["📅 ANO (todos os meses)"]
        mes_cmp = st.selectbox("Mês para comparar", opcoes_cmp, key="mes_cmp_r")
        eh_ano_cmp = mes_cmp == "📅 ANO (todos os meses)"
        meses_cmp_lista = [m for m in MESES if res_base.get(m)] if eh_ano_cmp else (
            [mes_cmp] if res_base.get(mes_cmp) else []
        )

        def _agregar(res_dict, meses_lista):
            rr = [res_dict.get(m) for m in meses_lista if res_dict.get(m)]
            if not rr: return None
            n = len(rr)
            sh_ciclo = sum(r["h_ciclo"] for r in rr)
            sh_labor = sum(r["h_labor"] for r in rr)
            sh_ativos= sum(r["h_ativos"] for r in rr)
            sh_todos = sum(r["h_todos"]  for r in rr)
            return {
                "tot_A":  round(sum(r["tot_A"] for r in rr) / n, 1),
                "tot_B":  round(sum(r["tot_B"] for r in rr) / n, 1),
                "tot_C":  round(sum(r["tot_C"] for r in rr) / n, 1),
                "total":  round(sum(r["total"]  for r in rr) / n, 1),
                "prod_labor_tot": sh_labor / sh_todos  if sh_todos  > 0 else 0,
                "prod_ciclo_tot": sh_ciclo / sh_todos  if sh_todos  > 0 else 0,
                "prod_labor_op":  sh_labor / sh_ativos if sh_ativos > 0 else 0,
            }

        if meses_cmp_lista:
            r_base_agg = _agregar(res_base, meses_cmp_lista)
            sufixo = " (méd.)" if eh_ano_cmp else ""
            cmp_rows = []
            for nm, res in todos.items():
                r_agg = _agregar(res, meses_cmp_lista)
                if not r_agg or not r_base_agg: continue
                is_base = "Base" in nm
                dA   = round(r_agg["tot_A"]  - r_base_agg["tot_A"],  1) if not is_base else "—"
                dB   = round(r_agg["tot_B"]  - r_base_agg["tot_B"],  1) if not is_base else "—"
                dC   = round(r_agg["tot_C"]  - r_base_agg["tot_C"],  1) if not is_base else "—"
                dT   = round(r_agg["total"]  - r_base_agg["total"],  1) if not is_base else "—"
                dL   = f'{r_agg["prod_labor_tot"] - r_base_agg["prod_labor_tot"]:+.1%}' if not is_base else "—"
                badge = f'<span class="cenario-badge">ANO</span>' if eh_ano_cmp and not is_base else ""
                cmp_rows.append({
                    "Cenário":          nm,
                    f"Turno A{sufixo}": r_agg["tot_A"],
                    f"Turno B{sufixo}": r_agg["tot_B"],
                    f"Turno C{sufixo}": r_agg["tot_C"],
                    f"Total{sufixo}":   r_agg["total"],
                    "Labor Tot.":       f'{r_agg["prod_labor_tot"]:.1%}',
                    "Ciclo Tot.":       f'{r_agg["prod_ciclo_tot"]:.1%}',
                    "ΔA":  f"{dA:+.1f}" if isinstance(dA, float) else dA,
                    "ΔB":  f"{dB:+.1f}" if isinstance(dB, float) else dB,
                    "ΔC":  f"{dC:+.1f}" if isinstance(dC, float) else dC,
                    "Δ Total": f"{dT:+.1f}" if isinstance(dT, float) else dT,
                    "Δ Labor": dL,
                })

            df_cmp = pd.DataFrame(cmp_rows)
            def _style_cmp(row):
                is_base = "Base" in str(row["Cenário"])
                if is_base:
                    return [f"background-color:{JD_VERDE_ESC};color:#FFFFFF;font-weight:700"] * len(row)
                styles = [""] * len(row)
                try:
                    d = float(str(row["Δ Total"]).replace("+", ""))
                    c_d = "#003D10" if d < 0 else ("#3D0000" if d > 0 else "")
                    t_d = "#B9F6CA" if d < 0 else ("#FF8A80" if d > 0 else "")
                    for i, col in enumerate(df_cmp.columns):
                        if col in ("ΔA","ΔB","ΔC","Δ Total","Δ Labor"):
                            styles[i] = f"background-color:{c_d};color:{t_d};font-weight:600"
                except: pass
                return styles
            st.dataframe(df_cmp.style.apply(_style_cmp, axis=1), use_container_width=True, hide_index=True)

            # Detalhamento por centro
            for nome_cen, dados_cen in st.session_state.cenarios.items():
                r_cen_res = dados_cen["resultados"]
                with st.expander(f"🔍 Detalhamento por centro — {nome_cen} vs Base"):
                    _m_ref2 = meses_cmp_lista[0] if meses_cmp_lista else None
                    det_rows = []
                    if _m_ref2 and res_base.get(_m_ref2) and r_cen_res.get(_m_ref2):
                        centros_set2 = sorted(set(
                            res_base[_m_ref2]["centros"].centro.tolist() +
                            r_cen_res[_m_ref2]["centros"].centro.tolist()
                        ))
                        for cen in centros_set2:
                            rb_c2 = res_base[_m_ref2]["centros"]
                            rc_c2 = r_cen_res[_m_ref2]["centros"]
                            rb_r2 = rb_c2[rb_c2.centro == cen]
                            rc_r2 = rc_c2[rc_c2.centro == cen]
                            if rb_r2.empty or rc_r2.empty: continue
                            rb2 = rb_r2.iloc[0]; rc2 = rc_r2.iloc[0]
                            mA = int(rb2.ativo_A) != int(rc2.ativo_A)
                            mB = int(rb2.ativo_B) != int(rc2.ativo_B)
                            mC = int(rb2.ativo_C) != int(rc2.ativo_C)
                            det_rows.append({
                                "Centro": cen,
                                "Ocup.A": f"{rb2.ocup_A:.0%}", "Base A": int(rb2.ativo_A), "Cen A": int(rc2.ativo_A),
                                "Ocup.B": f"{rb2.ocup_B:.0%}", "Base B": int(rb2.ativo_B), "Cen B": int(rc2.ativo_B),
                                "Ocup.C": f"{rb2.ocup_C:.0%}", "Base C": int(rb2.ativo_C), "Cen C": int(rc2.ativo_C),
                                "Mudou": "✅ Igual" if not (mA or mB or mC) else
                                         f"{'A ' if mA else ''}{'B ' if mB else ''}{'C' if mC else ''}alterado(s)",
                            })
                    if det_rows:
                        df_det2 = pd.DataFrame(det_rows)
                        def _sty_det2(row):
                            if "alterado" in str(row["Mudou"]):
                                return [f"background-color:#3D2D00;color:#FFE57F"] * len(row)
                            return [""] * len(row)
                        st.dataframe(df_det2.style.apply(_sty_det2, axis=1), use_container_width=True, hide_index=True)
                    # Se é cenário anual, mostra aviso
                    if dados_cen.get("eh_ano"):
                        st.markdown(f'<div class="aviso-ok">📅 Cenário anual — override aplicado em <b>todos os meses</b>. O detalhamento acima é do mês <b>{_m_ref2}</b>.</div>', unsafe_allow_html=True)

        # Exportação e gerenciamento
        st.markdown("---")
        st.markdown('<div class="jd-sub">📥 Exportar e gerenciar cenários</div>', unsafe_allow_html=True)
        col_exp, col_del = st.columns([3, 1])
        with col_exp:
            for nm, v in st.session_state.cenarios.items():
                if v.get("eh_ano"):
                    _m_exp = next((m for m in MESES if res_base.get(m)), None)
                    label_exp = f"📥 Exportar — {nm}  (ANO — todos os meses)"
                    fname_exp = f"cenario_{nm.replace(' ','_')}_ANO.xlsx"
                    _m_usado = _m_exp
                else:
                    _m_exp = v.get("mes", MESES[0])
                    label_exp = f"📥 Exportar — {nm}  ({_m_exp})"
                    fname_exp = f"cenario_{nm.replace(' ','_')}_{_m_exp}.xlsx"
                    _m_usado = _m_exp
                if _m_usado and res_base.get(_m_usado):
                    st.download_button(
                        label_exp,
                        data=exportar_cenario_vs_base(res_base, v["resultados"], _m_usado, nm),
                        file_name=fname_exp,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_cen_{nm}"
                    )
            # Exportar tabela comparativa
            if meses_cmp_lista and r_base_agg:
                def _exportar_cmp_full():
                    buf_c = BytesIO(); wb_c = openpyxl.Workbook()
                    brd_c = Border(
                        left=Side(style='thin',color='CCCCCC'), right=Side(style='thin',color='CCCCCC'),
                        top=Side(style='thin',color='CCCCCC'),  bottom=Side(style='thin',color='CCCCCC'))
                    JD_V_c = JD_VERDE_ESC.replace("#",""); JD_Y_c = JD_AMARELO.replace("#","")
                    def _ec2(c, bg="FFFFFF", fg="000000", bold=False, center=True):
                        c.font  = Font(name="Arial", bold=bold, color=fg, size=9)
                        c.fill  = PatternFill("solid", fgColor=bg)
                        c.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
                        c.border = brd_c
                    ws_c = wb_c.active; ws_c.title = "Comparação"
                    titulo_c = f"COMPARAÇÃO {'ANO' if eh_ano_cmp else mes_cmp.upper()}"
                    ws_c.merge_cells("A1:L1")
                    ct = ws_c.cell(1, 1, titulo_c)
                    ct.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
                    ct.fill = PatternFill("solid", fgColor=JD_V_c)
                    ct.alignment = Alignment(horizontal="center", vertical="center")
                    for i, h in enumerate(["Cenário","Turno A","Turno B","Turno C","Total",
                                            "Labor Tot.","Ciclo Tot.","ΔA","ΔB","ΔC","Δ Total","Δ Labor"], 1):
                        _ec2(ws_c.cell(2, i, h), JD_V_c, "FFFFFF", True)
                    todos_exp = {"📌 Base": res_base}
                    todos_exp.update({k: v["resultados"] for k, v in st.session_state.cenarios.items()})
                    for ri_e, (nm_e, res_e) in enumerate(todos_exp.items(), 3):
                        r_e = _agregar(res_e, meses_cmp_lista)
                        if not r_e: continue
                        is_base_e = "Base" in nm_e
                        bg_e = JD_V_c if is_base_e else ("EAF3FB" if ri_e % 2 == 0 else "FFFFFF")
                        fg_e = "FFFFFF" if is_base_e else "000000"
                        dT_e = round(r_e["total"] - r_base_agg["total"], 1) if not is_base_e else "—"
                        dA_e = round(r_e["tot_A"] - r_base_agg["tot_A"], 1) if not is_base_e else "—"
                        dB_e = round(r_e["tot_B"] - r_base_agg["tot_B"], 1) if not is_base_e else "—"
                        dC_e = round(r_e["tot_C"] - r_base_agg["tot_C"], 1) if not is_base_e else "—"
                        dL_e = f'{r_e["prod_labor_tot"] - r_base_agg["prod_labor_tot"]:+.1%}' if not is_base_e else "—"
                        vals_e = [nm_e, r_e["tot_A"], r_e["tot_B"], r_e["tot_C"], r_e["total"],
                                  f'{r_e["prod_labor_tot"]:.1%}', f'{r_e["prod_ciclo_tot"]:.1%}',
                                  f"{dA_e:+.1f}" if isinstance(dA_e, float) else dA_e,
                                  f"{dB_e:+.1f}" if isinstance(dB_e, float) else dB_e,
                                  f"{dC_e:+.1f}" if isinstance(dC_e, float) else dC_e,
                                  f"{dT_e:+.1f}" if isinstance(dT_e, float) else dT_e, dL_e]
                        for ci_e, v_e in enumerate(vals_e, 1):
                            _ec2(ws_c.cell(ri_e, ci_e, v_e), bg_e, fg_e, is_base_e, ci_e > 1)
                        ws_c.row_dimensions[ri_e].height = 14
                    for ci, w in enumerate([22, 8, 8, 8, 8, 10, 10, 6, 6, 6, 8, 10], 1):
                        ws_c.column_dimensions[get_column_letter(ci)].width = w
                    wb_c.save(buf_c); buf_c.seek(0)
                    return buf_c
                st.download_button(
                    f"📊 Exportar comparação ({'ANO' if eh_ano_cmp else mes_cmp})",
                    data=_exportar_cmp_full(),
                    file_name=f"comparacao_{'ano' if eh_ano_cmp else mes_cmp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_cmp_cen"
                )
        with col_del:
            if st.session_state.cenarios:
                dn = st.selectbox("Remover cenário", list(st.session_state.cenarios.keys()), key="del_c")
                if st.button("🗑️ Remover", type="secondary", key="btn_del_cen"):
                    del st.session_state.cenarios[dn]; st.rerun()
    else:
        st.info("Nenhum cenário criado ainda. Use o formulário acima para criar o primeiro.")
"""

if __name__ == "__main__":
    print("=" * 60)
    print("INSTRUÇÕES DE APLICAÇÃO")
    print("=" * 60)
    print()
    print("PASSO 1 — CSS EXTRA")
    print("  Adicione o conteúdo de CSS_EXTRA dentro do st.markdown(css)")
    print("  existente, logo antes do '</style>'")
    print()
    print("PASSO 2 — RESULTADOS ANO")
    print("  Dentro da TAB 4 (tab_res), localize:")
    print("    elif mes_r and mes_r == '📅 ANO (resumo anual)':")
    print("  Substitua TODO esse bloco (até o próximo elif/else no mesmo nível)")
    print("  pelo conteúdo de RESULTADOS_ANO")
    print()
    print("PASSO 3 — TAB CENÁRIOS")
    print("  Localize: 'with tab_cen:'")
    print("  Substitua TODO o bloco pelo conteúdo de CENARIOS_TAB")
    print()
    print("Pronto! O app terá:")
    print("  ✅ Cenários para ANO completo")
    print("  ✅ Visual rico com KPI cards, gauges SVG e barras por mês")
    print("  ✅ Agregação anual de cenários na tabela comparativa")
