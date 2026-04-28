"""
dashboard.py — Crítica da Base Cadastral EFPC
Plotly Dash 4.x — dcc.Location routing
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../src"))

import pandas as pd
import numpy as np
import plotly.graph_objects as go
from dash import Dash, html, dcc, Input, Output

from generate_data import generate_all
from validator import (validate_ativos, validate_assistidos,
                        validate_diferidos, population_summary)

# ── Data ──────────────────────────────────────────────────────────────────────
DATA_DIR  = os.path.join(os.path.dirname(__file__), "../data/raw")
BASE_PATH = os.path.join(DATA_DIR, "base_cadastral_2024.xlsx")

if not os.path.exists(BASE_PATH):
    os.makedirs(DATA_DIR, exist_ok=True)
    generate_all(output_dir=DATA_DIR)

xl         = pd.ExcelFile(BASE_PATH)
DF_A       = xl.parse("ATIVOS")
DF_S       = xl.parse("ASSISTIDOS")
DF_D       = xl.parse("DIFERIDOS")
ISS_A      = validate_ativos(DF_A)
ISS_S      = validate_assistidos(DF_S)
ISS_D      = validate_diferidos(DF_D)
ALL_ISSUES = pd.concat([ISS_A, ISS_S, ISS_D], ignore_index=True)
SUMM       = population_summary(DF_A, DF_S, DF_D)
REF        = pd.Timestamp("2024-12-31")

# ── Base theme (no yaxis — set per chart to avoid conflicts) ──────────────────
BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="DM Sans, sans-serif", color="#8b949e", size=11),
    margin=dict(l=12, r=12, t=28, b=12),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
)

XAXIS = dict(showgrid=False, zeroline=False,
             tickfont=dict(color="#8b949e", size=10), linecolor="#21262d")
YAXIS = dict(showgrid=True, gridcolor="#21262d", zeroline=False,
             tickfont=dict(color="#8b949e", size=10))
YAXIS_H = dict(autorange="reversed", showgrid=False,   # for horizontal bars
               tickfont=dict(color="#e6edf3", size=10), linecolor="#21262d")

def layout(height=260, xtitle="", ytitle="", horizontal=False):
    return dict(**BASE, height=height,
                xaxis={**XAXIS, "title": xtitle},
                yaxis={**(YAXIS_H if horizontal else YAXIS), "title": ytitle})

# ── Helpers ───────────────────────────────────────────────────────────────────
def ages(df):
    return pd.to_datetime(df["DT_NASCIMENTO"], dayfirst=True, errors="coerce").apply(
        lambda x: (REF - x).days / 365.25 if pd.notna(x) else np.nan)

def brl(v):
    return f"R$ {v:,.0f}".replace(",","X").replace(".",",").replace("X",".")

def kpi(label, value, cls="", delta=""):
    return html.Div([
        html.Div(label, className="kpi-label"),
        html.Div(value, className="kpi-value"),
        html.Div(delta, className="kpi-delta") if delta else None,
    ], className=f"kpi-card {cls}")

def card(title, sub, *children):
    return html.Div([
        html.Div(title, className="card-title"),
        html.Div(sub,   className="card-sub"),
        *children,
    ], className="card")

def badge(text, cls):
    return html.Span(text, className=f"badge badge-{cls}")

def G(fig):
    return dcc.Graph(figure=fig, config={"displayModeBar": False})

# ── Routes ────────────────────────────────────────────────────────────────────
PAGES = [
    ("/",               "📊", "Visão Geral"),
    ("/inconsistencias","🔍", "Inconsistências"),
    ("/ativos",         "👥", "Ativos"),
    ("/assistidos",     "🏦", "Assistidos"),
]

# ── App ───────────────────────────────────────────────────────────────────────
app    = Dash(__name__, assets_folder="assets",
              suppress_callback_exceptions=True,
              title="Crítica Cadastral EFPC")
server = app.server

def sidebar(current):
    return html.Div([
        html.Div([
            html.Div("CRÍTICA CADASTRAL", className="logo-title"),
            html.Div("EFPC · Data-base 31/12/2024", className="logo-sub"),
        ], className="sidebar-logo"),
        html.Div([
            html.Div("Análise", className="nav-section-label"),
            *[dcc.Link(
                [html.Span(icon, className="nav-icon"), label],
                href=href,
                className="nav-link active" if current == href else "nav-link",
                style={"textDecoration":"none"},
            ) for href, icon, label in PAGES],
        ], className="nav-section"),
        html.Div([
            html.Div(f"Ativos: {SUMM['n_ativos']:,}",    style={"marginBottom":"4px"}),
            html.Div(f"Assistidos: {SUMM['n_assistidos']:,}"),
            html.Div(f"Diferidos: {SUMM['n_diferidos']:,}"),
            html.Div("Resolução PREVIC 7/2022",
                     style={"marginTop":"10px","color":"#388bfd"}),
        ], className="sidebar-footer"),
    ], className="sidebar")

app.layout = html.Div([
    dcc.Location(id="url", refresh=False),
    html.Div([
        html.Div(id="sidebar-slot"),
        html.Div(id="content", className="main-content"),
    ], className="app-shell"),
], style={"background":"#0d1117","minHeight":"100vh"})

@app.callback(
    Output("sidebar-slot", "children"),
    Output("content",      "children"),
    Input("url", "pathname"),
)
def route(pathname):
    path = pathname or "/"
    sb   = sidebar(path)
    if path in ("/", ""):    return sb, _overview()
    if path == "/inconsistencias": return sb, _issues()
    if path == "/ativos":    return sb, _ativos()
    if path == "/assistidos":return sb, _assistidos()
    return sb, html.Div("Página não encontrada")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — Visão Geral
# ══════════════════════════════════════════════════════════════════════════════
def _overview():
    n_crit  = (ALL_ISSUES["SEVERIDADE"] == "CRITICO").sum()
    n_alert = (ALL_ISSUES["SEVERIDADE"] == "ALERTA").sum()
    total   = len(ALL_ISSUES)

    freq = (ALL_ISSUES.groupby(["CODIGO","SEVERIDADE"]).size()
            .reset_index(name="n").sort_values("n", ascending=False))

    fig_bar = go.Figure(go.Bar(
        x=freq["CODIGO"], y=freq["n"],
        marker_color=["#f85149" if s=="CRITICO" else "#d29922"
                      for s in freq["SEVERIDADE"]],
        text=freq["n"], textposition="outside",
        textfont=dict(size=10, color="#e6edf3"),
    ))
    fig_bar.update_layout(**layout(260, ytitle="Ocorrências"))

    fig_donut = go.Figure(go.Pie(
        labels=["Ativos","Assistidos","Diferidos"],
        values=[SUMM["n_ativos"],SUMM["n_assistidos"],SUMM["n_diferidos"]],
        hole=0.65,
        marker=dict(colors=["#388bfd","#3fb950","#d29922"]),
        textinfo="label+percent",
        textfont=dict(size=11, color="#e6edf3"),
    ))
    fig_donut.update_layout(**BASE, height=260, showlegend=False)
    fig_donut.add_annotation(
        text=(f"<b>{SUMM['n_total']:,}</b><br>"
              f"<span style='font-size:10px'>participantes</span>"),
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=14, color="#e6edf3"), align="center",
    )

    notice_cls = "notice-crit" if n_crit > 0 else "notice-info"
    notice_txt = (
        f"⛔  {n_crit} inconsistências CRÍTICAS — devem ser corrigidas antes da "
        f"avaliação atuarial (CPA 017/2019 IBA, Art. 4.2)."
        if n_crit > 0 else "✅ Nenhuma inconsistência crítica."
    )

    return html.Div([
        html.Div([
            html.Div("Visão Geral", className="page-title"),
            html.Div("Data-base: 31/12/2024 · Res. PREVIC 7/2022 + CPA 017/2019 IBA",
                     className="page-sub"),
        ], className="page-header"),
        html.Div(notice_txt, className=f"notice {notice_cls}"),
        html.Div([
            kpi("Total",          f"{SUMM['n_total']:,}"),
            kpi("Ativos",         f"{SUMM['n_ativos']:,}",     "blue"),
            kpi("Assistidos",     f"{SUMM['n_assistidos']:,}", "ok"),
            kpi("Diferidos",      f"{SUMM['n_diferidos']:,}"),
            kpi("CRÍTICOS",       f"{n_crit:,}", "crit",
                delta=f"{n_crit/total*100:.1f}% do total" if total else ""),
            kpi("ALERTAS",        f"{n_alert:,}", "alert",
                delta=f"{n_alert/total*100:.1f}% do total" if total else ""),
            kpi("Razão Ass./At.", f"{SUMM['razao_assistidos_ativos']:.3f}",
                delta="fundo maduro > 0.40"),
            kpi("Massa Salarial", brl(SUMM["ativos_total_massa"]),
                delta="por mês"),
        ], className="kpi-grid"),
        html.Div([
            card("Inconsistências por Código",
                 "Vermelho = CRÍTICO · Laranja = ALERTA", G(fig_bar)),
            card("Composição da População",
                 "Por situação regulamentar (Res. PREVIC 23/2023)", G(fig_donut)),
        ], className="chart-grid"),
    ])


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — Inconsistências
# ══════════════════════════════════════════════════════════════════════════════
def _issues():
    freq = (ALL_ISSUES.groupby(["CODIGO","SEVERIDADE","CAMPO"]).size()
            .reset_index(name="n").sort_values("n", ascending=False))

    fig_h = go.Figure(go.Bar(
        y=freq["CODIGO"], x=freq["n"], orientation="h",
        marker_color=["#f85149" if s=="CRITICO" else "#d29922"
                      for s in freq["SEVERIDADE"]],
        text=freq["n"], textposition="outside",
        textfont=dict(size=10, color="#e6edf3"),
    ))
    fig_h.update_layout(
        **layout(max(240, len(freq)*30), xtitle="Ocorrências", horizontal=True))

    rows = [
        html.Tr([
            html.Td(r["ID_PARTICIPANTE"],
                    style={"fontFamily":"DM Mono,monospace","fontSize":"11px"}),
            html.Td(r["GRUPO"]),
            html.Td(r["CAMPO"]),
            html.Td(r["VALOR_ATUAL"],
                    style={"fontFamily":"DM Mono,monospace","fontSize":"11px"}),
            html.Td(badge(r["SEVERIDADE"],
                    "crit" if r["SEVERIDADE"]=="CRITICO" else "alert")),
            html.Td(r["CODIGO"],
                    style={"fontFamily":"DM Mono,monospace","fontSize":"11px"}),
            html.Td(r["DESCRICAO"], className="wrap",
                    style={"fontSize":"11px","maxWidth":"360px"}),
        ])
        for _, r in ALL_ISSUES.iterrows()
    ]

    n_crit  = (ALL_ISSUES["SEVERIDADE"]=="CRITICO").sum()
    n_alert = (ALL_ISSUES["SEVERIDADE"]=="ALERTA").sum()

    return html.Div([
        html.Div([
            html.Div("Inconsistências Identificadas", className="page-title"),
            html.Div("Crítica da base cadastral — CPA 017/2019 IBA",
                     className="page-sub"),
        ], className="page-header"),
        html.Div([
            kpi("Total",    f"{len(ALL_ISSUES):,}"),
            kpi("Críticas", f"{n_crit:,}",  "crit",  delta="exigem correção"),
            kpi("Alertas",  f"{n_alert:,}", "alert", delta="exigem análise"),
        ], className="kpi-grid"),
        html.Div([
            card("Frequência por Código",
                 "Cada código = um tipo de inconsistência regulatória",
                 G(fig_h)),
        ], className="chart-grid single"),
        card("Detalhamento Completo",
             f"Todos os {len(ALL_ISSUES):,} registros com problema",
             html.Div([
                 html.Table([
                     html.Thead(html.Tr([
                         html.Th(t) for t in
                         ["ID","Grupo","Campo","Valor",
                          "Severidade","Código","Descrição"]
                     ])),
                     html.Tbody(rows),
                 ]),
             ], className="table-wrap",
                style={"maxHeight":"480px","overflowY":"auto"})),
    ])


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — Ativos
# ══════════════════════════════════════════════════════════════════════════════
def _ativos():
    df      = DF_A.copy()
    df["AG"]= ages(df).round(1)
    df_v    = df[df["SALARIO_CONTRIB"].notna() & (df["SALARIO_CONTRIB"]>0)]

    bins   = list(range(20, 80, 5))
    labels = [f"{b}–{b+4}" for b in bins[:-1]]
    df["FAIXA"] = pd.cut(df["AG"], bins=bins, labels=labels, right=False)
    faixa = df.groupby("FAIXA", observed=True).size().reset_index(name="n")

    fig_age = go.Figure(go.Bar(
        x=faixa["FAIXA"].astype(str), y=faixa["n"],
        marker_color="#388bfd", marker_opacity=0.85,
    ))
    fig_age.update_layout(**layout(230, xtitle="Faixa etária",
                                    ytitle="Participantes"))

    fig_sal = go.Figure(go.Histogram(
        x=df_v["SALARIO_CONTRIB"].clip(
            upper=df_v["SALARIO_CONTRIB"].quantile(0.97)),
        nbinsx=30, marker_color="#3fb950", marker_opacity=0.85,
    ))
    fig_sal.update_layout(**layout(230, xtitle="Salário (R$)",
                                    ytitle="Participantes"))

    df_sc = df_v[df_v["AG"].notna()]
    fig_sc = go.Figure(go.Scatter(
        x=df_sc["AG"], y=df_sc["SALARIO_CONTRIB"],
        mode="markers", marker=dict(color="#388bfd", opacity=0.4, size=5),
    ))
    fig_sc.update_layout(**layout(260, xtitle="Idade", ytitle="Salário (R$)"))

    cargo = (df.groupby("CARGO").size()
               .reset_index(name="n").sort_values("n", ascending=True))
    fig_cargo = go.Figure(go.Bar(
        y=cargo["CARGO"], x=cargo["n"], orientation="h",
        marker_color="#388bfd", marker_opacity=0.8,
    ))
    fig_cargo.update_layout(
        **layout(260, xtitle="Participantes", horizontal=True))

    return html.Div([
        html.Div([
            html.Div("Participantes Ativos", className="page-title"),
            html.Div(f"Análise demográfica e salarial — {len(df):,} participantes",
                     className="page-sub"),
        ], className="page-header"),
        html.Div([
            kpi("Total",          f"{len(df):,}",                     "blue"),
            kpi("Média Idade",    f"{df['AG'].mean():.1f} anos"),
            kpi("Salário Médio",  brl(df_v["SALARIO_CONTRIB"].mean())),
            kpi("Massa Salarial", brl(df_v["SALARIO_CONTRIB"].sum()),
                delta="por mês"),
        ], className="kpi-grid"),
        html.Div([
            card("Distribuição Etária",
                 "Número de ativos por faixa de idade", G(fig_age)),
            card("Distribuição Salarial",
                 "Salário de contribuição (P97 como limite)", G(fig_sal)),
        ], className="chart-grid"),
        html.Div([
            card("Idade × Salário",
                 "Relação entre idade e salário", G(fig_sc)),
            card("Distribuição por Cargo",
                 "Número de participantes por cargo", G(fig_cargo)),
        ], className="chart-grid"),
    ])


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — Assistidos
# ══════════════════════════════════════════════════════════════════════════════
def _assistidos():
    df      = DF_S.copy()
    df["AG"]= ages(df).round(1)
    df_v    = df[df["BENEFICIO_MENSAL"].notna() & (df["BENEFICIO_MENSAL"]>0)]

    bins   = list(range(55, 100, 5))
    labels = [f"{b}–{b+4}" for b in bins[:-1]]
    df["FAIXA"] = pd.cut(df["AG"], bins=bins, labels=labels, right=False)
    faixa = df.groupby("FAIXA", observed=True).size().reset_index(name="n")

    fig_age = go.Figure(go.Bar(
        x=faixa["FAIXA"].astype(str), y=faixa["n"],
        marker_color="#3fb950", marker_opacity=0.85,
    ))
    fig_age.update_layout(**layout(230, xtitle="Faixa etária",
                                    ytitle="Assistidos"))

    fig_ben = go.Figure(go.Histogram(
        x=df_v["BENEFICIO_MENSAL"].clip(
            upper=df_v["BENEFICIO_MENSAL"].quantile(0.97)),
        nbinsx=25, marker_color="#d29922", marker_opacity=0.85,
    ))
    fig_ben.update_layout(**layout(230, xtitle="Benefício (R$)",
                                    ytitle="Assistidos"))

    tipo = df.groupby("TIPO_BENEFICIO").size().reset_index(name="n")
    fig_tipo = go.Figure(go.Pie(
        labels=tipo["TIPO_BENEFICIO"], values=tipo["n"], hole=0.55,
        marker=dict(colors=["#3fb950","#388bfd","#d29922","#f85149"]),
        textinfo="label+percent",
        textfont=dict(size=11, color="#e6edf3"),
    ))
    fig_tipo.update_layout(**BASE, height=260, showlegend=False)

    df_sc = df_v[df_v["AG"].notna()]
    fig_sc = go.Figure(go.Scatter(
        x=df_sc["AG"], y=df_sc["BENEFICIO_MENSAL"],
        mode="markers", marker=dict(color="#3fb950", opacity=0.5, size=5),
    ))
    fig_sc.update_layout(**layout(260, xtitle="Idade",
                                   ytitle="Benefício (R$)"))

    ratio_cls = "alert" if SUMM["razao_assistidos_ativos"] > 0.40 else "ok"

    return html.Div([
        html.Div([
            html.Div("Assistidos (Beneficiários)", className="page-title"),
            html.Div(f"Análise demográfica e de benefícios — {len(df):,} assistidos",
                     className="page-sub"),
        ], className="page-header"),
        html.Div([
            kpi("Total",            f"{len(df):,}",                    "ok"),
            kpi("Média Idade",      f"{df['AG'].mean():.1f} anos"),
            kpi("Benefício Médio",  brl(df_v["BENEFICIO_MENSAL"].mean()),
                delta="por mês"),
            kpi("Total Benefícios", brl(df_v["BENEFICIO_MENSAL"].sum()),
                delta="por mês"),
            kpi("Razão Ass./At.",   f"{SUMM['razao_assistidos_ativos']:.3f}",
                ratio_cls, delta="fundo maduro > 0.40"),
        ], className="kpi-grid"),
        html.Div([
            card("Distribuição Etária",
                 "Número de assistidos por faixa de idade", G(fig_age)),
            card("Distribuição de Benefícios",
                 "Benefício mensal (P97 como limite)", G(fig_ben)),
        ], className="chart-grid"),
        html.Div([
            card("Por Tipo de Benefício",
                 "Aposentadoria programada, pensão, invalidez", G(fig_tipo)),
            card("Idade × Benefício",
                 "Relação entre idade e valor do benefício", G(fig_sc)),
        ], className="chart-grid"),
    ])


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=8050)