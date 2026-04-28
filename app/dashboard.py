"""
dashboard.py — Crítica da Base Cadastral EFPC
Plotly Dash application. Dark theme, 4 pages.
Run: python app/dashboard.py
Deploy: gunicorn app.dashboard:server
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../src"))

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from dash import Dash, html, dcc, Input, Output, State, callback, no_update
import dash

from generate_data import generate_all
from validator import (validate_ativos, validate_assistidos,
                        validate_diferidos, population_summary)

# ── Bootstrap data on startup ─────────────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), "../data/raw")
BASE_PATH = os.path.join(DATA_DIR, "base_cadastral_2024.xlsx")

if not os.path.exists(BASE_PATH):
    os.makedirs(DATA_DIR, exist_ok=True)
    generate_all(output_dir=DATA_DIR)

xl = pd.ExcelFile(BASE_PATH)
DF_A = xl.parse("ATIVOS")
DF_S = xl.parse("ASSISTIDOS")
DF_D = xl.parse("DIFERIDOS")

ISS_A = validate_ativos(DF_A)
ISS_S = validate_assistidos(DF_S)
ISS_D = validate_diferidos(DF_D)
ALL_ISSUES = pd.concat([ISS_A, ISS_S, ISS_D], ignore_index=True)
SUMM = population_summary(DF_A, DF_S, DF_D)

REF = pd.Timestamp("2024-12-31")

# ── Helpers ───────────────────────────────────────────────────────────────────
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="DM Sans, sans-serif", color="#8b949e", size=11),
    margin=dict(l=12, r=12, t=28, b=12),
    xaxis=dict(showgrid=False, zeroline=False,
               tickfont=dict(color="#8b949e", size=10),
               linecolor="#21262d"),
    yaxis=dict(showgrid=True, gridcolor="#21262d", zeroline=False,
               tickfont=dict(color="#8b949e", size=10)),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
)

def ages(df):
    return pd.to_datetime(df["DT_NASCIMENTO"], dayfirst=True, errors="coerce").apply(
        lambda x: (REF - x).days / 365.25 if pd.notna(x) else np.nan)

def fmt_brl(v):
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

# ── App ───────────────────────────────────────────────────────────────────────
app = Dash(
    __name__,
    assets_folder="assets",
    suppress_callback_exceptions=True,
    title="Crítica Cadastral EFPC",
)
server = app.server   # for gunicorn

# ── Sidebar ───────────────────────────────────────────────────────────────────
PAGES = [
    ("visao-geral",    "📊", "Visão Geral"),
    ("inconsistencias","🔍", "Inconsistências"),
    ("ativos",         "👥", "Ativos"),
    ("assistidos",     "🏦", "Assistidos"),
]

sidebar = html.Div([
    html.Div([
        html.Div("CRÍTICA CADASTRAL", className="logo-title"),
        html.Div("EFPC · Data-base 31/12/2024", className="logo-sub"),
    ], className="sidebar-logo"),

    html.Div([
        html.Div("Análise", className="nav-section-label"),
        *[
            html.Button([
                html.Span(icon, className="nav-icon"),
                label,
            ], id=f"nav-{page_id}", className="nav-link",
               **{"data-page": page_id})
            for page_id, icon, label in PAGES
        ],
    ], className="nav-section"),

    html.Div([
        html.Div(f"Ativos: {SUMM['n_ativos']:,}", style={"marginBottom":"4px"}),
        html.Div(f"Assistidos: {SUMM['n_assistidos']:,}"),
        html.Div(f"Diferidos: {SUMM['n_diferidos']:,}"),
        html.Div(f"Resolução PREVIC 7/2022",
                 style={"marginTop":"10px","color":"#388bfd"}),
    ], className="sidebar-footer"),
], className="sidebar")

# ── App layout ────────────────────────────────────────────────────────────────
app.layout = html.Div([
    dcc.Store(id="current-page", data="visao-geral"),
    html.Div([
        sidebar,
        html.Div(id="page-content", className="main-content"),
    ], className="app-shell"),
], style={"background":"#0d1117","minHeight":"100vh"})


# ── Navigation callback ───────────────────────────────────────────────────────
@app.callback(
    Output("current-page", "data"),
    [Input(f"nav-{p}", "n_clicks") for p, _, __ in PAGES],
    prevent_initial_call=True,
)
def set_page(*args):
    ctx = dash.callback_context
    if not ctx.triggered:
        return no_update
    btn_id = ctx.triggered[0]["prop_id"].split(".")[0]
    return btn_id.replace("nav-", "")


@app.callback(
    [Output(f"nav-{p}", "className") for p, _, __ in PAGES],
    Input("current-page", "data"),
)
def update_nav(page):
    return ["nav-link active" if p == page else "nav-link"
            for p, _, __ in PAGES]


# ── Page renderer ─────────────────────────────────────────────────────────────
@app.callback(
    Output("page-content", "children"),
    Input("current-page", "data"),
)
def render_page(page):
    if page == "visao-geral":   return page_visao_geral()
    if page == "inconsistencias": return page_inconsistencias()
    if page == "ativos":        return page_ativos()
    if page == "assistidos":    return page_assistidos()
    return html.Div("Página não encontrada")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1: Visão Geral
# ══════════════════════════════════════════════════════════════════════════════
def page_visao_geral():
    n_crit  = (ALL_ISSUES["SEVERIDADE"] == "CRITICO").sum()
    n_alert = (ALL_ISSUES["SEVERIDADE"] == "ALERTA").sum()

    # Issues donut
    freq = ALL_ISSUES.groupby(["CODIGO","SEVERIDADE"]).size().reset_index(name="n")
    freq_sorted = freq.sort_values("n", ascending=False).head(10)

    colors_bar = ["#f85149" if s == "CRITICO" else "#d29922"
                  for s in freq_sorted["SEVERIDADE"]]

    fig_bar = go.Figure(go.Bar(
        x=freq_sorted["CODIGO"],
        y=freq_sorted["n"],
        marker_color=colors_bar,
        text=freq_sorted["n"],
        textposition="outside",
        textfont=dict(size=10, color="#e6edf3"),
    ))
    fig_bar.update_layout(**PLOTLY_LAYOUT,
                           yaxis_title="Ocorrências",
                           height=260)

    # Population donut
    fig_donut = go.Figure(go.Pie(
        labels=["Ativos","Assistidos","Diferidos"],
        values=[SUMM["n_ativos"], SUMM["n_assistidos"], SUMM["n_diferidos"]],
        hole=0.65,
        marker=dict(colors=["#388bfd","#3fb950","#d29922"]),
        textinfo="label+percent",
        textfont=dict(size=11, color="#e6edf3"),
    ))
    fig_donut.update_layout(**PLOTLY_LAYOUT, height=260,
                             showlegend=False)
    fig_donut.add_annotation(
        text=f"<b>{SUMM['n_total']:,}</b><br><span style='font-size:10px'>participantes</span>",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=14, color="#e6edf3"),
        align="center",
    )

    notice_cls = "notice-crit" if n_crit > 0 else "notice-info"
    notice_txt = (f"⛔  {n_crit} inconsistências CRÍTICAS identificadas — "
                  f"registros devem ser corrigidos antes da avaliação atuarial "
                  f"(Ref: CPA 017/2019 IBA, Art. 4.2)."
                  if n_crit > 0 else
                  "✅ Nenhuma inconsistência crítica.")

    return html.Div([
        html.Div([
            html.Div("Visão Geral", className="page-title"),
            html.Div(f"Data-base: 31/12/2024 · Resolução PREVIC 7/2022 + CPA 017/2019 IBA",
                     className="page-sub"),
        ], className="page-header"),

        html.Div(notice_txt, className=f"notice {notice_cls}"),

        html.Div([
            kpi("Total", f"{SUMM['n_total']:,}"),
            kpi("Ativos", f"{SUMM['n_ativos']:,}", "blue"),
            kpi("Assistidos", f"{SUMM['n_assistidos']:,}", "ok"),
            kpi("Diferidos", f"{SUMM['n_diferidos']:,}"),
            kpi("CRÍTICOS", f"{n_crit:,}", "crit",
                delta=f"{n_crit/len(ALL_ISSUES)*100:.1f}% do total"),
            kpi("ALERTAS", f"{n_alert:,}", "alert",
                delta=f"{n_alert/len(ALL_ISSUES)*100:.1f}% do total"),
            kpi("Razão Ass./At.", f"{SUMM['razao_assistidos_ativos']:.3f}",
                delta="fundo maduro > 0.40"),
            kpi("Massa Salarial", fmt_brl(SUMM["ativos_total_massa"]),
                delta="por mês"),
        ], className="kpi-grid"),

        html.Div([
            card("Inconsistências por Código",
                 "Top 10 — vermelho = CRÍTICO, laranja = ALERTA",
                 dcc.Graph(figure=fig_bar, config={"displayModeBar": False})),
            card("Composição da População",
                 "Por situação regulamentar (Res. PREVIC 23/2023)",
                 dcc.Graph(figure=fig_donut, config={"displayModeBar": False})),
        ], className="chart-grid"),
    ])


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2: Inconsistências
# ══════════════════════════════════════════════════════════════════════════════
def page_inconsistencias():
    freq = (ALL_ISSUES.groupby(["CODIGO","SEVERIDADE","CAMPO"])
            .size().reset_index(name="n").sort_values("n", ascending=False))

    colors = ["#f85149" if s == "CRITICO" else "#d29922"
              for s in freq["SEVERIDADE"]]

    fig_h = go.Figure(go.Bar(
        y=freq["CODIGO"],
        x=freq["n"],
        orientation="h",
        marker_color=colors,
        text=freq["n"],
        textposition="outside",
        textfont=dict(size=10, color="#e6edf3"),
    ))
    fig_h.update_layout(
        **PLOTLY_LAYOUT,
        height=max(240, len(freq) * 28),
        xaxis_title="Ocorrências",
        yaxis=dict(autorange="reversed", showgrid=False,
                   tickfont=dict(color="#e6edf3", size=11),
                   linecolor="#21262d"),
    )

    # Table rows
    rows = []
    for _, r in ALL_ISSUES.iterrows():
        sev   = r["SEVERIDADE"]
        cls   = "crit" if sev == "CRITICO" else "alert"
        rows.append(html.Tr([
            html.Td(r["ID_PARTICIPANTE"],
                    style={"fontFamily":"DM Mono,monospace","fontSize":"11px"}),
            html.Td(r["GRUPO"]),
            html.Td(r["CAMPO"]),
            html.Td(r["VALOR_ATUAL"],
                    style={"fontFamily":"DM Mono,monospace","fontSize":"11px"}),
            html.Td(badge(sev, cls)),
            html.Td(r["CODIGO"],
                    style={"fontFamily":"DM Mono,monospace","fontSize":"11px"}),
            html.Td(r["DESCRICAO"], className="wrap",
                    style={"fontSize":"11px","maxWidth":"360px"}),
        ]))

    return html.Div([
        html.Div([
            html.Div("Inconsistências Identificadas", className="page-title"),
            html.Div("Crítica da base cadastral conforme CPA 017/2019 IBA",
                     className="page-sub"),
        ], className="page-header"),

        html.Div([
            kpi("Total", f"{len(ALL_ISSUES):,}"),
            kpi("Críticas",  f"{(ALL_ISSUES['SEVERIDADE']=='CRITICO').sum():,}",
                "crit", delta="exigem correção"),
            kpi("Alertas",   f"{(ALL_ISSUES['SEVERIDADE']=='ALERTA').sum():,}",
                "alert", delta="exigem análise"),
        ], className="kpi-grid"),

        html.Div([
            card("Frequência por Código",
                 "Cada código corresponde a um tipo de inconsistência regulatória",
                 dcc.Graph(figure=fig_h, config={"displayModeBar": False})),
        ], className="chart-grid single"),

        card("Detalhamento Completo",
             f"Todos os {len(ALL_ISSUES):,} registros com problema",
             html.Div([
                 html.Table([
                     html.Thead(html.Tr([
                         html.Th("ID"), html.Th("Grupo"), html.Th("Campo"),
                         html.Th("Valor"), html.Th("Severidade"),
                         html.Th("Código"), html.Th("Descrição"),
                     ])),
                     html.Tbody(rows),
                 ]),
             ], className="table-wrap",
                style={"maxHeight":"480px","overflowY":"auto"})),
    ])


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3: Ativos
# ══════════════════════════════════════════════════════════════════════════════
def page_ativos():
    df = DF_A.copy()
    df["IDADE"] = ages(df).round(1)
    df_valid = df[df["SALARIO_CONTRIB"].notna() & (df["SALARIO_CONTRIB"] > 0)]

    # Age distribution
    bins   = list(range(20, 80, 5))
    labels = [f"{b}–{b+4}" for b in bins[:-1]]
    df["FAIXA"] = pd.cut(df["IDADE"], bins=bins, labels=labels, right=False)
    faixa = df.groupby("FAIXA", observed=True).size().reset_index(name="n")

    fig_age = go.Figure(go.Bar(
        x=faixa["FAIXA"].astype(str),
        y=faixa["n"],
        marker_color="#388bfd",
        marker_opacity=0.85,
    ))
    fig_age.update_layout(**PLOTLY_LAYOUT, height=230,
                           yaxis_title="Participantes", xaxis_title="Faixa etária")

    # Salary distribution
    sal = df_valid["SALARIO_CONTRIB"]
    sal_clipped = sal.clip(upper=sal.quantile(0.97))
    fig_sal = go.Figure(go.Histogram(
        x=sal_clipped,
        nbinsx=30,
        marker_color="#3fb950",
        marker_opacity=0.85,
    ))
    fig_sal.update_layout(**PLOTLY_LAYOUT, height=230,
                           yaxis_title="Participantes", xaxis_title="Salário (R$)")

    # Scatter age × salary
    df_scatter = df_valid[df_valid["IDADE"].notna()].copy()
    fig_sc = go.Figure(go.Scatter(
        x=df_scatter["IDADE"],
        y=df_scatter["SALARIO_CONTRIB"],
        mode="markers",
        marker=dict(color="#388bfd", opacity=0.45, size=5),
    ))
    fig_sc.update_layout(**PLOTLY_LAYOUT, height=260,
                          xaxis_title="Idade", yaxis_title="Salário (R$)")

    # Cargo breakdown
    cargo = df.groupby("CARGO").size().reset_index(name="n").sort_values("n", ascending=True)
    fig_cargo = go.Figure(go.Bar(
        y=cargo["CARGO"], x=cargo["n"],
        orientation="h",
        marker_color="#388bfd", marker_opacity=0.8,
    ))
    fig_cargo.update_layout(
        **PLOTLY_LAYOUT, height=260,
        xaxis_title="Participantes",
        yaxis=dict(autorange="reversed", showgrid=False,
                   tickfont=dict(color="#e6edf3",size=10),
                   linecolor="#21262d"),
    )

    media_sal = df_valid["SALARIO_CONTRIB"].mean()
    media_idade = df["IDADE"].mean()
    media_tempo = df.assign(
        tempo=pd.to_datetime(df["DT_ADMISSAO_PLANO"], dayfirst=True,
                              errors="coerce").apply(
            lambda x: (REF - x).days / 365.25 if pd.notna(x) else np.nan)
    )["tempo"].mean()

    return html.Div([
        html.Div([
            html.Div("Participantes Ativos", className="page-title"),
            html.Div("Análise demográfica e salarial — 600 participantes",
                     className="page-sub"),
        ], className="page-header"),

        html.Div([
            kpi("Total", f"{len(df):,}", "blue"),
            kpi("Média de Idade", f"{media_idade:.1f} anos"),
            kpi("Tempo Médio no Plano", f"{media_tempo:.1f} anos"),
            kpi("Salário Médio", fmt_brl(media_sal)),
            kpi("Massa Salarial", fmt_brl(df_valid["SALARIO_CONTRIB"].sum()),
                delta="por mês"),
        ], className="kpi-grid"),

        html.Div([
            card("Distribuição Etária",
                 "Número de ativos por faixa de idade",
                 dcc.Graph(figure=fig_age, config={"displayModeBar": False})),
            card("Distribuição Salarial",
                 "Salário de contribuição (P97 como limite superior)",
                 dcc.Graph(figure=fig_sal, config={"displayModeBar": False})),
        ], className="chart-grid"),

        html.Div([
            card("Idade × Salário",
                 "Relação entre idade e salário de contribuição",
                 dcc.Graph(figure=fig_sc, config={"displayModeBar": False})),
            card("Distribuição por Cargo",
                 "Número de participantes por cargo",
                 dcc.Graph(figure=fig_cargo, config={"displayModeBar": False})),
        ], className="chart-grid"),
    ])


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4: Assistidos
# ══════════════════════════════════════════════════════════════════════════════
def page_assistidos():
    df = DF_S.copy()
    df["IDADE"] = ages(df).round(1)
    df_valid = df[df["BENEFICIO_MENSAL"].notna() & (df["BENEFICIO_MENSAL"] > 0)]

    # Age histogram
    bins   = list(range(55, 100, 5))
    labels = [f"{b}–{b+4}" for b in bins[:-1]]
    df["FAIXA"] = pd.cut(df["IDADE"], bins=bins, labels=labels, right=False)
    faixa = df.groupby("FAIXA", observed=True).size().reset_index(name="n")

    fig_age = go.Figure(go.Bar(
        x=faixa["FAIXA"].astype(str), y=faixa["n"],
        marker_color="#3fb950", marker_opacity=0.85,
    ))
    fig_age.update_layout(**PLOTLY_LAYOUT, height=230,
                           yaxis_title="Assistidos", xaxis_title="Faixa etária")

    # Benefit distribution
    fig_ben = go.Figure(go.Histogram(
        x=df_valid["BENEFICIO_MENSAL"].clip(upper=df_valid["BENEFICIO_MENSAL"].quantile(0.97)),
        nbinsx=25, marker_color="#d29922", marker_opacity=0.85,
    ))
    fig_ben.update_layout(**PLOTLY_LAYOUT, height=230,
                           yaxis_title="Assistidos", xaxis_title="Benefício Mensal (R$)")

    # By type
    tipo = df.groupby("TIPO_BENEFICIO").size().reset_index(name="n")
    fig_tipo = go.Figure(go.Pie(
        labels=tipo["TIPO_BENEFICIO"], values=tipo["n"],
        hole=0.55,
        marker=dict(colors=["#3fb950","#388bfd","#d29922","#f85149"]),
        textinfo="label+percent",
        textfont=dict(size=11, color="#e6edf3"),
    ))
    fig_tipo.update_layout(**PLOTLY_LAYOUT, height=260, showlegend=False)

    # Scatter age × benefit
    df_sc = df_valid[df_valid["IDADE"].notna()]
    fig_sc = go.Figure(go.Scatter(
        x=df_sc["IDADE"], y=df_sc["BENEFICIO_MENSAL"],
        mode="markers",
        marker=dict(color="#3fb950", opacity=0.5, size=5),
    ))
    fig_sc.update_layout(**PLOTLY_LAYOUT, height=260,
                          xaxis_title="Idade", yaxis_title="Benefício Mensal (R$)")

    return html.Div([
        html.Div([
            html.Div("Assistidos (Beneficiários)", className="page-title"),
            html.Div("Análise demográfica e de benefícios — 250 assistidos",
                     className="page-sub"),
        ], className="page-header"),

        html.Div([
            kpi("Total", f"{len(df):,}", "ok"),
            kpi("Média de Idade", f"{df['IDADE'].mean():.1f} anos"),
            kpi("Benefício Médio", fmt_brl(df_valid["BENEFICIO_MENSAL"].mean()),
                delta="por mês"),
            kpi("Total Benefícios", fmt_brl(df_valid["BENEFICIO_MENSAL"].sum()),
                delta="por mês"),
            kpi("Razão Ass./At.", f"{SUMM['razao_assistidos_ativos']:.3f}",
                "alert" if SUMM["razao_assistidos_ativos"] > 0.40 else "ok",
                delta="fundo maduro > 0.40"),
        ], className="kpi-grid"),

        html.Div([
            card("Distribuição Etária",
                 "Número de assistidos por faixa de idade",
                 dcc.Graph(figure=fig_age, config={"displayModeBar": False})),
            card("Distribuição de Benefícios",
                 "Benefício mensal (P97 como limite superior)",
                 dcc.Graph(figure=fig_ben, config={"displayModeBar": False})),
        ], className="chart-grid"),

        html.Div([
            card("Por Tipo de Benefício",
                 "Composição por tipo (aposentadoria, pensão, invalidez)",
                 dcc.Graph(figure=fig_tipo, config={"displayModeBar": False})),
            card("Idade × Benefício",
                 "Relação entre idade e valor do benefício mensal",
                 dcc.Graph(figure=fig_sc, config={"displayModeBar": False})),
        ], className="chart-grid"),
    ])


# ── Run ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=8050)
