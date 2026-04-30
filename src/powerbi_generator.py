"""
powerbi_generator.py  —  Crítica Cadastral EFPC
Gera um projeto PBIP/PBIR completo a partir de powerbi_data.xlsx.

Estrutura baseada no código-fonte real do powerbpy (v0.2.0):
  https://github.com/Russell-Shean/powerbpy

Schemas reais usados pelo powerbpy:
  visualContainer : 1.3.0
  page            : 1.2.0
  report          : 1.2.0
  pagesMetadata   : 1.0.0
  versionMetadata : 1.0.0
"""

import argparse, json, os, shutil, uuid
import pandas as pd

# Schemas extraídos do código-fonte e templates do powerbpy v0.2.0
S_VISUAL  = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.3.0/schema.json"
S_PAGE    = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.2.0/schema.json"
S_REPORT  = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.2.0/schema.json"
S_PAGES   = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json"
S_VERSION = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json"
S_DEFPROP = "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json"
S_PBISM   = "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json"

PAGE_W, PAGE_H = 1280, 720


def uid():
    return uuid.uuid4().hex[:20]

def write(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ── Visual base (from powerbpy/visual.py) ─────────────────────────────────────
def base_visual(vid, x, y, w, h, title=None):
    """Exact structure from powerbpy _Visual.__init__"""
    v = {
        "$schema": S_VISUAL,
        "name": vid,
        "position": {
            "x": x, "y": y, "z": 6000,
            "height": h, "width": w,
            "tabOrder": -1001
        },
        "visual": {
            "visualType": "GENERIC",
            "objects": {},
            "visualContainerObjects": {
                "general": [{"properties": {
                    "altText": {"expr": {"Literal": {"Value": "''"}}}
                }}],
                "title": [],
                "background": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "false"}}}
                }}]
            },
            "drillFilterOtherVisuals": True
        }
    }
    if title:
        v["visual"]["visualContainerObjects"]["title"].append({
            "properties": {
                "show": {"expr": {"Literal": {"Value": "true"}}},
                "text": {"expr": {"Literal": {"Value": f"'{title}'"}}}
            }
        })
    return v


# ── Card for 1-row KPI table ──────────────────────────────────────────────────
def card_visual(vid, x, y, w, h, title, table, column):
    """Card using Sum(column) — works when SUMARIO_KPI columns are numeric types."""
    v = base_visual(vid, x, y, w, h, title)
    v["visual"]["visualType"] = "card"
    v["visual"]["query"] = {
        "queryState": {
            "Values": {
                "projections": [{
                    "field": {"Aggregation": {
                        "Expression": {"Column": {
                            "Expression": {"SourceRef": {"Entity": table}},
                            "Property": column
                        }},
                        "Function": 0   # Sum — on a 1-row table returns the value
                    }},
                    "queryRef": f"Sum({table}.{column})",
                    "nativeQueryRef": column
                }]
            }
        },
        "sortDefinition": {"isDefaultSort": True}
    }
    v["visual"]["objects"]["categoryLabels"] = [{"properties": {
        "show": {"expr": {"Literal": {"Value": "false"}}}
    }}]
    return v


# ── Scatter chart ──────────────────────────────────────────────────────────────
def scatter_visual(vid, x, y, w, h, title, table, x_col, y_col):
    v = base_visual(vid, x, y, w, h, title)
    v["visual"]["visualType"] = "scatterChart"
    v["visual"]["query"] = {
        "queryState": {
            "Details": {"projections": [{
                "field": {"Column": {
                    "Expression": {"SourceRef": {"Entity": table}},
                    "Property": "ID_PARTICIPANTE"
                }},
                "queryRef": f"{table}.ID_PARTICIPANTE",
                "nativeQueryRef": "ID_PARTICIPANTE",
                "active": True
            }]},
            "X": {"projections": [{
                "field": {"Aggregation": {
                    "Expression": {"Column": {
                        "Expression": {"SourceRef": {"Entity": table}},
                        "Property": x_col
                    }},
                    "Function": 0
                }},
                "queryRef": f"Sum({table}.{x_col})",
                "nativeQueryRef": f"Sum of {x_col}"
            }]},
            "Y": {"projections": [{
                "field": {"Aggregation": {
                    "Expression": {"Column": {
                        "Expression": {"SourceRef": {"Entity": table}},
                        "Property": y_col
                    }},
                    "Function": 0
                }},
                "queryRef": f"Sum({table}.{y_col})",
                "nativeQueryRef": f"Sum of {y_col}"
            }]}
        },
        "sortDefinition": {"isDefaultSort": True}
    }
    return v


# ── Donut chart ────────────────────────────────────────────────────────────────
def donut_visual(vid, x, y, w, h, title, table, cat_col):
    v = base_visual(vid, x, y, w, h, title)
    v["visual"]["visualType"] = "donutChart"
    v["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [{
                "field": {"Column": {
                    "Expression": {"SourceRef": {"Entity": table}},
                    "Property": cat_col
                }},
                "queryRef": f"{table}.{cat_col}",
                "nativeQueryRef": cat_col,
                "active": True
            }]},
            "Y": {"projections": [{
                "field": {"Aggregation": {
                    "Expression": {"Column": {
                        "Expression": {"SourceRef": {"Entity": table}},
                        "Property": "ID_PARTICIPANTE"
                    }},
                    "Function": 5  # CountNonNull
                }},
                "queryRef": f"CountNonNull({table}.ID_PARTICIPANTE)",
                "nativeQueryRef": "Count of ID_PARTICIPANTE"
            }]}
        },
        "sortDefinition": {"isDefaultSort": True}
    }
    return v


# ── Page builder ───────────────────────────────────────────────────────────────
def build_page(pages_dir, page_id, display_name, visuals):
    pg_dir = os.path.join(pages_dir, page_id)
    for v in visuals:
        path = os.path.join(pg_dir, "visuals", v["name"], "visual.json")
        write(path, v)

    write(os.path.join(pg_dir, "page.json"), {
        "$schema": S_PAGE,
        "name": page_id,
        "displayName": display_name,
        "displayOption": "FitToPage",
        "height": PAGE_H,
        "width": PAGE_W
    })
    return page_id


# ── All pages ──────────────────────────────────────────────────────────────────
def build_all_pages(def_dir):
    pd_dir = os.path.join(def_dir, "pages")
    ids = []

    # ── Visão Geral ───────────────────────────────────────────────────────────
    p1 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total",      "n_total"),
        ("Ativos",     "n_ativos"),
        ("Assistidos", "n_assistidos"),
        ("Diferidos",  "n_diferidos"),
        ("Críticos",   "n_criticos"),
        ("Alertas",    "n_alertas"),
    ]):
        vis.append(card_visual(uid(), 10+i*200, 10, 190, 85,
                               title, "SUMARIO_KPI", col))
    # Horizontal bar: inconsistências por código (Sum of OCORRENCIAS=int64)
    vis.append(bar_visual(uid(), 10, 115, 630, 575,
                           "Inconsistências por Código",
                           "FREQ_INCONSISTENCIAS", "CODIGO", "OCORRENCIAS",
                           horizontal=True, agg_func=0))
    # Donut: composição da população — uses POPULACAO_TOTAL (all 3 groups combined)
    vis.append(donut_visual(uid(), 650, 115, 620, 280,
                             "Composição da População",
                             "POPULACAO_TOTAL", "SITUACAO_GRUPO"))
    ids.append(build_page(pd_dir, p1, "Visão Geral", vis))

    # ── Inconsistências ───────────────────────────────────────────────────────
    p2 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total",    "n_total"),
        ("Críticos", "n_criticos"),
        ("Alertas",  "n_alertas"),
    ]):
        vis.append(card_visual(uid(), 10+i*230, 10, 220, 85,
                               title, "SUMARIO_KPI", col))
    vis.append(bar_visual(uid(), 10, 115, 600, 575,
                           "Frequência por Código (CRITICO / ALERTA)",
                           "FREQ_INCONSISTENCIAS", "CODIGO", "OCORRENCIAS",
                           horizontal=True, agg_func=0))
    vis.append(table_visual(uid(), 620, 115, 650, 575,
                             "Detalhamento Completo", "INCONSISTENCIAS",
                             ["ID_PARTICIPANTE","GRUPO","CAMPO",
                              "SEVERIDADE","CODIGO","DESCRICAO"]))
    ids.append(build_page(pd_dir, p2, "Inconsistências", vis))

    # ── Ativos ────────────────────────────────────────────────────────────────
    p3 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total Ativos",   "n_ativos"),
        ("Salário Médio",  "ativos_media_salario"),
        ("Massa Salarial", "ativos_total_massa"),
    ]):
        vis.append(card_visual(uid(), 10+i*270, 10, 260, 85,
                               title, "SUMARIO_KPI", col))
    # Distribuição etária (bar)
    vis.append(bar_visual(uid(), 10, 115, 615, 270,
                           "Distribuição Etária",
                           "POP_ATIVOS", "FAIXA_ETARIA", "ID_PARTICIPANTE",
                           agg_func=5))
    # Distribuição salarial (bar — emulates histogram by faixa)
    vis.append(bar_visual(uid(), 640, 115, 630, 270,
                           "Distribuição Salarial por Faixa",
                           "POP_ATIVOS", "FAIXA_ETARIA", "SALARIO_CONTRIB",
                           agg_func=1))  # Avg salary per age band
    # Scatter: Idade x Salário
    vis.append(scatter_visual(uid(), 10, 400, 615, 290,
                               "Idade × Salário",
                               "POP_ATIVOS", "IDADE", "SALARIO_CONTRIB"))
    # Por Cargo (horizontal bar)
    vis.append(bar_visual(uid(), 640, 400, 630, 290,
                           "Por Cargo",
                           "POP_ATIVOS", "CARGO", "ID_PARTICIPANTE",
                           horizontal=True, agg_func=5))
    ids.append(build_page(pd_dir, p3, "Ativos", vis))

    # ── Assistidos ────────────────────────────────────────────────────────────
    p4 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total Assistidos",  "n_assistidos"),
        ("Benefício Médio",   "assistidos_media_beneficio"),
        ("Total Benefícios",  "assistidos_total_beneficio"),
        ("Razão Ass./At.",    "razao_assistidos_ativos"),
    ]):
        vis.append(card_visual(uid(), 10+i*300, 10, 290, 85,
                               title, "SUMARIO_KPI", col))
    # Distribuição etária
    vis.append(bar_visual(uid(), 10, 115, 615, 270,
                           "Distribuição Etária",
                           "POP_ASSISTIDOS", "FAIXA_ETARIA", "ID_PARTICIPANTE",
                           agg_func=5))
    # Distribuição de benefícios por faixa (avg)
    vis.append(bar_visual(uid(), 640, 115, 630, 270,
                           "Benefício Médio por Faixa Etária",
                           "POP_ASSISTIDOS", "FAIXA_ETARIA", "BENEFICIO_MENSAL",
                           agg_func=1))  # Avg benefit per age band
    # Donut: por tipo de benefício
    vis.append(donut_visual(uid(), 10, 400, 615, 290,
                             "Por Tipo de Benefício",
                             "POP_ASSISTIDOS", "TIPO_BENEFICIO"))
    # Scatter: Idade x Benefício
    vis.append(scatter_visual(uid(), 640, 400, 630, 290,
                               "Idade × Benefício Mensal",
                               "POP_ASSISTIDOS", "IDADE", "BENEFICIO_MENSAL"))
    ids.append(build_page(pd_dir, p4, "Assistidos", vis))
def bar_visual(vid, x, y, w, h, title, table,
               cat_col, val_col, horizontal=False, agg_func=5):
    """agg_func=5 is Count — safe for all column types including string."""
    vtype = "barChart" if horizontal else "clusteredColumnChart"
    # Function codes per powerbpy source:
    # 0=Sum, 1=Avg, 2=Min, 3=Max, 4=Last, 5=CountNonNull (displayed as "Count" in Desktop)
    func_name = {0:"Sum",1:"Avg",2:"Min",3:"Max",4:"Last",5:"CountNonNull"}.get(agg_func,"CountNonNull")
    v = base_visual(vid, x, y, w, h, title)
    v["visual"]["visualType"] = vtype
    v["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [{
                "field": {"Column": {
                    "Expression": {"SourceRef": {"Entity": table}},
                    "Property": cat_col
                }},
                "queryRef": f"{table}.{cat_col}",
                "nativeQueryRef": cat_col,
                "active": True
            }]},
            "Y": {"projections": [{
                "field": {"Aggregation": {
                    "Expression": {"Column": {
                        "Expression": {"SourceRef": {"Entity": table}},
                        "Property": val_col
                    }},
                    "Function": agg_func
                }},
                "queryRef": f"{func_name}({table}.{val_col})",
                "nativeQueryRef": f"{func_name} of {val_col}"
            }]}
        },
        "sortDefinition": {
            "sort": [{"field": {"Aggregation": {
                "Expression": {"Column": {
                    "Expression": {"SourceRef": {"Entity": table}},
                    "Property": val_col
                }},
                "Function": agg_func
            }}, "direction": "Descending"}],
            "isDefaultSort": True
        }
    }
    return v


# ── Table (from powerbpy/table.py pattern) ────────────────────────────────────
def table_visual(vid, x, y, w, h, title, table_name, columns):
    v = base_visual(vid, x, y, w, h, title)
    v["visual"]["visualType"] = "tableEx"
    v["visual"]["query"] = {
        "queryState": {
            "Values": {"projections": [{
                "field": {"Column": {
                    "Expression": {"SourceRef": {"Entity": table_name}},
                    "Property": col
                }},
                "queryRef": f"{table_name}.{col}",
                "nativeQueryRef": col
            } for col in columns]}
        },
        "sortDefinition": {"isDefaultSort": True}
    }
    return v


# ── Page builder ───────────────────────────────────────────────────────────────
def build_page(pages_dir, page_id, display_name, visuals):
    pg_dir = os.path.join(pages_dir, page_id)
    for v in visuals:
        path = os.path.join(pg_dir, "visuals", v["name"], "visual.json")
        write(path, v)

    # page.json — exact structure from powerbpy template
    write(os.path.join(pg_dir, "page.json"), {
        "$schema": S_PAGE,
        "name": page_id,
        "displayName": display_name,
        "displayOption": "FitToPage",
        "height": PAGE_H,
        "width": PAGE_W
    })
    return page_id


# ── All pages ──────────────────────────────────────────────────────────────────
def build_all_pages(def_dir):
    pd_dir = os.path.join(def_dir, "pages")
    ids = []

    # Visão Geral
    p1 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total",      "n_total"),
        ("Ativos",     "n_ativos"),
        ("Assistidos", "n_assistidos"),
        ("Diferidos",  "n_diferidos"),
        ("Críticos",   "n_criticos"),
        ("Alertas",    "n_alertas"),
    ]):
        vis.append(card_visual(uid(), 10+i*200, 10, 190, 85,
                               title, "SUMARIO_KPI", col))
    vis.append(bar_visual(uid(), 10, 115, 620, 570,
                           "Inconsistências por Código",
                           "FREQ_INCONSISTENCIAS", "CODIGO", "OCORRENCIAS",
                           horizontal=True, agg_func=0))  # OCORRENCIAS is int64
    vis.append(bar_visual(uid(), 650, 115, 620, 280,
                           "Composição da População",
                           "POPULACAO_TOTAL", "SITUACAO_GRUPO", "ID_PARTICIPANTE",
                           horizontal=False, agg_func=5))
    ids.append(build_page(pd_dir, p1, "Visão Geral", vis))

    # Inconsistências
    p2 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total",    "n_total"),
        ("Críticos", "n_criticos"),
        ("Alertas",  "n_alertas"),
    ]):
        vis.append(card_visual(uid(), 10+i*230, 10, 220, 85,
                               title, "SUMARIO_KPI", col))
    vis.append(bar_visual(uid(), 10, 115, 580, 575,
                           "Frequência por Código",
                           "FREQ_INCONSISTENCIAS", "CODIGO", "OCORRENCIAS",
                           horizontal=True, agg_func=0))  # OCORRENCIAS numeric
    vis.append(table_visual(uid(), 600, 115, 670, 575,
                             "Detalhamento", "INCONSISTENCIAS",
                             ["ID_PARTICIPANTE","GRUPO","CAMPO",
                              "SEVERIDADE","CODIGO","DESCRICAO"]))
    ids.append(build_page(pd_dir, p2, "Inconsistências", vis))

    # Ativos
    p3 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total Ativos",   "n_ativos"),
        ("Salário Médio",  "ativos_media_salario"),
        ("Massa Salarial", "ativos_total_massa"),
    ]):
        vis.append(card_visual(uid(), 10+i*270, 10, 260, 85,
                               title, "SUMARIO_KPI", col))
    vis.append(bar_visual(uid(), 10, 115, 620, 270,
                           "Distribuição Etária — Ativos",
                           "POP_ATIVOS", "FAIXA_ETARIA", "ID_PARTICIPANTE",
                           agg_func=5))
    vis.append(bar_visual(uid(), 650, 115, 620, 270,
                           "Por Cargo",
                           "POP_ATIVOS", "CARGO", "ID_PARTICIPANTE",
                           horizontal=True, agg_func=5))
    vis.append(bar_visual(uid(), 10, 400, 1260, 290,
                           "Salário por Faixa Etária",
                           "POP_ATIVOS", "FAIXA_ETARIA", "ID_PARTICIPANTE",
                           agg_func=5))
    ids.append(build_page(pd_dir, p3, "Ativos", vis))

    # Assistidos
    p4 = uid(); vis = []
    for i, (title, col) in enumerate([
        ("Total Assistidos",  "n_assistidos"),
        ("Benefício Médio",   "assistidos_media_beneficio"),
        ("Total Benefícios",  "assistidos_total_beneficio"),
        ("Razão Ass./At.",    "razao_assistidos_ativos"),
    ]):
        vis.append(card_visual(uid(), 10+i*205, 10, 195, 85,
                               title, "SUMARIO_KPI", col))
    vis.append(bar_visual(uid(), 10, 115, 620, 270,
                           "Distribuição Etária — Assistidos",
                           "POP_ASSISTIDOS", "FAIXA_ETARIA", "ID_PARTICIPANTE",
                           agg_func=5))
    vis.append(bar_visual(uid(), 650, 115, 620, 270,
                           "Por Tipo de Benefício",
                           "POP_ASSISTIDOS", "TIPO_BENEFICIO", "ID_PARTICIPANTE",
                           agg_func=5))
    vis.append(bar_visual(uid(), 10, 400, 1260, 290,
                           "Distribuição de Benefícios por Faixa",
                           "POP_ASSISTIDOS", "FAIXA_ETARIA", "ID_PARTICIPANTE",
                           agg_func=5))
    ids.append(build_page(pd_dir, p4, "Assistidos", vis))

    return ids


# ── TMDL ───────────────────────────────────────────────────────────────────────
def write_tmdl(base, xlsx_abs):
    sm_dir = os.path.join(base, "CriticaCadastral.SemanticModel")
    md = os.path.join(sm_dir, "definition")
    td = os.path.join(md, "tables")
    os.makedirs(td, exist_ok=True)
    os.makedirs(os.path.join(sm_dir, ".pbi"), exist_ok=True)
    os.makedirs(os.path.join(md, "cultures"), exist_ok=True)

    # .platform — required by Desktop to identify this as a SemanticModel
    with open(os.path.join(sm_dir, ".platform"), "w", encoding="utf-8") as f:
        json.dump({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "SemanticModel", "displayName": "CriticaCadastral"},
            "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}
        }, f, indent=2)

    # diagramLayout.json
    with open(os.path.join(sm_dir, "diagramLayout.json"), "w", encoding="utf-8") as f:
        json.dump({
            "version": "1.1.0",
            "diagrams": [{"ordinal": 0, "scrollPosition": {"x": 0, "y": 0},
                          "nodes": [], "name": "All tables", "zoomValue": 100,
                          "pinKeyFieldsToTop": False, "showExtraHeaderInfo": False,
                          "hideKeyFieldsWhenCollapsed": False, "tablesLocked": False}],
            "selectedDiagram": "All tables", "defaultDiagram": "All tables"
        }, f, indent=2)

    # .pbi/editorSettings.json
    with open(os.path.join(sm_dir, ".pbi", "editorSettings.json"), "w", encoding="utf-8") as f:
        json.dump({
            "version": "1.0", "parallelQueryLoading": True,
            "typeDetectionEnabled": True, "relationshipImportEnabled": True,
            "shouldNotifyUserOfNameConflictResolution": True
        }, f, indent=2)

    # cultures/pt-BR.tmdl
    with open(os.path.join(md, "cultures", "pt-BR.tmdl"), "w", encoding="utf-8") as f:
        f.write('cultureInfo pt-BR\n\n\tlinguisticMetadata =\n\t\t\t{\n\t\t\t  "Version": "1.0.0",\n\t\t\t  "Language": "pt-BR"\n\t\t\t}\n\t\tcontentType: json\n')

    # model.tmdl — matches powerbpy template structure
    with open(os.path.join(md, "model.tmdl"), "w", encoding="utf-8") as f:
        f.write(
            "model Model\n"
            "\tculture: pt-BR\n"
            "\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
            "\tsourceQueryCulture: pt-BR\n"
            "\tdataAccessOptions\n"
            "\t\tlegacyRedirects\n"
            "\t\treturnErrorValuesAsNull\n"
            "\n"
            "annotation PBI_ProTooling = [\"DevMode\"]\n"
            "\n"
            "annotation __PBI_TimeIntelligenceEnabled = 0\n"
            "\n"
            "annotation PBI_QueryOrder = []\n"
            "\n"
            "ref cultureInfo pt-BR\n"
        )

    # database.tmdl
    with open(os.path.join(md, "database.tmdl"), "w", encoding="utf-8") as f:
        f.write("database\n\tcompatibilityLevel: 1567\n")

    tables = {
        "POP_ATIVOS":          ["ID_PARTICIPANTE","SEXO","SITUACAO","CARGO",
                                 "SALARIO_CONTRIB","IDADE","FAIXA_ETARIA",
                                 "TEMPO_PLANO_ANOS","SITUACAO_GRUPO"],
        "POP_ASSISTIDOS":      ["ID_PARTICIPANTE","SEXO","TIPO_BENEFICIO",
                                 "BENEFICIO_MENSAL","IDADE","FAIXA_ETARIA","SITUACAO_GRUPO"],
        "POP_DIFERIDOS":       ["ID_PARTICIPANTE","SALDO_CONTA","SITUACAO_GRUPO"],
        "INCONSISTENCIAS":     ["ID_PARTICIPANTE","GRUPO","CAMPO",
                                 "VALOR_ATUAL","SEVERIDADE","CODIGO","DESCRICAO"],
        "FREQ_INCONSISTENCIAS":["CODIGO","SEVERIDADE","CAMPO","OCORRENCIAS","PCT_TOTAL"],
        "SUMARIO_KPI":         ["data_base","n_ativos","n_assistidos","n_diferidos",
                                 "n_total","n_criticos","n_alertas",
                                 "ativos_media_idade","ativos_media_salario","ativos_total_massa",
                                 "assistidos_media_idade","assistidos_media_beneficio",
                                 "assistidos_total_beneficio","razao_assistidos_ativos"],
        "POPULACAO_TOTAL":     ["ID_PARTICIPANTE","SITUACAO_GRUPO","FAIXA_ETARIA","IDADE"],
    }
    # Column types: string by default, numeric where needed for Sum/aggregation
    col_types = {
        # FREQ_INCONSISTENCIAS
        "OCORRENCIAS": "int64",
        "PCT_TOTAL":   "double",
        # POP_ATIVOS / POP_ASSISTIDOS / POP_DIFERIDOS
        "SALARIO_CONTRIB":   "double",
        "BENEFICIO_MENSAL":  "double",
        "SALDO_CONTA":       "double",
        "IDADE":             "double",
        "TEMPO_PLANO_ANOS":  "double",
        # SUMARIO_KPI — all numeric so cards can Sum the single row
        "n_ativos":                  "int64",
        "n_assistidos":              "int64",
        "n_diferidos":               "int64",
        "n_total":                   "int64",
        "n_criticos":                "int64",
        "n_alertas":                 "int64",
        "ativos_media_idade":        "double",
        "ativos_media_salario":      "double",
        "ativos_total_massa":        "double",
        "assistidos_media_idade":    "double",
        "assistidos_media_beneficio":"double",
        "assistidos_total_beneficio":"double",
        "razao_assistidos_ativos":   "double",
    }

    # Special M expressions for tables that need explicit type casting
    m_expressions = {}
    numeric_kpi_cols = [
        "n_ativos","n_assistidos","n_diferidos","n_total","n_criticos","n_alertas",
        "ativos_media_idade","ativos_media_salario","ativos_total_massa",
        "assistidos_media_idade","assistidos_media_beneficio",
        "assistidos_total_beneficio","razao_assistidos_ativos"
    ]
    # Build type list as plain string (no f-string to avoid brace escaping issues)
    kpi_type_pairs = ", ".join(
        '{"' + c + '", type number}' for c in numeric_kpi_cols
    )
    kpi_m_lines = [
        '    src = Excel.Workbook(File.Contents("' + xlsx_abs + '"), null, true),',
        '    sheet = src{[Item="SUMARIO_KPI",Kind="Sheet"]}[Data],',
        '    promoted = Table.PromoteHeaders(sheet, [PromoteAllScalars=true]),',
        '    typed = Table.TransformColumnTypes(promoted, { ' + kpi_type_pairs + ' })',
        'in',
        '    typed',
    ]
    m_expressions["SUMARIO_KPI"] = kpi_m_lines

    freq_m_lines = [
        '    src = Excel.Workbook(File.Contents("' + xlsx_abs + '"), null, true),',
        '    sheet = src{[Item="FREQ_INCONSISTENCIAS",Kind="Sheet"]}[Data],',
        '    promoted = Table.PromoteHeaders(sheet, [PromoteAllScalars=true]),',
        '    typed = Table.TransformColumnTypes(promoted, { {"OCORRENCIAS", Int64.Type}, {"PCT_TOTAL", type number} })',
        'in',
        '    typed',
    ]
    m_expressions["FREQ_INCONSISTENCIAS"] = freq_m_lines

    for tbl, cols in tables.items():
        # dataType must be explicit — Empty type is not allowed on import columns
        col_lines = "\n".join(
            f"    column {c}\n        dataType: {col_types.get(c, 'string')}"
            for c in cols
        )
        # Use custom M expression if defined, otherwise default
        if tbl in m_expressions:
            m_lines = m_expressions[tbl]
        else:
            m_lines = [
                '    src = Excel.Workbook(File.Contents("' + xlsx_abs + '"), null, true),',
                '    sheet = src{[Item="' + tbl + '",Kind="Sheet"]}[Data],',
                '    promoted = Table.PromoteHeaders(sheet, [PromoteAllScalars=true])',
                'in',
                '    promoted',
            ]
        # Build indented M block (no backticks — same format as powerbpy)
        indented_lines = "\n".join(
            ("                " if not line.startswith("in") else "            ") + line
            for line in m_lines
        )
        tmdl = (
            f'table {tbl}\n'
            f'    annotation PBI_ResultType = Table\n\n'
            f'    partition {tbl}-partition = m\n'
            f'        mode: import\n'
            f'        source =\n'
            f'            let\n'
            f'{indented_lines}\n\n'
            f'{col_lines}\n'
        )
        with open(os.path.join(td, f"{tbl}.tmdl"), "w", encoding="utf-8") as f:
            f.write(tmdl)

    write(os.path.join(sm_dir, "definition.pbism"),
          {"$schema": S_PBISM, "version": "4.0"})


# ── Root report files ──────────────────────────────────────────────────────────
def write_report_root(base, page_ids):
    rd = os.path.join(base, "CriticaCadastral.Report")
    dd = os.path.join(rd, "definition")
    os.makedirs(dd, exist_ok=True)

    # .platform — required by Desktop to identify this as a Report item
    with open(os.path.join(rd, ".platform"), "w", encoding="utf-8") as f:
        json.dump({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "Report", "displayName": "CriticaCadastral"},
            "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}
        }, f, indent=2)

    # StaticResources — CY24SU10 base theme (embedded, no external file dependency)
    theme_dir = os.path.join(rd, "StaticResources", "SharedResources", "BaseThemes")
    os.makedirs(theme_dir, exist_ok=True)
    # Theme content from powerbpy v0.2.0 dashboard_resources
    _CY24SU10_THEME = {"name":"CY24SU10","dataColors":["#118DFF","#12239E","#E66C37","#6B007B","#E044A7","#744EC2","#D9B300","#D64550","#197278","#1AAB40","#15C6F4","#4092FF","#FFA058","#BE5DC9","#F472D0","#B5A1FF","#C4A200","#FF8080","#00DBBC","#5BD667","#0091D5","#4668C5","#FF6300","#99008A","#EC008C","#533285","#99700A","#FF4141","#1F9A85","#25891C","#0057A2","#002050","#C94F0F","#450F54","#B60064","#34124F","#6A5A29","#1AAB40","#BA141A","#0C3D37","#0B511F"],"foreground":"#252423","foregroundNeutralSecondary":"#605E5C","foregroundNeutralTertiary":"#B3B0AD","background":"#FFFFFF","backgroundLight":"#F3F2F1","backgroundNeutral":"#C8C6C4","tableAccent":"#118DFF","good":"#1AAB40","neutral":"#D9B300","bad":"#D64554","maximum":"#118DFF","center":"#D9B300","minimum":"#DEEFFF","null":"#FF7F48","hyperlink":"#0078d4","visitedHyperlink":"#0078d4","textClasses":{"callout":{"fontSize":45,"fontFace":"DIN","color":"#252423"},"title":{"fontSize":12,"fontFace":"DIN","color":"#252423"},"header":{"fontSize":12,"fontFace":"Segoe UI Semibold","color":"#252423"},"label":{"fontSize":10,"fontFace":"Segoe UI","color":"#252423"}},"visualStyles":{"*":{"*":{"*":[{"wordWrap":True}],"line":[{"transparency":0}],"outline":[{"transparency":0}],"plotArea":[{"transparency":0}],"categoryAxis":[{"showAxisTitle":True,"gridlineStyle":"dotted","concatenateLabels":False}],"valueAxis":[{"showAxisTitle":True,"gridlineStyle":"dotted"}],"y2Axis":[{"show":True}],"title":[{"titleWrap":True}],"lineStyles":[{"strokeWidth":3}],"wordWrap":[{"show":True}],"background":[{"show":True,"transparency":0}],"border":[{"width":1}]}},"slicer":{"*":{"general":[{"responsive":True}]}},"clusteredColumnChart":{"*":{"general":[{"responsive":True}]}},"barChart":{"*":{"general":[{"responsive":True}]}},"page":{"*":{"background":[{"transparency":100}]}}}}
    with open(os.path.join(theme_dir, "CY24SU10.json"), "w", encoding="utf-8") as f:
        json.dump(_CY24SU10_THEME, f, indent=2)

    # definition.pbir
    write(os.path.join(rd, "definition.pbir"), {
        "$schema": S_DEFPROP,
        "version": "4.0",
        "datasetReference": {
            "byPath": {"path": "../CriticaCadastral.SemanticModel"}
        }
    })

    # version.json — exact content from powerbpy template
    write(os.path.join(dd, "version.json"), {
        "$schema": S_VERSION,
        "version": "2.0.0"
    })

    # report.json — exact structure from powerbpy template
    write(os.path.join(dd, "report.json"), {
        "$schema": S_REPORT,
        "themeCollection": {
            "baseTheme": {
                "name": "CY24SU10",
                "reportVersionAtImport": "5.59",
                "type": "SharedResources"
            }
        },
        "layoutOptimization": "None",
        "objects": {
            "section": [{"properties": {
                "verticalAlignment": {
                    "expr": {"Literal": {"Value": "'Top'"}}
                }
            }}]
        },
        "resourcePackages": [
            {
                "name": "SharedResources",
                "type": "SharedResources",
                "items": [{"name": "CY24SU10",
                           "path": "BaseThemes/CY24SU10.json",
                           "type": "BaseTheme"}]
            },
            {"name": "RegisteredResources", "type": "RegisteredResources", "items": []}
        ],
        "settings": {
            "useStylableVisualContainerHeader": True,
            "defaultDrillFilterOtherVisuals": True,
            "allowChangeFilterTypes": True,
            "useDefaultAggregateDisplayName": True
        },
        "slowDataSourceSettings": {
            "isCrossHighlightingDisabled": False,
            "isSlicerSelectionsButtonEnabled": False,
            "isFilterSelectionsButtonEnabled": False,
            "isFieldWellButtonEnabled": False,
            "isApplyAllButtonEnabled": False
        }
    })

    # pages.json — exact structure from powerbpy template
    write(os.path.join(dd, "pages", "pages.json"), {
        "$schema": S_PAGES,
        "pageOrder": page_ids,
        "activePageName": page_ids[0]
    })

    # .pbip
    write(os.path.join(base, "CriticaCadastral.pbip"), {
        "version": "1.0",
        "artifacts": [{"report": {"path": "CriticaCadastral.Report"}}]
    })

    # .gitignore
    with open(os.path.join(base, ".gitignore"), "w") as f:
        f.write(".pbi/\nlocalSettings.json\nCriticaCadastral.SemanticModel/.pbi/\n")


# ── Main ───────────────────────────────────────────────────────────────────────
def generate(input_path, output_dir):
    print("=" * 60)
    print("PBIP/PBIR GENERATOR  —  Crítica Cadastral EFPC")
    print(f"Input : {input_path}")
    print(f"Output: {output_dir}")
    print("=" * 60)

    xl  = pd.ExcelFile(input_path)
    iss = xl.parse("INCONSISTENCIAS")
    print(f"  Inconsistências: {len(iss)} ({(iss.SEVERIDADE=='CRITICO').sum()} críticas, "
          f"{(iss.SEVERIDADE=='ALERTA').sum()} alertas)")

    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    xlsx_abs = os.path.abspath(input_path).replace("\\", "/")

    print("\n[1/3] TMDL semantic model...")
    write_tmdl(output_dir, xlsx_abs)

    print("[2/3] PBIR pages & visuals...")
    def_dir  = os.path.join(output_dir, "CriticaCadastral.Report", "definition")
    page_ids = build_all_pages(def_dir)

    print("[3/3] Root files...")
    write_report_root(output_dir, page_ids)

    n_vis = sum(1 for _ in
                (f for r, _, fs in os.walk(output_dir)
                 for f in fs if f == "visual.json"))

    print(f"\n{'='*60}")
    print(f"Pronto! {output_dir}/CriticaCadastral.pbip")
    print(f"  Páginas : {len(page_ids)}")
    print(f"  Visuais : {n_vis}")
    print(f"\nComo abrir:")
    print(f"  File > Open > {output_dir}\\CriticaCadastral.pbip")
    print("=" * 60)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--input",  default="results/reports/powerbi_data.xlsx")
    p.add_argument("--output", default="powerbi/report")
    args = p.parse_args()
    generate(args.input, args.output)

if __name__ == "__main__":
    main()