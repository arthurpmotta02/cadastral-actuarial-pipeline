"""
validator.py
Crítica da base cadastral para avaliação atuarial de EFPC.

References:
  - Resolução PREVIC 7/2022, Art. 8
  - CPA 017/2019 IBA (Auditoria Atuarial)
  - Resolução PREVIC 23/2023
  - CPAO 035 IBA (Reservas Matemáticas)

Severity levels:
  CRITICO — record unusable in valuation without correction
  ALERTA  — actuary must review and formally justify
"""

import pandas as pd
import numpy as np
from datetime import date, datetime

REF_DATE       = date(2024, 12, 31)
SMN_2024       = 1412.00
IDADE_MIN      = 16
IDADE_MAX_ATIVO = 75
IDADE_MIN_APOS = 55
SITUACOES_OK   = {"ATIVO","ASSISTIDO","DIFERIDO","DESLIGADO_BPD"}
SEXOS_OK       = {"M","F"}

def _parse(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(val).strip(), fmt).date()
        except: pass
    try:
        return pd.to_datetime(val, dayfirst=True).date()
    except: return None

def _age(dt): return (REF_DATE - dt).days / 365.25

# ─────────────────────────────────────────────────────────────────────────────

def validate_ativos(df):
    issues = []

    def F(pid, campo, valor, sev, cod, desc):
        issues.append({"ID_PARTICIPANTE": pid, "GRUPO": "ATIVO",
                        "CAMPO": campo, "VALOR_ATUAL": str(valor),
                        "SEVERIDADE": sev, "CODIGO": cod, "DESCRICAO": desc})

    for _, r in df.iterrows():
        pid = r.get("ID_PARTICIPANTE","?")

        # ── C001: campos obrigatórios ────────────────────────────────────────
        for c in ["DT_NASCIMENTO","DT_ADMISSAO_PLANO","SALARIO_CONTRIB","SEXO","CPF","SITUACAO"]:
            if pd.isna(r.get(c)):
                F(pid, c, "NULO", "CRITICO", "C001",
                  f"Campo obrigatório '{c}' ausente. "
                  f"Impede cálculo do PMBaC. Ref: Art.8 Res.PREVIC 7/2022.")

        # ── C002/C003: data de nascimento ────────────────────────────────────
        dt_nasc = _parse(r.get("DT_NASCIMENTO"))
        if dt_nasc:
            if dt_nasc > REF_DATE:
                F(pid, "DT_NASCIMENTO", r["DT_NASCIMENTO"], "CRITICO", "C002",
                  "Data de nascimento posterior à data-base da avaliação.")
            else:
                idade = _age(dt_nasc)
                if idade < IDADE_MIN:
                    F(pid, "DT_NASCIMENTO", r["DT_NASCIMENTO"], "CRITICO", "C003",
                      f"Ativo com {idade:.1f} anos — abaixo do mínimo legal "
                      f"de {IDADE_MIN} anos (CLT Art.403). "
                      f"Erro de digitação ou CPF de dependente.")
                elif idade > IDADE_MAX_ATIVO:
                    F(pid, "DT_NASCIMENTO", r["DT_NASCIMENTO"], "ALERTA", "A001",
                      f"Ativo com {idade:.1f} anos — acima de {IDADE_MAX_ATIVO}. "
                      f"Confirmar com RH se permanece em atividade.")

        # ── C004/C005: data de admissão ──────────────────────────────────────
        dt_adm = _parse(r.get("DT_ADMISSAO_PLANO"))
        if dt_adm:
            if dt_adm > REF_DATE:
                F(pid, "DT_ADMISSAO_PLANO", r["DT_ADMISSAO_PLANO"], "CRITICO","C004",
                  "Admissão ao plano posterior à data-base. "
                  "Participante não deve constar na base de avaliação.")
            if dt_nasc and dt_adm < dt_nasc:
                F(pid, "DT_ADMISSAO_PLANO", r["DT_ADMISSAO_PLANO"], "CRITICO","C005",
                  "Admissão ao plano anterior ao nascimento — impossível. "
                  "Tempo de serviço calculado incorretamente → PMBaC errada.")
            if dt_nasc and (dt_adm - dt_nasc).days / 365.25 < IDADE_MIN:
                F(pid, "DT_ADMISSAO_PLANO", r["DT_ADMISSAO_PLANO"], "ALERTA","A002",
                  f"Admissão quando participante tinha menos de {IDADE_MIN} anos.")

        # ── C006: salário ────────────────────────────────────────────────────
        sal = r.get("SALARIO_CONTRIB")
        if pd.notna(sal):
            sal = float(sal)
            if sal < SMN_2024:
                F(pid, "SALARIO_CONTRIB", f"R$ {sal:,.2f}", "CRITICO","C006",
                  f"Salário (R$ {sal:,.2f}) abaixo do SMN 2024 "
                  f"(R$ {SMN_2024:,.2f}). No método PUC, salário errado "
                  f"propaga para o benefício projetado e para a PMBaC inteira.")
            elif sal > 100_000:
                F(pid, "SALARIO_CONTRIB", f"R$ {sal:,.2f}", "ALERTA","A003",
                  f"Salário (R$ {sal:,.2f}) acima de R$ 100.000. Confirmar com RH.")

        # ── C007: sexo ───────────────────────────────────────────────────────
        sexo = str(r.get("SEXO","")).strip().upper()
        if sexo not in SEXOS_OK:
            F(pid, "SEXO", r.get("SEXO"), "CRITICO","C007",
              f"Sexo inválido: '{sexo}'. A tábua biométrica (BR-EMS 2021) é "
              f"selecionada por sexo — erro aqui muda o fator ä₆₅ em até 16%.")

        # ── C008: situação ────────────────────────────────────────────────────
        sit = str(r.get("SITUACAO","")).strip().upper()
        if sit not in SITUACOES_OK:
            F(pid, "SITUACAO", r.get("SITUACAO"), "CRITICO","C008",
              f"Situação '{sit}' não reconhecida pela PREVIC. "
              f"Válidos: {sorted(SITUACOES_OK)}. Ref: Res.PREVIC 23/2023.")

    # ── C009: CPF duplicado (nível portfólio) ─────────────────────────────────
    dup = df["CPF"].value_counts()
    dup = dup[dup > 1].index
    for _, r in df[df["CPF"].isin(dup)].iterrows():
        F(r["ID_PARTICIPANTE"], "CPF", r["CPF"], "CRITICO","C009",
          f"CPF duplicado ({df['CPF'].value_counts()[r['CPF']]}x). "
          f"Risco de calcular PMBaC em duplicidade — passivo inflado.")

    return pd.DataFrame(issues) if issues else _empty()


def validate_assistidos(df):
    issues = []

    def F(pid, campo, valor, sev, cod, desc):
        issues.append({"ID_PARTICIPANTE": pid, "GRUPO": "ASSISTIDO",
                        "CAMPO": campo, "VALOR_ATUAL": str(valor),
                        "SEVERIDADE": sev, "CODIGO": cod, "DESCRICAO": desc})

    for _, r in df.iterrows():
        pid = r.get("ID_PARTICIPANTE","?")

        for c in ["DT_NASCIMENTO","BENEFICIO_MENSAL","DT_INICIO_BENEFICIO",
                   "SEXO","TIPO_BENEFICIO"]:
            if pd.isna(r.get(c)):
                F(pid, c, "NULO", "CRITICO","C010",
                  f"Campo obrigatório '{c}' ausente em assistido. "
                  f"Impede cálculo da PMBC.")

        ben = r.get("BENEFICIO_MENSAL")
        if pd.notna(ben):
            ben = float(ben)
            if ben <= 0:
                F(pid, "BENEFICIO_MENSAL", f"R$ {ben:,.2f}", "CRITICO","C011",
                  f"Benefício nulo ou negativo (R$ {ben:,.2f}). "
                  f"PMBC desse assistido seria zero — passivo subestimado, "
                  f"risco de falso superávit no plano.")
            elif ben < SMN_2024:
                F(pid, "BENEFICIO_MENSAL", f"R$ {ben:,.2f}", "ALERTA","A004",
                  f"Benefício (R$ {ben:,.2f}) abaixo do SMN. Verificar.")
            elif ben > 80_000:
                F(pid, "BENEFICIO_MENSAL", f"R$ {ben:,.2f}", "ALERTA","A005",
                  f"Benefício (R$ {ben:,.2f}) acima de R$ 80.000. Confirmar.")

        dt_nasc = _parse(r.get("DT_NASCIMENTO"))
        dt_ini  = _parse(r.get("DT_INICIO_BENEFICIO"))
        tipo    = str(r.get("TIPO_BENEFICIO","")).upper()
        if dt_nasc and dt_ini:
            idade_conc = (dt_ini - dt_nasc).days / 365.25
            if "PROG" in tipo and idade_conc < IDADE_MIN_APOS:
                F(pid, "DT_INICIO_BENEFICIO", r.get("DT_INICIO_BENEFICIO"),
                  "ALERTA","A006",
                  f"Aposentadoria programada concedida aos {idade_conc:.1f} anos "
                  f"(mínimo regulamentar: {IDADE_MIN_APOS}). "
                  f"Pode ser benefício de invalidez lançado com tipo errado.")

        sexo = str(r.get("SEXO","")).strip().upper()
        if sexo not in SEXOS_OK:
            F(pid, "SEXO", r.get("SEXO"), "CRITICO","C007",
              f"Sexo inválido: '{sexo}'. Impacta seleção de tábua biométrica.")

    return pd.DataFrame(issues) if issues else _empty()


def validate_diferidos(df):
    issues = []

    def F(pid, campo, valor, sev, cod, desc):
        issues.append({"ID_PARTICIPANTE": pid, "GRUPO": "DIFERIDO",
                        "CAMPO": campo, "VALOR_ATUAL": str(valor),
                        "SEVERIDADE": sev, "CODIGO": cod, "DESCRICAO": desc})

    for _, r in df.iterrows():
        pid = r.get("ID_PARTICIPANTE","?")
        saldo = r.get("SALDO_CONTA")
        if pd.notna(saldo) and float(saldo) <= 0:
            F(pid, "SALDO_CONTA", saldo, "CRITICO","C012",
              "Saldo de conta nulo ou negativo para diferido.")
        dt_nasc = _parse(r.get("DT_NASCIMENTO"))
        if dt_nasc and _age(dt_nasc) > 70:
            F(pid, "DT_NASCIMENTO", r.get("DT_NASCIMENTO"), "ALERTA","A007",
              f"Diferido com {_age(dt_nasc):.1f} anos sem ter requerido benefício. "
              f"Verificar se está vivo e notificado.")

    return pd.DataFrame(issues) if issues else _empty()


def _empty():
    return pd.DataFrame(columns=["ID_PARTICIPANTE","GRUPO","CAMPO",
                                   "VALOR_ATUAL","SEVERIDADE","CODIGO","DESCRICAO"])


def population_summary(df_a, df_s, df_d):
    def ages(df):
        return pd.to_datetime(df["DT_NASCIMENTO"], dayfirst=True,
                               errors="coerce").apply(
            lambda x: (pd.Timestamp(REF_DATE) - x).days / 365.25
            if pd.notna(x) else np.nan)

    return {
        "data_base":                   REF_DATE.strftime("%d/%m/%Y"),
        "n_ativos":                    len(df_a),
        "n_assistidos":                len(df_s),
        "n_diferidos":                 len(df_d),
        "n_total":                     len(df_a)+len(df_s)+len(df_d),
        "ativos_media_idade":          round(ages(df_a).mean(), 1),
        "ativos_media_salario":        round(df_a["SALARIO_CONTRIB"].mean(), 2),
        "ativos_total_massa":          round(df_a["SALARIO_CONTRIB"].sum(), 2),
        "assistidos_media_idade":      round(ages(df_s).mean(), 1),
        "assistidos_media_beneficio":  round(df_s["BENEFICIO_MENSAL"].mean(), 2),
        "assistidos_total_beneficio":  round(df_s["BENEFICIO_MENSAL"].sum(), 2),
        "razao_assistidos_ativos":     round(len(df_s)/len(df_a), 3),
    }
