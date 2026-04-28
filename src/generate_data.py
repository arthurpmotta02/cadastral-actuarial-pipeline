"""
generate_data.py
Synthetic EFPC participant base — calibrated to a mid-size Brazilian pension fund.
930 participants: 600 active, 250 retired, 80 vested deferred.
Inconsistencies injected to demonstrate the validation pipeline.
"""

import numpy as np
import pandas as pd
from datetime import date, timedelta
import random, os

np.random.seed(42)
random.seed(42)

REF_DATE = date(2024, 12, 31)

def _random_date(y0, y1):
    s = date(y0, 1, 1)
    e = date(y1, 12, 31)
    return s + timedelta(days=random.randint(0, (e - s).days))

def _cpf(i):
    return f"{i:09d}-{i % 99:02d}"

NOMES = ["Ana","Bruno","Carlos","Diana","Eduardo","Fernanda","Gabriel","Helena",
         "Igor","Juliana","Klaus","Larissa","Marcos","Natalia","Osvaldo","Patricia",
         "Quirino","Renata","Sergio","Tatiana","Ulisses","Vanessa","Wagner","Ximena",
         "Yuri","Zelia","Andre","Beatriz","Caio","Denise","Elisa","Fabio"]
SOBS  = ["Silva","Santos","Oliveira","Souza","Lima","Pereira","Costa","Ferreira",
         "Rodrigues","Alves","Nascimento","Carvalho","Gomes","Martins","Rocha"]
CARGOS = ["Analista Sr","Analista Pl","Analista Jr","Especialista","Coordenador",
          "Gerente","Assistente","Técnico","Auditor","Engenheiro","Contador"]

def _nome(): return f"{random.choice(NOMES)} {random.choice(SOBS)}"

def gerar_ativos(n=600):
    rows = []
    for i in range(n):
        dt_nasc    = _random_date(1960, 1999)
        dt_adm_emp = _random_date(max(1990, dt_nasc.year+22),
                                   min(2024, dt_nasc.year+42))
        dt_adm_pl  = dt_adm_emp + timedelta(days=random.randint(0, 180))
        tempo      = (REF_DATE - dt_adm_pl).days / 365
        salario    = round(max(6000, np.random.lognormal(np.log(18000), 0.55))
                           * (1 + 0.01 * tempo), 2)
        rows.append({
            "ID_PARTICIPANTE":       f"ACT{i+1:05d}",
            "CPF":                   _cpf(i),
            "NOME":                  _nome(),
            "SEXO":                  random.choice(["M","F","M","M","F"]),
            "DT_NASCIMENTO":         dt_nasc.strftime("%d/%m/%Y"),
            "DT_ADMISSAO_EMPRESA":   dt_adm_emp.strftime("%d/%m/%Y"),
            "DT_ADMISSAO_PLANO":     dt_adm_pl.strftime("%d/%m/%Y"),
            "SITUACAO":              "ATIVO",
            "CARGO":                 random.choice(CARGOS),
            "SALARIO_CONTRIB":       salario,
            "CONTRIB_NORMAL_PART":   round(salario * 0.06, 2),
            "CONTRIB_NORMAL_PATOC":  round(salario * 0.06, 2),
            "SALDO_CONTA":           round(salario * tempo * 0.15, 2),
            "GRUPO_CUSTEIO":         random.choice(["GC1","GC1","GC1","GC2"]),
            "PLANO":                 "BD-PRINCIPAL",
        })
    return pd.DataFrame(rows)

def gerar_assistidos(n=250):
    rows = []
    for i in range(n):
        dt_nasc = _random_date(1940, 1965)
        dt_apos = _random_date(max(1990, dt_nasc.year+55), 2024)
        benef   = round(max(3500, np.random.lognormal(np.log(12000), 0.4)), 2)
        tipo    = random.choice(["APOSENTADORIA_PROG","APOSENTADORIA_PROG",
                                  "APOSENTADORIA_PROG","PENSAO","INVALIDEZ"])
        rows.append({
            "ID_PARTICIPANTE":      f"ASS{i+1:05d}",
            "CPF":                  _cpf(10000+i),
            "NOME":                 _nome(),
            "SEXO":                 random.choice(["M","F","M","F","F"]),
            "DT_NASCIMENTO":        dt_nasc.strftime("%d/%m/%Y"),
            "DT_ADMISSAO_PLANO":    _random_date(1975, 1995).strftime("%d/%m/%Y"),
            "SITUACAO":             "ASSISTIDO",
            "TIPO_BENEFICIO":       tipo,
            "BENEFICIO_MENSAL":     benef,
            "DT_INICIO_BENEFICIO":  dt_apos.strftime("%d/%m/%Y"),
            "GRUPO_CUSTEIO":        "GC1",
            "PLANO":                "BD-PRINCIPAL",
        })
    return pd.DataFrame(rows)

def gerar_diferidos(n=80):
    rows = []
    for i in range(n):
        dt_nasc = _random_date(1970, 1995)
        dt_saida = _random_date(2010, 2023)
        saldo   = round(max(5000, np.random.lognormal(np.log(80000), 0.6)), 2)
        rows.append({
            "ID_PARTICIPANTE":    f"DIF{i+1:05d}",
            "CPF":                _cpf(20000+i),
            "NOME":               _nome(),
            "SEXO":               random.choice(["M","F"]),
            "DT_NASCIMENTO":      dt_nasc.strftime("%d/%m/%Y"),
            "DT_ADMISSAO_PLANO":  _random_date(2000, 2015).strftime("%d/%m/%Y"),
            "SITUACAO":           "DIFERIDO",
            "SALDO_CONTA":        saldo,
            "DT_SAIDA_EMPRESA":   dt_saida.strftime("%d/%m/%Y"),
            "GRUPO_CUSTEIO":      "GC1",
            "PLANO":              "BD-PRINCIPAL",
        })
    return pd.DataFrame(rows)

def inject_inconsistencies(df_a, df_s, df_d):
    n_a = len(df_a); n_s = len(df_s)

    # C001 — campos obrigatórios nulos
    df_a.loc[np.random.choice(n_a, int(n_a*0.02), replace=False), "SALARIO_CONTRIB"] = np.nan
    df_a.loc[np.random.choice(n_a, int(n_a*0.01), replace=False), "DT_NASCIMENTO"]   = np.nan

    # C003 — menor de 16
    for idx in np.random.choice(n_a, 4, replace=False):
        df_a.loc[idx, "DT_NASCIMENTO"] = date(2010, 6, 15).strftime("%d/%m/%Y")

    # C003b — suspeito >75 ativo
    for idx in np.random.choice(n_a, 4, replace=False):
        df_a.loc[idx, "DT_NASCIMENTO"] = date(1940, 3, 20).strftime("%d/%m/%Y")

    # C004 — admissão futura
    for idx in np.random.choice(n_a, 5, replace=False):
        df_a.loc[idx, "DT_ADMISSAO_PLANO"] = date(2025, 6, 1).strftime("%d/%m/%Y")

    # C005 — admissão antes do nascimento
    for idx in np.random.choice(n_a, 3, replace=False):
        s = df_a.loc[idx, "DT_NASCIMENTO"]
        if pd.notna(s):
            try:
                dt = pd.to_datetime(s, dayfirst=True)
                df_a.loc[idx, "DT_ADMISSAO_PLANO"] = (
                    dt - timedelta(days=500)).strftime("%d/%m/%Y")
            except: pass

    # C006 — salário abaixo do mínimo
    df_a.loc[np.random.choice(n_a, 6, replace=False), "SALARIO_CONTRIB"] = 800.0

    # C006b — salário acima de 100k (alerta)
    df_a.loc[np.random.choice(n_a, 4, replace=False), "SALARIO_CONTRIB"] = 250000.0

    # C007 — sexo inválido
    df_a.loc[np.random.choice(n_a, 5, replace=False), "SEXO"] = "X"

    # C008 — situação inválida
    df_a.loc[np.random.choice(n_a, 3, replace=False), "SITUACAO"] = "INATIVO"

    # C009 — CPF duplicado
    for pair in [(0,1),(2,3)]:
        df_a.loc[pair[1], "CPF"] = df_a.loc[pair[0], "CPF"]

    # C010/C011 — benefício nulo em assistido
    df_s.loc[np.random.choice(n_s, 4, replace=False), "BENEFICIO_MENSAL"] = 0.0

    # A006 — aposentadoria muito cedo
    for idx in np.random.choice(n_s, 5, replace=False):
        s = df_s.loc[idx, "DT_NASCIMENTO"]
        try:
            dt = pd.to_datetime(s, dayfirst=True)
            df_s.loc[idx, "DT_INICIO_BENEFICIO"] = (
                dt + timedelta(days=365*45)).strftime("%d/%m/%Y")
        except: pass

    return df_a, df_s, df_d

def generate_all(output_dir="../data/raw"):
    os.makedirs(output_dir, exist_ok=True)
    df_a = gerar_ativos(600)
    df_s = gerar_assistidos(250)
    df_d = gerar_diferidos(80)
    df_a, df_s, df_d = inject_inconsistencies(df_a, df_s, df_d)

    path = f"{output_dir}/base_cadastral_2024.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df_a.to_excel(w, sheet_name="ATIVOS",     index=False)
        df_s.to_excel(w, sheet_name="ASSISTIDOS", index=False)
        df_d.to_excel(w, sheet_name="DIFERIDOS",  index=False)

    print(f"Base gerada: {path}")
    print(f"  Ativos: {len(df_a)}  Assistidos: {len(df_s)}  Diferidos: {len(df_d)}")
    return df_a, df_s, df_d

if __name__ == "__main__":
    generate_all()
