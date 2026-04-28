#!/usr/bin/env python
"""
pipeline.py — Crítica da Base Cadastral EFPC
CLI entry point. Reads an Excel from HR and produces:
  - relatorio_critica_cadastral.xlsx  (actuarial report)
  - powerbi_data.xlsx                 (Power BI flat tables)

Usage:
  python pipeline.py
  python pipeline.py --input data/raw/base_rh.xlsx --output results/reports/

Resolução PREVIC 7/2022, Art. 8 + CPA 017/2019 IBA
"""

import argparse, os, sys, time
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from generate_data import generate_all
from validator import (validate_ativos, validate_assistidos,
                        validate_diferidos, population_summary)
from report_generator import build_report

LINE = "=" * 60

def load_base(path):
    xl = pd.ExcelFile(path)
    df_a = xl.parse("ATIVOS")     if "ATIVOS"     in xl.sheet_names else pd.DataFrame()
    df_s = xl.parse("ASSISTIDOS") if "ASSISTIDOS" in xl.sheet_names else pd.DataFrame()
    df_d = xl.parse("DIFERIDOS")  if "DIFERIDOS"  in xl.sheet_names else pd.DataFrame()
    return df_a, df_s, df_d

def run(input_path, output_dir):
    t0 = time.time()

    print(LINE)
    print("CRÍTICA DA BASE CADASTRAL — AVALIAÇÃO ATUARIAL")
    print("Res. PREVIC 7/2022 + CPA 017/2019 IBA")
    print(LINE)

    # ── 1. Load ───────────────────────────────────────────────────────────────
    print(f"\n[1/4] Lendo base: {input_path}")
    df_a, df_s, df_d = load_base(input_path)
    print(f"      Ativos: {len(df_a):,}  "
          f"Assistidos: {len(df_s):,}  "
          f"Diferidos: {len(df_d):,}")

    # ── 2. Validate ───────────────────────────────────────────────────────────
    print("\n[2/4] Executando crítica cadastral...")
    iss_a = validate_ativos(df_a)
    iss_s = validate_assistidos(df_s)
    iss_d = validate_diferidos(df_d)

    import pandas as pd2
    all_issues = pd.concat([iss_a, iss_s, iss_d], ignore_index=True)
    n_crit  = (all_issues["SEVERIDADE"] == "CRITICO").sum()
    n_alert = (all_issues["SEVERIDADE"] == "ALERTA").sum()

    print(f"      Total de ocorrências: {len(all_issues):,}")
    print(f"      CRÍTICAS: {n_crit:,}")
    print(f"      ALERTAS:  {n_alert:,}")

    if n_crit > 0:
        print(f"\n      ⚠  {n_crit} registros com inconsistências CRÍTICAS "
              f"serão excluídos da base limpa.")

    # ── 3. Summary ────────────────────────────────────────────────────────────
    print("\n[3/4] Calculando estatísticas populacionais...")
    summ = population_summary(df_a, df_s, df_d)
    print(f"      Razão assistidos/ativos: {summ['razao_assistidos_ativos']:.3f}")
    print(f"      Massa salarial mensal:   "
          f"R$ {summ['ativos_total_massa']:>15,.2f}")
    print(f"      Total benefícios/mês:   "
          f"R$ {summ['assistidos_total_beneficio']:>15,.2f}")

    # ── 4. Generate reports ───────────────────────────────────────────────────
    print(f"\n[4/4] Gerando relatórios em: {output_dir}")
    os.makedirs(output_dir, exist_ok=True)
    report_path = os.path.join(output_dir, "relatorio_critica_cadastral.xlsx")
    build_report(df_a, df_s, df_d, iss_a, iss_s, iss_d, summ,
                 output_path=report_path)

    elapsed = time.time() - t0
    print(LINE)
    print(f"Concluído em {elapsed:.1f}s")
    print(f"Relatório:        {report_path}")
    print(f"Dados Power BI:   "
          f"{report_path.replace('relatorio_critica_cadastral', 'powerbi_data')}")
    print(LINE)

    # ── Issues by code (console summary) ─────────────────────────────────────
    if len(all_issues) > 0:
        print("\nInconsistências por código:")
        freq = (all_issues.groupby(["CODIGO","SEVERIDADE"])
                .size().reset_index(name="N")
                .sort_values("N", ascending=False))
        for _, r in freq.iterrows():
            marker = "⛔" if r["SEVERIDADE"] == "CRITICO" else "⚠ "
            print(f"  {marker} {r['CODIGO']:6s}  {r['SEVERIDADE']:8s}  {r['N']:4d}")

    return report_path


def main():
    parser = argparse.ArgumentParser(
        description="Crítica da base cadastral EFPC — pipeline atuarial")
    parser.add_argument("--input",  default=None,
                        help="Caminho do Excel de entrada (padrão: gera base demo)")
    parser.add_argument("--output", default="results/reports",
                        help="Pasta de saída (padrão: results/reports)")
    parser.add_argument("--demo",   action="store_true",
                        help="Gera e usa base demo mesmo que --input seja fornecido")
    args = parser.parse_args()

    if args.demo or args.input is None:
        print("Gerando base de demonstração...")
        generate_all(output_dir="data/raw")
        input_path = "data/raw/base_cadastral_2024.xlsx"
    else:
        input_path = args.input

    run(input_path, args.output)


if __name__ == "__main__":
    main()
