# Como conectar o Power BI ao pipeline

## O que o pipeline gera

Ao rodar `python src/pipeline.py`, dois arquivos são gerados em `results/reports/`:

| Arquivo | Uso |
|---|---|
| `relatorio_critica_cadastral.xlsx` | Relatório atuarial para o atuário responsável |
| `powerbi_data.xlsx` | Dados flat para importar no Power BI Desktop |

---

## Passo a passo no Power BI Desktop

### 1. Importar o arquivo

1. Abra Power BI Desktop
2. **Home → Get Data → Excel Workbook**
3. Selecione `results/reports/powerbi_data.xlsx`
4. No Navigator, marque **todas as 6 abas**:
   - `POP_ATIVOS`
   - `POP_ASSISTIDOS`
   - `POP_DIFERIDOS`
   - `INCONSISTENCIAS`
   - `FREQ_INCONSISTENCIAS`
   - `SUMARIO_KPI`
5. Clique **Load**

---

### 2. Criar os relacionamentos

No **Model View** (ícone de diagrama na barra lateral):

Relacionamentos a criar manualmente:
- `POP_ATIVOS[ID_PARTICIPANTE]` → `INCONSISTENCIAS[ID_PARTICIPANTE]` (1:Many)
- `POP_ASSISTIDOS[ID_PARTICIPANTE]` → `INCONSISTENCIAS[ID_PARTICIPANTE]` (1:Many)
- `POP_DIFERIDOS[ID_PARTICIPANTE]` → `INCONSISTENCIAS[ID_PARTICIPANTE]` (1:Many)

---

### 3. Criar medidas DAX

Na aba **POP_ATIVOS**, botão direito → **New Measure**:

```dax
N Ativos = COUNTROWS(POP_ATIVOS)

N Assistidos = COUNTROWS(POP_ASSISTIDOS)

Razao Ass Ativos = DIVIDE([N Assistidos], [N Ativos])

Salario Medio = AVERAGE(POP_ATIVOS[SALARIO_CONTRIB])

Massa Salarial = SUM(POP_ATIVOS[SALARIO_CONTRIB])

Total Beneficios = SUM(POP_ASSISTIDOS[BENEFICIO_MENSAL])

N Criticos = CALCULATE(
    COUNTROWS(INCONSISTENCIAS),
    INCONSISTENCIAS[SEVERIDADE] = "CRITICO"
)

N Alertas = CALCULATE(
    COUNTROWS(INCONSISTENCIAS),
    INCONSISTENCIAS[SEVERIDADE] = "ALERTA"
)

Pct Criticos = DIVIDE([N Criticos], COUNTROWS(INCONSISTENCIAS))
```

---

### 4. Páginas sugeridas

**Página 1 — Visão Geral da População**
- Cards: N Ativos, N Assistidos, N Diferidos, Razão Ass./Ativos
- Gráfico de barras: Faixa etária × Contagem (por grupo)
- Gráfico de pizza: Distribuição por SITUACAO_GRUPO
- Tabela: Resumo KPI (de SUMARIO_KPI)

**Página 2 — Inconsistências**
- Cards: N Críticos, N Alertas (com formatação condicional em vermelho/laranja)
- Gráfico de barras: CODIGO × OCORRENCIAS (de FREQ_INCONSISTENCIAS)
- Slicer: SEVERIDADE
- Tabela: INCONSISTENCIAS com colunas ID, GRUPO, CAMPO, SEVERIDADE, DESCRICAO

**Página 3 — Ativos**
- Histograma de idades (agrupar FAIXA_ETARIA)
- Scatter: IDADE × SALARIO_CONTRIB
- Mapa de calor: FAIXA_ETARIA × GRUPO_CUSTEIO

**Página 4 — Assistidos**
- Distribuição de benefícios por TIPO_BENEFICIO
- Histograma de idades
- KPI: Benefício médio, total mensal

---

### 5. Atualizar dados (após novo ciclo)

Quando rodar `python src/pipeline.py` novamente com nova base:
1. Power BI Desktop → **Home → Refresh**
2. Os visuais atualizam automaticamente

> Para atualização automática sem abrir o Desktop, salve o `.pbix` em OneDrive
> e publique no Power BI Service — o refresh será automático a cada hora.

---

## Estrutura das tabelas

### POP_ATIVOS (19 colunas)
| Coluna | Tipo | Descrição |
|---|---|---|
| ID_PARTICIPANTE | texto | Identificador único |
| CPF | texto | CPF (mascarado em produção) |
| SEXO | texto | M/F |
| DT_NASCIMENTO | data | |
| DT_ADMISSAO_PLANO | data | |
| SITUACAO | texto | ATIVO |
| SALARIO_CONTRIB | número | Salário de contribuição |
| IDADE | número | Calculado em 31/12/2024 |
| FAIXA_ETARIA | texto | Bucket 10 anos |
| TEMPO_PLANO_ANOS | número | Tempo no plano em anos |

### INCONSISTENCIAS (7 colunas)
| Coluna | Tipo | Descrição |
|---|---|---|
| ID_PARTICIPANTE | texto | Liga com tabelas de população |
| GRUPO | texto | ATIVO / ASSISTIDO / DIFERIDO |
| CAMPO | texto | Campo com problema |
| VALOR_ATUAL | texto | Valor encontrado |
| SEVERIDADE | texto | CRITICO / ALERTA |
| CODIGO | texto | C001..C012, A001..A007 |
| DESCRICAO | texto | Explicação com impacto atuarial |

### SUMARIO_KPI (2 colunas: INDICADOR, VALOR)
Usada para cards no painel de visão geral.

### FREQ_INCONSISTENCIAS (5 colunas)
Pré-agregada para gráfico de barras — sem necessidade de DAX adicional.
