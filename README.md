# Crítica da Base Cadastral — Avaliação Atuarial EFPC

Pipeline Python que valida e prepara dados cadastrais de participantes de fundos de pensão (EFPC) antes da avaliação atuarial anual, conforme exigências regulatórias da PREVIC.

---

## Por que esse projeto existe

A Resolução PREVIC nº 7/2022 (Art. 8) exige que toda EFPC archive seus dados cadastrais em planilha eletrônica antes de cada avaliação atuarial. O CPA 017/2019 do IBA (Instituto Brasileiro de Atuária) define que **a crítica da base cadastral é o primeiro passo obrigatório de toda auditoria atuarial**:

> "O Atuário Independente deve verificar se as informações sobre os participantes e assistidos do plano de benefícios estão alinhadas com os registros internos."

Na prática esse processo é feito manualmente em Excel. Este pipeline automatiza as verificações em ~2 segundos, gera o relatório atuarial formatado, exporta os dados para Power BI e disponibiliza um dashboard interativo via Docker.

---

## Stack

| Camada | Tecnologia | Uso |
|---|---|---|
| Validação + relatório | Python, pandas, openpyxl | Pipeline CLI |
| Dashboard interativo | Plotly Dash, Flask/gunicorn | 4 páginas dark theme |
| Deploy empresarial | Docker, docker-compose | Self-hosted, dados na rede interna |
| BI corporativo | Power BI Desktop | `powerbi_data.xlsx` pronto para importar |

---

## Início rápido

### Opção A — Python direto

```bash
# Setup
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt

# Rodar pipeline (gera os Excels)
python src/pipeline.py --demo

# Rodar dashboard
python app/dashboard.py      # http://localhost:8050
```

### Opção B — Docker (recomendado para uso empresarial)

```bash
# Build e start
docker compose up -d

# Acessar
# http://localhost:8050

# Ver logs
docker compose logs -f

# Parar
docker compose down
```

> **Vantagem empresarial:** dados de participantes nunca saem da rede interna.
> Adequado para EFPCs com restrições LGPD e políticas de segurança da informação.

---

## Saídas geradas

```
python src/pipeline.py --demo
```

| Arquivo | Destinatário | Conteúdo |
|---|---|---|
| `results/reports/relatorio_critica_cadastral.xlsx` | Atuário responsável | 4 abas: sumário executivo, inconsistências detalhadas, base limpa, análise por tipo |
| `results/reports/powerbi_data.xlsx` | Analista de BI | 6 tabelas flat: população, inconsistências, KPIs — prontas para Power BI Desktop |

---

## Dashboard (Plotly Dash)

4 páginas no sidebar:

- **Visão Geral** — KPIs de população, gráfico de inconsistências por código, donut de composição
- **Inconsistências** — frequência por código + tabela completa com badges CRÍTICO/ALERTA
- **Ativos** — distribuição etária, salarial, scatter idade×salário, breakdown por cargo
- **Assistidos** — distribuição etária, de benefícios, por tipo, scatter idade×benefício

---

## Power BI

Ver `powerbi/INSTRUCOES_POWER_BI.md` para:
- Conectar o `powerbi_data.xlsx` no Power BI Desktop
- Criar relacionamentos entre tabelas
- Medidas DAX prontas (N Críticos, Razão Assistidos/Ativos, etc.)
- 4 páginas de dashboard sugeridas

---

## Validações implementadas (19 verificações)

| Código | Severidade | Verificação | Impacto Atuarial |
|---|---|---|---|
| C001 | CRÍTICO | Campo obrigatório ausente | Impede cálculo de PMBaC/PMBC |
| C002 | CRÍTICO | Data de nascimento no futuro | Impossível logicamente |
| C003 | CRÍTICO | Ativo com menos de 16 anos | CLT Art. 403 |
| C004 | CRÍTICO | Admissão ao plano após data-base | Não deveria constar na avaliação |
| C005 | CRÍTICO | Admissão anterior ao nascimento | Tempo de serviço e PMBaC incorretos |
| C006 | CRÍTICO | Salário abaixo do SMN 2024 (R$ 1.412) | No PUC, propaga para o benefício projetado inteiro |
| C007 | CRÍTICO | Código de sexo inválido | Tábua biométrica errada — ä₆₅ muda ±16% |
| C008 | CRÍTICO | Situação não reconhecida pela PREVIC | Grupo de custeio incorreto |
| C009 | CRÍTICO | CPF duplicado | PMBaC calculada em duplicidade |
| C010 | CRÍTICO | Campo obrigatório ausente em assistido | Impede cálculo da PMBC |
| C011 | CRÍTICO | Benefício nulo ou negativo | PMBC = 0 → passivo subestimado, risco de falso superávit |
| C012 | CRÍTICO | Saldo de conta nulo em diferido | Reserva de BPD incorreta |
| A001 | ALERTA | Ativo com mais de 75 anos | Confirmar permanência em atividade |
| A002 | ALERTA | Admissão quando participante tinha < 16 anos | Verificar com RH |
| A003 | ALERTA | Salário acima de R$ 100.000 | Confirmar com RH |
| A004 | ALERTA | Benefício abaixo do SMN | Verificar se correto |
| A005 | ALERTA | Benefício acima de R$ 80.000 | Confirmar com cadastro |
| A006 | ALERTA | Aposentadoria programada antes dos 55 anos | Pode ser invalidez com tipo errado |
| A007 | ALERTA | Diferido acima de 70 anos sem benefício | Verificar se está vivo e foi notificado |

---

## Resultados (base demo — 930 participantes)

| Métrica | Valor |
|---|---|
| Inconsistências críticas | 63 |
| Inconsistências alertas | 26 |
| Razão assistidos/ativos | 0.417 (fundo maduro) |
| Massa salarial mensal | R$ 14.7M |
| Total benefícios/mês | R$ 3.2M |
| Tempo de execução | ~2s |

---

## Estrutura do projeto

```
cadastral-actuarial-pipeline/
│
├── Dockerfile                    ← build da imagem Docker
├── docker-compose.yml            ← orquestração (1 comando para subir)
├── requirements.txt              ← dependências Python
├── README.md
│
├── src/
│   ├── pipeline.py               ← entry point CLI
│   ├── generate_data.py          ← gerador de base demo (930 participantes)
│   ├── validator.py              ← 19 verificações regulatórias
│   └── report_generator.py      ← Excel atuarial + Power BI flat tables
│
├── app/
│   ├── dashboard.py              ← Plotly Dash (4 páginas, dark theme)
│   └── assets/style.css         ← CSS customizado
│
├── powerbi/
│   └── INSTRUCOES_POWER_BI.md   ← guia completo com DAX e páginas sugeridas
│
└── data/
    ├── raw/                      ← base recebida do RH (.xlsx)
    └── processed/                ← base limpa exportada pelo pipeline
```

---

## Referências regulatórias

- **Resolução PREVIC nº 7/2022, Art. 8** — dados cadastrais devem ser arquivados em planilha eletrônica
- **Resolução PREVIC nº 23/2023** — norma consolidada EFPC (classificação ATIVO/ASSISTIDO/DIFERIDO)
- **CPA 017/2019 IBA** — Auditoria Atuarial e de Benefícios: crítica cadastral é o primeiro passo
- **CPAO 035 IBA** — Reservas Matemáticas (PMBaC, PMBC, método PUC)
- **CLT Art. 403** — idade mínima de trabalho (16 anos)

---

## Autor

Arthur Motta — Ciências Atuariais e Estatística, UFRJ
[GitHub](https://github.com/arthurpmotta02) · [LinkedIn](https://linkedin.com/in/arthurpmotta)
