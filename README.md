# Crítica da Base Cadastral — Avaliação Atuarial EFPC

Pipeline Python que valida e prepara dados cadastrais de participantes de fundos de pensão (EFPC) antes da avaliação atuarial anual, conforme exigências regulatórias da PREVIC.

> **Nota:** o `localhost:8050` mencionado abaixo é local — só você consegue acessar
> na sua máquina. Para outros acessarem é necessário rodar o Docker num servidor
> acessível na rede, ou usar deploy em nuvem.

---

## Por que esse projeto existe

A Resolução PREVIC nº 7/2022 (Art. 8) exige que toda EFPC archive seus dados cadastrais em planilha eletrônica antes de cada avaliação atuarial. O CPA 017/2019 do IBA define que **a crítica da base cadastral é o primeiro passo obrigatório de toda auditoria atuarial**:

> "O Atuário Independente deve verificar se as informações sobre os participantes e assistidos estão alinhadas com os registros internos e os dados utilizados pelo Atuário Responsável Técnico."

Na prática esse processo é feito manualmente em Excel. Este pipeline automatiza as verificações em ~2 segundos, gera o relatório atuarial formatado, exporta dados para Power BI e disponibiliza um dashboard interativo.

---

## Stack

| Camada | Tecnologia | Uso |
|---|---|---|
| Validação + relatório | Python · pandas · openpyxl | Pipeline CLI |
| Dashboard interativo | Plotly Dash 4 · Flask · gunicorn | 4 páginas dark theme |
| Deploy empresarial | Docker · docker-compose | Self-hosted, dados na rede interna |
| BI corporativo | Power BI Desktop | `powerbi_data.xlsx` pronto para importar |

---

## Início rápido

### Opção A — Python direto

```bash
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt

# Rodar pipeline — gera os Excels em results/reports/
python src/pipeline.py --demo

# Rodar dashboard — abre em http://localhost:8050
python app/dashboard.py
```

### Opção B — Docker (deploy empresarial)

```bash
# Build e start (requer Docker Desktop instalado e rodando)
docker compose up -d

# Acessar em http://localhost:8050
# Na rede interna da empresa: http://IP-DO-SERVIDOR:8050

# Ver logs
docker compose logs -f

# Parar
docker compose down
```

> **Por que Docker para uso empresarial?**
> Dados de participantes de EFPC são protegidos pela LGPD. Com Docker self-hosted,
> nenhum dado sai da rede interna — o container roda no servidor da própria empresa.

---

## Saídas geradas

| Arquivo | Destinatário | Conteúdo |
|---|---|---|
| `results/reports/relatorio_critica_cadastral.xlsx` | Atuário responsável | 4 abas: sumário executivo, inconsistências, base limpa, análise por tipo |
| `results/reports/powerbi_data.xlsx` | Analista de BI | 6 tabelas flat prontas para Power BI Desktop |

---

## Dashboard (Plotly Dash 4)

4 páginas navegáveis pelo sidebar:

| Página | Conteúdo |
|---|---|
| **Visão Geral** | KPIs de população, gráfico de inconsistências por código, donut de composição |
| **Inconsistências** | Frequência por código + tabela completa com badges CRÍTICO/ALERTA |
| **Ativos** | Distribuição etária, salarial, scatter idade×salário, breakdown por cargo |
| **Assistidos** | Distribuição etária, de benefícios, por tipo, scatter idade×benefício |

---

## Power BI

Ver `powerbi/INSTRUCOES_POWER_BI.md` para passo a passo completo com:
- Conectar `powerbi_data.xlsx` no Power BI Desktop
- Criar relacionamentos entre as 6 tabelas
- Medidas DAX prontas (N Críticos, Razão Assistidos/Ativos, etc.)
- 4 páginas de dashboard sugeridas

---

## Por que Power BI point-and-click não é estado da arte — e o que vem por aí

Esta seção existe porque o projeto entrega um `powerbi_data.xlsx` pronto para importação manual no Power BI Desktop, e é importante ser honesto sobre o que isso significa tecnicamente: é uma concessão à realidade do mercado corporativo brasileiro atual, não uma escolha de engenharia ideal. O argumento a seguir documenta o estado da arte de 2026 e o caminho que este projeto seguirá quando as ferramentas amadurecerem.

### O problema fundamental do Power BI como ferramenta point-and-click

Qualquer ferramenta de BI que exige que um humano arraste visuais, configure relacionamentos manualmente e salve um arquivo binário para gerar um relatório viola um princípio básico da engenharia de software moderna: **reprodutibilidade**. Se o processo de criação de um artefato não pode ser descrito integralmente em código, ele não pode ser versionado, testado, revisado em pull request, auditado, ou reproduzido de forma idêntica por outra pessoa em outra máquina. Ele existe apenas na cabeça de quem clicou.

Isso tem consequências práticas diretas no contexto de uma EFPC:

**1. Ausência de versionamento real.** Um arquivo `.pbix` é um ZIP binário. Fazer `git diff` num `.pbix` não produz nenhuma informação útil — o Git trata o arquivo como um blob opaco. Isso significa que não existe registro legível de quais medidas DAX foram alteradas entre a avaliação atuarial de dezembro e a de março, qual visual foi adicionado e por quê, ou quem mudou qual filtro. Em auditoria atuarial — onde rastreabilidade é um requisito regulatório explícito — isso é uma falha estrutural, não um detalhe cosmético.

**2. Irreprodutibildade sistêmica.** Se o analista que montou o relatório sair da empresa, a capacidade de reproduzir aquele dashboard exato vai junto. O conhecimento está no clique, não no código. Não existe um `README.md` que descreva "execute esses passos e você terá o mesmo resultado". Em contraste, o pipeline Python deste projeto pode ser executado por qualquer pessoa com `python src/pipeline.py --demo` e produzirá exatamente os mesmos outputs em qualquer máquina — isso é reprodutibilidade.

**3. Impossibilidade de CI/CD.** Ferramentas de BI point-and-click não se integram naturalmente a pipelines de integração contínua. Não existe um comando `powerbi build --validate` que você pode rodar no GitHub Actions para garantir que o relatório está consistente com os dados antes de cada deploy. O resultado é que mudanças nos dados (um novo campo no cadastro, uma nova situação regulatória adicionada pela PREVIC) exigem intervenção manual no relatório — um ponto de falha humano.

**4. Colaboração impossível em equipes.** Dois analistas não conseguem trabalhar simultaneamente no mesmo `.pbix` sem sobrescrever o trabalho um do outro. A solução usual — "fulano trabalha no relatório enquanto cicrano não mexe" — é o equivalente a desenvolver software sem controle de versão. Em qualquer equipe de engenharia de software isso seria inaceitável.

**5. Acoplamento a licença proprietária.** O `.pbix` só abre no Power BI Desktop, que só roda no Windows, que exige conta Microsoft ativa. O Dash roda em qualquer sistema operacional, em qualquer browser, sem licença, com `python app/dashboard.py`.

### O que o mercado está fazendo para resolver isso

A Microsoft reconheceu explicitamente esses problemas e está construindo uma solução. O estado da arte em 2026 é composto de três tecnologias complementares:

**PBIR — Power BI Enhanced Report Format**
Lançado em preview em 2024 e tornado padrão em janeiro de 2026 para o Power BI Service (com Power BI Desktop seguindo em março de 2026), o PBIR decompõe o arquivo monolítico `.pbix` em uma estrutura de pastas com arquivos JSON individuais para cada página, visual, bookmark e interação. Cada arquivo tem um schema JSON público e documentado pela Microsoft. Isso significa que, pela primeira vez, é possível fazer `git diff` em um relatório Power BI e ver exatamente qual propriedade de qual visual mudou. É Power BI finalmente tratando relatórios como código.

Fonte: [Microsoft Learn — Power BI Enhanced Report Format](https://learn.microsoft.com/en-us/power-bi/developer/embedded/projects-enhanced-report-format)

**PBIP — Power BI Project**
O formato de projeto que organiza o PBIR (relatório) e o TMDL (modelo semântico) em uma estrutura de diretórios versionável. Com PBIP, um relatório Power BI passa a ser um repositório Git como qualquer outro projeto de software — com histórico, branches, pull requests e merge conflicts resolvíveis.

Fonte: [Microsoft Learn — Power BI Desktop Projects](https://learn.microsoft.com/en-us/power-bi/developer/projects/projects-overview)

**`powerbpy` — geração programática de relatórios Power BI via Python**
Biblioteca Python open-source que gera a estrutura PBIP/PBIR inteiramente por código, sem abrir o Power BI Desktop. O resultado pode ser aberto e editado normalmente no Desktop, e versionado no Git como qualquer outro arquivo de texto.

Fonte: [powerbpy no PyPI](https://pypi.org/project/powerbpy/)

### Por que este projeto ainda não usa essas ferramentas

A resposta é direta: o PBIR ainda está em preview. A Microsoft adverte explicitamente que a especificação pode mudar antes do GA (previsto para Q3 2026), e o `powerbpy` é um projeto de desenvolvedor independente cujo roadmap depende da estabilidade do PBIR. Construir um pipeline de produção sobre uma especificação em preview seria trocar um problema (point-and-click não reprodutível) por outro (código que quebra quando a Microsoft muda a spec).

A escolha pragmática — exportar `powerbi_data.xlsx` bem estruturado para importação manual — é a decisão correta para o momento. Ela satisfaz o requisito real (analistas de EFPC que precisam de Power BI hoje) sem amarrar o projeto a uma tecnologia instável.

### O que este projeto fará quando o PBIR atingir GA

Assim que o PBIR sair de preview (previsão Q3 2026), este projeto será atualizado com:

```
src/
└── powerbi_generator.py   ← gera a estrutura PBIP/PBIR
                              inteiramente por Python, sem abrir o Desktop
                              usando powerbpy + manipulação direta dos JSON PBIR

powerbi/
├── INSTRUCOES_POWER_BI.md  ← mantido para quem preferir o fluxo manual
└── report/                 ← estrutura PBIP gerada automaticamente
    ├── definition.pbir
    ├── pages/
    └── visuals/
```

O pipeline passará a ser:

```bash
python src/pipeline.py --demo        # gera dados + relatório Excel
python src/powerbi_generator.py      # gera relatório Power BI por código
# Resultado: pasta powerbi/report/ pronta para abrir no Desktop
# ou fazer push via Fabric Git Integration
```

Isso fechará o ciclo de reprodutibilidade completa: de dados brutos a relatório Power BI, tudo em código, tudo versionado, tudo executável com um comando.

---

## Validações implementadas (19 verificações)

| Código | Severidade | Verificação | Impacto Atuarial |
|---|---|---|---|
| C001 | CRÍTICO | Campo obrigatório ausente | Impede cálculo de PMBaC/PMBC |
| C002 | CRÍTICO | Data de nascimento no futuro | Impossível logicamente |
| C003 | CRÍTICO | Ativo com menos de 16 anos | CLT Art. 403 |
| C004 | CRÍTICO | Admissão ao plano após data-base | Não deveria constar na avaliação |
| C005 | CRÍTICO | Admissão anterior ao nascimento | Tempo de serviço e PMBaC incorretos |
| C006 | CRÍTICO | Salário abaixo do SMN 2024 (R$ 1.412) | No método PUC, propaga para todo o benefício projetado |
| C007 | CRÍTICO | Código de sexo inválido | Tábua biométrica errada — ä₆₅ muda ±16% |
| C008 | CRÍTICO | Situação não reconhecida pela PREVIC | Grupo de custeio incorreto |
| C009 | CRÍTICO | CPF duplicado | PMBaC calculada em duplicidade — passivo inflado |
| C010 | CRÍTICO | Campo obrigatório ausente em assistido | Impede cálculo da PMBC |
| C011 | CRÍTICO | Benefício nulo ou negativo | PMBC = 0 → passivo subestimado, risco de falso superávit |
| C012 | CRÍTICO | Saldo de conta nulo em diferido | Reserva de BPD incorreta |
| A001 | ALERTA | Ativo com mais de 75 anos | Confirmar permanência em atividade |
| A002 | ALERTA | Admissão quando tinha < 16 anos | Verificar com RH |
| A003 | ALERTA | Salário acima de R$ 100.000 | Confirmar com RH |
| A004 | ALERTA | Benefício abaixo do SMN | Verificar se correto |
| A005 | ALERTA | Benefício acima de R$ 80.000 | Confirmar com cadastro |
| A006 | ALERTA | Aposentadoria programada antes dos 55 anos | Pode ser invalidez lançada com tipo errado |
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
├── Dockerfile                    ← imagem Docker para deploy empresarial
├── docker-compose.yml            ← sobe tudo com 1 comando
├── requirements.txt
├── README.md
│
├── src/
│   ├── pipeline.py               ← entry point CLI
│   ├── generate_data.py          ← gerador de base demo (930 participantes)
│   ├── validator.py              ← 19 verificações regulatórias
│   └── report_generator.py      ← Excel atuarial + Power BI flat tables
│
├── app/
│   ├── dashboard.py              ← Plotly Dash 4 (4 páginas, dark theme)
│   └── assets/style.css
│
├── powerbi/
│   └── INSTRUCOES_POWER_BI.md
│
└── data/
    ├── raw/                      ← base recebida do RH (.xlsx)
    └── processed/
```

---

## Referências regulatórias

- **Resolução PREVIC nº 7/2022, Art. 8** — dados cadastrais em planilha eletrônica
- **Resolução PREVIC nº 23/2023** — norma consolidada EFPC
- **CPA 017/2019 IBA** — Auditoria Atuarial: crítica cadastral é o primeiro passo
- **CPAO 035 IBA** — Reservas Matemáticas (PMBaC, PMBC, método PUC)
- **CLT Art. 403** — idade mínima de trabalho (16 anos)

---

## Autor

Arthur Motta — Ciências Atuariais e Estatística, UFRJ
[GitHub](https://github.com/arthurpmotta02) · [LinkedIn](https://linkedin.com/in/arthurpmotta)