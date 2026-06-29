<div align="center">

<img src="https://img.shields.io/badge/HanParal-v1.0-8B4A38?style=for-the-badge" alt="HanParal v1.0"/>
<img src="https://img.shields.io/badge/Python-3.8+-C07A5A?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.8+"/>
<img src="https://img.shields.io/badge/License-MIT-E0B090?style=for-the-badge" alt="MIT License"/>
<img src="https://img.shields.io/badge/Google_Colab-Ready-F9AB00?style=for-the-badge&logo=googlecolab&logoColor=white" alt="Google Colab"/>
<img src="https://img.shields.io/badge/DOI-10.5281%2Fzenodo.21048473-8B4A38?style=for-the-badge" alt="DOI"/>

# 🌸 HanParal

**Multilingual Literary Concordancer**

Python-based concordance tool for comparative translation analysis of culture-specific items in literary texts
·
Ferramenta de concordância multilíngue para análise tradutória de itens culturais específicos em textos literários

*Language-neutral · Corpus-agnostic · Thesis-ready*

</div>

---

## What it does / O que faz

**EN** — HanParal is a two-phase Google Colab workflow designed for translation studies researchers working with manually aligned multilingual literary corpora. It automates the most time-consuming parts of CSI analysis: finding every occurrence of a source-text item, preparing a structured annotation file, and generating thesis-ready summary tables and visualisations once annotation is complete.

**PT** — HanParal é um fluxo de trabalho em duas fases no Google Colab, desenvolvido para pesquisadoras e pesquisadores de estudos da tradução que trabalham com corpora literários multilíngues alinhados manualmente. A ferramenta automatiza as etapas mais trabalhosas da análise de ICEs: localizar todas as ocorrências de um item do texto-fonte, preparar um arquivo de anotação estruturado e gerar tabelas-resumo e visualizações prontas para a monografia ou dissertação.

---

## Workflow / Fluxo de trabalho

```
corpus.xlsx
     │
     ▼
Phase 1 / Fase 1
  ├── upload corpus
  ├── choose search mode: single_term · multi_term · category
  ├── KWIC context window (source + target texts)
  ├── optional: equivalent detection
  ├── prepare annotation table
  └── export annotated .xlsx ──► manual annotation in Excel
                                        │
                                        ▼
                               Phase 2 / Fase 2
                                 ├── upload annotated file
                                 ├── count strategies + percentages
                                 ├── generate thesis summary tables
                                 ├── strategy distribution chart
                                 ├── term × strategy heatmap (optional)
                                 └── export final analysis workbook .xlsx
```

---

## Input format / Formato de entrada

**EN** — Your corpus must be an `.xlsx` or `.xlsm` file with at least these columns:

**PT** — Seu corpus deve ser um arquivo `.xlsx` ou `.xlsm` com pelo menos estas colunas:

| `id` | `section` | `source_text` | `target_1` | `target_2` | `...` |
|------|-----------|---------------|------------|------------|-------|
| 0001 | Chapter 1 | 도청 앞에 사람들이... | As pessoas se reuniram... | People gathered... | ... |

**EN** — Optional columns: `term_source`, `category`, `strategy_target_1`, `strategy_target_2`, ...

**PT** — Colunas opcionais: `term_source`, `category`, `strategy_target_1`, `strategy_target_2`, ...

The number of target-text columns is fully configurable via `N_TARGETS`.
O número de colunas de texto-alvo é configurável via `N_TARGETS`.

---

## Features / Funcionalidades

### Phase 1 / Fase 1

**EN** | **PT**
--- | ---
Single-term, multi-term, and category-based search | Busca por termo único, múltiplos termos ou categoria
Token-level KWIC context windows | Janelas de contexto KWIC por ocorrência de token
Equivalent detection across all target texts | Detecção de equivalentes em todos os textos-alvo
Structured annotation table with Excel dropdowns | Tabela de anotação com menus suspensos no Excel
Category dropdown built into the output file | Menu de categoria integrado ao arquivo de saída

### Phase 2 / Fase 2

**EN** | **PT**
--- | ---
Strategy counts and percentages per language | Contagem e percentual de estratégias por língua
Strategy distribution chart (horizontal bar) | Gráfico de distribuição de estratégias (barras horizontais)
Term × strategy heatmap | Mapa de calor termo × estratégia
Thesis-ready summary tables (5 views) | Tabelas-resumo prontas para a monografia (5 visões)
Corpus health report | Relatório de saúde do corpus
Formatted Excel export with visual identity | Exportação em Excel formatado com identidade visual

---

## Quick start / Como usar

### EN

1. Open `notebooks/HanParal_Colab_v1_0_official.ipynb` in Google Colab
2. Run **cell 1** — install dependencies
3. Run **cell 2** — load all HanParal functions
4. Upload your corpus, or use `sample_data/sample_corpus.xlsx` to test
5. **Cell 4** — set `N_TARGETS` and choose a search mode
6. Run Phase 1, export the annotation file, annotate manually in Excel
7. Upload the annotated file and run Phase 2

### PT

1. Abra `notebooks/HanParal_Colab_v1_0_official.ipynb` no Google Colab
2. Execute a **célula 1** — instalar dependências
3. Execute a **célula 2** — carregar todas as funções do HanParal
4. Faça upload do seu corpus, ou use `sample_data/sample_corpus.xlsx` para testar
5. **Célula 4** — configure `N_TARGETS` e escolha o modo de busca
6. Execute a Fase 1, exporte o arquivo de anotação e anote manualmente no Excel
7. Faça upload do arquivo anotado e execute a Fase 2

### Configuring target languages / Configurando as línguas-alvo

```python
N_TARGETS = 3   # source + 3 target texts (default)
N_TARGETS = 2   # source + 2 target texts
N_TARGETS = 5   # source + 5 target texts
```

---

## Repository structure / Estrutura do repositório

```
HanParal-Concordancer/
├── notebooks/
│   └── HanParal_Colab_v1_0_official.ipynb   ← main workflow / fluxo principal
├── src/
│   └── hanparal_core.py                      ← all HanParal functions / todas as funções
├── sample_data/
│   ├── sample_corpus.xlsx                    ← invented test data / dados de teste fictícios
│   └── sample_annotated_occurrences.xlsx
├── README.md
├── CITATION.md
├── CITATION.cff
├── LICENSE
└── requirements.txt
```

---

## Installation / Instalação

```bash
pip install pandas openpyxl matplotlib numpy
```

**EN** — No installation needed for Google Colab — cell 1 handles it automatically.

**PT** — No Google Colab, nenhuma instalação manual é necessária — a célula 1 faz isso automaticamente.

---

## Data and copyright / Dados e direitos autorais

**EN** — The sample files contain entirely invented data. Do not upload or commit corpora derived from copyrighted literary texts to public repositories.

**PT** — Os arquivos de exemplo contêm dados completamente fictícios. Não faça upload nem commit de corpora derivados de textos literários protegidos por direitos autorais em repositórios públicos.

---

## Citation / Como citar

See [`CITATION.md`](CITATION.md) for ABNT and BibTeX formats.
Consulte [`CITATION.md`](CITATION.md) para os formatos ABNT e BibTeX.

---

## Background

**EN** — HanParal was developed as part of an undergraduate thesis in Translation Studies (Tradução) at the Universidade Federal da Paraíba (UFPB), Brazil. The original corpus involved culture-specific items across Portuguese, English and Spanish translations of a Korean literary source text, using the Aixelá (1996) framework. From v0.2 onwards, the tool was made fully language-neutral and can be used with any manually aligned multilingual parallel corpus.

**PT** — HanParal foi desenvolvido como parte de uma monografia de graduação em Tradução na Universidade Federal da Paraíba (UFPB), Brasil. O corpus original envolvia itens culturais específicos (ICEs) em traduções para o português, o inglês e o espanhol de um texto literário coreano de partida, com base no quadro teórico de Aixelá (1996). A partir da versão 0.2, a ferramenta foi tornada completamente independente de língua e pode ser utilizada com qualquer corpus paralelo multilíngue alinhado manualmente.

---

## License / Licença

MIT © 2026 Lívia Aguiar Correia Cavalcanti — see [`LICENSE`](LICENSE).

---

<div align="center">
<sub>built with 🌸 by 닐리 쌤</sub>
</div>
