# ✦ Generative UI Builder

> A Streamlit application that accepts multimodal inputs — text, CSV, images, PDFs, Word docs, and Python files — and uses an LLM to dynamically generate a structured, interactive dashboard in real time.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Running the App](#running-the-app)
- [How It Works](#how-it-works)
- [Supported Input Types](#supported-input-types)
- [Supported Chart & Component Types](#supported-chart--component-types)
- [Smart Chart Selection](#smart-chart-selection)
- [Project Structure](#project-structure)
- [Git Setup](#git-setup)
- [Notes & Limitations](#notes--limitations)

---

## Overview

Generative UI Builder lets you describe a dashboard in plain English — or simply upload a data file — and the LLM acts as a **UI architect**, returning a strict JSON layout that the app instantly renders as interactive charts, metric cards, tables, and text blocks.

All charts are rendered with **Plotly**, giving you real pie charts, scatter plots with trendlines, histograms, and area charts — all styled to match the dark editorial theme.

---

## Features

- **Multimodal input** — text prompts, CSV, images (PNG/JPG/WEBP), plain text, Python, PDF, and Word (`.docx`)
- **Generative UI** — the LLM outputs structured JSON; the app renders it live
- **Smart chart selection** — data is analyzed for time series, correlations, distributions, and cardinality; the LLM is guided to pick the most appropriate chart type automatically
- **Plotly-powered charts** — real pie/donut charts, scatter with trendline, histogram, area, bar, and line
- **Streaming responses** — live JSON preview while the model is generating
- **Python file tooling** — upload a `.py` file to auto-fill code gaps or generate full Markdown documentation
- **Dark editorial theme** — DM Serif / DM Sans / DM Mono typographic system

---

## Requirements

- Python 3.9 or higher
- An **OpenAI API key** with access to `gpt-5.4-mini`

### Python Dependencies

All dependencies are listed in `requirements.txt`. Install them in one command:

```bash
pip install -r requirements.txt
```

| Package | Purpose |
|---|---|
| `streamlit` | Web app framework |
| `openai` | OpenAI API client |
| `pandas` | Data loading and analysis |
| `plotly` | Interactive chart rendering |
| `pillow` | Image handling |
| `python-dotenv` | Loads `.env` API key at startup |
| `pydantic` | JSON schema validation |
| `pypdf` | PDF text extraction |
| `python-docx` | Word document parsing |
| `statsmodels` | OLS trendline in scatter charts *(optional)* |

---

## Installation

**1. Clone or download the project**

```bash
git clone <your-repo-url>
cd generative-ui-builder
```

**2. Create and activate a virtual environment (recommended)**

```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows
```

**3. Install dependencies**

```bash
pip install -r requirements.txt
```

---

## Configuration

Create a `.env` file in the project root:

```
OPENAI_API_KEY=sk-...your-key-here...
```

Alternatively, export it as an environment variable before running:

```bash
export OPENAI_API_KEY=sk-...your-key-here...   # macOS / Linux
set OPENAI_API_KEY=sk-...your-key-here...      # Windows CMD
```

The app reads the key with `python-dotenv` at startup. If no key is found, a warning is shown and generation is blocked until one is set.

---

## Running the App

```bash
streamlit run UI_app.py
```

The app opens at `http://localhost:8501` by default.

---

## How It Works

```
User input (text / file)
        │
        ▼
  build_messages()
  ┌─────────────────────────────────────┐
  │  System prompt (UI architect rules) │
  │  + data summary with chart hints    │
  │  + user prompt / file content       │
  └─────────────────────────────────────┘
        │
        ▼  (streaming)
  OpenAI gpt-5.4-mini
        │
        ▼
  Raw JSON string
        │
        ▼
  parse_ui_response()   ← strips fences, validates with Pydantic
        │
        ▼
  UIResponse (title, layout, components[])
        │
        ▼
  render_component()    ← Plotly charts, st.metric, st.dataframe, st.markdown
        │
        ▼
  Interactive Streamlit dashboard
```

---

## Supported Input Types

| Input | What gets sent to the model |
|---|---|
| **Text prompt** | Raw prompt; model invents realistic sample data |
| **CSV file** | Statistical summary + column types + chart-suitability hints |
| **Image** (PNG/JPG/WEBP) | Base64-encoded image via multimodal API |
| **Plain text** (`.txt`) | First 4,000 characters of the file |
| **Python** (`.py`) | Structural summary (functions, classes, imports) + source excerpt |
| **PDF** | Extracted text (up to 6,000 chars) + page count + metadata |
| **Word** (`.docx`) | Paragraph text + headings + table count (up to 6,000 chars) |

---

## Supported Chart & Component Types

The LLM can emit any of these component types in its JSON output:

| Type | Description | Rendered with |
|---|---|---|
| `text` | Markdown text block | `st.markdown` |
| `metric` | KPI card with optional delta | `st.metric` |
| `table` | Tabular data | `st.dataframe` |
| `bar_chart` | Grouped/horizontal bar chart | Plotly Express |
| `line_chart` | Line chart with markers | Plotly Express |
| `area_chart` | Filled area chart | Plotly Express |
| `pie_chart` | Pie or donut chart (set `hole: 0.4`) | Plotly Express |
| `scatter_chart` | Scatter plot with optional trendline & color grouping | Plotly Express |
| `histogram` | Distribution histogram with configurable bins | Plotly Express |
| `image` | Image from a public URL | `st.image` |

---

## Smart Chart Selection

When a CSV is uploaded, the app analyzes the data and sends **chart-suitability hints** to the model alongside the data summary. The model then follows these rules:

| Data pattern detected | Chart chosen |
|---|---|
| Datetime column + numeric | `line_chart` or `area_chart` |
| Two or more numeric columns | `scatter_chart` (correlation) |
| Categorical ≤ 6 unique values | `pie_chart` (donut style) |
| Categorical 7–12 unique values | `bar_chart` (horizontal if long labels) |
| Categorical > 12 unique values | `bar_chart` (top-N) + `table` |
| Continuous numeric distribution | `histogram` |
| Mixed / no time axis | `bar_chart` (grouped) + `scatter_chart` |

This ensures the **most informative** chart always appears automatically, without any manual configuration.

---

## Project Structure

```
generative-ui-builder/
├── UI_app.py          # Main application (single file)
├── requirements.txt   # All Python dependencies
├── .env               # API key — never committed (listed in .gitignore)
├── .env.example       # Safe template to commit; shows required keys without values
├── .gitignore         # Excludes secrets, venvs, caches, data files, and IDE folders
└── README.md          # This file
```

---

## Git Setup

A `.gitignore` is included to keep secrets, generated files, and local tooling out of version control.

**What is ignored:**

| Category | Examples |
|---|---|
| Secrets | `.env`, `*.pem`, `*.key` |
| Virtual environments | `venv/`, `.venv/`, `env/` |
| Python cache | `__pycache__/`, `*.pyc`, `*.pyo` |
| Build & dist | `dist/`, `build/`, `*.egg-info/` |
| Streamlit secrets | `.streamlit/secrets.toml` |
| Data & uploads | `*.csv`, `*.xlsx`, `uploads/` |
| IDE settings | `.vscode/`, `.idea/`, `.DS_Store` |
| Logs & temp | `*.log`, `tmp/`, `temp/` |

**Recommended: commit an `.env.example`** so collaborators know which keys to set:

```bash
# .env.example  ← safe to commit
OPENAI_API_KEY=your-api-key-here
```

**First-time git setup:**

```bash
git init
git add UI_app.py requirements.txt README.md .gitignore .env.example
git commit -m "Initial commit"
```

---

## Notes & Limitations

- **Model name** — hardcoded to `gpt-5.4-mini` in `MODEL_NAME`. Change this constant to swap models.
- **Data privacy** — CSV contents and file text are sent to the OpenAI API. Do not upload sensitive or personal data.
- **Invented data** — for pure text prompts (no file), the model generates plausible sample data. This is intentional for demo/prototyping purposes.
- **Image URLs** — the `image` component type requires a publicly accessible URL. Base64 images from uploads are not re-rendered in the generated layout.
- **Pie chart** — set `"hole": 0.4` in the JSON content for a donut chart; `0.0` renders a full pie.
- **Plotly trendlines** — the `scatter_chart` type adds an OLS trendline automatically when no color grouping is present. This requires `statsmodels` to be installed (`pip install statsmodels`). If not installed, the chart renders without the trendline.
- **Temperature slider** — available in the sidebar (0.0–1.0). Lower values produce more consistent, deterministic JSON; higher values produce more creative layouts.