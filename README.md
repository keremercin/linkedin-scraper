# 🔎 linkedin-scraper

[![CI](https://github.com/keremercin/linkedin-scraper/actions/workflows/ci.yml/badge.svg)](https://github.com/keremercin/linkedin-scraper/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![FastAPI](https://img.shields.io/badge/FastAPI-API-009688)

Professional LinkedIn URL enrichment pipeline for contact/company datasets.

## What it does
- builds targeted search queries (`name + company + title hint + site:linkedin.com/in`)
- extracts likely LinkedIn profile URLs from search results
- supports both batch enrichment and API usage
- returns hit-rate style outcome metrics

> Note: Always use responsibly and follow platform terms + privacy requirements.

---

## API

- `GET /health`
- `POST /v1/enrich`

Swagger: `http://localhost:8400/docs`

Example request:
```json
{
  "name": "Jane Doe",
  "company": "Acme",
  "title_hint": "CEO"
}
```

---

## Batch enrichment

Input file should include at least:
- `name`
- `company`

Run:
```bash
python scripts/run_enrichment.py \
  --input data/sample_input.csv \
  --output data/sample_output.csv \
  --title-hint "CEO"
```

---

## Quickstart

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[dev]

# API
uvicorn linkedin_scraper.api.main:app --reload --port 8400

# tests
pytest
```

---

## Project structure

```text
src/linkedin_scraper/
├─ api/main.py
├─ providers.py
├─ search.py
└─ pipeline.py
```

---

## Engineering quality

- unit tests for query/parsing/pipeline/API
- CI workflow (lint + test)
- modular architecture (provider/search/pipeline split)

---

## Docs
- `docs/CASE_STUDY.md`
