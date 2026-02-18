.PHONY: install run-api test lint enrich

install:
	python -m venv .venv && . .venv/bin/activate && pip install -e .[dev]

run-api:
	uvicorn linkedin_scraper.api.main:app --reload --port 8400

enrich:
	python scripts/run_enrichment.py --input data/sample_input.csv --output data/sample_output.csv --title-hint "CEO"

test:
	pytest

lint:
	ruff check src tests scripts
