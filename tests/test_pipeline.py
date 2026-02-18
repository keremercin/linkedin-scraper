from pathlib import Path

import pandas as pd

from linkedin_scraper import pipeline


class DummyProvider:
    def search_linkedin(self, query: str, timeout: int = 20):
        if "Acme" in query:
            return ["https://www.linkedin.com/in/acme-ceo"]
        return []


def test_enrich_pipeline(monkeypatch, tmp_path: Path) -> None:
    inp = tmp_path / "input.csv"
    out = tmp_path / "output.csv"
    pd.DataFrame([
        {"name": "Jane", "company": "Acme"},
        {"name": "Bob", "company": "Unknown"},
    ]).to_csv(inp, index=False)

    monkeypatch.setattr(pipeline, "DuckDuckGoProvider", DummyProvider)

    stats = pipeline.enrich_linkedin_urls(str(inp), str(out))
    assert stats["rows"] == 2
    assert out.exists()

    df = pd.read_csv(out)
    assert "linkedin_url" in df.columns
