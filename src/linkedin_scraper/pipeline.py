from pathlib import Path

import pandas as pd

from linkedin_scraper.providers import DuckDuckGoProvider
from linkedin_scraper.search import build_query


def load_table(path: str | Path) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(p)
    return pd.read_csv(p)


def save_table(df: pd.DataFrame, path: str | Path) -> None:
    p = Path(path)
    if p.suffix.lower() in {".xlsx", ".xls"}:
        df.to_excel(p, index=False)
    else:
        df.to_csv(p, index=False)


def enrich_linkedin_urls(
    input_path: str,
    output_path: str,
    name_col: str = "name",
    company_col: str = "company",
    title_hint: str = "",
) -> dict:
    df = load_table(input_path)
    provider = DuckDuckGoProvider()

    out_urls = []
    for _, row in df.iterrows():
        query = build_query(str(row.get(name_col, "")), str(row.get(company_col, "")), title_hint)
        urls = provider.search_linkedin(query)
        out_urls.append(urls[0] if urls else None)

    df["linkedin_url"] = out_urls
    save_table(df, output_path)

    hit_rate = float(pd.Series(out_urls).notna().mean()) if len(out_urls) else 0.0
    return {
        "rows": len(df),
        "matched": int(pd.Series(out_urls).notna().sum()),
        "hit_rate": round(hit_rate, 4),
        "output_path": output_path,
    }
