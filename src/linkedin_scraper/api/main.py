from fastapi import FastAPI
from pydantic import BaseModel, Field

from linkedin_scraper.providers import DuckDuckGoProvider
from linkedin_scraper.search import build_query

app = FastAPI(title="LinkedIn Enrichment API", version="0.2.0")
provider = DuckDuckGoProvider()


class EnrichRequest(BaseModel):
    name: str = Field(min_length=1)
    company: str = Field(min_length=1)
    title_hint: str = ""


@app.get("/health")
def health() -> dict:
    return {"status": "ok", "service": "linkedin-enrichment"}


@app.post("/v1/enrich")
def enrich(req: EnrichRequest) -> dict:
    query = build_query(req.name, req.company, req.title_hint)
    urls = provider.search_linkedin(query)
    return {
        "query": query,
        "linkedin_url": urls[0] if urls else None,
        "candidates": urls[:5],
    }
