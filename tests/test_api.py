from fastapi.testclient import TestClient

from linkedin_scraper.api.main import app, provider


class DummyProvider:
    def search_linkedin(self, query: str, timeout: int = 20):
        return ["https://www.linkedin.com/in/demo-profile"]


def test_health() -> None:
    client = TestClient(app)
    r = client.get("/health")
    assert r.status_code == 200


def test_enrich(monkeypatch) -> None:
    monkeypatch.setattr(type(provider), "search_linkedin", lambda self, query, timeout=20: ["https://www.linkedin.com/in/demo-profile"])
    client = TestClient(app)
    r = client.post("/v1/enrich", json={"name": "Jane", "company": "Acme", "title_hint": "CEO"})
    assert r.status_code == 200
    body = r.json()
    assert body["linkedin_url"].startswith("https://www.linkedin.com/in/")
