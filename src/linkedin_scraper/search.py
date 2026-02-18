from bs4 import BeautifulSoup


def build_query(name: str, company: str, title_hint: str = "") -> str:
    parts = [name.strip(), company.strip(), title_hint.strip(), 'site:linkedin.com/in']
    return " ".join([p for p in parts if p])


def extract_linkedin_urls(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "linkedin.com/in/" in href:
            links.append(href)
    # preserve order, unique
    seen = set()
    out = []
    for u in links:
        if u not in seen:
            out.append(u)
            seen.add(u)
    return out
