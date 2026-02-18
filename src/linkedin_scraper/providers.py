import requests

from linkedin_scraper.search import extract_linkedin_urls


class DuckDuckGoProvider:
    """Simple HTML search provider for demo/portfolio purposes.

    Note: Search engines may change HTML structure/rate limits.
    """

    endpoint = "https://duckduckgo.com/html/"

    def search_linkedin(self, query: str, timeout: int = 20) -> list[str]:
        r = requests.get(self.endpoint, params={"q": query}, timeout=timeout)
        r.raise_for_status()
        return extract_linkedin_urls(r.text)
