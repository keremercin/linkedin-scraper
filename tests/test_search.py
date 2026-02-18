from linkedin_scraper.search import build_query, extract_linkedin_urls


def test_build_query() -> None:
    q = build_query("Jane Doe", "Acme", "CEO")
    assert "Jane Doe" in q
    assert "Acme" in q
    assert "site:linkedin.com/in" in q


def test_extract_linkedin_urls() -> None:
    html = '''
    <html><body>
      <a href="https://www.linkedin.com/in/jane-doe-123">Jane</a>
      <a href="https://example.com">X</a>
      <a href="https://www.linkedin.com/in/john-doe-456">John</a>
    </body></html>
    '''
    urls = extract_linkedin_urls(html)
    assert len(urls) == 2
    assert "jane-doe" in urls[0]
