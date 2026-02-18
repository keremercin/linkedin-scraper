import argparse

from linkedin_scraper.pipeline import enrich_linkedin_urls


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--name-col", default="name")
    parser.add_argument("--company-col", default="company")
    parser.add_argument("--title-hint", default="")
    args = parser.parse_args()

    out = enrich_linkedin_urls(
        input_path=args.input,
        output_path=args.output,
        name_col=args.name_col,
        company_col=args.company_col,
        title_hint=args.title_hint,
    )
    print(out)


if __name__ == "__main__":
    main()
