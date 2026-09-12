"""Group 40 Kafue Town Council scraper and dataset builder.

The scraper intentionally uses only the official council website. It records
page-level text and link metadata, then adds a small curated fact layer for
stable, assignment-relevant fields. Output is pipe-delimited.
"""
from __future__ import annotations

import csv
import re
from datetime import date
from pathlib import Path
from urllib.parse import urljoin

from html.parser import HTMLParser
from urllib.request import Request, urlopen

BASE_URL = "https://www.kafuecouncil.gov.zm/"
OUT = Path(__file__).with_name("db-unza26-csc4792-kafue_council_information.csv")
TODAY = date.today().isoformat()

PAGES = {
    "district": "?page_id=2759",
    "quick_facts": "?page_id=2541",
    "finance": "?page_id=1409",
    "planning": "?page_id=1411",
    "engineering": "?page_id=1415",
    "administration": "?page_id=1421",
    "mandate": "?page_id=169",
    "services": "?page_id=831",
    "about_cdf": "?page_id=2102",
    "publications": "?page_id=195",
    "news": "?page_id=187",
}

POSTS = {
    "post_3023": "?p=3023",
    "post_2902": "?p=2902",
    "post_2473": "?p=2473",
    "post_1856": "?p=1856",
    "post_1758": "?p=1758",
    "post_1671": "?p=1671",
    "post_1446": "?p=1446",
    "post_1094": "?p=1094",
    "post_1059": "?p=1059",
}

FIELDS = ["record_type", "record_id", "title", "category", "description", "value", "unit", "date", "source_url", "source_name", "retrieved_at"]

def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()

class PageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.title_parts: list[str] = []
        self.text_parts: list[str] = []
        self.links: list[tuple[str, str]] = []
        self._in_title = False
        self._href = ""
        self._link_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_map = dict(attrs)
        if tag == "title":
            self._in_title = True
        if tag == "a":
            self._href = attrs_map.get("href") or ""
            self._link_parts = []

    def handle_endtag(self, tag: str) -> None:
        if tag == "title":
            self._in_title = False
        if tag == "a" and self._href:
            self.links.append((clean(" ".join(self._link_parts)), self._href))
            self._href, self._link_parts = "", []

    def handle_data(self, data: str) -> None:
        if self._in_title:
            self.title_parts.append(data)
        self.text_parts.append(data)
        if self._href:
            self._link_parts.append(data)

def fetch(url: str) -> tuple[str, PageParser]:
    request = Request(url, headers={"User-Agent": "CSC4792-Group40-research/1.0"})
    with urlopen(request, timeout=30) as response:
        final_url = response.geturl()
        html = response.read().decode("utf-8", errors="replace")
    parser = PageParser()
    parser.feed(html)
    return final_url, parser

def main() -> None:
    rows: list[dict[str, str]] = []
    for key, path in {**PAGES, **POSTS}.items():
        url = urljoin(BASE_URL, path)
        final_url, parser = fetch(url)
        title = clean(" ".join(parser.title_parts)) or key
        body_text = clean(" ".join(parser.text_parts))
        rows.append({"record_type": "page", "record_id": key, "title": title, "category": "official_web_page", "description": body_text[:5000], "value": "", "unit": "", "date": "", "source_url": final_url, "source_name": "Kafue Town Council", "retrieved_at": TODAY})
        for i, (label, raw_href) in enumerate(parser.links):
            href = urljoin(final_url, raw_href)
            if not href or href.endswith("#") or not label:
                continue
            if any(ext in href.lower() for ext in (".pdf", ".xlsx", ".xls", ".docx")):
                rows.append({"record_type": "resource", "record_id": f"{key}_resource_{i}", "title": label, "category": "linked_document", "description": "Official council resource linked from the council website.", "value": href, "unit": "url", "date": "", "source_url": final_url, "source_name": "Kafue Town Council", "retrieved_at": TODAY})

    # Stable assignment-relevant facts curated from the official pages.
    facts = [
        ("district_area", "District area", "district_profile", "Approximately 4471 square kilometres", "4471", "km2"),
        ("constituencies", "Constituencies", "district_profile", "Kafue has one constituency", "1", "constituency"),
        ("wards", "Wards", "district_profile", "Kafue Constituency has 18 wards", "18", "wards"),
        ("population", "District population", "district_profile", "Population reported on the Quick Facts page", "219574", "people"),
        ("province", "Province", "district_profile", "Kafue District is in Lusaka Province", "Lusaka", "province"),
        ("finance_budgeting", "Finance department function", "finance", "Coordinates annual council budgets and regular budget performance reviews", "", ""),
        ("finance_revenue", "Finance department objective", "finance", "Maximises the council revenue base and manages financial resources", "", ""),
        ("planning_cdf", "Planning department function", "planning", "Coordinates Constituency Development Fund activities", "", ""),
        ("planning_gis", "Planning department function", "planning", "Manages geographical information systems and development planning", "", ""),
        ("engineering_water", "Engineering service", "engineering", "Provides rural water and sanitation services in the district", "", ""),
        ("engineering_roads", "Engineering service", "engineering", "Constructs and maintains roads and drainages", "", ""),
        ("mandate_services", "Council mandate", "mandate", "Provides municipal services to spur socio-economic development", "", ""),
        ("cosp_motorbikes", "COSP motorbike handover", "official_news", "Nine motorbikes were handed over to Kafue Town Council for agriculture, health, community development, education, water and infrastructure implementers", "9", "motorbikes"),
        ("cosp_irrigation", "COSP irrigation target", "official_news", "Chiansi Outgrower Small Holder Support Project intends to develop an irrigation scheme covering 600 hectares", "600", "hectares"),
        ("cash_for_work_beneficiaries", "2026 Revised Cash for Work", "official_news", "National programme coverage and target reported by the council news post", "1500000", "beneficiaries"),
    ]
    for rid, title, cat, desc, value, unit in facts:
        source = BASE_URL + (PAGES.get("quick_facts") if rid in {"district_area", "constituencies", "wards", "population", "province"} else PAGES.get("finance") if rid.startswith("finance") else PAGES.get("planning") if rid.startswith("planning") else PAGES.get("engineering") if rid.startswith("engineering") else PAGES.get("mandate") if rid.startswith("mandate") else POSTS.get("post_2902") if rid.startswith("cosp") else POSTS.get("post_3023"))
        rows.append({"record_type": "fact", "record_id": rid, "title": title, "category": cat, "description": desc, "value": value, "unit": unit, "date": "", "source_url": urljoin(BASE_URL, source or ""), "source_name": "Kafue Town Council", "retrieved_at": TODAY})

    with OUT.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS, delimiter="|", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} records to {OUT}")

if __name__ == "__main__":
    main()
