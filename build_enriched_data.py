"""Build the public, row-level Group 40 data release.

The downloaded council workbooks contain personal fields (NRC numbers,
guardian names and phone numbers).  This script deliberately excludes those
fields and keeps only public project, grant, loan, performance, finance and
document metadata that can be used for the assignment.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from pypdf import PdfReader
except ImportError:  # pragma: no cover
    PdfReader = None


SCRIPT_DIR = Path(__file__).resolve().parent
if SCRIPT_DIR.name == "work":
    ROOT = SCRIPT_DIR.parent
    SOURCE_DIR = ROOT / "work" / "source_docs"
    OUT_DIR = ROOT / "outputs" / "group40_kafue"
else:
    # In the submitted repository the script is kept beside the CSV files.
    ROOT = SCRIPT_DIR
    SOURCE_DIR = ROOT / "source_docs"
    OUT_DIR = ROOT
RETRIEVED_AT = "2026-09-11"
COLLECTION_PAGE = "https://www.kafuecouncil.gov.zm/?page_id=195"
BASE_URL = "https://www.kafuecouncil.gov.zm/"

CDF_FILES = {
    "kafue_2022_cdf_submission.xlsx": "2022",
    "kafue_2023_cdf_submission.xlsx": "2023",
    "kafue_2024_cdf_submission.xlsx": "2024",
    "kafue_2024_community_projects.xlsx": "2024",
}

PERFORMANCE_FILE = "kafue_2024_2025_performance.xlsx"
EXPENDITURE_FILE = "kafue_2024_cdf_expenditure.xlsx"

PDF_FILES = {
    "kafue_2025_approved_cdf_projects.pdf": "Approved CDF community projects 2025",
    "kafue_2025_cdf_all_applications.pdf": "CDF applications 2025",
    "kafue_2025_budget.pdf": "Kafue Town Council budget 2025",
    "kafue_2022_financial_statements.pdf": "Financial statements 2022",
    "kafue_2023_financial_statements.pdf": "Financial statements 2023",
    "kafue_2024_financial_statements.pdf": "Financial statements 2024",
    "kafue_2026_budget_performance.pdf": "Budget performance report June 2026",
    "kafue_idp_2024_2034.pdf": "Integrated Development Plan 2024-2034",
    "kafue_2025_capital_projects.pdf": "Capital projects 2025",
    "kafue_2025_council_minutes.pdf": "Council minutes 30 May 2025",
}

CDF_COLUMNS = [
    "record_type", "record_id", "source_year", "source_file", "source_sheet",
    "row_number", "category", "entity_name", "description", "province",
    "district", "constituency", "ward", "zone", "sector", "project_site",
    "status", "amount_requested", "amount_recommended", "amount_approved",
    "amount_disbursed", "annual_fees", "termly_fees", "date_approved",
    "date_disbursed", "source_url", "source_name", "retrieved_at",
]

PERFORMANCE_COLUMNS = [
    "record_type", "record_id", "source_year", "source_file", "source_sheet",
    "row_number", "programme", "key_output", "key_indicator", "annual_target",
    "achieved_output", "variance", "status", "comment", "source_url",
    "source_name", "retrieved_at",
]

FINANCE_COLUMNS = [
    "record_type", "record_id", "reporting_month", "source_file", "source_sheet",
    "row_number", "metric", "monthly_amount", "cumulative_amount", "notes",
    "source_url", "source_name", "retrieved_at",
]

DOCUMENT_COLUMNS = [
    "record_type", "record_id", "source_file", "document_type", "page_number",
    "page_text", "source_url", "source_name", "retrieved_at",
]


def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def safe_number(value) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = text.replace(",", "").replace("ZMW", "").replace("K", "")
    return text.strip()


def normal_header(value) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def find_col(headers: list[str], include: tuple[str, ...], exclude: tuple[str, ...] = ()) -> str | None:
    for header in headers:
        h = normal_header(header)
        if all(term in h for term in include) and not any(term in h for term in exclude):
            return header
    return None


def first_col(headers: list[str], candidates: list[tuple[tuple[str, ...], tuple[str, ...]]]) -> str | None:
    for include, exclude in candidates:
        found = find_col(headers, include, exclude)
        if found:
            return found
    return None


def row_value(row: pd.Series, col: str | None) -> str:
    return clean_text(row.get(col, "")) if col else ""


def source_label(filename: str) -> str:
    return filename.replace("_", " ").replace(".xlsx", "").replace(".pdf", "").strip()


def category_for_sheet(sheet: str) -> str:
    s = normal_header(sheet)
    if "community" in s or "project" in s:
        return "community_project"
    if "empowerment" in s or "grant" in s:
        return "empowerment_grant"
    if "loan" in s:
        return "cdf_loan"
    if "skill" in s:
        return "skills_development_summary"
    if "secondary" in s or "boarding" in s:
        return "secondary_bursary_summary"
    return "cdf_record"


def public_row(sheet: str, row: pd.Series, filename: str, year: str, excel_row: int, index: int) -> dict | None:
    category = category_for_sheet(sheet)
    headers = [str(c) for c in row.index]

    # Individual bursary and skills rows contain personal information.  Keep
    # useful, non-identifying annual summaries for these sheets instead.
    if category.endswith("_summary"):
        return None

    name_col = first_col(headers, [
        (("project", "name"), ()),
        (("name", "of", "group"), ()),
        (("group", "name"), ()),
        (("business", "name"), ()),
        (("company", "name"), ()),
    ])
    description_col = first_col(headers, [
        (("project", "description"), ()),
        (("type", "of", "venture"), ()),
        (("type", "of", "business"), ()),
    ])
    sector_col = first_col(headers, [(("sector",), ("financial" ,))])
    project_site_col = first_col(headers, [(("project", "site"), ()), (("location",), ())])
    province_col = first_col(headers, [(("province",), ())])
    district_col = first_col(headers, [(("district",), ())])
    constituency_col = first_col(headers, [(("constituency",), ())])
    ward_col = first_col(headers, [(("ward",), ())])
    zone_col = first_col(headers, [(("zone",), ())])
    status_col = first_col(headers, [(("status",), ()), (("repayment",), ())])
    requested_col = first_col(headers, [(("amount", "requested"), ()), (("requested",), ())])
    recommended_col = first_col(headers, [(("amount", "recommended"), ()), (("recommended",), ())])
    approved_col = first_col(headers, [(("approved", "amount"), ()), (("amount", "approved"), ())])
    disbursed_col = first_col(headers, [(("disbursed",), ()), (("bank", "amount"), ())])
    approved_date_col = first_col(headers, [(("date", "of", "approval"), ()), (("approval", "date"), ())])
    disbursed_date_col = first_col(headers, [(("date", "of", "disbursement"), ()), (("disbursement", "date"), ())])
    annual_fees_col = first_col(headers, [(("annual", "fees"), ())])
    termly_fees_col = first_col(headers, [(("termly", "fees"), ())])

    values = {
        "record_type": "cdf_record",
        "record_id": f"{year}_{Path(filename).stem}_{normal_header(sheet).replace(' ', '_')}_{index:04d}",
        "source_year": year,
        "source_file": filename,
        "source_sheet": clean_text(sheet),
        "row_number": str(excel_row),
        "category": category,
        "entity_name": row_value(row, name_col),
        "description": row_value(row, description_col),
        "province": row_value(row, province_col),
        "district": row_value(row, district_col),
        "constituency": row_value(row, constituency_col),
        "ward": row_value(row, ward_col),
        "zone": row_value(row, zone_col),
        "sector": row_value(row, sector_col),
        "project_site": row_value(row, project_site_col),
        "status": row_value(row, status_col),
        "amount_requested": safe_number(row.get(requested_col, "")) if requested_col else "",
        "amount_recommended": safe_number(row.get(recommended_col, "")) if recommended_col else "",
        "amount_approved": safe_number(row.get(approved_col, "")) if approved_col else "",
        "amount_disbursed": safe_number(row.get(disbursed_col, "")) if disbursed_col else "",
        "annual_fees": safe_number(row.get(annual_fees_col, "")) if annual_fees_col else "",
        "termly_fees": safe_number(row.get(termly_fees_col, "")) if termly_fees_col else "",
        "date_approved": row_value(row, approved_date_col),
        "date_disbursed": row_value(row, disbursed_date_col),
        "source_url": COLLECTION_PAGE,
        "source_name": source_label(filename),
        "retrieved_at": RETRIEVED_AT,
    }
    # Skip genuinely empty rows and rows that are only an index.
    if not any(values[key] for key in ("entity_name", "description", "ward", "sector", "amount_approved", "amount_requested")):
        return None
    return values


def read_public_cdf_records() -> tuple[pd.DataFrame, list[dict]]:
    records: list[dict] = []
    summaries: list[dict] = []
    for filename, year in CDF_FILES.items():
        path = SOURCE_DIR / filename
        if not path.exists():
            continue
        book = pd.ExcelFile(path)
        for sheet in book.sheet_names:
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
            # The council templates place the field names on the fifth row.
            header_index = 4 if len(raw) > 4 else 0
            headers = [clean_text(v) or f"column_{i}" for i, v in enumerate(raw.iloc[header_index].tolist())]
            data = raw.iloc[header_index + 1:].copy()
            data.columns = headers
            data = data.dropna(how="all")
            category = category_for_sheet(sheet)
            if category.endswith("_summary"):
                # No names, NRCs, addresses or phone numbers are exported.
                summaries.append({
                    "record_type": "cdf_summary",
                    "record_id": f"{year}_{Path(filename).stem}_{normal_header(sheet).replace(' ', '_')}_summary",
                    "source_year": year,
                    "source_file": filename,
                    "source_sheet": clean_text(sheet),
                    "row_number": "",
                    "category": category,
                    "entity_name": "",
                    "description": f"{len(data)} source rows retained as a non-identifying count; individual beneficiary fields excluded.",
                    "province": "",
                    "district": "Kafue",
                    "constituency": "Kafue",
                    "ward": "",
                    "zone": "",
                    "sector": "",
                    "project_site": "",
                    "status": "",
                    "amount_requested": "",
                    "amount_recommended": "",
                    "amount_approved": "",
                    "amount_disbursed": "",
                    "annual_fees": "",
                    "termly_fees": "",
                    "date_approved": "",
                    "date_disbursed": "",
                    "source_url": COLLECTION_PAGE,
                    "source_name": source_label(filename),
                    "retrieved_at": RETRIEVED_AT,
                })
                continue
            for idx, (_, row) in enumerate(data.iterrows(), start=1):
                record = public_row(sheet, row, filename, year, header_index + 2 + idx, idx)
                if record:
                    records.append(record)
    all_records = summaries + records
    return pd.DataFrame(all_records, columns=CDF_COLUMNS), summaries


def find_table_header(raw: pd.DataFrame, max_rows: int = 100) -> int | None:
    for i in range(min(max_rows, len(raw))):
        text = " ".join(normal_header(v) for v in raw.iloc[i].tolist() if clean_text(v))
        hits = sum(token in text for token in ("programme", "key output", "annual target", "variance", "project", "no"))
        if hits >= 2:
            return i
    return None


def performance_col(headers: list[str], *terms: str) -> str | None:
    for h in headers:
        nh = normal_header(h)
        if all(term in nh for term in terms):
            return h
    return None


def read_performance_records() -> pd.DataFrame:
    path = SOURCE_DIR / PERFORMANCE_FILE
    records: list[dict] = []
    if not path.exists():
        return pd.DataFrame(columns=PERFORMANCE_COLUMNS)
    book = pd.ExcelFile(path)
    for sheet in book.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        header_index = find_table_header(raw)
        if header_index is None:
            continue
        headers = [clean_text(v) or f"column_{i}" for i, v in enumerate(raw.iloc[header_index].tolist())]
        data = raw.iloc[header_index + 1:].copy()
        data.columns = headers
        data = data.dropna(how="all")
        programme = performance_col(headers, "programme")
        output = performance_col(headers, "key", "output")
        indicator = performance_col(headers, "key", "indicator")
        target = performance_col(headers, "annual", "target")
        achieved = performance_col(headers, "achieved") or performance_col(headers, "acheived")
        variance = performance_col(headers, "variance")
        comment = performance_col(headers, "comment")
        status = performance_col(headers, "status")
        for idx, (_, row) in enumerate(data.iterrows(), start=1):
            vals = {
                "record_type": "performance_record",
                "record_id": f"performance_{normal_header(sheet).replace(' ', '_')}_{idx:04d}",
                "source_year": "2024-2025",
                "source_file": PERFORMANCE_FILE,
                "source_sheet": clean_text(sheet),
                "row_number": str(header_index + 2 + idx),
                "programme": row_value(row, programme),
                "key_output": row_value(row, output),
                "key_indicator": row_value(row, indicator),
                "annual_target": row_value(row, target),
                "achieved_output": row_value(row, achieved),
                "variance": row_value(row, variance),
                "status": row_value(row, status),
                "comment": row_value(row, comment),
                "source_url": COLLECTION_PAGE,
                "source_name": source_label(PERFORMANCE_FILE),
                "retrieved_at": RETRIEVED_AT,
            }
            if any(vals[k] for k in ("programme", "key_output", "key_indicator", "status")):
                records.append(vals)
    return pd.DataFrame(records, columns=PERFORMANCE_COLUMNS)


def numeric_tokens(values: list[str]) -> list[str]:
    out = []
    for value in values:
        text = clean_text(value).replace(",", "")
        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
            out.append(text)
    return out


def read_finance_records() -> pd.DataFrame:
    path = SOURCE_DIR / EXPENDITURE_FILE
    records: list[dict] = []
    if not path.exists():
        return pd.DataFrame(columns=FINANCE_COLUMNS)
    book = pd.ExcelFile(path)
    metric_words = ("balance", "receipt", "payment", "expenditure", "transfer", "bank statement", "cash book")
    for sheet in book.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        for idx, (_, row) in enumerate(raw.iterrows(), start=1):
            values = [clean_text(v) for v in row.tolist()]
            text_values = [v for v in values if v]
            if not text_values:
                continue
            label = next((v for v in text_values if any(w in v.lower() for w in metric_words)), "")
            if not label:
                continue
            nums = numeric_tokens(values)
            if not nums:
                continue
            records.append({
                "record_type": "finance_metric",
                "record_id": f"finance_{normal_header(sheet).replace(' ', '_')}_{idx:04d}",
                "reporting_month": clean_text(sheet),
                "source_file": EXPENDITURE_FILE,
                "source_sheet": clean_text(sheet),
                "row_number": str(idx),
                "metric": label,
                "monthly_amount": nums[0] if nums else "",
                "cumulative_amount": nums[1] if len(nums) > 1 else "",
                "notes": "Values are preserved as reported; currency and column meaning should be checked against the source sheet.",
                "source_url": COLLECTION_PAGE,
                "source_name": source_label(EXPENDITURE_FILE),
                "retrieved_at": RETRIEVED_AT,
            })
    return pd.DataFrame(records, columns=FINANCE_COLUMNS)


def read_document_pages() -> pd.DataFrame:
    records: list[dict] = []
    for filename, title in PDF_FILES.items():
        path = SOURCE_DIR / filename
        if not path.exists() or PdfReader is None:
            continue
        reader = PdfReader(str(path))
        for page_number, page in enumerate(reader.pages, start=1):
            text = clean_text(page.extract_text() or "")
            # Keep the page-level text for traceability, bounded so a single
            # page cannot dominate the CSV. Pipes are escaped for the format.
            text = text.replace("|", "/")[:8000]
            records.append({
                "record_type": "document_page",
                "record_id": f"{Path(filename).stem}_page_{page_number:04d}",
                "source_file": filename,
                "document_type": title,
                "page_number": str(page_number),
                "page_text": text,
                "source_url": COLLECTION_PAGE,
                "source_name": title,
                "retrieved_at": RETRIEVED_AT,
            })
    return pd.DataFrame(records, columns=DOCUMENT_COLUMNS)


def write_pipe(df: pd.DataFrame, filename: str) -> None:
    path = OUT_DIR / filename
    df.fillna("").astype(str).to_csv(path, sep="|", index=False, encoding="utf-8")
    print(f"{filename}: {len(df):,} rows, {len(df.columns)} columns")


def write_manifest() -> None:
    rows = []
    for filename in sorted(SOURCE_DIR.glob("*")):
        if filename.is_file():
            rows.append({
                "source_file": filename.name,
                "file_type": filename.suffix.lower().lstrip("."),
                "file_size_bytes": str(filename.stat().st_size),
                "source_url": COLLECTION_PAGE,
                "source_page": COLLECTION_PAGE,
                "retrieved_at": RETRIEVED_AT,
                "notes": "Downloaded from the official Kafue Town Council publications page; raw source files are kept outside the submission CSV release.",
            })
    write_pipe(pd.DataFrame(rows), "db-unza26-csc4792-source_manifest.csv")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cdf, summaries = read_public_cdf_records()
    performance = read_performance_records()
    finance = read_finance_records()
    pages = read_document_pages()
    write_pipe(cdf, "db-unza26-csc4792-kafue_cdf_records.csv")
    write_pipe(performance, "db-unza26-csc4792-kafue_performance_records.csv")
    write_pipe(finance, "db-unza26-csc4792-kafue_finance_records.csv")
    write_pipe(pages, "db-unza26-csc4792-kafue_document_pages.csv")
    write_manifest()
    print(f"Non-identifying CDF summaries: {len(summaries)}")


if __name__ == "__main__":
    main()
