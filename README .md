# Group 40 - Kafue Town Council dataset

This package is a reproducible first release for CSC 4792 Group 40.

## Contents

- `db-unza26-csc4792-kafue_council_information.csv` - pipe-delimited curated dataset.
- `scrape_kafue.py` - scraper and normalisation pipeline for the official council website.
- `group40_kafue_dataset.ipynb` - documented Jupyter Notebook version of the workflow.

## Scope

The dataset covers Kafue district profile facts, council mandate and departmental functions, CDF and planning resources, publication metadata, and official news items. Each record preserves the source URL and retrieval date. The site publishes several linked PDF, XLSX, XLS, and DOCX resources; these are represented as resource metadata rows so the source files can be downloaded and parsed in a subsequent enrichment pass.

## Format

CSV files use `|` as the separator, as required by the assignment. The naming convention follows `db-unza26-csc4792-[description].csv`.

## Source

Official website: https://www.kafuecouncil.gov.zm/

Retrieved: 2026-09-11 (Africa/Johannesburg)
