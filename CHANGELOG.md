# Changelog

All notable changes to this repository are documented here.  
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [Unreleased]

---

## [2026-04] — 2026-04-29

### Fixed
- Removed tracked root `.Rhistory` file (`git rm --cached .Rhistory`)

### Changed
- `.gitignore` — added root `.Rhistory` explicitly; added data file extensions
  (`.xlsx`, `.xls`, `.csv`, `.tsv`, `.parquet`, `.accdb`, `.mdb`, `.rds`, etc.)