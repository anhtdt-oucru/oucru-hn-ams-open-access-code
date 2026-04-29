## Summary

<!-- What does this PR do? Link to the related issue or project if applicable. -->

## Type of change

- [ ] `feat` — new script or functionality
- [ ] `fix` — correcting a bug or wrong output
- [ ] `docs` — documentation and comments only
- [ ] `refactor` — restructuring code without changing behaviour
- [ ] `chore` — renaming, moving files, updating configs
- [ ] `style` — formatting or naming, no logic change

## Data safety checklist ⚠️

> This repository is public. Every item below must be checked before merging.

- [ ] **No data files committed** — no `.xlsx`, `.xls`, `.csv`, `.tsv`, `.parquet`, `.accdb`, `.mdb`, `.sav`, `.dta`, `.rds`, `.rda` files in this PR
- [ ] **No `.Rhistory` or `.RData` files committed**
- [ ] **No hardcoded paths** pointing to internal servers, network drives, or patient-linked directories (e.g. `//hospital-server/...`, `C:/Users/staff/...`)
- [ ] **No credentials or tokens** in any script (database passwords, API keys, etc.)
- [ ] If a `reference/` file is included: **confirmed it contains no patient identifiers**, admission numbers, or internal site mappings

## Code checklist

- [ ] Code runs end-to-end without errors on a clean R session
- [ ] `renv.lock` updated if new packages were added (`renv::snapshot()`)
- [ ] README updated if the script structure or usage changed
- [ ] Follows the two-layer convention: data-source logic → `core/`, study analysis → `projects/`
