# OUCRU HN AMS — Open Access Code

Open-source code repository for the Antimicrobial Stewardship (AMS) team at the Oxford University Clinical Research Unit, Hanoi (OUCRU HN).

This repository centralises reusable data pipelines and project-specific analysis code that would otherwise be scattered across local drives. It serves two audiences: **team members** looking for shared cleaning tools, and **external researchers** seeking code associated with published studies.

---

## Repository structure

```
.
├── core/                        # Reusable pipelines by data source
│   └── <source_name>/
│       ├── R/
│       │   ├── load.R           # Data loading
│       │   ├── clean.R          # Cleaning and standardisation
│       │   └── map.R            # Reference mapping
│       └── reference/           # Lookup and mapping files for this source
│
├── projects/                    # Project-specific analysis code
│   └── <project_id_name>/
│       ├── scripts/             # Analysis scripts
│       ├── renv.lock            # Package snapshot (reproducibility)
│       └── README.md            # Project description and publication link
│
└── README.md
```

### `core/`

Contains generalized, reusable scripts for loading, cleaning, and mapping each data source. Code here is **data-source-centric** — it is shared across projects and maintained independently of any single study.

### `projects/`

Contains analysis code tied to a specific study or publication. Each subfolder is a self-contained R project with its own `renv.lock` for reproducibility. Code here is **project-centric** — it sources functions from `core/` and applies them to a specific research question.

---

## Contributing

This repository follows a two-layer structure. When adding new code, please place it in the correct layer:

- **Data-source logic** (loading, cleaning, mapping) → `core/<source_name>/`
- **Study/project analysis** → `projects/<project_id_name>/`

**Commit message convention:** 

| Type | Use for |
|---|---|
| `feat` | New script or functionality |
| `fix` | Correcting a bug or wrong output |
| `docs` | Documentation and comments only |
| `refactor` | Restructuring code without changing behaviour |
| `chore` | Maintenance — renaming, moving files, updating configs |
| `style` | Formatting or naming — no logic change |

New contributors should be added to `CODEOWNERS` by the repository maintainer.

---

## Maintainers

| GitHub handle | Role | Projects |
|---|---|---|
| [@anhtdt-oucru](https://github.com/anhtdt-oucru) | Repo admin | All (default owner) |
| [@yendh-oucru](https://github.com/yendh-oucru) | Code owner | 60HN – Antibiogram |
| [@tranglnh-oucru](https://github.com/tranglnh-oucru) | Code owner | 60HN – Individual Report |

---

## About OUCRU HN AMS

The AMS team at OUCRU Hanoi conducts research on antimicrobial resistance and stewardship in Vietnamese hospital settings. This repository supports open, reproducible science by making analysis code available alongside study publications.

For enquiries about this repository or the underlying research, please open a GitHub Issue or contact the maintainer directly.
