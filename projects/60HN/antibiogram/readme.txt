================================================================================
 Chuẩn hóa Danh mục Vi sinh vật (Antibiogram Organism Name Standardizer)
 Project: 60HN | Version: 3 (Local Desktop)
================================================================================

WHAT IT DOES
------------
A local Shiny app that standardizes microbiology organism names in your Excel
files against a reference catalogue (DTH_danh_muc_vsv.xlsx). It flags
unrecognized names and lets you review, correct, and export the cleaned data.

HOW TO RUN
----------
1. Make sure R (>= 4.4.0) is installed: https://cran.r-project.org
2. Double-click run_app.bat — it will:
   - Find Rscript automatically
   - Restore all required packages from renv.lock (first run only, takes a few minutes)
   - Launch the app in your default browser

FILES
-----
  app.R                  Main Shiny application
  setup_and_run.R        Called by run_app.bat — handles renv restore + app launch
  run_app.bat            Double-click launcher (Windows)
  renv.lock              Pinned package versions (do not edit manually)
  renv/                  renv internals (do not edit manually)
  DTH_danh_muc_vsv.xlsx  Reference catalogue of standardized organism names
  output/                Folder for exported results

REQUIREMENTS
------------
  - Windows OS
  - R >= 4.4.0
  - Internet connection on first run (for package installation)

NOTES
-----
  - On first run, renv::restore() installs ~65 packages — this is normal.
  - Subsequent runs skip restoration and launch immediately.
  - Do not move app.R or renv.lock out of this folder.
================================================================================