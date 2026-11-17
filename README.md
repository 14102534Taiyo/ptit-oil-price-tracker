# PTIT Oil Price Structure & Retail Tracker

Auto-synced Thailand Oil Price Structure (EPPO) & Daily Retail Price Dashboard (PTTOR • Bangchak • Caltex), built with the same Git-based JSON ETL + Next.js Static Export architecture as `ptit-focus-dmf`.

## Architecture
1. **Auto-Sync ETL (`.github/workflows/eppo_retail_sync.yml` & `.github/scripts/sync_eppo.py`)**
   - Runs automatically at `05:05`, `11:00`, and `17:30` (Bangkok Time) or manually via `workflow_dispatch`.
   - Downloads new/updated `.xlsx` files from EPPO, parses all 10 products, and saves per-year JSON caches (`data/cache/eppo/YYYY.json` & `frontend/public/data/eppo/YYYY.json`).
   - Commits updated JSON files back to `main` with `[skip ci]`.
2. **GitHub Pages Deployment (`.github/workflows/deploy_pages.yml`)**
   - Automatically builds Next.js static export (`frontend/out`) after each successful sync and deploys to GitHub Pages.

## Local Development
```bash
# 1. Sync latest EPPO data into JSON
python .github/scripts/sync_eppo.py

# 2. Run Next.js dev server
cd frontend
npm install
npm run dev
```
