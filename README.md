# Russell Toolkit

Internal tools for Russell — a small, modular FastAPI app that hosts utilities like the Invoice Analyzer and Directory Admin, with room to grow into document processing and other ops tools.

The first tool, **Invoice Analyzer**, helps Accounts Payable break software license charges down by **Branch + License**.

Built as a lightweight FastAPI app so it runs on:
- macOS
- Windows
- barebones Linux hardware (no heavy frontend build required)

## What it does

1. Choose vendor mode (`Hexnode`, `Adobe`, `Integricom Licensing`, `Integricom Support Hours`, or `Generic`)
2. Upload one invoice file (optional for Generic, recommended for Hexnode, required for Adobe, Integricom Licensing, and Integricom Support Hours)
3. Upload one or more CSV exports from vendors (Microsoft, Adobe, etc.) when needed
4. Auto-detect common columns for branch, license, and amount
5. Aggregate totals into a pivot-style breakdown
6. Download `breakdown.csv`

Hexnode mode includes:
- fixed $2/device cost
- `Default User` usernames mapped to `Home Office`
- invoice reconciliation (`invoice total - source total`) added to `Home Office`

Adobe mode includes:
- parsing per-license pricing from invoice line items
- a living SQLite user directory (`app/data/adobe_users.sqlite3`)
- automatic detection of new users and users missing from each new export
- prompt-driven user enrichment (branch) for new users
- allocation by user/product from Adobe `users.csv` export using saved user directory records
- invoice reconciliation added as `Adobe Invoice Adjustment` under `Home Office`

Integricom Licensing mode includes:
- parsing line-item quantities and totals from Integricom invoice PDFs
- a living SQLite user directory (`app/data/integricom_users.sqlite3`)
- Microsoft user export ingestion (`User principal name`, `Office`, `Licenses`)
- dynamic user-based allocation for workstation/cloud-backup/M365 core lines
- fixed-template allocation for network/security/site charges
- prompt-driven branch assignment for extra branch-tethered quantities (instead of auto-charging Home Office)
- Dropbox forced to `Home Office` allocation
- per-user editable branch table with dedicated save button
- optional no-CSV workflow via Microsoft Entra sync (Admin app)

Integricom Support Hours mode includes:
- parsing invoice time-detail blocks (`Charge To`) and including only entries where `Bill = Y`
- block-level allocation (one branch per billable support block)
- automatic branch guess from charge summary text
- low-confidence default to `Home Office` with a review queue
- editable branch review table with re-run support to update totals/export

## Docker (Recommended Packaging)

Run from this folder:

```bash
cd /Users/prestonpierce/Documents/TestCodexProj/russell-toolkit
docker compose up --build -d
```

Open:
- App launcher: [http://127.0.0.1:8080](http://127.0.0.1:8080)
- Invoice Analyzer: [http://127.0.0.1:8080/apps/invoice-analyzer](http://127.0.0.1:8080/apps/invoice-analyzer)
- Admin: [http://127.0.0.1:8080/apps/admin](http://127.0.0.1:8080/apps/admin)

Useful commands:

```bash
docker compose logs -f
docker compose down
```

Data persistence:
- `./app/data` is mounted into the container at `/app/app/data`.
- Your Adobe/Integricom SQLite living DB files persist across container restarts/rebuilds.

## Run locally

```bash
cd /Users/prestonpierce/Documents/TestCodexProj/russell-toolkit
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8080 --reload
```

Open:
- App launcher: [http://127.0.0.1:8080](http://127.0.0.1:8080)
- Invoice Analyzer: [http://127.0.0.1:8080/apps/invoice-analyzer](http://127.0.0.1:8080/apps/invoice-analyzer)
- Admin: [http://127.0.0.1:8080/apps/admin](http://127.0.0.1:8080/apps/admin)

## Windows quick start (PowerShell)

```powershell
cd C:\path\to\russell-toolkit
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8080 --reload
```

## API endpoint

- `POST /api/analyze` (multipart form)
  - `vendor_type`: `hexnode`, `adobe`, `integricom`, `integricom_support`, or `generic` (defaults to `generic`)
  - `csv_files`: one or more CSV files (required for `hexnode`, `adobe`, and `generic`; optional for `integricom`; not used for `integricom_support`)
  - `invoice_file`: invoice file (required for `hexnode`, `adobe`, `integricom`, and `integricom_support`)
  - `adobe_user_updates`: optional JSON array used when Adobe mode prompts for new user details
  - `integricom_user_updates`: optional JSON array used when Integricom mode saves branch edits during analyze
  - `integricom_branch_item_updates`: optional JSON array used when Integricom mode prompts for extra branch-tethered assignments
  - `integricom_support_updates`: optional JSON array used when Integricom Support Hours mode applies branch review edits
- `POST /api/adobe/users/save` (JSON array) to save Adobe user branch edits
- `POST /api/integricom/users/save` (JSON array) to save Integricom user branch edits
- `GET /api/adobe/users` and `GET /api/integricom/users` to list active directory users for admin tooling
- `POST /api/adobe/users/deactivate` and `POST /api/integricom/users/deactivate` with `{"emails":[...]}` to deactivate users
- `POST /api/integricom/sync/entra` to pull Integricom users directly from Microsoft Entra (Graph app-only)

## Microsoft Entra Sync (Integricom)

Set these environment variables where the app runs:
- `ENTRA_TENANT_ID`
- `ENTRA_CLIENT_ID`
- `ENTRA_CLIENT_SECRET`

Docker option:
- copy `.env.example` to `.env`
- fill values, then `docker compose up -d --build`

Expected Entra app permissions (Application):
- `User.Read.All`
- `LicenseAssignment.Read.All`

Then use Admin app:
- [http://127.0.0.1:8080/apps/admin](http://127.0.0.1:8080/apps/admin)
- Choose `Integricom`
- Click `Sync from Entra`

## Expected CSV fields

The app can detect many common aliases. These are best:
- `Branch`
- `License` or `Product`
- `Amount`

Fallback supported:
- `Quantity` + `Unit Price` (if `Amount` is missing)

## Sample files

- `/Users/prestonpierce/Documents/TestCodexProj/russell-toolkit/samples/microsoft-export.csv`
- `/Users/prestonpierce/Documents/TestCodexProj/russell-toolkit/samples/adobe-export.csv`

## Tests

```bash
cd /Users/prestonpierce/Documents/TestCodexProj/russell-toolkit
source .venv/bin/activate
pip install -r requirements-dev.txt
pytest -q
```

## Next recommended enhancements

1. Add invoice parsers for PDF/email bill formats to auto-validate totals.
2. Save recurring column mappings per vendor.
3. Export directly to AP spreadsheet format.
