# JoSAA IIT / NIT / IIIT Cut-off Book — Vercel Deployment

Everything this app needs is **inside this folder** — Excel, Word template,
ready-made book, parsed-data cache, logo, templates. No absolute `C:\...`
paths are used, so it runs the same on Vercel (Linux) or locally.

The app **auto-scans the Excel on page load** — no button click needed.
A bundled JSON cache (`data/parsed_cache.json`, keyed by the Excel's content
hash) makes the first load instant even after a cold start (~0.2s instead of
~7s), a localStorage cache paints the list the moment the page opens, and the
last-opened institute is restored on reload. The **Refresh Excel** button
force-re-reads the workbook.

## Files

| File | Purpose |
|---|---|
| `api/index.py` | Vercel entry point (exports the Flask `app`) |
| `vercel.json` | Build/routes config, `maxDuration: 300` |
| `requirements.txt` | Python deps (pywin32 installs on Windows only) |
| `web_app.py` | Flask app — all routes |
| `parser.py` / `book_updater.py` / `institute_docx.py` / `metadata_db.py` | Excel parsing, book rebuild, per-institute Word |
| `word_com.py` | MS Word page-fit/index fix — **Windows only**, auto-skipped on Vercel |
| `data/iit,iiit,nit 2026.xlsx` | Rank data (source of truth) |
| `data/ENGINNERING FINAL BOOK sachin sir.docx` | Word template |
| `data/ENGINNERING FINAL BOOK 2027 - UPDATED IIT NIT IIIT.docx` | Ready-made book, served immediately by **Book DOCX** |
| `data/parsed_cache.json` | Pre-built parse cache → instant scan on load |
| `templates/`, `static/` | UI + logo |

## Deploy to Vercel

1. Extract this zip.
2. Either:
   - **Dashboard**: https://vercel.com/new → *Upload* / drag this folder
     (it contains `vercel.json` at the root), or import a Git repo with it; **or**
   - **CLI**: `npm i -g vercel` → run `vercel` inside this folder → `vercel --prod`.
3. Login on the deployed site: user `josaa2027`, password `josaa@2027`.

## What works on Vercel

- Login, Scan (Excel bundled), institute cut-off tables, **Download Word** (per
  institute, ranks straight from Excel), **Book DOCX** (serves the bundled
  ready-made book instantly).
- **Build Full Book** runs synchronously in one request (Excel+template → new
  book in `/tmp`). Needs a few minutes — if the plan's function timeout is
  shorter, the bundled ready-made book is still downloadable.

## What is skipped on Vercel

- MS Word steps (page-fit + index page numbers) — needs desktop Word, so
  `word_error: "Skipped: MS Word is not available on this server."` is shown in
  the report. Run `run.bat` on a Windows machine with Word for those steps.

## Run locally (optional)

```
pip install -r requirements.txt
python web_app.py     # http://localhost:5000
```
