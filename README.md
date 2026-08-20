# Job Scraper — aryan2k1

A daily job scraper that checks a list of target companies for new
analyst/associate/graduate postings in London and produces an Excel tracker.
It runs automatically on **GitHub Actions** every weekday morning — no server to
keep running — and publishes the tracker as a downloadable build artifact.

## How it works

Each run queries four sources, filters to relevant early-career roles, dedupes
by URL, and writes/updates `job_tracker.xlsx`:

| Source | What it needs | What it covers |
|---|---|---|
| **Adzuna API** | Free API key | UK job boards, aggregated |
| **Reed API** | Free API key (optional) | Reed.co.uk (strong UK finance coverage) |
| **Greenhouse** | Nothing — public API | Companies on the Greenhouse ATS |
| **Lever** | Nothing — public API | Companies on the Lever ATS |

Relevance is controlled by `INCLUDE_KEYWORDS` / `EXCLUDE_KEYWORDS`, and the
company list by `TARGETS`, all in `scraper.py`.

## Project structure

```
job-scraper/
├── scraper.py                      # the scraper: sources, filters, Excel output
├── requirements.txt                # Python dependencies
├── .env.example                    # template for local API keys (copy to .env)
├── .gitignore
└── .github/workflows/daily_scrape.yml   # daily GitHub Actions schedule
```

The scraped `job_tracker.xlsx` is **not** committed to the repo — it is uploaded
as an Actions artifact so your job-search data stays private.

## Setup

### 1. Get your API keys (both free)

- **Adzuna** — https://developer.adzuna.com → register → copy your
  `Application ID` and `Application Key`.
- **Reed** (optional) — https://www.reed.co.uk/developers → register → copy your key.

Greenhouse and Lever need no key.

### 2. Run it on GitHub Actions (recommended)

1. Push this repo to GitHub.
2. **Settings → Secrets and variables → Actions → New repository secret**, add:

   | Name | Value |
   |---|---|
   | `ADZUNA_APP_ID` | your Adzuna Application ID |
   | `ADZUNA_APP_KEY` | your Adzuna Application Key |
   | `REED_API_KEY` | your Reed key (optional) |

3. **Actions → enable workflows** if prompted.

It runs every weekday at **08:00 UTC** (edit the `cron` line in the workflow to
change it). After each run: **Actions → latest run → Artifacts →
job-tracker-…** to download the Excel.

### 3. Run it locally

```bash
pip install -r requirements.txt
cp .env.example .env          # then paste your keys into .env
python scraper.py             # writes job_tracker.xlsx
```

## Excel tracker columns

| Column | Meaning |
|---|---|
| Company / Tier | Target firm and priority (1 or 2) |
| Role / Title | Job title |
| URL | Direct link to the posting |
| Location / Found / Source | City, date first seen, which API found it |
| Status / Applied? / Notes | Your own tracking fields |

## Adjusting it

- `TARGETS` — add/remove companies (with their Adzuna name and Greenhouse/Lever slug).
- `INCLUDE_KEYWORDS` / `EXCLUDE_KEYWORDS` — tune which titles count as relevant.
- `cron` in the workflow — change the schedule ([crontab.guru](https://crontab.guru), UTC).

## Tech stack

- **Python 3.11**
- **requests** — HTTP calls to the job APIs
- **openpyxl** — builds the formatted Excel tracker
- **GitHub Actions** — daily scheduling (cron) + artifact storage

## Notes

Only compliant public APIs are used. Sites that block scraping (LinkedIn,
Indeed, Totaljobs) are intentionally not included — set up their native email
alerts separately if you want that coverage.
