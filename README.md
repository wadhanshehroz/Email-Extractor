# Google Maps Lead Scraper — Setup Guide

Scrapes business listings from Google Maps (name, phone, address, website,
rating) for any search you give it, then visits each business's website to
try to find a contact email, and saves everything neatly into an Excel file.

## 1. Files in this project

```
gmaps_scraper/
├── gmaps_scraper.py   # the scraper itself
├── .env.example       # config template (copy to .env)
└── README.md          # this file
```

## 2. One-time setup

You said you've already installed `playwright`, `pandas`, `openpyxl`, and
`python-dotenv` — good. There's one more required step: Playwright needs to
download an actual browser binary (Chromium) the first time you use it.

```bash
playwright install chromium
```

(If that command isn't found, try `python -m playwright install chromium`.)

## 3. Configure your search

Copy the example env file and edit it:

```bash
cp .env.example .env
```

Open `.env` and set:

```
SEARCH_QUERY="dentists in Lahore"
MAX_RESULTS=50
HEADLESS=false
WEBSITE_VISIT_DELAY=2
OUTPUT_FILE=leads.xlsx
```

- **SEARCH_QUERY**: exactly what you'd type into Google Maps' search bar —
  combine a keyword and a location, e.g. `"gyms in Karachi"`, `"law firms
  in New York"`.
- **MAX_RESULTS**: how many listings to collect. Start small (10–20) to test
  before running a big batch.
- **HEADLESS**: keep as `false` the first few runs so you can *watch* the
  browser and confirm it's working. Switch to `true` once you trust it — it
  runs faster and in the background.
- **WEBSITE_VISIT_DELAY**: pause (in seconds) between visiting each business
  website. Keep this at 2+ seconds so you don't hammer other people's
  servers or get temporarily blocked.

## 4. Run it

```bash
python gmaps_scraper.py
```

What happens, step by step:
1. Opens Google Maps and searches your query.
2. Scrolls the results panel to load listings up to `MAX_RESULTS`.
3. Clicks into each listing and reads its name, category, address, phone,
   website, and rating.
4. Visits each business's website (and tries `/contact`, `/contact-us` too)
   looking for a `mailto:` link or an email-looking string on the page.
5. Writes everything into `leads.xlsx` (or whatever you set `OUTPUT_FILE`
   to), with readable column widths.

You'll see live progress printed in the terminal, e.g.:

```
[12/50] Riverside Dental Clinic -> https://riversidedental.com
[12/50] Riverside Dental Clinic: contact@riversidedental.com
```

## 5. Output

An Excel file with these columns:

| name | category | phone | email | website | address | rating | maps_url |
|------|----------|-------|-------|---------|---------|--------|----------|

Rows where no email could be found will just have a blank `email` cell —
not every business publishes one publicly.

## 6. Things worth knowing

- **Selectors may need updates over time.** Google changes its internal
  Maps HTML class names periodically. If the script suddenly finds 0
  results, open Google Maps in a normal browser, right-click a listing →
  "Inspect", and compare against the selectors defined near the top of
  `gmaps_scraper.py` (they're commented so you know what each one is for).
- **Email accuracy isn't guaranteed.** The script pulls `mailto:` links
  first (most reliable), then falls back to scanning page text for
  anything that looks like an email — this occasionally catches junk
  (e.g. a support email meant for a different purpose). Spot-check before
  using the list for outreach.
- **Be a good citizen.** This only reads publicly visible information, but
  scraping at scale can still violate a site's Terms of Service or get
  your IP rate-limited/blocked. Keep `MAX_RESULTS` and visit delays
  reasonable, and don't hit the same targets repeatedly in a short window.
- **CAPTCHAs.** If Google detects unusual automated traffic, it may show a
  CAPTCHA. Running with `HEADLESS=false` and not overusing the script from
  the same IP helps avoid this. There's no captcha-bypass logic here on
  purpose — if you hit one, solve it manually in the visible browser
  window and the script can usually continue.

## 7. Extending it

Ideas if you want to take this further:
- Add a `keywords` list in `.env` and loop `SEARCH_QUERY` over multiple
  cities/categories in one run, appending to the same Excel file.
- Add retry logic around `page.goto()` calls for flaky connections.
- Swap the Excel output for a live Google Sheet using `gspread` if you want
  cloud-synced results instead of a local file.
