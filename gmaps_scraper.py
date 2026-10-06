"""
Google Maps Lead Scraper
-------------------------
Scrapes business listings from Google Maps (name, address, phone, website,
rating, category) for a given search query, then visits each business's
website to try to find a contact email, and saves everything to an Excel file.

Usage:
    1. cp .env.example .env      (then edit SEARCH_QUERY etc. inside .env)
    2. python gmaps_scraper.py

Notes:
    - Google Maps' HTML structure changes periodically. If selectors stop
      matching, see the "SELECTOR MAINTENANCE" note near the top of the file.
    - This scrapes PUBLICLY visible business info. Be a good citizen: keep
      MAX_RESULTS reasonable, don't hammer sites, and check Google's Terms of
      Service / robots.txt before using this for anything beyond personal /
      small-scale research use.
"""

import os
import re
import time
import sys
from dataclasses import dataclass, field, asdict

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
load_dotenv()

SEARCH_QUERY = os.getenv("SEARCH_QUERY", "restaurants in Lahore")
MAX_RESULTS = int(os.getenv("MAX_RESULTS", "50"))
HEADLESS = os.getenv("HEADLESS", "false").strip().lower() == "true"
WEBSITE_VISIT_DELAY = float(os.getenv("WEBSITE_VISIT_DELAY", "2"))
OUTPUT_FILE = os.getenv("OUTPUT_FILE", "leads.xlsx")

# SELECTOR MAINTENANCE:
# Google frequently changes internal class names (things like "DUwDvf") but
# the "data-item-id" attributes and role="feed"/role="article" structure has
# stayed fairly stable. If scraping breaks, open Google Maps in a real
# browser, right-click a listing -> Inspect, and update these selectors.
FEED_SELECTOR = 'div[role="feed"]'
CARD_SELECTOR = 'div[role="feed"] a[href^="https://www.google.com/maps/place"]'
NAME_SELECTOR = "h1"
ADDRESS_SELECTOR = 'button[data-item-id="address"]'
PHONE_SELECTOR = 'button[data-item-id^="phone:tel:"]'
WEBSITE_SELECTOR = 'a[data-item-id="authority"]'
RATING_SELECTOR = 'div[jsaction*="pane.rating"] span[aria-hidden="true"]'
CATEGORY_SELECTOR = 'button[jsaction*="category"]'

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
# Common junk/false-positive domains found in tracking pixels, image files, etc.
EMAIL_BLOCKLIST_DOMAINS = ("sentry.io", "example.com", "wixpress.com", "godaddy.com")


@dataclass
class Business:
    name: str = ""
    category: str = ""
    address: str = ""
    phone: str = ""
    website: str = ""
    rating: str = ""
    email: str = ""
    maps_url: str = ""


# ---------------------------------------------------------------------------
# STEP 1: Scrape listing cards from the Google Maps results feed
# ---------------------------------------------------------------------------
def scrape_google_maps(page, query: str, max_results: int) -> list[Business]:
    url = f"https://www.google.com/maps/search/{query.replace(' ', '+')}"
    page.goto(url, timeout=60000)

    # Dismiss cookie consent dialog if it appears (varies by region)
    try:
        page.get_by_role("button", name=re.compile("Accept all|I agree", re.I)).click(timeout=4000)
    except PWTimeout:
        pass

    try:
        page.wait_for_selector(FEED_SELECTOR, timeout=15000)
    except PWTimeout:
        print("Could not find the results feed. Google may have changed its layout, "
              "or your query returned a single result / no results.")
        return []

    seen_urls = set()
    businesses: list[Business] = []
    stagnant_scrolls = 0

    # Scroll the results feed to lazy-load more listings
    while len(seen_urls) < max_results and stagnant_scrolls < 6:
        cards = page.query_selector_all(CARD_SELECTOR)
        before = len(seen_urls)
        for card in cards:
            href = card.get_attribute("href")
            if href and href not in seen_urls:
                seen_urls.add(href)
            if len(seen_urls) >= max_results:
                break

        if len(seen_urls) == before:
            stagnant_scrolls += 1
        else:
            stagnant_scrolls = 0

        page.hover(FEED_SELECTOR)
        page.mouse.wheel(0, 2000)
        time.sleep(1.2)

    print(f"Found {len(seen_urls)} listing links. Visiting each for details...")

    for i, listing_url in enumerate(list(seen_urls)[:max_results], start=1):
        biz = Business(maps_url=listing_url)
        try:
            page.goto(listing_url, timeout=30000)
            page.wait_for_selector(NAME_SELECTOR, timeout=10000)
            time.sleep(0.8)  # let side panel fully render

            biz.name = _safe_text(page, NAME_SELECTOR)
            biz.address = _safe_text(page, ADDRESS_SELECTOR)
            biz.phone = _safe_text(page, PHONE_SELECTOR).replace("Phone: ", "")
            biz.rating = _safe_text(page, RATING_SELECTOR)
            biz.category = _safe_text(page, CATEGORY_SELECTOR)

            website_el = page.query_selector(WEBSITE_SELECTOR)
            if website_el:
                biz.website = website_el.get_attribute("href") or ""

            print(f"[{i}/{len(seen_urls)}] {biz.name or 'Unnamed'} -> "
                  f"{biz.website or 'no website'}")
        except PWTimeout:
            print(f"[{i}/{len(seen_urls)}] Timed out loading listing, skipping.")
        except Exception as e:
            print(f"[{i}/{len(seen_urls)}] Error scraping listing: {e}")

        businesses.append(biz)

    return businesses


def _safe_text(page, selector: str) -> str:
    el = page.query_selector(selector)
    if not el:
        return ""
    return (el.inner_text() or "").strip()


# ---------------------------------------------------------------------------
# STEP 2: Visit each business website and try to extract an email
# ---------------------------------------------------------------------------
def find_email_on_website(page, url: str) -> str:
    if not url:
        return ""

    pages_to_try = [url]
    # Many sites only put contact info on a /contact page
    for suffix in ("/contact", "/contact-us", "/contactus"):
        pages_to_try.append(url.rstrip("/") + suffix)

    for target in pages_to_try:
        try:
            page.goto(target, timeout=15000, wait_until="domcontentloaded")
        except Exception:
            continue

        # 1. Check mailto: links first (most reliable)
        mailto_el = page.query_selector('a[href^="mailto:"]')
        if mailto_el:
            href = mailto_el.get_attribute("href") or ""
            email = href.replace("mailto:", "").split("?")[0].strip()
            if _is_valid_email(email):
                return email

        # 2. Fall back to regex over visible page text
        try:
            body_text = page.inner_text("body")
        except Exception:
            body_text = ""
        matches = EMAIL_REGEX.findall(body_text)
        for m in matches:
            if _is_valid_email(m):
                return m

    return ""


def _is_valid_email(email: str) -> bool:
    email = email.lower()
    if not EMAIL_REGEX.fullmatch(email):
        return False
    return not any(bad in email for bad in EMAIL_BLOCKLIST_DOMAINS)


# ---------------------------------------------------------------------------
# STEP 3: Save results to Excel
# ---------------------------------------------------------------------------
def save_to_excel(businesses: list[Business], output_file: str):
    df = pd.DataFrame([asdict(b) for b in businesses])
    # Nice column order
    cols = ["name", "category", "phone", "email", "website", "address", "rating", "maps_url"]
    df = df[[c for c in cols if c in df.columns]]

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Leads")
        worksheet = writer.sheets["Leads"]
        # Auto-fit column widths roughly based on content length
        for i, col in enumerate(df.columns, start=1):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = min(max_len, 60)

    print(f"\nSaved {len(df)} rows to {output_file}")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    print(f"Searching Google Maps for: '{SEARCH_QUERY}'  (max {MAX_RESULTS} results)")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
        )
        page = context.new_page()

        businesses = scrape_google_maps(page, SEARCH_QUERY, MAX_RESULTS)

        if not businesses:
            print("No businesses found. Exiting.")
            browser.close()
            sys.exit(1)

        print("\nNow visiting business websites to look for emails...")
        for i, biz in enumerate(businesses, start=1):
            if biz.website:
                email = find_email_on_website(page, biz.website)
                biz.email = email
                print(f"[{i}/{len(businesses)}] {biz.name}: "
                      f"{email if email else 'no email found'}")
            time.sleep(WEBSITE_VISIT_DELAY)

        browser.close()

    save_to_excel(businesses, OUTPUT_FILE)


if __name__ == "__main__":
    main()
