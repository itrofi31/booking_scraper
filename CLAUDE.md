# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Directory structure

```
booking_scraper/
├── booking/                  ← Booking.com scraper
│   ├── scraper.py
│   ├── properties.json
│   ├── scraper.log           (auto-created)
│   └── booking_*.xlsx        (auto-created)
├── airbnb/                   ← Airbnb scraper
│   ├── airbnb_scraper.py
│   ├── airbnb_properties.json
│   ├── airbnb_scraper.log    (auto-created)
│   └── airbnb_*.xlsx         (auto-created)
├── requirements.txt
├── CLAUDE.md
└── bin/python3               (venv)
```

## Commands

```bash
# Install dependencies (once)
./bin/python3 -m pip install playwright openpyxl playwright-stealth
./bin/playwright install chromium

# ── Booking.com ──────────────────────────────────────────────────────────────
# All properties, default dates
./bin/python3 booking/scraper.py

# Single property, single date pair
./bin/python3 booking/scraper.py --property ozone --checkin 2026-08-01 --checkout 2026-08-08

# From dates file
./bin/python3 booking/scraper.py --dates-file dates.txt --workers 4

# Visible browser (for debugging)
./bin/python3 booking/scraper.py --property ozone --visible

# ── Airbnb ────────────────────────────────────────────────────────────────────
# All properties, default dates
./bin/python3 airbnb/airbnb_scraper.py

# Single property, single date pair
./bin/python3 airbnb/airbnb_scraper.py --property ozone --checkin 2026-08-01 --checkout 2026-08-08

# Visible browser (for debugging)
./bin/python3 airbnb/airbnb_scraper.py --property ozone --visible
```

Booking workers default 4, max 8. Airbnb workers default 3, max 6.
Output files auto-generated in the same folder as the script.

## Architecture

Single-file scraper: `scraper.py` + config: `properties.json`.

**Data flow:**
1. `main()` → parses CLI args, calls `run()`
2. `run()` → builds job list per date period, calls `scrape_batch()` twice (Pass 1: all periods, Pass 2: retry failures)
3. `scrape_batch()` → parallel asyncio workers, each gets its own browser context via `new_context()`
4. `scrape_property()` → navigates to Booking.com hotel page with date params in URL, extracts offers via `page.evaluate()` (inline JS), returns `PropertyResult`
5. `export_excel()` → writes openpyxl workbook with two sheets: "Все тарифы" (all offers, grouped by date then property) and "Сводка" (summary per property per date)

**Key design decisions:**

- Each worker gets a fresh Chromium context (isolated cookies/session) — avoids cross-contamination
- URL includes `checkin`, `checkout`, `selected_currency=THB`, `sb_price_type=total` directly — calendar interaction is a fallback only
- All DOM extraction happens in a single `page.evaluate()` JS blob to minimize round-trips; `pageCancels` is collected page-wide because cancellation policy elements are not always inside the row
- "Execution context was destroyed" happens when Booking.com JS redirects mid-evaluate — handled by 3-attempt retry loop with full re-navigation (not just sleep)
- Two-pass retry (Pass 1 scrapes everything, Pass 2 retries all failures) prevents one slow retry from blocking other periods
- Price sanity check: skip any price < `nights × 300฿` (catches spurious DOM elements)

**`properties.json` structure:**
```json
{
  "default_dates": [{"checkin": "...", "checkout": "..."}],
  "ozone": {
    "label": "Ozone 1BR",
    "url": "https://www.booking.com/hotel/th/...",
    "competitors": [{"label": "...", "url": "..."}]
  },
  "cassia": { ... }
}
```
Own property has `is_own: true` in jobs; competitor rows get `is_own: false`. The "vs Мой мин %" column compares like-for-like: refundable offers vs own refundable minimum, non-refundable vs own non-refundable minimum (fallback to overall minimum if the specific type has no price).

**DOM selectors (may break on Booking.com HTML updates):**
- Modern room rows: `[data-testid="availability-table-row"]`
- Legacy fallback: `#hprt-table tr`
- Cancellation text: `[data-testid="policy-title"]` → `htmlToText()` (strips tags to get date after `<strong>`)
- `classifyCancel()`: non-refundable if text contains `non-refund`, `no refund`, `not refund`, or `reschedule`

**Excel layout:**
- Row 1: column headers (frozen via `freeze_panes = "C2"`)
- Each date period: dark header row, then property groups with thick blue outer border
- Own property rows: green background (`★` prefix), competitor rows: white/light blue
- Column order: Объект | Тип номера | Гостей | Цена ฿ | ADR ฿ | До скидки ฿ | Скидка % | Возвратность | vs Мой мин %
- ADR = `price_final / nights`

**Debugging scraping failures:**
- Check `scraper.log` for per-property errors
- Common errors: `no_room_table` (page didn't load table), `no_offers` (table found but no prices parsed), `context_destroyed` (Booking redirect)
- Run with `--visible` to watch the browser; uncomment `page.screenshot()` in `scrape_property()` for `no_room_table` cases
- If all URLs fail with `ERR_NAME_NOT_RESOLVED` — network issue, not a code bug
