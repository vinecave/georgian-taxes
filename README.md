# Georgian Small Business Tax Calculator

A Google Apps Script for Google Sheets that helps Georgian small business owners (1% tax status) log income, automatically convert foreign currency to GEL via the National Bank of Georgia exchange rates, and track annual revenue toward the 500,000 GEL threshold.

**Landing page:** [vinecave.github.io/georgian-taxes](https://vinecave.github.io/georgian-taxes/)

## Features

- **1% tax auto-calculation** on every income entry
- **NBG exchange rates** fetched automatically for USD, EUR, and RUB on the income date
- **Running sum** that accumulates across entries, with support for an initial carry-over balance
- **Threshold warnings** — Running Sum cell turns orange at 450,000 GEL and red at 500,000 GEL
- **rs.ge compatible** — all numeric values use dot decimal separator (not comma), stored as plain text to prevent locale conversion
- **Currencies supported:** GEL, USD, EUR, RUB

## Sheet Columns

| Column | Field | Description |
|--------|-------|-------------|
| A | Month | Which month this income belongs to (Jan–Dec) |
| B | Date Filled | Date the row was created (auto-filled) |
| C | Income Date | Actual date income was received |
| D | Currency | GEL / USD / EUR / RUB |
| E | Gross Revenue | Amount in original currency |
| F | Amount GEL | Converted to GEL via exchange rate |
| G | Running Sum | Cumulative annual revenue in GEL |
| H | Tax 1% | Amount GEL × 0.01 |
| I | Exchange Rate | NBG rate per unit of currency on the income date |

## Setup

1. Open a Google Sheet (or use the template from the [landing page](https://vinecave.github.io/georgian-taxes/))
2. Go to **Extensions → Apps Script**
3. Create two files in the Apps Script editor:
   - `Code.gs` — paste the contents of [Code.gs](Code.gs)
   - `Dialog.html` — click **+** next to Files, choose **HTML**, name it `Dialog`, paste the contents of [Dialog.html](Dialog.html)
4. Save both files (Ctrl+S / Cmd+S)
5. Close the Apps Script tab and **reload the spreadsheet**
6. A **Tax Calculator** menu will appear in the menu bar

## Usage

1. Click **Tax Calculator → Add Income**
2. On the first run, Google will ask you to authorize the script — follow the prompts and click Allow
3. Fill in the dialog:
   - **Month** — select from dropdown
   - **Income Date** — the actual date of income
   - **Currency** — GEL, USD, EUR, or RUB
   - **Gross Revenue** — amount in the selected currency
4. If this is your first entry (no existing data), an **Initial Running Sum** field appears so you can carry over a prior balance
5. Click **Add Row** — the script fetches the exchange rate, calculates everything, and appends a new row

## Exchange Rates

Rates are fetched from the [National Bank of Georgia](https://nbg.gov.ge) public API:

```
https://nbg.gov.ge/gw/api/ct/monetarypolicy/currencies/en/json/?currencies={CURRENCY}&date={YYYY-MM-DD}
```

- GEL transactions use a hardcoded rate of 1.0 (no API call)
- RUB is quoted per 100 units by NBG — the script handles this automatically

## Privacy

- No data is sent to any third-party service by the script
- The script reads and writes only to your own Google Sheet
- The only external request is to the NBG API for exchange rates — no personal or financial data is included
- All data stays in your Google account

## File Structure

```
Code.gs       — Server-side Apps Script (menu, dialog, API calls, row writing)
Dialog.html   — Modal dialog UI (HTML form served via HtmlService)
index.html    — GitHub Pages landing page with setup instructions
```

## License

Free and open source. Provided as-is with no warranty. You are responsible for verifying tax calculations before filing.
