# Pocket Budget

A budgeting app for iPhone that runs offline from your Home Screen and mirrors everything to an Excel workbook.

## What it does

- **Profile:** your name, currency, how often you're paid and your take-home pay.
- **Debts:** sorted by type (credit card, student loan, auto loan, mortgage, personal loan, medical, line of credit, buy now pay later). For each one it records the amount owed, interest rate (APR), minimum payment and due day. It also shows:
  - monthly interest
  - payoff time at the minimum payment
  - total interest you'd pay
  - a warning if the minimum doesn't cover the interest
  - which debt to pay down first (the highest APR)
- **Income:** one-time or recurring (weekly, every 2 weeks, twice a month, monthly, quarterly, yearly).
- **Expenses:** logged by date, category and how you paid.
- **Recurring bills:** amount, how often, due day and whether it's on autopay.
- **Overview:** the money you can spend this month, worked out as:

  `income − recurring bills − debt minimums − money spent = expendable`

  It also shows a per-day amount for the rest of the month, what's due in the next 14 days, and spending by category.

## The Excel workbook

**Export Excel** creates `PocketBudget-YYYY-MM-DD.xlsx` with these sheets: **Summary, Profile, Debts, Income, Expenses, Bills**. The totals use real formulas, so if you edit the file in Excel or Numbers the totals update. On iPhone, the share sheet lets you save the file to Files, iCloud Drive or OneDrive.

**Import .xlsx** (on the Profile tab) loads a workbook and replaces the data in the app. Use it to restore a backup, move to a new phone, or bring back changes you made in the spreadsheet.

Your data stays on your phone (browser storage). Nothing is sent to a server.

## Installing on your iPhone

The app has to be served over HTTPS. The simplest free option is GitHub Pages:

1. Push this repo to GitHub. Go to **Settings → Pages**, choose your branch, and set the folder to `/ (root)`.
2. On your iPhone, open `https://<your-user>.github.io/<repo>/budget-app/` in **Safari**.
3. Tap **Share → Add to Home Screen**. It opens full screen like a native app and works offline.

Any static host works too (Netlify, Cloudflare Pages, Vercel).

To try it on your computer:

```bash
cd budget-app && python3 -m http.server 8000
```

Then open http://localhost:8000.

## Files

| File | Purpose |
| --- | --- |
| `index.html`, `styles.css`, `app.js` | The app |
| `sw.js`, `manifest.webmanifest`, `icons/` | Offline support and Home Screen install |
| `vendor/xlsx.full.min.js` | [SheetJS](https://sheetjs.com) 0.18.5 (Apache-2.0), used to read and write Excel files |

**Back up regularly:** if you clear Safari's website data, the app's data is deleted with it. Export the workbook now and then.
