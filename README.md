# Personal Expense Dashboard

A responsive GitHub Pages website that turns your Excel expense log into an external dashboard.

## What this version does

- You keep using Excel for data entry.
- The website reads `.xlsx`, `.xls`, or `.csv` files directly in the browser.
- Your expense data is not uploaded to a server by the website.
- Filters: month, category, payment method, search.
- KPIs: total spend, transaction count, daily average, largest expense.
- Charts: daily spend, category split, top categories.
- Tables: top expenses and the filtered expense log.

## GitHub Pages setup

1. Create a GitHub repository, for example `expense-dashboard`.
2. Upload everything in this folder to the repository.
3. On GitHub: **Settings → Pages → Source → GitHub Actions**.
4. Push to `main`; the included workflow publishes the site.
5. Your site will normally be at `https://YOUR-USERNAME.github.io/expense-dashboard/`.

GitHub documents GitHub Pages deployment through Actions, including `configure-pages@v5`, `upload-pages-artifact@v4`, and `deploy-pages@v4`.

## Daily use

Open the website, click **Load Excel**, and select your expense workbook. Enter new expenses in Excel as usual. The next time you open the site, load the updated workbook again.

## Important privacy note

Do not commit a personal expense workbook to a public GitHub repository. GitHub Pages sites are publicly available on the internet. This project deliberately keeps the Excel file outside the website so your financial data can remain private.

## Expected Excel columns

The site recognizes these columns:

`Date | Expense Description | Expense Type | Amount (INR) | Payment Method | Notes`

It also accepts common variations such as `Expense Name`, `Category`, `Amount`, and `Payment`.
