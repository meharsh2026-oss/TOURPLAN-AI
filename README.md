# TourPlan AI

Automated Tour Plan generator. Upload an existing Tour Plan / PJP workbook,
pick a target month and year, and TourPlan AI rewrites the same workbook —
dates, weekdays, Sunday off-days and the monthly VALUE total — entirely in
your browser. Nothing is uploaded to a server.

## How it works

1. **Upload** — drop in the `.xlsx` workbook you already use.
2. **Detect** — the app scans the workbook for the `Date | Day | Town Name |
   Work With | DB NAME | TC | PC | VALUE` header row, the data block, and the
   `SUM(...)` total-formula row. Nothing about row/column position is
   hardcoded — it's inferred from the file itself, so it adapts to minor
   layout differences between workbooks.
3. **Choose a month/year** — defaults to the month after the one detected in
   the upload, but any month/year can be selected.
4. **Generate** — the app rewrites the data block in place:
   - Every date is replaced with the corresponding date in the target month.
   - The Day column is recalculated from the actual weekday.
   - Sundays are cleared (Town Name / Work With / DB NAME / TC / PC / VALUE),
     matching the convention already used in the source file.
   - Town Name / Work With / DB NAME / TC / PC / VALUE for working days are
     carried over from the source workbook's own rows (cycled if the target
     month has more working days than the source had).
   - The monthly VALUE total is rebalanced to exactly **250000** by adjusting
     only the **last working day** of the month — every other VALUE is left
     untouched.
   - If the target month is shorter than the template (e.g. 28 vs 31 rows),
     the extra row slots are cleared, mirroring how the source workbook
     itself leaves trailing rows blank for short months.
5. **Download** — `TourPlan_<Month>_<Year>.xlsx`, built from the original
   workbook so headers, fonts, fills, merges, and column widths are
   preserved.

## Project structure

```
/
├── index.html     UI markup
├── style.css      Dark glassmorphic design system
├── app.js         UI orchestration (upload → generate → download)
├── excel.js       Workbook structure detection + mutation (SheetJS-based)
├── planner.js     Pure calendar/date logic (weekdays, leap years, etc.)
├── value.js       Pure VALUE-redistribution logic (the 250000 rule)
└── README.md
```

`excel.js` and `app.js` use native ES modules (`import`/`export`), loaded via
`<script type="module">` — no build step is required.

## Running locally

Because the app uses ES modules, open it through a local web server rather
than the `file://` protocol (browsers block module imports over `file://`).

```bash
# any static file server works, for example:
npx serve .
# or
python3 -m http.server 8000
```

Then visit the printed local URL.

## Deploying to GitHub Pages

1. Push this folder to a GitHub repository.
2. In the repository settings, enable **Pages** → **Deploy from a branch**
   → select the branch and `/ (root)` folder.
3. GitHub Pages serves static files over HTTPS by default, which satisfies
   the ES module same-origin requirement — no further configuration needed.

## Business rules implemented

| Rule | Behavior |
|---|---|
| Date rewrite | Every date in the data block is replaced with the matching date in the target month/year. |
| Weekday calculation | The Day column is derived from the actual JS `Date` weekday — never copied from the source. |
| Month length | 28/29/30/31-day months are all supported; trailing template rows beyond the month's last day are cleared. |
| Leap years | Handled natively via `new Date(year, month + 1, 0).getDate()` — no manual leap-year table. |
| Sunday rule | On a Sunday, Date and Day are kept; Town Name, Work With, DB NAME, TC, PC and VALUE are cleared. |
| Monthly VALUE rule | The sum of VALUE across all working days (Sundays ignored) is rebalanced to exactly 250000 by adjusting **only** the last working day's VALUE. |
| Template fidelity | Headers, column order, row order, fonts, fills and merged cells are preserved — the original workbook is mutated, not rebuilt from scratch. |

## Notes on the Excel engine

The app loads **ExcelJS** (browser bundle) instead of any SheetJS-family
library. ExcelJS keeps the full OOXML document model in memory — styles,
merges, borders, fonts, fills, row heights, column widths, hidden
rows/columns, print settings, formulas — and only changes what the code
actually touches. Every write in `excel.js` sets `cell.value` and nothing
else, so every other property already on a cell (inherited from the
uploaded template) survives the round trip untouched. This is what fixes
the formatting loss (grey headers, fills, merged cells, borders, fonts,
row heights, column widths) that happened with the previous
xlsx-js-style-based implementation.

`excel.js` is now async: `parseWorkbook()` and `exportWorkbook()` both
return Promises (`workbook.xlsx.load()` / `workbook.xlsx.writeBuffer()`
under the hood), since ExcelJS's read/write APIs are asynchronous. `app.js`
awaits both.

## Error handling

- Uploading a non-`.xlsx` file is rejected with an inline message.
- A workbook that doesn't contain a recognizable `Date | Day | ...` header
  row is rejected with a descriptive message instead of failing silently.
- A workbook with no working-day rows to use as a template is rejected the
  same way.
- If a selected month has more days than the template has row slots, the
  result panel shows a warning explaining exactly how many days could not
  be added.
