# BCM Implementation Toolkit (web)

A website that replaces data entry in the Almarai BCM Toolkit Template v0.2
(`template/Almarai_BCM_Toolkit_Template_v0.2.xlsx`, a blank copy of the Bakery v0.2 workbook).
Every sheet of the workbook has a page; automatic columns are recalculated live, and
**Export to Excel** produces a populated copy of the original workbook.

Open `bcm-toolkit/index.html` (or `/bcm-toolkit/` on the deployed site). No build step,
no server code - it is a static site.

## How it maps to the workbook

| Sheet | Page | What you fill |
|---|---|---|
| 00 Document Control | form | version, dates, status, sign-offs, version history |
| 01 Instructions | reference | read-only |
| 02 Dashboard | live view | - (all counters calculated) |
| 03 Business Unit Profile | form | sections 1-11 (v0.2 order: identity, leadership, internal, external, climate, interested parties, scope, legal, assets, risk appetite, approval) |
| 04 Products & Services | table (45 rows, 5-49) | BU, site, product, owner, scope, corporate BIA |
| 05 Processes & Activities | table (207 rows) | activities per product code |
| 06 Application & Equipment | grid | application / service / equipment lists that feed dropdowns |
| 07 BIA | one block per in-scope P&S (40 blocks) | interruption Y/N, 1-5 impact scores over 9 horizons, BU critical, MBCO, RTO, RPO, date |
| 08 Dependency Register | table (205 rows) | dependencies per BIA |
| 09 Risk Assumption Map | grid | assumptions per dependency category |
| 10 Risk Assessment | table (57 rows) | scope, assumption, ERM ref, controls, likelihood, impact, remarks |
| 11 Strategies & Solutions | table (60 rows) | strategy, requirement flags, owner, dates |
| 12 Requirements Register | table (300 rows) | requirements per S&S |
| 13 IT DR Register | table (200 rows) | systems, RTO/RPO required vs capable, DR tests |
| 14 Recovery & Response | table (100 rows, 5-104) | actions, plans, exercises |
| 15 BCM Risk Register | live view | - (consolidated automatically) |
| 16 CAPA Register | table (100 rows) | findings through verified closure |
| 17 BCMS Progress (PDCA) | table | manual override, owner & next action |
| 18 ISO Docs & Records | table (20 docs) | status and evidence metadata |
| 19 ISO 22301 Map | table | owner, evidence, verification fields |
| 20 ERM RA Guidance | reference | read-only (heat map drives risk ratings) |
| 21 Analysis | live view | - |
| 99 Engine | (internal lists) | used for dropdowns |

Row capacities are the workbook's own (e.g. 45 P&S, 40 BIA blocks, 57 risks), so every
row exported has a place in the template. Sheets hidden in the workbook (01, 17, 18, 19, 21)
are listed at the end of the menu.

**Paste from Excel:** select a block of cells in any Excel workbook, copy, click the first
target cell on a page and press Ctrl+V. Values fill across columns and down rows (rows are
added as needed), dropdown values are matched case-insensitively, dates such as
`15/01/2026`, `15-Jan-26` or `2026-01-15` are recognised, and automatic/label columns are skipped.

## Export / import

* **Export to Excel** (`js/xlsx-io.js`) opens the embedded template, writes each input
  value into its exact cell and leaves formulas, formatting, data validation, tables and
  comments untouched. It sets *full calculation on load*, so Excel refreshes every
  automatic column (IDs, ratings, statuses, dashboard) when the file is opened.
* **Import Excel** reads the same input cells from any filled copy of the toolkit, so an
  existing workbook can be continued online.
* Data is kept in the browser's local storage. Use *Download backup (.json)* to move it
  between computers.

## Files

* `js/schema.js` - every page, column and input cell (single source of truth for export/import)
* `js/engine.js` - JavaScript ports of the workbook formulas (live calculation on the site)
* `js/xlsx-io.js` - Excel export/import (uses `vendor/jszip.min.js`, MIT)
* `js/app.js`, `css/toolkit.css`, `index.html` - the user interface
* `js/template.js`, `js/template-data.js` - generated from the workbook

## Updating the template

Start from a filled or blank copy of the new workbook version. Make an empty template
(inputs cleared, cached results, hyperlinks and saved paths removed), then regenerate the
data files:

```
node tools/dump_inputs.js > /tmp/inputs.json
python3 tools/make_blank_template.py <workbook>.xlsx /tmp/inputs.json template/<name>.xlsx
python3 tools/build_template_data.py template/<name>.xlsx
```

If the new version moves or adds input cells, update `js/schema.js` to match.
