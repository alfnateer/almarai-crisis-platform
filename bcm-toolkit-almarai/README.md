# Almarai BCM Toolkit (web)

This is the Almarai edition of the toolkit website. The generic edition (template v5.1)
lives in `../bcm-toolkit/`.

**Import Excel** accepts two kinds of workbook:
* the Almarai BCM Toolkit Template v0.3
* v0.2, blank or filled. Business-unit copies such as
  `03. Bakery_Almarai_BCM_Toolkit_Template_v0.2_...` work, and any extra sheets of their own
  are ignored.

The embedded template has the larger register capacities of the Bakery copy (57 risks,
205 dependencies), so data from either v0.2 variant fits. To build a one-file version, run
`python3 tools/build_single_html.py "dist/Almarai BCM Toolkit.html"`.

The site replaces data entry in the Almarai BCM Toolkit Template **v0.3**
(`template/Almarai_BCM_Toolkit_Template_v0.3.xlsx`). It is built in two steps:
1. **v0.2.1** (`tools/fix_template.py`): a blank copy of the Bakery v0.2 workbook with its
   formula defects fixed, dropdown cells in yellow and red highlights on missing critical
   entries.
2. **v0.3** (`tools/cleanup_template.py`): Document Control removed (ISO 22301 7.5 now reads
   18. ISO Docs & Records), the always-blank BCM Risk Register columns T-W removed, and
   label/SharePoint metadata, the add-in, printer settings and unused named ranges stripped.
Every sheet of the workbook has a page; automatic columns are recalculated live, and
**Export to Excel** produces a populated copy of the original workbook.

Open `bcm-toolkit-almarai/index.html` (or `/bcm-toolkit-almarai/` on the deployed site). No build step,
no server code - it is a static site.

## How it maps to the workbook

| Sheet | Page | What you fill |
|---|---|---|
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
| 15 BCM Risk Register | live view | - (consolidated automatically, columns A-T) |
| 16 CAPA Register | table (100 rows) | findings through verified closure |
| 17 BCMS Progress (PDCA) | table | manual override, owner & next action |
| 18 ISO Docs & Records | table (20 docs) | status and evidence metadata |
| 19 ISO 22301 Map | table | owner, evidence, verification fields |
| 20 ERM RA Guidance | guidance page | read-only: risk appetite, likelihood / impact scales, heat map with live risk counts, impact criteria per category, controls, treatment |
| 21 Analysis | live view | - |
| 99 Engine | (internal lists) | used for dropdowns |

Row capacities are the workbook's own (e.g. 45 P&S, 40 BIA blocks, 57 risks), so every
row exported has a place in the template. Sheets hidden in the workbook (01, 17, 18, 19, 21)
are listed at the end of the menu.

**Paste from Excel:** select a block of cells in any Excel workbook, copy, click the first
target cell on a page and press Ctrl+V. Values fill across columns and down rows (rows are
added as needed), dropdown values are matched case-insensitively, dates such as
`15/01/2026`, `15-Jan-26` or `2026-01-15` are recognised, and automatic/label columns are skipped.

**Checks & Gaps:** the same rules as the template's red highlights run live on the site -
required entries left empty in a row that is in use, dropdown values not in their list,
RTO not shorter than MTPD, and S&S requirement flags set to Yes without a matching row in
12. Requirements Register (amber when rows exist but are not all Ready). The Checks & Gaps
page lists them with links to each cell.

**CAPA pre-population:** findings in the registers (requirement flags not captured,
overdue strategies, critical/significant risks without a strategy, failed or overdue
exercises, IT DR gaps, overdue BIA reviews, single points of failure, ISO documents past
review) are offered as CAPA rows on the CAPA page - pre-filled with source, reference,
finding, severity, owner, target date and status Open - and go into the exported workbook.

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
* `js/checks.js` - missing-entry checks and CAPA suggestions
* `js/xlsx-io.js` - Excel export/import (uses `vendor/jszip.min.js`, MIT)
* `js/app.js`, `css/toolkit.css`, `index.html` - the user interface
* `js/template.js`, `js/template-data.js` - generated from the workbook

## Updating the template

Start from a filled or blank copy of the new workbook version. Make an empty template
(inputs cleared, cached results, hyperlinks and saved paths removed), then regenerate the
data files:

```
node tools/dump_inputs.js > /tmp/inputs.json
python3 tools/make_blank_template.py <workbook>.xlsx /tmp/inputs.json template/source/<name>_blank.xlsx
python3 tools/fix_template.py template/source/<name>_blank.xlsx template/source/<name>_v0.2.1.xlsx
python3 tools/cleanup_template.py template/source/<name>_v0.2.1.xlsx template/<name>_v0.3.xlsx
python3 tools/build_template_data.py template/<name>_v0.3.xlsx
```

`fix_template.py` asserts on every formula it changes, so a new workbook version that
already differs fails loudly; adjust the script (or drop fixes that are no longer needed).

If the new version moves or adds input cells, update `js/schema.js` to match.
