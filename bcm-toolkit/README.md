# BCM Implementation Toolkit (web)

A website that replaces data entry in the **BCM Toolkit Template v5.1**
(`template/BCM_Toolkit_Template_v5.1.xlsx`). It is a generic (not company-specific) blank copy
of template v5, with v5's logic errors fixed, dropdown cells in yellow and red highlights on
missing critical entries. `tools/fix_template.py` lists every change.
Every data sheet of the workbook has a page. Automatic columns are recalculated live, and
**Export to Excel** produces a populated copy of the embedded workbook.

Open `bcm-toolkit/index.html` (or `/bcm-toolkit/` on the deployed site). It is a static
site with no build step and no server code. `tools/build_single_html.py` bundles everything,
including the template, into one HTML file that also works offline.

## How it maps to the workbook

| Sheet | Page | What you fill |
|---|---|---|
| 02 Dashboard | live view | nothing (all counters calculated) |
| 03 Business Unit Profile | form | sections 1-11 (identity, leadership, internal, external, climate, interested parties, scope, legal, assets, risk appetite grid, approval) |
| 04 Products & Services | table (45 rows) | BU, site, product, owner, scope, corporate BIA criticality / MBCO / MTPD |
| 05 Processes & Activities | table (207 rows) | activities per product code |
| 06 Application & Equipment | grid | application / service / equipment lists that feed dropdowns |
| 07 BIA | one block per in-scope P&S (40 blocks) | interruption Y/N, 1-5 impact scores over 9 horizons, BU critical, RTO, RPO, date |
| 08 Dependency Register | table (200 rows) | dependencies per BIA |
| 09 Risk Assumption Map | grid | assumptions per dependency category |
| 10 Risk Assessment | table (30 rows) | scope, assumption, ERM ref, controls, likelihood, impact, owner, assessment / review date, approver, evidence |
| 11 Strategies & Solutions | table (60 rows) | strategy, requirement flags, owner, dates |
| 12 Requirements Register | table (300 rows) | requirements per S&S |
| 13 IT DR Register | table (200 rows) | systems, RTO/RPO required vs capable, DR tests |
| 14 Recovery & Response | table (100 rows) | actions, plans, exercises |
| 15 BCM Risk Register | live view | nothing (consolidated automatically) |
| 16 CAPA Register | table (100 rows) | findings through verified closure |
| 03 rows 137+ (support lists) | Support Lists | dropdown values (named ranges) and duration units |

Pages that exist only on the website: **Checks & Gaps**, **How to use** (workflow, colours,
risk scoring, durations, import/export) and **Recovery Targets** (MBCO / MTPD / RTO / RPO per
in-scope product or service, with a check that RTO is within MTPD).

Row capacities are the workbook's own (for example 45 P&S, 40 BIA blocks, 30 risks), so
every exported row has a place in the template.

**Paste from Excel:** select a block of cells in any Excel workbook, copy, click the first
target cell on a page and press Ctrl+V. In tables the values fill across columns and down
rows (rows are added as needed). In forms and support lists they fill the section's fields
in order. Dropdown values are matched case-insensitively, dates such as `15/01/2026`,
`15-Jan-26` or `2026-01-15` are recognised, and automatic/label columns are skipped.

**Checks & Gaps:** the site runs the same rules as the template's red highlights, live:
* required entries left empty in a row that is in use
* dropdown values not in their list
* RTO not shorter than MTPD
* S&S requirement flags set to Yes with no matching row in 12. Requirements Register
  (amber when rows exist but are not all Ready)

The Checks & Gaps page lists them with a link to each cell.

**CAPA pre-population:** the site turns findings in the registers into suggested CAPA rows
on the CAPA page:
* requirement flags not captured
* overdue strategies
* critical/significant risks without a strategy
* failed or overdue exercises
* IT DR gaps
* overdue BIA reviews
* risk assessments past their review date
* single points of failure

Each suggestion is pre-filled with source, reference, finding, severity, owner, target date
and status Open, and goes into the exported workbook.

**Support lists:** the dropdown values live in 03. Business Unit Profile (rows 138+, one
named range per list). Edit them on the Support Lists page and the site's dropdowns and the
exported workbook both use the new values. Lists marked "used in formulas" are matched by
calculations.

## Export / import

* **Export to Excel** (`js/xlsx-io.js`) opens the embedded template and writes each input
  value into its exact cell. Formulas, formatting, data validation, conditional formatting,
  tables and comments are left untouched. It sets *full calculation on load*, so Excel
  refreshes every automatic column (IDs, ratings, statuses, register, dashboard) when the
  file is opened.
* **Import Excel** reads the same input cells from any filled copy of the toolkit.
  Support-list defaults are kept where the file has none. Files encrypted with a sensitivity
  label are detected, and the site explains how to save an unencrypted copy.
* **Download blank template** (under the ⋯ menu) gives the empty v5.1 workbook.
* Data is kept in the browser's local storage. Use *Download backup (.json)* to move it
  between computers.

## Files

* `js/schema.js`: every page, column, support list and input cell (single source of truth
  for export/import)
* `js/engine.js`: JavaScript ports of the workbook formulas (live calculation on the site)
* `js/checks.js`: missing-entry checks and CAPA suggestions
* `js/xlsx-io.js`: Excel export/import (uses `vendor/jszip.min.js`, MIT)
* `js/app.js`, `css/toolkit.css`, `index.html`: the user interface
* `js/template.js`, `js/template-data.js`: generated from the workbook
* `template/source/BCM_Toolkit_Template_v5_blank.xlsx`: v5 emptied, before fixes

## Updating the template

Start from a filled or blank copy of the new workbook version. First make an empty
template: inputs cleared; cached results, hyperlinks and saved paths removed; and company
traces dropped (sensitivity labels and footers, SharePoint metadata, add-ins, custom
fonts). Then apply the fixes and regenerate the data files:

```
node tools/dump_inputs.js > /tmp/inputs.json
python3 tools/make_blank_template.py <workbook>.xlsx /tmp/inputs.json template/source/<name>_blank.xlsx
python3 tools/fix_template.py template/source/<name>_blank.xlsx template/<name>.xlsx
python3 tools/build_template_data.py template/<name>.xlsx
python3 tools/build_single_html.py dist/BCM_Toolkit.html   # optional one-file version
```

`fix_template.py` asserts on every formula it changes. If a new workbook version already
differs, the script fails loudly; adjust it, or drop fixes that are no longer needed.

If the new version moves or adds input cells, update `js/schema.js` to match.
