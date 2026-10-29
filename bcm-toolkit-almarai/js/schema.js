/*
 * BCM Toolkit schema - mirrors the workbook "Almarai BCM Toolkit Template v0.2".
 *
 * Every input the website collects maps to exactly one cell in the workbook, so the
 * export can write it straight back into the original template. Columns marked
 * `auto` are formulas in Excel: the website computes them for display (engine.js)
 * and never writes them - Excel recalculates them when the exported file is opened.
 */
(function (root) {
  'use strict';

  // Column helpers -----------------------------------------------------------
  function colToNum(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }
  function numToCol(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function colRange(a, b) {
    var out = [];
    for (var i = colToNum(a); i <= colToNum(b); i++) out.push(numToCol(i));
    return out;
  }

  var S = {
    INS: '01. Instructions',
    DASH: '02. Dashboard',
    BUP: '03. Business Unit Profile',
    PS: '04. Products & Services',
    PROC: '05. Processes & Activities',
    APP: '06. Application & Equipment',
    BIA: '07. BIA',
    DEP: '08. Dependency Register',
    MAP: '09. Risk Assumption Map',
    RA: '10. Risk Assessment',
    SS: '11. Strategies & Solutions',
    REQ: '12. Requirements Register',
    DR: '13. IT DR Register',
    RR: '14. Recovery & Response',
    REG: '15. BCM Risk Register',
    CAPA: '16. CAPA Register',
    PDCA: '17. BCMS Progress (PDCA)',
    DOCS: '18. ISO Docs & Records',
    ISO: '19. ISO 22301 Map',
    ERM: '20. ERM RA Guidance',
    ANA: '21. Analysis',
    ENG: '99. Engine'
  };

  // BIA layout: 40 blocks of 7 rows starting at row 6.
  var BIA = {
    firstRow: 6, blockRows: 7, blocks: 40,
    scoreCols: colRange('J', 'R'),
    horizons: ['Up to 6 Hours', 'Up to 18 Hours', 'Up to 36 Hours', 'Up to 3 Days', 'Up to 1 Week',
      'Up to 2 Weeks', 'Up to 3 Weeks', 'Up to 1 month', 'More than 1 Month'],
    horizonHours: [6, 18, 36, 72, 168, 336, 504, 672],
    blockCols: ['V', 'Y', 'Z', 'AC', 'AE', 'AF'],
    top: function (k) { return 6 + 7 * (k - 1); }
  };

  /*
   * Sheet definitions. kind:
   *   form      - labelled single cells / small grids (sections)
   *   table     - one row per record, fixed row range
   *   bia       - 40 seven-row BIA blocks
   *   grid      - free grid of inputs (Risk Assumption Map, application lists)
   *   pdca, isomap, docs - tables with fixed rows and some inputs
   *   view      - fully calculated (Dashboard, Risk Register, Analysis)
   *   reference - read-only reference page (Instructions)
   *   erm       - ERM risk assessment guidance (read-only, from 20. ERM RA Guidance)
   */
  var sheets = [
    { id: 'guide', sheet: S.INS, nav: '01', title: 'Instructions', group: 'Hidden sheets in the workbook', kind: 'reference',
      iso: 'How the toolkit works - governance, colours & flow' },
    { id: 'dash', sheet: S.DASH, nav: '02', title: 'Dashboard', group: 'Governance', hiddenInWorkbook: true, kind: 'view',
      iso: 'ISO 22301: 9.1 Monitoring, measurement, analysis and evaluation' },
    { id: 'checks', sheet: 'Checks', nav: 'QA', title: 'Checks & Gaps', group: 'Governance', kind: 'checks',
      iso: 'Critical entries left empty, requirement flags not captured, and findings to raise as CAPA' },
    {
      id: 'bup', sheet: S.BUP, nav: '03', title: 'Business Unit Profile', group: 'Plan - Context & Scope',
      iso: 'ISO 22301 Clause 4 - Context of the organisation', kind: 'form',
      sections: [
        { title: 'Section 1 - Organization Identity', rowsFields: { col: 'B', rows: [5, 6, 7, 8, 9, 10, 11, 12, 13, 14], labelCol: 'A', type: 'textarea' } },
        { title: 'Section 2 - Leadership & Governance', grid: { rows: [18, 19, 20, 21, 22, 23, 24, 25], cols: [
          { c: 'A', label: 'Field / Role' }, { c: 'B', label: 'Name' }, { c: 'C', label: 'Title' }, { c: 'D', label: 'Contact' }] } },
        { title: 'Section 3 - Internal Context', rowsFields: { col: 'B', rows: [28, 29, 30, 31, 32, 33, 34, 35], labelCol: 'A', type: 'textarea' } },
        { title: 'Section 4 - External Context', rowsFields: { col: 'B', rows: [38, 39, 40, 41, 42, 43, 44, 45, 46], labelCol: 'A', type: 'textarea' } },
        { title: 'Section 5 - Climate Relevance for BCMS Context and Interested Parties', rowsFields: { col: 'B', rows: [50, 51, 52, 53, 54], labelCol: 'A', type: 'textarea' } },
        { title: 'Section 6 - Interested Parties & Requirements', grid: { rows: [58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68], labelCol: 'A', labelTitle: 'Interested Party', cols: [
          { c: 'B', label: 'Requirements / Expectations', type: 'textarea' }, { c: 'C', label: 'How Addressed', type: 'textarea' },
          { c: 'D', label: 'Priority', list: ['Critical', 'High', 'Medium', 'Low'] }] } },
        { title: 'Section 7 - BCMS Scope Definition', rowsFields: { col: 'B', rows: [71, 72, 73, 74, 75, 76], labelCol: 'A', type: 'textarea' } },
        { title: 'Section 8 - Legal & Regulatory Requirements', grid: { rows: [80, 81, 82, 83, 84, 85, 86, 87, 88, 89], cols: [
          { c: 'A', label: 'Requirement Type' }, { c: 'B', label: 'Specific Requirement', type: 'textarea' },
          { c: 'C', label: 'Compliance Status', list: ['Compliant', 'Partially Compliant', 'Non-Compliant', 'Not Assessed'] }, { c: 'D', label: 'Owner' }] } },
        { title: 'Section 9 - Critical Assets & Resources Overview', rowsFields: { col: 'B', rows: [92, 93, 94, 95, 96, 97, 98, 99, 100, 101], labelCol: 'A', type: 'textarea' } },
        { title: 'Section 10 - Risk Appetite & Tolerance (aligned with ERM)', fields: [
          { addr: 'B105', label: 'Risk appetite & tolerance (one merged cell in the workbook, covering Financial to Customer Impact)', type: 'textarea' }
        ] },
        { title: 'Section 11 - Profile Approval', grids: [
          { caption: 'Profile Completed By', rows: [117], cols: [{ c: 'A', label: 'Position / Title' }, { c: 'B', label: 'Name' }, { c: 'C', label: 'Date', type: 'date' }, { c: 'D', label: 'Signature' }] },
          { caption: 'Profile Reviewed By', rows: [121], cols: [{ c: 'A', label: 'Position / Title' }, { c: 'B', label: 'Name' }, { c: 'C', label: 'Date', type: 'date' }, { c: 'D', label: 'Signature' }] },
          { caption: 'Profile Approved By', rows: [125], cols: [{ c: 'A', label: 'Position / Title' }, { c: 'B', label: 'Name' }, { c: 'C', label: 'Date', type: 'date' }, { c: 'D', label: 'Signature' }] }
        ], fields: [
          { addr: 'B127', label: 'Completed Date', type: 'date' },
          { addr: 'D127', label: 'Next Review Date', auto: true }
        ] }
      ]
    },
    {
      id: 'ps', sheet: S.PS, nav: '04', title: 'Products & Services', group: 'Plan - Context & Scope',
      iso: 'Scope: Almarai Products & Services | ISO 22301: 4.3 Scope; 8.2.2 Prioritised activities', kind: 'table',
      first: 5, last: 49, numCol: 'A', usedBy: ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'],
      note: 'Steward issues P&S IDs. In-scope rows (In Scope = Yes) get a BIA order and become BIA blocks in 07. BIA, in the order listed here.',
      cols: [
        { c: 'B', label: 'P&S ID', auto: true, w: 80 },
        { c: 'C', label: 'Business Unit', w: 140, required: true },
        { c: 'D', label: 'Division / Site', w: 140, required: true },
        { c: 'E', label: 'Product / Service', w: 180, required: true },
        { c: 'F', label: 'Sub-Category', w: 130 },
        { c: 'G', label: 'Owner', w: 130, required: true },
        { c: 'H', label: 'EVP', w: 120, required: true },
        { c: 'I', label: 'In Scope (Y/N)', type: 'select', options: 'lstYN', w: 90 },
        { c: 'J', label: 'Corporate BIA Critical', type: 'select', options: 'lstCorpCrit', w: 120, group: 'Corporate BIA' },
        { c: 'K', label: 'Corporate MBCO (if provided)', type: 'textarea', w: 180, group: 'Corporate BIA' },
        { c: 'L', label: 'Corporate MTPD (if provided)', w: 120, group: 'Corporate BIA', placeholder: 'e.g. 3 days', duration: true },
        { c: 'M', label: 'Corporate MTPD (hours)', auto: true, w: 90, group: 'Corporate BIA' },
        { c: 'N', label: 'BIA Order', auto: true, w: 70 },
        { c: 'O', label: 'BU BIA MTPD', auto: true, w: 80 }
      ]
    },
    {
      id: 'proc', sheet: S.PROC, nav: '05', title: 'Processes & Activities', group: 'Plan - Context & Scope',
      iso: 'Value chain activities per product / service | ISO 22301: 8.2.2', kind: 'table',
      first: 5, last: 211, numCol: 'A', usedBy: ['B', 'D', 'E', 'F', 'G', 'H', 'I', 'L'],
      cols: [
        { c: 'B', label: 'Product Code', type: 'select', options: 'psIds', w: 100 },
        { c: 'C', label: 'Product & Services', auto: true, w: 160 },
        { c: 'D', label: 'Activities performed within the process', type: 'textarea', w: 240, requiredIf: 'B' },
        { c: 'E', label: 'Frequency of Activity', type: 'select', options: ['Daily', 'Weekly', 'Monthly', 'Yearly'], w: 110 },
        { c: 'F', label: 'Peak Period', w: 120 },
        { c: 'G', label: 'Rationale for Peak Period', type: 'textarea', w: 180 },
        { c: 'H', label: 'Current Staff to Perform Activity', type: 'number', w: 100, positiveIf: 'B' },
        { c: 'I', label: 'Activity Can be Performed Remotely?', type: 'select', options: ['Yes', 'No'], w: 110 },
        { c: 'J', label: 'BIA level MTPD in hours', auto: true, w: 90 },
        { c: 'K', label: 'BIA Level RTO in hours', auto: true, w: 90 },
        { c: 'L', label: 'Remark', type: 'textarea', w: 160 }
      ]
    },
    {
      id: 'app', sheet: S.APP, nav: '06', title: 'Application & Equipment', group: 'Plan - Context & Scope',
      iso: 'Step 3c - Application / equipment list | ISO 22301: 8.2.2 Dependencies & supporting resources', kind: 'grid',
      note: 'List the IT applications and services and the OT equipment / building / utilities. These lists feed the dropdowns in 08. Dependency Register and 13. IT DR Register.',
      rows: { first: 10, last: 34 }, numCol: 'F',
      cols: [
        { c: 'G', label: 'ICT - Application' },
        { c: 'H', label: 'ICT - Service' },
        { c: 'I', label: 'Equipment / Building / Utilities' }
      ]
    },
    {
      id: 'bia', sheet: S.BIA, nav: '07', title: 'Business Impact Analysis', group: 'Do - BIA & Risk',
      iso: 'Step 2 - BIA (product & service level) | ISO 22301: 8.2.2 Business impact analysis', kind: 'bia'
    },
    {
      id: 'dep', sheet: S.DEP, nav: '08', title: 'Dependency Register', group: 'Do - BIA & Risk',
      iso: 'Step 2b - one row per dependency | ISO 22301: 8.2.2 Dependencies & supporting resources', kind: 'table',
      first: 5, last: 209, numCol: 'A', usedBy: ['C', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
      cols: [
        { c: 'B', label: 'Dep ID', auto: true, w: 80 },
        { c: 'C', label: 'BIA # (select)', type: 'select', options: 'biaIds', w: 100 },
        { c: 'D', label: 'Product / Service (auto)', auto: true, w: 150 },
        { c: 'E', label: 'Product / Service ID', auto: true, w: 90 },
        { c: 'F', label: 'Category (select)', type: 'select', options: 'lstDepCat', w: 150 },
        { c: 'G', label: 'Sub category (ICT / Equipment)', type: 'select', options: 'depSubCat', w: 150 },
        { c: 'H', label: 'ICT / Equipment details', type: 'select', options: 'depDetail', w: 160, free: true },
        { c: 'I', label: 'Dependency Name', w: 160 },
        { c: 'J', label: 'Description / Detail', type: 'textarea', w: 200 },
        { c: 'K', label: 'SPOF? (Y/N)', type: 'select', options: 'lstSPOF', w: 80 },
        { c: 'L', label: 'Owner', w: 120 },
        { c: 'M', label: 'Remarks', type: 'textarea', w: 160 }
      ]
    },
    {
      id: 'map', sheet: S.MAP, nav: '09', title: 'Risk Assumption Map', group: 'Do - BIA & Risk',
      iso: "Step 3 - Risk assumptions per dependency category (per the organisation's ERM Framework)", kind: 'grid', keepDefaults: true,
      note: 'Risk assumptions offered in 10. Risk Assessment for each dependency category. Excel only offers the first five rows of each column in its dropdown; the website suggests all of them.',
      rows: { first: 5, last: 24 },
      cols: [
        { c: 'C', label: 'People' }, { c: 'D', label: 'ICT' }, { c: 'E', label: 'Data / Info' }, { c: 'F', label: '3rd Party' },
        { c: 'G', label: 'Equipment / Building / Utilities' }, { c: 'H', label: 'Pre-Service' }, { c: 'I', label: 'Post-Service' }
      ]
    },
    {
      id: 'ra', sheet: S.RA, nav: '10', title: 'Risk Assessment', group: 'Do - BIA & Risk',
      iso: "Step 3 - BCM risk assessment by dependency | ISO 22301: 8.2.3 Risk assessment", kind: 'table',
      first: 5, last: 61, numCol: 'A', usedBy: ['C', 'D', 'E', 'G', 'I', 'J', 'K', 'M', 'R'],
      note: 'Likelihood x Impact uses the ERM heat map in 20. ERM RA Guidance. A risk gets its Risk ID once a Risk Assumption is chosen.',
      cols: [
        { c: 'B', label: 'Risk ID', auto: true, w: 80 },
        { c: 'C', label: 'BU', type: 'select', options: 'businessUnits', w: 130 },
        { c: 'D', label: 'Division', type: 'select', options: 'divisionSites', w: 130 },
        { c: 'E', label: 'Dependency categories', type: 'select', options: 'lstDepCat', w: 150 },
        { c: 'F', label: 'Applicable BIA (P&S)', auto: true, w: 170 },
        { c: 'G', label: 'Risk Assumption (select or type)', type: 'select', options: 'riskAssumptions', free: true, w: 190 },
        { c: 'H', label: 'Business Disruption Scenario (auto)', auto: true, w: 220 },
        { c: 'I', label: 'ERM / Site Risk Register Ref#', type: 'select', options: ['No applicable ERM reference'], free: true, w: 150 },
        { c: 'J', label: 'Implemented Controls', type: 'textarea', w: 200 },
        { c: 'K', label: 'Residual Likelihood (1-5)', type: 'select', options: ['1', '2', '3', '4', '5'], numeric: true, w: 80 },
        { c: 'L', label: 'Likelihood Rating', auto: true, w: 100 },
        { c: 'M', label: 'Residual Impact (1-5)', type: 'select', options: ['1', '2', '3', '4', '5'], numeric: true, w: 80 },
        { c: 'N', label: 'Impact Rating', auto: true, w: 100 },
        { c: 'O', label: 'Risk Reference (ERM)', auto: true, w: 70 },
        { c: 'P', label: 'Risk Rating', auto: true, w: 100, badge: true },
        { c: 'Q', label: 'BCP Priority (1=Critical..4=Low)', auto: true, w: 80 },
        { c: 'R', label: 'Remarks', type: 'textarea', w: 160 },
        { c: 'S', label: 'S&S # (auto)', auto: true, w: 120 },
        { c: 'T', label: 'R&R IDs (auto)', auto: true, w: 120 },
        { c: 'U', label: 'Risks (New / Existing)', auto: true, w: 110, badge: true }
      ]
    },
    {
      id: 'ss', sheet: S.SS, nav: '11', title: 'Strategies & Solutions', group: 'Do - Strategies & Plans',
      iso: 'Step 4 - Continuity strategies & solutions | ISO 22301: 8.3', kind: 'table',
      first: 5, last: 64, numCol: 'A', usedBy: ['B', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'W'],
      note: 'Requirement flags (People ... Fund) set to Yes need matching rows in 12. Requirements Register with Readiness = Ready before the strategy can be Completed.',
      cols: [
        { c: 'B', label: 'Risk ID', type: 'select', options: 'riskIds', w: 90 },
        { c: 'C', label: 'S&S ID', auto: true, w: 80 },
        { c: 'D', label: 'Dependency', auto: true, w: 120 },
        { c: 'E', label: 'Risk Assumption', auto: true, w: 160 },
        { c: 'F', label: 'Disruption Scenario', auto: true, w: 200 },
        { c: 'G', label: 'Products / Services', auto: true, w: 160 },
        { c: 'H', label: 'Strategy', type: 'select', options: 'lstStrategy', w: 200 },
        { c: 'I', label: 'Strategy Details / Do nothing Justification', type: 'textarea', w: 220 },
        { c: 'J', label: 'People', type: 'select', options: 'lstYN', w: 70, group: 'Requirements' },
        { c: 'K', label: 'ICT', type: 'select', options: 'lstYN', w: 70, group: 'Requirements' },
        { c: 'L', label: '3rd Party', type: 'select', options: 'lstYN', w: 70, group: 'Requirements' },
        { c: 'M', label: 'Data Backup', type: 'select', options: 'lstYN', w: 70, group: 'Requirements' },
        { c: 'N', label: 'Equipment / Building / Utilities', type: 'select', options: 'lstYN', w: 90, group: 'Requirements' },
        { c: 'O', label: 'Fund', type: 'select', options: 'lstYN', w: 70, group: 'Requirements' },
        { c: 'P', label: 'Action Owner', w: 120 },
        { c: 'Q', label: 'Start Date', type: 'date', w: 130 },
        { c: 'R', label: 'Due Date', type: 'date', w: 130 },
        { c: 'S', label: 'Completion Date', type: 'date', w: 130 },
        { c: 'T', label: 'Status', auto: true, w: 100, badge: true },
        { c: 'U', label: '1:1 Check', auto: true, w: 90, badge: true },
        { c: 'V', label: 'R&R IDs (auto)', auto: true, w: 110 },
        { c: 'W', label: 'Remarks', type: 'textarea', w: 160 }
      ]
    },
    {
      id: 'req', sheet: S.REQ, nav: '12', title: 'Requirements Register', group: 'Do - Strategies & Plans',
      iso: 'Step 4b - one row per continuity requirement | ISO 22301: 8.3', kind: 'table',
      first: 5, last: 304, numCol: 'A', usedBy: ['C', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'],
      cols: [
        { c: 'B', label: 'Req ID', auto: true, w: 85 },
        { c: 'C', label: 'S&S ID', type: 'select', options: 'ssIds', w: 90 },
        { c: 'D', label: 'Risk ID (auto)', auto: true, w: 80 },
        { c: 'E', label: 'Category', type: 'select', options: 'lstReqCat', w: 120 },
        { c: 'F', label: 'Item / Resource / Role', w: 170 },
        { c: 'G', label: 'Detail / Specification', type: 'textarea', w: 200 },
        { c: 'H', label: 'Primary Location', w: 130 },
        { c: 'I', label: 'Alternate Location(s)', w: 130 },
        { c: 'J', label: 'Supporting Documents', w: 140 },
        { c: 'K', label: 'Owner', w: 120 },
        { c: 'L', label: 'Readiness', type: 'select', options: ['Ready', 'In Progress'], w: 100 }
      ]
    },
    {
      id: 'dr', sheet: S.DR, nav: '13', title: 'IT DR Register', group: 'Do - Strategies & Plans',
      iso: 'One row per critical system - Required vs DR-capable RTO/RPO | ISO 22301: 8.3.2 ICT readiness (ISO/IEC 27031)', kind: 'table',
      first: 5, last: 204, numCol: 'A', usedBy: ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'R'],
      cols: [
        { c: 'B', label: 'DR ID', auto: true, w: 80 },
        { c: 'C', label: 'Dependency category', type: 'select', options: 'drCat', w: 150 },
        { c: 'D', label: 'Dependency area', type: 'select', options: 'drArea', w: 140 },
        { c: 'E', label: 'Dependent System / Application', type: 'select', options: 'drSystem', free: true, w: 170 },
        { c: 'F', label: 'Business P&S supported', w: 160 },
        { c: 'G', label: 'Required RTO in hours', type: 'number', w: 90 },
        { c: 'H', label: 'Required RPO in hours', type: 'number', w: 90 },
        { c: 'I', label: 'DR RTO capable (hrs)', type: 'number', w: 90 },
        { c: 'J', label: 'DR RPO capable (hrs)', type: 'number', w: 90 },
        { c: 'K', label: 'DR Strategy', type: 'select', options: 'lstDRStrategy', w: 150 },
        { c: 'L', label: 'DR Site / Location', w: 130 },
        { c: 'M', label: 'Last DR Test', type: 'date', w: 130, notFuture: true },
        { c: 'N', label: 'Test Scenario', type: 'textarea', w: 170 },
        { c: 'O', label: 'Result', type: 'select', options: 'lstTestResult', w: 130 },
        { c: 'P', label: 'Next Test Due', auto: true, w: 100 },
        { c: 'Q', label: 'Status (auto)', auto: true, w: 150, badge: true },
        { c: 'R', label: 'Gaps / Remarks', type: 'textarea', w: 170 },
        { c: 'S', label: 'RTO Compliance', auto: true, w: 120, badge: true },
        { c: 'T', label: 'RPO Compliance', auto: true, w: 120, badge: true }
      ]
    },
    {
      id: 'rr', sheet: S.RR, nav: '14', title: 'Recovery & Response', group: 'Do - Strategies & Plans',
      iso: 'Step 5 - Recovery & Response plan | ISO 22301: 8.4 Plans & procedures; 8.5 Exercise programme', kind: 'table',
      first: 5, last: 104, numCol: 'A', usedBy: ['C', 'E', 'F', 'L', 'M', 'N', 'R'],
      cols: [
        { c: 'B', label: 'R&R ID', auto: true, w: 80 },
        { c: 'C', label: 'S&S ID', type: 'select', options: 'ssIds', w: 90 },
        { c: 'D', label: 'Product / Service', auto: true, w: 150 },
        { c: 'E', label: 'Response & Recovery Actions', type: 'textarea', w: 240, group: 'Immediate Response' },
        { c: 'F', label: 'Supporting Plans / Procedures', type: 'textarea', w: 180, group: 'Immediate Response' },
        { c: 'G', label: 'Dependency', auto: true, w: 110, group: 'Risk Context' },
        { c: 'H', label: 'Risk Assumption', auto: true, w: 150, group: 'Risk Context' },
        { c: 'I', label: 'Disruption Scenario', auto: true, w: 180, group: 'Risk Context' },
        { c: 'J', label: 'Strategy', auto: true, w: 150, group: 'Strategy & Resources' },
        { c: 'K', label: 'Solution Details', auto: true, w: 180, group: 'Strategy & Resources' },
        { c: 'L', label: 'Planned Exercise Date', type: 'date', w: 130, group: 'Exercise' },
        { c: 'M', label: 'Actual Exercise Date', type: 'date', w: 130, group: 'Exercise' },
        { c: 'N', label: 'Exercise Result', type: 'select', options: ['Pass', 'Pass with Findings', 'Fail'], w: 130, group: 'Exercise' },
        { c: 'O', label: 'Next Exercise Due', auto: true, w: 100, group: 'Exercise' },
        { c: 'P', label: 'Exercise Status', auto: true, w: 130, badge: true, group: 'Exercise' },
        { c: 'Q', label: 'Duplicate Check', auto: true, w: 90, badge: true },
        { c: 'R', label: 'Remarks', type: 'textarea', w: 160 }
      ]
    },
    { id: 'reg', sheet: S.REG, nav: '15', title: 'BCM Risk Register', group: 'Check - Monitor & Improve', kind: 'view',
      iso: 'Step 6 - Consolidated register (auto view for ERM & audit) | ISO 22301: 9.1' },
    {
      id: 'capa', sheet: S.CAPA, nav: '16', title: 'CAPA Register', group: 'Check - Monitor & Improve',
      iso: 'ISO 22301 clause 10.1 - findings, root cause, corrective action, verification and effectiveness', kind: 'table',
      first: 6, last: 105, numCol: 'A', usedBy: ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'U', 'V', 'W', 'X'],
      note: 'Open findings from risks, exercises, IT DR, audits and documented-information gaps are tracked here through verified closure. Closure requires evidence and an effectiveness decision.',
      cols: [
        { c: 'B', label: 'CAPA ID', auto: true, w: 90 },
        { c: 'C', label: 'Source Type', type: 'select', options: ['Risk', 'Recovery Plan', 'IT DR', 'ISO Document', 'Audit', 'Management Review', 'BIA', 'Dependency', 'Other'], w: 130 },
        { c: 'D', label: 'Source Ref', type: 'select', options: 'capaSourceRef', free: true, w: 110, requiredIf: 'F' },
        { c: 'E', label: 'Finding Date', type: 'date', w: 130, notFuture: true },
        { c: 'F', label: 'Finding', type: 'textarea', w: 220 },
        { c: 'G', label: 'Severity', type: 'select', options: ['Critical', 'High', 'Medium', 'Low'], w: 90 },
        { c: 'H', label: 'Immediate Correction', type: 'textarea', w: 170 },
        { c: 'I', label: 'Root Cause', type: 'textarea', w: 170 },
        { c: 'J', label: 'Corrective Action', type: 'textarea', w: 170 },
        { c: 'K', label: 'Owner', w: 120 },
        { c: 'L', label: 'Target Date', type: 'date', w: 130, notBefore: 'E' },
        { c: 'M', label: 'Status', type: 'select', options: ['Open', 'Investigation', 'Action in Progress', 'Pending Verification', 'Closed', 'Cancelled'], w: 150 },
        { c: 'N', label: 'Evidence Ref', w: 120 },
        { c: 'O', label: 'Verification Owner', w: 120 },
        { c: 'P', label: 'Verification Date', type: 'date', w: 130, notBefore: 'E' },
        { c: 'Q', label: 'Effectiveness', type: 'select', options: ['Effective', 'Partially Effective', 'Ineffective', 'Pending'], w: 130 },
        { c: 'R', label: 'Closure Date', type: 'date', w: 130, notBefore: 'P' },
        { c: 'S', label: 'Days Open', auto: true, w: 70 },
        { c: 'T', label: 'Overdue', auto: true, w: 90, badge: true },
        { c: 'U', label: 'Risk ID', type: 'select', options: 'riskIds', w: 90 },
        { c: 'V', label: 'S&S ID', type: 'select', options: 'ssIds', w: 90 },
        { c: 'W', label: 'R&R ID', type: 'select', options: 'rrIds', w: 90 },
        { c: 'X', label: 'Remarks', type: 'textarea', w: 160 }
      ]
    },
    {
      id: 'pdca', sheet: S.PDCA, nav: '17', title: 'BCMS Progress (PDCA)', group: 'Hidden sheets in the workbook',
      iso: 'ISO 22301: 4.4 BCMS; 10.2 Continual improvement', kind: 'pdca', keepDefaults: true,
      rows: { first: 5, last: 19 }
    },
    {
      id: 'docs', sheet: S.DOCS, nav: '18', title: 'ISO Docs & Records', group: 'Hidden sheets in the workbook',
      iso: 'Documents & records the toolkit does not generate - the certification gap register', kind: 'table', fixedRows: true, keepDefaults: true,
      first: 5, last: 24, usedBy: [],
      cols: [
        { c: 'A', label: 'Ref', fixed: true, w: 70 },
        { c: 'B', label: 'ISO Clause', fixed: true, w: 60 },
        { c: 'C', label: 'Document / Record', fixed: true, w: 190 },
        { c: 'D', label: 'What it must contain', fixed: true, w: 240 },
        { c: 'E', label: 'Owner', w: 120 },
        { c: 'F', label: 'Status', type: 'select', options: ['Not Started', 'Drafting', 'Under Review', 'Approved', 'Evidence Filed', 'Not Applicable'], w: 130 },
        { c: 'G', label: 'Target Date', type: 'date', w: 130 },
        { c: 'H', label: 'Location / Link', w: 150 },
        { c: 'I', label: 'Frequency', w: 120 },
        { c: 'J', label: 'Remarks', type: 'textarea', w: 160 },
        { c: 'K', label: 'Document Version', w: 90 },
        { c: 'L', label: 'Approval Date', type: 'date', w: 130 },
        { c: 'M', label: 'Review / Next Review Date', type: 'date', w: 130 },
        { c: 'N', label: 'Approver', w: 120 },
        { c: 'O', label: 'Controlled Repository Reference', w: 150 },
        { c: 'P', label: 'Evidence Verification Owner', w: 130 },
        { c: 'Q', label: 'Evidence Verification Date', type: 'date', w: 130 },
        { c: 'R', label: 'Confidentiality / Access Classification', type: 'select', options: ['Public', 'Internal', 'Confidential', 'Restricted'], w: 130 },
        { c: 'S', label: 'Expiry / Next-Review Alert', auto: true, w: 130, badge: true }
      ]
    },
    {
      id: 'iso', sheet: S.ISO, nav: '19', title: 'ISO 22301 Map', group: 'Hidden sheets in the workbook',
      iso: 'ISO 22301:2019 full clause map - status pulls from PDCA and ISO Docs where wired', kind: 'isomap', keepDefaults: true,
      rows: { first: 5, last: 43 },
      inputCols: ['E', 'F', 'G', 'J', 'K', 'L', 'M', 'N', 'O'],
      cols: [
        { c: 'A', label: 'Clause', fixed: true, w: 60 },
        { c: 'B', label: 'Requirement (summary)', fixed: true, w: 220 },
        { c: 'C', label: 'Where it lives', fixed: true, w: 240 },
        { c: 'D', label: 'Status', auto: true, w: 100, badge: true },
        { c: 'E', label: 'Owner', w: 110 },
        { c: 'F', label: 'Evidence / Document', w: 160 },
        { c: 'G', label: 'Notes / Gap Action', type: 'textarea', w: 200 },
        { c: 'H', label: 'Navigate to', fixed: true, w: 150 },
        { c: 'I', label: 'Objective Completion Criterion', fixed: true, w: 240 },
        { c: 'J', label: 'Internal Evidence ID(s)', w: 150 },
        { c: 'K', label: 'External Evidence ID(s)', w: 110 },
        { c: 'L', label: 'Last Verified Date', type: 'date', w: 130 },
        { c: 'M', label: 'Assessor', w: 110 },
        { c: 'N', label: 'Manual-Override Justification', type: 'textarea', w: 170 },
        { c: 'O', label: 'Linked Finding / CAPA', w: 110 }
      ]
    },
    { id: 'erm', sheet: S.ERM, nav: '20', title: 'ERM RA Guidance', group: 'Reference', kind: 'erm',
      iso: 'ISO 22301: 8.2.3 Risk criteria - Almarai ERM framework: appetite, likelihood, impact, heat map and treatment' },
    { id: 'ana', sheet: S.ANA, nav: '21', title: 'Analysis', group: 'Hidden sheets in the workbook', kind: 'view',
      iso: 'In-scope products & services with corporate vs BIA recovery targets' }
  ];

  // PDCA manual override and next action are inputs.
  var PDCA_OVERRIDE = ['Yes', 'Partial', 'No', 'Not Assessed'];

  /*
   * Enumerate every input cell -> {sheet, addr, type, numeric}. This is the single
   * source of truth for export (write) and import (read).
   */
  function inputCells(tpl) {
    var out = [];
    function add(sheet, addr, type, numeric) { out.push({ sheet: sheet, addr: addr, type: type || 'text', numeric: !!numeric }); }
    sheets.forEach(function (sh) {
      if (sh.kind === 'form') {
        sh.sections.forEach(function (sec) {
          (sec.fields || []).forEach(function (f) { if (!f.auto) add(sh.sheet, f.addr, f.type); });
          if (sec.rowsFields) sec.rowsFields.rows.forEach(function (r) { add(sh.sheet, sec.rowsFields.col + r, sec.rowsFields.type); });
          var grids = sec.grids || (sec.grid ? [sec.grid] : []);
          grids.forEach(function (g) {
            g.rows.forEach(function (r) {
              g.cols.forEach(function (c) {
                if (c.autoRows && c.autoRows.indexOf(r) >= 0) return;
                add(sh.sheet, c.c + r, c.type);
              });
            });
          });
        });
      } else if (sh.kind === 'table') {
        for (var r = sh.first; r <= sh.last; r++) {
          sh.cols.forEach(function (c) {
            if (c.auto || c.fixed) return;
            add(sh.sheet, c.c + r, c.type === 'date' ? 'date' : (c.type === 'number' ? 'number' : 'text'), c.numeric || c.type === 'number');
          });
        }
      } else if (sh.kind === 'grid') {
        for (var g = sh.rows.first; g <= sh.rows.last; g++) sh.cols.forEach(function (c) { add(sh.sheet, c.c + g); });
      } else if (sh.kind === 'bia') {
        for (var k = 1; k <= BIA.blocks; k++) {
          var t = BIA.top(k);
          for (var i = 0; i < BIA.blockRows; i++) {
            add(sh.sheet, 'H' + (t + i));
            BIA.scoreCols.forEach(function (c) { add(sh.sheet, c + (t + i), 'text', true); });
          }
          add(sh.sheet, 'V' + t); add(sh.sheet, 'Y' + t); add(sh.sheet, 'Z' + t);
          add(sh.sheet, 'AC' + t); add(sh.sheet, 'AE' + t, 'number', true); add(sh.sheet, 'AF' + t, 'date');
        }
      } else if (sh.kind === 'pdca') {
        for (var p = sh.rows.first; p <= sh.rows.last; p++) { add(sh.sheet, 'E' + p); add(sh.sheet, 'H' + p); }
      } else if (sh.kind === 'isomap') {
        var formulas = (tpl && tpl.formulas && tpl.formulas[sh.sheet]) || [];
        var fset = {};
        formulas.forEach(function (a) { fset[a] = true; });
        var sectionRows = isoSectionRows(tpl);
        for (var q = sh.rows.first; q <= sh.rows.last; q++) {
          if (sectionRows[q]) continue;
          sh.inputCols.forEach(function (c) {
            if (fset[c + q]) return;
            var col = sh.cols.filter(function (x) { return x.c === c; })[0];
            add(sh.sheet, c + q, col && col.type === 'date' ? 'date' : 'text');
          });
        }
      }
    });
    return out;
  }

  // Clause section header rows in the ISO map (e.g. "4 Context of the organization").
  function isoSectionRows(tpl) {
    var rows = {};
    var st = (tpl && tpl.static && tpl.static[S.ISO]) || {};
    for (var r = 5; r <= 43; r++) {
      var a = st['A' + r];
      if (a && typeof a === 'string' && /^\d+ [A-Z]/.test(a)) rows[r] = a;
    }
    return rows;
  }

  var SCHEMA = {
    S: S, BIA: BIA, sheets: sheets, PDCA_OVERRIDE: PDCA_OVERRIDE,
    inputCells: inputCells, isoSectionRows: isoSectionRows,
    colToNum: colToNum, numToCol: numToCol, colRange: colRange,
    byId: function (id) { return sheets.filter(function (s) { return s.id === id; })[0]; }
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = SCHEMA;
  else root.BCM_SCHEMA = SCHEMA;
})(this);
