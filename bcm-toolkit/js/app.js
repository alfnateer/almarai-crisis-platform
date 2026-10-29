/*
 * BCM Implementation Toolkit - web front end.
 * Renders every sheet of the toolkit as a page, keeps the data in the browser
 * (localStorage), recalculates the automatic columns live and exports a populated
 * copy of the original Excel workbook.
 */
(function () {
  'use strict';

  var T = window.BCM_TEMPLATE_DATA;
  var SC = window.BCM_SCHEMA;
  var EN = window.BCM_ENGINE;
  var IO = window.BCM_XLSX;
  var CK = window.BCM_CHECKS;
  var S = SC.S;
  var STORE_KEY = 'bcm-toolkit-v5.1';

  var INPUTS = SC.inputCells(T);
  var SPEC = {};
  INPUTS.forEach(function (c) { SPEC[c.sheet + '!' + c.addr] = c; });
  var SHEET_DEF = {};
  SC.sheets.forEach(function (d) { SHEET_DEF[d.sheet] = d; });

  // ------------------------------------------------------------------ state
  function freshState() {
    var values = {};
    INPUTS.forEach(function (c) {
      if (!c.keep) return;
      var v = (T.static[c.sheet] || {})[c.addr];
      if (v === undefined || v === null || v === '') return;
      (values[c.sheet] = values[c.sheet] || {})[c.addr] = v;
    });
    return { app: 'bcm-toolkit', template: T.source, version: 1, values: values, ui: { rows: {} }, updated: null };
  }
  function loadState() {
    try {
      var raw = localStorage.getItem(STORE_KEY);
      if (raw) {
        var s = JSON.parse(raw);
        if (s && s.values) { s.ui = s.ui || { rows: {} }; s.ui.rows = s.ui.rows || {}; return s; }
      }
    } catch (e) { /* storage unavailable - start fresh */ }
    return freshState();
  }
  var state = loadState();
  var saveTimer = null;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      state.updated = new Date().toISOString();
      try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); setSaved('Saved in this browser'); }
      catch (e) { setSaved('Not saved - browser storage unavailable. Use Backup.'); }
    }, 250);
  }
  function setSaved(msg) { var el = document.getElementById('save-status'); if (el) el.textContent = msg; }

  var book = {
    get: function (sheet, addr) {
      var key = sheet + '!' + addr;
      if (SPEC[key]) { var sv = state.values[sheet]; return sv && sv[addr] !== undefined ? sv[addr] : ''; }
      var st = T.static[sheet];
      return st && st[addr] !== undefined ? st[addr] : '';
    }
  };
  function setValue(sheet, addr, v) {
    var spec = SPEC[sheet + '!' + addr];
    if (!spec) return;
    if (typeof v === 'string') {
      if (spec.numeric && v.trim() !== '' && !isNaN(Number(v))) v = Number(v);
      else if (spec.type === 'number' && v.trim() !== '' && !isNaN(Number(v))) v = Number(v);
    }
    var sv = state.values[sheet] = state.values[sheet] || {};
    if (v === '' || v === null || v === undefined) delete sv[addr];
    else sv[addr] = v;
    save();
    scheduleCalc();
  }

  var calc, issues = [], issueAt = {};
  // Recalculate everything and re-run the data-quality checks.
  function recompute() {
    calc = EN.compute(book, { tpl: T });
    issues = CK.missing(val).concat(offListIssues());
    issueAt = {};
    issues.forEach(function (x) { var k = x.sheet + '|' + x.addr; if (!issueAt[k] || x.level === 'red') issueAt[k] = x; });
  }
  // Dropdown cells holding a value that is not in their list (e.g. typed or imported text).
  function offListIssues() {
    var out = [];
    SC.sheets.forEach(function (def) {
      if (def.kind !== 'table' || !def.usedBy) return;
      var cols = def.cols.filter(function (c) { return !c.auto && !c.fixed && c.options && !c.free; });
      if (!cols.length) return;
      for (var r = def.first; r <= def.last; r++) {
        cols.forEach(function (c) {
          var v = book.get(def.sheet, c.c + r);
          if (v === '') return;
          var opts = optionsFor(c.options, def.sheet, r).map(String);
          if (opts.length && opts.indexOf(String(v)) < 0) {
            var idc = def.cols[0] && def.cols[0].auto ? def.cols[0].c : null;
            out.push({ id: def.id, sheet: def.sheet, addr: c.c + r, row: r, level: 'red', ref: (idc && val(def.sheet, idc + r)) || 'row ' + r,
              what: c.label + ': "' + v + '" is not in the dropdown list' });
          }
        });
      }
    });
    return out;
  }
  function val(sheet, addr) {
    var c = calc[sheet];
    if (c && addr in c) return c[addr];
    return book.get(sheet, addr);
  }
  var calcTimer = null;
  function scheduleCalc() {
    clearTimeout(calcTimer);
    calcTimer = setTimeout(function () { recompute(); refreshCalc(); }, 120);
  }

  // ------------------------------------------------------------------ utils
  function esc(s) {
    return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function fmt(v) {
    if (v === undefined || v === null) return '';
    if (typeof v === 'number') return Math.abs(v - Math.round(v)) < 1e-9 ? String(Math.round(v)) : String(Math.round(v * 100) / 100);
    return String(v);
  }
  var RED = ['Critical', 'Overdue', 'No - review', 'Red - Action Required', 'DUPLICATE', 'Failed', 'GAP - no DR', 'Missing', 'No', 'Incomplete',
    'RTO Not Complied', 'RPO Not Complied', 'Test Overdue', 'Retest Required', 'Recovery Not Compliant', 'Not Evidenced', 'Ineffective', '#VALUE!', '#N/A', '#REF!'];
  var AMBER = ['Significant', 'Partial', 'Amber - Attention', 'In Progress', 'Incomplete Exercise', 'Result Required', 'Pass with Findings',
    'Not Tested', 'Recovery Targets Incomplete', 'Review Date Missing', 'Due \u226430 Days', 'RTO not set', 'Moderate', 'Planned', 'Scheduled', 'Not Scheduled', 'Under Evaluation'];
  var GREEN = ['Low', 'Yes', 'Complete', 'Completed', 'Passed', 'OK', 'On Track', 'Green - On Track', 'Tested & Current', 'RTO Complied',
    'RPO Complied', 'Current', 'In Place', 'MONITOR', 'N/A'];
  function tone(v) {
    v = String(v);
    if (v === '') return 'none';
    if (/^ESCALATE/.test(v)) return 'red';
    if (/^REVIEW/.test(v)) return 'amber';
    if (RED.indexOf(v) >= 0) return 'red';
    if (AMBER.indexOf(v) >= 0) return 'amber';
    if (GREEN.indexOf(v) >= 0) return 'green';
    return 'grey';
  }
  function badge(v) { return v === '' || v === undefined ? '' : '<span class="badge ' + tone(v) + '">' + esc(fmt(v)) + '</span>'; }
  function toast(msg, kind) {
    var el = document.getElementById('toast');
    el.textContent = msg;
    el.className = 'toast show ' + (kind || '');
    clearTimeout(el._t);
    el._t = setTimeout(function () { el.className = 'toast'; }, kind === 'err' ? 15000 : 4200);
  }
  // Support lists live in 03. Business Unit Profile (rows 138+), so edits on the
  // Support Lists page change the dropdowns here and in the exported workbook.
  var LIST_DEF = {};
  SC.LISTS.forEach(function (l) { LIST_DEF[l.name] = l; });
  function listValues(name) {
    var l = LIST_DEF[name];
    if (!l) return null;
    var out = [];
    for (var r = l.first; r <= l.last; r++) { var v = book.get(S.BUP, l.col + r); if (v !== '') out.push(String(v)); }
    return out;
  }
  function nonEmpty(sheet, col, a, b) {
    var out = [];
    for (var r = a; r <= b; r++) { var v = book.get(sheet, col + r); if (v !== '') out.push(String(v)); }
    return out;
  }
  // Options for a dropdown (static list name, array, or dynamic source by row).
  function optionsFor(src, sheet, row) {
    if (Array.isArray(src)) return src;
    var lv = listValues(src);
    if (lv) return lv;
    var L = calc.lists;
    switch (src) {
      case 'psIds': case 'biaIds': case 'riskIds': case 'ssIds': case 'rrIds': case 'businessUnits': case 'divisionSites':
        return L[src];
      case 'depSubCat':
        var f = book.get(S.DEP, 'F' + row);
        return f === 'ICT' ? ['Application', 'Service'] : (f === 'Equipment / Building / Utilities' ? ['Equipment / Building / Utilities'] : []);
      case 'depDetail':
        return assetList(book.get(S.DEP, 'F' + row), book.get(S.DEP, 'G' + row));
      case 'drCat':
        return nonEmpty(S.APP, 'A', 10, 11);
      case 'drArea':
        var c = book.get(S.DR, 'C' + row);
        return c === 'ICT' ? ['Application', 'Service'] : (c === 'Equipment / Building / Utilities' ? ['Equipment / Building / Utilities'] : []);
      case 'drSystem':
        return assetList(book.get(S.DR, 'C' + row), book.get(S.DR, 'D' + row));
      case 'riskAssumptions':
        var cat = book.get(S.RA, 'E' + row);
        // Column of 09. Risk Assumption Map whose header (row 4) is the category.
        for (var mc = 3; mc <= 9; mc++) {
          var hd = book.get(S.MAP, SC.numToCol(mc) + '4');
          if (cat !== '' && String(hd).toLowerCase() === String(cat).toLowerCase()) return nonEmpty(S.MAP, SC.numToCol(mc), 5, 24);
        }
        return [];
      case 'capaSourceRef':
        var t = book.get(S.CAPA, 'C' + row);
        return { Risk: L.riskIds, 'Recovery Plan': L.rrIds, 'IT DR': L.drIds, BIA: L.biaIds, Dependency: L.depIds }[t] || [];
    }
    return [];
  }
  function assetList(cat, sub) {
    if (cat === 'ICT' && sub === 'Application') return nonEmpty(S.APP, 'G', 10, 34);
    if (cat === 'ICT' && sub === 'Service') return nonEmpty(S.APP, 'H', 10, 34);
    if (cat === 'Equipment / Building / Utilities' && sub === 'Equipment / Building / Utilities') return nonEmpty(S.APP, 'I', 10, 34);
    return [];
  }

  // ------------------------------------------------------------------ input widgets
  function key(sheet, addr) { return esc(sheet + '|' + addr); }
  function inputHtml(sheet, addr, col, row, extraClass) {
    col = col || {};
    var v = book.get(sheet, addr);
    var k = key(sheet, addr);
    var cls = 'in' + (extraClass ? ' ' + extraClass : '');
    var ph = col.placeholder ? ' placeholder="' + esc(col.placeholder) + '"' : '';
    var type = col.type || 'text';
    if (type === 'select' || col.options) {
      if (col.free) {
        return '<input class="' + cls + '" data-k="' + k + '" data-opt="' + esc(JSON.stringify(col.options)) + '" data-row="' + row + '" value="' + esc(fmt(v)) + '"' + ph + ' autocomplete="off">';
      }
      return '<select class="' + cls + '" data-k="' + k + '" data-opt="' + esc(JSON.stringify(col.options)) + '" data-row="' + row + '">' +
        selectOptions(optionsFor(col.options, sheet, row), v) + '</select>';
    }
    if (col.list) {
      return '<input class="' + cls + '" data-k="' + k + '" data-opt="' + esc(JSON.stringify(col.list)) + '" data-row="' + row + '" value="' + esc(fmt(v)) + '" autocomplete="off">';
    }
    if (type === 'date') return '<input type="date" class="' + cls + '" data-k="' + k + '" value="' + esc(v) + '">';
    if (type === 'number') return '<input type="number" step="any" class="' + cls + '" data-k="' + k + '" value="' + esc(fmt(v)) + '"' + ph + '>';
    if (type === 'textarea') return '<textarea class="' + cls + '" data-k="' + k + '" rows="1"' + ph + '>' + esc(v) + '</textarea>';
    return '<input class="' + cls + '" data-k="' + k + '" value="' + esc(fmt(v)) + '"' + ph + '>';
  }
  function selectOptions(opts, v) {
    var sv = fmt(v);
    var html = '<option value=""></option>';
    var found = false;
    opts.forEach(function (o) {
      var sel = String(o) === sv;
      if (sel) found = true;
      html += '<option' + (sel ? ' selected' : '') + ' value="' + esc(o) + '">' + esc(o) + '</option>';
    });
    if (sv !== '' && !found) html += '<option selected value="' + esc(sv) + '">' + esc(sv) + ' (not in list)</option>';
    return html;
  }
  function calcCell(sheet, addr, isBadge) {
    var v = val(sheet, addr);
    return '<td class="auto" data-calc="' + key(sheet, addr) + '"' + (isBadge ? ' data-badge="1"' : '') + '>' + (isBadge ? badge(v) : esc(fmt(v))) + '</td>';
  }

  // ------------------------------------------------------------------ validation
  function today() { return EN.todaySerial(); }
  function todayIso() { return EN.isoFromSerial(EN.todaySerial()); }
  function validate(def, col, row) {
    if (!def || !col || col.auto || col.fixed) return '';
    var sheet = def.sheet, addr = col.c + row;
    var v = book.get(sheet, addr);
    var hasId = def.cols && def.cols[0] && def.cols[0].auto ? val(sheet, def.cols[0].c + row) !== '' : false;
    if (col.required && hasId && String(v).trim() === '') return 'Required once the row has an ID';
    if (col.requiredIf && book.get(sheet, col.requiredIf + row) !== '' && String(v).trim() === '') return 'Required';
    if (col.positiveIf && book.get(sheet, col.positiveIf + row) !== '' && !(typeof v === 'number' && v > 0)) return 'Enter a number greater than 0';
    if (col.duration && v !== '' && typeof val(sheet, 'M' + row) !== 'number') return 'Use "<number> <unit>", e.g. 6 hours, 3 days, 2 weeks, 1 month';
    var s = EN.serial(v);
    if (col.notFuture && s !== null && s > today()) return 'Date cannot be in the future';
    if (col.notBefore && s !== null) {
      var o = EN.serial(book.get(sheet, col.notBefore + row));
      if (o !== null && s < o) return 'Cannot be before ' + col.notBefore + row;
    }
    return '';
  }

  // ------------------------------------------------------------------ navigation
  var current = 'dash';
  function renderNav() {
    var groups = [], seen = {};
    SC.sheets.forEach(function (d) { if (!seen[d.group]) { seen[d.group] = 1; groups.push(d.group); } });
    // Sheets hidden in the workbook are listed last.
    groups.sort(function (x, y) { return (/^Hidden/.test(x) ? 1 : 0) - (/^Hidden/.test(y) ? 1 : 0); });
    var html = '';
    groups.forEach(function (g) {
      html += '<div class="nav-group">' + esc(g) + '</div>';
      SC.sheets.filter(function (d) { return d.group === g; }).forEach(function (d) {
        html += '<a href="#' + d.id + '" class="nav-item' + (d.id === current ? ' active' : '') + '" data-id="' + d.id + '">' +
          '<span class="nav-num">' + d.nav + '</span><span class="nav-title">' + esc(d.title) + '</span>' +
          '<span class="nav-count" data-count="' + d.id + '"></span></a>';
      });
    });
    document.getElementById('nav').innerHTML = html;
    updateCounts();
  }
  function usedRows(def) {
    var out = [];
    if (!def.usedBy) return out;
    for (var r = def.first; r <= def.last; r++) {
      if (def.usedBy.some(function (c) { return book.get(def.sheet, c + r) !== ''; })) out.push(r);
    }
    return out;
  }
  function updateCounts() {
    SC.sheets.forEach(function (d) {
      var el = document.querySelector('[data-count="' + d.id + '"]');
      if (!el) return;
      var n = '';
      if (d.kind === 'table' && !d.fixedRows) n = usedRows(d).length || '';
      else if (d.id === 'bia') n = calc.lists.biaIds.length || '';
      else if (d.id === 'reg') n = calc.register.length || '';
      else if (d.id === 'checks') n = issues.filter(function (x) { return x.level === 'red'; }).length || '';
      if (d.id === 'capa') { var sug = CK.newFindings(val, todayIso()).length; if (sug) n = (n ? n + ' \u00b7 ' : '') + sug + ' new'; }
      el.textContent = n;
    });
  }
  function go(id) {
    if (!SC.byId(id)) id = 'dash';
    current = id;
    if (location.hash.slice(1) !== id && history.replaceState) history.replaceState(null, '', '#' + id);
    document.querySelectorAll('.nav-item').forEach(function (a) { a.classList.toggle('active', a.getAttribute('data-id') === id); });
    document.body.classList.remove('nav-open');
    renderPage();
    document.getElementById('main').scrollTop = 0;
    window.scrollTo(0, 0);
  }

  // ------------------------------------------------------------------ pages
  function pageHeader(def, extra) {
    var inBook = T.sheets.indexOf(def.sheet) >= 0;
    var eyebrow = def.kind === 'lists' ? 'Workbook support lists \u00b7 ' + esc(def.sheet) + ' rows 137+' : (inBook ? 'Sheet ' + esc(def.nav) + ' \u00b7 ' + esc(def.sheet) : 'Website page - not exported');
    return '<header class="page-head"><div class="eyebrow">' + eyebrow + '</div>' +
      '<h1>' + esc(def.title) + '</h1>' + (def.iso ? '<p class="iso">' + esc(def.iso) + '</p>' : '') +
      (def.note ? '<p class="note">' + esc(def.note) + '</p>' : '') + (extra || '') + '</header>';
  }
  function renderPage() {
    var def = SC.byId(current);
    var html;
    switch (def.kind) {
      case 'form': html = renderForm(def); break;
      case 'table': html = renderTable(def); break;
      case 'grid': html = renderGrid(def); break;
      case 'bia': html = renderBia(def); break;
      case 'lists': html = renderLists(def); break;
      case 'help': html = renderHelp(def); break;
      case 'checks': html = renderChecks(def); break;
      case 'view': html = def.id === 'dash' ? renderDashboard(def) : (def.id === 'reg' ? renderRegister(def) : renderTargets(def)); break;
    }
    var main = document.getElementById('page');
    main.innerHTML = html;
    autosizeAll(main);
    refreshValidation();
  }

  // ---- forms (00, 03)
  function renderForm(def) {
    var html = pageHeader(def);
    def.sections.forEach(function (sec) {
      html += '<section class="card"><h2>' + esc(sec.title) + '</h2>';
      if (sec.rowsFields) {
        html += '<div class="form-grid">';
        sec.rowsFields.rows.forEach(function (r) {
          var label = book.get(def.sheet, sec.rowsFields.labelCol + r);
          html += '<label class="field"><span>' + esc(label) + '</span>' + inputHtml(def.sheet, sec.rowsFields.col + r, { type: sec.rowsFields.type }, r) + '</label>';
        });
        html += '</div>';
      }
      var grids = sec.grids || (sec.grid ? [sec.grid] : []);
      grids.forEach(function (g) {
        if (g.caption) html += '<h3>' + esc(g.caption) + '</h3>';
        html += '<div class="table-wrap"><table class="sheet-table"><thead><tr>' + (g.labelCol ? '<th>' + esc(g.labelTitle || '') + '</th>' : '');
        g.cols.forEach(function (c) { html += '<th>' + esc(c.label) + '</th>'; });
        html += '</tr></thead><tbody>';
        g.rows.forEach(function (r) {
          html += '<tr>' + (g.labelCol ? '<td class="label-cell">' + esc(book.get(def.sheet, g.labelCol + r)) + '</td>' : '');
          g.cols.forEach(function (c) {
            if (c.autoRows && c.autoRows.indexOf(r) >= 0) html += calcCell(def.sheet, c.c + r);
            else html += '<td>' + inputHtml(def.sheet, c.c + r, c, r) + '</td>';
          });
          html += '</tr>';
        });
        html += '</tbody></table></div>';
      });
      if (sec.fields) {
        html += '<div class="form-grid">';
        sec.fields.forEach(function (f) {
          html += '<label class="field"><span>' + esc(f.label) + '</span>' +
            (f.auto ? '<div class="auto-box" data-calc="' + key(def.sheet, f.addr) + '">' + esc(fmt(val(def.sheet, f.addr))) + '</div>' : inputHtml(def.sheet, f.addr, f, 0)) +
            (f.hint ? '<small>' + esc(f.hint) + '</small>' : '') + '</label>';
        });
        html += '</div>';
      }
      html += '</section>';
    });
    return html;
  }

  // ---- record tables
  function visibleRows(def) {
    if (def.fixedRows) { var all = []; for (var r = def.first; r <= def.last; r++) all.push(r); return all; }
    var used = usedRows(def);
    var lastUsed = used.length ? used[used.length - 1] : def.first - 1;
    var want = Math.max(lastUsed + 1, def.first + (state.ui.rows[def.id] || 3) - 1);
    want = Math.min(want, def.last);
    var rows = [];
    for (var q = def.first; q <= want; q++) rows.push(q);
    return rows;
  }
  // The workbook's cell notes on a column header (guidance for that column).
  function noteFor(sheet, col, beforeRow) {
    var notes = (T.comments || {})[sheet] || {}, best = '';
    for (var r = 1; r < beforeRow; r++) if (notes[col + r]) best = notes[col + r];
    return best;
  }
  function headerRows(cols, extraFirst, sheet, firstRow) {
    var hasGroups = cols.some(function (c) { return c.group; });
    var top = '', bottom = '';
    if (hasGroups) {
      top = '<tr class="group-row">' + (extraFirst ? '<th></th>' : '');
      var i = 0;
      while (i < cols.length) {
        var g = cols[i].group || '', span = 1;
        while (i + span < cols.length && (cols[i + span].group || '') === g) span++;
        top += '<th colspan="' + span + '"' + (g ? ' class="grp"' : '') + '>' + esc(g) + '</th>';
        i += span;
      }
      top += '<th></th></tr>';
    }
    bottom = '<tr>' + (extraFirst ? '<th class="num">#</th>' : '');
    cols.forEach(function (c) {
      var note = sheet ? noteFor(sheet, c.c, firstRow) : '';
      var tip = (c.auto ? 'Automatic (formula)' : '') + (note ? (c.auto ? ' - ' : '') + note : '');
      bottom += '<th style="min-width:' + (c.w || 120) + 'px"' + (c.auto ? ' class="auto-h"' : '') + (tip ? ' title="' + esc(tip) + '"' : '') + '>' +
        esc(c.label) + (note ? ' <span class="info">i</span>' : '') + '</th>';
    });
    bottom += '<th></th></tr>';
    return top + bottom;
  }
  function renderTable(def) {
    var rows = visibleRows(def);
    var html = pageHeader(def, legend());
    if (def.id === 'capa') html += capaPanel();
    html += '<div class="card flush"><div class="table-wrap tall"><table class="sheet-table records" data-sheet="' + esc(def.id) + '"><thead>' +
      headerRows(def.cols, !!def.numCol, def.sheet, def.first) + '</thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr data-row="' + r + '">' + (def.numCol ? '<td class="num">' + esc(book.get(def.sheet, def.numCol + r)) + '</td>' : '');
      def.cols.forEach(function (c) {
        if (c.auto) html += calcCell(def.sheet, c.c + r, c.badge);
        else if (c.fixed) html += '<td class="fixed">' + esc(fmt(book.get(def.sheet, c.c + r))) + '</td>';
        else html += '<td>' + inputHtml(def.sheet, c.c + r, c, r) + '</td>';
      });
      html += '<td class="row-tools">' + (def.fixedRows ? '' : '<button class="icon-btn" data-clear-row="' + r + '" title="Clear this row">\u00d7</button>') + '</td></tr>';
    });
    html += '</tbody></table></div>';
    if (!def.fixedRows) {
      var cap = def.last - def.first + 1;
      html += '<div class="table-foot"><button class="btn ghost" data-add-row="' + def.id + '"' + (rows.length >= cap ? ' disabled' : '') + '>+ Add row</button>' +
        '<span class="muted">' + usedRows(def).length + ' of ' + cap + ' rows used (workbook capacity)</span></div>';
    }
    html += '</div>';
    return html;
  }
  // ---- CAPA suggestions (pre-population from toolkit findings)
  function capaPanel() {
    var list = CK.newFindings(val, todayIso());
    if (!list.length) return '<div class="card ok-note">No new findings to raise. Suggested CAPAs appear here when the registers show gaps ' +
      '(requirement flags not captured, overdue strategies, unmitigated critical risks, failed or overdue exercises, DR gaps, overdue BIA reviews, single points of failure).</div>';
    var html = '<section class="card suggest"><div class="suggest-head"><div><h2>Suggested CAPAs from toolkit findings (' + list.length + ')</h2>' +
      '<p class="muted">Pre-filled with source, finding, severity, owner where known, target date and status Open. Add them, then complete root cause and corrective action.</p></div>' +
      '<div class="suggest-actions"><button class="btn" data-capa-add="selected">Add selected</button><button class="btn primary" data-capa-add="all">Add all ' + list.length + '</button></div></div>' +
      '<div class="table-wrap" style="max-height:300px"><table class="sheet-table"><thead><tr><th style="width:34px"><input type="checkbox" data-capa-all checked></th><th>Source Type</th><th>Source Ref</th><th>Finding</th><th>Severity</th><th>Owner</th><th>Target Date</th></tr></thead><tbody>';
    list.forEach(function (f, i) {
      html += '<tr><td><input type="checkbox" data-capa-pick="' + i + '" checked></td><td>' + esc(f.type) + '</td><td>' + esc(f.ref) + '</td><td>' + esc(f.finding) +
        '</td><td>' + badge(f.severity === 'Critical' || f.severity === 'High' ? f.severity : f.severity) + '</td><td>' + esc(f.owner || '') + '</td><td>' + esc(f.target) + '</td></tr>';
    });
    return html + '</tbody></table></div></section>';
  }
  function addCapas(list) {
    var C = S.CAPA, added = 0, full = 0;
    var inputCols = SC.byId('capa').cols.filter(function (c) { return !c.auto; }).map(function (c) { return c.c; });
    list.forEach(function (f) {
      var row = null;
      for (var r = 6; r <= 105 && row === null; r++) {
        if (inputCols.every(function (c) { return book.get(C, c + r) === ''; })) row = r;
      }
      if (row === null) { full++; return; }
      var put = { C: f.type, D: f.ref, E: todayIso(), F: f.finding, G: f.severity, K: f.owner || '', L: f.target, M: 'Open',
        U: f.riskId || '', V: f.ssId || '', W: f.rrId || '', X: 'Suggested by toolkit check' };
      Object.keys(put).forEach(function (c) { if (put[c] !== '') setValue(C, c + row, put[c]); });
      added++;
    });
    recompute();
    var y = window.scrollY;
    renderPage();
    window.scrollTo(0, y);
    updateCounts();
    toast('Added ' + added + ' CAPA' + (added === 1 ? '' : 's') + (full ? ' - ' + full + ' not added: CAPA register is full (100 rows)' : ''), full ? 'err' : 'ok');
  }

  // ---- Checks & gaps page
  function renderChecks(def) {
    var red = issues.filter(function (x) { return x.level === 'red'; }), amber = issues.filter(function (x) { return x.level === 'amber'; });
    var sug = CK.newFindings(val, todayIso());
    var html = pageHeader(def, '<p class="note">Red: a critical entry is missing in a row that is in use, or a requirement flag is Yes without a matching row in 12. Requirements Register. ' +
      'Amber: requirements captured but not all Ready. The same rules highlight the cells in red / amber on each page and in the exported Excel file.</p>');
    html += '<div class="tiles wide">' + tile('Missing critical entries', red.length, '', red.length > 0) + tile('Needs attention', amber.length) +
      tile('Suggested CAPAs not yet raised', sug.length, '', sug.length > 0) + '</div>';
    if (sug.length) html += '<p><a class="btn primary" href="#capa">Review and add suggested CAPAs \u2192</a></p>';
    var bySheet = {};
    issues.forEach(function (x) { (bySheet[x.sheet] = bySheet[x.sheet] || []).push(x); });
    if (!issues.length) return html + '<div class="card ok-note">No gaps found. Every row in use has its critical entries.</div>';
    Object.keys(bySheet).sort().forEach(function (sheet) {
      var l = bySheet[sheet];
      html += '<section class="card flush"><h2 class="card-pad">' + esc(sheet) + ' <span class="muted">(' + l.length + ')</span></h2><div class="table-wrap" style="max-height:360px"><table class="sheet-table"><thead><tr>' +
        '<th style="width:90px">Level</th><th style="width:110px">Record</th><th style="width:70px">Cell</th><th>What is missing</th><th style="width:80px"></th></tr></thead><tbody>';
      l.forEach(function (x) {
        html += '<tr><td>' + (x.level === 'red' ? '<span class="badge red">Missing</span>' : '<span class="badge amber">Attention</span>') + '</td><td>' + esc(x.ref) + '</td><td>' + esc(x.addr) +
          '</td><td>' + esc(x.what) + '</td><td><a href="#" data-goto="' + esc(x.id + '|' + x.sheet + '|' + x.addr) + '">Go to</a></td></tr>';
      });
      html += '</tbody></table></div></section>';
    });
    return html;
  }
  function gotoCell(spec) {
    var p = spec.split('|');
    go(p[0]);
    var el = document.querySelector('[data-k="' + cssEsc(p[1] + '|' + p[2]) + '"]');
    if (!el) return;
    el.scrollIntoView({ block: 'center', inline: 'center' });
    el.focus({ preventScroll: true });
    el.classList.add('flash');
    setTimeout(function () { el.classList.remove('flash'); }, 1600);
  }

  function legend() {
    return '<div class="legend"><span><i class="sw white"></i>You fill these</span><span><i class="sw yellow"></i>Dropdown</span>' +
      '<span><i class="sw blue"></i>Automatic - calculated like the Excel formulas</span><span><i class="sw red"></i>Critical entry missing - see Checks &amp; Gaps</span>' +
      '<span>Tip: copy a block of cells in Excel and paste it here with Ctrl+V - it fills across columns and rows</span></div>';
  }

  // ---- simple grids (06, 09)
  function renderGrid(def) {
    var html = pageHeader(def, legend()) + '<div class="card flush"><div class="table-wrap tall"><table class="sheet-table"><thead><tr><th class="num">#</th>';
    def.cols.forEach(function (c) { html += '<th style="min-width:180px">' + esc(c.label) + '</th>'; });
    html += '</tr></thead><tbody>';
    for (var r = def.rows.first; r <= def.rows.last; r++) {
      html += '<tr><td class="num">' + (r - def.rows.first + 1) + '</td>';
      def.cols.forEach(function (c) { html += '<td>' + inputHtml(def.sheet, c.c + r, {}, r) + '</td>'; });
      html += '</tr>';
    }
    return html + '</tbody></table></div></div>';
  }

  // ---- BIA
  function scoreTone(v) { return typeof v === 'number' ? ' s' + v : ''; }
  function renderBia(def) {
    var html = pageHeader(def, legend() +
      '<p class="note">One block per in-scope product / service, in the order of 04. Products &amp; Services (BIA Order). Score each impact type 1-5 across the nine horizons; ' +
      'MTPD is the first horizon where any impact reaches 4. Keep RTO below MTPD. Scoring anchors are on the <a href="#help">How to use</a> page.</p>');
    var blocks = [];
    for (var k = 1; k <= SC.BIA.blocks; k++) if (val(S.BIA, 'E' + SC.BIA.top(k)) !== '') blocks.push(k);
    if (!blocks.length) {
      return html + '<div class="card empty">No BIA blocks yet. Mark products / services as <b>In Scope = Yes</b> in <a href="#ps">04. Products &amp; Services</a> - each one becomes a BIA block here.</div>';
    }
    html += '<div class="chips">';
    blocks.forEach(function (k) { html += '<a class="chip" href="#bia-' + k + '" data-jump="bia-' + k + '">' + esc(val(S.BIA, 'B' + SC.BIA.top(k))) + ' \u00b7 ' + esc(val(S.BIA, 'E' + SC.BIA.top(k))) + '</a>'; });
    html += '</div>';
    var scoreCol = { type: 'select', options: 'lstScore', numeric: true };
    var yn = { type: 'select', options: 'lstYN' };
    blocks.forEach(function (k) {
      var t = SC.BIA.top(k);
      html += '<section class="card bia-block" id="bia-' + k + '"><div class="bia-head"><div><div class="eyebrow">' + esc(val(S.BIA, 'B' + t)) + ' \u00b7 ' + esc(val(S.BIA, 'F' + t)) + '</div>' +
        '<h2>' + esc(val(S.BIA, 'E' + t)) + '</h2><div class="muted">' + esc(val(S.BIA, 'C' + t)) + ' \u00b7 ' + esc(val(S.BIA, 'D' + t)) + '</div></div>' +
        '<div class="kpis"><div><span>Max impact</span><b data-calc="' + key(S.BIA, 'S' + t) + '">' + esc(fmt(val(S.BIA, 'S' + t))) + '</b></div>' +
        '<div><span>MTPD (hrs)</span><b data-calc="' + key(S.BIA, 'AA' + t) + '">' + esc(fmt(val(S.BIA, 'AA' + t))) + '</b></div>' +
        '<div><span>Critical - final</span><b data-calc="' + key(S.BIA, 'W' + t) + '" data-badge="1">' + badge(val(S.BIA, 'W' + t)) + '</b></div></div></div>';
      html += '<div class="table-wrap"><table class="sheet-table bia-table"><thead><tr><th>Dependency</th><th>Will it interrupt the operation?</th><th>Impact Type</th>';
      SC.BIA.horizons.forEach(function (h) { html += '<th class="hz">' + esc(h) + '</th>'; });
      html += '<th class="auto-h" style="min-width:130px">Max Impact Type (auto)</th></tr></thead><tbody>';
      for (var i = 0; i < 7; i++) {
        var r = t + i;
        html += '<tr><td class="label-cell">' + esc(book.get(S.BIA, 'G' + r)) + '</td><td>' + inputHtml(S.BIA, 'H' + r, yn, r) + '</td>' +
          '<td class="label-cell">' + esc(book.get(S.BIA, 'I' + r)) + '</td>';
        SC.BIA.scoreCols.forEach(function (c) { html += '<td class="score">' + inputHtml(S.BIA, c + r, scoreCol, r, 'score-in' + scoreTone(book.get(S.BIA, c + r))) + '</td>'; });
        html += calcCell(S.BIA, 'T' + r) + '</tr>';
      }
      html += '</tbody></table></div>';
      function af(label, inner) { return '<label class="field"><span>' + label + '</span>' + inner + '</label>'; }
      function ab(addr, isBadge) { var v = val(S.BIA, addr); return '<div class="auto-box" data-calc="' + key(S.BIA, addr) + '"' + (isBadge ? ' data-badge="1"' : '') + '>' + (isBadge ? badge(v) : esc(fmt(v))) + '</div>'; }
      html += '<div class="form-grid cols-4">' +
        af('Max Impact Type (auto)', '<div class="auto-box">' + esc(maxTypes(t)) + '</div>') +
        af('Corporate Critical (auto)', ab('U' + t, true)) +
        af('BU Critical (Y/N)', inputHtml(S.BIA, 'V' + t, yn, t)) +
        af('Critical - Final (auto)', ab('W' + t, true)) +
        af('MBCO - Corporate (auto)', ab('X' + t)) +
        af('MBCO - BU', inputHtml(S.BIA, 'Y' + t, { type: 'textarea' }, t)) +
        af('Rationale / Supporting Reasons', inputHtml(S.BIA, 'Z' + t, { type: 'textarea' }, t)) +
        af('CEO MTPD (hrs, auto)', ab('AB' + t)) +
        af('RTO', inputHtml(S.BIA, 'AC' + t, { placeholder: 'e.g. 2 days' }, t) + '<small data-rto-warn="' + t + '"></small>') +
        af('RTO in hours (auto)', ab('AD' + t)) +
        af('RPO in hours (fill if Data/ICT = Yes)', inputHtml(S.BIA, 'AE' + t, { type: 'number' }, t)) +
        af('Assessment Date', inputHtml(S.BIA, 'AF' + t, { type: 'date' }, t)) +
        af('Valid Until (auto +12 mo)', ab('AG' + t)) +
        af('Risk IDs (auto)', ab('AH' + t)) +
        af('S&amp;S # (auto)', ab('AI' + t)) +
        af('R&amp;R IDs (auto)', ab('AJ' + t)) + '</div></section>';
    });
    return html;
  }
  function maxTypes(t) {
    var out = [];
    for (var i = 0; i < 7; i++) { var v = val(S.BIA, 'T' + (t + i)); if (v !== '') out.push(v); }
    return out.join(', ');
  }

  // ---- Dashboard
  function tile(label, value, sub, warn) {
    return '<div class="tile' + (warn ? ' warn' : '') + '"><div class="tile-label">' + esc(label) + '</div><div class="tile-value" data-dash>' + esc(value) + '</div>' + (sub ? '<div class="tile-sub">' + esc(sub) + '</div>' : '') + '</div>';
  }
  function renderDashboard(def) {
    var D = calc.dashboard;
    var bu = book.get(S.BUP, 'B6') || book.get(S.BUP, 'B5');
    var raIncomplete = 0;
    for (var ri = 5; ri <= 34; ri++) if (val(S.RA, 'Z' + ri) === 'Incomplete') raIncomplete++;
    var html = '<header class="page-head"><div class="eyebrow">Sheet 02 \u00b7 02. Dashboard</div><h1>BCM Programme Dashboard' + (bu ? ' - ' + esc(bu) : '') + '</h1>' +
      '<p class="iso">Live counters - review weekly - red items need action, not explanation | ISO 22301: 9.1</p></header>';
    html += '<section class="card intro"><div><h2>Fill the toolkit online, export the workbook</h2><p>Work through the sheets in the left menu (same order as the Excel toolkit). ' +
      'Everything you type is kept in this browser and recalculated live. <b>Export to Excel</b> writes your data into the original <i>' + esc(T.source) +
      '</i> so every formula, dropdown and format stays intact. <b>Import Excel</b> loads an existing filled toolkit.</p></div>' +
      '<div class="intro-actions"><button class="btn primary" data-action="export">Export to Excel</button><button class="btn" data-action="import">Import Excel</button></div></section>';
    html += '<div class="dash-grid">';
    html += '<section class="card"><h2>Coverage</h2><div class="tiles">' +
      tile('P&S in scope', D.psInScope) + tile('BIAs completed', D.biaCompleted) + tile('Critical P&S (final)', D.criticalPs) +
      tile('P&S-risk associations', D.psRiskAssoc) + tile('Dependencies registered', D.dependencies) +
      tile('BIA reviews overdue', D.biaOverdue, '', D.biaOverdue > 0) + '</div></section>';
    html += '<section class="card"><h2>Risk &amp; Readiness</h2><div class="tiles">' +
      tile('Risks assessed', D.risks) + tile('Critical / Significant risks', D.critSig, '', D.critSig > 0) + tile('Priority 1-2 risks (BCP)', D.prio12) +
      tile('Solutions defined / overdue', D.ssDefined + ' / ' + D.ssOverdue, '', D.ssOverdue > 0) + tile('Requirement rows (5b)', D.reqRows) + '</div></section>';
    html += '<section class="card"><h2>Governance &amp; Compliance</h2><div class="tiles">' +
      tile('BCP rows (RR plan)', D.bcpRows) + tile('Exercises done / overdue', D.exercisesDone + ' / ' + D.exercisesOverdue, '', D.exercisesOverdue > 0) +
      tile('DR systems / tested-current', D.drSystems + ' / ' + D.drCurrent) + tile('Duplicate S&S mappings', D.ssDuplicates, '', D.ssDuplicates > 0) +
      tile('Risk assessments incomplete', raIncomplete, '', raIncomplete > 0) + '</div></section>';
    html += '<section class="card"><h2>Corrective Action &amp; Improvement</h2><div class="tiles">' +
      tile('Open CAPAs', D.capaOpen) + tile('Overdue CAPAs', D.capaOverdue, '', D.capaOverdue > 0) +
      tile('Effectiveness pending', D.capaPending) + tile('Critical / High open', D.capaCritHigh, '', D.capaCritHigh > 0) + '</div></section>';
    var redN = issues.filter(function (x) { return x.level === 'red'; }).length, amberN = issues.length - redN, sugN = CK.newFindings(val, todayIso()).length;
    html += '<section class="card"><h2>Data quality <a class="muted small-link" href="#checks">open Checks &amp; Gaps \u2192</a></h2><div class="tiles">' +
      tile('Missing critical entries', redN, '', redN > 0) + tile('Needs attention', amberN) + tile('Suggested CAPAs not yet raised', sugN, '', sugN > 0) + '</div></section>';
    html += '</div><div class="dash-grid two">';
    html += '<section class="card"><h2>Status snapshot</h2>' + statusSnapshot() + '</section>';
    html += '<section class="card"><h2>Risk heat map <span class="muted">(residual, from 10. Risk Assessment)</span></h2>' + heatMap() + '</section></div>';
    return html;
  }
  function heatMap() {
    var counts = {};
    for (var r = 5; r <= 34; r++) {
      var k = val(S.RA, 'K' + r), m = val(S.RA, 'M' + r);
      if (val(S.RA, 'B' + r) !== '' && k !== '' && m !== '') counts[k + '-' + m] = (counts[k + '-' + m] || 0) + 1;
    }
    var lik = ['Almost Certain', 'Frequent', 'Occasional', 'Infrequent', 'Rare'];
    var imp = ['Negligible', 'Minor', 'Medium', 'Critical', 'Catastrophic'];
    var html = '<table class="heat"><tbody>';
    // ratings: >= 20 Critical, >= 12 Significant, >= 7 Moderate, else Low
    for (var l = 5; l >= 1; l--) {
      html += '<tr><th>' + lik[5 - l] + '</th>';
      for (var i = 1; i <= 5; i++) {
        // Risk score = likelihood x impact, rated as in 10. Risk Assessment (P).
        var ref = l * i;
        var t = ref >= 20 ? 'red' : (ref >= 12 ? 'orange' : (ref >= 7 ? 'yellow' : 'green'));
        var n = counts[l + '-' + i] || 0;
        html += '<td class="h-' + t + '" title="Score ' + ref + ' (' + lik[5 - l] + ' x ' + imp[i - 1] + ')">' + (n ? '<b>' + n + '</b>' : '') + '<span>' + ref + '</span></td>';
      }
      html += '</tr>';
    }
    html += '<tr><th></th>' + imp.map(function (x) { return '<th class="imp">' + x + '</th>'; }).join('') + '</tr></tbody></table>';
    return html;
  }

  // ---- Register (15) and Analysis (21)
  function renderRegister(def) {
    var cols = SC.colRange('A', 'X');
    var badgeCols = { G: 1, M: 1, O: 1, P: 1, Q: 1, X: 1 };
    var html = pageHeader(def, '<p class="note">Read-only - built automatically from 10. Risk Assessment, 11. Strategies &amp; Solutions and 14. Recovery &amp; Response.</p>');
    if (!calc.register.length) return html + '<div class="card empty">No risks yet. Risks appear here once a Risk Assumption is selected in <a href="#ra">10. Risk Assessment</a>.</div>';
    html += '<div class="card flush"><div class="table-wrap tall"><table class="sheet-table"><thead><tr>';
    cols.forEach(function (c) { html += '<th style="min-width:' + (c === 'A' ? 80 : 130) + 'px">' + esc(book.get(S.REG, c + '5')) + '</th>'; });
    html += '</tr></thead><tbody>';
    calc.register.forEach(function (row) {
      html += '<tr>';
      cols.forEach(function (c) { html += '<td class="auto">' + (badgeCols[c] ? badge(row[c]) : esc(fmt(row[c]))) + '</td>'; });
      html += '</tr>';
    });
    return html + '</tbody></table></div></div>';
  }
  function renderTargets(def) {
    var cols = [['B', 'P&S ID'], ['C', 'Business Unit'], ['D', 'Division / Site'], ['E', 'Product / Service'], ['F', 'Corporate MBCO'], ['G', 'Corporate MTPD'],
      ['I', 'BIA MTPD (h)'], ['J', 'RTO (h)'], ['K', 'RPO (h)']];
    var html = pageHeader(def, '<p class="note">Read-only summary of the recovery targets of every in-scope product / service: corporate targets from 04. Products &amp; Services ' +
      'and the MTPD, RTO and RPO of its BIA block in 07. BIA. RTO must be shorter than MTPD.</p>');
    html += '<div class="card flush"><div class="table-wrap tall"><table class="sheet-table"><thead><tr><th class="num">#</th>';
    cols.forEach(function (c) { html += '<th style="min-width:110px">' + esc(c[1]) + '</th>'; });
    html += '<th style="min-width:130px">RTO within MTPD</th></tr></thead><tbody>';
    if (!calc.analysis.length) html += '<tr><td colspan="' + (cols.length + 2) + '" class="muted">No in-scope products / services yet.</td></tr>';
    calc.analysis.forEach(function (row, i) {
      html += '<tr><td class="num">' + (i + 1) + '</td>';
      cols.forEach(function (c) { html += '<td class="auto">' + esc(fmt(row[c[0]])) + '</td>'; });
      var ok = typeof row.I === 'number' && typeof row.J === 'number' ? (row.J < row.I ? 'Yes' : 'No - review') : (row.J === '' ? 'RTO not set' : '');
      html += '<td class="auto">' + badge(ok) + '</td></tr>';
    });
    return html + '</tbody></table></div></div>';
  }
  function statusSnapshot() {
    function counts(sheet, col, a, b) {
      var o = {};
      for (var r = a; r <= b; r++) { var v = val(sheet, col + r); if (v !== '') o[v] = (o[v] || 0) + 1; }
      return o;
    }
    var groups = [['Strategies (11)', counts(S.SS, 'T', 5, 64)], ['Exercises (14)', counts(S.RR, 'P', 5, 104)], ['IT DR (13)', counts(S.DR, 'Q', 5, 204)],
      ['Risk assessment source (10)', counts(S.RA, 'Z', 5, 34)], ['CAPA (16)', counts(S.CAPA, 'M', 6, 105)]];
    var html = '<table class="mini"><tbody>';
    groups.forEach(function (gp) {
      var keys = Object.keys(gp[1]);
      html += '<tr><td>' + esc(gp[0]) + '</td><td>' + (keys.length ? keys.map(function (k) { return badge(k) + ' <b>' + gp[1][k] + '</b>'; }).join(' ') : '<span class="muted">-</span>') + '</td></tr>';
    });
    return html + '</tbody></table>';
  }

  // ---- Support lists (03. Business Unit Profile rows 138+)
  function renderLists(def) {
    var html = pageHeader(def, legend());
    html += '<div class="lists-grid">';
    SC.LISTS.forEach(function (l) {
      html += '<section class="card"><h2>' + esc(l.label) + ' <span class="muted">' + esc(l.name) + '</span>' + (l.core ? ' <span class="badge amber">used in formulas</span>' : '') + '</h2>';
      for (var r = l.first; r <= l.last; r++) html += '<div class="list-row">' + inputHtml(def.sheet, l.col + r, { type: l.numeric ? 'number' : 'text' }, r) + '</div>';
      html += '</section>';
    });
    var U = SC.DURATION_UNITS;
    html += '<section class="card"><h2>Duration units <span class="muted">DurationUnits</span> <span class="badge amber">used in formulas</span></h2>' +
      '<p class="muted">Unit word and hours per unit, used to convert entries such as "3 days" into hours (MBCO / MTPD / RTO).</p><table class="mini"><tbody>';
    for (var d = U.first; d <= U.last; d++) {
      html += '<tr><td>' + inputHtml(def.sheet, U.unitCol + d, {}, d) + '</td><td>' + inputHtml(def.sheet, U.hoursCol + d, { type: 'number' }, d) + '</td></tr>';
    }
    return html + '</tbody></table></section></div>';
  }

  // ---- Help
  function renderHelp(def) {
    var h = pageHeader(def);
    h += '<section class="card help"><h2>Workflow</h2><ol>' +
      '<li><b>Plan - context &amp; scope:</b> complete <a href="#bup">03. Business Unit Profile</a>, then list products and services in <a href="#ps">04</a> and mark those in scope. Each in-scope item gets a BIA block in <a href="#bia">07. BIA</a>.</li>' +
      '<li><b>Analyse:</b> list activities in <a href="#proc">05</a> and applications / equipment in <a href="#app">06</a>; score impacts in <a href="#bia">07</a>; record dependencies in <a href="#dep">08</a>.</li>' +
      '<li><b>Assess risk:</b> pick assumptions from <a href="#map">09. Risk Assumption Map</a> in <a href="#ra">10. Risk Assessment</a>, rate residual likelihood and impact, and complete owner, dates, approver and evidence so the source is <i>Complete</i>.</li>' +
      '<li><b>Treat:</b> define strategies in <a href="#ss">11</a>; every requirement flag set to <i>Yes</i> needs rows in <a href="#req">12. Requirements Register</a>. Record IT DR in <a href="#dr">13</a> and response &amp; recovery plans and exercises in <a href="#rr">14</a>.</li>' +
      '<li><b>Check &amp; improve:</b> review <a href="#reg">15. BCM Risk Register</a>, <a href="#checks">Checks &amp; Gaps</a> and raise corrective actions in <a href="#capa">16. CAPA Register</a> (suggested CAPAs can be added with one click).</li></ol></section>';
    h += '<section class="card help"><h2>Colours</h2>' + legend() + '<p>Amber cells need attention (for example requirements captured but not all Ready, or a risk assessment that is still Incomplete). The exported workbook shows the same red / amber highlights.</p></section>';
    h += '<section class="card help"><h2>Risk scoring</h2><p>Risk score = residual likelihood (1-5) &times; residual impact (1-5). Rating: <span class="badge red">Critical</span> 20-25 (BCP priority 1), ' +
      '<span class="badge orange">Significant</span> 12-19 (priority 2), <span class="badge amber">Moderate</span> 7-11 (priority 3), <span class="badge green">Low</span> 1-6 (priority 4).</p>' +
      '<table class="mini"><tbody><tr><th>Score</th><th>Likelihood</th><th>Impact</th></tr>' +
      [['1', 'Rare', 'Negligible'], ['2', 'Infrequent', 'Minor'], ['3', 'Occasional', 'Medium'], ['4', 'Frequent', 'Critical'], ['5', 'Almost Certain', 'Catastrophic']]
        .map(function (x) { return '<tr><td class="num">' + x[0] + '</td><td>' + x[1] + '</td><td>' + x[2] + '</td></tr>'; }).join('') + '</tbody></table>' +
      '<p><b>BIA impact scores</b> (07. BIA) use the same 1-5 scale per impact type and horizon; MTPD is the first horizon where any impact reaches 4.</p></section>';
    h += '<section class="card help"><h2>Durations</h2><p>MBCO, MTPD and RTO accept a number of hours or "&lt;number&gt; &lt;unit&gt;", for example <i>6 hours</i>, <i>3 days</i>, <i>2 weeks</i>, <i>1 month</i>. ' +
      'Units are defined on the <a href="#lists">Support Lists</a> page.</p></section>';
    h += '<section class="card help"><h2>Excel import / export</h2><ul>' +
      '<li><b>Export to Excel</b> writes your entries into the embedded template <i>' + esc(T.source) + '</i>; formulas, dropdowns, formats and highlights stay intact and Excel recalculates everything on opening.</li>' +
      '<li><b>Import Excel</b> reads any filled copy of this toolkit (same sheet layout). Files protected with a sensitivity label / IRM encryption must be saved unprotected first.</li>' +
      '<li><b>Blank template</b> downloads the empty workbook.</li>' +
      '<li><b>Paste from Excel:</b> copy a block of cells, click the first target cell and press Ctrl+V. Dropdown values are matched, dates are recognised, automatic columns are skipped.</li>' +
      '<li>Data stays in this browser only. Use <b>Backup</b> to save a .json copy and <b>Restore</b> to load it on another computer.</li></ul></section>';
    return h;
  }

  // ------------------------------------------------------------------ live refresh
  function refreshCalc() {
    var def = SC.byId(current);
    if (def.kind === 'view') {
      // Views are cheap to re-render and have no inputs.
      var y = window.scrollY;
      renderPage();
      window.scrollTo(0, y);
    } else {
      document.querySelectorAll('[data-calc]').forEach(function (el) {
        var p = el.getAttribute('data-calc').split('|');
        var v = val(p[0], p[1]);
        if (el.getAttribute('data-badge')) el.innerHTML = badge(v);
        else el.textContent = fmt(v);
      });
      if (def.kind === 'bia') {
        document.querySelectorAll('.bia-block').forEach(function (sec) {
          var t = SC.BIA.top(+sec.id.slice(4));
          var box = sec.querySelector('.form-grid .auto-box');
          if (box) box.textContent = maxTypes(t);
        });
      }
      refreshValidation();
    }
    updateCounts();
  }
  function refreshValidation() {
    var def = SC.byId(current);
    if (def.kind === 'table') {
      document.querySelectorAll('.records tr[data-row]').forEach(function (tr) {
        var r = +tr.getAttribute('data-row');
        def.cols.forEach(function (c) {
          if (c.auto || c.fixed) return;
          var el = tr.querySelector('[data-k="' + cssEsc(def.sheet + '|' + c.c + r) + '"]');
          if (!el) return;
          var msg = validate(def, c, r), lvl = msg ? 'red' : '';
          var iss = issueAt[def.sheet + '|' + c.c + r];
          if (!msg && iss) { msg = iss.what; lvl = iss.level; }
          el.classList.toggle('invalid', lvl === 'red');
          el.classList.toggle('attention', lvl === 'amber');
          el.title = msg;
        });
      });
    }
    if (def.kind === 'bia') {
      document.querySelectorAll('[data-rto-warn]').forEach(function (el) {
        var t = +el.getAttribute('data-rto-warn');
        var rto = val(S.BIA, 'AD' + t), mtpd = val(S.BIA, 'AA' + t);
        var msg = '';
        if (rto === '#VALUE!') msg = 'Use "<number> <unit>", e.g. 12 hours or 2 days';
        else if (typeof rto === 'number' && typeof mtpd === 'number' && rto >= mtpd) msg = 'RTO should be shorter than MTPD (' + mtpd + ' hrs) - keep a buffer';
        el.textContent = msg;
        el.className = msg ? 'warn-text' : '';
      });
      document.querySelectorAll('.bia-block [data-k]').forEach(function (el) {
        var k = el.getAttribute('data-k'), iss = issueAt[k];
        var s = el.type === 'date' ? EN.serial(el.value) : null;
        var msg = s !== null && s > today() ? 'Assessment date cannot be in the future' : (iss ? iss.what : '');
        el.classList.toggle('invalid', !!msg);
        el.title = msg;
      });
    }
  }
  function cssEsc(s) { return window.CSS && CSS.escape ? CSS.escape(s) : s.replace(/["\\]/g, '\\$&'); }
  function autosize(el) {
    if (el.tagName !== 'TEXTAREA') return;
    el.style.height = 'auto';
    // Keep table rows compact; long text scrolls inside the box (forms allow taller boxes).
    el.style.height = Math.min(el.scrollHeight + 2, el.closest('.sheet-table') ? 110 : 240) + 'px';
  }
  function autosizeAll(root) { root.querySelectorAll('textarea').forEach(autosize); }

  // ------------------------------------------------------------------ events
  function onInput(e) {
    var el = e.target;
    if (el.hasAttribute && el.hasAttribute('data-capa-all')) {
      document.querySelectorAll('[data-capa-pick]').forEach(function (cb) { cb.checked = el.checked; });
      return;
    }
    var k = el.getAttribute && el.getAttribute('data-k');
    if (!k) return;
    var p = k.split('|');
    setValue(p[0], p[1], el.value);
    autosize(el);
    if (el.classList.contains('score-in')) {
      el.className = el.className.replace(/\bs[1-5]\b/g, '').trim();
      if (/^[1-5]$/.test(el.value)) el.classList.add('s' + el.value);
    }
  }
  function onFocus(e) {
    var el = e.target;
    var opt = el.getAttribute && el.getAttribute('data-opt');
    if (!opt) return;
    var src = JSON.parse(opt), row = +el.getAttribute('data-row');
    var sheet = el.getAttribute('data-k').split('|')[0];
    var opts = optionsFor(src, sheet, row);
    if (el.tagName === 'SELECT') {
      var html = selectOptions(opts, el.value);
      if (el._opts !== html) { el.innerHTML = html; el._opts = html; }
    } else {
      var dl = document.getElementById('active-list');
      dl.innerHTML = opts.map(function (o) { return '<option value="' + esc(o) + '">'; }).join('');
      el.setAttribute('list', 'active-list');
    }
  }
  function onClick(e) {
    var t = e.target.closest('button, a');
    if (!t) return;
    if (t.hasAttribute('data-add-row')) {
      var def = SC.byId(t.getAttribute('data-add-row'));
      var shown = visibleRows(def).length;
      state.ui.rows[def.id] = shown + 1;
      save();
      var y = window.scrollY;
      renderPage();
      window.scrollTo(0, y);
      var last = document.querySelector('.records tbody tr:last-child [data-k]');
      if (last) last.focus();
      return;
    }
    if (t.hasAttribute('data-clear-row')) {
      var d = SC.byId(current), r = +t.getAttribute('data-clear-row');
      if (!confirm('Clear all entries in this row? IDs of other rows do not change.')) return;
      d.cols.forEach(function (c) { if (!c.auto && !c.fixed) setValue(d.sheet, c.c + r, ''); });
      recompute();
      var yy = window.scrollY;
      renderPage();
      window.scrollTo(0, yy);
      return;
    }
    if (t.hasAttribute('data-goto')) { e.preventDefault(); gotoCell(t.getAttribute('data-goto')); return; }
    if (t.hasAttribute('data-capa-add')) {
      var list = CK.newFindings(val, todayIso());
      if (t.getAttribute('data-capa-add') === 'selected') {
        var picked = {};
        document.querySelectorAll('[data-capa-pick]').forEach(function (cb) { if (cb.checked) picked[cb.getAttribute('data-capa-pick')] = 1; });
        list = list.filter(function (f, i) { return picked[i]; });
      }
      if (!list.length) { toast('Nothing selected'); return; }
      addCapas(list);
      return;
    }
    if (t.hasAttribute('data-jump')) {
      e.preventDefault();
      var target = document.getElementById(t.getAttribute('data-jump'));
      if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    var action = t.getAttribute('data-action');
    if (action) {
      e.preventDefault();
      var menu = t.closest('details');
      if (menu) menu.removeAttribute('open');
      actions[action]();
    }
  }

  // ------------------------------------------------------------------ paste from Excel
  // Excel puts a block of cells on the clipboard as tab-separated text; cells holding
  // tabs, line breaks or quotes are wrapped in quotes with "" escaping.
  function parseClipboard(text) {
    var rows = [], row = [], cell = '', i = 0, q = false;
    text = text.replace(/\r\n?/g, '\n');
    while (i < text.length) {
      var ch = text[i];
      if (q) {
        if (ch === '"' && text[i + 1] === '"') { cell += '"'; i += 2; continue; }
        if (ch === '"') { q = false; i++; continue; }
        cell += ch; i++; continue;
      }
      if (ch === '"' && cell === '') { q = true; i++; continue; }
      if (ch === '\t') { row.push(cell); cell = ''; i++; continue; }
      if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; i++; continue; }
      cell += ch; i++;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }
  var MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function isoDate(y, m, d) {
    if (y < 100) y += 2000;
    var dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return y + '-' + pad2(m) + '-' + pad2(d);
  }
  // Dates as Excel copies them: 2026-01-15, 15/01/2026 (day first), 15-Jan-26, Jan 15 2026, or a serial number.
  function parseDate(s) {
    s = s.trim();
    var m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return isoDate(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/.exec(s))) return isoDate(+m[3], +m[2], +m[1]);
    if ((m = /^(\d{1,2})[\s\-\/]([A-Za-z]{3})[A-Za-z]*[\s\-\/,]+(\d{2,4})$/.exec(s)) && MONTHS[m[2].toLowerCase()]) return isoDate(+m[3], MONTHS[m[2].toLowerCase()], +m[1]);
    if ((m = /^([A-Za-z]{3})[A-Za-z]*[\s\-]+(\d{1,2}),?\s+(\d{2,4})$/.exec(s)) && MONTHS[m[1].toLowerCase()]) return isoDate(+m[3], MONTHS[m[1].toLowerCase()], +m[2]);
    if (/^\d{4,5}(\.\d+)?$/.test(s)) return IO.excelToIso(Math.floor(Number(s)));
    return null;
  }
  // Fit a pasted text to the field: match dropdown options, parse dates and numbers.
  function normalizeFor(el, raw) {
    var v = raw.trim();
    if (v === '') return { value: '' };
    if (el.type === 'date') { var d = parseDate(v); return d ? { value: d } : { error: 'date' }; }
    if (el.type === 'number') {
      var n = Number(v.replace(/,/g, ''));
      return isNaN(n) ? { error: 'number' } : { value: String(n) };
    }
    var opt = el.getAttribute('data-opt');
    if (opt) {
      var row = +el.getAttribute('data-row'), sheet = el.getAttribute('data-k').split('|')[0];
      var opts = optionsFor(JSON.parse(opt), sheet, row);
      for (var i = 0; i < opts.length; i++) if (String(opts[i]).toLowerCase() === v.toLowerCase()) return { value: String(opts[i]) };
      if (/^\d+(\.0+)?$/.test(v)) for (var j = 0; j < opts.length; j++) if (String(opts[j]) === String(Number(v))) return { value: String(opts[j]) };
      if (el.tagName === 'SELECT' && opts.length) return { value: v, offList: true };
    }
    return { value: el.tagName === 'TEXTAREA' ? raw.replace(/^\s+|\s+$/g, '') : v };
  }
  function onPaste(e) {
    var el = e.target;
    if (!el.getAttribute || !el.getAttribute('data-k')) return;
    var text = (e.clipboardData || window.clipboardData).getData('text/plain');
    if (!text) return;
    var data = parseClipboard(text);
    // A single value typed into a text box keeps the normal paste behaviour.
    if (data.length === 1 && data[0].length === 1 && el.tagName !== 'SELECT' && el.type !== 'date') return;
    var table = el.closest('table');
    e.preventDefault();
    if (!table) {
      // Form fields and support lists: fill the fields of this section in order,
      // so a column (or row) copied from Excel lands field after field.
      var box = el.closest('section') || document.getElementById('page');
      var fields = Array.prototype.slice.call(box.querySelectorAll('[data-k]'));
      var at = fields.indexOf(el), n = { done: 0, bad: 0, offList: 0, overflow: 0, skipped: 0 };
      [].concat.apply([], data).forEach(function (raw, i) {
        var f = fields[at + i];
        if (!f) { if (raw.trim() !== '') n.overflow++; return; }
        var res = normalizeFor(f, raw);
        if (res.error) { n.bad++; return; }
        if (res.offList) n.offList++;
        var q = f.getAttribute('data-k').split('|');
        setValue(q[0], q[1], res.value);
        n.done++;
      });
      return pasteDone(n, 'beyond the last field of this section');
    }
    var td = el.closest('td'), tr = el.closest('tr');
    var cIdx = Array.prototype.indexOf.call(tr.children, td);
    var def = SC.byId(current);
    var startRow = tr.getAttribute('data-row');
    // Record tables grow to fit the pasted block (within the workbook's capacity).
    if (table.classList.contains('records') && startRow && !def.fixedRows) {
      var needed = +startRow + data.length - def.first + 1;
      if (needed > visibleRows(def).length) {
        state.ui.rows[def.id] = Math.min(needed, def.last - def.first + 1);
        var focusKey = el.getAttribute('data-k');
        renderPage();
        el = document.querySelector('[data-k="' + cssEsc(focusKey) + '"]');
        table = el.closest('table'); td = el.closest('td'); tr = el.closest('tr');
      }
    }
    var bodyRows = Array.prototype.filter.call(tr.parentNode.children, function (x) { return x.tagName === 'TR' && !x.classList.contains('section-row'); });
    var rIdx = bodyRows.indexOf(tr);
    var done = 0, skipped = 0, bad = 0, offList = 0, overflow = 0;
    data.forEach(function (cells, i) {
      var row = bodyRows[rIdx + i];
      if (!row) { overflow += cells.length; return; }
      cells.forEach(function (raw, j) {
        var cell = row.children[cIdx + j];
        var input = cell && cell.querySelector('[data-k]');
        if (!input) { if (raw.trim() !== '') skipped++; return; }
        var res = normalizeFor(input, raw);
        if (res.error) { bad++; return; }
        if (res.offList) offList++;
        var p = input.getAttribute('data-k').split('|');
        setValue(p[0], p[1], res.value);
        done++;
      });
    });
    pasteDone({ done: done, skipped: skipped, bad: bad, offList: offList, overflow: overflow }, 'beyond the last row of this sheet');
  }
  function pasteDone(n, overflowText) {
    recompute();
    var y = window.scrollY;
    renderPage();
    window.scrollTo(0, y);
    updateCounts();
    var msg = 'Pasted ' + n.done + ' cell' + (n.done === 1 ? '' : 's');
    var notes = [];
    if (n.skipped) notes.push(n.skipped + ' on automatic/label columns skipped');
    if (n.bad) notes.push(n.bad + ' not a valid date/number');
    if (n.offList) notes.push(n.offList + ' not in the dropdown list (kept, marked "not in list")');
    if (n.overflow) notes.push(n.overflow + ' ' + overflowText);
    toast(msg + (notes.length ? ' - ' + notes.join('; ') : ''), n.bad || n.overflow ? 'err' : 'ok');
  }

  // ------------------------------------------------------------------ actions
  function download(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  function fileSafe(s) { return String(s || '').replace(/[^A-Za-z0-9\u0600-\u06FF]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 40); }
  function stamp() { var d = new Date(); var p = function (x) { return x < 10 ? '0' + x : x; }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()); }
  function pickFile(accept, cb) {
    var input = document.getElementById('file-input');
    input.accept = accept;
    input.value = '';
    input.onchange = function () { if (input.files[0]) cb(input.files[0]); };
    input.click();
  }
  var busy = false;
  var actions = {
    export: async function () {
      if (busy) return;
      busy = true;
      toast('Building the Excel workbook\u2026');
      try {
        var blob = await IO.exportWorkbook(window.BCM_TEMPLATE_B64, INPUTS, function (sheet, addr) { return book.get(sheet, addr); });
        var bu = fileSafe(book.get(S.BUP, 'B7') || book.get(S.BUP, 'B6'));
        download(blob, 'BCM_Toolkit_v5.1' + (bu ? '_' + bu : '') + '_' + stamp() + '.xlsx');
        toast('Exported. Excel recalculates all formulas when you open the file.', 'ok');
      } catch (err) {
        console.error(err);
        toast('Export failed: ' + err.message, 'err');
      } finally { busy = false; }
    },
    import: function () {
      pickFile('.xlsx,.xlsm', async function (file) {
        try {
          var res = await IO.importWorkbook(await file.arrayBuffer(), INPUTS);
          var n = 0;
          Object.keys(res.values).forEach(function (s) { n += Object.keys(res.values[s]).length; });
          var known = Object.keys(res.values).length;
          if (!known) throw new Error('No toolkit sheets found in this file.');
          if (!confirm('Replace the data in this browser with ' + n + ' values from "' + file.name + '"?' +
            (res.missingSheets.length ? '\n\nSheets not found in the file (left empty): ' + res.missingSheets.join(', ') : ''))) return;
          // Template defaults (support lists, risk assumption map) stay where the file has none.
          state = freshState();
          Object.keys(res.values).forEach(function (sh) {
            var dst = state.values[sh] = state.values[sh] || {};
            Object.keys(res.values[sh]).forEach(function (a) { dst[a] = res.values[sh][a]; });
          });
          state.importedFrom = file.name;
          save();
          recompute();
          renderNav();
          renderPage();
          toast('Imported ' + n + ' values from ' + file.name, 'ok');
        } catch (err) {
          console.error(err);
          toast('Import failed: ' + err.message, 'err');
        }
      });
    },
    backup: function () {
      download(new Blob([JSON.stringify(state, null, 1)], { type: 'application/json' }), 'BCM_Toolkit_backup_' + stamp() + '.json');
    },
    restore: function () {
      pickFile('.json,application/json', function (file) {
        file.text().then(function (txt) {
          var s = JSON.parse(txt);
          if (!s || s.app !== 'bcm-toolkit' || !s.values) throw new Error('Not a BCM Toolkit backup file.');
          if (!confirm('Replace the data in this browser with the backup "' + file.name + '"?')) return;
          s.ui = s.ui || { rows: {} }; s.ui.rows = s.ui.rows || {};
          state = s;
          save();
          recompute();
          renderNav();
          renderPage();
          toast('Backup restored', 'ok');
        }).catch(function (err) { toast('Restore failed: ' + err.message, 'err'); });
      });
    },
    reset: function () {
      if (!confirm('Clear ALL data in this browser and start from the blank toolkit? Export or back up first if you need it.')) return;
      state = freshState();
      save();
      recompute();
      renderNav();
      renderPage();
      toast('Started a blank toolkit');
    },
    menu: function () { document.body.classList.toggle('nav-open'); },
    template: function () {
      var bin = atob(window.BCM_TEMPLATE_B64);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      download(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), window.BCM_TEMPLATE_NAME);
    }
  };

  // ------------------------------------------------------------------ boot
  function route() {
    var h = (location.hash || '').slice(1);
    var m = /^bia-(\d+)$/.exec(h);
    if (m) return;
    go(h || 'dash');
  }
  document.addEventListener('DOMContentLoaded', function () {
    recompute();
    var page = document.getElementById('page');
    page.addEventListener('input', onInput);
    page.addEventListener('change', onInput);
    page.addEventListener('focusin', onFocus);
    page.addEventListener('paste', onPaste);
    page.addEventListener('mousedown', function (e) { if (e.target.tagName === 'SELECT') onFocus(e); });
    document.addEventListener('click', onClick);
    window.addEventListener('hashchange', route);
    renderNav();
    route();
    setSaved(state.updated ? 'Saved in this browser' : 'Not saved yet');
  });

  window.BCM_APP = { getState: function () { return state; }, getCalc: function () { return calc; }, val: val, actions: actions };
})();
