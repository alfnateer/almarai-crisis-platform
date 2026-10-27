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
  var S = SC.S;
  var STORE_KEY = 'almarai-bcm-toolkit-v6.4';

  var INPUTS = SC.inputCells(T);
  var SPEC = {};
  INPUTS.forEach(function (c) { SPEC[c.sheet + '!' + c.addr] = c; });
  var SHEET_DEF = {};
  SC.sheets.forEach(function (d) { SHEET_DEF[d.sheet] = d; });

  // ------------------------------------------------------------------ state
  function freshState() {
    var values = {};
    INPUTS.forEach(function (c) {
      var def = SHEET_DEF[c.sheet];
      if (!def || !def.keepDefaults) return;
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

  var calc = EN.compute(book, { tpl: T });
  function val(sheet, addr) {
    var c = calc[sheet];
    if (c && addr in c) return c[addr];
    return book.get(sheet, addr);
  }
  var calcTimer = null;
  function scheduleCalc() {
    clearTimeout(calcTimer);
    calcTimer = setTimeout(function () { calc = EN.compute(book, { tpl: T }); refreshCalc(); }, 120);
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
  var RED = ['Critical', 'Overdue', 'Red - Action Required', 'DUPLICATE', 'Failed', 'GAP - no DR', 'Missing', 'No', 'Incomplete',
    'RTO Not Complied', 'RPO Not Complied', 'Test Overdue', 'Retest Required', 'Recovery Not Compliant', 'Not Evidenced', 'Ineffective', '#VALUE!', '#N/A'];
  var AMBER = ['Significant', 'Partial', 'Amber - Attention', 'In Progress', 'Incomplete Exercise', 'Result Required', 'Pass with Findings',
    'Not Tested', 'Recovery Targets Incomplete', 'Review Date Missing', 'Due \u226430 Days', 'Moderate', 'Planned', 'Scheduled', 'Not Scheduled', 'Under Evaluation'];
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
  function engineList(col) {
    var out = [], st = T.static[S.ENG] || {};
    for (var r = 4; r <= 12; r++) { var v = st[col + r]; if (v !== undefined && v !== '') out.push(String(v)); }
    return out;
  }
  var LISTS = {
    lstYN: engineList('A'), lstScore: engineList('C'), lstDepCat: engineList('D'), lstCorpCrit: engineList('E'),
    lstStrategy: engineList('H'), lstTestResult: engineList('I'), lstDRStrategy: engineList('J'), lstSPOF: engineList('L'),
    lstReqCat: engineList('M')
  };
  function nonEmpty(sheet, col, a, b) {
    var out = [];
    for (var r = a; r <= b; r++) { var v = book.get(sheet, col + r); if (v !== '') out.push(String(v)); }
    return out;
  }
  // Options for a dropdown (static list name, array, or dynamic source by row).
  function optionsFor(src, sheet, row) {
    if (Array.isArray(src)) return src;
    if (LISTS[src]) return LISTS[src];
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
        var idx = LISTS.lstDepCat.indexOf(cat);
        return idx < 0 ? [] : nonEmpty(S.MAP, SC.numToCol(3 + idx), 5, 24);
      case 'capaSourceRef':
        var t = book.get(S.CAPA, 'C' + row);
        return { Risk: L.riskIds, 'Recovery Plan': L.rrIds, 'IT DR': L.drIds, 'ISO Document': L.docIds, BIA: L.biaIds, Dependency: L.depIds }[t] || [];
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
      el.textContent = n;
    });
  }
  function go(id) {
    if (!SC.byId(id)) id = 'dash';
    current = id;
    document.querySelectorAll('.nav-item').forEach(function (a) { a.classList.toggle('active', a.getAttribute('data-id') === id); });
    document.body.classList.remove('nav-open');
    renderPage();
    document.getElementById('main').scrollTop = 0;
    window.scrollTo(0, 0);
  }

  // ------------------------------------------------------------------ pages
  function pageHeader(def, extra) {
    return '<header class="page-head"><div class="eyebrow">Sheet ' + esc(def.nav) + ' \u00b7 ' + esc(def.sheet) + '</div>' +
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
      case 'pdca': html = renderPdca(def); break;
      case 'isomap': html = renderIsoMap(def); break;
      case 'reference': html = pageHeader(def) + '<div class="card ref-wrap">' + (T.reference[def.sheet] || '') + '</div>'; break;
      case 'view': html = def.id === 'dash' ? renderDashboard(def) : (def.id === 'reg' ? renderRegister(def) : renderAnalysis(def)); break;
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
  function headerRows(cols, extraFirst) {
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
      bottom += '<th style="min-width:' + (c.w || 120) + 'px"' + (c.auto ? ' class="auto-h" title="Automatic (formula)"' : '') + '>' + esc(c.label) + '</th>';
    });
    bottom += '<th></th></tr>';
    return top + bottom;
  }
  function renderTable(def) {
    var rows = visibleRows(def);
    var html = pageHeader(def, legend());
    html += '<div class="card flush"><div class="table-wrap tall"><table class="sheet-table records" data-sheet="' + esc(def.id) + '"><thead>' +
      headerRows(def.cols, !!def.numCol) + '</thead><tbody>';
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
  function legend() {
    return '<div class="legend"><span><i class="sw white"></i>You fill these</span><span><i class="sw blue"></i>Automatic - calculated like the Excel formulas</span></div>';
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
      'MTPD is the first horizon where any impact reaches 4. Keep RTO below MTPD. Anchors are in 20. ERM RA Guidance.</p>');
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

  // ---- PDCA
  function renderPdca(def) {
    var html = pageHeader(def, '<p class="note">Auto Status is derived from the registers. Use Manual Override only with evidence; Final = override if set, otherwise auto.</p>');
    html += '<div class="card flush"><div class="table-wrap"><table class="sheet-table"><thead><tr><th>Phase</th><th>Step</th><th>BCMS Element</th><th class="auto-h">Auto Status</th><th>Manual Override</th><th class="auto-h">Final</th><th>What it means / evidence</th><th>Owner &amp; Next Action</th></tr></thead><tbody>';
    var phase = '';
    for (var r = def.rows.first; r <= def.rows.last; r++) {
      var a = book.get(def.sheet, 'A' + r);
      if (a) phase = a;
      html += '<tr><td class="label-cell phase-' + esc(String(phase).toLowerCase()) + '">' + esc(a) + '</td><td class="num">' + esc(book.get(def.sheet, 'B' + r)) + '</td>' +
        '<td class="label-cell">' + esc(book.get(def.sheet, 'C' + r)) + '</td>' + calcCell(def.sheet, 'D' + r, true) +
        '<td>' + inputHtml(def.sheet, 'E' + r, { type: 'select', options: SC.PDCA_OVERRIDE }, r) + '</td>' + calcCell(def.sheet, 'F' + r, true) +
        '<td class="fixed">' + esc(book.get(def.sheet, 'G' + r)) + '</td><td>' + inputHtml(def.sheet, 'H' + r, { type: 'textarea' }, r) + '</td></tr>';
    }
    return html + '</tbody></table></div></div>';
  }

  // ---- ISO 22301 map
  function renderIsoMap(def) {
    var sections = SC.isoSectionRows(T);
    var formulas = {};
    (T.formulas[def.sheet] || []).forEach(function (a) { formulas[a] = 1; });
    var html = pageHeader(def, legend()) + '<div class="card flush"><div class="table-wrap tall"><table class="sheet-table"><thead>' + headerRows(def.cols, false) + '</thead><tbody>';
    for (var r = def.rows.first; r <= def.rows.last; r++) {
      if (sections[r]) { html += '<tr class="section-row"><td colspan="' + (def.cols.length + 1) + '">' + esc(sections[r]) + '</td></tr>'; continue; }
      html += '<tr>';
      def.cols.forEach(function (c) {
        var addr = c.c + r;
        if (c.c === 'D' || (c.c === 'I' && formulas[addr])) html += calcCell(def.sheet, addr, c.c === 'D');
        else if (SPEC[def.sheet + '!' + addr]) html += '<td>' + inputHtml(def.sheet, addr, c, r) + '</td>';
        else html += '<td class="fixed">' + esc(fmt(book.get(def.sheet, addr))) + '</td>';
      });
      html += '<td></td></tr>';
    }
    return html + '</tbody></table></div></div>';
  }

  // ---- Dashboard
  function tile(label, value, sub, warn) {
    return '<div class="tile' + (warn ? ' warn' : '') + '"><div class="tile-label">' + esc(label) + '</div><div class="tile-value" data-dash>' + esc(value) + '</div>' + (sub ? '<div class="tile-sub">' + esc(sub) + '</div>' : '') + '</div>';
  }
  function renderDashboard(def) {
    var D = calc.dashboard;
    var bu = book.get(S.BUP, 'B6') || book.get(S.BUP, 'B5');
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
      tile('DR systems / tested-current', D.drSystems + ' / ' + D.drCurrent) + tile('External evidence not evidenced', D.extNotEvidenced, '', D.extNotEvidenced > 0) +
      tile('ISO clauses missing', D.isoMissing, '', D.isoMissing > 0) + tile('Duplicate S&S mappings', D.ssDuplicates, '', D.ssDuplicates > 0) + '</div></section>';
    html += '<section class="card"><h2>Corrective Action &amp; Improvement</h2><div class="tiles">' +
      tile('Open CAPAs', D.capaOpen) + tile('Overdue CAPAs', D.capaOverdue, '', D.capaOverdue > 0) +
      tile('Effectiveness pending', D.capaPending) + tile('Critical / High open', D.capaCritHigh, '', D.capaCritHigh > 0) + '</div></section>';
    html += '</div><div class="dash-grid two">';
    html += '<section class="card"><h2>PDCA snapshot</h2><table class="mini"><tbody>';
    for (var r = 5; r <= 19; r++) html += '<tr><td>' + esc(book.get(S.PDCA, 'C' + r)) + '</td><td>' + badge(val(S.PDCA, 'F' + r)) + '</td></tr>';
    html += '</tbody></table></section>';
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
    for (var l = 5; l >= 1; l--) {
      html += '<tr><th>' + lik[5 - l] + '</th>';
      for (var i = 1; i <= 5; i++) {
        var ref = Number(book.get(S.ERM, SC.numToCol(4 + i) + (84 - l)));
        var t = ref >= 20 ? 'red' : (ref >= 12 ? 'orange' : (ref >= 7 ? 'yellow' : 'green'));
        var n = counts[l + '-' + i] || 0;
        html += '<td class="h-' + t + '" title="Risk reference ' + ref + '">' + (n ? '<b>' + n + '</b>' : '') + '<span>' + ref + '</span></td>';
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
  function renderAnalysis(def) {
    var labels = { B: 'P&S ID', C: 'Business Unit', D: 'Division / Site', E: 'Product / Service', F: 'Corporate MBCO (if provided)', G: 'Corporate MTPD (if provided)',
      I: book.get(S.ANA, 'I1') || 'BIA MTPD', J: book.get(S.ANA, 'J1') || 'Dependency RTO', K: book.get(S.ANA, 'K1') || 'Dependency RPO' };
    var cols = ['B', 'C', 'D', 'E', 'F', 'G', 'I', 'J', 'K'];
    var html = pageHeader(def, '<p class="note">Read-only. Columns J and K reproduce the workbook formulas, which read the first dependency row of each P&amp;S in 08. Dependency Register (SPOF and Owner columns).</p>');
    html += '<div class="card flush"><div class="table-wrap tall"><table class="sheet-table"><thead><tr><th class="num">Sl #</th>';
    cols.forEach(function (c) { html += '<th style="min-width:120px">' + esc(labels[c]) + '</th>'; });
    html += '</tr></thead><tbody>';
    if (!calc.analysis.length) html += '<tr><td colspan="10" class="muted">No in-scope products / services yet.</td></tr>';
    calc.analysis.forEach(function (row, i) {
      html += '<tr><td class="num">' + (i + 1) + '</td>';
      cols.forEach(function (c) { html += '<td class="auto">' + esc(fmt(row[c])) + '</td>'; });
      html += '</tr>';
    });
    return html + '</tbody></table></div></div>';
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
          var msg = validate(def, c, r);
          el.classList.toggle('invalid', !!msg);
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
      document.querySelectorAll('.bia-block input[type=date]').forEach(function (el) {
        var s = EN.serial(el.value);
        var bad = s !== null && s > today();
        el.classList.toggle('invalid', bad);
        el.title = bad ? 'Assessment date cannot be in the future' : '';
      });
    }
  }
  function cssEsc(s) { return window.CSS && CSS.escape ? CSS.escape(s) : s.replace(/["\\]/g, '\\$&'); }
  function autosize(el) {
    if (el.tagName !== 'TEXTAREA') return;
    el.style.height = 'auto';
    el.style.height = Math.min(el.scrollHeight + 2, 240) + 'px';
  }
  function autosizeAll(root) { root.querySelectorAll('textarea').forEach(autosize); }

  // ------------------------------------------------------------------ events
  function onInput(e) {
    var el = e.target;
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
      calc = EN.compute(book, { tpl: T });
      var yy = window.scrollY;
      renderPage();
      window.scrollTo(0, yy);
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
        download(blob, 'BCM_Implementation_Toolkit_v6.4' + (bu ? '_' + bu : '') + '_' + stamp() + '.xlsx');
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
          state = freshState();
          state.values = res.values;
          state.importedFrom = file.name;
          save();
          calc = EN.compute(book, { tpl: T });
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
          calc = EN.compute(book, { tpl: T });
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
      calc = EN.compute(book, { tpl: T });
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
    var page = document.getElementById('page');
    page.addEventListener('input', onInput);
    page.addEventListener('change', onInput);
    page.addEventListener('focusin', onFocus);
    page.addEventListener('mousedown', function (e) { if (e.target.tagName === 'SELECT') onFocus(e); });
    document.addEventListener('click', onClick);
    window.addEventListener('hashchange', route);
    renderNav();
    route();
    setSaved(state.updated ? 'Saved in this browser' : 'Not saved yet');
  });

  window.BCM_APP = { getState: function () { return state; }, getCalc: function () { return calc; }, actions: actions };
})();
