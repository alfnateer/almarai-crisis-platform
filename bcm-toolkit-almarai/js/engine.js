/*
 * Calculation engine - JavaScript ports of the workbook formulas, so the website
 * can show the same automatic values (IDs, ratings, statuses, dashboard) live.
 * Excel remains the source of truth: exported files are recalculated on open.
 */
(function (root) {
  'use strict';

  var SCHEMA = root.BCM_SCHEMA || (typeof require !== 'undefined' ? require('./schema.js') : null);
  var S = SCHEMA.S, BIA = SCHEMA.BIA;

  // ---------------------------------------------------------------- helpers
  function pad3(n) { n = String(n); while (n.length < 3) n = '0' + n; return n; }
  function blank(v) { return v === undefined || v === null || v === ''; }
  function str(v) { return blank(v) ? '' : String(v); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function num(v) {
    if (isNum(v)) return v;
    if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return Number(v);
    return null;
  }
  // Dates are held as ISO strings (yyyy-mm-dd); serial = days since epoch.
  function serial(v) {
    if (typeof v !== 'string') return null;
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
    if (!m) return null;
    return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5;
  }
  function isoFromSerial(s) {
    var d = new Date(s * 864e5);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }
  function pad2(n) { return n < 10 ? '0' + n : String(n); }
  function todaySerial() {
    var d = new Date();
    return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5;
  }
  // EDATE(d, months): same day n months later, clamped to month end.
  function edate(iso, months) {
    var s = serial(iso);
    if (s === null) return '';
    var d = new Date(s * 864e5);
    var y = d.getUTCFullYear(), m = d.getUTCMonth() + months, day = d.getUTCDate();
    y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
    var last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    return y + '-' + pad2(m + 1) + '-' + pad2(Math.min(day, last));
  }
  // DATE(YEAR+1, MONTH, DAY): rolls over (29 Feb -> 1 Mar).
  function addYearRoll(iso) {
    var s = serial(iso);
    if (s === null) return '';
    var d = new Date(s * 864e5);
    return isoFromSerial(Date.UTC(d.getUTCFullYear() + 1, d.getUTCMonth(), d.getUTCDate()) / 864e5);
  }
  // "3 days" -> 72 (Excel LET/TEXTBEFORE/SWITCH duration parser).
  function durationHours(v) {
    if (blank(v)) return '';
    var s = String(v);
    var i = s.indexOf(' ');
    if (i < 0) return '#VALUE!';
    var n = Number(s.slice(0, i));
    if (s.slice(0, i).trim() === '' || isNaN(n)) return '#VALUE!';
    var unit = s.slice(i + 1).toLowerCase();
    var f = { hour: 1, hours: 1, day: 24, days: 24, week: 168, weeks: 168, month: 24 * 365.25 / 12, months: 24 * 365.25 / 12 }[unit];
    return f === undefined ? '' : n * f;
  }
  function uniq(arr) {
    var seen = {}, out = [];
    arr.forEach(function (x) { var k = String(x); if (!seen[k]) { seen[k] = 1; out.push(x); } });
    return out;
  }
  function depKey(s) {
    var k = String(s || '').replace(/ /g, '').replace(/-/g, '').replace(/\//g, '').toLowerCase();
    return k === 'data' ? 'datainfo' : k;
  }
  function wild(prefix) { return function (v) { return typeof v === 'string' && v.indexOf(prefix) === 0; }; }
  function count(arr, pred) { var n = 0; arr.forEach(function (x) { if (pred(x)) n++; }); return n; }
  function range(a, b) { var o = []; for (var i = a; i <= b; i++) o.push(i); return o; }

  // ---------------------------------------------------------------- engine
  /*
   * book.get(sheet, addr) returns the stored input value, or the template's static
   * value for non-input cells (labels). compute() returns calc[sheet][addr].
   */
  function compute(book, opts) {
    opts = opts || {};
    var T = opts.today !== undefined ? opts.today : todaySerial();
    var calc = {};
    Object.keys(S).forEach(function (k) { calc[S[k]] = {}; });
    function g(sheet, addr) { var c = calc[sheet]; if (c && addr in c) return c[addr]; var v = book.get(sheet, addr); return blank(v) ? '' : v; }
    function set(sheet, addr, v) { calc[sheet][addr] = v; }
    function eqStr(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }

    // ---- 03 BU Profile
    set(S.BUP, 'D119', blank(g(S.BUP, 'B119')) ? '' : edate(g(S.BUP, 'B119'), 12));

    // ---- 04 Products & Services (B, M, N)
    var PS_ROWS = range(5, 49);
    PS_ROWS.forEach(function (r) {
      var a = g(S.PS, 'A' + r);
      set(S.PS, 'B' + r, blank(g(S.PS, 'E' + r)) ? '' : 'PS-' + pad3(a));
      set(S.PS, 'M' + r, durationHours(g(S.PS, 'L' + r)));
    });
    PS_ROWS.forEach(function (r) {
      var a = num(g(S.PS, 'A' + r));
      if (g(S.PS, 'I' + r) !== 'Yes' || a === null) { set(S.PS, 'N' + r, ''); return; }
      var n = count(PS_ROWS, function (q) { var aq = num(g(S.PS, 'A' + q)); return aq !== null && aq < a && g(S.PS, 'I' + q) === 'Yes'; });
      set(S.PS, 'N' + r, n + 1);
    });
    function psRowByName(name) {
      for (var i = 0; i < PS_ROWS.length; i++) if (eqStr(g(S.PS, 'E' + PS_ROWS[i]), name)) return PS_ROWS[i];
      return null;
    }
    function psRowById(id) {
      if (blank(id)) return null;
      for (var i = 0; i < PS_ROWS.length; i++) if (eqStr(g(S.PS, 'B' + PS_ROWS[i]), id)) return PS_ROWS[i];
      return null;
    }

    // ---- 07 BIA blocks
    var BIA_ROWS = range(6, 285);
    function blockMaxFirstCol(t) {
      // MTPD: first horizon (J..Q) whose max across the block reaches 4.
      var any = false;
      for (var r = t; r < t + 7; r++) BIA.scoreCols.forEach(function (c) { if (num(g(S.BIA, c + r)) !== null) any = true; });
      if (!any) return '';
      for (var j = 0; j < 8; j++) {
        var mx = null;
        for (var q = t; q < t + 7; q++) { var x = num(g(S.BIA, BIA.scoreCols[j] + q)); if (x !== null && (mx === null || x > mx)) mx = x; }
        if (mx !== null && mx >= 4) return BIA.horizonHours[j];
      }
      return '';
    }
    for (var k = 1; k <= BIA.blocks; k++) {
      var t = BIA.top(k);
      var psRow = null;
      PS_ROWS.forEach(function (r) { if (psRow === null && g(S.PS, 'N' + r) === k) psRow = r; });
      var name = psRow ? str(g(S.PS, 'E' + psRow)) : '';
      set(S.BIA, 'C' + t, psRow ? str(g(S.PS, 'C' + psRow)) : '');
      set(S.BIA, 'D' + t, psRow ? str(g(S.PS, 'D' + psRow)) : '');
      set(S.BIA, 'E' + t, name);
      set(S.BIA, 'B' + t, name === '' ? '' : 'BIA-' + pad3(k));
      var byName = name === '' ? null : psRowByName(name);
      set(S.BIA, 'F' + t, byName ? str(g(S.PS, 'B' + byName)) : '');
      var scores = [];
      // Max impact and impact type look at J..Q (the "More than 1 Month" column is excluded).
      var SQ = BIA.scoreCols.slice(0, 8);
      for (var r = t; r < t + 7; r++) SQ.forEach(function (c) { var x = num(g(S.BIA, c + r)); if (x !== null) scores.push(x); });
      var smax = scores.length ? Math.max.apply(null, scores) : '';
      set(S.BIA, 'S' + t, smax);
      for (var rr = t; rr < t + 7; rr++) {
        var rowScores = [];
        SQ.forEach(function (c) { var x = num(g(S.BIA, c + rr)); if (x !== null) rowScores.push(x); });
        set(S.BIA, 'T' + rr, smax === '' || !rowScores.length ? '' : (Math.max.apply(null, rowScores) > 3 ? str(g(S.BIA, 'I' + rr)) : ''));
      }
      set(S.BIA, 'AA' + t, name === '' ? '' : blockMaxFirstCol(t));
      set(S.BIA, 'AD' + t, durationHours(g(S.BIA, 'AC' + t)));
      set(S.BIA, 'AG' + t, blank(g(S.BIA, 'AF' + t)) ? '' : addYearRoll(g(S.BIA, 'AF' + t)));
    }
    // U, W, X, AB need the block names first.
    for (k = 1; k <= BIA.blocks; k++) {
      t = BIA.top(k);
      var nm = g(S.BIA, 'E' + t);
      if (nm === '') { ['U', 'W', 'X', 'AB'].forEach(function (c) { set(S.BIA, c + t, ''); }); continue; }
      var matches = PS_ROWS.filter(function (r) { return eqStr(g(S.PS, 'E' + r), nm); }).map(function (r) { return g(S.PS, 'J' + r); });
      var nth = 0;
      for (var b = 1; b <= k; b++) if (eqStr(g(S.BIA, 'E' + BIA.top(b)), nm)) nth++;
      var u = matches[nth - 1];
      u = (blank(u) || u === 0) ? '' : u;
      set(S.BIA, 'U' + t, u);
      var v = g(S.BIA, 'V' + t);
      set(S.BIA, 'W' + t, (u === 'Yes' || v === 'Yes') ? 'Yes' : ((u === 'No' && v === 'No') ? 'No' : ((u === 'Under Evaluation' && v === 'No') ? 'Under Evaluation' : '')));
      var pr = psRowByName(nm);
      set(S.BIA, 'X' + t, pr ? str(g(S.PS, 'K' + pr)) : '');
      var f = g(S.BIA, 'F' + t);
      var pi = psRowById(f);
      set(S.BIA, 'AB' + t, f === '' || !pi ? '' : g(S.PS, 'M' + pi));
    }
    // P&S O: BU BIA MTPD from the matching BIA block.
    PS_ROWS.forEach(function (r) {
      var nm2 = g(S.PS, 'E' + r);
      if (nm2 === '') { set(S.PS, 'O' + r, ''); return; }
      var hit = null;
      for (var i = 0; i < BIA_ROWS.length; i++) if (eqStr(g(S.BIA, 'E' + BIA_ROWS[i]), nm2)) { hit = BIA_ROWS[i]; break; }
      set(S.PS, 'O' + r, hit === null ? '' : blockMaxFirstCol(hit));
    });
    function biaTopById(id) {
      for (var b = 1; b <= BIA.blocks; b++) if (!blank(id) && eqStr(g(S.BIA, 'B' + BIA.top(b)), id)) return BIA.top(b);
      return null;
    }
    function biaTopByPs(psid) {
      for (var b = 1; b <= BIA.blocks; b++) if (!blank(psid) && eqStr(g(S.BIA, 'F' + BIA.top(b)), psid)) return BIA.top(b);
      return null;
    }
    // Fill-down (Excel SCAN) of the merged block columns for row-level filters.
    function scanDown(col) {
      var acc = '', out = {};
      BIA_ROWS.forEach(function (r) { var x = g(S.BIA, col + r); if (x !== '') acc = x; out[r] = acc; });
      return out;
    }
    var scanB = scanDown('B'), scanC = scanDown('C'), scanD = scanDown('D'), scanE = scanDown('E');

    // ---- 05 Processes
    range(5, 211).forEach(function (r) {
      var code = g(S.PROC, 'B' + r);
      var p = psRowById(code);
      set(S.PROC, 'C' + r, p ? str(g(S.PS, 'E' + p)) : '');
      var bt = code === '' ? null : biaTopByPs(code);
      var aa = bt ? g(S.BIA, 'AA' + bt) : '';
      var ad = bt ? g(S.BIA, 'AD' + bt) : '';
      set(S.PROC, 'J' + r, code === '' ? '' : (aa === 0 ? '' : aa));
      set(S.PROC, 'K' + r, code === '' ? '' : (ad === 0 ? '' : ad));
    });

    // ---- 08 Dependency Register
    var DEP_ROWS = range(5, 209);
    DEP_ROWS.forEach(function (r) {
      set(S.DEP, 'B' + r, blank(g(S.DEP, 'I' + r)) ? '' : 'DEP-' + pad3(g(S.DEP, 'A' + r)));
      var c = g(S.DEP, 'C' + r);
      var bt = c === '' ? null : biaTopById(c);
      set(S.DEP, 'D' + r, bt ? str(g(S.BIA, 'E' + bt)) : '');
      set(S.DEP, 'E' + r, bt ? str(g(S.BIA, 'F' + bt)) : '');
    });

    // ---- 10 Risk Assessment (part 1)
    var RA_ROWS = range(5, 61);
    var LIK = ['Rare', 'Infrequent', 'Occasional', 'Frequent', 'Almost Certain'];
    var IMP = ['Negligible', 'Minor', 'Medium', 'Critical', 'Catastrophic'];
    function heat(l, i) {
      // '20. ERM RA Guidance'!E79:I83 - row 79 is likelihood 5.
      var row = 79 + (5 - l), col = SCHEMA.numToCol(4 + i);
      var x = num(book.get(S.ERM, col + row));
      return x === null ? '' : x;
    }
    function applicable(bu, div, dep, field) {
      // RA col F (field = scanE, names) and Risk Register col I (field = scanB, BIA ids).
      if (bu === '') return '';
      var out = [];
      if (div === '') {
        BIA_ROWS.forEach(function (r) { if (field[r] !== '' && eqStr(scanC[r], bu)) out.push(field[r]); });
      } else {
        if (dep === '') return '';
        var key = depKey(dep);
        BIA_ROWS.forEach(function (r) {
          if (field[r] !== '' && eqStr(scanC[r], bu) && eqStr(scanD[r], div) && depKey(g(S.BIA, 'G' + r)) === key && g(S.BIA, 'H' + r) === 'Yes') out.push(field[r]);
        });
      }
      return uniq(out).join(',');
    }
    RA_ROWS.forEach(function (r) {
      var G = g(S.RA, 'G' + r);
      set(S.RA, 'B' + r, G === '' ? '' : 'RA-' + pad3(g(S.RA, 'A' + r)));
      set(S.RA, 'F' + r, applicable(g(S.RA, 'C' + r), g(S.RA, 'D' + r), g(S.RA, 'E' + r), scanE));
      set(S.RA, 'H' + r, G === '' ? '' : 'Inability to deliver the associated products/services due to [' + G + ']');
      var K = num(g(S.RA, 'K' + r)), M = num(g(S.RA, 'M' + r));
      set(S.RA, 'L' + r, K === null ? '' : (LIK[K - 1] || '#VALUE!'));
      set(S.RA, 'N' + r, M === null ? '' : (IMP[M - 1] || '#VALUE!'));
      var O = (K === null || M === null) ? '' : heat(K, M);
      set(S.RA, 'O' + r, O);
      var P = O === '' ? '' : (O >= 20 ? 'Critical' : (O >= 12 ? 'Significant' : (O >= 7 ? 'Moderate' : 'Low')));
      set(S.RA, 'P' + r, P);
      set(S.RA, 'Q' + r, P === '' ? '' : ({ Critical: 1, Significant: 2, Moderate: 3 }[P] || 4));
      // "Risks": New when there is no ERM reference, otherwise Existing (every table row).
      set(S.RA, 'U' + r, g(S.RA, 'I' + r) === 'No applicable ERM reference' ? 'New' : 'Existing');
    });
    // A risk record is complete when BU, Division, Dependency, Applicable P&S, Risk Assumption,
    // ERM ref, Controls, Likelihood and Impact are filled (template v0.2.1).
    var RA_REQUIRED = ['C', 'D', 'E', 'F', 'G', 'I', 'J', 'K', 'M'];
    function raComplete(r) { return RA_REQUIRED.every(function (c) { return g(S.RA, c + r) !== ''; }); }
    function raRowById(id) {
      if (blank(id)) return null;
      for (var i = 0; i < RA_ROWS.length; i++) if (eqStr(g(S.RA, 'B' + RA_ROWS[i]), id)) return RA_ROWS[i];
      return null;
    }

    // ---- 14 R&R ids, 11 S&S, 12 Requirements, then statuses
    var RR_ROWS = range(5, 104), SS_ROWS = range(5, 64), REQ_ROWS = range(5, 304);
    RR_ROWS.forEach(function (r) { set(S.RR, 'B' + r, blank(g(S.RR, 'C' + r)) ? '' : 'RR-' + pad3(g(S.RR, 'A' + r))); });
    SS_ROWS.forEach(function (r) {
      var B = g(S.SS, 'B' + r);
      set(S.SS, 'C' + r, B === '' ? '' : 'SS-' + pad3(g(S.SS, 'A' + r)));
      var ra = raRowById(B);
      set(S.SS, 'D' + r, B === '' ? '' : (ra ? str(g(S.RA, 'E' + ra)) : ''));
      set(S.SS, 'E' + r, B === '' ? '' : (ra ? str(g(S.RA, 'G' + ra)) : ''));
      set(S.SS, 'F' + r, B === '' ? '' : (ra ? str(g(S.RA, 'H' + ra)) : ''));
      set(S.SS, 'G' + r, B === '' ? '' : (ra ? str(g(S.RA, 'F' + ra)) : ''));
    });
    SS_ROWS.forEach(function (r) {
      var B = g(S.SS, 'B' + r);
      if (B === '') { set(S.SS, 'U' + r, ''); return; }
      var n = count(SS_ROWS, function (q) { return eqStr(g(S.SS, 'B' + q), B) && eqStr(g(S.SS, 'G' + q), g(S.SS, 'G' + r)) && eqStr(g(S.SS, 'H' + q), g(S.SS, 'H' + r)); });
      set(S.SS, 'U' + r, n > 1 ? 'DUPLICATE' : 'OK');
    });
    function ssRowById(id) {
      if (blank(id)) return null;
      for (var i = 0; i < SS_ROWS.length; i++) if (eqStr(g(S.SS, 'C' + SS_ROWS[i]), id)) return SS_ROWS[i];
      return null;
    }
    function rrIdsForSs(ssList) {
      return uniq(RR_ROWS.filter(function (q) {
        var b = g(S.RR, 'B' + q);
        return b !== '' && ssList.some(function (s) { return eqStr(s, g(S.RR, 'C' + q)); });
      }).map(function (q) { return g(S.RR, 'B' + q); }));
    }
    SS_ROWS.forEach(function (r) {
      var C = g(S.SS, 'C' + r);
      if (C === '') { set(S.SS, 'V' + r, ''); return; }
      var ids = rrIdsForSs([C]);
      set(S.SS, 'V' + r, ids.length ? ids.join(',') : 'Not Created');
    });
    REQ_ROWS.forEach(function (r) {
      set(S.REQ, 'B' + r, blank(g(S.REQ, 'F' + r)) ? '' : 'REQ-' + pad3(g(S.REQ, 'A' + r)));
      var C = g(S.REQ, 'C' + r);
      var sr = ssRowById(C);
      set(S.REQ, 'D' + r, C === '' ? '' : (sr ? str(g(S.SS, 'B' + sr)) : ''));
    });
    var FLAG_COLS = ['J', 'K', 'L', 'M', 'N', 'O'];
    var FLAG_CATS = ['People', 'ICT', '3rd Party', 'Data Backup', 'Equipment / Building / Utilities', 'Fund'];
    SS_ROWS.forEach(function (r) {
      var B = g(S.SS, 'B' + r), C = g(S.SS, 'C' + r);
      if (B === '') { set(S.SS, 'T' + r, ''); return; }
      var yes = 0, ok = 0;
      FLAG_COLS.forEach(function (fc, i) {
        if (g(S.SS, fc + r) !== 'Yes') return;
        yes++;
        var total = count(REQ_ROWS, function (q) { return eqStr(g(S.REQ, 'C' + q), C) && eqStr(g(S.REQ, 'E' + q), FLAG_CATS[i]); });
        var ready = count(REQ_ROWS, function (q) { return eqStr(g(S.REQ, 'C' + q), C) && eqStr(g(S.REQ, 'E' + q), FLAG_CATS[i]) && g(S.REQ, 'L' + q) === 'Ready'; });
        if (total > 0 && ready === total) ok++;
      });
      var reqOK = ok === yes;
      var V = g(S.SS, 'V' + r);
      var core = ['H', 'I', 'P', 'Q', 'R'].every(function (c) { return g(S.SS, c + r) !== ''; }) && V !== '' && V !== 'Not Created';
      var Sd = serial(g(S.SS, 'S' + r)), Q = serial(g(S.SS, 'Q' + r)), R = serial(g(S.SS, 'R' + r));
      var st;
      if (g(S.SS, 'S' + r) !== '') st = (Sd !== null && Sd <= T && core && reqOK) ? 'Completed' : 'Incomplete';
      else if (Q === null) st = 'Planned';
      else if (T < Q) st = 'Planned';
      else if (R === null) st = 'In Progress';
      else st = T <= R ? 'In Progress' : 'Overdue';
      set(S.SS, 'T' + r, st);
    });
    RR_ROWS.forEach(function (r) {
      var C = g(S.RR, 'C' + r);
      var sr = ssRowById(C);
      function ss(col) { return C === '' ? '' : (sr ? str(g(S.SS, col + sr)) : ''); }
      set(S.RR, 'D' + r, ss('G')); set(S.RR, 'G' + r, ss('D')); set(S.RR, 'H' + r, ss('E'));
      set(S.RR, 'I' + r, ss('F')); set(S.RR, 'J' + r, ss('H')); set(S.RR, 'K' + r, ss('I'));
      set(S.RR, 'O' + r, blank(g(S.RR, 'M' + r)) ? '' : edate(g(S.RR, 'M' + r), 12));
    });
    RR_ROWS.forEach(function (r) {
      var C = g(S.RR, 'C' + r);
      if (C === '') { set(S.RR, 'Q' + r, ''); return; }
      var n = count(RR_ROWS, function (q) { return eqStr(g(S.RR, 'C' + q), C) && eqStr(g(S.RR, 'D' + q), g(S.RR, 'D' + r)) && eqStr(g(S.RR, 'E' + q), g(S.RR, 'E' + r)); });
      set(S.RR, 'Q' + r, n > 1 ? 'DUPLICATE' : 'OK');
    });
    RR_ROWS.forEach(function (r) {
      var C = g(S.RR, 'C' + r);
      if (C === '') { set(S.RR, 'P' + r, ''); return; }
      var sr = ssRowById(C);
      var ssStatus = sr ? g(S.SS, 'T' + sr) : '';
      var L = serial(g(S.RR, 'L' + r)), N = g(S.RR, 'N' + r);
      var st;
      if (g(S.RR, 'L' + r) === '') st = 'Not Scheduled';
      else if (g(S.RR, 'M' + r) === '') st = (L !== null && L < T) ? 'Overdue' : 'Scheduled';
      else if (N === '') st = 'Result Required';
      else if (g(S.RR, 'E' + r) === '' || g(S.RR, 'F' + r) === '') st = 'Incomplete Exercise';
      else if (/^#/.test(g(S.RR, 'Q' + r))) st = g(S.RR, 'Q' + r);
      else if (g(S.RR, 'Q' + r) === '') st = 'Incomplete Exercise';
      else if (N === 'Fail') st = 'Failed';
      else if (N === 'Pass with Findings' || ssStatus !== 'Completed') st = 'Pass with Findings';
      else st = 'Passed';
      set(S.RR, 'P' + r, st);
    });
    // RA S, T
    function ssIdsForRisk(id) {
      return uniq(SS_ROWS.filter(function (q) { return eqStr(g(S.SS, 'B' + q), id) && g(S.SS, 'C' + q) !== ''; }).map(function (q) { return g(S.SS, 'C' + q); }));
    }
    RA_ROWS.forEach(function (r) {
      var B = g(S.RA, 'B' + r);
      if (B === '') { set(S.RA, 'S' + r, ''); set(S.RA, 'T' + r, ''); return; }
      var ss = ssIdsForRisk(B);
      set(S.RA, 'S' + r, ss.length ? ss.join(',') : 'Not Created');
      var rr = rrIdsForSs(ss);
      set(S.RA, 'T' + r, rr.length ? rr.join(',') : 'Not Created');
    });

    // ---- 15 BCM Risk Register (spilled view)
    var regIds = RA_ROWS.map(function (r) { return g(S.RA, 'B' + r); }).filter(function (x) { return x !== ''; });
    var register = [];
    var SS_ORD = ['Completed', 'Planned', 'In Progress', 'Incomplete', 'Overdue'];
    var EX_ORD = ['Passed', 'Scheduled', 'Not Scheduled', 'Pass with Findings', 'Result Required', 'Incomplete Exercise', 'Overdue', 'Failed'];
    function worst(list, ord) {
      var mx = 0;
      list.forEach(function (s) { var i = ord.indexOf(s); mx = Math.max(mx, i < 0 ? 4 : i + 1); });
      return ord[mx - 1];
    }
    regIds.forEach(function (id, idx) {
      var r = raRowById(id);
      var row = { A: id };
      row.B = str(g(S.RA, 'G' + r)); row.C = str(g(S.RA, 'H' + r)); row.D = str(g(S.RA, 'E' + r));
      row.E = str(g(S.RA, 'L' + r)); row.F = str(g(S.RA, 'N' + r)); row.G = str(g(S.RA, 'P' + r)); row.H = g(S.RA, 'Q' + r);
      row.I = applicable(g(S.RA, 'C' + r), g(S.RA, 'D' + r), g(S.RA, 'E' + r), scanB);
      var names = row.I === '' ? [] : row.I.split(',').map(function (x) { var bt = biaTopById(x.trim()); return bt ? g(S.BIA, 'E' + bt) : ''; });
      row.J = row.I === '' ? '' : uniq(names.filter(function (x) { return x !== ''; })).join(', ');
      var ss = ssIdsForRisk(id);
      var rr = rrIdsForSs(ss);
      row.K = rr.length ? rr.join(',') : 'Not Created';
      row.L = ss.length ? ss.join(',') : 'Not Created';
      var ssStat = SS_ROWS.filter(function (q) { return eqStr(g(S.SS, 'B' + q), id) && g(S.SS, 'C' + q) !== ''; }).map(function (q) { return g(S.SS, 'T' + q); });
      row.M = ssStat.length ? worst(ssStat, SS_ORD) : 'Not Created';
      var dues = SS_ROWS.filter(function (q) { return eqStr(g(S.SS, 'B' + q), id) && g(S.SS, 'C' + q) !== '' && g(S.SS, 'T' + q) !== 'Completed' && serial(g(S.SS, 'R' + q)) !== null; })
        .map(function (q) { return serial(g(S.SS, 'R' + q)); });
      row.N = dues.length ? isoFromSerial(Math.min.apply(null, dues)) : '';
      var exStat = RR_ROWS.filter(function (q) { return g(S.RR, 'B' + q) !== '' && ss.some(function (s) { return eqStr(s, g(S.RR, 'C' + q)); }); }).map(function (q) { return g(S.RR, 'P' + q); });
      row.O = exStat.length ? worst(exStat, EX_ORD) : 'Not Created';
      row.R = str(g(S.RA, 'R' + r));
      // S: Risks (New / Existing); T: Source Completion (template v0.3).
      row.S = str(g(S.RA, 'U' + r));
      row.T = raComplete(r) ? 'Complete' : 'Incomplete';
      var critical = row.G === 'Critical' || row.H === 1;
      var unrated = row.G === '';
      var dueOver = row.N !== '' && serial(row.N) < T;
      if (row.T === 'Incomplete') {
        row.P = 'Red - Action Required';
        row.Q = 'ESCALATE: Source risk record incomplete';
      } else {
        if (unrated || row.M === 'Overdue' || row.O === 'Failed' || row.O === 'Overdue' || (critical && (row.M === 'Not Created' || row.O === 'Not Created'))) row.P = 'Red - Action Required';
        else if (['Not Created', 'Incomplete', 'In Progress'].indexOf(row.M) >= 0 || ['Not Created', 'Result Required', 'Not Scheduled'].indexOf(row.O) >= 0 || dueOver) row.P = 'Amber - Attention';
        else row.P = 'Green - On Track';
        if (unrated && row.M === 'Not Created') row.Q = 'ESCALATE: Unrated risk has no strategy';
        else if (critical && row.M === 'Not Created') row.Q = 'ESCALATE: Critical risk has no strategy';
        else if (critical && row.O === 'Not Created') row.Q = 'ESCALATE: Critical risk has no recovery plan';
        else if (row.M === 'Overdue') row.Q = 'ESCALATE: Strategy overdue';
        else if (dueOver) row.Q = 'ESCALATE: Strategy action overdue';
        else if (row.O === 'Overdue') row.Q = 'ESCALATE: Exercise overdue';
        else if (row.O === 'Failed') row.Q = 'ESCALATE: Exercise failed';
        else if (row.O === 'Result Required') row.Q = 'ESCALATE: Exercise result required';
        else if (row.O === 'Not Scheduled' || row.O === 'Not Created') row.Q = 'REVIEW: Exercise not scheduled';
        else if (row.P === 'Amber - Attention') row.Q = 'REVIEW: Treatment needs attention';
        else row.Q = 'MONITOR';
      }
      register.push(row);
      Object.keys(row).forEach(function (c) { set(S.REG, c + (6 + idx), row[c]); });
    });

    // ---- 07 BIA linked records (from the register)
    for (k = 1; k <= BIA.blocks; k++) {
      t = BIA.top(k);
      var bid = g(S.BIA, 'B' + t);
      ['AH', 'AI', 'AJ'].forEach(function (c, i) {
        if (bid === '') { set(S.BIA, c + t, ''); return; }
        var src = ['A', 'L', 'K'][i];
        var hits = register.filter(function (row) { return (',' + row.I.replace(/ /g, '') + ',').indexOf(',' + bid + ',') >= 0; }).map(function (row) { return row[src]; }).filter(function (x) { return x !== '' && x !== 'Not Created'; });
        set(S.BIA, c + t, hits.length ? uniq(hits).join(', ') : 'Not Created');
      });
    }

    // ---- 13 IT DR Register
    range(5, 204).forEach(function (r) {
      set(S.DR, 'A' + r, r - 4);
      set(S.DR, 'B' + r, (blank(g(S.DR, 'C' + r)) && blank(g(S.DR, 'D' + r))) ? '' : 'DR-' + pad3(r - 4));
      set(S.DR, 'P' + r, blank(g(S.DR, 'M' + r)) ? '' : edate(g(S.DR, 'M' + r), 12));
      var E = g(S.DR, 'E' + r);
      var G = num(g(S.DR, 'G' + r)), H = num(g(S.DR, 'H' + r)), I = num(g(S.DR, 'I' + r)), J = num(g(S.DR, 'J' + r));
      var Sx = E === '' ? '' : ((G === null || I === null) ? 'Not Assessed' : (I <= G ? 'RTO Complied' : 'RTO Not Complied'));
      var Tx = E === '' ? '' : ((H === null || J === null) ? 'Not Assessed' : (J <= H ? 'RPO Complied' : 'RPO Not Complied'));
      set(S.DR, 'S' + r, Sx); set(S.DR, 'T' + r, Tx);
      var P = serial(g(S.DR, 'P' + r)), O = g(S.DR, 'O' + r);
      var q;
      if (E === '') q = '';
      else if (g(S.DR, 'K' + r) === 'None (gap)') q = 'GAP - no DR';
      else if (Sx === 'Not Assessed' || Tx === 'Not Assessed') q = 'Recovery Targets Incomplete';
      else if (g(S.DR, 'M' + r) === '') q = 'Not Tested';
      else if (P !== null && P < T) q = 'Test Overdue';
      else if (O === '') q = 'Not Tested';
      else if (O !== 'Pass') q = 'Retest Required';
      else if (Sx !== 'RTO Complied' || Tx !== 'RPO Complied') q = 'Recovery Not Compliant';
      else q = 'Tested & Current';
      set(S.DR, 'Q' + r, q);
    });

    // ---- 16 CAPA
    var CAPA_ROWS = range(6, 105);
    CAPA_ROWS.forEach(function (r) {
      set(S.CAPA, 'B' + r, blank(g(S.CAPA, 'F' + r)) ? '' : 'CAPA-' + pad3(g(S.CAPA, 'A' + r)));
      var E = serial(g(S.CAPA, 'E' + r)), R = serial(g(S.CAPA, 'R' + r));
      set(S.CAPA, 'S' + r, g(S.CAPA, 'E' + r) === '' ? '' : (E === null ? '#VALUE!' : (g(S.CAPA, 'R' + r) !== '' ? (R === null ? '#VALUE!' : R - E) : T - E)));
      var M = g(S.CAPA, 'M' + r), L = serial(g(S.CAPA, 'L' + r));
      set(S.CAPA, 'T' + r, (M === 'Closed' || M === 'Cancelled' || g(S.CAPA, 'L' + r) === '') ? '' : ((L !== null && L < T) ? 'Overdue' : 'On Track'));
    });

    // ---- 18 ISO Docs & Records
    var DOC_ROWS = range(5, 24);
    DOC_ROWS.forEach(function (r) {
      var A = g(S.DOCS, 'A' + r), F = g(S.DOCS, 'F' + r), M = serial(g(S.DOCS, 'M' + r));
      var s;
      if (A === '') s = '';
      else if (F === '' || F === 'Not Started' || F === 'Missing') s = 'Not Evidenced';
      else if (F === 'Not Applicable') s = 'N/A';
      else if (M === null) s = 'Review Date Missing';
      else if (M < T) s = 'Overdue';
      else if (M <= T + 30) s = 'Due ≤30 Days';
      else s = 'Current';
      set(S.DOCS, 'S' + r, s);
    });
    function docRow(ext) {
      for (var i = 0; i < DOC_ROWS.length; i++) if (g(S.DOCS, 'A' + DOC_ROWS[i]) === ext) return DOC_ROWS[i];
      return null;
    }
    function docComplete(r) {
      return g(S.DOCS, 'F' + r) === 'Evidence Filed' && g(S.DOCS, 'K' + r) !== '' && serial(g(S.DOCS, 'L' + r)) !== null &&
        serial(g(S.DOCS, 'M' + r)) !== null && g(S.DOCS, 'N' + r) !== '' && g(S.DOCS, 'O' + r) !== '' && g(S.DOCS, 'P' + r) !== '' &&
        serial(g(S.DOCS, 'Q' + r)) !== null && g(S.DOCS, 'R' + r) !== '' && g(S.DOCS, 'S' + r) !== 'Overdue';
    }
    function pdcaDocStatus(ext) {
      var r = docRow(ext);
      var s = r ? g(S.DOCS, 'F' + r) : '';
      if (s === '' || s === 'Not Started') return 'Not Assessed';
      if (s === 'Not Applicable') return 'Yes';
      return docComplete(r) ? 'Yes' : 'Partial';
    }
    function isoDocStatus(ext) {
      var r = docRow(ext);
      if (!r) return '#N/A';
      var s = g(S.DOCS, 'F' + r);
      if (s === '' || s === 'Not Started') return 'Not Assessed';
      if (s === 'Not Applicable') return 'N/A';
      if (s === 'Missing') return 'Missing';
      return docComplete(r) ? 'In Place' : 'Partial';
    }

    // ---- 17 PDCA
    var biaIds = [];
    for (k = 1; k <= BIA.blocks; k++) { var bb = g(S.BIA, 'B' + BIA.top(k)); if (bb !== '') biaIds.push({ id: bb, t: BIA.top(k) }); }
    var raIds = RA_ROWS.filter(function (r) { return /^RA-/.test(g(S.RA, 'B' + r)); });
    var ssIds = SS_ROWS.filter(function (r) { return /^SS-/.test(g(S.SS, 'C' + r)); });
    var rrIds = RR_ROWS.filter(function (r) { return /^RR-/.test(g(S.RR, 'B' + r)); });
    var drIds = range(5, 204).filter(function (r) { return /^DR-/.test(g(S.DR, 'B' + r)); });
    var capaIds = CAPA_ROWS.filter(function (r) { return /^CAPA-/.test(g(S.CAPA, 'B' + r)); });
    var reqCount = count(REQ_ROWS, function (r) { return /^REQ-/.test(g(S.REQ, 'B' + r)); });
    var pd = {};
    (function () {
      var scope = count(PS_ROWS, function (r) { return g(S.PS, 'I' + r) === 'Yes'; });
      var fields = count(range(5, 14), function (r) { return g(S.BUP, 'B' + r) !== ''; });
      pd[5] = (scope === 0 && fields === 0) ? 'Not Assessed' : ((scope > 0 && fields === 10) ? 'Yes' : 'Partial');
      pd[6] = pdcaDocStatus('EXT-01');
      var ps = count(PS_ROWS, function (r) { return /^PS-/.test(g(S.PS, 'B' + r)); });
      var crit = count(PS_ROWS, function (r) { return g(S.PS, 'J' + r) === 'Yes'; });
      pd[7] = ps === 0 ? 'Not Assessed' : (crit === 0 ? 'Partial' : 'Yes');
      var n = biaIds.length;
      var complete = count(biaIds, function (b) { return g(S.BIA, 'E' + b.t) !== '' && serial(g(S.BIA, 'AF' + b.t)) !== null && serial(g(S.BIA, 'AG' + b.t)) !== null; });
      var overdue = count(biaIds, function (b) { var x = serial(g(S.BIA, 'AG' + b.t)); return x !== null && x < T; });
      pd[8] = n === 0 ? 'Not Assessed' : ((complete === n && overdue === 0) ? 'Yes' : 'Partial');
      var covered = count(biaIds, function (b) { return DEP_ROWS.some(function (r) { return eqStr(g(S.DEP, 'C' + r), b.id); }); });
      pd[9] = n === 0 ? 'Not Assessed' : (covered === n ? 'Yes' : 'Partial');
      var done = count(RA_ROWS, function (r) { return g(S.RA, 'B' + r) !== '' && raComplete(r); });
      pd[10] = raIds.length === 0 ? 'Not Assessed' : (done === raIds.length ? 'Yes' : 'Partial');
      var ssComplete = count(SS_ROWS, function (r) {
        return g(S.SS, 'C' + r) !== '' && g(S.SS, 'B' + r) !== '' && g(S.SS, 'F' + r) !== '' && g(S.SS, 'H' + r) !== '' && g(S.SS, 'I' + r) !== '' &&
          g(S.SS, 'P' + r) !== '' && serial(g(S.SS, 'Q' + r)) !== null && serial(g(S.SS, 'R' + r)) !== null && g(S.SS, 'T' + r) !== '';
      });
      pd[11] = ssIds.length === 0 ? 'Not Assessed' : (ssComplete === ssIds.length ? 'Yes' : 'Partial');
      pd[12] = ssIds.length === 0 ? 'Not Assessed' : (reqCount === 0 ? 'Partial' : (reqCount >= ssIds.length ? 'Yes' : 'Partial'));
      var rrComplete = count(RR_ROWS, function (r) { return ['B', 'C', 'D', 'E', 'F'].every(function (c) { return g(S.RR, c + r) !== ''; }); });
      var rrAmp = rrIds.length;
      pd[13] = rrAmp === 0 ? 'Not Assessed' : (rrComplete === rrAmp ? 'Yes' : 'Partial');
      var drComplete = count(range(5, 204), function (r) {
        return g(S.DR, 'B' + r) !== '' && g(S.DR, 'E' + r) !== '' && g(S.DR, 'F' + r) !== '' &&
          ['G', 'H', 'I', 'J'].every(function (c) { return isNum(g(S.DR, c + r)); }) &&
          g(S.DR, 'K' + r) !== '' && g(S.DR, 'L' + r) !== '' && serial(g(S.DR, 'M' + r)) !== null && g(S.DR, 'N' + r) !== '' &&
          g(S.DR, 'O' + r) !== '' && serial(g(S.DR, 'P' + r)) !== null;
      });
      var drCurrent = count(range(5, 204), function (r) { return g(S.DR, 'Q' + r) === 'Tested & Current'; });
      pd[14] = drIds.length === 0 ? 'Not Assessed' : ((drComplete === drIds.length && drCurrent === drIds.length) ? 'Yes' : 'Partial');
      var exComplete = count(RR_ROWS, function (r) {
        return g(S.RR, 'B' + r) !== '' && g(S.RR, 'F' + r) !== '' && serial(g(S.RR, 'L' + r)) !== null && serial(g(S.RR, 'M' + r)) !== null && g(S.RR, 'N' + r) !== '';
      });
      var exOverdue = count(RR_ROWS, function (r) { return g(S.RR, 'P' + r) === 'Overdue'; });
      pd[15] = rrAmp === 0 ? 'Not Assessed' : ((exComplete === rrAmp && exOverdue === 0) ? 'Yes' : 'Partial');
      pd[16] = pdcaDocStatus('EXT-15');
      pd[17] = pdcaDocStatus('EXT-16');
      pd[18] = pdcaDocStatus('EXT-17');
      var cancelled = count(capaIds, function (r) { return g(S.CAPA, 'M' + r) === 'Cancelled'; });
      var closedComplete = count(CAPA_ROWS, function (r) {
        return g(S.CAPA, 'B' + r) !== '' && g(S.CAPA, 'M' + r) === 'Closed' && serial(g(S.CAPA, 'E' + r)) !== null && g(S.CAPA, 'F' + r) !== '' &&
          g(S.CAPA, 'K' + r) !== '' && serial(g(S.CAPA, 'L' + r)) !== null && g(S.CAPA, 'N' + r) !== '' && g(S.CAPA, 'O' + r) !== '' &&
          serial(g(S.CAPA, 'P' + r)) !== null && g(S.CAPA, 'Q' + r) !== '' && serial(g(S.CAPA, 'R' + r)) !== null;
      });
      var capaOver = count(CAPA_ROWS, function (r) { return g(S.CAPA, 'T' + r) === 'Overdue'; });
      var ineff = count(CAPA_ROWS, function (r) { return g(S.CAPA, 'Q' + r) === 'Ineffective'; });
      pd[19] = capaIds.length === 0 ? 'Not Assessed' : ((capaOver > 0 || ineff > 0) ? 'No' : (closedComplete + cancelled === capaIds.length ? 'Yes' : 'Partial'));
    })();
    range(5, 19).forEach(function (r) {
      set(S.PDCA, 'D' + r, pd[r]);
      var E = g(S.PDCA, 'E' + r);
      set(S.PDCA, 'F' + r, E !== '' ? E : (pd[r] === '' ? 'Not Assessed' : pd[r]));
    });

    // ---- 19 ISO 22301 map
    var wiring = (opts.tpl && opts.tpl.isoMap) || { D: {}, Iformula: [] };
    function fromPdca(row) {
      var s = g(S.PDCA, 'F' + row);
      if (/^#/.test(s)) return s;
      return { Yes: 'In Place', Partial: 'Partial', No: 'Missing', 'Not Assessed': 'Not Assessed' }[s] || 'Not Assessed';
    }
    var roles = 0;
    range(18, 25).forEach(function (r) { ['B', 'C', 'D'].forEach(function (c) { if (g(S.BUP, c + r) !== '') roles++; }); });
    Object.keys(wiring.D).forEach(function (rs) {
      var w = wiring.D[rs], r = +rs, s;
      if (w.type === 'pdca') s = fromPdca(w.row);
      else if (w.type === 'ext') s = isoDocStatus(w.ext);
      else if (w.type === 'pdcaext') {
        var p = g(S.PDCA, 'F' + w.row), e = isoDocStatus(w.ext);
        if (/^#/.test(p)) s = p;
        else if (p === 'Not Assessed' && e === 'Not Assessed') s = 'Not Assessed';
        else if (p === 'No' || e === 'Missing') s = 'Missing';
        else if (p === 'Yes' && (e === 'In Place' || e === 'N/A')) s = 'In Place';
        else s = 'Partial';
      } else if (w.type === 'roles51') s = roles === 0 ? 'Not Assessed' : (roles < 24 ? 'Partial' : 'In Place');
      else if (w.type === 'roles53') {
        var owned = count(SS_ROWS, function (q) { return g(S.SS, 'C' + q) !== '' && g(S.SS, 'P' + q) !== ''; });
        s = (roles === 0 && ssIds.length === 0) ? 'Not Assessed' : ((roles === 24 && owned === ssIds.length) ? 'In Place' : 'Partial');
      } else if (w.type === 'docs75') {
        // 7.5: every listed document in 18. ISO Docs & Records (not N/A) filed and in date.
        var listed = 0, started = 0, filed = 0;
        DOC_ROWS.forEach(function (q) {
          if (g(S.DOCS, 'A' + q) === '') return;
          var st = g(S.DOCS, 'F' + q);
          if (st !== 'Not Applicable') listed++;
          if (st !== '' && st !== 'Not Started' && st !== 'Not Applicable') started++;
          if (st === 'Evidence Filed' && g(S.DOCS, 'S' + q) !== 'Overdue') filed++;
        });
        s = started === 0 ? 'Not Assessed' : (filed === listed ? 'In Place' : 'Partial');
      }
      set(S.ISO, 'D' + r, s);
    });
    (wiring.Iformula || []).forEach(function (r) {
      var b = str(g(S.ISO, 'B' + r));
      set(S.ISO, 'I' + r, 'Complete when ' + b.charAt(0).toLowerCase() + b.slice(1) + '; evidence is approved, current and retrievable.');
    });

    // ---- 02 Dashboard
    var D = {};
    D.psInScope = count(PS_ROWS, function (r) { return g(S.PS, 'I' + r) === 'Yes'; });
    D.biaCompleted = count(biaIds, function (b) { return serial(g(S.BIA, 'AF' + b.t)) !== null; });
    D.criticalPs = count(biaIds, function (b) { return g(S.BIA, 'W' + b.t) === 'Yes'; });
    D.psRiskAssoc = 0;
    RA_ROWS.forEach(function (r) { var f = g(S.RA, 'F' + r); if (f !== '') D.psRiskAssoc += f.split(',').length; });
    D.dependencies = count(DEP_ROWS, function (r) { return /^DEP-/.test(g(S.DEP, 'B' + r)); });
    D.biaOverdue = count(biaIds, function (b) { var x = serial(g(S.BIA, 'AG' + b.t)); return x !== null && x < T; });
    D.risks = raIds.length;
    D.critSig = count(RA_ROWS, function (r) { var p = g(S.RA, 'P' + r); return p === 'Critical' || p === 'Significant'; });
    D.prio12 = count(RA_ROWS, function (r) { var q = g(S.RA, 'Q' + r); return q === 1 || q === 2; });
    D.ssDefined = ssIds.length;
    D.ssOverdue = count(SS_ROWS, function (r) { return g(S.SS, 'T' + r) === 'Overdue'; });
    D.reqRows = reqCount;
    D.bcpRows = count(RR_ROWS, function (r) { return /^SS-/.test(g(S.RR, 'C' + r)); });
    D.exercisesDone = count(RR_ROWS, function (r) { return serial(g(S.RR, 'M' + r)) !== null; });
    D.exercisesOverdue = count(RR_ROWS, function (r) { return g(S.RR, 'P' + r) === 'Overdue'; });
    D.drSystems = drIds.length;
    D.drCurrent = count(range(5, 204), function (r) { return g(S.DR, 'Q' + r) === 'Tested & Current'; });
    D.extNotEvidenced = count(DOC_ROWS, function (r) {
      var a = String(g(S.DOCS, 'A' + r)), f = g(S.DOCS, 'F' + r);
      return a.slice(0, 4) === 'EXT-' && ((f !== 'Evidence Filed' && f !== 'Not Applicable') || g(S.DOCS, 'S' + r) === 'Overdue');
    });
    D.isoMissing = count(range(5, 43), function (r) { return g(S.ISO, 'D' + r) === 'Missing'; });
    D.ssDuplicates = count(SS_ROWS, function (r) { return g(S.SS, 'U' + r) === 'DUPLICATE'; });
    function capaOpen(r) { var m = g(S.CAPA, 'M' + r); return m !== 'Closed' && m !== 'Cancelled'; }
    D.capaOpen = count(capaIds, capaOpen);
    D.capaOverdue = count(CAPA_ROWS, function (r) { return g(S.CAPA, 'T' + r) === 'Overdue'; });
    // Excel COUNTIFS over the whole table (blank rows have Status <> Closed).
    D.capaPending = count(CAPA_ROWS, function (r) { return g(S.CAPA, 'M' + r) === 'Pending Verification'; }) +
      count(CAPA_ROWS, function (r) { return g(S.CAPA, 'M' + r) === 'Closed' && g(S.CAPA, 'Q' + r) === 'Pending'; });
    D.capaCritHigh = count(CAPA_ROWS, function (r) { var sv = g(S.CAPA, 'G' + r); return capaOpen(r) && (sv === 'Critical' || sv === 'High'); });
    calc.dashboard = D;
    // The same counters at their cells in 02. Dashboard.
    [['B5', D.psInScope], ['B6', D.biaCompleted], ['B7', D.criticalPs], ['B8', D.psRiskAssoc], ['B9', D.dependencies], ['B10', D.biaOverdue],
      ['E5', D.risks], ['E6', D.critSig], ['E7', D.prio12], ['E8', D.ssDefined + ' / ' + D.ssOverdue], ['E9', D.reqRows],
      ['H5', D.bcpRows], ['H6', D.exercisesDone + ' / ' + D.exercisesOverdue], ['H7', D.drSystems + ' / ' + D.drCurrent],
      ['H8', D.extNotEvidenced], ['H9', D.isoMissing], ['H10', D.ssDuplicates],
      ['B30', D.capaOpen], ['E30', D.capaOverdue], ['H30', D.capaPending], ['B31', D.capaCritHigh]
    ].forEach(function (x) { set(S.DASH, x[0], x[1]); });
    calc.register = register;

    // ---- 21 Analysis
    var analysis = [];
    PS_ROWS.filter(function (r) { return g(S.PS, 'I' + r) === 'Yes'; }).forEach(function (r) {
      var id = g(S.PS, 'B' + r);
      var bt = biaTopByPs(id);
      analysis.push({
        B: id, C: str(g(S.PS, 'C' + r)), D: str(g(S.PS, 'D' + r)), E: str(g(S.PS, 'E' + r)),
        F: str(g(S.PS, 'K' + r)), G: str(g(S.PS, 'L' + r)),
        // BIA MTPD, RTO and RPO (hours) of the P&S's BIA block.
        I: bt ? g(S.BIA, 'AA' + bt) : '',
        J: bt ? g(S.BIA, 'AD' + bt) : '',
        K: bt ? g(S.BIA, 'AE' + bt) : ''
      });
    });
    analysis.forEach(function (row, i) { ['B', 'C', 'D', 'E', 'F', 'G', 'I', 'J', 'K'].forEach(function (c) { set(S.ANA, c + (2 + i), row[c]); }); });
    calc.analysis = analysis;

    // Lists used by dropdowns.
    calc.lists = {
      psIds: PS_ROWS.map(function (r) { return g(S.PS, 'B' + r); }).filter(Boolean),
      biaIds: biaIds.map(function (b) { return b.id; }),
      riskIds: raIds.map(function (r) { return g(S.RA, 'B' + r); }),
      ssIds: ssIds.map(function (r) { return g(S.SS, 'C' + r); }),
      rrIds: rrIds.map(function (r) { return g(S.RR, 'B' + r); }),
      drIds: drIds.map(function (r) { return g(S.DR, 'B' + r); }),
      depIds: DEP_ROWS.map(function (r) { return g(S.DEP, 'B' + r); }).filter(Boolean),
      docIds: DOC_ROWS.map(function (r) { return g(S.DOCS, 'A' + r); }).filter(Boolean),
      businessUnits: uniq(PS_ROWS.map(function (r) { return str(g(S.PS, 'C' + r)); }).filter(Boolean)).sort(),
      divisionSites: uniq(PS_ROWS.map(function (r) { return str(g(S.PS, 'D' + r)); }).filter(Boolean)).sort()
    };
    return calc;
  }

  var ENGINE = {
    compute: compute, serial: serial, isoFromSerial: isoFromSerial, todaySerial: todaySerial,
    edate: edate, durationHours: durationHours, depKey: depKey
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = ENGINE;
  else root.BCM_ENGINE = ENGINE;
})(this);
