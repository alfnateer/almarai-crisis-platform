/*
 * Data-quality checks and CAPA suggestions.
 *
 * missing(): required entries left empty in a filled row - the same rules as the red
 * conditional formatting added to the template (tools/fix_template.py), plus S&S
 * requirement flags set to Yes without matching rows in 12. Requirements Register.
 * findings(): issues that warrant a corrective action, offered as CAPA rows.
 */
(function (root) {
  'use strict';

  var SCHEMA = root.BCM_SCHEMA || (typeof require !== 'undefined' ? require('./schema.js') : null);
  var ENGINE = root.BCM_ENGINE || (typeof require !== 'undefined' ? require('./engine.js') : null);
  var S = SCHEMA.S, BIA = SCHEMA.BIA;

  function range(a, b) { var o = []; for (var i = a; i <= b; i++) o.push(i); return o; }

  // sheet, rows, key column (row is in use when it is filled), required columns
  var REQUIRED = [
    { id: 'ps', sheet: S.PS, rows: range(5, 49), key: 'E', cols: ['C', 'D', 'G', 'H', 'I'] },
    { id: 'proc', sheet: S.PROC, rows: range(5, 211), key: 'B', cols: ['D', 'E', 'H', 'I'] },
    { id: 'bia', sheet: S.BIA, rows: range(1, BIA.blocks).map(BIA.top), key: 'E', cols: ['V', 'AC', 'AF'], keyAuto: true },
    { id: 'dep', sheet: S.DEP, rows: range(5, 209), key: 'I', cols: ['C', 'F', 'K', 'L'] },
    { id: 'ra', sheet: S.RA, rows: range(5, 61), key: 'G', cols: ['C', 'D', 'E', 'I', 'J', 'K', 'M'] },
    { id: 'ss', sheet: S.SS, rows: range(5, 64), key: 'B', cols: ['H', 'I', 'P', 'Q', 'R'] },
    { id: 'req', sheet: S.REQ, rows: range(5, 304), key: 'F', cols: ['C', 'E', 'K', 'L'] },
    { id: 'dr', sheet: S.DR, rows: range(5, 204), key: 'E', cols: ['C', 'D', 'G', 'H', 'I', 'J', 'K', 'M'] },
    { id: 'rr', sheet: S.RR, rows: range(5, 104), key: 'C', cols: ['E', 'F', 'L'] },
    { id: 'capa', sheet: S.CAPA, rows: range(6, 105), key: 'F', cols: ['C', 'D', 'E', 'G', 'K', 'L', 'M'] }
  ];
  var FLAG_COLS = ['J', 'K', 'L', 'M', 'N', 'O'];

  function label(sheet, col, row) {
    var def = SCHEMA.sheets.filter(function (d) { return d.sheet === sheet; })[0];
    var c = def && def.cols ? def.cols.filter(function (x) { return x.c === col; })[0] : null;
    if (c) return c.label;
    return { V: 'BU Critical (Y/N)', AC: 'RTO', AF: 'Assessment Date' }[col] || col + row;
  }

  // val(sheet, addr): the value shown on the site (inputs, or calculated values).
  function missing(val) {
    var out = [];
    REQUIRED.forEach(function (r) {
      r.rows.forEach(function (row) {
        var key = val(r.sheet, r.key + row);
        if (key === '' || key === undefined) return;
        r.cols.forEach(function (c) {
          if (val(r.sheet, c + row) === '') {
            out.push({ id: r.id, sheet: r.sheet, addr: c + row, row: row, level: 'red', what: label(r.sheet, c, row) + ' is empty',
              ref: refFor(val, r, row) });
          }
        });
      });
    });
    // BIA: RTO must be shorter than MTPD.
    range(1, BIA.blocks).forEach(function (k) {
      var t = BIA.top(k), rto = val(S.BIA, 'AD' + t), mtpd = val(S.BIA, 'AA' + t);
      if (typeof rto === 'number' && typeof mtpd === 'number' && rto >= mtpd) {
        out.push({ id: 'bia', sheet: S.BIA, addr: 'AC' + t, row: t, level: 'red', ref: val(S.BIA, 'B' + t),
          what: 'RTO (' + rto + ' h) is not shorter than MTPD (' + mtpd + ' h)' });
      }
    });
    // S&S requirement flags: Yes needs rows of that category in 12. Requirements Register.
    range(5, 64).forEach(function (row) {
      var ssId = val(S.SS, 'C' + row);
      if (ssId === '') return;
      FLAG_COLS.forEach(function (c) {
        if (val(S.SS, c + row) !== 'Yes') return;
        var cat = val(S.SS, c + '4');
        var total = 0, ready = 0;
        range(5, 304).forEach(function (q) {
          if (String(val(S.REQ, 'C' + q)).toLowerCase() === String(ssId).toLowerCase() &&
              String(val(S.REQ, 'E' + q)).toLowerCase() === String(cat).toLowerCase()) {
            total++;
            if (val(S.REQ, 'L' + q) === 'Ready') ready++;
          }
        });
        if (total === 0) {
          out.push({ id: 'ss', sheet: S.SS, addr: c + row, row: row, level: 'red', ref: ssId, flag: true,
            what: cat + ' = Yes but no "' + cat + '" requirement for ' + ssId + ' in 12. Requirements Register' });
        } else if (ready < total) {
          out.push({ id: 'ss', sheet: S.SS, addr: c + row, row: row, level: 'amber', ref: ssId, flag: true,
            what: cat + ' requirements for ' + ssId + ': ' + ready + ' of ' + total + ' Ready' });
        }
      });
    });
    return out;
  }
  function refFor(val, r, row) {
    var idCol = { ps: 'B', proc: 'B', bia: 'B', dep: 'B', ra: 'B', ss: 'C', req: 'B', dr: 'B', rr: 'B', capa: 'B' }[r.id];
    var id = val(r.sheet, idCol + row);
    return id || ('row ' + row);
  }

  var SEVERITY_BY_PRIORITY = { 1: 'Critical', 2: 'High', 3: 'Medium', 4: 'Low' };
  var TARGET_DAYS = { Critical: 14, High: 30, Medium: 60, Low: 90 };

  // Suggested CAPA rows for findings in the registers.
  function findings(val, todayIso) {
    var out = [];
    function riskOf(riskId) {
      for (var r = 5; r <= 61; r++) if (val(S.RA, 'B' + r) === riskId) return r;
      return null;
    }
    function sevForRisk(riskId, fallback) {
      var r = riskOf(riskId);
      var p = r ? val(S.RA, 'Q' + r) : '';
      return SEVERITY_BY_PRIORITY[p] || fallback || 'Medium';
    }
    function add(f) {
      f.target = ENGINE.isoFromSerial(ENGINE.serial(todayIso) + (TARGET_DAYS[f.severity] || 60));
      f.key = f.ref + '|' + f.finding;
      out.push(f);
    }
    // Strategies: requirement flags not captured, overdue strategies.
    range(5, 64).forEach(function (row) {
      var ssId = val(S.SS, 'C' + row), riskId = val(S.SS, 'B' + row);
      if (ssId === '') return;
      FLAG_COLS.forEach(function (c) {
        if (val(S.SS, c + row) !== 'Yes') return;
        var cat = val(S.SS, c + '4'), total = 0;
        range(5, 304).forEach(function (q) {
          if (String(val(S.REQ, 'C' + q)).toLowerCase() === String(ssId).toLowerCase() && String(val(S.REQ, 'E' + q)).toLowerCase() === String(cat).toLowerCase()) total++;
        });
        if (!total) add({ type: 'Risk', ref: ssId, riskId: riskId, ssId: ssId, severity: sevForRisk(riskId), owner: val(S.SS, 'P' + row),
          finding: ssId + ': ' + cat + ' requirement flagged Yes but not captured in 12. Requirements Register' });
      });
      if (val(S.SS, 'T' + row) === 'Overdue') add({ type: 'Risk', ref: ssId, riskId: riskId, ssId: ssId, severity: sevForRisk(riskId), owner: val(S.SS, 'P' + row),
        finding: ssId + ': strategy implementation overdue (due ' + val(S.SS, 'R' + row) + ')' });
    });
    // Critical / significant risks without a strategy.
    range(5, 61).forEach(function (row) {
      var id = val(S.RA, 'B' + row), p = val(S.RA, 'P' + row);
      if (id !== '' && (p === 'Critical' || p === 'Significant') && val(S.RA, 'S' + row) === 'Not Created') {
        add({ type: 'Risk', ref: id, riskId: id, severity: p === 'Critical' ? 'Critical' : 'High',
          finding: id + ': ' + p + ' risk "' + val(S.RA, 'G' + row) + '" has no continuity strategy (11. Strategies & Solutions)' });
      }
    });
    // Exercises.
    range(5, 104).forEach(function (row) {
      var id = val(S.RR, 'B' + row), st = val(S.RR, 'P' + row);
      if (id === '') return;
      var ssId = val(S.RR, 'C' + row), riskId = '';
      range(5, 64).forEach(function (q) { if (val(S.SS, 'C' + q) === ssId) riskId = val(S.SS, 'B' + q); });
      if (st === 'Failed') add({ type: 'Recovery Plan', ref: id, rrId: id, ssId: ssId, riskId: riskId, severity: 'High', finding: id + ': exercise failed' });
      if (st === 'Overdue') add({ type: 'Recovery Plan', ref: id, rrId: id, ssId: ssId, riskId: riskId, severity: 'Medium', finding: id + ': planned exercise (' + val(S.RR, 'L' + row) + ') not carried out' });
    });
    // IT DR.
    var drSev = { 'GAP - no DR': 'High', 'Retest Required': 'High', 'Recovery Not Compliant': 'High', 'Test Overdue': 'Medium' };
    range(5, 204).forEach(function (row) {
      var id = val(S.DR, 'B' + row), st = val(S.DR, 'Q' + row);
      if (id !== '' && drSev[st]) add({ type: 'IT DR', ref: id, severity: drSev[st], finding: id + ' ' + val(S.DR, 'E' + row) + ': ' + st });
    });
    // BIA reviews.
    range(1, BIA.blocks).forEach(function (k) {
      var t = BIA.top(k), id = val(S.BIA, 'B' + t), until = ENGINE.serial(val(S.BIA, 'AG' + t));
      if (id !== '' && until !== null && until < ENGINE.serial(todayIso)) add({ type: 'BIA', ref: id, severity: 'Medium',
        finding: id + ' ' + val(S.BIA, 'E' + t) + ': BIA review overdue (valid until ' + val(S.BIA, 'AG' + t) + ')' });
    });
    // Single points of failure.
    range(5, 209).forEach(function (row) {
      var id = val(S.DEP, 'B' + row);
      if (id !== '' && val(S.DEP, 'K' + row) === 'Yes') add({ type: 'Dependency', ref: id, severity: 'Medium', owner: val(S.DEP, 'L' + row),
        finding: id + ': single point of failure - ' + val(S.DEP, 'I' + row) + ' (' + val(S.DEP, 'D' + row) + ')' });
    });
    // ISO documents past their review date.
    range(5, 24).forEach(function (row) {
      var ref = val(S.DOCS, 'A' + row);
      if (ref !== '' && val(S.DOCS, 'S' + row) === 'Overdue') add({ type: 'ISO Document', ref: ref, severity: 'Medium', owner: val(S.DOCS, 'E' + row),
        finding: ref + ' ' + val(S.DOCS, 'C' + row) + ': review date passed' });
    });
    return out;
  }

  // Findings not yet in the CAPA register (same Source Ref and Finding).
  function newFindings(val, todayIso) {
    var have = {};
    range(6, 105).forEach(function (r) { have[val(S.CAPA, 'D' + r) + '|' + val(S.CAPA, 'F' + r)] = true; });
    return findings(val, todayIso).filter(function (f) { return !have[f.key]; });
  }

  var CHECKS = { missing: missing, findings: findings, newFindings: newFindings, REQUIRED: REQUIRED };
  if (typeof module !== 'undefined' && module.exports) module.exports = CHECKS;
  else root.BCM_CHECKS = CHECKS;
})(this);
