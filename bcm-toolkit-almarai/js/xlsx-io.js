/*
 * Excel import / export for the BCM Toolkit.
 *
 * Export opens the original toolkit workbook (embedded in template.js) and writes
 * each input value into its own cell, leaving every formula, style, dropdown,
 * table and comment untouched. Excel is told to fully recalculate on open, so all
 * automatic columns (IDs, ratings, statuses, dashboard) are refreshed from the data.
 *
 * Import reads the same input cells back from any copy of the toolkit workbook.
 */
(function (root) {
  'use strict';

  var NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var EXCEL_EPOCH_OFFSET = 25569; // days between 1899-12-30 and 1970-01-01

  function colToNum(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }
  function splitAddr(addr) {
    var m = /^([A-Z]+)(\d+)$/.exec(addr);
    return { col: m[1], row: +m[2], c: colToNum(m[1]) };
  }
  function parse(xml) { return new DOMParser().parseFromString(xml, 'application/xml'); }
  function serialize(doc) {
    var xml = new XMLSerializer().serializeToString(doc);
    return /^<\?xml/.test(xml) ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
  }
  function children(el, name) {
    var out = [];
    for (var n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && n.localName === name) out.push(n);
    return out;
  }
  function isoToExcel(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    if (!m) return null;
    return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5 + EXCEL_EPOCH_OFFSET;
  }
  function excelToIso(n) {
    var d = new Date(Math.round((n - EXCEL_EPOCH_OFFSET) * 864e5));
    var p = function (x) { return x < 10 ? '0' + x : String(x); };
    return d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate());
  }

  // Map sheet name -> zip path using workbook.xml + its relationships.
  async function sheetPaths(zip) {
    var wb = parse(await zip.file('xl/workbook.xml').async('string'));
    var rels = parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    var targets = {};
    Array.prototype.forEach.call(rels.getElementsByTagNameNS(NS_PKG_REL, 'Relationship'), function (r) {
      targets[r.getAttribute('Id')] = r.getAttribute('Target');
    });
    var out = {};
    Array.prototype.forEach.call(wb.getElementsByTagNameNS(NS, 'sheet'), function (s) {
      var t = targets[s.getAttributeNS(NS_R, 'id')];
      if (!t) return;
      out[s.getAttribute('name')] = t.charAt(0) === '/' ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
    });
    return out;
  }

  // ------------------------------------------------------------------ export
  function findOrCreateRow(doc, sheetData, rowIndex, rowCache) {
    if (rowCache[rowIndex]) return rowCache[rowIndex];
    var row = doc.createElementNS(NS, 'row');
    row.setAttribute('r', String(rowIndex));
    var before = null;
    var rows = children(sheetData, 'row');
    for (var i = 0; i < rows.length; i++) {
      if (+rows[i].getAttribute('r') > rowIndex) { before = rows[i]; break; }
    }
    sheetData.insertBefore(row, before);
    rowCache[rowIndex] = row;
    return row;
  }
  function findOrCreateCell(doc, row, addr, colNum) {
    var cells = children(row, 'c');
    var before = null;
    for (var i = 0; i < cells.length; i++) {
      var r = cells[i].getAttribute('r');
      if (r === addr) return cells[i];
      if (colToNum(splitAddr(r).col) > colNum) { before = cells[i]; break; }
    }
    var c = doc.createElementNS(NS, 'c');
    c.setAttribute('r', addr);
    // Borrow the style of the left neighbour so new cells look like the row.
    var prev = before ? before.previousSibling : row.lastChild;
    while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
    if (prev && prev.getAttribute('s')) c.setAttribute('s', prev.getAttribute('s'));
    row.insertBefore(c, before);
    row.removeAttribute('spans');
    return c;
  }
  function clearCell(c) {
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    c.removeAttribute('cm');
    c.removeAttribute('vm');
  }
  function writeCell(doc, c, value, spec) {
    clearCell(c);
    if (value === undefined || value === null || value === '') return;
    var numeric = null;
    if (spec.type === 'date') numeric = isoToExcel(String(value));
    else if (typeof value === 'number' && isFinite(value)) numeric = value;
    else if ((spec.numeric || spec.type === 'number') && String(value).trim() !== '' && !isNaN(Number(value))) numeric = Number(value);
    if (numeric !== null) {
      var v = doc.createElementNS(NS, 'v');
      v.textContent = String(numeric);
      c.appendChild(v);
      return;
    }
    c.setAttribute('t', 'inlineStr');
    var is = doc.createElementNS(NS, 'is');
    var t = doc.createElementNS(NS, 't');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = String(value);
    is.appendChild(t);
    c.appendChild(is);
  }

  /*
   * cells: [{sheet, addr, type, numeric}] (SCHEMA.inputCells), getValue(sheet, addr)
   * Returns a Blob of the populated workbook.
   */
  async function exportWorkbook(templateB64, cells, getValue) {
    var zip = await root.JSZip.loadAsync(templateB64, { base64: true });
    var paths = await sheetPaths(zip);
    var bySheet = {};
    cells.forEach(function (c) { (bySheet[c.sheet] = bySheet[c.sheet] || []).push(c); });

    for (var name in bySheet) {
      var path = paths[name];
      if (!path) throw new Error('Template is missing sheet "' + name + '"');
      var doc = parse(await zip.file(path).async('string'));
      var sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
      var rowCache = {};
      children(sheetData, 'row').forEach(function (r) { rowCache[+r.getAttribute('r')] = r; });
      bySheet[name].forEach(function (spec) {
        var a = splitAddr(spec.addr);
        var value = getValue(spec.sheet, spec.addr);
        var row = rowCache[a.row];
        if (!row && (value === '' || value === undefined || value === null)) return;
        row = findOrCreateRow(doc, sheetData, a.row, rowCache);
        var cell = null;
        var existing = children(row, 'c');
        for (var i = 0; i < existing.length; i++) if (existing[i].getAttribute('r') === spec.addr) { cell = existing[i]; break; }
        if (!cell && (value === '' || value === undefined || value === null)) return;
        if (!cell) cell = findOrCreateCell(doc, row, spec.addr, a.c);
        writeCell(doc, cell, value, spec);
      });
      zip.file(path, serialize(doc), { createFolders: false });
    }

    // Force a full recalculation when Excel opens the file and drop the stale
    // calculation chain (Excel rebuilds it).
    var wbXml = parse(await zip.file('xl/workbook.xml').async('string'));
    var calcPr = wbXml.getElementsByTagNameNS(NS, 'calcPr')[0];
    if (!calcPr) {
      calcPr = wbXml.createElementNS(NS, 'calcPr');
      var wbEl = wbXml.documentElement;
      var after = wbXml.getElementsByTagNameNS(NS, 'definedNames')[0] || wbXml.getElementsByTagNameNS(NS, 'sheets')[0];
      wbEl.insertBefore(calcPr, after ? after.nextSibling : null);
    }
    calcPr.setAttribute('fullCalcOnLoad', '1');
    zip.file('xl/workbook.xml', serialize(wbXml), { createFolders: false });
    if (zip.file('xl/calcChain.xml')) {
      zip.remove('xl/calcChain.xml');
      var ct = await zip.file('[Content_Types].xml').async('string');
      zip.file('[Content_Types].xml', ct.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''), { createFolders: false });
      var rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
      zip.file('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/, ''), { createFolders: false });
    }
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
  }

  // ------------------------------------------------------------------ import
  function textOf(el) {
    // Concatenate <t> nodes (handles rich-text runs) but skip phonetic <rPh>.
    var out = '';
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.localName === 'rPh') continue;
        if (c.localName === 't') out += c.textContent;
        else walk(c);
      }
    })(el);
    return out;
  }

  /*
   * Read the input cells from a toolkit workbook (ArrayBuffer).
   * Returns {values: {sheet: {addr: value}}, missingSheets: [...]}
   */
  // Encrypted workbooks are OLE compound files instead of zip packages.
  function encryptionError(buffer) {
    var b = new Uint8Array(buffer);
    if (!(b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0)) return null;
    var marker = 'DRMEncryptedDataSpace', utf16 = [];
    for (var i = 0; i < marker.length; i++) utf16.push(marker.charCodeAt(i), 0);
    outer: for (var p = 0; p < b.length - utf16.length; p++) {
      for (var j = 0; j < utf16.length; j++) if (b[p + j] !== utf16[j]) continue outer;
      return new Error('This workbook is protected by a Microsoft 365 sensitivity label (encrypted), so only Excel signed in ' +
        'with your account can open it. In Excel, save an unencrypted copy (change the sensitivity label to one without ' +
        'encryption, or File > Info > Protect Workbook > Unrestricted Access) and import that copy.');
    }
    return new Error('This workbook is password-protected or in the old .xls format. In Excel, remove the password ' +
      '(File > Info > Protect Workbook > Encrypt with Password, clear it) or save as .xlsx, then import that copy.');
  }

  async function importWorkbook(buffer, cells) {
    var encrypted = encryptionError(buffer);
    if (encrypted) throw encrypted;
    var zip = await root.JSZip.loadAsync(buffer);
    if (!zip.file('xl/workbook.xml')) throw new Error('This file is not an Excel workbook (.xlsx).');
    var paths = await sheetPaths(zip);
    var shared = [];
    if (zip.file('xl/sharedStrings.xml')) {
      var sst = parse(await zip.file('xl/sharedStrings.xml').async('string'));
      Array.prototype.forEach.call(sst.getElementsByTagNameNS(NS, 'si'), function (si) { shared.push(textOf(si)); });
    }
    var bySheet = {};
    cells.forEach(function (c) { (bySheet[c.sheet] = bySheet[c.sheet] || []).push(c); });
    var values = {}, missing = [];
    for (var name in bySheet) {
      if (!paths[name]) { missing.push(name); continue; }
      var doc = parse(await zip.file(paths[name]).async('string'));
      var raw = {};
      Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS, 'c'), function (c) { raw[c.getAttribute('r')] = c; });
      var out = values[name] = {};
      bySheet[name].forEach(function (spec) {
        var c = raw[spec.addr];
        if (!c) return;
        var t = c.getAttribute('t') || 'n';
        var vEl = children(c, 'v')[0];
        var v = vEl ? vEl.textContent : null;
        var value;
        if (t === 's') value = v === null ? '' : (shared[+v] || '');
        else if (t === 'inlineStr') { var is = children(c, 'is')[0]; value = is ? textOf(is) : ''; }
        else if (t === 'str') value = v === null ? '' : v;
        else if (t === 'b') value = v === '1' ? 'TRUE' : 'FALSE';
        else if (t === 'e') value = '';
        else value = v === null ? '' : Number(v);
        if (value === '' || value === null) return;
        if (spec.type === 'date') {
          if (typeof value === 'number') value = excelToIso(value);
          else if (/^\d{4}-\d{2}-\d{2}/.test(value)) value = value.slice(0, 10);
        } else if (typeof value === 'number' && !(spec.numeric || spec.type === 'number')) {
          value = String(value);
        }
        out[spec.addr] = value;
      });
    }
    return { values: values, missingSheets: missing };
  }

  root.BCM_XLSX = { exportWorkbook: exportWorkbook, importWorkbook: importWorkbook, isoToExcel: isoToExcel, excelToIso: excelToIso };
})(this);
