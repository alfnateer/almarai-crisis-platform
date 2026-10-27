#!/usr/bin/env python3
"""Make an empty export template from a filled BCM Toolkit workbook.

    node tools/dump_inputs.js > /tmp/inputs.json
    python3 tools/make_blank_template.py filled.xlsx /tmp/inputs.json template/<name>.xlsx

Clears every input cell listed in inputs.json (the website's schema) except those
marked keep (default lists such as the Risk Assumption Map), drops the
cached results of all formulas (and the cells their dynamic arrays spilled into),
removes shared strings that are no longer used, removes the calculation chain and
asks Excel to recalculate on open. Formulas, styles, validations, tables, comments
and sheet layout are kept byte-for-byte apart from those edits.
"""
import json
import re
import sys
import zipfile

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
Q = lambda t: "{%s}%s" % (NS, t)


def col_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split(addr):
    m = re.match(r"([A-Z]+)(\d+)$", addr)
    return col_num(m.group(1)), int(m.group(2))


def in_range(addr, ref):
    a, b = (ref.split(":") + [ref])[:2]
    c, r = split(addr)
    c1, r1 = split(a)
    c2, r2 = split(b)
    return c1 <= c <= c2 and r1 <= r <= r2


def clear(c):
    for ch in list(c):
        c.remove(ch)
    for a in ("t", "cm", "vm"):
        c.attrib.pop(a, None)


def main(src, inputs_path, dst):
    inputs = json.load(open(inputs_path))
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}

    wb = etree.fromstring(files["xl/workbook.xml"])
    rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
    target = {r.get("Id"): r.get("Target") for r in rels.iter("{%s}Relationship" % NS_PR)}
    sheet_path = {}
    for s in wb.iter(Q("sheet")):
        t = target[s.get("{%s}id" % NS_R)]
        sheet_path[s.get("name")] = t.lstrip("/") if t.startswith("/") else "xl/" + t

    by_sheet = {}
    for c in inputs:
        if not c.get("keep"):  # sheets of defaults (e.g. the risk assumption map) stay filled
            by_sheet.setdefault(c["sheet"], set()).add(c["addr"])

    stats = {}
    for name, path in sheet_path.items():
        doc = etree.fromstring(files[path])
        wanted = by_sheet.get(name, set())
        cells = list(doc.iter(Q("c")))
        spill = [f.getparent() for f in doc.iter(Q("f")) if f.get("t") == "array" and f.get("ref")]
        spill_refs = [(c.get("r"), c.find(Q("f")).get("ref")) for c in spill]
        n_in = n_f = n_sp = 0
        for c in cells:
            addr = c.get("r")
            f = c.find(Q("f"))
            if addr in wanted:
                clear(c)
                n_in += 1
            elif f is not None:
                v = c.find(Q("v"))
                if v is not None:
                    c.remove(v)
                c.attrib.pop("t", None)
                n_f += 1
            elif c.find(Q("v")) is not None or c.find(Q("is")) is not None:
                if any(addr != a and in_range(addr, ref) for a, ref in spill_refs):
                    clear(c)
                    n_sp += 1
        # Hyperlinks on cleared cells (e.g. mailto: contacts) go too, with their relationships.
        dropped_rids = set()
        for hl in list(doc.iter(Q("hyperlink"))):
            if hl.get("ref") in wanted:
                rid = hl.get("{%s}id" % NS_R)
                if rid:
                    dropped_rids.add(rid)
                parent = hl.getparent()
                parent.remove(hl)
                if len(parent) == 0:
                    parent.getparent().remove(parent)
        rels_path = path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if dropped_rids and rels_path in files:
            srels = etree.fromstring(files[rels_path])
            for r in list(srels):
                if r.get("Id") in dropped_rids:
                    srels.remove(r)
            files[rels_path] = etree.tostring(srels, xml_declaration=True, encoding="UTF-8", standalone=True)
        stats[name] = (n_in, n_f, n_sp)
        files[path] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Rebuild shared strings without the ones no sheet uses any more.
    if "xl/sharedStrings.xml" in files:
        sst = etree.fromstring(files["xl/sharedStrings.xml"])
        items = list(sst.iter(Q("si")))
        used, sheet_docs = set(), {}
        for path in sheet_path.values():
            d = etree.fromstring(files[path])
            sheet_docs[path] = d
            for c in d.iter(Q("c")):
                if c.get("t") == "s":
                    used.add(int(c.find(Q("v")).text))
        remap = {old: new for new, old in enumerate(sorted(used))}
        for si in items:
            sst.remove(si)
        for old in sorted(used):
            sst.append(items[old])
        sst.set("count", str(len(used)))
        sst.set("uniqueCount", str(len(used)))
        files["xl/sharedStrings.xml"] = etree.tostring(sst, xml_declaration=True, encoding="UTF-8", standalone=True)
        for path, d in sheet_docs.items():
            for c in d.iter(Q("c")):
                if c.get("t") == "s":
                    v = c.find(Q("v"))
                    v.text = str(remap[int(v.text)])
            files[path] = etree.tostring(d, xml_declaration=True, encoding="UTF-8", standalone=True)
        print("shared strings: %d -> %d" % (len(items), len(used)))

    # Recalculate on open; the calculation chain is rebuilt by Excel.
    calc = wb.find(Q("calcPr"))
    if calc is None:
        calc = etree.SubElement(wb, Q("calcPr"))
    calc.set("fullCalcOnLoad", "1")
    files["xl/workbook.xml"] = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)
    # Drop the saved-from path (OneDrive / desktop folder of whoever saved it last).
    files["xl/workbook.xml"] = re.sub(rb"<mc:AlternateContent[^>]*>\s*<mc:Choice Requires=\"x15\">\s*<x15ac:absPath[^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>", b"", files["xl/workbook.xml"])
    if "xl/calcChain.xml" in files:
        del files["xl/calcChain.xml"]
        files["[Content_Types].xml"] = re.sub(rb'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b"", files["[Content_Types].xml"])
        files["xl/_rels/workbook.xml.rels"] = re.sub(rb'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', b"", files["xl/_rels/workbook.xml.rels"])

    if "docProps/core.xml" in files:
        files["docProps/core.xml"] = re.sub(rb"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", b"<cp:lastModifiedBy></cp:lastModifiedBy>", files["docProps/core.xml"])

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in files:
                zout.writestr(info.filename, files[info.filename])
    for name, (a, b, c) in stats.items():
        print("%-32s inputs cleared %5d  formula caches dropped %5d  spill cells cleared %4d" % (name, a, b, c))


if __name__ == "__main__":
    main(*sys.argv[1:4])
