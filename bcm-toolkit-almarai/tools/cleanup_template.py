#!/usr/bin/env python3
"""Clean up the fixed Almarai template (v0.2.1) into v0.3.

    python3 tools/cleanup_template.py template/source/Almarai_BCM_Toolkit_Template_v0.2.1.xlsx template/Almarai_BCM_Toolkit_Template_v0.3.xlsx

Run on the output of fix_template.py. Every edit asserts that it found what it expects.

Structure:
  00. Document Control   sheet removed. ISO 22301 7.5 (19. ISO Map D23) now reads the
                         controlled documents in 18. ISO Docs & Records; evidence texts that
                         named the sheet point to 18. ISO Docs & Records
  15. BCM Risk Register  columns T-W removed (Assessment Date, Review Date, Approver,
                         Evidence Reference: not captured in 10. Risk Assessment, always
                         blank). Source Completion moves from X to T
Package clean-up:
  sensitivity-label / SharePoint metadata (customXml, docMetadata, custom properties),
  add-in task pane (webextensions), printer settings, cached sheet list in app.xml,
  and defined names nothing refers to
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fix_template import NS, NS_PR, NS_R, Q, Sheet, col_name, col_num, set_formula, split  # noqa: E402

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
DOC = "00. Document Control"
ISO_DOCS = "'18. ISO Docs & Records'!"

# Evidence texts in 17. PDCA / 19. ISO Map that named 00. Document Control.
TEXTS = {
    "00. Document Control; change-control evidence tracked in 18. ISO Docs & Records — EXT-04":
        "Change-control evidence tracked in 18. ISO Docs & Records — EXT-04",
    "00. Document Control / 18. ISO Docs & Records — EXT-04": "18. ISO Docs & Records — EXT-04",
    "00. Document Control": "18. ISO Docs & Records",
    "00. Document Control, 01. Instructions, and controlled external records tracked in 18. ISO Docs & Records":
        "Controlled documents and records tracked in 18. ISO Docs & Records (version, approval, review, repository) and 01. Instructions",
    "17. BCMS Progress (PDCA), 16. CAPA Register, and 00. Document Control version history":
        "17. BCMS Progress (PDCA), 16. CAPA Register and document versions in 18. ISO Docs & Records",
    "00. Document Control; 16. CAPA Register; 17. BCMS Progress (PDCA)":
        "18. ISO Docs & Records; 16. CAPA Register; 17. BCMS Progress (PDCA)",
}


def rng(c):
    return "%s$%s$5:$%s$24" % (ISO_DOCS, c, c)


# 7.5 Control documented information: every listed document (not N/A) filed and in date.
ISO75 = ('IF(SUMPRODUCT(--({A}<>""),--({F}<>""),--({F}<>"Not Started"),--({F}<>"Not Applicable"))=0,"Not Assessed",'
         'IF(SUMPRODUCT(--({A}<>""),--({F}="Evidence Filed"),--({S}<>"Overdue"))='
         'SUMPRODUCT(--({A}<>""),--({F}<>"Not Applicable")),"In Place","Partial"))').format(A=rng("A"), F=rng("F"), S=rng("S"))


def drop_part(files, name):
    """Remove a part with its own rels file and content-type override."""
    assert name in files, name
    del files[name]
    d, b = name.rsplit("/", 1) if "/" in name else ("", name)
    rels = (d + "/" if d else "") + "_rels/" + b + ".rels"
    files.pop(rels, None)
    ct = etree.fromstring(files["[Content_Types].xml"])
    for o in ct.findall("{%s}Override" % NS_CT):
        if o.get("PartName") == "/" + name:
            ct.remove(o)
    files["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)


def drop_rel(files, rels_name, target_pattern):
    r = etree.fromstring(files[rels_name])
    gone = [x for x in r if re.search(target_pattern, x.get("Target"))]
    for x in gone:
        r.remove(x)
    files[rels_name] = etree.tostring(r, xml_declaration=True, encoding="UTF-8", standalone=True)
    return [x.get("Id") for x in gone]


def resolve(base, target):
    parts = (base.rsplit("/", 1)[0] + "/" + target).split("/") if not target.startswith("/") else target.lstrip("/").split("/")
    out = []
    for p in parts:
        if p == "..":
            out.pop()
        elif p:
            out.append(p)
    return "/".join(out)


def remove_sheet(files, wb, name):
    rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
    sheets = wb.find(Q("sheets"))
    el = [s for s in sheets if s.get("name") == name]
    assert len(el) == 1, name
    el = el[0]
    index = list(sheets).index(el)
    rid = el.get("{%s}id" % NS_R)
    target = [r for r in rels if r.get("Id") == rid][0].get("Target")
    part = resolve("xl/x", target)
    # Parts used only by this sheet: comments, VML notes, printer settings.
    srels = "xl/worksheets/_rels/" + part.rsplit("/", 1)[1] + ".rels"
    if srels in files:
        for r in etree.fromstring(files[srels]):
            dep = resolve(part, r.get("Target"))
            if dep in files:
                drop_part(files, dep)
    drop_part(files, part)
    sheets.remove(el)
    drop_rel(files, "xl/_rels/workbook.xml.rels", re.escape(target) + "$")
    # Sheet-scoped names and the active tab follow the sheet order.
    for dn in wb.iter(Q("definedName")):
        if dn.get("localSheetId") is not None:
            k = int(dn.get("localSheetId"))
            assert k != index, dn.get("name")
            if k > index:
                dn.set("localSheetId", str(k - 1))
    for v in wb.iter(Q("workbookView")):
        for a in ("activeTab", "firstSheet"):
            if v.get(a) is not None and int(v.get(a)) >= index:
                v.set(a, str(max(0, int(v.get(a)) - 1)))


def delete_columns(sheet, first, last):
    """Delete columns first..last of a sheet (cells, widths, merges, CF) and shift the rest left."""
    a, b = col_num(first), col_num(last)
    width = b - a + 1

    def shift(col):
        n = col_num(col)
        return col_name(n - width) if n > b else col

    def shift_ref(ref):
        return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", lambda m: m.group(1) + shift(m.group(2)) + m.group(3) + m.group(4), ref)

    for row in sheet.rows.values():
        for c in list(row.findall(Q("c"))):
            n, _ = split(c.get("r"))
            if a <= n <= b:
                row.remove(c)
            else:
                c.set("r", shift_ref(c.get("r")))
                f = c.find(Q("f"))
                if f is not None and f.get("ref"):
                    f.set("ref", shift_ref(f.get("ref")))
        if row.get("spans"):
            lo, hi = row.get("spans").split(":")
            row.set("spans", "%s:%d" % (lo, min(int(hi), col_num(shift(col_name(int(hi)))))))
    cols = sheet.doc.find(Q("cols"))
    for c in list(cols):
        lo, hi = int(c.get("min")), int(c.get("max"))
        assert lo == hi, "column ranges not expected"
        if a <= lo <= b:
            cols.remove(c)
        elif lo > b:
            c.set("min", str(lo - width))
            c.set("max", str(hi - width))
    for m in sheet.doc.iter(Q("mergeCell")):
        m.set("ref", shift_ref(m.get("ref")))
    for cf in sheet.doc.iter(Q("conditionalFormatting")):
        cf.set("sqref", shift_ref(cf.get("sqref")))
    dim = sheet.doc.find(Q("dimension"))
    dim.set("ref", shift_ref(dim.get("ref")))


def main(src, dst):
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    wb = etree.fromstring(files["xl/workbook.xml"])
    rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
    target = {r.get("Id"): r.get("Target") for r in rels.iter("{%s}Relationship" % NS_PR)}
    path = {s.get("name"): resolve("xl/x", target[s.get("{%s}id" % NS_R)]) for s in wb.iter(Q("sheet"))}

    # 00. Document Control --------------------------------------------------------------
    iso = Sheet(etree.fromstring(files[path["19. ISO 22301 Map"]]))
    f = iso.formula("D23")
    assert f is not None and DOC in f.text
    set_formula(iso, "D23", ISO75, keep_attrs=False)
    files[path["19. ISO 22301 Map"]] = etree.tostring(iso.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    sst = etree.fromstring(files["xl/sharedStrings.xml"])
    done = set()
    for si in sst.findall(Q("si")):
        text = "".join(si.itertext())
        if text in TEXTS:
            for ch in list(si):
                si.remove(ch)
            etree.SubElement(si, Q("t")).text = TEXTS[text]
            done.add(text)
    assert done == set(TEXTS), set(TEXTS) - done
    files["xl/sharedStrings.xml"] = etree.tostring(sst, xml_declaration=True, encoding="UTF-8", standalone=True)
    remove_sheet(files, wb, DOC)
    for name, p in path.items():
        if name != DOC and p in files:
            assert ("'%s'!" % DOC).encode() not in files[p], name

    # 15. BCM Risk Register: drop the unlinked columns T-W ---------------------------------
    reg = Sheet(etree.fromstring(files[path["15. BCM Risk Register"]]))
    for col in "TUVW":
        assert reg.formula(col + "6").text == 'IF(_xlfn.ANCHORARRAY(A6)="","","")', col
    assert "Source Completion" in "".join(sst.findall(Q("si"))[int(reg.cell("X5").find(Q("v")).text)].itertext())
    delete_columns(reg, "T", "W")
    for col in "PQ":
        f = reg.formula(col + "6")
        assert f.text.count("_xlfn.ANCHORARRAY(X6)") == 1, col
        f.text = f.text.replace("_xlfn.ANCHORARRAY(X6)", "_xlfn.ANCHORARRAY(T6)")
    assert reg.formula("T6") is not None and "RiskAssessment" in reg.formula("T6").text
    files[path["15. BCM Risk Register"]] = etree.tostring(reg.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    for p in files:
        if p.startswith("xl/worksheets/sheet") and p != path["15. BCM Risk Register"]:
            assert not re.search(rb"BCM Risk Register'!\$?[T-X]", files[p]), p

    # Package clean-up --------------------------------------------------------------------
    for p in [p for p in files if p.startswith("customXml/") and not p.startswith("customXml/_rels/")]:
        drop_part(files, p)
    drop_rel(files, "xl/_rels/workbook.xml.rels", r"customXml/")
    for p in ("docMetadata/LabelInfo.xml", "docProps/custom.xml"):
        drop_part(files, p)
    drop_rel(files, "_rels/.rels", r"docMetadata/|docProps/custom\.xml")
    for p in [p for p in files if p.startswith("xl/webextensions/") and "/_rels/" not in p]:
        drop_part(files, p)
    drop_rel(files, "_rels/.rels", r"webextensions/")
    for p in [p for p in files if p.startswith("xl/printerSettings/")]:
        drop_part(files, p)
    for p in [p for p in files if p.startswith("xl/worksheets/_rels/")]:
        ids = drop_rel(files, p, r"printerSettings/")
        sheet = p.replace("_rels/", "").replace(".rels", "")
        for rid in ids:
            files[sheet] = re.sub(rb'<pageSetup([^>]*?) r:id="%s"' % rid.encode(), rb"<pageSetup\1", files[sheet])
    if not any(p.endswith(".bin") for p in files):
        ct = etree.fromstring(files["[Content_Types].xml"])
        for d in ct.findall("{%s}Default" % NS_CT):
            if d.get("Extension") == "bin" and "printerSettings" in d.get("ContentType"):
                ct.remove(d)
        files["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    app = etree.fromstring(files["docProps/app.xml"])
    for t in list(app):
        if etree.QName(t).localname in ("HeadingPairs", "TitlesOfParts"):
            app.remove(t)
    files["docProps/app.xml"] = etree.tostring(app, xml_declaration=True, encoding="UTF-8", standalone=True)
    # Defined names nothing refers to (formulas, validations, tables, other names).
    text = b" ".join(v for k, v in files.items() if k.startswith(("xl/worksheets/", "xl/tables/")))
    names = wb.find(Q("definedNames"))
    removed = []
    for dn in list(names):
        n = dn.get("name")
        if n.startswith("_xlnm."):
            continue
        others = " ".join(x.text or "" for x in names if x is not dn)
        if not re.search(r"\b%s\b" % re.escape(n), text.decode("utf-8", "ignore") + " " + others):
            names.remove(dn)
            removed.append(n)
    files["xl/workbook.xml"] = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)
    for p in files:
        if p.endswith(".rels") or p == "[Content_Types].xml":
            assert b"customXml" not in files[p] and b"webextension" not in files[p] and b"printerSettings" not in files[p], p

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ["[Content_Types].xml"] + sorted(k for k in files if k != "[Content_Types].xml"):
            z.writestr(n, files[n])
    print("removed sheet:", DOC)
    print("removed Register columns T-W (Source Completion now T)")
    print("removed defined names:", ", ".join(removed))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
