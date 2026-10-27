#!/usr/bin/env python3
"""Apply the toolkit's formula fixes, dropdown colouring and missing-entry highlights.

    python3 tools/fix_template.py template/in.xlsx template/out.xlsx

Run on the blank template produced by make_blank_template.py. Every edit asserts that
it found what it expects, so running it on a different workbook version fails loudly
instead of producing a half-fixed file.

Fixes (v0.2 defects):
  15. BCM Risk Register  S header renamed (it shows New/Existing); T-W blank (columns no
                         longer captured); X Source Completion computed from the Risk
                         Assessment fields, so P/Q work again
  17. PDCA               ID prefixes PS- / RR- (were P&S- / R&R-); step 6 computed from
                         the Risk Assessment fields (was #REF!)
  02. Dashboard          DR counters cover the whole IT DR register (were rows 5-34)
  16. CAPA               Overdue ignores Cancelled CAPAs again
  07. BIA                CEO MTPD blank instead of #N/A
  14. R&R / 11. S&S      duplicate checks use SUMPRODUCT (COUNTIFS fails on text > 255 chars)
  21. Analysis           exact-match lookups on every row, no #N/A; J/K show BIA RTO / RPO
  10. Risk Assessment    stray =57/3 in G3 removed
Additions:
  - every dropdown (list validation) cell filled yellow
  - red highlight on required entries left empty in a filled row, on RTO >= MTPD, and on
    S&S requirement flags set to Yes without a matching row in 12. Requirements Register
    (amber when rows exist but are not all Ready)
"""
import copy
import html
import re
import sys
import zipfile

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
Q = lambda t: "{%s}%s" % (NS, t)

YELLOW = "FFFFFF99"
RED_FONT, RED_FILL = "FF9C0006", "FFFFC7CE"
AMBER_FONT, AMBER_FILL = "FF9C5700", "FFFFEB9C"


def col_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def col_name(n):
    s = ""
    while n:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


def split(addr):
    m = re.match(r"([A-Z]+)(\d+)$", addr)
    return col_num(m.group(1)), int(m.group(2))


def expand(sqref):
    for part in sqref.split():
        a, b = (part.split(":") + [part])[:2]
        c1, r1 = split(a)
        c2, r2 = split(b)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                yield "%s%d" % (col_name(c), r)


class Sheet:
    def __init__(self, doc):
        self.doc = doc
        self.data = doc.find(Q("sheetData"))
        self.rows = {int(r.get("r")): r for r in self.data.findall(Q("row"))}

    def cell(self, addr, create=False):
        c, r = split(addr)
        row = self.rows.get(r)
        if row is None:
            if not create:
                return None
            row = etree.Element(Q("row"), r=str(r))
            later = [k for k in self.rows if k > r]
            if later:
                self.rows[min(later)].addprevious(row)
            else:
                self.data.append(row)
            self.rows[r] = row
        before = None
        for x in row.findall(Q("c")):
            if x.get("r") == addr:
                return x
            if before is None and split(x.get("r"))[0] > c:
                before = x
        if not create:
            return None
        new = etree.Element(Q("c"), r=addr)
        prev = before.getprevious() if before is not None else (row.findall(Q("c")) or [None])[-1]
        if prev is not None and prev.get("s"):
            new.set("s", prev.get("s"))
        if before is not None:
            before.addprevious(new)
        else:
            row.append(new)
        row.attrib.pop("spans", None)
        return new

    def formula(self, addr):
        c = self.cell(addr)
        f = c.find(Q("f")) if c is not None else None
        return f


def set_formula(sheet, addr, text, keep_attrs=True):
    c = sheet.cell(addr, create=True)
    f = c.find(Q("f"))
    attrs = dict(f.attrib) if (f is not None and keep_attrs) else {}
    for ch in list(c):
        c.remove(ch)
    for a in ("t", "vm") + (() if keep_attrs else ("cm",)):
        c.attrib.pop(a, None)
    nf = etree.SubElement(c, Q("f"), **attrs)
    nf.text = text


def set_text(sheet, addr, text):
    c = sheet.cell(addr, create=True)
    for ch in list(c):
        c.remove(ch)
    c.set("t", "inlineStr")
    is_ = etree.SubElement(c, Q("is"))
    t = etree.SubElement(is_, Q("t"))
    t.text = text


def sub_formulas(sheet, col, pattern, repl, expect_min=1):
    """Regex-replace in every formula (masters and plain) of a column."""
    n = 0
    for row in sheet.rows.values():
        for c in row.findall(Q("c")):
            if re.match(r"%s\d+$" % col, c.get("r")):
                f = c.find(Q("f"))
                if f is not None and f.text:
                    new, k = re.subn(pattern, repl, f.text)
                    if k:
                        f.text = new
                        n += 1
    assert n >= expect_min, (col, pattern, n)
    return n


def sub_table(files, table_name, column, pattern, repl):
    for name in files:
        if name.startswith("xl/tables/") and name.endswith(".xml"):
            t = files[name].decode("utf-8")
            if 'name="%s"' % table_name not in t and 'displayName="%s"' % table_name not in t:
                continue
            m = re.search(r'(<tableColumn [^>]*name="%s"[^>]*>\s*<calculatedColumnFormula>)(.*?)(</calculatedColumnFormula>)' % re.escape(column), t, re.S)
            assert m, (table_name, column)
            body = html.unescape(m.group(2))
            new, k = re.subn(pattern, repl, body)
            assert k, (table_name, column, pattern)
            t = t[:m.start(2)] + html.escape(new, quote=False) + t[m.end(2):]
            files[name] = t.encode("utf-8")
            return
    raise AssertionError("table %s not found" % table_name)


# --------------------------------------------------------------------- styles
class Styles:
    def __init__(self, xml):
        self.doc = etree.fromstring(xml)
        self.fills = self.doc.find(Q("fills"))
        self.xfs = self.doc.find(Q("cellXfs"))
        self.dxfs = self.doc.find(Q("dxfs"))
        if self.dxfs is None:
            self.dxfs = etree.Element(Q("dxfs"), count="0")
            self.doc.find(Q("cellStyles")).addnext(self.dxfs)
        self.yellow_fill = self._add_fill(YELLOW)
        self.yellow_xf = {}
        self.red_dxf = self._add_dxf(RED_FONT, RED_FILL)
        self.amber_dxf = self._add_dxf(AMBER_FONT, AMBER_FILL)

    def _add_fill(self, rgb):
        fill = etree.SubElement(self.fills, Q("fill"))
        pf = etree.SubElement(fill, Q("patternFill"), patternType="solid")
        etree.SubElement(pf, Q("fgColor"), rgb=rgb)
        etree.SubElement(pf, Q("bgColor"), indexed="64")
        self.fills.set("count", str(len(self.fills)))
        return len(self.fills) - 1

    def _add_dxf(self, font, fill):
        d = etree.SubElement(self.dxfs, Q("dxf"))
        f = etree.SubElement(d, Q("font"))
        etree.SubElement(f, Q("color"), rgb=font)
        fl = etree.SubElement(d, Q("fill"))
        pf = etree.SubElement(fl, Q("patternFill"))
        etree.SubElement(pf, Q("bgColor"), rgb=fill)
        self.dxfs.set("count", str(len(self.dxfs)))
        return len(self.dxfs) - 1

    def yellow(self, s):
        s = int(s or 0)
        if s not in self.yellow_xf:
            xf = copy.deepcopy(self.xfs[s])
            xf.set("fillId", str(self.yellow_fill))
            xf.set("applyFill", "1")
            self.xfs.append(xf)
            self.xfs.set("count", str(len(self.xfs)))
            self.yellow_xf[s] = len(self.xfs) - 1
        return self.yellow_xf[s]

    def bytes(self):
        return etree.tostring(self.doc, xml_declaration=True, encoding="UTF-8", standalone=True)


CF_ORDER = ["conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
            "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors",
            "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
            "webPublishItems", "tableParts", "extLst"]


def add_cf(sheet, sqref, formula, dxf):
    doc = sheet.doc
    prios = [int(r.get("priority")) for r in doc.iter(Q("cfRule"))]
    cf = etree.Element(Q("conditionalFormatting"), sqref=sqref)
    rule = etree.SubElement(cf, Q("cfRule"), type="expression", dxfId=str(dxf), priority=str(max(prios + [0]) + 1))
    etree.SubElement(rule, Q("formula")).text = formula
    existing = doc.findall(Q("conditionalFormatting"))
    if existing:
        existing[-1].addnext(cf)
        return
    for tag in CF_ORDER[1:]:
        el = doc.find(Q(tag))
        if el is not None:
            el.addprevious(cf)
            return
    doc.append(cf)


# --------------------------------------------------------------------- main
def main(src, dst):
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    wb = etree.fromstring(files["xl/workbook.xml"])
    rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
    target = {r.get("Id"): r.get("Target") for r in rels.iter("{%s}Relationship" % NS_PR)}
    path = {}
    for s in wb.iter(Q("sheet")):
        t = target[s.get("{%s}id" % NS_R)]
        path[s.get("name")] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    sheets = {name: Sheet(etree.fromstring(files[p])) for name, p in path.items()}
    S = lambda n: sheets[n]
    styles = Styles(files["xl/styles.xml"])
    RA = "INDEX(RiskAssessment[],0,%d)"
    # Risk Assessment fields that make a risk record complete: BU, Division, Dependency,
    # Applicable P&S, Risk Assumption, ERM ref, Controls, Likelihood, Impact.
    REQ_RA = [3, 4, 5, 6, 7, 9, 10, 11, 13]

    # 15. BCM Risk Register ---------------------------------------------------
    reg = S("15. BCM Risk Register")
    set_text(reg, "S5", "Risks - New / Existing (auto)")
    for col, label in (("T", "Assessment Date"), ("U", "Review Date"), ("V", "Approver"), ("W", "Evidence Reference")):
        set_text(reg, col + "5", label + " (not captured in 10. Risk Assessment)")
        assert reg.formula(col + "6").get("t") == "array"
        reg.formula(col + "6").text = 'IF(_xlfn.ANCHORARRAY(A6)="","","")'
    checks = ",".join('INDEX(INDEX(RiskAssessment[],0,%d),_xlpm.r)<>""' % k for k in REQ_RA)
    assert reg.formula("X6").get("t") == "array"
    reg.formula("X6").text = (
        '_xlfn.MAP(_xlfn.ANCHORARRAY(A6),_xlfn.LAMBDA(_xlpm.id,IF(_xlpm.id="","",_xlfn.LET(_xlpm.r,'
        '_xlfn.XMATCH(_xlpm.id,INDEX(RiskAssessment[],0,2)),IF(AND(' + checks + '),"Complete","Incomplete")))))')

    # 17. PDCA ----------------------------------------------------------------
    pdca = S("17. BCMS Progress (PDCA)")
    for addr, a, b in (("D7", '"P&S-*"', '"PS-*"'), ("D13", '"R&R-*"', '"RR-*"'), ("D15", '"R&R-*"', '"RR-*"')):
        f = pdca.formula(addr)
        assert a in f.text, addr
        f.text = f.text.replace(a, b)
    done = ",".join('--(%s<>"")' % (RA % k) for k in REQ_RA)
    pdca.formula("D10").text = (
        '_xlfn.LET(_xlpm.ids,INDEX(RiskAssessment[],0,2),_xlpm.n,COUNTIF(_xlpm.ids,"RA-*"),'
        '_xlpm.done,SUMPRODUCT(--(_xlpm.ids<>""),' + done + '),'
        'IF(_xlpm.n=0,"Not Assessed",IF(_xlpm.done=_xlpm.n,"Yes","Partial")))')

    # 02. Dashboard -------------------------------------------------------------
    f = S("02. Dashboard").formula("H7")
    assert f.text.count("$34") == 2
    f.text = f.text.replace("$B$5:$B$34", "$B$5:$B$204").replace("$Q$5:$Q$34", "$Q$5:$Q$204")

    # 16. CAPA: Overdue ignores Cancelled ---------------------------------------
    pat, rep = r'OR\(M(\d+)="Closed",L', r'OR(M\1="Closed",M\1="Cancelled",L'
    sub_formulas(S("16. CAPA Register"), "T", pat, rep)
    sub_table(files, "CAPARegister", "Overdue", pat, rep)

    # 07. BIA: CEO MTPD without #N/A --------------------------------------------
    n = sub_formulas(S("07. BIA"), "AB", r"^VLOOKUP\(F(\d+),(.*)\)$", r'IF($F\1="","",IFERROR(VLOOKUP($F\1,\2),""))')
    assert n == 40, n

    # Duplicate checks: SUMPRODUCT instead of COUNTIFS (255-character limit) ------
    dup = r"COUNTIFS\((\$\w\$\d+:\$\w\$\d+),(\$\w\d+),(\$\w\$\d+:\$\w\$\d+),(\$\w\d+),(\$\w\$\d+:\$\w\$\d+),(\$\w\d+)\)"
    dup_rep = r"SUMPRODUCT(--(\1=\2),--(\3=\4),--(\5=\6))"
    sub_formulas(S("14. Recovery & Response"), "Q", dup, dup_rep)
    sub_table(files, "RecoveryResponsePlans", "Duplicate Check", dup, dup_rep)
    sub_formulas(S("11. Strategies & Solutions"), "U", dup, dup_rep)
    sub_table(files, "StrategiesSolutions", "1:1 Check", dup, dup_rep)

    # 21. Analysis ----------------------------------------------------------------
    ana = S("21. Analysis")
    set_text(ana, "J1", "BIA RTO (hours)")
    set_text(ana, "K1", "BIA RPO (hours)")
    def look(r, src_sheet, key_col, val_col, lo, hi):
        # Exact match; a blank source cell stays blank (VLOOKUP would show 0).
        x = "INDEX('%s'!$%s$%d:$%s$%d,MATCH($B%d,'%s'!$%s$%d:$%s$%d,0))" % (
            src_sheet, val_col, lo, val_col, hi, r, src_sheet, key_col, lo, key_col, hi)
        return 'IF($B%d="","",IFERROR(IF(%s="","",%s),""))' % (r, x, x)
    for r in range(2, 47):
        for col, src in (("C", "C"), ("D", "D"), ("E", "E"), ("F", "K"), ("G", "L")):
            set_formula(ana, "%s%d" % (col, r), look(r, "04. Products & Services", "B", src, 4, 49), keep_attrs=False)
        for col, src in (("I", "AA"), ("J", "AD"), ("K", "AE")):
            set_formula(ana, "%s%d" % (col, r), look(r, "07. BIA", "F", src, 4, 285), keep_attrs=False)

    # 10. Risk Assessment: stray formula -------------------------------------------
    g3 = S("10. Risk Assessment").cell("G3")
    assert g3 is not None and g3.find(Q("f")).text == "57/3"
    g3.getparent().remove(g3)

    # Dropdown cells in yellow ------------------------------------------------------
    painted = 0
    for name, sh in sheets.items():
        ranges = []
        for dv in sh.doc.iter(Q("dataValidation")):
            if dv.get("type") == "list":
                ranges.append(dv.get("sqref"))
        for dv in sh.doc.iter("{%s}dataValidation" % NS_X14):
            if dv.get("type") == "list":
                ref = dv.find("{%s}sqref" % NS_XM)
                ranges.append(ref.text)
        for sq in ranges:
            for addr in expand(sq):
                c = sh.cell(addr, create=True)
                c.set("s", str(styles.yellow(c.get("s"))))
                painted += 1

    # Missing-entry highlights ------------------------------------------------------
    red, amber = styles.red_dxf, styles.amber_dxf
    REQ = "'12. Requirements Register'"
    ss = S("11. Strategies & Solutions")
    rule = [r for r in ss.doc.iter(Q("cfRule")) if r.getparent().get("sqref") == "J5:O64"]
    assert len(rule) == 1
    rule[0].find(Q("formula")).text = 'AND(J5="Yes",$C5<>"",COUNTIFS(%s!$C$5:$C$304,$C5,%s!$E$5:$E$304,J$4)=0)' % (REQ, REQ)
    rule[0].set("dxfId", str(red))
    add_cf(ss, "J5:O64", 'AND(J5="Yes",$C5<>"",COUNTIFS(%s!$C$5:$C$304,$C5,%s!$E$5:$E$304,J$4,%s!$L$5:$L$304,"<>Ready")>0)' % (REQ, REQ, REQ), amber)
    required = [
        ("04. Products & Services", "C5:D49 G5:I49", "$E5", "C5"),
        ("05. Processes & Activities", "D5:E211 H5:I211", "$B5", "D5"),
        ("07. BIA", "V6:V285 AC6:AC285 AF6:AF285", "$E6", "V6"),
        ("08. Dependency Register", "C5:C209 F5:F209 K5:L209", "$I5", "C5"),
        ("10. Risk Assessment", "C5:E61 I5:K61 M5:M61", "$G5", "C5"),
        ("11. Strategies & Solutions", "H5:I64 P5:R64", "$B5", "H5"),
        ("12. Requirements Register", "C5:C304 E5:E304 K5:L304", "$F5", "C5"),
        ("13. IT DR Register", "C5:D204 G5:K204 M5:M204", "$E5", "C5"),
        ("14. Recovery & Response", "E5:F104 L5:L104", "$C5", "E5"),
        ("16. CAPA Register", "C6:E105 G6:G105 K6:M105", "$F6", "C6"),
    ]
    for name, sq, key, first in required:
        add_cf(S(name), sq, 'AND(%s<>"",%s="")' % (key, first), red)
    add_cf(S("07. BIA"), "AC6:AD285", "AND(ISNUMBER($AD6),ISNUMBER($AA6),$AD6>=$AA6)", red)

    for name, sh in sheets.items():
        files[path[name]] = etree.tostring(sh.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["xl/styles.xml"] = styles.bytes()
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, files[info.filename])
    print("dropdown cells painted yellow:", painted)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
