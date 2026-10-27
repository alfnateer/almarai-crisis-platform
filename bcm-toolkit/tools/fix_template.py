#!/usr/bin/env python3
"""Turn the blank BCM Toolkit Template v5 into v5.1: fix its logic errors, make it
generic, colour dropdowns and highlight missing critical entries.

    python3 tools/fix_template.py template/source/BCM_Toolkit_Template_v5_blank.xlsx template/BCM_Toolkit_Template_v5.1.xlsx

Run on the blank produced by make_blank_template.py. Every edit asserts that it found
what it expects, so running it on another workbook version fails loudly instead of
producing a half-fixed file.

Logic fixes (v5 defects):
  10. Risk Assessment   Risk Score was INDEX(#REF!, ...) after the ERM guidance sheet was
                        removed - now Likelihood x Impact as the header states, so Risk
                        Rating / BCP Priority / Register / Dashboard work again
  10. Risk Assessment   Source Completion was an empty input - now calculated (the BCM
                        Risk Register's escalation depends on it)
  13. IT DR Register    empty leftover validation rule on column F removed
  02. Dashboard         DR counters covered rows 5-34 only - now the whole register
  05. Processes         Product & Services shows blank (not 0) when no product code
  16. CAPA Register     Overdue counted Cancelled CAPAs
  14. R&R / 11. S&S     duplicate checks used COUNTIFS, which fails on text > 255 chars
Generic content:
  "Annual Revenue (SAR)" -> local currency; "Food Safety" -> "Product / Service Safety";
  company examples in the column notes replaced with neutral ones
Additions:
  every dropdown (list validation) cell filled yellow; red highlight on required entries
  left empty in a row in use, on RTO >= MTPD and on S&S requirement flags set to Yes
  without a matching row in 12. Requirements Register (amber when not all Ready)
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
def replace_shared_string(files, old, new):
    sst = files["xl/sharedStrings.xml"].decode("utf-8")
    pat = re.compile(r"(<t(?: [^>]*)?>)" + re.escape(html.escape(old, quote=False)) + r"(</t>)")
    sst, n = pat.subn(lambda m: m.group(1) + html.escape(new, quote=False) + m.group(2), sst)
    assert n >= 1, old
    files["xl/sharedStrings.xml"] = sst.encode("utf-8")


def replace_comment(files, old, new):
    n = 0
    for name in files:
        if re.match(r"xl/comments\d+\.xml$", name):
            t = files[name].decode("utf-8")
            if html.escape(old, quote=False) in t:
                t = t.replace(html.escape(old, quote=False), html.escape(new, quote=False))
                files[name] = t.encode("utf-8")
                n += 1
    assert n >= 1, old


def set_table_formula(files, table_name, column, formula):
    """Give a table column a calculated-column formula (adds or replaces it)."""
    for name in files:
        if not (name.startswith("xl/tables/") and name.endswith(".xml")):
            continue
        t = files[name].decode("utf-8")
        if 'displayName="%s"' % table_name not in t and ' name="%s"' % table_name not in t:
            continue
        esc_col = re.escape(html.escape(column, quote=True))
        m = re.search(r'<tableColumn [^>]*name="%s"[^>]*?(/>|>.*?</tableColumn>)' % esc_col, t, re.S)
        assert m, (table_name, column)
        el = m.group(0)
        f = "<calculatedColumnFormula>%s</calculatedColumnFormula>" % html.escape(formula, quote=False)
        if el.endswith("/>"):
            new = el[:-2] + ">" + f + "</tableColumn>"
        else:
            new = re.sub(r"<calculatedColumnFormula[^>]*>.*?</calculatedColumnFormula>", lambda _: f, el, flags=re.S)
            if new == el:
                new = el.replace("</tableColumn>", f + "</tableColumn>")
        files[name] = (t[:m.start()] + new + t[m.end():]).encode("utf-8")
        return
    raise AssertionError("table %s not found" % table_name)


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

    # Generic wording ----------------------------------------------------------------
    replace_shared_string(files, "Annual Revenue (SAR)", "Annual Revenue (local currency)")
    replace_shared_string(files, "Food Safety", "Product / Service Safety")
    replace_comment(files, "e.g. Protein, Bakery, International", "e.g. Retail, Manufacturing, International")
    replace_comment(files, "e.g Protein, Bakery, International", "e.g. Retail, Manufacturing, International")
    replace_comment(files, "e.g Poultry Farming, Poultry Manufacturing, Bakery Alkharj, Bakery Hail, Jordan, Egypt",
                    "e.g. Head Office, Plant A, Plant B, Distribution Centre, Country X")
    replace_comment(files, "e.g. Poultry Farming, Bakry Alkharj, Jordan", "e.g. Plant A, Distribution Centre, Country X")

    # 10. Risk Assessment: Risk Score = Likelihood x Impact -----------------------------
    ra = S("10. Risk Assessment")
    score = 'IF(OR($K{r}="",$M{r}=""),"",$K{r}*$M{r})'
    for r in range(5, 35):
        f = ra.formula("O%d" % r)
        assert f is not None and "#REF!" in f.text, r
        f.text = score.format(r=r)
    set_table_formula(files, "RiskAssessment", "Risk Score (Likelihood × Impact)", score.format(r=5))

    # 10. Risk Assessment: Source Completion calculated --------------------------------
    done = ('IF($B{r}="","",IF(AND($C{r}<>"",$D{r}<>"",$E{r}<>"",$F{r}<>"",$G{r}<>"",$H{r}<>"",$J{r}<>"",'
            '$K{r}<>"",$M{r}<>"",$U{r}<>"",ISNUMBER($V{r}),ISNUMBER($W{r}),$X{r}<>"",$Y{r}<>""),"Complete","Incomplete"))')
    for r in range(5, 35):
        c = ra.cell("Z%d" % r, create=True)
        assert c.find(Q("f")) is None and c.find(Q("v")) is None, r
        set_formula(ra, "Z%d" % r, done.format(r=r), keep_attrs=False)
    set_table_formula(files, "RiskAssessment", "Source Completion", done.format(r=5))

    # 13. IT DR Register: an empty leftover validation on F (no rule) ------------------
    dr = S("13. IT DR Register")
    dvs = [dv for dv in dr.doc.iter(Q("dataValidation")) if dv.get("sqref") == "F5:F204"]
    assert len(dvs) == 1 and dvs[0].get("type") is None and dvs[0].find(Q("formula1")) is None
    parent = dvs[0].getparent()
    parent.remove(dvs[0])
    parent.set("count", str(len(parent)))

    # 02. Dashboard: DR counters over the whole register --------------------------------
    f = S("02. Dashboard").formula("H7")
    assert f.text.count("$34") == 2
    f.text = f.text.replace("$B$5:$B$34", "$B$5:$B$204").replace("$Q$5:$Q$34", "$Q$5:$Q$204")

    # 05. Processes: blank product code -> blank (VLOOKUP of an empty cell can match an
    # empty P&S row and show 0 outside Excel) ---------------------------------------------
    pat, rep = r"^IFERROR\(VLOOKUP\(B(\d+),", r'IF(B\1="","",IFERROR(VLOOKUP(B\1,'
    tail, tail_rep = r'\),""\)$', r'),""))'
    sub_formulas(S("05. Processes & Activities"), "C", pat + r"(.*)" + tail, rep + r"\2" + tail_rep, expect_min=207)
    sub_table(files, "ProcessesActivities", "Product &amp; Services", pat + r"(.*)" + tail, rep + r"\2" + tail_rep)

    # 16. CAPA: Overdue ignores Cancelled -------------------------------------------------
    pat, rep = r'OR\(M(\d+)="Closed",L', r'OR(M\1="Closed",M\1="Cancelled",L'
    sub_formulas(S("16. CAPA Register"), "T", pat, rep)
    sub_table(files, "CAPARegister", "Overdue", pat, rep)

    # Duplicate checks: SUMPRODUCT instead of COUNTIFS (255-character limit) -----------
    dup = r"COUNTIFS\((\$\w\$\d+:\$\w\$\d+),(\$\w\d+),(\$\w\$\d+:\$\w\$\d+),(\$\w\d+),(\$\w\$\d+:\$\w\$\d+),(\$\w\d+)\)"
    dup_rep = r"SUMPRODUCT(--(\1=\2),--(\3=\4),--(\5=\6))"
    sub_formulas(S("14. Recovery & Response"), "Q", dup, dup_rep)
    sub_table(files, "RecoveryResponsePlans", "Duplicate Check", dup, dup_rep)
    sub_formulas(S("11. Strategies & Solutions"), "U", dup, dup_rep)
    sub_table(files, "StrategiesSolutions", "1:1 Check", dup, dup_rep)

    # Dropdown cells in yellow ------------------------------------------------------------
    painted = 0
    for name, sh in sheets.items():
        ranges = [dv.get("sqref") for dv in sh.doc.iter(Q("dataValidation")) if dv.get("type") == "list"]
        for dv in sh.doc.iter("{%s}dataValidation" % NS_X14):
            if dv.get("type") == "list":
                ranges.append(dv.find("{%s}sqref" % NS_XM).text)
        for sq in ranges:
            for addr in expand(sq):
                c = sh.cell(addr, create=True)
                c.set("s", str(styles.yellow(c.get("s"))))
                painted += 1

    # Missing-entry highlights ------------------------------------------------------------
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
        ("08. Dependency Register", "C5:C204 F5:F204 K5:L204", "$I5", "C5"),
        ("10. Risk Assessment", "C5:E34 J5:K34 M5:M34 U5:Y34", "$G5", "C5"),
        ("11. Strategies & Solutions", "H5:I64 P5:R64", "$B5", "H5"),
        ("12. Requirements Register", "C5:C304 E5:E304 K5:L304", "$F5", "C5"),
        ("13. IT DR Register", "C5:D204 G5:K204 M5:M204", "$E5", "C5"),
        ("14. Recovery & Response", "E5:F104 L5:L104", "$C5", "E5"),
        ("16. CAPA Register", "C6:E105 G6:G105 K6:M105", "$F6", "C6"),
    ]
    for name, sq, key, first in required:
        add_cf(S(name), sq, 'AND(%s<>"",%s="")' % (key, first), red)
    add_cf(S("07. BIA"), "AC6:AD285", "AND(ISNUMBER($AD6),ISNUMBER($AA6),$AD6>=$AA6)", red)
    add_cf(ra, "Z5:Z34", '$Z5="Incomplete"', amber)

    for name, sh in sheets.items():
        files[path[name]] = etree.tostring(sh.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["xl/styles.xml"] = styles.bytes()
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, files[info.filename])
    print("dropdown cells painted yellow:", painted)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
