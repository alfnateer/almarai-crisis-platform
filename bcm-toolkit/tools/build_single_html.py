#!/usr/bin/env python3
"""Bundle the website into one self-contained HTML file (works offline, from e-mail or
a shared drive - no server needed):

    python3 tools/build_single_html.py dist/BCM_Toolkit.html

Inlines css/toolkit.css and every <script src> (JSZip, the embedded template, schema,
engine, checks, Excel I/O and app) and drops the link back to the crisis platform.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main(dst):
    page = (ROOT / "index.html").read_text(encoding="utf-8")
    css = (ROOT / "css" / "toolkit.css").read_text(encoding="utf-8")
    page = page.replace('<link rel="stylesheet" href="css/toolkit.css">', "<style>\n" + css + "\n</style>")

    def script(m):
        code = (ROOT / m.group(1)).read_text(encoding="utf-8")
        return "<script>\n" + code.replace("</script", "<\\/script") + "\n</script>"

    page, n = re.subn(r'<script src="([^"]+)"></script>', script, page)
    assert n >= 8 and "href=\"css/" not in page, n
    page = page.replace('        <a href="../index.html">Crisis Management Platform</a>\n', "")
    out = Path(dst)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(page, encoding="utf-8")
    print("wrote", out, "(%d KB, %d scripts inlined)" % (out.stat().st_size // 1024, n))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(ROOT / "dist" / "BCM_Toolkit.html"))
