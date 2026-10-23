"""Turn the original red-text letters into fill-in templates.

Reads originals/*.docx, swaps every red run for {{placeholders}} (keeping the
run's font, size and bold), makes the text black, and writes:

  templates/*.docx    the placeholder templates
  templates.js        the same templates as base64, so index.html also works
                      when opened straight from disk (no web server needed)

Run from the repo root:  python3 tools/build_templates.py

Spec format, per template: one entry per non-blank red run, in document order
(body first, then headers/footers by part name). An entry is:
  "text"   new text; the original run's leading/trailing whitespace is kept
  "=text"  exact new text (no whitespace kept), "=" alone empties the run
  DEL      delete the whole paragraph that holds this run
SPANS clears every run after red run i up to and including red run j.
"""
import base64
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DEL = object()

CHILD_NAME = "{{child}} {{sName}}"

SPECS = {
    "1_Sponsorship_Letter": {
        "title": "Sponsorship Letter",
        "red": [
            "{{pTitle}}", "{{pName}}", "{{pAge}}", "{{pRel}}", "{{sTitle}}", "{{sName}}",
            "{{address}}", "{{spouseRel}}", "{{spTitle}}", "{{spName}}",
            CHILD_NAME + "{{siblingClause}}", "=", "=", "=",          # 10-13
            CHILD_NAME, CHILD_NAME + ".", "{{aProfession}}", "{{incomeMain}}",
            CHILD_NAME, "{{his}}", "{{school}}", "{{airfield}}", "{{childs}}",
            "{{he}}", "{{child}}", CHILD_NAME, "{{child}}",
            "{{propCents}}", "{{propPlace}}, {{propDistrict}},",
            "{{propValue}}/- ({{propWords}}", ")",
            "{{goldSov}}", "{{goldValue}}/- ({{goldWords}}", ")",
            "{{incomeValue}}/- ({{incomeWords}}", ")", "{{incomeVia}}",
            "{{child}}", "{{his}}", "{{he}}", "{{his}}",
            "{{child}}", "{{he}}", "{{he}}", "{{his}}",
            CHILD_NAME, "{{his}}", "{{his}}",
            "{{pName}}", "({{pRelCap}})", "{{addressBlock}}", DEL,
            "{{footMob}}", "{{footEmail}}",
        ],
        "spans": {10: 13},
        "black": [("income derived from my salary.", "income derived from my {{incomeShort}}."),
                  ("I am working as a ", "I am working as ")],
    },
    "2_Self_Affidavit": {
        "title": "Self Affidavit",
        "red": [
            "{{sTitle}}. {{sName}} {{sdo}} {{pTitle}}. {{pName}}, {{address}}",
            "{{school}},", "{{airfield}}", "{{passport}}",
            "{{sName}}", "{{addressBlock}}",
            "{{footMob}}", "{{footEmail}}", "=", "=",
        ],
    },
    "3_Affidavit_of_Support": {
        "title": "Affidavit of Support",
        "red": [
            "{{pTitle}}", "{{pName}}", "{{pAge}}",
            "{{pRel}} of {{sTitle}}. {{sName}}, {{pProfession}},",
            "{{sName}}", "{{school}}", ", {{airfield}}", "{{address}}",
            "{{childs}}", "{{his}}", "={{sa1}}", "={{sa2}}", "={{sa3}}",
            "{{childs}}", "={{sa1}}", "={{sa2}}", "={{sa3}}", "=", "=",
            CHILD_NAME, "{{his}}", "{{his}}", CHILD_NAME, "{{child}}",
            "{{childs}}", "{{he}}", "{{his}}", "{{childs}}", "{{he}}", "{{he}}", "{{his}}",
            "{{pName}}", "({{pRelCap}})", "{{addressBlock}}", DEL, DEL, DEL,
            "{{footMob}}", "{{footEmail}}",
        ],
    },
    "4_Financial_Support_Letter": {
        "title": "Financial Support Letter",
        "red": [
            "{{sTitle}}. {{sName}}",
            "{{pTitle}}. {{pName}}, {{pRel}} of {{sTitle}}. {{sName}},",
            "{{child}}", "{{his}}", "{{He}}", "{{school}}", "{{airfield}}",
            "{{child}}", "{{his}}", "{{his}}",
            "{{pName}} ({{pRelCap}})", "{{addressBlock}}", DEL, DEL, DEL,
            "{{footMob}}", "{{footEmail}}",
        ],
    },
}

# The blank "  /    /2026" date in every letter becomes {{date}}; when the
# date is left blank the app fills in the same "    /    /YYYY" gap.
DATE_RE = re.compile(r"^(\s*)/\s*/2026(\s*)$")


def w(tag):
    return "{%s}%s" % (W, tag)


def run_text(r):
    out = ""
    for c in r:
        tag = etree.QName(c).localname
        if tag == "t":
            out += c.text or ""
        elif tag == "br":
            out += "\n"
        elif tag == "tab":
            out += "\t"
    return out


def set_run_text(r, text):
    for c in list(r):
        if etree.QName(c).localname != "rPr":
            r.remove(c)
    for i, line in enumerate(text.split("\n")):
        if i:
            etree.SubElement(r, w("br"))
        if line:
            t = etree.SubElement(r, w("t"))
            t.text = line
            t.set(XML_SPACE, "preserve")


def is_red(r):
    c = r.find("w:rPr/w:color", NS)
    return c is not None and c.get(w("val"), "").lower() == "ff0000"


def fit_date(r):
    """Put {{date}} in place of the blank date. A written date is wider than
    the blank, so drop the padding after it and a few of the spaces that push
    "DATE :" to the right, or the date wraps onto the next line."""
    set_run_text(r, "{{date}}")
    runs = list(r.getparent().iter(w("r")))
    i = runs.index(r)
    for after in runs[i + 1:]:
        if not run_text(after).strip():
            set_run_text(after, "")
    if i:
        before = run_text(runs[i - 1])
        set_run_text(runs[i - 1], re.sub(r" {20,}", lambda m: m.group(0)[:-10], before, count=1).rstrip() + " ")


def new_text(orig, spec):
    if spec.startswith("="):
        return spec[1:]
    lead = orig[: len(orig) - len(orig.lstrip())]
    trail = orig[len(orig.rstrip()):]
    return lead + spec + trail


def build(name, spec):
    src = ROOT / "originals" / f"{name}.docx"
    zin = zipfile.ZipFile(src)
    part_names = ["word/document.xml"] + sorted(
        n for n in zin.namelist() if re.match(r"word/(header|footer)\d+\.xml$", n)
    )
    roots = {p: etree.fromstring(zin.read(p)) for p in part_names}

    reds = [r for p in part_names for r in roots[p].iter(w("r")) if is_red(r) and run_text(r).strip()]
    if len(reds) != len(spec["red"]):
        sys.exit(f"{name}: found {len(reds)} red runs but spec has {len(spec['red'])}")

    for i, (r, s) in enumerate(zip(reds, spec["red"])):
        if s is DEL:
            p = r.getparent()
            p.getparent().remove(p)
            continue
        set_run_text(r, new_text(run_text(r), s))
        if i in spec.get("spans", {}):
            end = reds[spec["spans"][i]]
            siblings = list(r.getparent().iter(w("r")))
            for other in siblings[siblings.index(r) + 1: siblings.index(end) + 1]:
                set_run_text(other, "")

    for root in roots.values():
        for r in root.iter(w("r")):
            text = run_text(r)
            if DATE_RE.match(text):
                fit_date(r)
            for old, new in spec.get("black", []):
                if old in text:
                    set_run_text(r, text.replace(old, new))
        # Everything left red (blank runs, paragraph marks) becomes default colour.
        for c in root.iter(w("color")):
            if c.get(w("val"), "").lower() == "ff0000":
                c.getparent().remove(c)

    out = ROOT / "templates" / f"{name}.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in roots:
                data = etree.tostring(roots[item.filename], xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    return out


def main():
    (ROOT / "templates").mkdir(exist_ok=True)
    bundle = []
    for name, spec in SPECS.items():
        out = build(name, spec)
        bundle.append({"file": name, "title": spec["title"],
                       "data": base64.b64encode(out.read_bytes()).decode()})
        print("built", out.relative_to(ROOT))
    js = "// Generated by tools/build_templates.py - do not edit by hand.\n"
    js += "window.LETTER_TEMPLATES = " + json.dumps(bundle, indent=1) + ";\n"
    (ROOT / "templates.js").write_text(js)
    print("built templates.js")


if __name__ == "__main__":
    main()
