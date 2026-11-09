"""shared v1.0 (panelbook 2.0.0). Copy cached formula results from a LibreOffice-recalculated copy into the openpyxl-built file,
so previewers show numbers while the original formatting stays untouched.
usage: python inject.py built.xlsx recalculated.xlsx out.xlsx"""
import sys, zipfile, posixpath
from lxml import etree
import openpyxl

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"

src, calc, out = sys.argv[1:4]
vals = openpyxl.load_workbook(calc, data_only=True)

zin = zipfile.ZipFile(src)
wbx = etree.fromstring(zin.read("xl/workbook.xml"))
rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
rid2t = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{PNS}}}Relationship")}
path2name = {}
for sh in wbx.find(f"{{{NS}}}sheets"):
    t = rid2t[sh.get(f"{{{RNS}}}id")]
    p = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
    path2name[p] = sh.get("name")

n_set = 0
zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename in path2name:
        ws = vals[path2name[item.filename]]
        root = etree.fromstring(data)
        for c in root.iter(f"{{{NS}}}c"):
            f = c.find(f"{{{NS}}}f")
            if f is None:
                continue
            v = c.find(f"{{{NS}}}v")
            if v is None:
                v = etree.SubElement(c, f"{{{NS}}}v")
            val = ws[c.get("r")].value
            if val is None:
                val = ""
            if isinstance(val, bool):
                c.set("t", "b"); v.text = "1" if val else "0"
            elif isinstance(val, (int, float)):
                if "t" in c.attrib:
                    del c.attrib["t"]
                v.text = repr(float(val)) if isinstance(val, float) else str(val)
            elif isinstance(val, str) and val.startswith("#") and val.endswith(("!", "?", "A")) and len(val) < 9:
                c.set("t", "e"); v.text = val
            else:
                c.set("t", "str"); v.text = str(val)
            n_set += 1
        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    zout.writestr(item, data)
zout.close()
print("cached values written:", n_set)
