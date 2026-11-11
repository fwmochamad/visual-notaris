"""Menghasilkan template akta untuk kedua edisi dari templates/akta.json.
   python scripts/gen_templates.py
"""
import json, pathlib

root = pathlib.Path(__file__).resolve().parent.parent
data = json.loads((root / "templates" / "akta.json").read_text(encoding="utf-8"))
tpl = {k: v for k, v in data.items() if not k.startswith("_")}
label = data["_label"]
import re
dipakai = {t for v in tpl.values() for baris in v["isi"] for t in re.findall(r"\{([A-Z0-9_]+)\}", baris)}
kurang = sorted(dipakai - set(label))
if kurang:
    raise SystemExit("Isian tanpa label lengkap di _label: " + ", ".join(kurang))

# ---- Edisi Modern (ES5) ----
js = ("/* DIHASILKAN OTOMATIS dari templates/akta.json oleh scripts/gen_templates.py - jangan diedit manual. */\n"
      "(function (root) {\n root.AktaraTemplates = " + json.dumps(tpl, ensure_ascii=False, indent=1) +
      ";\n root.AktaraLabel = " + json.dumps(label, ensure_ascii=False, indent=1) + ";\n})(this);\n")
(root / "addin" / "src" / "templates.js").write_text(js, encoding="utf-8")

# ---- Edisi Klasik (VBA) ----
def vb(s):
    return '"' + s.replace('"', '""') + '"'

lines = [
    'Attribute VB_Name = "AktaraTemplate"',
    "' DIHASILKAN OTOMATIS dari templates/akta.json oleh scripts/gen_templates.py - jangan diedit manual.",
    "Option Explicit",
    "",
    "Public Function DaftarTemplate() As Variant",
    "    DaftarTemplate = Array(" + ", ".join(vb(k) for k in tpl) + ")",
    "End Function",
    "",
    "Public Function NamaTemplate(ByVal kode As String) As String",
    "    Select Case kode",
] + ["    Case %s: NamaTemplate = %s" % (vb(k), vb(t["nama"])) for k, t in tpl.items()] + [
    "    End Select",
    "End Function",
    "",
    "' Label lengkap (tanpa singkatan) untuk setiap isian",
    "Public Function LabelIsian(ByVal tag As String) As String",
    "    Select Case tag",
] + ["    Case %s: LabelIsian = %s" % (vb(k), vb(v)) for k, v in label.items()] + [
    "    Case Else: LabelIsian = Replace(tag, \"_\", \" \")",
    "    End Select",
    "End Function",
    "",
    "' Mengisi judul(), isi() dan sifat untuk kode template. False bila kode tidak dikenal.",
    "Public Function AmbilTemplate(ByVal kode As String, ByRef judul() As String, ByRef isi() As String, ByRef sifat As String) As Boolean",
    "    Select Case kode",
]
for k, t in tpl.items():
    lines.append("    Case " + vb(k))
    lines.append("        sifat = " + vb(t["sifat"]))
    lines.append("        ReDim judul(0 To %d)" % (len(t["judul"]) - 1))
    for i, j in enumerate(t["judul"]):
        lines.append("        judul(%d) = %s" % (i, vb(j)))
    lines.append("        ReDim isi(0 To %d)" % (len(t["isi"]) - 1))
    for i, j in enumerate(t["isi"]):
        lines.append("        isi(%d) = %s" % (i, vb(j)))
lines += ["    Case Else", "        Exit Function", "    End Select", "    AmbilTemplate = True", "End Function", ""]
(root / "klasik" / "src" / "AktaraTemplate.bas").write_bytes("\r\n".join(lines).encode("cp1252"))
print("templates.js dan AktaraTemplate.bas dibuat:", ", ".join(tpl))

# ---- Ribbon Edisi Klasik (menu Akta Baru berisi semua template) ----
def xa(t):
    return t.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
menu = "\n".join('            <button id="tpl_%s" tag="%s" label="%s" imageMso="FileNew" onAction="AkTemplate"/>' % (k, k, xa(t["nama"])) for k, t in tpl.items())
ribbon = (root / "klasik" / "ribbon" / "customUI.template.xml").read_text(encoding="utf-8").replace("<!--MENU_TEMPLATE-->", menu)
(root / "klasik" / "ribbon" / "customUI.xml").write_text(ribbon, encoding="utf-8")
print("ribbon Klasik dibuat")
