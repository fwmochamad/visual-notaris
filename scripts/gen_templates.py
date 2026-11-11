"""Menghasilkan template akta untuk kedua edisi dari templates/akta.json.
   python scripts/gen_templates.py
"""
import json, pathlib

root = pathlib.Path(__file__).resolve().parent.parent
data = json.loads((root / "templates" / "akta.json").read_text(encoding="utf-8"))
tpl = {k: v for k, v in data.items() if not k.startswith("_")}

# ---- Edisi Modern (ES5) ----
js = ("/* DIHASILKAN OTOMATIS dari templates/akta.json oleh scripts/gen_templates.py - jangan diedit manual. */\n"
      "(function (root) { root.AktaraTemplates = " + json.dumps(tpl, ensure_ascii=False, indent=1) + "; })(this);\n")
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
