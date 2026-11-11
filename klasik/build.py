"""Membangun klasik/build/Aktara-tanpa-makro.dotm (ribbon tanpa makro).
Makro VBA dari src/*.bas diimpor sekali di Word lalu disimpan (lihat README)."""
import zipfile, pathlib
root = pathlib.Path(__file__).parent
out = root / "build" / "Aktara-tanpa-makro.dotm"  # jangan timpa dist/Aktara.dotm (berisi makro)
out.parent.mkdir(exist_ok=True)
CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.template.macroEnabledTemplate.main+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rIdAktaraUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
</Relationships>"""
DOC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>Aktara Edisi Klasik - template add-in. Simpan di folder STARTUP Word.</w:t></w:r></w:p>
</w:body></w:document>"""
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", DOC)
    z.write(root / "ribbon" / "customUI.xml", "customUI/customUI.xml")
print("dibuat:", out)
