"""Mengganti ribbon (customUI) di dist/Aktara.dotm tanpa menyentuh makro VBA di dalamnya.
   python klasik/update_ribbon.py"""
import zipfile, pathlib, shutil
root = pathlib.Path(__file__).parent
src = root / "dist" / "Aktara.dotm"
tmp = src.with_suffix(".tmp")
ui = (root / "ribbon" / "customUI.xml").read_bytes()
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = ui if item.filename.lower() == "customui/customui.xml" else zin.read(item.filename)
        zout.writestr(item, data)
shutil.move(tmp, src)
print("ribbon diperbarui di", src)
