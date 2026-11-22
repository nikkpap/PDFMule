"""Writes the build's Windows version resource (win_version.txt) and stamps the manifest, from APP_VERSION
in PDFMule.pyw. Usage: python tools/build_info.py  → prints the version."""
import re
from pathlib import Path

root = Path(__file__).resolve().parents[1]
src = (root / "PDFMule.pyw").read_text(encoding="utf-8")
ver = re.search(r'^APP_VERSION = "([0-9][0-9A-Za-z.\-]*)"', src, re.M).group(1)
nums = (re.findall(r"\d+", ver.split("-")[0]) + ["0", "0", "0", "0"])[:4]
t = ", ".join(nums)
dotted = ".".join(nums)
(root / "win_version.txt").write_text(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers=({t}), prodvers=({t}), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('040904B0', [
      StringStruct('CompanyName', 'Nikolaos K. Paridis (aka nikkpap) - ALU DEV TEAM'),
      StringStruct('FileDescription', 'PDF Mule - local PDF toolbox'),
      StringStruct('FileVersion', '{ver}'),
      StringStruct('InternalName', 'PDFMule'),
      StringStruct('LegalCopyright', '(c) 2026 Nikolaos K. Paridis (aka nikkpap)'),
      StringStruct('OriginalFilename', 'PDFMule_v{ver}-windows-x64.exe'),
      StringStruct('ProductName', 'PDF Mule'),
      StringStruct('ProductVersion', '{ver}')])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
""", encoding="utf-8")
man = root / "PDFMule.manifest"
man.write_text(re.sub(r'(assemblyIdentity version=")[0-9.]+(")', rf"\g<1>{dotted}\g<2>", man.read_text(encoding="utf-8")),
               encoding="utf-8")
print(ver)
