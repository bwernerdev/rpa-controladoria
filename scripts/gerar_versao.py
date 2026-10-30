"""Gera build/version.txt (recurso de versao do Windows) a partir de src/versao.py."""
import sys
from pathlib import Path

raiz = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(raiz / "src"))
from versao import NOME, VERSAO  # noqa: E402

partes = [int(p) for p in VERSAO.split(".")] + [0] * (4 - len(VERSAO.split(".")))
tupla = tuple(partes[:4])
modelo = f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={tupla}, prodvers={tupla}, mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041604b0', [
      StringStruct('CompanyName', ''),
      StringStruct('FileDescription', '{NOME} - preenchimento do SRP'),
      StringStruct('FileVersion', '{VERSAO}'),
      StringStruct('InternalName', 'RPA-SRP'),
      StringStruct('OriginalFilename', 'RPA-SRP.exe'),
      StringStruct('ProductName', '{NOME}'),
      StringStruct('ProductVersion', '{VERSAO}')])]),
    VarFileInfo([VarStruct('Translation', [1046, 1200])])
  ]
)
"""
(raiz / "build").mkdir(exist_ok=True)
(raiz / "build" / "version.txt").write_text(modelo, encoding="utf-8")
print("build/version.txt gerado para", VERSAO)
