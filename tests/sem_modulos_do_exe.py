"""Roda a suite com os modulos que o build.ps1 tira do exe bloqueados (import deles da ImportError), para pegar codigo
novo que passe a depender de um deles. Uso: .venv\Scripts\python.exe tests\sem_modulos_do_exe.py
1) importa todos os modulos de src com tudo bloqueado; 2) roda a suite liberando so a fonte (PIL._imagingft): os testes
de OCR desenham texto com Tahoma para simular a tela, o app nao. unittest/doctest ficam fora: a suite precisa deles."""
import importlib
import sys
import unittest
from pathlib import Path

BLOQUEADOS = ("PIL._avif", "PIL.AvifImagePlugin", "PIL._webp", "PIL.WebPImagePlugin", "PIL._imagingft",
              "PIL._imagingcms", "PIL.ImageCms", "setuptools", "pydoc_data", "pydoc", "pefile", "lib2to3", "xmlrpc", "darkdetect")
SO_NOS_TESTES = ("PIL._imagingft",)

pasta = Path(__file__).parent
src = pasta.parent / "src"
sys.path.insert(0, str(src))
for nome in BLOQUEADOS:
    sys.modules[nome] = None
for arquivo in sorted(src.glob("*.py")):
    importlib.import_module(arquivo.stem)
print(f"src: {len(list(src.glob('*.py')))} modulos importados sem {', '.join(BLOQUEADOS)}")
for nome in SO_NOS_TESTES + ("PIL.ImageFont",):  # o ImageFont guardou a fonte como faltando
    sys.modules.pop(nome, None)
import PIL
if hasattr(PIL, "ImageFont"):
    del PIL.ImageFont
resultado = unittest.TextTestRunner(verbosity=1).run(unittest.defaultTestLoader.discover(str(pasta)))
sys.exit(not resultado.wasSuccessful())
