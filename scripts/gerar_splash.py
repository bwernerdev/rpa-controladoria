"""Gera build/splash.png (tela de abertura do exe) com a versao atual de src/versao.py. Chamado pelo build.ps1."""
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

raiz = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(raiz / "src"))
from icone import desenhar  # noqa: E402
from versao import NOME, VERSAO  # noqa: E402

FONTES = Path(r"C:\Windows\Fonts")


def fonte(arquivo, tamanho):
    try:
        return ImageFont.truetype(str(FONTES / arquivo), tamanho)
    except OSError:
        return ImageFont.load_default(tamanho)


imagem = Image.new("RGB", (420, 170), (246, 247, 249))
d = ImageDraw.Draw(imagem)
d.rectangle((0, 0, 419, 169), outline=(201, 206, 214))
icone = desenhar(96)
imagem.paste(icone, (30, 37), icone)
d.text((148, 54), NOME, font=fonte("seguisb.ttf", 25), fill=(31, 35, 40))
d.text((148, 90), "Carregando...", font=fonte("segoeui.ttf", 14), fill=(107, 114, 128))
d.text((148, 110), f"v{VERSAO}", font=fonte("segoeui.ttf", 13), fill=(160, 166, 175))
destino = raiz / "build" / "splash.png"
destino.parent.mkdir(exist_ok=True)
imagem.save(destino)
print(f"build/splash.png gerado (v{VERSAO})")
