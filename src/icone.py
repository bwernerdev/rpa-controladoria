"""Icone do app, desenhado em codigo. Executado como script, grava assets/icone.ico."""
import sys
from pathlib import Path

from PIL import Image, ImageDraw


def desenhar(tamanho=256):
    escala = 4
    lado = tamanho * escala
    imagem = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
    d = ImageDraw.Draw(imagem)
    d.rounded_rectangle((0, 0, lado - 1, lado - 1), radius=lado // 5, fill=(15, 23, 42, 255))
    # cabeca do robo
    m = lado // 5
    d.rounded_rectangle((m, lado * 0.30, lado - m, lado * 0.80), radius=lado // 10, fill=(37, 99, 235, 255))
    # olhos
    r = lado // 14
    for cx in (lado * 0.38, lado * 0.62):
        cy = lado * 0.52
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(255, 255, 255, 255))
    # antena
    d.line((lado / 2, lado * 0.30, lado / 2, lado * 0.16), fill=(37, 99, 235, 255), width=lado // 28)
    d.ellipse((lado / 2 - r * 0.8, lado * 0.16 - r * 1.6, lado / 2 + r * 0.8, lado * 0.16), fill=(34, 211, 238, 255))
    # boca (indicador)
    d.rounded_rectangle((lado * 0.36, lado * 0.66, lado * 0.64, lado * 0.70), radius=lado // 60, fill=(147, 197, 253, 255))
    return imagem.resize((tamanho, tamanho), Image.LANCZOS)


if __name__ == "__main__":
    destino = Path(__file__).resolve().parent.parent / "assets"
    destino.mkdir(exist_ok=True)
    desenhar(256).save(destino / "icone.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("assets/icone.ico gerado")
    sys.exit(0)
