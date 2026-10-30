"""OCR da quantidade e leitura do aviso do SRP (precisa do Tesseract em vendor/tesseract; senao os testes sao pulados)."""
import pathlib
import unittest

from PIL import Image, ImageDraw, ImageFont

from simulador import core

IMAGENS = pathlib.Path(__file__).resolve().parent / "imagens"
TESSERACT = core.caminho_tesseract()


@unittest.skipUnless(TESSERACT, "Tesseract nao encontrado em vendor/tesseract")
class QuantidadeNaCelula(unittest.TestCase):
    def test_celulas_reais_do_srp(self):
        """Recortes reais salvos pelo robo; 45.0000 era lido como 4 antes do ajuste do canal mais escuro."""
        for arquivo, esperado in (("celula_45.png", 45.0), ("celula_1.png", 1.0), ("celula_2.png", 2.0)):
            with self.subTest(arquivo=arquivo):
                confere, lidas = core.conferir_recorte(Image.open(IMAGENS / arquivo).convert("RGB"), TESSERACT, esperado)
                self.assertTrue(confere, f"lido {lidas}, esperado {esperado}")

    def test_bateria_sintetica(self):
        fonte = ImageFont.truetype("C:/Windows/Fonts/tahoma.ttf", 11)
        for valor in (1, 4, 7, 10, 11, 25, 38, 45, 58, 99, 100, 250, 1000, 0.5, 2.5, 12.25):
            texto = f"{valor:.4f}"
            imagem = Image.new("RGB", (92, 19), (255, 255, 255))
            desenho = ImageDraw.Draw(imagem)
            desenho.rectangle((0, 0, 91, 18), outline=(170, 170, 170))
            desenho.text((86 - desenho.textlength(texto, font=fonte), 3), texto, font=fonte, fill=(0, 0, 0))
            with self.subTest(valor=texto):
                confere, lidas = core.conferir_recorte(imagem, TESSERACT, float(valor))
                self.assertTrue(confere, f"lido {lidas}")


@unittest.skipUnless(TESSERACT, "Tesseract nao encontrado em vendor/tesseract")
class AvisoDoSRP(unittest.TestCase):
    @staticmethod
    def janela_de_aviso(linhas):
        """Desenha uma janela parecida com o aviso do SRP: icone vermelho puro, duas linhas de texto e botao OK."""
        imagem = Image.new("RGB", (900, 400), (236, 236, 236))
        d = ImageDraw.Draw(imagem)
        d.rectangle((250, 120, 640, 280), fill=(255, 255, 255), outline=(120, 120, 120))
        d.ellipse((275, 175, 315, 215), fill=(255, 0, 0))
        fonte = ImageFont.truetype("C:/Windows/Fonts/tahoma.ttf", 11)
        for i, linha in enumerate(linhas):
            d.text((335, 180 + 16 * i), linha, font=fonte, fill=(0, 0, 0))
        d.rounded_rectangle((570, 245, 625, 268), radius=8, outline=(60, 60, 140))
        d.text((590, 250), "OK", font=fonte, fill=(0, 0, 0))
        return imagem

    def test_conta_correta(self):
        info = core.ler_aviso(self.janela_de_aviso(["Conta Contabil invalida para o item 900135 filial 10/100",
                                                     "conta 301 c.custo 7300. Conta correta=8712"]), TESSERACT)
        self.assertIsNotNone(info)
        self.assertEqual((info["item"], info["custo"], info["conta_correta"]), ("900135", "7300", "8712"))
        self.assertIn("Conta correta=8712", info["texto"])

    def test_aviso_de_saldo_nao_tem_conta(self):
        info = core.ler_aviso(self.janela_de_aviso(["O saldo do item 900266 e insuficiente 45"]), TESSERACT)
        self.assertEqual(info["item"], "900266")
        self.assertIsNone(info["conta_correta"])

    def test_sem_aviso(self):
        self.assertIsNone(core.ler_aviso(Image.new("RGB", (400, 200), (240, 240, 240)), TESSERACT))


if __name__ == "__main__":
    unittest.main()
