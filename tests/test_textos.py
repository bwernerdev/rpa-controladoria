"""Textos da tela: plural certo (`qtd`) e mensagens do robo com acento desde a fonte (v1.14.0), lidas sem acento por
quem as interpreta (cores do console, logs da SNAC0236), para os logs antigos (sem acento) continuarem valendo."""
import unittest

from simulador import core


class Textos(unittest.TestCase):
    def test_plural(self):
        self.assertEqual(core.qtd(1, "nota", "notas"), "1 nota")
        self.assertEqual(core.qtd(0, "nota", "notas"), "0 notas")
        self.assertEqual(core.qtd(15, "item", "itens"), "15 itens")

    def test_sem_acentos(self):
        self.assertEqual(core.sem_acentos("Não iniciei: o cursor está na Série. ATENÇÃO, último"),
                         "Nao iniciei: o cursor esta na Serie. ATENCAO, ultimo")
        self.assertEqual(core.sem_acentos("Nao iniciei"), "Nao iniciei")  # texto antigo fica igual

    def test_cor_do_console_vale_com_e_sem_acento(self):
        try:
            import visual
        except ImportError:  # sem customtkinter (a janela nao roda)
            self.skipTest("sem customtkinter")
        for texto in ("Não iniciei: nenhum campo do SRP em foco", "Nao iniciei: nenhum campo do SRP em foco"):
            self.assertEqual(visual.tipo_mensagem(texto), "erro")
        for texto in ("ATENÇÃO: não li na tela os rótulos", "ATENCAO: nao li na tela os rotulos"):
            self.assertEqual(visual.tipo_mensagem(texto), "aviso")

    def test_mensagens_nascem_com_acento(self):
        import pathlib
        fonte = (pathlib.Path(core.__file__).parent / "snac0236.py").read_text(encoding="utf-8")
        self.assertIn("PAUSADO: confira a integração no SRP", fonte)
        self.assertNotIn("PAUSADO: confira a integracao", fonte)


if __name__ == "__main__":
    unittest.main()
