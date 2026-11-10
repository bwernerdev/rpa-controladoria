"""Textos da tela: plural certo (`qtd`) e acentos nas mensagens do robo mostradas no status e na faixa de pausa."""
import unittest

from simulador import core


class Textos(unittest.TestCase):
    def test_plural(self):
        self.assertEqual(core.qtd(1, "nota", "notas"), "1 nota")
        self.assertEqual(core.qtd(0, "nota", "notas"), "0 notas")
        self.assertEqual(core.qtd(15, "item", "itens"), "15 itens")

    def test_acentos(self):
        self.assertEqual(core.com_acentos("o cursor nao saiu do lugar. PAUSADO: termine a nota a mao."),
                         "o cursor não saiu do lugar. PAUSADO: termine a nota a mão.")
        self.assertEqual(core.com_acentos("Nao esta no Item da proxima linha"), "Não está no Item da próxima linha")
        self.assertEqual(core.com_acentos("NAO iniciei; a correta e 8712"), "NÃO iniciei; a correta é 8712")

    def test_acentos_nao_mexe_no_que_nao_conhece(self):
        # "e" e "esta" soltos sao ambiguos; nome de arquivo e caminho ficam como estao
        texto = r"Termine esta nota e abra proxima.xlsx em C:\logs\ultima"
        self.assertEqual(core.com_acentos(texto), texto)


if __name__ == "__main__":
    unittest.main()
