"""Leitura/validacao da planilha e regras de tela (campo ciano, coluna do foco, Conta pelo vao)."""
import pathlib
import tempfile
import unittest

from openpyxl import Workbook
from PIL import Image, ImageDraw

from simulador import CENTROS, core


class Planilha(unittest.TestCase):
    def criar(self, linhas, cabecalho=("Item", "C/Custo", "Conta", "Quantidade")):
        livro = Workbook()
        folha = livro.active
        folha.append(list(cabecalho))
        for linha in linhas:
            folha.append(list(linha))
        caminho = pathlib.Path(tempfile.mkdtemp()) / "teste.xlsx"
        livro.save(caminho)
        return caminho

    def test_le_e_normaliza(self):
        regs = core.ler_planilha(self.criar([(1001, "20", 3101, 5), ("0023", 30, None, 2.5), (None, None, None, None)]))
        self.assertEqual(len(regs), 2)  # linha vazia ignorada
        self.assertEqual(regs[0], {"item": "1001", "custo": "20", "conta": "3101", "quantidade": "5", "linha": 2})
        self.assertEqual(regs[1]["item"], "0023")  # zero a esquerda preservado quando a celula e texto
        self.assertEqual(regs[1]["quantidade"], "2.5")

    def test_cabecalho_faltando(self):
        with self.assertRaises(ValueError):
            core.ler_planilha(self.criar([(1, 2, 3)], cabecalho=("Item", "C/Custo", "Conta")))

    def test_validacao(self):
        erros = core.validar(core.ler_planilha(self.criar([(1, "", 3, 4), (5, 6, 7, "x")])))
        self.assertTrue(any("custo em branco" in e for e in erros))
        self.assertTrue(any("quantidade invalida" in e for e in erros))

    def test_itens_repetidos(self):
        regs = core.ler_planilha(self.criar([(10, 1, 1, 1), (20, 1, 1, 1), (10, 1, 1, 1)]))
        self.assertEqual(core.itens_repetidos(regs), {"10": [2, 4]})


class Tela(unittest.TestCase):
    def test_maior_bloco_ciano_ignora_pixels_soltos(self):
        imagem = Image.new("RGB", (800, 500), (240, 240, 240))
        d = ImageDraw.Draw(imagem)
        for x, y in ((10, 10), (300, 40), (700, 450)):
            d.rectangle((x, y, x + 2, y + 2), fill=(0, 255, 255))
        d.rectangle((100, 20, 104, 480), fill=(0, 255, 255))  # traco fino e longo
        d.rectangle((500, 300, 560, 315), fill=(0, 255, 255))
        d.rectangle((600, 200, 710, 216), fill=(0, 255, 255))  # campo em foco (maior)
        esq, topo, dir_, base = core.maior_caixa_ciano(imagem)
        self.assertTrue(600 <= esq <= 604 and 706 <= dir_ <= 711 and 200 <= topo <= 204)
        self.assertIsNone(core.maior_caixa_ciano(Image.new("RGB", (50, 50), (255, 255, 255))))

    def test_icone_do_aviso_ignora_outros_vermelhos(self):
        imagem = Image.new("RGB", (900, 500), (240, 240, 240))
        d = ImageDraw.Draw(imagem)
        d.rectangle((20, 400, 600, 440), fill=(255, 0, 0))   # faixa vermelha grande (nao e o icone)
        d.rectangle((700, 20, 703, 300), fill=(255, 0, 0))   # traco vermelho fino
        self.assertIsNone(core.icone_aviso(imagem))
        d.ellipse((300, 100, 322, 122), fill=(255, 0, 0))    # icone ~22 px
        esq, topo, dir_, base = core.icone_aviso(imagem)
        self.assertTrue(298 <= esq <= 306 and 98 <= topo <= 106 and 316 <= dir_ <= 324)
        self.assertLessEqual(core.maior_caixa_cor(imagem, "vermelho")[0], 24)  # o maior vermelho continua sendo a faixa

    def test_coluna_do_foco(self):
        self.assertEqual(core.qual_coluna((424, 503, 466, 514), CENTROS), "custo")
        self.assertEqual(core.qual_coluna((475, 503, 499, 514), CENTROS), "conta")
        self.assertEqual(core.qual_coluna((656, 503, 740, 514), CENTROS), "quantidade")
        self.assertIsNone(core.qual_coluna((1200, 503, 1250, 514), CENTROS))

    def test_conta_pelo_vao(self):
        custo = (424, 503, 466, 514)
        self.assertTrue(core.conta_editavel(custo, (475, 503, 499, 514)))
        self.assertFalse(core.conta_editavel(custo, (656, 503, 740, 514)))


class Progresso(unittest.TestCase):
    def test_salva_e_so_vale_para_a_mesma_planilha(self):
        from unittest import mock
        arquivo = pathlib.Path(tempfile.mkdtemp()) / "progresso.json"
        regs = [{"linha": i, "item": f"10{i}", "custo": "7300", "conta": "301", "quantidade": "1"} for i in (2, 3, 4)]
        with mock.patch.object(core, "arquivo_progresso", lambda: arquivo):
            self.assertEqual(core.ler_progresso(regs), (set(), None))
            core.gravar_progresso(regs, {2, 3})
            feitas, quando = core.ler_progresso(regs)
            self.assertEqual(feitas, {2, 3})
            self.assertTrue(quando)
            outra = [dict(r) for r in regs]
            outra[2]["quantidade"] = "9"  # mudou uma celula: o andamento nao vale
            self.assertEqual(core.ler_progresso(outra), (set(), None))
            core.gravar_progresso(regs, set())
            self.assertFalse(arquivo.exists())


class Logs(unittest.TestCase):
    def test_apaga_so_os_antigos(self):
        import os
        import time
        from unittest import mock
        pasta = pathlib.Path(tempfile.mkdtemp())
        (pasta / "ocr").mkdir()
        velho, novo, recorte = pasta / "velho.log", pasta / "novo.log", pasta / "ocr" / "velho.png"
        for arquivo in (velho, novo, recorte):
            arquivo.write_text("x")
        antigo = time.time() - 31 * 86400
        os.utime(velho, (antigo, antigo))
        os.utime(recorte, (antigo, antigo))
        with mock.patch.object(core, "pasta_logs", lambda: pasta):
            self.assertEqual(core.limpar_logs(30), 2)
        self.assertEqual([p.name for p in pasta.rglob("*.*")], ["novo.log"])


if __name__ == "__main__":
    unittest.main()
