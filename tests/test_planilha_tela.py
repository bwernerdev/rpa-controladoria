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

    def test_faixa_vermelha_fixa_do_cabecalho_nao_e_aviso(self):
        """Cabecalho do SRP: faixa vermelha cortada pelo logo; o pedaco da esquerda (17x21, na borda) parecia icone."""
        tela = Image.new("RGB", (1366, 768), (240, 240, 240))
        d = ImageDraw.Draw(tela)
        d.rectangle((0, 49, 19, 74), fill=(255, 0, 0))       # pedaco da faixa a esquerda do logo
        d.rectangle((51, 49, 457, 74), fill=(255, 0, 0))     # resto da faixa
        d.rectangle((300, 60, 330, 66), fill=(250, 250, 250))  # texto branco sobre a faixa
        fixos = core.caixas_cor(tela, "vermelho")
        self.assertIsNone(core.icone_aviso(tela))            # encostado na borda: descartado mesmo sem calibracao
        d.rectangle((200, 49, 222, 74), fill=(255, 0, 0))     # o mesmo tipo de pedaco longe da borda
        fixos = core.caixas_cor(tela, "vermelho")
        self.assertIsNone(core.icone_aviso(tela, fixos))     # ja estava na tela: nao e aviso
        d.ellipse((600, 380, 622, 402), fill=(255, 0, 0))    # aviso que aparece depois
        self.assertIsNotNone(core.icone_aviso(tela, fixos))

    def test_checagem_antes_de_comecar(self):
        tela = Image.new("RGB", (1366, 768), (240, 240, 240))
        d = ImageDraw.Draw(tela)
        d.rectangle((0, 49, 19, 74), fill=(255, 0, 0))        # faixa do cabecalho na borda: nao e aviso
        d.rectangle((44, 503, 118, 514), fill=(0, 255, 255))  # foco no Item
        d.text((50, 480), "Item  C/Custo  Conta  Quantidade", fill=(0, 0, 0))
        janela = (0, 0, 1366, 768)
        self.assertIsNone(core.problema_tela(janela, tela))
        self.assertIn("minimizada", core.problema_tela((-32000, -32000, -31840, -31972), tela))
        self.assertIn("vazia", core.problema_tela(janela, Image.new("RGB", (1366, 768), (0, 0, 0))))
        d.ellipse((600, 380, 622, 402), fill=(255, 0, 0))     # aviso do SRP ja aberto
        self.assertIn("aviso aberto", core.problema_tela(janela, tela))

    def test_coluna_do_foco(self):
        self.assertEqual(core.qual_coluna((424, 503, 466, 514), CENTROS), "custo")
        self.assertEqual(core.qual_coluna((475, 503, 499, 514), CENTROS), "conta")
        self.assertEqual(core.qual_coluna((656, 503, 740, 514), CENTROS), "quantidade")
        self.assertIsNone(core.qual_coluna((1200, 503, 1250, 514), CENTROS))

    def test_conta_pelo_vao(self):
        custo = (424, 503, 466, 514)
        self.assertTrue(core.conta_editavel(custo, (475, 503, 499, 514)))
        self.assertFalse(core.conta_editavel(custo, (656, 503, 740, 514)))


class PintarLinhas(unittest.TestCase):
    def planilha(self, pasta):
        caminho = pathlib.Path(pasta) / "notas.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Notas"
        ws.append(["CHAVE NF", "ITEM", "QTD FISCAL"])
        for i in range(4):
            ws.append(["1" * 44, "900135", 100 + i])
        wb.create_sheet("Outra")
        wb.active = 0
        wb.save(caminho)
        return caminho

    def test_pinta_so_as_linhas_pedidas_nas_colunas_do_cabecalho(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            self.assertEqual(core.pintar_linhas(caminho, [3, 2, 3]), "arquivo")
            ws = load_workbook(caminho).active
            verde = lambda c: (c.fill.fgColor.rgb or "")[-6:] == core.VERDE_CONCLUIDA  # noqa: E731
            self.assertTrue(all(verde(ws.cell(row=l, column=c)) for l in (2, 3) for c in (1, 2, 3)))
            self.assertFalse(any(verde(ws.cell(row=l, column=c)) for l in (1, 4, 5) for c in (1, 2, 3)))
            self.assertFalse(verde(ws.cell(row=2, column=4)))  # fora do cabecalho
            self.assertEqual((ws.title, ws.cell(row=3, column=3).value), ("Notas", 101))  # valores e aba ativa ficam
            self.assertEqual(sorted(p.name for p in pathlib.Path(pasta).iterdir()), ["notas.xlsx"])  # sem temporario

    def test_planilha_travada_pinta_no_excel_aberto(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            chamadas, original_replace, original_excel = [], core.os.replace, core._pintar_no_excel

            def travado(*_a):
                raise PermissionError("aberta no Excel")
            core.os.replace, core._pintar_no_excel = travado, lambda *a: chamadas.append(a)
            try:
                self.assertEqual(core.pintar_linhas(caminho, [4]), "excel")
            finally:
                core.os.replace, core._pintar_no_excel = original_replace, original_excel
            self.assertEqual(chamadas, [(caminho, "Notas", [4], 3, core.VERDE_CONCLUIDA, {}, 1)])
            self.assertEqual(sorted(p.name for p in pathlib.Path(pasta).iterdir()), ["notas.xlsx"])

    def test_serie_com_ponto_outro_verde_e_comentario_na_chave(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            self.assertEqual(core.pintar_linhas(caminho, [2, 3], core.VERDE_SERIE_PONTO, {2: "lançada com '.'"}), "arquivo")
            ws = load_workbook(caminho).active
            self.assertEqual(ws.cell(row=3, column=2).fill.fgColor.rgb[-6:], core.VERDE_SERIE_PONTO)
            self.assertTrue(core.fundo_verde(ws.cell(row=2, column=1).fill))  # ainda conta como feita ao recarregar
            self.assertEqual(ws.cell(row=2, column=1).comment.text, "lançada com '.'")  # na coluna CHAVE NF
            self.assertIsNone(ws.cell(row=3, column=1).comment)


class PlanilhaSolta(unittest.TestCase):
    def test_ler_nao_deixa_o_arquivo_travado(self):
        # o Excel salva por cima trocando o arquivo: com o arquivo ainda aberto pelo robo, o salvar falhava (06/11/2026)
        import snac0236
        with tempfile.TemporaryDirectory() as pasta:
            for ler in (core.ler_planilha, snac0236.ler_planilha):
                caminho = pathlib.Path(pasta) / "p.xlsx"
                wb = Workbook()
                wb.active.append(["Item", "C/Custo", "Conta", "Quantidade"])
                wb.save(caminho)
                try:
                    ler(caminho)
                except ValueError:
                    pass  # cabecalho da outra tela: o que importa e o arquivo solto depois
                caminho.rename(caminho.with_name("trocado.xlsx"))  # falha no Windows se o arquivo ficou aberto
                caminho.with_name("trocado.xlsx").unlink()


class FecharNoExcel(unittest.TestCase):
    class Livro:
        def __init__(self, saved, read_only=False):
            self.Saved, self.ReadOnly, self.fechada = saved, read_only, None

        def Close(self, SaveChanges):
            self.fechada = SaveChanges

    def fechar(self, livro):
        original = core._pasta_aberta
        core._pasta_aberta = lambda caminho: livro
        try:
            return core.fechar_no_excel("planilha.xlsx")
        finally:
            core._pasta_aberta = original

    def test_nao_aberta_nao_faz_nada(self):
        self.assertFalse(self.fechar(None))

    def test_salva_so_se_tiver_alteracao(self):
        alterada, salva = self.Livro(saved=False), self.Livro(saved=True)
        self.assertTrue(self.fechar(alterada))
        self.assertTrue(self.fechar(salva))
        self.assertEqual((alterada.fechada, salva.fechada), (True, False))

    def test_somente_leitura_com_alteracao_nao_fecha(self):
        livro = self.Livro(saved=False, read_only=True)
        with self.assertRaises(RuntimeError):
            self.fechar(livro)
        self.assertIsNone(livro.fechada)


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

    def test_compacta_logs_antigos_e_prints_por_7_dias(self):
        """Logs de execucao com mais de 2 dias vao para o zip do mes (e continuam legiveis); prints de telas ficam 7
        dias; zip sem log novo ha 30 dias e apagado."""
        import os
        import time
        import zipfile
        from unittest import mock
        pasta = pathlib.Path(tempfile.mkdtemp())
        (pasta / "telas").mkdir()
        def criar(nome, dias, texto="x"):
            arquivo = pasta / nome
            arquivo.write_text(texto, encoding="utf-8")
            quando = time.time() - dias * 86400
            os.utime(arquivo, (quando, quando))
            return arquivo
        criar("20261101-100000-snac0236.log", 3, "OK Nota ...000001 (linha 2): integrada com sucesso.\n")
        criar("20261102-090000.log", 5)
        criar("20261109-100000-snac0236.log", 0)  # recente: fica solto
        criar("atualizacao.log", 5)  # nao e log de execucao: nao vai para o zip
        criar("telas/20261101-100000-icms.png", 8)
        criar("telas/20261107-100000-icms.png", 2)
        criar("arquivo-202609.zip", 31)
        with mock.patch.object(core, "pasta_logs", lambda: pasta):
            self.assertEqual(core.limpar_logs(30), 2)  # o print de 8 dias e o zip velho
            self.assertEqual(sorted(p.name for p in pasta.rglob("*.*")),
                             ["20261107-100000-icms.png", "20261109-100000-snac0236.log", "arquivo-202611.zip",
                              "atualizacao.log"])
            with zipfile.ZipFile(pasta / "arquivo-202611.zip") as zipado:
                self.assertEqual(sorted(zipado.namelist()), ["20261101-100000-snac0236.log", "20261102-090000.log"])
            self.assertEqual([n for n, _ in core.ler_logs("*-snac0236.log")],
                             ["20261101-100000-snac0236.log", "20261109-100000-snac0236.log"])
            # um log do mesmo mes entra no mesmo zip
            criar("20261103-080000.log", 3)
            self.assertEqual(core.compactar_logs(), 1)
            with zipfile.ZipFile(pasta / "arquivo-202611.zip") as zipado:
                self.assertEqual(len(zipado.namelist()), 3)


if __name__ == "__main__":
    unittest.main()
