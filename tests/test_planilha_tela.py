"""Leitura/validacao da planilha e regras de tela (campo ciano, coluna do foco, Conta pelo vao)."""
import pathlib
import tempfile
import unittest
from datetime import datetime

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
        self.assertTrue(any("quantidade inválida" in e for e in erros))

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
        esq, topo, dir_, _base = core.maior_caixa_ciano(imagem)
        self.assertTrue(600 <= esq <= 604 and 706 <= dir_ <= 711 and 200 <= topo <= 204)
        self.assertIsNone(core.maior_caixa_ciano(Image.new("RGB", (50, 50), (255, 255, 255))))

    def test_icone_do_aviso_ignora_outros_vermelhos(self):
        imagem = Image.new("RGB", (900, 500), (240, 240, 240))
        d = ImageDraw.Draw(imagem)
        d.rectangle((20, 400, 600, 440), fill=(255, 0, 0))   # faixa vermelha grande (nao e o icone)
        d.rectangle((700, 20, 703, 300), fill=(255, 0, 0))   # traco vermelho fino
        self.assertIsNone(core.icone_aviso(imagem))
        d.ellipse((300, 100, 322, 122), fill=(255, 0, 0))    # icone ~22 px
        esq, topo, dir_, _base = core.icone_aviso(imagem)
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
            self.assertTrue(all(verde(ws.cell(row=n, column=c)) for n in (2, 3) for c in (1, 2, 3)))
            self.assertFalse(any(verde(ws.cell(row=n, column=c)) for n in (1, 4, 5) for c in (1, 2, 3)))
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
            self.assertEqual(chamadas, [(caminho, "Notas", [4], 3, core.VERDE_CONCLUIDA, {}, 1, None, None)])
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


class AbaLancados(unittest.TestCase):
    """SNAC0236 (usuario, 20/11/2026): as linhas das notas integradas vao tambem para a aba "Lançados", com a data."""
    A, B, C = "1" * 44, "2" * 44, "3" * 44
    QUANDO = datetime(2026, 11, 20, 14, 30)

    def planilha(self, pasta):
        caminho = pathlib.Path(pasta) / "notas.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Notas"
        ws.append(["CHAVE NF", "ITEM", "DATA EMISSÃO", "QTD FISCAL"])
        ws.append([self.A, "900135", datetime(2026, 11, 1), 2])
        ws.append([None, "900136", None, 3])  # chave em branco: mais um item da nota A
        ws.append([self.B, "900137", datetime(2026, 11, 2), 4])
        ws.append([self.C, "900138", datetime(2026, 11, 3), 5])
        ws["C2"].number_format = ws["C4"].number_format = "dd/mm/yyyy"
        ws.column_dimensions["A"].width = 50
        wb.save(caminho)
        return caminho

    def ler(self, caminho):
        from openpyxl import load_workbook
        return load_workbook(caminho)

    def test_cria_a_aba_ao_iniciar_so_uma_vez(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            self.assertEqual(core.pintar_linhas(caminho, [], lancados={}), "arquivo")
            livro = self.ler(caminho)
            self.assertEqual(livro.sheetnames, ["Notas", "Lançados"])
            self.assertEqual(livro.active.title, "Notas")  # o robo continua lendo a aba dos dados
            lancados = livro["Lançados"]
            self.assertEqual([c.value for c in lancados[1]], ["CHAVE NF", "ITEM", "DATA EMISSÃO", "QTD FISCAL", "LANÇADO EM"])
            self.assertEqual((lancados.max_row, lancados.column_dimensions["A"].width), (1, 50))
            self.assertIsNone(core.pintar_linhas(caminho, [], lancados={}))  # ja existe: nem regrava
            self.assertIsNone(core.pintar_linhas(caminho, []))  # sem pedido de Lançados: nada a fazer

    def test_copia_as_notas_integradas_verdes_com_a_data(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            core.pintar_linhas(caminho, [2, 3], lancados=dict.fromkeys([2, 3], self.QUANDO))
            core.pintar_linhas(caminho, [5], core.VERDE_SERIE_PONTO, {5: "lançada com '.'"},
                               lancados={5: datetime(2026, 11, 20, 15, 0)})
            livro = self.ler(caminho)
            notas, lancados = livro["Notas"], livro["Lançados"]
            self.assertEqual([[c.value for c in linha] for linha in lancados.iter_rows(min_row=2)], [
                [self.A, "900135", datetime(2026, 11, 1), 2, self.QUANDO],
                [None, "900136", None, 3, self.QUANDO],
                [self.C, "900138", datetime(2026, 11, 3), 5, datetime(2026, 11, 20, 15, 0)]])
            self.assertEqual(lancados["C2"].number_format, "dd/mm/yyyy")  # formato da celula vai junto
            self.assertEqual(lancados["E2"].number_format, core.FORMATO_LANCADO_EM)
            self.assertTrue(all(core.fundo_verde(lancados.cell(row=n, column=c).fill) for n in (2, 3, 4) for c in (1, 2, 3, 4)))
            self.assertEqual(lancados["B4"].fill.fgColor.rgb[-6:], core.VERDE_SERIE_PONTO)
            self.assertEqual(lancados["A4"].comment.text, "lançada com '.'")
            # a aba dos dados continua com todas as linhas, verdes (usuario: "mantenha a cor verde")
            self.assertEqual(notas.max_row, 5)
            self.assertTrue(all(core.fundo_verde(notas.cell(row=n, column=1).fill) for n in (2, 3, 5)))
            self.assertFalse(core.fundo_verde(notas.cell(row=4, column=1).fill))
            self.assertEqual(livro.active.title, "Notas")

    def test_nota_ja_copiada_nao_repete(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            core.pintar_linhas(caminho, [2, 3], lancados=dict.fromkeys([2, 3], self.QUANDO))
            core.pintar_linhas(caminho, [2, 3, 4], lancados=dict.fromkeys([2, 3, 4], self.QUANDO))  # Conferir pinta de novo
            lancados = self.ler(caminho)["Lançados"]
            self.assertEqual([lancados.cell(row=n, column=2).value for n in range(2, lancados.max_row + 1)],
                             ["900135", "900136", "900137"])

    def test_aba_lancados_com_outras_colunas_so_avisa(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            livro = self.ler(caminho)
            livro.create_sheet("Lançados").append(["OUTRA", "COISA"])
            livro.save(caminho)
            avisos = []
            self.assertEqual(core.pintar_linhas(caminho, [4], lancados={4: self.QUANDO}, avisos=avisos), "arquivo")
            livro = self.ler(caminho)
            self.assertEqual(len(avisos), 1)
            self.assertIn("colunas diferentes", avisos[0])
            self.assertEqual(livro["Lançados"].max_row, 1)  # nada copiado
            self.assertTrue(core.fundo_verde(livro["Notas"]["A4"].fill))  # mas pintou de verde

    def test_planilha_salva_com_a_aba_lancados_aberta(self):
        """Se o usuario salvar com a "Lançados" ativa, o robo continua pintando (e lendo) a aba dos dados."""
        with tempfile.TemporaryDirectory() as pasta:
            caminho = self.planilha(pasta)
            core.pintar_linhas(caminho, [2, 3], lancados=dict.fromkeys([2, 3], self.QUANDO))
            livro = self.ler(caminho)
            livro.active = 1
            livro.save(caminho)
            core.pintar_linhas(caminho, [4], lancados={4: self.QUANDO})
            livro = self.ler(caminho)
            self.assertTrue(core.fundo_verde(livro["Notas"]["A4"].fill))
            self.assertEqual(livro["Lançados"].max_row, 4)


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

    def test_gravar_json_troca_o_arquivo_de_uma_vez(self):
        """Grava pelo temporario + os.replace: o arquivo antigo so some quando o novo esta inteiro, sem sobra .tmp."""
        import json
        from unittest import mock
        pasta = pathlib.Path(tempfile.mkdtemp())
        arquivo = pasta / "config.json"
        core.gravar_json(arquivo, {"tema": "claro"})
        core.gravar_json(arquivo, {"tema": "escuro", "aba": "snac0236"}, indent=2)
        self.assertEqual(json.loads(arquivo.read_text(encoding="utf-8")), {"tema": "escuro", "aba": "snac0236"})
        with mock.patch.object(core.os, "replace", side_effect=OSError("disco cheio")), self.assertRaises(OSError):
            core.gravar_json(arquivo, {"tema": "auto"})  # falha no meio
        self.assertEqual(json.loads(arquivo.read_text(encoding="utf-8"))["tema"], "escuro")  # o antigo ficou inteiro
        self.assertEqual(sorted(p.name for p in pasta.iterdir() if p.suffix != ".tmp"), ["config.json"])


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


class LinhaDeComando(unittest.TestCase):
    """preencher.py: planilha que nao abre = mensagem curta e saida 1, sem o rastro do Python."""

    def rodar(self, planilha):
        from unittest import mock

        import preencher
        with mock.patch("sys.argv", ["preencher.py", str(planilha), "--dry-run"]), self.assertRaises(SystemExit) as saida:
            preencher.main()
        return str(saida.exception.code)

    def test_planilha_que_nao_existe(self):
        self.assertEqual(self.rodar("nao_existe.xlsx"), "Planilha não encontrada: nao_existe.xlsx")

    def test_arquivo_que_nao_e_planilha(self):
        falsa = pathlib.Path(tempfile.mkdtemp()) / "falsa.xlsx"
        falsa.write_text("nao e planilha", encoding="utf-8")
        self.assertTrue(self.rodar(falsa).startswith(f"Não consegui ler a planilha {falsa}: "))


if __name__ == "__main__":
    unittest.main()
