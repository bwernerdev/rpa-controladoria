"""SNAC0236: planilha (formatos, agrupamento por chave) e preenchimento com um SRP simulado (Enter entre campos)."""
import contextlib
import pathlib
import tempfile
import threading
import unittest
from datetime import datetime
from unittest import mock

from PIL import Image

import simulador  # noqa: F401  (poe src no sys.path)
import core
import snac0236

CHAVE_A = "1" * 40 + "0001"  # chaves ficticias
CHAVE_B = "2" * 40 + "0002"

# palavras da tela (texto, esq, topo, larg, alt), como o palavras_tela devolve
PALAVRAS = [
    ("Chave", 10, 100, 40, 12), ("Acesso", 55, 100, 45, 12), ("NF-e", 105, 100, 25, 12),
    ("Filial", 10, 130, 35, 12), ("Física:", 50, 130, 40, 12), ("Filial", 500, 130, 35, 12), ("Pagamento:", 540, 130, 70, 12),
    ("Data", 10, 160, 30, 12), ("Emissão:", 45, 160, 55, 12), ("Entrada:", 200, 160, 50, 12),
    ("Valor", 400, 190, 35, 12), ("Total", 440, 190, 35, 12), ("Documento:", 480, 190, 70, 12),
    ("Frete:", 10, 220, 40, 12),
    ("Integrar", 500, 300, 50, 12), ("Coleta", 600, 300, 40, 12), ("Frete", 645, 300, 35, 12),
    ("Item", 10, 370, 30, 12), ("Documento", 45, 370, 70, 12), ("Recebimento", 120, 370, 90, 12),
    ("tem", 240, 400, 30, 12), ("Descrição", 400, 400, 60, 12), ("Spm", 900, 400, 30, 12), ("Nat", 990, 385, 30, 12),
    ("Oper", 990, 400, 30, 12), ("(ltd.", 1050, 400, 40, 12), ("Fiscal", 1095, 400, 40, 12), ("Unitario", 1180, 400, 60, 12),
    ("Fiscal", 1245, 400, 40, 12),
]
CAIXAS = {"chave": (140, 100, 400, 112), "filial1": (100, 130, 130, 142), "filial2": (140, 130, 170, 142),
          "data": (110, 160, 170, 172), "valor_total": (560, 190, 650, 202), "frete": (60, 220, 110, 232)}
INTEGRAR = (500, 300, 550, 312)


class SRP0236:
    """Formulario como lista de campos; Enter avanca um. Campos sem nome sao os que o robo pula com Enter."""

    def __init__(self, recusar=(), integrar_volta_para_chave=True, extras_filial=0, sem_cor=()):
        self.sem_cor = set(sem_cor)  # indices do cabecalho cujo foco nao fica ciano (ex.: a lista Lemg)
        self.extras_filial = extras_filial  # campos a mais entre a 2a parte da filial e a Data (ex.: Filial Fiscal editavel)
        self.recusar = set(recusar)  # nomes de campo em que o Enter nao avanca (aviso)
        self.volta = integrar_volta_para_chave
        self.lock = threading.Lock()
        self.indice, self.linha_grade, self.alerta = 0, 0, False
        self.atual, self.integradas = {}, []

    def campos(self):
        cab = (["chave", None, None, None, None, "filial1", "filial2", None] + [None] * self.extras_filial + ["data"] + [None] * 6
               + ["valor_total", "frete"] + [None] * 9)
        rateio = ["oper", "natoper", "qtd", None, "unit", "icms"] + [None] * 6 + ["icms_terceiro"]
        return cab, rateio

    def onde(self):
        """(nome do campo ou None, caixa)."""
        cab, rateio = self.campos()
        if self.indice < len(cab):
            nome = cab[self.indice]
            return nome, CAIXAS.get(nome) or (700, 20 + self.indice * 15, 760, 30 + self.indice * 15)
        resto = self.indice - len(cab)
        ciclo = 1 + len(rateio)
        k, j = divmod(resto, ciclo)
        if j == 0:
            return f"item{k}", (235, 430 + 20 * k, 290, 442 + 20 * k)
        colunas = {"oper": 885, "natoper": 975, "qtd": 1070, "unit": 1200}  # x da coluna na grade (linha k)
        nome = rateio[j - 1]
        if nome in colunas:
            x = colunas[nome]
            return f"{nome}{k}", (x, 430 + 20 * k, x + 60, 442 + 20 * k)
        return (f"{nome}{k}" if nome else None), (1400, 500 + 15 * j, 1460, 510 + 15 * j)

    def titulo(self):
        """Titulo da coluna acima do campo em foco (o robo le antes de digitar Valor Unit e ICMS)."""
        nome = (self.onde()[0] or "").rstrip("0123456789")
        return [{"unit": ("Unitario Fiscal", 0, 0, 50, 10), "icms": ("Vlr ICMS Item Doc Forn", 0, 0, 50, 10)}.get(
            nome, ("Outra", 0, 0, 30, 10))]

    def caixa(self, _janela=None):
        return None if self.indice in self.sem_cor else self.onde()[1]

    def tela(self, _janela=None, _fixos=()):
        return self.caixa(), self.alerta

    def teclas(self, texto, pause=0):
        with self.lock:
            nome = self.onde()[0]
            base = nome.rstrip("0123456789") if nome else None
            if texto == "{ENTER}":
                if base in self.recusar:
                    self.alerta = True
                    return
                self.indice += 1
            elif nome:
                self.atual[nome] = self.atual.get(nome, "") + texto.replace("{,}", ",")

    def clicar(self, x, y):
        with self.lock:
            if INTEGRAR[0] <= x <= INTEGRAR[2] and INTEGRAR[1] <= y <= INTEGRAR[3]:
                self.integradas.append(self.atual)
                self.atual = {}
                if self.volta:
                    self.indice, self.alerta = 0, False


@contextlib.contextmanager
def ambiente(srp, palavras=PALAVRAS):
    pasta = pathlib.Path(tempfile.mkdtemp())
    trocas = {
        "send_keys": srp.teclas, "caixa_foco": srp.caixa, "estado_tela": srp.tela,
        "capturar": lambda janela: Image.new("RGB", (1600, 900), "white"),
        # tela inteira: as palavras; recorte acima do foco: o titulo da coluna; recorte do Integrar: o botao
        "palavras_tela": lambda imagem, t: (list(palavras) if imagem.size == (1600, 900) else
                                            srp.titulo() if imagem.height < 60 else [("Integrar", 300, 80, 50, 12)]),
        "caminho_tesseract": lambda: pathlib.Path("tesseract-falso"),
        "titulo_janela_ativa": lambda: "Corporativo - Desktop Viewer", "mouse_no_canto": lambda: False,
        "janela_citrix": lambda trecho: simulador.Janela(), "problema_tela": lambda retangulo, imagem: None,
        "pasta_logs": lambda: pasta, "caixas_cor": lambda imagem, cor: [],
    }
    with contextlib.ExitStack() as pilha:
        for nome, valor in trocas.items():
            pilha.enter_context(mock.patch.object(core, nome, valor))
        pilha.enter_context(mock.patch.object(snac0236, "clicar", srp.clicar))
        pilha.enter_context(mock.patch.object(snac0236, "ler_estado", srp.tela))
        pilha.enter_context(mock.patch.object(snac0236, "foco_tela", srp.caixa))
        yield pasta


def nota(chave, itens, filial=("30", "46"), data="01/11/2026", total="47407.28", frete="CIF"):
    regs = [{"linha": 2 + i, "chave": chave, "filial1": filial[0], "filial2": filial[1], "data": data, "valor_total": total,
             "frete": frete, **item} for i, item in enumerate(itens)]
    notas, erros = snac0236.agrupar_notas(regs)
    assert not erros, erros
    return notas[0]


def item(codigo, oper="5101", qtd="2", unit="10.5", icms="1.89", icms_t="21"):
    return {"item": codigo, "oper": oper, "qtd": qtd, "unit": unit, "icms": icms, "icms_terceiro": icms_t}


def rodar(srp, notas, com_continuar=True, ao_pausar=None, intervalo=0):
    avisos, estados = [], []
    retomar = threading.Event() if com_continuar else None

    def estado(linha, nome):
        estados.append((linha, nome))
        if nome == "pausado" and com_continuar:
            def soltar():
                if ao_pausar:
                    ao_pausar(srp)
                retomar.set()
            threading.Timer(0.05, soltar).start()

    with mock.patch.object(snac0236.time, "sleep", lambda s: None if s >= 0.1 else threading.Event().wait(s)):
        codigo = snac0236.executar(notas, "Desktop Viewer", 0, 0.3, 0, threading.Event(), avisos.append, estado, retomar, intervalo)
    return codigo, avisos, estados


class Planilha(unittest.TestCase):
    def gravar(self, linhas):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["CHAVE NF", "FILIAL FISCAL", "DATA EMISSÃO", "VALOR TOTAL DOCUMENTO", "FRETE", "ITEM", "OPER SPM",
                   "QTD FISCAL", "VALOR UNIT", "VALOR DO ICMS", "ICMS TERCEIRO (BASE DE CALCULO NF)"])
        for linha in linhas:
            ws.append(linha)
        caminho = pathlib.Path(tempfile.mkdtemp()) / "nota.xlsx"
        wb.save(caminho)
        return caminho

    def test_formatos_e_agrupamento(self):
        caminho = self.gravar([
            [CHAVE_A, "30 46", datetime(2026, 11, 1), 47407.28, "cif", 900135, "5101", 2, 10.5, 1.89, 21],
            [CHAVE_A, None, None, None, None, "900266", 5101, 1.25, "3,75", None, None],
            [CHAVE_B, "30 47", "02/11/2026", "1200", "FOB", 900135, 5102, 1, 1200, 0, 0],
        ])
        registros, notas, erros = snac0236.carregar(caminho)
        self.assertEqual(erros, [])
        self.assertEqual([len(n["itens"]) for n in notas], [2, 1])
        a = notas[0]
        self.assertEqual((a["filial1"], a["filial2"], a["data"], a["valor_total"], a["frete"]), ("30", "46", "01/11/2026", "47407.28", "CIF"))
        self.assertEqual((a["itens"][1]["qtd"], a["itens"][1]["unit"], a["itens"][1]["icms"]), ("1.25", "3.75", ""))
        self.assertEqual(notas[1]["frete"], "FOB")

    def test_erros(self):
        caminho = self.gravar([
            [1.11e43, "3046", "31/02/2026", "47.407,28", 150, "900135", None, 1, 1, 0, 0],
            [CHAVE_A, "30 46", "01/11/2026", 10, None, "900135", "5101", 1, 1, 0, 0],
            [CHAVE_A, "30 47", "01/11/2026", 10, None, "900266", "5101", 1, 1, 0, 0],
            [CHAVE_B, "30 46", "01/11/2026", 10, None, "900135", "5101", 1, 1, 0, 0],
            [CHAVE_A, "30 46", "01/11/2026", 10, None, "900135", "5101", 1, 1, 0, 0],
        ])
        _, _, erros = snac0236.carregar(caminho)
        texto = "\n".join(erros)
        for trecho in ("linha 2: a CHAVE NF esta como numero", "linha 2: FILIAL FISCAL precisa de 2 partes", "linha 2: data invalida",
                       "linha 2: VALOR TOTAL DOCUMENTO: valor invalido", "linha 4: FILIAL FISCAL diferente", "linha 2: FRETE invalido: 150", "linha 3: FRETE em branco",
                       "linha 6: a CHAVE NF ...110001 ja apareceu antes", "OPER SPM em branco"):
            self.assertIn(trecho, texto)

    def test_coluna_faltando(self):
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.append(["CHAVE NF", "ITEM"])
        caminho = pathlib.Path(tempfile.mkdtemp()) / "x.xlsx"
        wb.save(caminho)
        with self.assertRaises(ValueError):
            snac0236.ler_planilha(caminho)


class PassoAPasso(unittest.TestCase):
    def test_anuncia_e_espera_cada_acao(self):
        srp = SRP0236()
        inicio = __import__("time").time()
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])], intervalo=0.01)
        self.assertEqual(codigo, 0, avisos)
        passos = [a for a in avisos if a.startswith("Passo a passo:")]
        enters = sum(n for _, n in snac0236.SEQ_CABECALHO + snac0236.SEQ_ITEM)
        self.assertEqual(len(passos), 12 + enters + 1)  # 12 campos, cada Enter e o Integrar
        self.assertIn("Passo a passo: digitar '30' em Filial Fisica (1a parte)", passos)
        self.assertEqual(passos[-1], "Passo a passo: clicar em Integrar")
        self.assertGreater(__import__("time").time() - inicio, 0.01 * len(passos) * 0.9)


class Tela(unittest.TestCase):
    def test_foco_unido(self):
        imagem = Image.new("RGB", (400, 100), "white")
        for x0, x1 in ((50, 60), (140, 160)):  # sobras do ciano a esquerda e a direita do texto selecionado
            imagem.paste((0, 255, 255), (x0, 20, x1, 40))
        imagem.paste((0, 255, 255), (300, 2, 330, 4))  # traco fino no alto da tela: nao e foco
        caixa = snac0236.foco_unido(imagem)
        self.assertLessEqual(caixa[0], 52)
        self.assertGreaterEqual(caixa[2], 157)
        self.assertIsNone(snac0236.foco_unido(Image.new("RGB", (50, 50), "white")))
        selecionado = Image.new("RGB", (400, 100), "white")  # campo branco com o texto selecionado (Data preenchida)
        selecionado.paste((72, 100, 141), (60, 20, 160, 40))
        selecionado.paste((72, 100, 141), (0, 0, 1, 100))  # borda de 1 px da janela, mesmo azul: nao e foco
        caixa = snac0236.foco_unido(selecionado)
        self.assertTrue(caixa and caixa[0] <= 64 and caixa[2] >= 156, caixa)

    def test_rotulo_mais_perto_a_esquerda(self):
        rotulos, lista = snac0236.localizar_rotulos(PALAVRAS), snac0236.frases(PALAVRAS)
        longe = (150, 160, 190, 172)  # Data com o comeco um pouco afastado do rotulo
        self.assertTrue(snac0236.posicao_ok("data", longe, rotulos, None, lista))
        self.assertFalse(snac0236.posicao_ok("data", (260, 160, 300, 172), rotulos, None, lista))  # campo Entrada
        self.assertFalse(snac0236.posicao_ok("valor_total", (700, 185, 760, 195), rotulos, None, lista))  # longe demais
        self.assertFalse(snac0236.posicao_ok("data", None, rotulos, None, lista))

    def test_rotulos(self):
        rotulos = snac0236.localizar_rotulos(PALAVRAS)
        self.assertEqual(set(rotulos), {"chave", "filial", "data", "valor_total", "frete", "item", "integrar", "oper", "qtd", "unit"})
        self.assertEqual(rotulos["item"][:2], (240, 400))  # cabecalho da coluna, nao o titulo "Item Documento Recebimento"
        self.assertEqual(rotulos["frete"][:2], (10, 220))  # rotulo, nao o botao "Coleta Frete"
        self.assertTrue(snac0236.posicao_ok("chave", CAIXAS["chave"], rotulos))
        self.assertFalse(snac0236.posicao_ok("chave", CAIXAS["data"], rotulos))
        self.assertTrue(snac0236.posicao_ok("filial2", CAIXAS["filial2"], rotulos, CAIXAS["filial1"]))
        self.assertIsNone(snac0236.posicao_ok("data", CAIXAS["data"], {}))
        lista = snac0236.frases(PALAVRAS)
        self.assertEqual(snac0236.onde_esta(CAIXAS["frete"], lista), "frete")
        self.assertEqual(snac0236.onde_esta((235, 430, 290, 442), lista), "tem")


class Fluxo(unittest.TestCase):
    def test_duas_notas(self):
        srp = SRP0236()
        notas = [nota(CHAVE_A, [item("900135"), item("900266", qtd="1.25")]), nota(CHAVE_B, [item("900135")])]
        notas[1]["itens"][0]["linha"] = 4
        with ambiente(srp):
            codigo, avisos, estados = rodar(srp, notas)
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(len(srp.integradas), 2)
        a = srp.integradas[0]
        self.assertEqual((a["chave"], a["filial1"], a["filial2"], a["data"], a["valor_total"], a["frete"]),
                         (CHAVE_A, "30", "46", "01/11/2026", "47407.28", "CIF"))
        self.assertEqual((a["item0"], a["oper0"], a["qtd0"], a["unit0"], a["icms0"], a["icms_terceiro0"]),
                         ("900135", "5101", "2", "10.5", "1.89", "21"))
        self.assertEqual((a["item1"], a["qtd1"]), ("900266", "1.25"))
        self.assertEqual(srp.integradas[1]["chave"], CHAVE_B)
        self.assertEqual(estados[-1], (4, "ok"))
        self.assertTrue(any("PAUSADO: confira a integracao" in a for a in avisos))

    def test_enter_a_mais_ate_o_campo_certo(self):
        srp = SRP0236(extras_filial=2)  # o SRP pede 2 Enter a mais antes da Data
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(srp.integradas[0]["data"], "01/11/2026")
        self.assertEqual(srp.integradas[0]["valor_total"], "47407.28")

    def test_nao_chega_no_campo_pausa_sem_digitar(self):
        srp = SRP0236(extras_filial=6)  # alem do limite (previstos + 4)
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])], com_continuar=False)
        self.assertEqual(codigo, 1)
        self.assertTrue(any("nao cheguei no campo Data Emissao com 6 Enter" in a for a in avisos), avisos)
        self.assertNotIn("data", srp.atual)

    def test_passa_por_campo_sem_cor(self):
        srp = SRP0236(sem_cor={10, 11})  # dois campos sem ciano entre a Data e o Valor Total (lista Lemg...)
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(srp.integradas[0]["valor_total"], "47407.28")

    def test_cursor_fora_da_chave_nao_digita(self):
        srp = SRP0236()
        srp.indice = 8  # Data Emissao
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 1)
        self.assertIn("Nada foi digitado", avisos[-1])
        self.assertEqual(srp.atual, {})

    def test_aviso_pausa_e_usuario_termina(self):
        srp = SRP0236(recusar={"data"})

        def usuario(srp):  # termina a nota a mao e volta o cursor para a Chave
            srp.atual, srp.indice, srp.alerta = {}, 0, False
            srp.recusar.clear()

        notas = [nota(CHAVE_A, [item("900135")]), nota(CHAVE_B, [item("900266")])]
        notas[1]["itens"][0]["linha"] = 3
        with ambiente(srp):
            codigo, avisos, estados = rodar(srp, notas, ao_pausar=usuario)
        self.assertEqual(codigo, 0, avisos)
        self.assertTrue(any("apareceu um aviso do SRP" in a and "Data Emissao" in a for a in avisos))
        self.assertIn((2, "pausado"), estados)
        self.assertEqual([n["chave"] for n in srp.integradas], [CHAVE_B])

    def test_sem_continuar_interrompe(self):
        srp = SRP0236(recusar={"qtd"})
        with ambiente(srp):
            codigo, avisos, estados = rodar(srp, [nota(CHAVE_A, [item("900135")])], com_continuar=False)
        self.assertEqual(codigo, 1)
        self.assertIn("Interrompido", avisos[-1])
        self.assertIn((2, "erro"), estados)

    def test_sem_rotulos_so_conta_enters(self):
        srp = SRP0236()
        with ambiente(srp, palavras=[("Integrar", 500, 300, 50, 12)]):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertTrue(any("nao li na tela os rotulos" in a for a in avisos))
        self.assertEqual(srp.integradas[0]["icms_terceiro0"], "21")


if __name__ == "__main__":
    unittest.main()
