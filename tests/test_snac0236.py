"""SNAC0236: planilha (formatos, agrupamento por chave) e preenchimento com um SRP simulado (Enter entre campos)."""
import contextlib
import pathlib
import re
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
    # Rateio (grade de tributos e botao); no simulador a tela inteira sempre devolve estas palavras
    ("Ident", 800, 560, 40, 12), ("Tributo", 800, 575, 50, 12), ("Base", 990, 575, 35, 12), ("Calculo", 1030, 575, 50, 12),
    ("Aliquota", 1100, 575, 55, 12), ("Lancamento", 1190, 575, 75, 12), ("COFINS", 150, 600, 50, 12),
    ("ALIQUOTA", 205, 600, 60, 12), ("ZERO", 270, 600, 35, 12), ("ICMS", 150, 660, 35, 12), ("TERCEIROS", 190, 660, 75, 12),
    ("Tributos", 700, 800, 55, 12), ("OK", 760, 800, 20, 12),
]
TRIBUTOS_OK = (700, 800, 780, 812)
AVISO = [(p, 300 + 60 * i, 450, 55, 12) for i, p in enumerate(
    "Este documento possui frete com ICMS PAGO no LTCC, por isso foi alterado o Frete para CIF.".split())]
RESUMO = [(p, 300 + 60 * i, 480, 55, 12) for i, p in enumerate("ae dos Tributos Documento Confirmar Cancelar".split())]
LTCC = [(p, 300 + 60 * i, 450, 55, 12) for i, p in enumerate("Informe o Centro de Custo e Conta Contabil para Frete LTCC".split())] + [
    ("Frete", 1000, 470, 50, 12), ("LTCC", 1055, 470, 45, 12),  # cabecalho da grade (nao e o botao)
    ("Voltar", 430, 700, 90, 12), ("Frete", 1000, 700, 50, 12), ("LTC", 1055, 700, 45, 12)]
BOTAO_LTCC = (1000, 700, 1100, 712)
REFERENCIA = [(p, 300 + 70 * i, 450, 65, 12) for i, p in enumerate(
    "Nota Fiscal de Referência emitida pelo Faturamento".split())] + [
    ("Escolha", 300, 480, 60, 12), ("a", 365, 480, 10, 12), ("Nota", 380, 480, 35, 12), ("Fiscal", 420, 480, 45, 12),
    ("Faturamento", 470, 480, 90, 12),
    ("030", 400, 520, 30, 12), ("00765", 450, 520, 50, 12), ("0000087445", 520, 520, 90, 12), ("1", 630, 520, 10, 12),
    ("99999", 660, 520, 50, 12), ("1", 730, 520, 10, 12), ("OK", 700, 760, 20, 12)]
BOTAO_OK_REF = (700, 760, 720, 772)
SUCESSO = [(p, 300 + 80 * i, 450, 75, 12) for i, p in enumerate("Documento Integrado com Sucesso.".split())]
BARRA = (250, 400, 1300, 430)  # barra de titulo azul das janelas pos-Integrar
AREA = (230, 395, 1320, 900)  # snac0236.area_dialogo da BARRA na tela 1600x900
LINHAS_TRIBUTO = {"199": 600, "198": 630, "114": 660}  # y de cada linha da grade de tributos
COLUNAS_TRIBUTO = {"ident": 800, "base": 1000, "aliquota": 1100, "lancamento": 1200}
CAIXAS = {"chave": (140, 100, 400, 112), "filial1": (100, 130, 130, 142), "filial2": (140, 130, 170, 142),
          "data": (110, 160, 170, 172), "valor_total": (560, 190, 650, 202), "frete": (60, 220, 110, 232)}
INTEGRAR = (500, 300, 550, 312)


class SRP0236:
    """Formulario como lista de campos; Enter avanca um. Campos sem nome sao os que o robo pula com Enter."""

    def __init__(self, recusar=(), integrar_volta_para_chave=True, extras_filial=0, sem_cor=(), parado=(), extras_rateio=0,
                 colunas_tributo=("ident", "base", "aliquota")):  # como no SRP real: o Lancamento nao para
        self.colunas_tributo = colunas_tributo  # campos de cada linha de tributo em que o Enter para
        self.extras_rateio = extras_rateio  # Enter a mais (alem dos 7) ate o Rateio abrir
        self.parado = set(parado)  # nomes de campo em que o Enter nao faz nada (nem aviso)
        self.sem_cor = set(sem_cor)  # indices do cabecalho cujo foco nao fica ciano (ex.: a lista Lemg)
        self.extras_filial = extras_filial  # campos a mais entre a 2a parte da filial e a Data (ex.: Filial Fiscal editavel)
        self.recusar = set(recusar)  # nomes de campo em que o Enter nao avanca (aviso)
        self.volta = integrar_volta_para_chave
        self.lock = threading.Lock()
        self.indice, self.linha_grade, self.alerta = 0, 0, False
        self.atual, self.integradas, self.reintegrou = {}, [], 0
        self.com_aviso_frete, self.aviso_frete, self.fechou_aviso = False, False, 0
        self.resumo, self.confirmou = False, 0  # "Resumo dos Tributos do Documento", depois do aviso do frete
        self.ltcc, self.sucesso, self.sucessos = False, False, 0  # janela do Frete LTCC e "Integrado com Sucesso"
        self.sem_barra, self.recorte_vazio = False, False  # barra azul nao achada / recorte sem as palavras da janela
        self.leituras, self.recortes = 0, 0
        self.cliques_chave, self.avisos_no_1o, self.enter_no_ltcc = 0, False, 0
        self.ignora_cliques = False  # cliques no Item e no Ident Tributo que nao mudam o foco
        self.clique_item_passa = False  # o foco passa por outro campo antes de chegar no Item (SRP real, 04/11 19:20)
        self.cliques_item, self.cliques_ident, self.enters = 0, 0, 0
        # conteudo que o SRP ja traz no campo (Ident do ICMS TERCEIROS = 1, 05/11/2026). Chegando pelo Enter o texto vem
        # selecionado (digitar substitui); pelo clique o cursor fica no fim (digitar acrescenta).
        self.iniciais = {"t114ident": "1"}
        # Tributos OK deixa o foco no Item da mesma linha (SRP real, 05/11/2026); seta para baixo vai a linha de baixo
        self.ok_desce = False  # True: o Tributos OK ja desce para o Item da linha de baixo
        self.ignora_seta = False  # seta para baixo que nao faz nada
        self.setas = 0
        # janela "Nota fiscal de referencia do faturamento" depois do 1o Integrar (SRP real, 05/11/2026)
        self.com_referencia, self.referencia, self.sem_ok_ref = False, False, False  # sem_ok_ref: OCR nao le o OK
        self.ok_referencia, self.enters_referencia = 0, 0
        self.ref_no_caminho = None  # Enter depois do Frete em que abre a janela da NF de referencia (SRP real, 05/11/2026)
        self.selecionado = True

    def campos(self):
        cab = (["chave", None, None, None, None, "filial1", "filial2", None] + [None] * self.extras_filial + ["data"] + [None] * 6
               + ["valor_total", "frete"] + [None] * 9)
        tributos = [f"t{cod}{col}" for cod in LINHAS_TRIBUTO for col in self.colunas_tributo]
        rateio = ["oper", "natoper", "qtd", "qtdfisica", "unit", "icms"] + [None] * (6 + self.extras_rateio) + tributos
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
        if nome and nome.startswith("t1"):  # grade de tributos do Rateio
            x, y = COLUNAS_TRIBUTO[nome[4:]], LINHAS_TRIBUTO[nome[1:4]]
            return f"{nome}{k}", (x, y, x + 60, y + 12)
        if j > 5:  # depois do Valor Unit a grade rola: o foco fica parado na borda direita (como no SRP real em 4K)
            return (f"{nome}{k}" if nome else None), (1400, 500, 1460, 510)
        return (f"{nome}{k}" if nome else None), (1400, 500 + 15 * j, 1460, 510 + 15 * j)

    def titulo(self):
        """Titulo da coluna acima do campo em foco (o robo le antes de digitar Valor Unit e ICMS)."""
        nome = (self.onde()[0] or "").rstrip("0123456789")
        return [{"icms": ("Vir oms tem doc forn", 0, 0, 50, 10),
                 "t114ident": ("ICMS TERCEIROS", 0, 0, 60, 10)}.get(  # recorte do rotulo da linha (nota seguinte)
            nome, ("Outra", 0, 0, 30, 10))]

    def caixa(self, _janela=None):
        if self.referencia:  # foco na celula Filial da grade da janela
            return (450, 520, 500, 532)
        nome = self.onde()[0] or ""
        # Valor Unit: no SRP real fica amarelo (sem ciano)
        return None if self.indice in self.sem_cor or nome.startswith("unit") else self.onde()[1]

    def captura(self, _janela=None):
        """Tela muda a cada campo (no SRP real a grade rola ao entrar no Valor Unit)."""
        return Image.new("RGB", (1600, 900), ((self.indice * 97) % 256,) * 3)

    def tela(self, _janela=None, _fixos=(), _imagem=None):
        return self.caixa(), self.alerta

    def janela_por_cima(self):
        return self.referencia or self.aviso_frete or self.resumo or self.ltcc or self.sucesso

    def palavras_janelas(self):
        return (([p for p in REFERENCIA if not (self.sem_ok_ref and p[0] == "OK")] if self.referencia else [])
                + (AVISO if self.aviso_frete else []) + (RESUMO if self.resumo else []) + (LTCC if self.ltcc else [])
                + (SUCESSO if self.sucesso else []))

    def barras(self, imagem):
        self.leituras += 1
        return [BARRA] if self.janela_por_cima() and not self.sem_barra and imagem.size == (1600, 900) else []

    def recorte(self):
        """Palavras das janelas no recorte AREA (coordenadas do recorte)."""
        self.recortes += 1
        if self.recorte_vazio:
            return []
        return [(p, x - AREA[0], y - AREA[1], w, h) for p, x, y, w, h in self.palavras_janelas()]

    def teclas(self, texto, pause=0):
        with self.lock:
            nome = self.onde()[0]
            base = nome.rstrip("0123456789") if nome else None
            if texto == "{ENTER}" and self.referencia:  # no SRP real a janela fechou no 7o Enter
                self.enters_referencia += 1
                if self.enters_referencia >= 7:
                    self.referencia = False
                return
            if texto == "{ENTER}" and self.aviso_frete:  # OK no aviso do frete: vem o Resumo dos Tributos
                self.aviso_frete, self.fechou_aviso, self.resumo = False, self.fechou_aviso + 1, True
                return
            if texto == "{ENTER}" and self.resumo:  # Confirmar no Resumo: vem a janela do Frete LTCC
                self.resumo, self.confirmou, self.ltcc = False, self.confirmou + 1, True
                return
            if texto == "{ENTER}" and self.ltcc:  # no SRP real o Enter roda sem fim nessa janela
                self.enter_no_ltcc += 1
                return
            if texto == "{ENTER}" and self.sucesso:  # OK no "Integrado com Sucesso": tela vazia, foco no CNPJ
                self.sucesso, self.sucessos, self.indice = False, self.sucessos + 1, 1
                return
            if texto == "{DOWN}":
                self.setas += 1
                cab, rateio = self.campos()
                resto = self.indice - len(cab)
                if not self.ignora_seta and resto >= 0 and resto % (1 + len(rateio)) == 0:
                    self.indice += 1 + len(rateio)  # Item da linha de baixo
                return
            if texto == "{ENTER}":
                self.enters += 1
                if base in self.recusar:
                    self.alerta = True
                    return
                if base in self.parado:
                    return
                self.indice += 1
                self.selecionado = True
                cab = self.campos()[0]
                if self.ref_no_caminho and self.indice == cab.index("frete") + self.ref_no_caminho:
                    self.referencia = True
            elif nome:
                valor = "" if self.selecionado else self.atual.get(nome, self.iniciais.get(base, ""))
                for tecla in re.findall(r"\{[^}]*\}|[^{]+", texto):
                    if tecla == "{END}":
                        self.selecionado = False
                    elif tecla.startswith("{BACKSPACE"):
                        partes = tecla.strip("{}").split()
                        valor = valor[:max(len(valor) - (int(partes[1]) if len(partes) > 1 else 1), 0)]
                    else:
                        valor += "," if tecla == "{,}" else tecla
                self.selecionado = False
                self.atual[nome] = valor

    def rateio_aberto(self):
        return (self.onde()[0] or "").startswith("t1")

    def clicar(self, x, y):
        with self.lock:
            chave = CAIXAS["chave"]
            if chave[0] <= x <= chave[2] and chave[1] <= y <= chave[3]:  # clique no campo Chave Acesso NF-e
                self.indice, self.cliques_chave = 0, self.cliques_chave + 1
                return
            if self.referencia and BOTAO_OK_REF[0] <= x <= BOTAO_OK_REF[2] and BOTAO_OK_REF[1] <= y <= BOTAO_OK_REF[3]:
                self.referencia, self.ok_referencia = False, self.ok_referencia + 1
                return
            if self.ltcc and BOTAO_LTCC[0] <= x <= BOTAO_LTCC[2] and BOTAO_LTCC[1] <= y <= BOTAO_LTCC[3]:
                self.ltcc, self.sucesso = False, True
                return
            if TRIBUTOS_OK[0] <= x <= TRIBUTOS_OK[2] and TRIBUTOS_OK[1] <= y <= TRIBUTOS_OK[3] and self.rateio_aberto():
                cab, rateio = self.campos()  # fecha o Rateio: foco no Item da mesma linha da grade
                k = (self.indice - len(cab)) // (1 + len(rateio))
                self.indice = len(cab) + (k + int(self.ok_desce)) * (1 + len(rateio))
                return
            cab, rateio = self.campos()
            if self.indice < len(cab) and self.indice > cab.index("frete") and 235 <= x <= 290 and 430 <= y <= 442:
                self.cliques_item += 1  # 1a linha da grade, coluna Item
                if self.clique_item_passa:
                    self.indice = cab.index("frete") + 3

                    def chegar():
                        with self.lock:
                            self.indice = len(cab)
                    threading.Timer(0.1, chegar).start()
                elif not self.ignora_cliques:
                    self.indice = len(cab)
                return
            if self.rateio_aberto():
                for nome in rateio:
                    if nome and nome.startswith("t1"):
                        cx, cy = COLUNAS_TRIBUTO[nome[4:]], LINHAS_TRIBUTO[nome[1:4]]
                        if cx <= x <= cx + 60 and cy <= y <= cy + 12:
                            self.cliques_ident += 1
                            if not self.ignora_cliques:
                                self.selecionado = False
                                k = (self.indice - len(cab)) // (1 + len(rateio))
                                self.indice = len(cab) + k * (1 + len(rateio)) + 1 + rateio.index(nome)
                            return
            if INTEGRAR[0] <= x <= INTEGRAR[2] and INTEGRAR[1] <= y <= INTEGRAR[3]:
                if not self.atual:  # 2o clique (depois dos 7 Enter): conta, mas nao e outra nota
                    self.reintegrou += 1
                    self.aviso_frete = self.com_aviso_frete
                    self.indice, self.alerta = 0, False
                    return
                self.integradas.append(self.atual)
                self.referencia = self.com_referencia
                self.aviso_frete = self.avisos_no_1o  # janelas logo depois do 1o Integrar (sem o 2o)
                self.atual = {}
                if self.volta:
                    self.indice, self.alerta = 0, False


@contextlib.contextmanager
def ambiente(srp, palavras=PALAVRAS):
    pasta = pathlib.Path(tempfile.mkdtemp())
    trocas = {
        "send_keys": srp.teclas, "caixa_foco": srp.caixa, "estado_tela": srp.tela,
        "capturar": srp.captura,
        # tela inteira: as palavras; recorte acima do foco: o titulo da coluna; recorte do Integrar: o botao
        "palavras_tela": lambda imagem, t, escala=2: (list(palavras) + srp.palavras_janelas() if imagem.size == (1600, 900) else
                                            srp.recorte() if imagem.size == (AREA[2] - AREA[0], AREA[3] - AREA[1]) else
                                            srp.titulo() if imagem.height < 60 or imagem.width < 120 else
                                            [("Tributos", 300, 80, 55, 12), ("OK", 360, 80, 20, 12)] if srp.rateio_aberto() else
                                            [("Integrar", 300, 80, 50, 12)]),
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
        pilha.enter_context(mock.patch.object(snac0236, "barras_dialogo", srp.barras))
        pilha.enter_context(mock.patch.object(snac0236, "foco_tela", srp.caixa))
        pilha.enter_context(mock.patch.object(snac0236, "arquivo_posicoes", lambda: pasta / "posicoes_0236.json"))
        yield pasta


def nota(chave, itens, filial=("30", "46"), data="01/11/2026", total="47407.28", frete="CIF"):
    regs = [{"linha": 2 + i, "chave": chave, "filial1": filial[0], "filial2": filial[1], "data": data, "valor_total": total,
             "frete": frete, **item} for i, item in enumerate(itens)]
    notas, erros = snac0236.agrupar_notas(regs)
    assert not erros, erros
    return notas[0]


def item(codigo, oper="5101", qtd="2", unit="10.5", icms="1.89", icms_t="0", aliq="0"):
    return {"item": codigo, "oper": oper, "qtd": qtd, "unit": unit, "icms": icms, "icms_terceiro": icms_t, "aliquota_icms": aliq}


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
                   "QTD FISCAL", "VALOR UNIT", "VALOR DO ICMS", "BASE CALC. ICMS", "ALIQUOTA"])
        for linha in linhas:
            ws.append(linha)
        caminho = pathlib.Path(tempfile.mkdtemp()) / "nota.xlsx"
        wb.save(caminho)
        return caminho

    def test_formatos_e_agrupamento(self):
        caminho = self.gravar([
            [CHAVE_A, "30 46", datetime(2026, 11, 1), 47407.28, "cif", 900135, "5101", 2, 10.5, 180, 1500.5, 12],
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
        self.assertEqual(snac0236.valores_rateio(a["itens"][0]), {"ident": "1", "base": "1500.5", "aliquota": "12"})
        self.assertEqual(snac0236.valores_rateio(a["itens"][1]),
                         {"ident": "0", "base": "0", "aliquota": "0"})

    def test_chave_em_branco_e_item_da_nota_de_cima(self):
        caminho = self.gravar([
            [CHAVE_A, "30 46", datetime(2026, 11, 1), 47407.28, "CIF", 900135, 5101, 2, 10.5, 180, 1500.5, 12],
            [None, None, None, None, None, 900266, 5101, 1, 3.75, 0, 0, 0],
            [CHAVE_B, "30 47", "02/11/2026", 1200, "FOB", 900135, 5102, 1, 1200, 0, 0, 0],
            [None, None, None, None, None, 900266, 5102, 3, 1, 0, 0, 0],
        ])
        _, notas, erros = snac0236.carregar(caminho)
        self.assertEqual(erros, [])
        self.assertEqual([[i["item"] for i in n["itens"]] for n in notas], [["900135", "900266"], ["900135", "900266"]])
        self.assertEqual(snac0236.valores_rateio(notas[0]["itens"][1]), {"ident": "0", "base": "0", "aliquota": "0"})

    def test_chave_em_branco_com_cabecalho(self):
        caminho = self.gravar([
            [CHAVE_A, "30 46", "01/11/2026", 10, "CIF", 900135, 5101, 1, 1, 0, 0, 0],
            [None, "30 47", None, None, None, 900266, 5101, 1, 1, 0, 0, 0],
        ])
        _, _, erros = snac0236.carregar(caminho)
        self.assertTrue(any("linha 3: CHAVE NF em branco com o cabecalho preenchido" in e for e in erros), erros)

    def test_erros(self):
        caminho = self.gravar([
            [1.11e43, "3046", "31/02/2026", "47.407,28", 150, "900135", None, 1, 1, 0, 0],
            [CHAVE_A, "30 46", "01/11/2026", 10, None, "900135", "5101", 1, 1, 0, 5, 0],
            [CHAVE_A, "30 47", "01/11/2026", 10, None, "900266", "5101", 1, 1, 0, 0],
            [CHAVE_B, "30 46", "01/11/2026", 10, None, "900135", "5101", 1, 1, 0, 0],
            [CHAVE_A, "30 46", "01/11/2026", 10, None, "900135", "5101", 1, 1, 0, 0],
            [CHAVE_A, "30 46", "01/11/2026", 10, "CIF", "900135", "5101", 1, 1, 0, 1500, 0.12],
            [CHAVE_A, "30 46", "01/11/2026", 10, "CIF", "900135", "5101", 1, 1, 0, 1500, 120],
        ])
        _, _, erros = snac0236.carregar(caminho)
        texto = "\n".join(erros)
        for trecho in ("linha 2: a CHAVE NF esta como numero", "linha 2: FILIAL FISCAL precisa de 2 partes", "linha 2: data invalida",
                       "linha 2: VALOR TOTAL DOCUMENTO: valor invalido", "linha 4: FILIAL FISCAL diferente", "linha 2: FRETE invalido: 150", "linha 3: FRETE em branco",
                       "linha 6: a CHAVE NF ...110001 ja apareceu antes", "OPER SPM em branco",
                       "linha 3: BASE CALC. ICMS e ALIQUOTA precisam ser os dois 0",
                       "linha 7: ALIQUOTA 0.12 parece porcentagem do Excel", "linha 8: ALIQUOTA invalida: 120"):
            self.assertIn(trecho, texto)

    def gravar_cabecalho(self, base):
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.append(["CHAVE NF", "FILIAL FISCAL", "DATA EMISSÃO", "VALOR TOTAL DOCUMENTO", "FRETE", "ITEM", "OPER SPM",
                          "QTD FISCAL", "VALOR UNIT", "VALOR DO ICMS"] + base + ["ALIQUOTA"])
        wb.active.append([CHAVE_A, "30 46", "01/11/2026", 10, "CIF", "900135", "5101", 1, 1, 180] + [1500] * len(base) + [12])
        caminho = pathlib.Path(tempfile.mkdtemp()) / "nota.xlsx"
        wb.save(caminho)
        return caminho

    def test_nome_antigo_da_base(self):
        _, notas, erros = snac0236.carregar(self.gravar_cabecalho(["ICMS TERCEIROS (BASE DE CALCULO NF)"]))
        self.assertEqual(erros, [])
        self.assertEqual(snac0236.valores_rateio(notas[0]["itens"][0]), {"ident": "1", "base": "1500", "aliquota": "12"})

    def test_base_em_duas_colunas(self):
        with self.assertRaises(ValueError) as erro:
            snac0236.ler_planilha(self.gravar_cabecalho(["ICMS TERCEIROS (BASE DE CALCULO NF)", "BASE CALC. ICMS"]))
        self.assertIn("Duas colunas para a base", str(erro.exception))

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
        # 11 campos + 4 do ICMS TERCEIROS, os Enter combinados, os Enter ate o Rateio e na linha, Tributos OK e Integrar
        self.assertGreater(len(passos), 15 + enters + 2)
        self.assertIn("Passo a passo: clicar em Tributos OK", passos)
        self.assertIn("Passo a passo: digitar '30' em Filial Fisica (1a parte)", passos)
        self.assertEqual(passos[-9:], ["Passo a passo: clicar em Integrar"] + [f"Passo a passo: Enter {i} de 7 depois do Integrar" for i in range(1, 8)] + ["Passo a passo: clicar em Integrar de novo"])
        self.assertGreater(__import__("time").time() - inicio, 0.01 * len(passos) * 0.9)


class Tela(unittest.TestCase):
    def test_foco_data_selecionada_com_texto(self):
        # Data Emissao preenchida pelo SRP: sem ciano, fundo azul da selecao com a data em branco por cima
        imagem = Image.new("RGB", (800, 300), "white")
        imagem.paste((72, 100, 141), (400, 100, 560, 138))
        for x in range(410, 550, 7):  # tracos brancos do texto
            imagem.paste((255, 255, 255), (x, 108, x + 2, 130))
        caixa = snac0236.foco_unido(imagem)
        self.assertGreaterEqual(caixa[3] - caixa[1], 30, caixa)  # o campo inteiro, nao so uma faixa no alto
        self.assertLessEqual(abs((caixa[1] + caixa[3]) / 2 - 119), 3, caixa)

    def test_foco_data_selecionada_tela_pequena(self):
        # Citrix em 1366x768 (05/11/2026): campo de 14 px, selecao de 11 px e a data em tracos brancos de 1 px
        imagem = Image.new("RGB", (1366, 768), "white")
        imagem.paste((72, 100, 141), (300, 0, 1366, 1))  # borda da janela no mesmo azul: nao conta
        imagem.paste((72, 100, 141), (140, 212, 190, 223))
        for x in range(142, 188, 3):  # digitos: tracos verticais de 1 px
            imagem.paste((255, 255, 255), (x, 214, x + 1, 221))
        imagem.paste((255, 255, 255), (142, 217, 188, 218))  # e um horizontal no meio
        caixa = snac0236.foco_unido(imagem)
        self.assertIsNotNone(caixa)
        self.assertLessEqual(abs((caixa[1] + caixa[3]) / 2 - 217.5), 2, caixa)
        self.assertLessEqual(abs(caixa[0] - 140), 3, caixa)
        self.assertLessEqual(abs(caixa[2] - 190), 3, caixa)

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


class RateioTelaPequena(unittest.TestCase):
    def test_ident_pelo_tributo_sozinho(self):
        # 1366x768 (05/11/2026): o OCR leu so a 2a linha do titulo "Ident / Tributo"
        palavras = [("Nome", 147, 408, 30, 9), ("Tributo", 180, 408, 29, 9), ("Tributo", 272, 408, 33, 9),
                    ("Tribut", 454, 407, 32, 10), ("Base", 524, 412, 30, 9), ("Calculo", 557, 412, 39, 9)]
        self.assertEqual(snac0236.localizar_rateio(palavras)["ident"], (272, 408, 305, 417))

    def test_junta_leituras_em_tela_pequena(self):
        leituras = {2: [("ICMS", 59, 467, 30, 8), ("TERCEIROS", 92, 467, 78, 8), ("Base", 524, 412, 30, 9),
                        ("Calculo", 557, 412, 39, 9), ("Aliquota", 613, 413, 38, 9), ("Lancamento", 700, 413, 57, 11),
                        ("Tributos", 365, 563, 40, 8), ("OK", 408, 563, 14, 8)],
                    4: [("Ident", 278, 393, 23, 8)]}
        imagem = Image.new("RGB", (1366, 768), "white")
        with mock.patch.object(core, "palavras_tela", lambda img, t, escala=2: leituras[escala]):
            achados = snac0236.ler_rateio(imagem, None)
        self.assertEqual(set(achados), {"terceiros", "ident", "base", "aliquota", "lancamento", "tributos_ok"})
        self.assertEqual(achados["ident"], (278, 393, 301, 401))

    def test_descricao_do_3o_item_nao_e_o_ident(self):
        # 1366x768 (06/11/2026): no Enter 7 o foco vai a Descricao do 3o item, na altura da linha ICMS TERCEIROS do
        # Rateio salvo; tomada pelo Ident, a conferencia falhava e o robo lia a tela inteira (~11 s a mais)
        salvo = {"terceiros": (59, 467, 170, 475), "base": (524, 412, 596, 421), "aliquota": (613, 413, 651, 422),
                 "lancamento": (700, 413, 757, 424), "ident": (275, 413, 307, 421), "tributos_ok": (367, 567, 424, 575)}
        descricao, ident = (203, 457, 410, 468), (275, 465, 302, 476)
        self.assertTrue(snac0236.na_linha(descricao, salvo["terceiros"]))  # mesma altura: so a largura separa
        self.assertFalse(snac0236.celula_ident(descricao, salvo))
        self.assertTrue(snac0236.celula_ident(ident, salvo))


class TituloAcima(unittest.TestCase):
    def recorte(self, caixa, topo_linha1=None):
        from PIL import Image
        tamanhos, original = [], snac0236.core.palavras_tela
        snac0236.core.palavras_tela = lambda img, _t, **_k: tamanhos.append(img.size) or [("icms", 0, 0, 10, 5)]
        try:
            self.assertEqual(snac0236.titulo_acima(Image.new("RGB", (1366, 768)), caixa, "tess", topo_linha1), "icms")
        finally:
            snac0236.core.palavras_tela = original
        return tamanhos[0]

    def test_3o_item_le_ate_acima_da_1a_linha(self):
        # 1366x768 (06/11/2026): linhas da grade a cada 18 px, campo de 10 px; 1a linha em 421, 3a em 457
        self.assertEqual(self.recorte((710, 421, 800, 431)), (110, 38))  # 1o item: 4 alturas acima, como antes
        self.assertEqual(self.recorte((710, 457, 800, 467), topo_linha1=421), (110, 74))  # 3o: desde 4 alturas acima da 1a


class TempoPrevisto(unittest.TestCase):
    def notas(self, *itens):
        return [{"itens": [{}] * n} for n in itens]

    def test_sem_nota_feita_usa_os_tempos_de_referencia(self):
        faltam = self.notas(1, 2, 1)
        esperado = 3 * snac0236.SEG_NOTA + snac0236.SEG_ITEM_EXTRA
        self.assertEqual(snac0236.tempo_previsto([], faltam, 0), esperado)

    def test_ritmo_medido_corrige_a_referencia(self):
        feitas, faltam = self.notas(1, 1), self.notas(1, 3)
        # 2 notas de 1 item levaram o dobro da referencia: o resto tambem vai levar o dobro
        gasto = 2 * 2 * snac0236.SEG_NOTA
        esperado = 2 * (2 * snac0236.SEG_NOTA + 2 * snac0236.SEG_ITEM_EXTRA)
        self.assertAlmostEqual(snac0236.tempo_previsto(feitas, faltam, gasto), esperado)

    def test_texto_duracao(self):
        self.assertEqual(snac0236.texto_duracao(6720), "~1 h 52 min")
        self.assertEqual(snac0236.texto_duracao(840), "~14 min")
        self.assertEqual(snac0236.texto_duracao(42), "~40 s")
        self.assertEqual(snac0236.texto_duracao(-3), "~5 s")


class NfReferencia(unittest.TestCase):
    def test_botao_ok_e_a_frase_ok_mais_baixa(self):
        self.assertEqual(snac0236.botao_ok_referencia(snac0236.frases(REFERENCIA)), BOTAO_OK_REF)
        self.assertTrue(snac0236.nf_referencia(snac0236.normalizar(" ".join(p[0] for p in REFERENCIA))))


class BarraDialogo(unittest.TestCase):
    def test_resumo_dos_tributos_variantes_do_ocr(self):
        self.assertTrue(snac0236.resumo_tributos("ae dos tributos documento confirmar cancelar"))
        self.assertTrue(snac0236.resumo_tributos("resume dos tributes do documento cofins see conftirmar | cancelar |"))
        self.assertFalse(snac0236.resumo_tributos("tributos ok integrar cancelar"))

    def test_barra_da_janela_por_cima(self):
        imagem = Image.new("RGB", (1600, 900), (240, 240, 240))
        imagem.paste((45, 132, 197), (2, 10, 1598, 40))  # barra da tela principal (largura toda): fica de fora
        imagem.paste((45, 132, 197), (300, 300, 900, 330))  # janela por cima
        imagem.paste((45, 132, 197), (100, 600, 400, 603))  # borda fina: fica de fora
        barras = snac0236.barras_dialogo(imagem)
        self.assertEqual(len(barras), 1, barras)
        self.assertLessEqual(max(abs(a - b) for a, b in zip(barras[0], (300, 300, 900, 330))), 2, barras)
        area = snac0236.area_dialogo(imagem, barras)
        self.assertLessEqual(area[0], 300)
        self.assertGreaterEqual(area[2], 898)
        self.assertGreaterEqual(area[3], 300 + 0.6 * 600)  # janela de ate 0,6 da largura cabe no recorte

    def test_sem_janela_por_cima(self):
        imagem = Image.new("RGB", (1600, 900), (240, 240, 240))
        imagem.paste((45, 132, 197), (2, 10, 1598, 40))
        self.assertEqual(snac0236.barras_dialogo(imagem), [])


class Fluxo(unittest.TestCase):
    def test_duas_notas(self):
        srp = SRP0236()
        notas = [nota(CHAVE_A, [item("900135"), item("900266", qtd="1.25")]), nota(CHAVE_B, [item("900135")])]
        notas[1]["itens"][0]["linha"] = 4
        with ambiente(srp):
            codigo, avisos, estados = rodar(srp, notas)
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(len(srp.integradas), 2)
        self.assertEqual(srp.reintegrou, 2)  # Integrar, 7 Enter e Integrar de novo em cada nota
        a = srp.integradas[0]
        self.assertEqual((a["chave"], a["filial1"], a["filial2"], a["data"], a["valor_total"], a["frete"]),
                         (CHAVE_A, "30", "46", "01/11/2026", "47407.28", "CIF"))
        self.assertEqual((a["item0"], a["oper0"], a["qtd0"], a["unit0"], a["icms0"], a["t114ident0"]),
                         ("900135", "5101", "2", "10.5", "1.89", "0"))
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

    def test_cliques_no_item_e_no_ident_terceiros(self):
        srp = SRP0236()
        with ambiente(srp) as pasta:
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")]), nota(CHAVE_B, [item("900266")])])
            log = next(pasta.glob("*snac0236.log")).read_text(encoding="utf-8")
        self.assertEqual(codigo, 0, avisos)
        # do Frete ao Item so Enter (o clique pulava a janela da NF de referencia, 05/11/2026)
        self.assertEqual((srp.cliques_item, srp.cliques_ident), (0, 2))
        # o Ident ja vinha com 1 e o clique nao seleciona: sem apagar antes ficaria "10" (sem imposto)
        self.assertIn("Item: cheguei com", log)
        self.assertIn("cliquei no Ident Tributo do ICMS TERCEIROS", log)
        self.assertEqual([n["t114ident0"] for n in srp.integradas], ["0", "0"])
        self.assertNotIn("t199ident0", srp.integradas[0])

    def test_nf_referencia_entre_o_frete_e_o_item(self):
        srp = SRP0236()
        srp.ref_no_caminho = 4  # abre no 4o Enter depois do Frete
        with ambiente(srp) as pasta:
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
            log = next(pasta.glob("*snac0236.log")).read_text(encoding="utf-8")
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.ok_referencia, srp.enters_referencia), (1, 0))  # OK clicado, nenhum Enter na grade dela
        self.assertIn("NF de referencia: cliquei em OK", log)
        self.assertEqual(srp.integradas[0]["item0"], "900135")

    def test_nf_referencia_no_caminho_sem_o_ok_lido(self):
        srp = SRP0236()
        srp.ref_no_caminho, srp.sem_ok_ref = 4, True
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.ok_referencia, srp.enters_referencia), (0, 7))
        self.assertEqual(srp.integradas[0]["item0"], "900135")

    def test_sem_janela_depois_do_1o_integrar_nao_le_a_tela_inteira_a_cada_enter(self):
        srp = SRP0236()
        srp.com_aviso_frete = True  # janelas so depois do 2o Integrar: 7 Enter sem janela antes
        with ambiente(srp) as pasta:
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
            log = next(pasta.glob("*snac0236.log")).read_text(encoding="utf-8")
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(srp.reintegrou, 1)
        self.assertEqual(log.count("tela inteira, janelas=[]"), 1)  # so a 1a leitura depois do 1o clique
        self.assertGreaterEqual(log.count("nao li (nenhuma barra de janela)"), 7)

    def test_cliques_que_nao_andam_seguem_com_enter(self):
        sem, com = SRP0236(), SRP0236()
        com.ignora_cliques = True
        for srp in (sem, com):
            with ambiente(srp):
                codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
            self.assertEqual(codigo, 0, avisos)
        self.assertEqual((com.cliques_item, com.cliques_ident), (0, 1))
        self.assertEqual(com.integradas[0]["t114ident0"], "0")
        self.assertGreater(com.enters, sem.enters)  # Enter pelas linhas de tributo que o clique pularia

    def test_janelas_pos_integrar_lidas_no_recorte(self):
        srp = SRP0236()
        srp.com_aviso_frete = True
        with ambiente(srp) as pasta:
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
            log = next(pasta.glob("*snac0236.log")).read_text(encoding="utf-8")
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.fechou_aviso, srp.confirmou, srp.sucessos), (1, 1, 1))
        # frete, resumo, ltcc (texto + botao binarizado) e sucesso no recorte, nao na tela inteira
        self.assertGreaterEqual(srp.recortes, 5)
        self.assertIn(f"recorte {AREA}, janelas=['ltcc']", log)

    def test_janelas_pos_integrar_sem_barra_azul(self):
        # barra de titulo nao achada (outra cor): le a tela inteira antes de dar Enter e na ultima espera
        srp = SRP0236()
        srp.com_aviso_frete, srp.sem_barra = True, True
        with ambiente(srp) as pasta:
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
            log = next(pasta.glob("*snac0236.log")).read_text(encoding="utf-8")
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.fechou_aviso, srp.confirmou, srp.sucessos, srp.enter_no_ltcc), (1, 1, 1, 0))
        self.assertEqual(srp.recortes, 0)
        self.assertIn("sem barra de titulo azul: daqui em diante leio sempre a tela inteira", log)

    def test_janelas_pos_integrar_recorte_sem_texto(self):
        # barra achada, mas o recorte nao reconhece a janela: le a tela inteira (nunca Enter as cegas com janela aberta)
        srp = SRP0236()
        srp.com_aviso_frete, srp.recorte_vazio = True, True
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.fechou_aviso, srp.confirmou, srp.sucessos, srp.enter_no_ltcc), (1, 1, 1, 0))

    def test_aviso_frete_cif_depois_do_integrar(self):
        srp = SRP0236()
        srp.com_aviso_frete = True
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")]), nota(CHAVE_B, [item("900266")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.reintegrou, srp.fechou_aviso, srp.aviso_frete), (2, 2, False))
        self.assertEqual((srp.confirmou, srp.resumo), (2, False))
        self.assertEqual((srp.sucessos, srp.ltcc, srp.sucesso), (2, False, False))  # clicou Frete LTCC e deu OK
        self.assertTrue(any("Integrado com Sucesso" in a for a in avisos))
        # integrada com sucesso: clica na Chave e segue sem pausar; so a ultima nota termina com Concluido
        self.assertEqual(srp.cliques_chave, 2)  # depois de cada nota integrada, tambem a ultima
        self.assertFalse([a for a in avisos if "PAUSADO" in a], avisos)
        self.assertEqual([n["chave"] for n in srp.integradas], [CHAVE_A, CHAVE_B])

    def test_janelas_logo_depois_do_1o_integrar(self):
        srp = SRP0236()
        srp.avisos_no_1o = True  # aviso, Resumo, Frete LTCC e Sucesso aparecem durante os 7 Enter
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")]), nota(CHAVE_B, [item("900266")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.fechou_aviso, srp.confirmou, srp.sucessos, srp.reintegrou, srp.enter_no_ltcc), (2, 2, 2, 0, 0))
        self.assertEqual(srp.cliques_chave, 2)  # depois de cada nota integrada, tambem a ultima
        self.assertFalse([a for a in avisos if "PAUSADO" in a], avisos)

    def test_rateio_sem_imposto(self):
        srp = SRP0236(extras_rateio=3)  # mais Enter ate o Rateio que os 7 contados
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        nota_srp = srp.integradas[0]
        self.assertEqual({n: v for n, v in nota_srp.items() if n.startswith("t")},
                         {"t114ident0": "0", "t114base0": "0", "t114aliquota0": "0"})
        self.assertFalse([n for n in nota_srp if n.startswith(("t199", "t198"))])  # COFINS e PIS: nao mexe

    def test_rateio_com_imposto(self):
        srp = SRP0236()
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135", icms_t="1500.5", aliq="12")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual({n: v for n, v in srp.integradas[0].items() if n.startswith("t")},
                         {"t114ident0": "1", "t114base0": "1500.5", "t114aliquota0": "12"})

    def test_rateio_reaproveita_posicoes(self):
        srp = SRP0236()
        notas = [nota(CHAVE_A, [item("900135")]), nota(CHAVE_B, [item("900266")])]
        notas[1]["itens"][0]["linha"] = 3
        with ambiente(srp) as pasta:
            codigo, avisos, _ = rodar(srp, notas)
            log = next(pasta.glob("*snac0236.log")).read_text(encoding="utf-8")
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(log.count("Rateio lido no Enter"), 1)  # tela inteira so na 1a nota
        self.assertIn("Rateio: posicoes da nota anterior", log)
        self.assertEqual([n["t114ident0"] for n in srp.integradas], ["0", "0"])

    def test_rateio_posicoes_da_execucao_anterior(self):
        srp = SRP0236()
        with ambiente(srp) as pasta:
            rodar(srp, [nota(CHAVE_A, [item("900135")])])
            for log in pasta.glob("*snac0236.log"):
                log.unlink()
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_B, [item("900266")])])  # outra execucao (outro planilha)
            log = next(pasta.glob("*snac0236.log")).read_text(encoding="utf-8")
        self.assertEqual(codigo, 0, avisos)
        self.assertNotIn("Rateio lido no Enter", log)  # a 1a nota ja usou as posicoes salvas
        self.assertIn("posicoes do Rateio da execucao anterior", log)

    def test_rateio_dois_itens(self):
        srp = SRP0236()
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135"), item("900266", icms_t="30", aliq="18")])])
        self.assertEqual(codigo, 0, avisos)
        nota_srp = srp.integradas[0]
        self.assertEqual((nota_srp["t114ident0"], nota_srp["t114ident1"], nota_srp["t114base1"], nota_srp["t114aliquota1"]),
                         ("0", "1", "30", "18"))
        # Tributos OK deixa o foco no Item da mesma linha: seta para baixo uma vez, e o 2o item na linha de baixo
        self.assertEqual(srp.setas, 1)
        self.assertEqual((nota_srp["item0"], nota_srp["item1"]), ("900135", "900266"))

    def test_tributos_ok_ja_desce(self):
        srp = SRP0236()
        srp.ok_desce = True  # sem seta: o foco ja esta no Item da linha de baixo
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135"), item("900266")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(srp.setas, 0)
        self.assertEqual((srp.integradas[0]["item0"], srp.integradas[0]["item1"]), ("900135", "900266"))

    def test_seta_que_nao_desce_pausa(self):
        srp = SRP0236()
        srp.ignora_seta = True  # nunca digita o 2o item por cima do 1o
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135"), item("900266")])], com_continuar=False)
        self.assertEqual(codigo, 1)
        self.assertTrue(any("seta para baixo" in a for a in avisos), avisos)
        self.assertEqual(srp.atual.get("item0"), "900135")
        self.assertNotIn("item1", srp.atual)

    def test_nf_referencia_clica_ok(self):
        srp = SRP0236()
        srp.com_referencia, srp.com_aviso_frete = True, True
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.ok_referencia, srp.enters_referencia), (1, 0))  # OK clicado, nenhum Enter na grade dela
        self.assertEqual((srp.reintegrou, srp.fechou_aviso, srp.confirmou, srp.sucessos), (1, 1, 1, 1))

    def test_nf_referencia_sem_o_ok_lido_fecha_com_enter(self):
        srp = SRP0236()
        srp.com_referencia, srp.sem_ok_ref, srp.com_aviso_frete = True, True, True
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual((srp.ok_referencia, srp.enters_referencia), (0, 7))  # no simulador fecha no 7o Enter
        self.assertEqual((srp.reintegrou, srp.sucessos), (1, 1))

    def test_rateio_sem_o_campo_pausa(self):
        srp = SRP0236(colunas_tributo=("ident", "base"))  # Aliquota nao para: pausa, sem Tributos OK
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])], com_continuar=False)
        self.assertEqual(codigo, 1)
        self.assertTrue(any("Aliquota" in a for a in avisos), avisos)

    def test_valor_unit_amarelo(self):
        srp = SRP0236()  # o Valor Unit nao fica ciano: o robo confere a entrada pela rolagem da grade
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135", unit="8.5858")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertEqual(srp.integradas[0]["unit0"], "8.5858")
        self.assertNotIn("qtdfisica0", srp.integradas[0])

    def test_enter_nao_entra_no_valor_unit(self):
        srp = SRP0236(parado={"qtdfisica"})  # Enter no Qtd.Fisica nao anda e a grade nao rola: nao digita o Valor Unit
        with ambiente(srp):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135", unit="8.5858")])], com_continuar=False)
        self.assertEqual(codigo, 1)
        self.assertTrue(any("nao saiu do Qtd.Fisica" in a for a in avisos), avisos)
        self.assertNotIn("qtdfisica0", srp.atual)

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
        rateio = PALAVRAS[PALAVRAS.index(("Ident", 800, 560, 40, 12)):]  # o Rateio e lido quando abre
        with ambiente(srp, palavras=[("Integrar", 500, 300, 50, 12)] + rateio):
            codigo, avisos, _ = rodar(srp, [nota(CHAVE_A, [item("900135")])])
        self.assertEqual(codigo, 0, avisos)
        self.assertTrue(any("nao li na tela os rotulos" in a for a in avisos))
        self.assertEqual(srp.integradas[0]["t114ident0"], "0")


if __name__ == "__main__":
    unittest.main()
