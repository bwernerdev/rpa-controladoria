"""SRP simulado para testar core.executar sem Citrix: campos travados, Tab recusado (aviso) e um usuario que reage a pausa."""
import contextlib
import pathlib
import sys
import tempfile
import threading
from unittest import mock

RAIZ = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "src"))
import core  # noqa: E402

CAIXAS = {"item": (44, 0, 118, 11), "custo": (424, 0, 466, 11), "conta": (475, 0, 499, 11), "quantidade": (656, 0, 740, 11)}
CENTROS = {"item": 81, "custo": 437, "conta": 488, "quantidade": 692}


class SRP:
    """`travas`: por linha, {'custo': bool, 'conta': bool}. `recusar`: {(indice_linha, campo)} onde o Tab nao avanca."""

    def __init__(self, travas, recusar=()):
        self.travas, self.recusar = travas, set(recusar)
        self.linha, self.campo, self.valores, self.lock = 0, "item", [{}], threading.Lock()
        self.alerta = False  # aviso do SRP na tela (icone vermelho) depois de um Tab recusado

    def ordem(self):
        t = self.travas[min(self.linha, len(self.travas) - 1)]
        return [c for c in core.ORDEM if not t.get(c)]

    def tela(self, _janela=None, _fixos=()):
        """(foco, ha aviso): o aviso some quando o usuario resolve o campo recusado."""
        return self.caixa(), self.alerta and (self.linha, self.campo) in self.recusar

    def caixa(self, _janela=None):
        e, _, d, _ = CAIXAS[self.campo]
        y = 503 + 19 * self.linha
        return (e, y, d, y + 11)

    def teclas(self, texto, pause=0):
        with self.lock:
            if texto == "{TAB}":
                if (self.linha, self.campo) in self.recusar:
                    self.alerta = True
                    return
                ordem = self.ordem()
                i = ordem.index(self.campo)
                if i + 1 < len(ordem):
                    self.campo = ordem[i + 1]
                else:
                    self.linha += 1
                    self.campo = "item"
                    self.valores.append({})
            else:
                self.valores[self.linha][self.campo] = self.valores[self.linha].get(self.campo, "") + texto


class Janela:
    def set_focus(self):
        pass


def registros(n, conta="301"):
    return [{"linha": i + 2, "item": f"10{i}", "custo": "7300", "conta": conta, "quantidade": str(i + 1)} for i in range(n)]


@contextlib.contextmanager
def ambiente(srp, com_cabecalho=True, extras=None):
    """Troca as dependencias de tela/teclado do core pelo SRP simulado; logs e aprendizado vao para uma pasta temporaria."""
    pasta = pathlib.Path(tempfile.mkdtemp())
    trocas = {
        "send_keys": srp.teclas, "caixa_foco": srp.caixa, "estado_tela": srp.tela, "capturar": lambda janela: None,
        "localizar_colunas": (lambda imagem, t: dict(CENTROS)) if com_cabecalho else (lambda imagem, t: None),
        "caminho_tesseract": lambda: pathlib.Path("tesseract-falso"), "titulo_janela_ativa": lambda: "Corporativo - Desktop Viewer",
        "mouse_no_canto": lambda: False, "janela_citrix": lambda trecho: Janela(), "ler_aviso": lambda imagem, t, fixos=(): None,
        "pasta_logs": lambda: pasta, "arquivo_aprendizado": lambda: pasta / "aprendizado.json",
    }
    trocas.update(extras or {})
    with contextlib.ExitStack() as pilha:
        for nome, valor in trocas.items():
            pilha.enter_context(mock.patch.object(core, nome, valor))
        yield pasta


def executar(srp, regs, retomar=None, ao_pausar=None, conferir=False, render=0.3, pausar=None):
    """Roda core.executar; `ao_pausar(srp, retomar, linha)` e chamado (em outra thread) a cada pausa."""
    avisos, estados = [], []

    def estado(linha, nome):
        estados.append((linha, nome))
        if nome == "pausado" and ao_pausar:
            threading.Thread(target=ao_pausar, args=(srp, retomar, linha), daemon=True).start()

    codigo = core.executar(regs, "Desktop Viewer", 0, render, 0, threading.Event(), avisos.append, estado, conferir, retomar,
                           pausar)
    return codigo, avisos, estados
