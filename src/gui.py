"""Janela do robo do SRP: arrastar a planilha, simular, preencher, retomar e recarregar."""
import ctypes
import json
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import ImageTk

from core import (VK_F8, VK_F9, aplicar_atualizacao, aplicar_sugestoes, atalhos_globais, baixar_atualizacao, buscar_atualizacao, caminho_tesseract, citrix_aberto, contas_suspeitas, descrever, executar,
                  gravar_progresso, itens_repetidos, ler_planilha, ler_progresso, pasta_base, pasta_logs, validar)
from icone import desenhar
from versao import NOME, VERSAO

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    RaizTk = TkinterDnD.Tk
except Exception:  # sem arrastar e soltar: o botao de escolher continua funcionando
    DND_FILES, RaizTk = None, tk.Tk

PALETAS = {
    "claro": {
        "fundo": "#f6f7f9", "cartao": "#ffffff", "borda": "#e2e5ea", "borda_forte": "#c9ced6", "texto": "#1f2328",
        "suave": "#6b7280", "destaque": "#2f5bd3", "destaque_hover": "#2749ad", "no_destaque": "#ffffff",
        "ok": "#1a7f37", "ok_fundo": "#f0f7f2", "erro": "#c62828", "erro_fundo": "#fdf2f2", "aviso": "#9a6700",
        "azul_fundo": "#eef3fd", "cinza_fundo": "#f1f3f5", "botao": "#e8ebef", "botao_hover": "#dde1e7", "console": "#fbfbfc", "console_texto": "#4b5260",
        "aviso_fundo": "#fff7e0", "aviso_borda": "#efd08a", "zebra": "#f9fafb",
    },
    "escuro": {
        "fundo": "#17191e", "cartao": "#1f2229", "borda": "#2e333c", "borda_forte": "#454c58", "texto": "#e6e8eb",
        "suave": "#9aa1ab", "destaque": "#6b93f0", "destaque_hover": "#86a7f3", "no_destaque": "#0f1115",
        "ok": "#57c27a", "ok_fundo": "#1c2a21", "erro": "#f07a7a", "erro_fundo": "#2e1f21", "aviso": "#d9a441",
        "azul_fundo": "#222c40", "cinza_fundo": "#262a32", "botao": "#2b3039", "botao_hover": "#353b46", "console": "#1b1e24", "console_texto": "#b8bec7",
        "aviso_fundo": "#2b2517", "aviso_borda": "#6b5524", "zebra": "#232730",
    },
}
TEMAS = {"auto": "Automático (igual ao Windows)", "claro": "Claro", "escuro": "Escuro"}
FILTROS = (("todas", "Todas", None), ("pendentes", "Pendentes", {"pendente", "executando"}), ("pausadas", "Pausadas", {"pausado"}),
           ("erros", "Erros", {"erro"}), ("concluidas", "Concluídas", {"ok"}))
MARCA_TROCADA = " ✎"  # na celula da Conta trocada pela sugestao do SRP
COR = dict(PALETAS["claro"])
FONTE = "Segoe UI"
FONTE_FORTE = "Segoe UI Semibold"
ESTADOS = {"pendente": "○  Pendente", "executando": "●  Em andamento", "pausado": "●  Pausada", "ok": "●  Concluída", "erro": "●  Erro"}
ARQUIVO_CONFIG = pasta_base() / "config.json"
ESCALA = 1.0  # DPI do monitor / 96; tamanhos em pixels passam por px()


def px(valor):
    return max(1, round(valor * ESCALA))


def fonte(pontos, forte=False, mono=False):
    """Fonte com tamanho em pixels da escala atual (tamanho negativo = pixels no Tk): nao depende do cache do
    tk scaling, que deixava parte das fontes no tamanho do monitor anterior."""
    familia = "Consolas" if mono else (FONTE_FORTE if forte else FONTE)
    return (familia, -max(1, round(pontos * ESCALA * 96 / 72)))


def ativar_dpi():
    """Declara ao Windows que o app desenha na escala de cada monitor (sem isso, a 150% a janela fica desfocada).
    Precisa rodar antes de criar qualquer janela."""
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # por monitor (v2)
        return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def dpi_da_janela(janela):
    try:
        return ctypes.windll.user32.GetDpiForWindow(int(janela.wm_frame(), 16)) or 96
    except (AttributeError, OSError, ValueError):
        return 96


def area_util():
    """(esq, topo, dir, base) da area de trabalho do monitor principal, sem a barra de tarefas."""
    try:
        import ctypes.wintypes
        ret = ctypes.wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(ret), 0)  # SPI_GETWORKAREA
        return ret.left, ret.top, ret.right, ret.bottom
    except (AttributeError, OSError):
        return 0, 0, 1366, 728


def tema_do_windows():
    """'claro' ou 'escuro' conforme o modo de cor dos apps no Windows (claro se nao der para ler)."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as chave:
            return "claro" if winreg.QueryValueEx(chave, "AppsUseLightTheme")[0] else "escuro"
    except OSError:
        return "claro"


def tipo_mensagem(texto):
    """Cor da mensagem no console: 'erro', 'aviso', 'srp' ou None."""
    if re.search(r"Interrompido|Nao iniciei|^Erro|falhou|nao encontrad", texto):
        return "erro"
    if ": SRP: " in texto:
        return "srp"
    if re.search(r"AVISO|ATENCAO|PAUSADO|Pausado|Ainda pausado|^Aviso|^F9", texto):
        return "aviso"
    return None


def duracao(segundos):
    segundos = int(segundos)
    return f"{segundos // 3600}:{segundos // 60 % 60:02d}:{segundos % 60:02d}" if segundos >= 3600 else f"{segundos // 60:02d}:{segundos % 60:02d}"


def ler_config():
    try:
        return json.loads(ARQUIVO_CONFIG.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def gravar_config(dados):
    try:
        ARQUIVO_CONFIG.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


class BarraFina(tk.Frame):
    """Barra de progresso de 4 px; aceita configure(value=, maximum=) e barra["value"] como a ttk.Progressbar."""

    def __init__(self, pai):
        super().__init__(pai, bg=COR["borda"], height=px(4))
        self.preenchido = tk.Frame(self, bg=COR["destaque"], height=px(4))
        self._valor, self._maximo = 0, 1
        self._desenhar()

    def _desenhar(self):
        fracao = max(0.0, min(1.0, self._valor / self._maximo)) if self._maximo else 0.0
        self.preenchido.place(x=0, y=0, relheight=1, relwidth=fracao)

    def configure(self, cnf=None, **opcoes):
        if "value" in opcoes:
            self._valor = opcoes.pop("value")
        if "maximum" in opcoes:
            self._maximo = opcoes.pop("maximum") or 1
        self._desenhar()
        return super().configure(cnf, **opcoes) if (cnf or opcoes) else None

    config = configure

    def __setitem__(self, chave, valor):
        self.configure(**{chave: valor})

    def __getitem__(self, chave):
        return {"value": self._valor, "maximum": self._maximo}.get(chave) if chave in ("value", "maximum") else super().__getitem__(chave)


class App(RaizTk):
    def __init__(self):
        ativar_dpi()
        super().__init__()
        self.dpi = 96
        self._aplicar_escala(self.winfo_fpixels("1i"))
        self.config_usuario = ler_config()
        self.tema_escolhido = self.config_usuario.get("tema") if self.config_usuario.get("tema") in TEMAS else "auto"
        self.tema = tema_do_windows() if self.tema_escolhido == "auto" else self.tema_escolhido
        COR.update(PALETAS[self.tema])
        self.title(f"{NOME} v{VERSAO}")
        # cabe na area util do monitor (sem a barra de tarefas) mesmo a 150% ou em notebook, centralizada
        esq, topo, dir_, base = area_util()
        largura, altura = min(px(880), dir_ - esq - px(20)), min(px(700), base - topo - px(50))
        self.geometry(f"{largura}x{altura}+{esq + (dir_ - esq - largura) // 2}+{topo + max((base - topo - altura) // 2 - px(15), 0)}")
        self.minsize(px(800), px(600))
        self._icone = ImageTk.PhotoImage(desenhar(64))
        self.iconphoto(True, self._icone)
        self.registros, self.iids, self.estados = [], {}, {}
        self.originais = []  # registros como estao no Excel: o andamento e salvo por eles, mesmo com contas sugeridas aceitas
        self.trocas = {}  # linha -> conta da planilha, nas linhas em que o usuario aceitou a conta sugerida
        self.zebra = {}  # linha -> "par"/"impar" (fundo alternado na tabela)
        self.filtro = "todas"
        self.console_aberto = bool(self.config_usuario.get("console_aberto", False))
        self.execucao = None  # andamento da execucao em curso (painel, estimativa e resumo do fim)
        self.ultima_mensagem, self.ultimo_erro, self.srp_por_linha = "", "", {}
        self.dica = None
        self.pulso = False
        self.caminho_atual = None
        self.historico = []
        self.resumo = None  # (linhas, com_conta, erros)
        self.status_atual = ("Aguardando planilha", "suave")
        self.citrix = None
        self.ocupado = False
        self.fila = queue.Queue()
        self.parar = threading.Event()
        self.retomar = threading.Event()  # botao Continuar durante uma pausa
        self.pausar = threading.Event()  # F9: pausa ao terminar a linha atual
        self.soltar_atalhos = None
        self.pausado = False
        self.thread = None
        self.concluidas = 0
        self.limite = tk.StringVar(value="0")
        self.pausa = tk.StringVar(value="0.25")
        self.conferir = tk.BooleanVar(value=caminho_tesseract() is not None)
        if DND_FILES:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._soltou)
        self._construir()
        self.bind("<Configure>", self._talvez_mudou_monitor)
        threading.Thread(target=self._vigiar_citrix, daemon=True).start()
        self.after(100, self._drenar)
        self.after(1000, self._tique)
        self.atualizacao_pendente = None
        if getattr(sys, "frozen", False):  # so no exe; rodando pelo codigo-fonte nao se atualiza
            self.after(2000, lambda: threading.Thread(
                target=lambda: self.fila.put(("atualizacao", buscar_atualizacao(VERSAO))), daemon=True).start())

    # ---------- atualizacao ----------
    def _oferecer_atualizacao(self, info):
        if not info:
            return
        if self.ocupado:  # nao interrompe um preenchimento: pergunta quando terminar
            self.atualizacao_pendente = info
            return
        self.atualizacao_pendente = None
        notas = re.sub(r"[*#`]", "", info["notas"]).strip()
        notas = notas if len(notas) <= 700 else notas[:700].rsplit("\n", 1)[0] + "\n..."
        if not messagebox.askyesno(
                "Atualização disponível",
                f"Versão {info['versao']} disponível (você tem a {VERSAO}).\n\n{notas}\n\n"
                "Atualizar agora? O app fecha e abre de novo sozinho em alguns segundos. "
                "Configurações, contas aprendidas, andamento e logs são mantidos."):
            self.escrever(f"Atualizacao {info['versao']} adiada; o app pergunta de novo na proxima abertura.")
            return
        self.ocupado = True
        self._atualizar_botoes()
        self._status(f"Baixando a versão {info['versao']}...", "destaque")
        self.escrever(f"Baixando a versao {info['versao']} do GitHub...")

        def baixar():
            try:
                pacote = baixar_atualizacao(info, lambda feito, total: self.fila.put(("baixando", feito, total)))
                self.fila.put(("baixado", pacote))
            except Exception as erro:
                self.fila.put(("falha_atualizacao", str(erro)))
        threading.Thread(target=baixar, daemon=True).start()

    def _evento_atualizacao(self, item):
        if item[0] == "atualizacao":
            self._oferecer_atualizacao(item[1])
        elif item[0] == "baixando" and item[2]:
            self._status(f"Baixando atualização... {item[1] * 100 // item[2]}%", "destaque")
        elif item[0] == "baixado":
            self.escrever("Atualizacao baixada e conferida. O app vai fechar e abrir de novo na versao nova.")
            self._status("Instalando a atualização...", "destaque")
            self.update()
            aplicar_atualizacao(item[1], pasta_base())
            self.after(500, self.destroy)
        elif item[0] == "falha_atualizacao":
            self.ocupado = False
            self._atualizar_botoes()
            self._status("Atualização falhou: veja o console", "erro")
            self.escrever(f"A atualizacao falhou: {item[1]}. O app continua na versao {VERSAO}.")

    # ---------- visual ----------
    def _aplicar_escala(self, dpi):
        global ESCALA
        self.dpi = round(dpi)
        ESCALA = self.dpi / 96
        self.tk.call("tk", "scaling", self.dpi / 72)  # fontes em pontos acompanham o DPI

    def _talvez_mudou_monitor(self, evento):
        """Ao arrastar para um monitor de outra escala, refaz a janela no tamanho certo (o tkinter nao faz sozinho)."""
        if evento.widget is not self:
            return
        novo = dpi_da_janela(self)
        if novo == self.dpi:
            return
        proporcao = novo / self.dpi
        largura, altura = self.winfo_width(), self.winfo_height()
        self._aplicar_escala(novo)
        self.iids = {}
        self._construir()
        self.minsize(px(800), px(600))
        self.geometry(f"{round(largura * proporcao)}x{round(altura * proporcao)}")

    def _construir(self):
        for filho in self.winfo_children():
            filho.destroy()
        self.configure(bg=COR["fundo"])
        self._estilos()
        self._montar()
        self._repovoar()

    def _estilos(self):
        estilo = ttk.Style(self)
        estilo.theme_use("clam")
        estilo.configure(".", font=fonte(10), background=COR["fundo"], foreground=COR["texto"])
        base = dict(padding=(px(16), px(8)), borderwidth=1, relief="flat", focusthickness=0, lightcolor=COR["cartao"],
                    darkcolor=COR["cartao"])
        estilo.configure("Primario.TButton", **{**base, "lightcolor": COR["destaque"], "darkcolor": COR["destaque"]},
                         font=fonte(10, forte=True), background=COR["destaque"], foreground=COR["no_destaque"],
                         bordercolor=COR["destaque"])
        estilo.map("Primario.TButton", background=[("disabled", COR["cinza_fundo"]), ("active", COR["destaque_hover"])],
                   foreground=[("disabled", COR["suave"])], bordercolor=[("disabled", COR["borda"]), ("active", COR["destaque_hover"])],
                   lightcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["destaque_hover"])],
                   darkcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["destaque_hover"])])
        for nome, cor_texto in (("Secundario", COR["texto"]), ("Perigo", COR["erro"])):
            estilo.configure(f"{nome}.TButton", **{**base, "lightcolor": COR["botao"], "darkcolor": COR["botao"]},
                             font=fonte(10), background=COR["botao"], foreground=cor_texto, bordercolor=COR["botao"])
            estilo.map(f"{nome}.TButton", background=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])],
                       foreground=[("disabled", COR["borda_forte"])],
                       bordercolor=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])],
                       lightcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])],
                       darkcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])])
        estilo.configure("Link.TButton", font=fonte(9), padding=(px(6), px(2)), background=COR["fundo"], foreground=COR["destaque"],
                         borderwidth=0, relief="flat", focusthickness=0)
        estilo.map("Link.TButton", foreground=[("active", COR["destaque_hover"]), ("disabled", COR["suave"])],
                   background=[("active", COR["fundo"])])
        estilo.configure("Treeview", font=fonte(10), rowheight=px(26), background=COR["cartao"], fieldbackground=COR["cartao"],
                         foreground=COR["texto"], borderwidth=0, lightcolor=COR["cartao"], darkcolor=COR["cartao"],
                         bordercolor=COR["cartao"])
        estilo.configure("Treeview.Heading", font=fonte(9), background=COR["cartao"], foreground=COR["suave"],
                         relief="flat", padding=(px(8), px(6)), borderwidth=0, lightcolor=COR["cartao"], darkcolor=COR["borda"],
                         bordercolor=COR["borda"])
        estilo.map("Treeview.Heading", background=[("active", COR["cinza_fundo"])])
        estilo.map("Treeview", background=[("selected", COR["azul_fundo"])], foreground=[("selected", COR["texto"])])
        estilo.configure("Barra.Horizontal.TProgressbar", troughcolor=COR["borda"], background=COR["destaque"],
                         borderwidth=0, thickness=4, lightcolor=COR["destaque"], darkcolor=COR["destaque"])
        estilo.configure("Vertical.TScrollbar", background=COR["cinza_fundo"], troughcolor=COR["cartao"],
                         bordercolor=COR["cartao"], arrowcolor=COR["suave"], relief="flat", lightcolor=COR["cinza_fundo"],
                         darkcolor=COR["cinza_fundo"], gripcount=0)
        estilo.map("Vertical.TScrollbar", background=[("active", COR["borda_forte"])])
        estilo.configure("TEntry", padding=(px(6), px(3)), fieldbackground=COR["cartao"], foreground=COR["texto"],
                         bordercolor=COR["borda_forte"], lightcolor=COR["cartao"], darkcolor=COR["cartao"],
                         insertcolor=COR["texto"])
        estilo.map("TEntry", bordercolor=[("focus", COR["destaque"])], lightcolor=[("focus", COR["destaque"])])
        estilo.configure("TCheckbutton", background=COR["fundo"], foreground=COR["texto"], font=fonte(9),
                         indicatorbackground=COR["cartao"], indicatorforeground=COR["destaque"],
                         upperbordercolor=COR["borda_forte"], lowerbordercolor=COR["borda_forte"], focusthickness=0)
        estilo.map("TCheckbutton", background=[("active", COR["fundo"])],
                   indicatorbackground=[("selected", COR["cartao"]), ("active", COR["cinza_fundo"])])
        estilo.configure("TRadiobutton", background=COR["fundo"], foreground=COR["texto"], font=fonte(9),
                         indicatorbackground=COR["cartao"], indicatorforeground=COR["destaque"],
                         upperbordercolor=COR["borda_forte"], lowerbordercolor=COR["borda_forte"], focusthickness=0)
        estilo.map("TRadiobutton", background=[("active", COR["fundo"])],
                   indicatorbackground=[("selected", COR["cartao"]), ("active", COR["cinza_fundo"])])

    def _cartao(self, pai):
        return tk.Frame(pai, bg=COR["cartao"], highlightbackground=COR["borda"], highlightcolor=COR["borda"],
                        highlightthickness=1)

    def _link(self, pai, texto, comando, cor="destaque", bg=None):
        """Texto clicavel (sem borda de botao)."""
        rotulo = tk.Label(pai, text=texto, bg=bg or COR["fundo"], fg=COR[cor], font=fonte(9), cursor="hand2")
        rotulo.bind("<Button-1>", lambda _e: comando())
        rotulo.bind("<Enter>", lambda _e: rotulo.configure(fg=COR["destaque_hover"] if cor == "destaque" else COR["texto"]))
        rotulo.bind("<Leave>", lambda _e: rotulo.configure(fg=COR[cor]))
        return rotulo

    def _montar(self):
        margem = px(20)
        topo = tk.Frame(self, bg=COR["fundo"])
        topo.pack(fill="x", padx=margem, pady=(12, 10))
        lado = tk.Frame(topo, bg=COR["fundo"])
        lado.pack(side="right", anchor="n")
        titulos = tk.Frame(topo, bg=COR["fundo"])
        titulos.pack(side="left", fill="x", expand=True)
        tk.Label(titulos, text=NOME, bg=COR["fundo"], fg=COR["texto"], font=fonte(13, forte=True), anchor="w").pack(fill="x")
        tk.Label(titulos, text="Preenchimento de requisições do SRP a partir da planilha", bg=COR["fundo"],
                 fg=COR["suave"], font=fonte(9), anchor="w").pack(fill="x")
        self.b_opcoes = self._link(lado, "⚙  Opções", self.abrir_opcoes, cor="suave")
        self.b_opcoes.pack(side="right")
        self.chip_citrix = tk.Label(lado, text="", bg=COR["fundo"], fg=COR["suave"], font=fonte(9))
        self.chip_citrix.pack(side="right", padx=(0, 16))
        tk.Frame(self, bg=COR["borda"], height=1).pack(fill="x")

        corpo = tk.Frame(self, bg=COR["fundo"])
        corpo.pack(fill="both", expand=True, padx=margem, pady=(14, 12))

        # planilha: area grande de arrastar ate carregar; depois, cartao compacto
        self.zona = tk.Frame(corpo, bg=COR["cartao"], highlightbackground=COR["borda_forte"], highlightcolor=COR["borda_forte"],
                             highlightthickness=1)
        self.zona.pack(fill="x")
        if DND_FILES:
            self.zona.drop_target_register(DND_FILES)
            self.zona.dnd_bind("<<Drop>>", self._soltou)

        # faixa de pausa e painel da execucao (aparecem so quando servem)
        self.area_avisos = tk.Frame(corpo, bg=COR["fundo"])
        self.area_avisos.pack(fill="x")
        self.faixa_pausa = tk.Frame(self.area_avisos, bg=COR["aviso_fundo"], highlightbackground=COR["aviso_borda"],
                                    highlightcolor=COR["aviso_borda"], highlightthickness=1)
        textos = tk.Frame(self.faixa_pausa, bg=COR["aviso_fundo"])
        textos.pack(side="left", fill="both", expand=True, padx=px(14), pady=px(10))
        self.pausa_titulo = tk.Label(textos, text="", bg=COR["aviso_fundo"], fg=COR["aviso"], font=fonte(11, forte=True), anchor="w")
        self.pausa_titulo.pack(fill="x")
        self.pausa_srp = tk.Label(textos, text="", bg=COR["aviso_fundo"], fg=COR["texto"], font=fonte(10), anchor="w",
                                  justify="left", wraplength=px(560))
        self.pausa_instrucao = tk.Label(textos, text="", bg=COR["aviso_fundo"], fg=COR["suave"], font=fonte(9), anchor="w",
                                        justify="left", wraplength=px(560))
        self.pausa_instrucao.pack(fill="x", pady=(px(2), 0))
        self.b_continuar_faixa = ttk.Button(self.faixa_pausa, text="Continuar", style="Primario.TButton", command=self.continuar)
        self.b_continuar_faixa.pack(side="right", padx=px(14))

        self.painel = self._cartao(self.area_avisos)
        direita = tk.Frame(self.painel, bg=COR["cartao"])
        direita.pack(side="right", padx=px(14), pady=px(10))
        self.painel_tempo = tk.Label(direita, text="", bg=COR["cartao"], fg=COR["texto"], font=fonte(13, forte=True), anchor="e")
        self.painel_tempo.pack(fill="x")
        self.painel_resta = tk.Label(direita, text="", bg=COR["cartao"], fg=COR["suave"], font=fonte(9), anchor="e")
        self.painel_resta.pack(fill="x")
        esquerda = tk.Frame(self.painel, bg=COR["cartao"])
        esquerda.pack(side="left", fill="x", expand=True, padx=px(14), pady=px(10))
        self.painel_linha = tk.Label(esquerda, text="", bg=COR["cartao"], fg=COR["texto"], font=fonte(13, forte=True), anchor="w")
        self.painel_linha.pack(fill="x")
        self.painel_detalhe = tk.Label(esquerda, text="", bg=COR["cartao"], fg=COR["suave"], font=fonte(9), anchor="w")
        self.painel_detalhe.pack(fill="x")

        # filtros da tabela (com a contagem de cada estado)
        filtros = tk.Frame(corpo, bg=COR["fundo"])
        filtros.pack(fill="x", pady=(10, 6))
        self.botoes_filtro = {}
        for chave, titulo, _ in FILTROS:
            rotulo = tk.Label(filtros, text=titulo, bg=COR["fundo"], fg=COR["suave"], font=fonte(9), cursor="hand2",
                              padx=px(10), pady=px(3))
            rotulo.pack(side="left", padx=(0, px(4)))
            rotulo.bind("<Button-1>", lambda _e, c=chave: self.filtrar(c))
            self.botoes_filtro[chave] = (rotulo, titulo)

        rodape = tk.Frame(corpo, bg=COR["fundo"])
        rodape.pack(side="bottom", fill="x", pady=(6, 0))
        self.b_console = self._link(rodape, "", self.alternar_console)
        self.b_console.pack(side="left")
        self.b_copiar = self._link(rodape, "Copiar", self.copiar_console)
        self._link(rodape, "Abrir pasta de logs", lambda: os.startfile(pasta_logs())).pack(side="left", padx=(px(14), 0))
        tk.Label(rodape, text=f"v{VERSAO}", bg=COR["fundo"], fg=COR["suave"], font=fonte(8)).pack(side="right")
        self.caixa_saida = self._cartao(corpo)
        self.saida = tk.Text(self.caixa_saida, height=6, state="disabled", wrap="word", bg=COR["console"], fg=COR["console_texto"],
                             font=fonte(9, mono=True), relief="flat", borderwidth=0, padx=px(10), pady=px(6),
                             insertbackground=COR["texto"], highlightthickness=0)
        self.saida.pack(fill="x")
        for tipo, cor in (("erro", "erro"), ("aviso", "aviso"), ("srp", "destaque")):
            self.saida.tag_configure(tipo, foreground=COR[cor])

        self.botoes = tk.Frame(corpo, bg=COR["fundo"])
        self.botoes.pack(side="bottom", fill="x", pady=(12, 10))
        self.b_iniciar = ttk.Button(self.botoes, text="Iniciar preenchimento", style="Primario.TButton", command=self.iniciar)
        self.b_iniciar.pack(side="left")
        self.b_continuar = ttk.Button(self.botoes, text="Continuar", style="Secundario.TButton", command=self.continuar)
        self.b_continuar.pack(side="left", padx=(8, 0))
        self.b_simular = ttk.Button(self.botoes, text="Simular", style="Secundario.TButton", command=self.simular)
        self.b_simular.pack(side="left", padx=(8, 0))
        self.b_parar = ttk.Button(self.botoes, text="Parar", style="Perigo.TButton", command=self.parar.set)
        self.b_parar.pack(side="right")

        progresso = tk.Frame(corpo, bg=COR["fundo"])
        progresso.pack(side="bottom", fill="x", pady=(12, 0))
        self.status = tk.Label(progresso, text="", bg=COR["fundo"], fg=COR["suave"], font=fonte(9), anchor="w")
        self.status.pack(fill="x")
        self.barra = BarraFina(progresso)
        self.barra.pack(fill="x", pady=(4, 0))

        cartao = self._cartao(corpo)
        cartao.pack(fill="both", expand=True)
        colunas = ("linha", "item", "custo", "conta", "quantidade", "status")
        self.tabela = ttk.Treeview(cartao, columns=colunas, show="headings", height=3)
        for coluna, titulo, largura, ancora in zip(colunas, ("Linha", "Item", "C/Custo", "Conta", "Quantidade", "Status"),
                                                   tuple(px(v) for v in (56, 120, 100, 90, 100, 140)), ("center", "w", "w", "w", "center", "w")):
            self.tabela.heading(coluna, text=titulo, anchor=ancora)
            self.tabela.column(coluna, width=largura, minwidth=px(50), anchor=ancora, stretch=True)
        rolagem = ttk.Scrollbar(cartao, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.tabela.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
        # ordem importa: tags configuradas depois tem prioridade (o fundo do estado vence o da zebra)
        self.tabela.tag_configure("par", background=COR["zebra"])
        self.tabela.tag_configure("impar", background=COR["cartao"])
        self.tabela.tag_configure("ok", foreground=COR["ok"])
        self.tabela.tag_configure("erro", background=COR["erro_fundo"], foreground=COR["erro"])
        self.tabela.tag_configure("executando", background=COR["azul_fundo"], foreground=COR["texto"])
        self.tabela.tag_configure("pausado", background=COR["aviso_fundo"], foreground=COR["aviso"])
        self.tabela.tag_configure("pendente", foreground=COR["texto"])
        self.tabela.bind("<Motion>", self._dica_conta)
        self.tabela.bind("<Leave>", lambda _e: self._esconder_dica())

    def _pintar_zona(self):
        """Antes da planilha: area grande de arrastar. Depois: cartao com nome, resumo e acoes."""
        for filho in self.zona.winfo_children():
            filho.destroy()
        self.b_recarregar = None
        if not self.caminho_atual:
            self.zona.configure(cursor="hand2", highlightbackground=COR["borda_forte"])
            titulo = tk.Label(self.zona, text="Arraste a planilha aqui", bg=COR["cartao"], fg=COR["texto"],
                              font=fonte(11, forte=True), cursor="hand2")
            titulo.pack(pady=(px(16), 0))
            sub = tk.Label(self.zona, text="ou clique para escolher o arquivo (.xlsx)", bg=COR["cartao"],
                           fg=COR["suave"], font=fonte(9), cursor="hand2")
            sub.pack(pady=(0, px(16)))
            for parte in (self.zona, titulo, sub):
                parte.bind("<Button-1>", lambda _e: self.escolher())
                parte.bind("<Enter>", lambda _e: self.zona.configure(highlightbackground=COR["destaque"]))
                parte.bind("<Leave>", lambda _e: self.zona.configure(highlightbackground=COR["borda_forte"]))
            return
        self.zona.configure(cursor="", highlightbackground=COR["borda"])
        self.zona.unbind("<Button-1>")
        self.zona.unbind("<Enter>")
        self.zona.unbind("<Leave>")
        acoes = tk.Frame(self.zona, bg=COR["cartao"])
        acoes.pack(side="right", padx=px(14))
        self.b_recarregar = self._link(acoes, "Recarregar", self.recarregar, bg=COR["cartao"])
        self.b_recarregar.pack(side="right")
        self._link(acoes, "Trocar planilha", self.escolher, bg=COR["cartao"]).pack(side="right", padx=(0, px(14)))
        tk.Label(self.zona, text="▦", bg=COR["cartao"], fg=COR["ok"], font=fonte(20)).pack(side="left", padx=(px(14), px(8)))
        textos = tk.Frame(self.zona, bg=COR["cartao"])
        textos.pack(side="left", fill="x", expand=True, pady=px(8))
        caminho = Path(self.caminho_atual)
        tk.Label(textos, text=caminho.name, bg=COR["cartao"], fg=COR["texto"], font=fonte(10, forte=True),
                 anchor="w").pack(fill="x")
        linha = tk.Frame(textos, bg=COR["cartao"])
        linha.pack(fill="x")
        partes = []
        try:
            partes.append(f"salva em {time.strftime('%d/%m %H:%M', time.localtime(caminho.stat().st_mtime))}")
        except OSError:
            pass
        if self.resumo:
            linhas, com_conta, erros = self.resumo
            partes[:0] = [f"{linhas} linhas", f"{com_conta} com conta"]
        tk.Label(linha, text="  ·  ".join(partes), bg=COR["cartao"], fg=COR["suave"], font=fonte(9), anchor="w").pack(side="left")
        if self.resumo and self.resumo[2]:
            tk.Label(linha, text=f"  ·  {self.resumo[2]} erro(s) na planilha", bg=COR["cartao"], fg=COR["erro"],
                     font=fonte(9)).pack(side="left")

    def _tags(self, linha, nome):
        """Zebra so nas linhas sem fundo proprio: no Treeview a cor de fundo da zebra venceria a do estado."""
        return (nome,) if nome in ("erro", "executando", "pausado") else (self.zebra.get(linha, "impar"), nome)

    def _valores(self, r):
        conta = (r["conta"] or "—") + (MARCA_TROCADA if r["linha"] in self.trocas else "")
        return (r["linha"], r["item"], r["custo"], conta, r["quantidade"], ESTADOS[self.estados.get(r["linha"], "pendente")])

    def _repovoar(self):
        """Redesenha o estado atual (usado ao trocar de tema ou de monitor)."""
        for indice, r in enumerate(self.registros):
            self.zebra[r["linha"]] = "par" if indice % 2 else "impar"
            self.iids[r["linha"]] = self.tabela.insert("", "end", values=self._valores(r),
                                                       tags=self._tags(r["linha"], self.estados.get(r["linha"], "pendente")))
        self._pintar_zona()
        self.barra.configure(maximum=max(len(self.registros), 1), value=self.concluidas)
        self._status(*self.status_atual)
        self._pintar_citrix()
        self.saida.configure(state="normal")
        for texto in self.historico[-200:]:
            self.saida.insert("end", texto + "\n", tipo_mensagem(texto) or ())
        self.saida.see("end")
        self.saida.configure(state="disabled")
        self._mostrar_console()
        self.filtrar(self.filtro)
        self._mostrar_avisos()
        self._atualizar_botoes()

    def _pintar_citrix(self):
        if self.citrix is None:
            self.chip_citrix.configure(text="● Procurando o Citrix…", fg=COR["destaque"] if self.pulso else COR["borda_forte"])
        elif self.citrix:
            self.chip_citrix.configure(text="● Citrix conectado", fg=COR["ok"])
        else:
            self.chip_citrix.configure(text="● Citrix não encontrado, procurando…", fg=COR["erro"] if self.pulso else COR["borda_forte"])

    # ---------- tabela: filtro e dica da conta trocada ----------
    def filtrar(self, chave):
        self.filtro = chave
        alvo = dict((c, estados) for c, _, estados in FILTROS)[chave]
        posicao = 0
        for r in self.registros:
            iid = self.iids.get(r["linha"])
            if not iid:
                continue
            if alvo is None or self.estados.get(r["linha"], "pendente") in alvo:
                self.tabela.move(iid, "", posicao)
                posicao += 1
            else:
                self.tabela.detach(iid)
        self._contar_filtros()

    def _contar_filtros(self):
        for chave, estados in ((c, e) for c, _, e in FILTROS):
            rotulo, titulo = self.botoes_filtro[chave]
            total = len(self.registros) if estados is None else sum(
                1 for r in self.registros if self.estados.get(r["linha"], "pendente") in estados)
            ativo = chave == self.filtro
            rotulo.configure(text=f"{titulo}  {total}", bg=COR["azul_fundo"] if ativo else COR["fundo"],
                             fg=COR["destaque"] if ativo else (COR["erro"] if chave == "erros" and total else COR["suave"]),
                             font=fonte(9, forte=ativo))

    def _dica_conta(self, evento):
        """Ao passar o mouse na Conta trocada, mostra a conta que estava na planilha."""
        iid = self.tabela.identify_row(evento.y)
        linha = next((l for l, i in self.iids.items() if i == iid), None) if iid else None
        if not linha or linha not in self.trocas or self.tabela.identify_column(evento.x) != "#4":
            self._esconder_dica()
            return
        texto = f"Na planilha: {self.trocas[linha]}. Trocada pela conta que o SRP informou como correta."
        if self.dica and self.dica.texto == texto:
            self.dica.geometry(f"+{evento.x_root + 14}+{evento.y_root + 12}")
            return
        self._esconder_dica()
        self.dica = tk.Toplevel(self)
        self.dica.texto = texto
        self.dica.overrideredirect(True)
        tk.Label(self.dica, text=texto, bg=COR["texto"], fg=COR["fundo"], font=fonte(9), padx=px(8), pady=px(4)).pack()
        self.dica.geometry(f"+{evento.x_root + 14}+{evento.y_root + 12}")

    def _esconder_dica(self):
        if self.dica:
            self.dica.destroy()
            self.dica = None

    # ---------- console ----------
    def alternar_console(self):
        self.console_aberto = not self.console_aberto
        self.config_usuario["console_aberto"] = self.console_aberto
        gravar_config(self.config_usuario)
        self._mostrar_console()

    def _mostrar_console(self):
        if self.console_aberto:
            self.caixa_saida.pack(side="bottom", fill="x", before=self.botoes)
            self.b_copiar.pack(side="left", padx=(px(14), 0), after=self.b_console)
            self.saida.see("end")
        else:
            self.caixa_saida.pack_forget()
            self.b_copiar.pack_forget()
        self.b_console.configure(text=("Detalhes ▾" if self.console_aberto else "Detalhes ▸"))

    def copiar_console(self):
        self.clipboard_clear()
        self.clipboard_append("\n".join(self.historico))
        self._status("Detalhes copiados", "ok")

    # ---------- opcoes ----------
    def abrir_opcoes(self):
        janela = tk.Toplevel(self)
        janela.title("Opções")
        janela.configure(bg=COR["fundo"])
        janela.transient(self)
        janela.resizable(False, False)
        janela.iconphoto(False, self._icone)
        margem = px(16)
        corpo = tk.Frame(janela, bg=COR["fundo"])
        corpo.pack(fill="both", padx=margem, pady=margem)
        tk.Label(corpo, text="Preenchimento", bg=COR["fundo"], fg=COR["texto"], font=fonte(10, forte=True), anchor="w").pack(fill="x")
        ttk.Checkbutton(corpo, text="Conferir a quantidade na tela (OCR)", variable=self.conferir).pack(anchor="w", pady=(px(6), 0))
        for rotulo, variavel, ajuda in (("Limite de linhas", self.limite, "0 = todas"),
                                        ("Pausa entre teclas (s)", self.pausa, "padrão 0.25")):
            linha = tk.Frame(corpo, bg=COR["fundo"])
            linha.pack(fill="x", pady=(px(8), 0))
            tk.Label(linha, text=rotulo, bg=COR["fundo"], fg=COR["texto"], font=fonte(9), width=22, anchor="w").pack(side="left")
            ttk.Entry(linha, textvariable=variavel, width=6, justify="center", font=fonte(10)).pack(side="left")
            tk.Label(linha, text=ajuda, bg=COR["fundo"], fg=COR["suave"], font=fonte(8)).pack(side="left", padx=(px(8), 0))
        tk.Label(corpo, text="Tema", bg=COR["fundo"], fg=COR["texto"], font=fonte(10, forte=True),
                 anchor="w").pack(fill="x", pady=(px(16), 0))
        escolha = tk.StringVar(value=self.tema_escolhido)

        def mudar_tema():
            janela.destroy()  # a janela principal e refeita no tema novo
            self.escolher_tema(escolha.get())
        for chave, titulo in TEMAS.items():
            ttk.Radiobutton(corpo, text=titulo, value=chave, variable=escolha, command=mudar_tema).pack(anchor="w", pady=(px(4), 0))
        ttk.Button(corpo, text="Fechar", style="Secundario.TButton", command=janela.destroy).pack(anchor="e", pady=(px(16), 0))
        janela.update_idletasks()
        janela.geometry(f"+{self.b_opcoes.winfo_rootx() + self.b_opcoes.winfo_width() - janela.winfo_width()}"
                        f"+{self.b_opcoes.winfo_rooty() + self.b_opcoes.winfo_height() + px(6)}")
        janela.grab_set()

    # ---------- execucao: painel, faixa de pausa e resumo ----------
    def _mostrar_avisos(self):
        pausa = self.ocupado and self.pausado and self.execucao
        if pausa:
            self.faixa_pausa.pack(fill="x", pady=(px(10), 0))
        else:
            self.faixa_pausa.pack_forget()
        if self.ocupado and self.execucao and not pausa:  # na pausa a faixa ocupa o lugar do painel (sobra espaco para a tabela)
            self.painel.pack(fill="x", pady=(px(10), 0))
        else:
            self.painel.pack_forget()
        if pausa:
            linha = self.execucao.get("linha")
            r = next((r for r in self.registros if r["linha"] == linha), None)
            self.pausa_titulo.configure(text=f"Pausado na linha {linha}" + (f"  ·  item {r['item']}" if r else ""))
            srp = self.srp_por_linha.get(linha)
            if srp:
                self.pausa_srp.configure(text=f"SRP: “{srp}”")
                self.pausa_srp.pack(fill="x", pady=(px(4), 0), before=self.pausa_instrucao)
            else:
                self.pausa_srp.pack_forget()
            instrucao = re.sub(r"^Linha \d+: ", "", self.execucao.get("instrucao", ""))
            self.pausa_instrucao.configure(text=instrucao[:1].upper() + instrucao[1:])
        self._atualizar_painel()

    def _atualizar_painel(self):
        e = self.execucao
        if not (e and self.ocupado):
            return
        agora = time.time()
        pausado = e["tempo_pausado"] + (agora - e["pausado_desde"] if e["pausado_desde"] else 0)
        decorrido = agora - e["inicio"]
        linha = e.get("linha")
        r = next((r for r in self.registros if r["linha"] == linha), None)
        if r:
            self.painel_linha.configure(text=f"Linha {linha}   ·   {e['feitas'] + 1 if e['feitas'] < e['total'] else e['total']} de {e['total']}")
            self.painel_detalhe.configure(text=f"Item {r['item']}  ·  C/Custo {r['custo']}  ·  Conta {r['conta'] or '—'}  ·  "
                                               f"Quantidade {r['quantidade']}")
        else:
            self.painel_linha.configure(text="Preparando…")
            self.painel_detalhe.configure(text="Clique no campo Item da primeira linha livre no SRP")
        self.painel_tempo.configure(text=duracao(decorrido))
        faltam = e["total"] - e["feitas"]
        if e["feitas"] >= 2 and faltam > 0:
            resta = (decorrido - pausado) / e["feitas"] * faltam
            texto = f"~{max(1, round(resta / 60))} min restantes" if resta >= 90 else f"~{max(5, round(resta / 5) * 5)} s restantes"
        else:
            texto = "calculando o tempo restante…" if faltam else "terminando…"
        self.painel_resta.configure(text=texto + (f"  ·  {e['pausas']} pausa(s)" if e["pausas"] else ""))

    def _registrar_estado_execucao(self, linha, nome, anterior):
        e = self.execucao
        if not e:
            return
        if anterior == "pausado" and nome != "pausado" and e["pausado_desde"]:
            e["tempo_pausado"] += time.time() - e["pausado_desde"]
            e["pausado_desde"] = None
        if nome == "executando":
            e["linha"] = linha
        elif nome == "pausado":
            e["linha"] = linha
            e["instrucao"] = self.ultima_mensagem
            if anterior != "pausado":
                e["pausas"] += 1
                e["pausado_desde"] = time.time()
        elif nome == "ok" and anterior != "ok":
            e["feitas"] += 1
        elif nome == "erro":
            e["erros"] += 1

    def _resumo_final(self):
        e, self.execucao = self.execucao, None
        self._mostrar_avisos()
        if not e:
            return
        completo = e["feitas"] >= e["total"]
        janela = tk.Toplevel(self)
        janela.title("Resumo")
        janela.configure(bg=COR["fundo"])
        janela.transient(self)
        janela.resizable(False, False)
        janela.iconphoto(False, self._icone)
        margem = px(18)
        tk.Label(janela, text="Preenchimento concluído" if completo else "Preenchimento interrompido", bg=COR["fundo"],
                 fg=COR["ok"] if completo else COR["erro"], font=fonte(13, forte=True), anchor="w").pack(fill="x", padx=margem,
                                                                                                       pady=(margem, 0))
        if not completo and self.ultimo_erro:
            tk.Label(janela, text=self.ultimo_erro, bg=COR["fundo"], fg=COR["texto"], font=fonte(9), anchor="w", justify="left",
                     wraplength=px(440)).pack(fill="x", padx=margem, pady=(px(4), 0))
        numeros = tk.Frame(janela, bg=COR["fundo"])
        numeros.pack(fill="x", padx=margem, pady=(px(14), 0))
        for valor, legenda, cor in ((f"{e['feitas']}/{e['total']}", "linhas concluídas", "texto"),
                                    (str(e["erros"]), "com erro", "erro" if e["erros"] else "texto"),
                                    (str(e["pausas"]), "pausas", "aviso" if e["pausas"] else "texto"),
                                    (duracao(time.time() - e["inicio"]), "tempo total", "texto")):
            caixa = self._cartao(numeros)
            caixa.pack(side="left", fill="x", expand=True, padx=(0, px(8)))
            tk.Label(caixa, text=valor, bg=COR["cartao"], fg=COR[cor], font=fonte(15, forte=True)).pack(padx=px(14), pady=(px(8), 0))
            tk.Label(caixa, text=legenda, bg=COR["cartao"], fg=COR["suave"], font=fonte(8)).pack(padx=px(14), pady=(0, px(8)))
        botoes = tk.Frame(janela, bg=COR["fundo"])
        botoes.pack(fill="x", padx=margem, pady=margem)

        def nova():
            janela.destroy()
            self.escolher()
        ttk.Button(botoes, text="Fechar", style="Primario.TButton", command=janela.destroy).pack(side="right")
        ttk.Button(botoes, text="Nova planilha", style="Secundario.TButton", command=nova).pack(side="right", padx=(0, px(8)))
        ttk.Button(botoes, text="Abrir log", style="Secundario.TButton", command=self.abrir_ultimo_log).pack(side="right",
                                                                                                       padx=(0, px(8)))
        janela.update_idletasks()
        janela.geometry(f"+{self.winfo_rootx() + max((self.winfo_width() - janela.winfo_width()) // 2, 0)}"
                        f"+{self.winfo_rooty() + px(120)}")
        janela.lift()

    def abrir_ultimo_log(self):
        logs = sorted((p for p in pasta_logs().glob("*.log") if p.stem[:8].isdigit()), key=lambda p: p.stat().st_mtime)
        os.startfile(logs[-1] if logs else pasta_logs())

    def escolher_tema(self, escolha):
        self.tema_escolhido = escolha
        self.config_usuario["tema"] = escolha
        gravar_config(self.config_usuario)
        self._aplicar_tema(tema_do_windows() if escolha == "auto" else escolha)

    def _aplicar_tema(self, tema):
        if tema == self.tema:
            return
        self.tema = tema
        COR.update(PALETAS[tema])
        self.iids = {}
        self._construir()

    def _tique(self):
        """A cada segundo: relogio e estimativa do painel, pulso do Citrix procurado e tema automatico."""
        try:
            self._atualizar_painel()
            if self.citrix is not True:
                self.pulso = not self.pulso
                self._pintar_citrix()
            if self.tema_escolhido == "auto" and not self.ocupado and int(time.time()) % 5 == 0:
                self._aplicar_tema(tema_do_windows())
        finally:
            self.after(1000, self._tique)

    # ---------- comportamento ----------
    def escrever(self, texto):
        self.historico.append(texto)
        tipo = tipo_mensagem(texto)
        self.saida.configure(state="normal")
        self.saida.insert("end", texto + "\n", tipo or ())
        self.saida.see("end")
        self.saida.configure(state="disabled")
        if texto.startswith("Continuando"):
            return
        self.ultima_mensagem = texto
        srp = re.match(r'Linha (\d+): SRP: "(.*)"$', texto)
        if srp:  # texto do aviso lido em segundo plano: entra na faixa de pausa
            self.srp_por_linha[int(srp.group(1))] = srp.group(2)
            self._mostrar_avisos()
        elif tipo == "erro":
            self.ultimo_erro = texto
            if not self.console_aberto:  # com o console fechado, o motivo aparece no status
                self._status(texto if len(texto) <= 140 else texto[:137] + "...", "erro")

    def _status(self, texto, cor="suave"):
        self.status_atual = (texto, cor)
        self.status.configure(text=texto, fg=COR[cor])

    def _vigiar_citrix(self):
        while True:
            if not self.ocupado:
                self.fila.put(("citrix", citrix_aberto()))
            time.sleep(3)

    def _drenar(self):
        try:
            while True:
                item = self.fila.get_nowait()
                if item is None:
                    if self.soltar_atalhos:
                        self.soltar_atalhos()
                        self.soltar_atalhos = None
                    self.ocupado = False
                    self.pausado = False
                    self.pausar.clear()
                    self._atualizar_botoes()
                    self._fim_da_execucao()
                    self._resumo_final()
                    if self.atualizacao_pendente:
                        self.after(1500, lambda: self._oferecer_atualizacao(self.atualizacao_pendente))
                elif isinstance(item, tuple) and item[0] in ("atualizacao", "baixando", "baixado", "falha_atualizacao"):
                    self._evento_atualizacao(item)
                elif isinstance(item, tuple) and item[0] == "tecla":
                    self._tecla(item[1])
                elif isinstance(item, tuple) and item[0] == "citrix":
                    if item[1] != self.citrix:
                        self.citrix = item[1]
                        self._pintar_citrix()
                elif isinstance(item, tuple):
                    self._estado_linha(*item)
                else:
                    self.escrever(item)
        except queue.Empty:
            pass
        self.after(100, self._drenar)

    def _estado_linha(self, linha, nome):
        anterior = self.estados.get(linha)
        self.estados[linha] = nome
        if anterior == "ok" and nome != "ok":  # conferencia em segundo plano reprovou uma linha ja marcada como concluida
            self.concluidas = max(self.concluidas - 1, 0)
            self.barra["value"] = self.concluidas
        self._registrar_estado_execucao(linha, nome, anterior)
        iid = self.iids.get(linha)
        if not iid:
            return
        self.tabela.set(iid, "status", ESTADOS[nome])
        self.tabela.item(iid, tags=self._tags(linha, nome))
        if self.filtro == "todas":
            self._contar_filtros()
        else:
            self.filtrar(self.filtro)
        if nome in ("executando", "pausado"):
            self.pausado = nome == "pausado"
            self._atualizar_botoes()
            self._mostrar_avisos()
            try:
                self.tabela.see(iid)
            except tk.TclError:  # linha escondida pelo filtro
                pass
        if nome == "executando":
            self._status(f"Preenchendo linha {linha}...", "destaque")
        elif nome == "pausado":
            self._status(f"Pausado na linha {linha}: corrija no SRP e clique em Continuar", "aviso")
            self.bell()
            self.deiconify()
            self.lift()
        elif nome == "ok":
            self.concluidas += 1
            self.barra["value"] = self.concluidas
            self._status(f"{self.concluidas} de {len(self.registros)} concluídas", "ok")
        elif nome == "erro":
            self._status(f"Interrompido na linha {linha}", "erro")
        if (anterior == "ok") != (nome == "ok") and self.registros:  # guarda o andamento para o app reabrir dali
            gravar_progresso(self.originais, {l for l, e in self.estados.items() if e == "ok"})

    def _proxima_linha(self):
        for r in self.registros:
            if self.estados.get(r["linha"]) != "ok":
                return r["linha"]
        return None

    def _atualizar_botoes(self):
        pronto = bool(self.registros) and not self.ocupado
        retomar = pronto and 0 < self.concluidas < len(self.registros)
        self.b_iniciar.configure(state="normal" if pronto else "disabled")
        self.b_simular.configure(state="normal" if pronto else "disabled")
        self.b_parar.configure(state="normal" if self.ocupado else "disabled")
        if self.ocupado:
            self.b_continuar.configure(state="normal" if self.pausado else "disabled",
                                       text="Continuar (corrigido no SRP)" if self.pausado else "Continuar")
        else:
            self.b_continuar.configure(state="normal" if retomar else "disabled",
                                       text=f"Continuar da linha {self._proxima_linha()}" if retomar else "Continuar")
        self.b_continuar_faixa.configure(text="Continuar (corrigido no SRP)")
        if self.b_recarregar:
            self.b_recarregar.configure(state="disabled" if self.ocupado else "normal", disabledforeground=COR["borda_forte"])

    def continuar(self):
        """Durante uma pausa, libera o robo; depois de uma interrupcao, retoma pelas linhas nao concluidas."""
        if self.ocupado and self.pausado:
            self.escrever("Continuando: o robo volta ao Citrix e retoma do campo onde o cursor estiver.")
            self.pausado = False
            self._atualizar_botoes()
            self.retomar.set()
        elif not self.ocupado:
            self.iniciar(True)

    def _tecla(self, nome):
        """Atalhos globais durante a execucao: F9 pausa (ao fim da linha) ou continua; F8 para."""
        if not self.ocupado:
            return
        if nome == "parar":
            self.escrever("F8: parando.")
            self.parar.set()
        elif self.pausado:
            self.continuar()
        elif not self.pausar.is_set():
            self.pausar.set()
            self.escrever("F9: o robo pausa ao terminar a linha atual.")
            self._status("Pausando ao fim da linha...", "aviso")

    def _fim_da_execucao(self):
        if self.registros and self.concluidas == len(self.registros):
            self._status("Tudo concluído", "ok")

    def _soltou(self, evento):
        arquivos = self.tk.splitlist(evento.data)
        if arquivos:
            self.carregar(arquivos[0])

    def escolher(self):
        caminho = filedialog.askopenfilename(title="Planilha do SRP", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if caminho:
            self.carregar(caminho)

    def recarregar(self):
        if not self.caminho_atual:
            return
        if self.concluidas and not messagebox.askokcancel(
                "Recarregar", f"Recarregar apaga o andamento ({self.concluidas} linha(s) concluída(s)).\nContinuar?"):
            return
        if self.registros:
            gravar_progresso(self.originais, set())
        self.carregar(self.caminho_atual)

    def carregar(self, caminho):
        if self.ocupado:
            return
        self.tabela.delete(*self.tabela.get_children())
        for iid in self.iids.values():  # as escondidas pelo filtro nao aparecem em get_children
            if self.tabela.exists(iid):
                self.tabela.delete(iid)
        self.registros, self.iids, self.estados, self.concluidas, self.resumo = [], {}, {}, 0, None
        self.trocas, self.zebra, self.filtro = {}, {}, "todas"
        self.caminho_atual = caminho
        self.barra.configure(value=0, maximum=1)
        try:
            registros = ler_planilha(caminho)
        except Exception as erro:
            messagebox.showerror("Planilha", str(erro))
            self._atualizar_botoes()
            return
        erros = validar(registros)
        for indice, r in enumerate(registros):
            self.zebra[r["linha"]] = "par" if indice % 2 else "impar"
            self.iids[r["linha"]] = self.tabela.insert("", "end", values=self._valores(r), tags=self._tags(r["linha"], "pendente"))
        self.resumo = (len(registros), sum(1 for r in registros if r["conta"]), len(erros))
        self._pintar_zona()
        self.escrever(f"{len(registros)} linha(s) lidas de {caminho}")
        if erros:
            self._status("Corrija a planilha e clique em Recarregar", "erro")
            self.escrever("Planilha invalida, corrija no Excel, salve e clique em Recarregar:\n  " + "\n  ".join(erros))
            messagebox.showerror("Planilha inválida", "\n".join(erros[:10]) + "\n\nCorrija no Excel, salve e clique em 'Recarregar'.")
        elif not registros:
            self._status("Planilha vazia", "aviso")
        else:
            self.registros = registros
            self.originais = [dict(r) for r in registros]
            self.barra.configure(maximum=len(registros))
            self._status("Pronta para preencher", "ok")
            feitas, quando = ler_progresso(self.originais)
            if feitas:
                for linha in sorted(feitas):
                    self._estado_linha(linha, "ok")
                if len(feitas) < len(registros):
                    self.escrever(f"Andamento salvo em {quando}: {len(feitas)} de {len(registros)} linha(s) ja concluidas nesta "
                                  f"planilha. Use 'Continuar da linha {self._proxima_linha()}' ou 'Iniciar' para comecar do zero.")
                    self._status(f"Andamento anterior: continuar da linha {self._proxima_linha()}", "aviso")
                else:
                    self.escrever(f"Esta planilha ja foi preenchida inteira em {quando}. 'Iniciar' preenche tudo de novo.")
                    self._status("Planilha ja concluida antes", "aviso")
            repetidos = itens_repetidos(registros)
            if repetidos:
                lista = "; ".join(f"{item} (linhas {', '.join(map(str, linhas))})" for item, linhas in list(repetidos.items())[:5])
                self.escrever(f"AVISO: {len(repetidos)} item(ns) repetido(s): {lista}{' ...' if len(repetidos) > 5 else ''}. "
                              "O SRP soma as quantidades de itens iguais na mesma linha, entao a conferencia da quantidade "
                              "vai interromper. Para esta planilha, desmarque 'Conferir a quantidade' em Opcoes.")
                self._status("Itens repetidos: veja o aviso", "aviso")
            suspeitas = contas_suspeitas(registros)
            if suspeitas:
                aceitas = self._sugerir_contas(suspeitas)
                for linha, antes, depois in aplicar_sugestoes(registros, aceitas):
                    self.trocas[linha] = antes
                    self.tabela.set(self.iids[linha], "conta", depois + MARCA_TROCADA)
                    self.escrever(f"Linha {linha}: Conta {antes} trocada por {depois} (sugestao do SRP aceita; a planilha no "
                                  "Excel nao muda).")
                recusadas = [s for s in suspeitas if s[0] not in aceitas]
                for linha, item, conta, correta in recusadas:
                    self.escrever(f"AVISO: linha {linha}: item {item} com Conta {conta}, mas o SRP ja informou que a correta e {correta}. "
                                  "O robo vai pausar nela para voce corrigir no SRP.")
                if recusadas:
                    self._status(f"{len(recusadas)} conta(s) que o SRP ja recusou: veja o aviso", "aviso")
                else:
                    self._status(f"Pronta para preencher ({len(aceitas)} conta(s) sugerida(s) aceita(s))", "ok")
        self.filtrar("todas")
        self._atualizar_botoes()

    def _sugerir_contas(self, suspeitas):
        """Mostra as linhas com Conta que o SRP ja recusou e a conta que ele informou; devolve as linhas que o
        usuario marcou para trocar (nada muda sem o clique dele)."""
        janela = tk.Toplevel(self)
        janela.title("Contas que o SRP já recusou")
        janela.configure(bg=COR["fundo"])
        janela.transient(self)
        janela.iconphoto(False, self._icone)
        marcadas = {linha for linha, *_ in suspeitas}
        escolha = set()
        margem = px(16)
        tk.Label(janela, text=f"{len(suspeitas)} linha(s) com uma Conta que o SRP já recusou para o mesmo Item e C/Custo",
                 bg=COR["fundo"], fg=COR["texto"], font=fonte(11, forte=True), anchor="w").pack(fill="x", padx=margem,
                                                                                               pady=(margem, px(2)))
        tk.Label(janela, text="Clique numa linha para marcar ou desmarcar. As marcadas usam a conta que o SRP informou "
                              "como correta; as outras ficam como estão e o robô pausa nelas. A planilha no Excel não muda.",
                 bg=COR["fundo"], fg=COR["suave"], font=fonte(9), anchor="w", justify="left",
                 wraplength=px(560)).pack(fill="x", padx=margem, pady=(0, px(10)))
        cartao = self._cartao(janela)
        cartao.pack(fill="both", expand=True, padx=margem)
        colunas = ("usar", "linha", "item", "conta", "sugerida")
        tabela = ttk.Treeview(cartao, columns=colunas, show="headings", height=min(len(suspeitas), 10), selectmode="none")
        for coluna, titulo, largura in zip(colunas, ("Usar", "Linha", "Item", "Conta na planilha", "Conta sugerida"),
                                           (50, 60, 120, 130, 130)):
            tabela.heading(coluna, text=titulo, anchor="center")
            tabela.column(coluna, width=px(largura), anchor="center")
        for linha, item, conta, correta in suspeitas:
            tabela.insert("", "end", iid=str(linha), values=("✓", linha, item, conta, correta))
        if len(suspeitas) > 10:
            rolagem = ttk.Scrollbar(cartao, orient="vertical", command=tabela.yview)
            tabela.configure(yscrollcommand=rolagem.set)
            rolagem.pack(side="right", fill="y")
        tabela.pack(fill="both", expand=True, padx=1, pady=1)

        def alternar(evento):
            iid = tabela.identify_row(evento.y)
            if iid:
                marcadas.symmetric_difference_update({int(iid)})
                tabela.set(iid, "usar", "✓" if int(iid) in marcadas else "")

        def usar():
            escolha.update(marcadas)
            janela.destroy()
        tabela.bind("<Button-1>", alternar)
        botoes = tk.Frame(janela, bg=COR["fundo"])
        botoes.pack(fill="x", padx=margem, pady=margem)
        ttk.Button(botoes, text="Usar as marcadas", style="Primario.TButton", command=usar).pack(side="right")
        ttk.Button(botoes, text="Manter a planilha", style="Secundario.TButton", command=janela.destroy).pack(side="right",
                                                                                                         padx=(0, px(8)))
        janela.update_idletasks()
        janela.geometry(f"+{self.winfo_rootx() + max((self.winfo_width() - janela.winfo_width()) // 2, 0)}"
                        f"+{self.winfo_rooty() + px(80)}")
        janela.grab_set()
        self.wait_window(janela)
        return escolha

    def selecionadas(self, retomar=False):
        try:
            limite = int(self.limite.get() or 0)
            pausa = float(self.pausa.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Opções", "O limite deve ser um número inteiro e a pausa deve ser um número.")
            return None
        base = [r for r in self.registros if self.estados.get(r["linha"]) != "ok"] if retomar else self.registros
        return (base[:limite] if limite > 0 else base), pausa

    def simular(self):
        escolha = self.selecionadas()
        if escolha:
            for r in escolha[0]:
                self.escrever(f"linha {r['linha']}: {descrever(r)}")
            self.escrever("Simulacao: nada foi enviado. A Conta so e digitada se estiver editavel na tela.")

    def iniciar(self, retomar=False):
        escolha = self.selecionadas(retomar)
        if not escolha:
            return
        if not citrix_aberto():
            self.citrix = False
            self._pintar_citrix()
            messagebox.showwarning("Citrix não encontrado",
                                   "Não achei a janela do Citrix (Desktop Viewer).\nAbra o SRP e tente de novo.")
            return
        registros, pausa = escolha
        if not registros:
            messagebox.showinfo("Nada a fazer", "Não há linhas pendentes.")
            return
        if retomar:
            aviso = (f"Continuar a partir da linha {registros[0]['linha']} ({len(registros)} linha(s))?\n\n"
                     f"Antes: confira no SRP se a linha {registros[0]['linha']} ficou pela metade e limpe-a se precisar.\n"
                     "Depois do OK você terá 5 segundos para clicar no campo Item dela.")
        else:
            aviso = (f"Preencher {len(registros)} linha(s) no SRP?\n\n"
                     "Depois do OK você terá 5 segundos para clicar no campo Item da primeira linha livre.")
        suspeitas = contas_suspeitas(registros)
        if suspeitas:
            aviso += (f"\n\nATENÇÃO: {len(suspeitas)} linha(s) com Conta que o SRP já recusou (linhas "
                      + ", ".join(str(l) for l, *_ in suspeitas[:8]) + "). O robô vai pausar nelas.")
        if self.conferir.get() and itens_repetidos(self.registros):
            aviso += ("\n\nATENÇÃO: a planilha tem itens repetidos e a conferência da quantidade está ligada. "
                      "Como o SRP soma itens iguais, a conferência vai interromper nesses itens.")
        if not messagebox.askokcancel("Iniciar", aviso):
            return
        self.parar.clear()
        self.execucao = {"inicio": time.time(), "total": len(registros), "feitas": 0, "erros": 0, "pausas": 0,
                         "tempo_pausado": 0.0, "pausado_desde": None, "linha": None, "instrucao": ""}
        self.srp_por_linha, self.ultimo_erro = {}, ""
        if not retomar:
            self.concluidas = 0
            self.estados = {}
            gravar_progresso(self.originais, set())
            for linha, iid in self.iids.items():
                self.tabela.set(iid, "status", ESTADOS["pendente"])
                self.tabela.item(iid, tags=self._tags(linha, "pendente"))
            self.filtrar(self.filtro)
        else:
            for r in registros:
                self._estado_linha(r["linha"], "pendente")
        self.barra.configure(value=self.concluidas, maximum=len(self.registros))
        self._status("Iniciando...", "destaque")
        self.ocupado = True
        self.pausar.clear()
        self._atualizar_botoes()
        self._mostrar_avisos()
        self.soltar_atalhos, falhas = atalhos_globais({VK_F9: lambda: self.fila.put(("tecla", "pausa")),
                                                       VK_F8: lambda: self.fila.put(("tecla", "parar"))})
        if falhas:
            self.escrever("Aviso: F8/F9 ja estao em uso por outro programa; use os botoes ou o mouse no canto.")
        else:
            self.escrever("Atalhos: F9 pausa ao fim da linha (e continua), F8 para.")
        self.thread = threading.Thread(target=self._rodar, args=(registros, pausa, self.conferir.get()), daemon=True)
        self.thread.start()

    def _rodar(self, registros, pausa, conferir):
        try:
            executar(registros, "Desktop Viewer", pausa, 4.0, 5, self.parar, self.fila.put,
                     lambda linha, nome: self.fila.put((linha, nome)), conferir, self.retomar, self.pausar)
        finally:
            self.fila.put(None)


def fechar_abertura():
    """Fecha a tela de abertura do .exe (PyInstaller --splash); no codigo-fonte nao existe."""
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass


if __name__ == "__main__":
    app = App()
    app.after(0, fechar_abertura)
    app.mainloop()
