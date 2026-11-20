"""Visual comum das abas (v1.14.0; antes tudo ficava no gui.py e a aba SNAC0236 recebia o modulo gui inteiro como
`ui`): tema (PALETAS/COR), escala (ESCALA, px, esp, fonte), janelas do Windows (DPI, area util, tema, barra de titulo
escura, piscar na barra de tarefas), textos da tela (ESTADOS, tipo_mensagem, duracao) e os widgets proprios (Botao,
BarraFina, Cartao, Console, dialogo, icones). O COR e um dict unico: trocar o tema faz COR.update(...) e refaz a tela.
`ESCALA` muda com o monitor: use sempre visual.px()/esp()/fonte(), nunca copie o valor."""
import ctypes
import math
import re
import sys
import tkinter as tk
import types

from PIL import ImageTk

# customtkinter (v1.13.0; usuario, 10/11/2026: visual limitado): botoes arredondados (`Botao`) e a barra (`BarraFina`).
# O app cuida sozinho do tema (PALETAS) e da escala (px/fonte), entao: sem o darkdetect (o platform.release() dele
# custava ~0,1 s na abertura), escala do customtkinter fixa em 1 e sem os lacos dele rodando junto com o robo
# (tema a cada 30 ms e DPI a cada 100 ms; a troca de monitor e o _talvez_mudou_monitor).
if sys.modules.get("darkdetect") is None:  # o exe sai sem ele (build.ps1)
    sys.modules["darkdetect"] = types.SimpleNamespace(theme=lambda: "Light")
import customtkinter as ctk  # noqa: E402

ctk.deactivate_automatic_dpi_awareness()
ctk.set_appearance_mode("light")
ctk.AppearanceModeTracker.update = classmethod(lambda cls: None)
ctk.ScalingTracker.check_dpi_scaling = classmethod(lambda cls: None)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    RaizTk = TkinterDnD.Tk
except Exception:  # sem arrastar e soltar: o botao de escolher continua funcionando
    DND_FILES, RaizTk = None, tk.Tk


PALETAS = {
    "claro": {
        "fundo": "#f6f7f9", "cartao": "#ffffff", "borda": "#e2e5ea", "borda_forte": "#c9ced6", "texto": "#1f2328",
        "suave": "#6b7280", "destaque": "#2f5bd3", "destaque_hover": "#2749ad", "no_destaque": "#ffffff",
        "ok": "#177233", "ok_fundo": "#d9f2e1", "erro": "#c62828", "erro_fundo": "#fdf2f2", "aviso": "#9a6700",
        "azul_fundo": "#eef3fd", "cinza_fundo": "#f1f3f5", "botao": "#e8ebef", "botao_hover": "#dde1e7", "console": "#fbfbfc", "console_texto": "#4b5260",
        "aviso_fundo": "#fff7e0", "aviso_borda": "#efd08a", "zebra": "#f9fafb", "ok_serie_fundo": "#b4e0c1",
        "ok_serie_texto": "#0f5c27", "desativado": "#858b94", "tracejado": "#878e98",
    },
    "escuro": {
        "fundo": "#17191e", "cartao": "#1f2229", "borda": "#2e333c", "borda_forte": "#454c58", "texto": "#e6e8eb",
        "suave": "#9aa1ab", "destaque": "#6b93f0", "destaque_hover": "#86a7f3", "no_destaque": "#0f1115",
        "ok": "#57c27a", "ok_fundo": "#1f3d2a", "erro": "#f07a7a", "erro_fundo": "#2e1f21", "aviso": "#d9a441",
        "azul_fundo": "#222c40", "cinza_fundo": "#262a32", "botao": "#2b3039", "botao_hover": "#353b46", "console": "#1b1e24", "console_texto": "#b8bec7",
        "aviso_fundo": "#2b2517", "aviso_borda": "#6b5524", "zebra": "#232730", "ok_serie_fundo": "#2c5a3b",
        "ok_serie_texto": "#8fe0a8", "desativado": "#737a87", "tracejado": "#646c7a",
    },
}
TEMAS = {"auto": "Automático", "claro": "Claro", "escuro": "Escuro"}  # nomes no seletor das Opcoes
COR = dict(PALETAS["claro"])
FONTE = "Segoe UI"
FONTE_FORTE = "Segoe UI Semibold"
ESTADOS = {"pendente": "○  Pendente", "executando": "●  Em andamento", "pausado": "●  Pausada", "ok": "●  Concluída", "erro": "●  Erro"}
ESCALA = 1.0  # DPI do monitor / 96; tamanhos em pixels passam por px()


def px(valor):
    return max(1, round(valor * ESCALA))


def esp(passos):
    """Espacamento (margens e vaos) na grade de 4 px da tela (v1.14.0): esp(0.5)=2, esp(1)=4, esp(2)=8, esp(3)=12,
    esp(4)=16, esp(5)=20, esp(6)=24, ja na escala do monitor. Antes havia 2, 3, 4, 6, 8, 10, 12, 14, 16, 18, 20 e 24
    soltos, e parte sem px() (ficava menor que o resto a 150%)."""
    return px(passos * 4)


# escala de tipos (v1.14.0): os tamanhos da tela, em pontos, por nome
TIPO = {"legenda": 8, "pequeno": 9, "corpo": 10, "subtitulo": 11, "titulo": 13, "numero": 15}


def fonte(tamanho, forte=False, mono=False):
    """Fonte com tamanho em pixels da escala atual (tamanho negativo = pixels no Tk): nao depende do cache do
    tk scaling, que deixava parte das fontes no tamanho do monitor anterior. `tamanho` = nome do `TIPO` (ou pontos)."""
    pontos = TIPO[tamanho] if isinstance(tamanho, str) else tamanho
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


def barra_titulo(janela):
    """Barra de titulo do Windows no tema do app (escura no escuro). Atributo 20 do DWM (Windows 10 20H1+ e 11; 19
    nas builds antigas); sem ele, a barra fica como esta. O SetWindowPos com FRAMECHANGED redesenha a moldura na
    hora (sem isso a cor so muda quando a janela perde e ganha o foco)."""
    try:
        janela.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(janela.winfo_id())
        if not hwnd:
            return
        valor = ctypes.c_int(int(COR["fundo"] == PALETAS["escuro"]["fundo"]))
        for atributo in (20, 19):
            if not ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, atributo, ctypes.byref(valor), ctypes.sizeof(valor)):
                break
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x27)  # FRAMECHANGED|NOZORDER|NOSIZE|NOMOVE
    except (AttributeError, OSError, tk.TclError):
        pass


def piscar_barra(janela):
    """Pisca o botao do app na barra de tarefas ate o usuario voltar para ele (v1.14.0): na pausa o usuario esta no
    Citrix, e o lift() sozinho o Windows costuma barrar (a janela nao rouba o foco). FlashWindowEx com
    FLASHW_ALL | FLASHW_TIMERNOFG."""
    try:
        import ctypes.wintypes as wt

        class FLASHWINFO(ctypes.Structure):
            _fields_ = [("cbSize", wt.UINT), ("hwnd", wt.HWND), ("dwFlags", wt.DWORD), ("uCount", wt.UINT),
                        ("dwTimeout", wt.DWORD)]
        hwnd = ctypes.windll.user32.GetParent(janela.winfo_id())
        if hwnd:
            info = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd, 0x3 | 0xC, 0, 0)
            ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
    except (AttributeError, OSError, tk.TclError):
        pass


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


class BarraFina(ctk.CTkProgressBar):
    """Barra de progresso fina de pontas redondas; aceita configure(value=, maximum=) e barra["value"] como a
    ttk.Progressbar (a do clam ignora a espessura). So redesenha quando a fracao muda."""

    def __init__(self, pai):
        self._valor, self._maximo, self._fracao, self._cor = 0, 1, None, COR["destaque"]
        super().__init__(pai, height=px(6), corner_radius=px(3), border_width=0, fg_color=COR["borda"],
                         progress_color=COR["destaque"], mode="determinate")
        self._desenhar()

    def _desenhar(self):
        fracao = max(0.0, min(1.0, self._valor / self._maximo)) if self._maximo else 0.0
        if fracao != self._fracao:
            self._fracao = fracao
            self.set(fracao)
            cor = COR["destaque"] if fracao > 0 else COR["borda"]  # em 0 o customtkinter ainda desenha a ponta redonda
            if cor != self._cor:
                self._cor = cor
                ctk.CTkProgressBar.configure(self, progress_color=cor)

    def configure(self, require_redraw=False, **opcoes):
        if "value" in opcoes or "maximum" in opcoes:
            self._valor = opcoes.pop("value", self._valor)
            self._maximo = opcoes.pop("maximum", self._maximo) or 1
            self._desenhar()
        return super().configure(require_redraw, **opcoes) if (require_redraw or opcoes) else None

    config = configure

    def __setitem__(self, chave, valor):
        self.configure(**{chave: valor})

    def __getitem__(self, chave):
        return {"value": self._valor, "maximum": self._maximo}.get(chave) if chave in ("value", "maximum") else super().__getitem__(chave)


class Botao(ctk.CTkButton):
    """Botao de cantos redondos (customtkinter) que aceita os style= dos ttk.Button de antes ("Primario.TButton",
    "Secundario.TButton", "Perigo.TButton") e state=. Largura = texto + 16 px de cada lado (o CTkButton nao tem
    padding). O configure so repassa o que mudou: o _atualizar_botoes reconfigura a cada estado."""

    def __init__(self, pai, text="", style="Secundario.TButton", command=None, state="normal"):
        self._tk_pai, self._estilo, self._estado, self._texto = pai.tk, style, state, text
        super().__init__(pai, command=command, corner_radius=px(6), border_width=0, height=px(34), **self._opcoes())
        self._aplicadas = self._opcoes()

    def _opcoes(self):
        tipo = self._estilo.split(".")[0]
        primario = tipo == "Primario"
        letra = fonte("corpo", forte=primario)
        largura = max(px(72), int(self._tk_pai.call("font", "measure", letra, self._texto)) + px(32))
        opcoes = dict(text=self._texto, font=letra, width=largura, state=self._estado,
                      fg_color=COR["destaque"] if primario else COR["botao"],
                      hover_color=COR["destaque_hover"] if primario else COR["botao_hover"],
                      text_color=COR["no_destaque"] if primario else COR["erro"] if tipo == "Perigo" else COR["texto"],
                      text_color_disabled=COR["suave"] if primario else COR["desativado"])
        if self._estado == "disabled":
            opcoes["fg_color"] = COR["cinza_fundo"]
        return opcoes

    def configure(self, require_redraw=False, **opcoes):
        if not hasattr(self, "_aplicadas"):  # chamadas do proprio CTkButton durante o __init__ (cursor)
            return super().configure(require_redraw, **opcoes)
        self._estilo = opcoes.pop("style", self._estilo)
        self._estado = opcoes.pop("state", self._estado)
        self._texto = opcoes.pop("text", self._texto)
        novas = self._opcoes()
        mudou = {k: v for k, v in novas.items() if self._aplicadas.get(k) != v}
        self._aplicadas = novas
        if mudou or opcoes or require_redraw:
            super().configure(require_redraw, **mudou, **opcoes)

    config = configure


class Cartao(ctk.CTkFrame):
    """Cartao de cantos redondos (customtkinter) que aceita as opcoes de tk.Frame que o app usa: bg (fundo),
    highlightbackground (borda), highlightthickness (0 = sem borda); highlightcolor e cursor sao ignorados (o
    CTkFrame e coberto pelo proprio Canvas). Filhos tk com o fundo do cartao precisam de margem de ~4 px para nao
    cobrir os cantos. Arrastar e soltar: registra o cartao e o Canvas de dentro (e ele que recebe o mouse)."""

    def __init__(self, pai, bg=None, borda=None, raio=8):
        super().__init__(pai, corner_radius=px(raio), border_width=1, fg_color=bg or COR["cartao"],
                         border_color=borda or COR["borda"], width=px(40), height=px(40))
        self._pronto = True

    def configure(self, require_redraw=False, **opcoes):
        if not hasattr(self, "_pronto"):
            return super().configure(require_redraw, **opcoes)
        traduzidas = {}
        for chave, valor in opcoes.items():
            if chave in ("bg", "background"):
                traduzidas["fg_color"] = valor
            elif chave == "highlightbackground":
                traduzidas["border_color"] = valor
            elif chave == "highlightthickness":
                traduzidas["border_width"] = 1 if valor else 0
            elif chave not in ("highlightcolor", "cursor"):
                traduzidas[chave] = valor
        return super().configure(require_redraw, **traduzidas) if traduzidas or require_redraw else None

    config = configure

    def drop_target_register(self, *tipos):
        tk.BaseWidget.drop_target_register(self, *tipos)
        self._canvas.drop_target_register(*tipos)

    def dnd_bind(self, sequencia=None, funcao=None, add=None):
        self._canvas.dnd_bind(sequencia, funcao, add)
        return tk.BaseWidget.dnd_bind(self, sequencia, funcao, add)


class Console(ctk.CTkTextbox):
    """Detalhes (log da execucao): texto so leitura com barra de rolagem fina que so aparece quando o texto passa das
    `linhas` visiveis (v1.13.4). O CTkTextbox confere a barra num laco a cada 200 ms; aqui so confere depois de
    escrever/apagar e quando o tamanho muda (nada rodando em segundo plano). Cores por tipo: tags erro/aviso/srp."""

    def __init__(self, pai, linhas=6):
        letra = fonte("pequeno", mono=True)
        altura = linhas * int(pai.tk.call("font", "metrics", letra, "-linespace")) + 2 * px(6) + 2 * px(3) + px(4)
        super().__init__(pai, height=altura, corner_radius=0, border_width=0, border_spacing=px(3), wrap="word",
                         state="disabled", fg_color=COR["console"], text_color=COR["console_texto"], font=letra,
                         scrollbar_button_color=COR["borda_forte"], scrollbar_button_hover_color=COR["suave"],
                         padx=esp(2), pady=esp(2))
        self._conferindo = None
        self._textbox.bind("<Configure>", lambda _e: self._conferir_logo(), add="+")
        for tipo, cor in (("erro", "erro"), ("aviso", "aviso"), ("srp", "destaque")):
            self.tag_config(tipo, foreground=COR[cor])

    def _check_if_scrollbars_needed(self, event=None, continue_loop=False):
        self._conferindo = None
        if self.winfo_exists():
            super()._check_if_scrollbars_needed(event, False)

    def _conferir_logo(self):
        if self._conferindo is None:
            self._conferindo = self.after(50, self._check_if_scrollbars_needed)

    def insert(self, index, text, tags=None):
        self._conferir_logo()
        return super().insert(index, text, tags)

    def delete(self, index1, index2=None):
        self._conferir_logo()
        return super().delete(index1, index2)


def icone_planilha(lado, cor, fundo):
    """Icone de planilha desenhado (v1.13.3; o caractere da fonte saia serrilhado): folha de cantos redondos na `cor`,
    com a linha do cabecalho e a grade vazadas na cor do `fundo`; desenhado 4x e reduzido para suavizar."""
    from PIL import Image, ImageDraw
    grande = lado * 4
    imagem = Image.new("RGBA", (grande, grande), (0, 0, 0, 0))
    desenho = ImageDraw.Draw(imagem)
    x0 = y0 = grande * 0.08
    x1 = y1 = grande * 0.92
    desenho.rounded_rectangle((x0, y0, x1, y1), radius=grande * 0.16, fill=cor)
    traco = max(4, grande // 16)
    cabecalho = y0 + (y1 - y0) * 0.32
    for fracao in (0.32, 0.55, 0.78):
        y = y0 + (y1 - y0) * fracao
        desenho.line([(x0, y), (x1, y)], fill=fundo, width=traco)
    x = x0 + (x1 - x0) * 0.38
    desenho.line([(x, cabecalho), (x, y1)], fill=fundo, width=traco)
    return ImageTk.PhotoImage(imagem.resize((lado, lado), Image.LANCZOS))


def icone_dialogo(lado, tipo):
    """Circulo colorido com o simbolo do tipo de dialogo ('info', 'pergunta', 'aviso', 'erro'), desenhado so com
    formas (o exe sai sem o FreeType do PIL, entao nada de texto na imagem); 4x e reduzido para suavizar."""
    from PIL import Image, ImageDraw
    cor = {"aviso": COR["aviso"], "erro": COR["erro"]}.get(tipo, COR["destaque"])
    g = lado * 4
    imagem = Image.new("RGBA", (g, g), (0, 0, 0, 0))
    d = ImageDraw.Draw(imagem)
    d.ellipse((0, 0, g - 1, g - 1), fill=cor)
    branco, traco = COR["no_destaque"], max(4, g // 9)
    if tipo == "erro":
        a, b = g * 0.32, g * 0.68
        d.line([(a, a), (b, b)], fill=branco, width=traco)
        d.line([(a, b), (b, a)], fill=branco, width=traco)
    elif tipo == "aviso":  # "!": traco e ponto
        d.rounded_rectangle((g / 2 - traco / 2, g * 0.22, g / 2 + traco / 2, g * 0.6), radius=traco / 2, fill=branco)
        d.ellipse((g / 2 - traco * 0.6, g * 0.68, g / 2 + traco * 0.6, g * 0.68 + traco * 1.2), fill=branco)
    else:  # "i": ponto e traco
        d.ellipse((g / 2 - traco * 0.6, g * 0.2, g / 2 + traco * 0.6, g * 0.2 + traco * 1.2), fill=branco)
        d.rounded_rectangle((g / 2 - traco / 2, g * 0.42, g / 2 + traco / 2, g * 0.78), radius=traco / 2, fill=branco)
    return ImageTk.PhotoImage(imagem.resize((lado, lado), Image.LANCZOS))


def dialogo(dono, titulo, texto, botoes, tipo="info", lista=None, padrao=-1, cancelar=None, titulo_janela=None):
    """Janela de mensagem no tema do app (v1.14.0; as messagebox do Windows ficavam claras no tema escuro e eram
    paredes de texto). `botoes`: [(rotulo, valor, estilo)] da esquerda para a direita; `padrao` = indice do botao
    do Enter; Esc e o X devolvem `cancelar`. `lista`: itens mostrados num cartao (ate 12; o resto vira "e mais N").
    Modal: devolve o valor do botao clicado."""
    raiz = dono.winfo_toplevel()
    janela = tk.Toplevel(raiz)
    janela.withdraw()
    janela.title(titulo_janela or titulo)
    janela.configure(bg=COR["fundo"])
    janela.transient(raiz)
    janela.resizable(False, False)
    icone_app = getattr(raiz, "_icone", None)
    if icone_app:
        janela.iconphoto(False, icone_app)
    resposta = {"valor": cancelar}
    margem = esp(5)
    corpo = tk.Frame(janela, bg=COR["fundo"])
    corpo.pack(fill="both", expand=True, padx=margem, pady=(margem, 0))
    imagem = icone_dialogo(px(28), tipo)
    simbolo = tk.Label(corpo, image=imagem, bg=COR["fundo"])
    simbolo.image = imagem
    simbolo.grid(row=0, column=0, rowspan=3, sticky="n", padx=(0, esp(4)), pady=(esp(0.5), 0))
    tk.Label(corpo, text=titulo, bg=COR["fundo"], fg=COR["texto"], font=fonte("titulo", forte=True), anchor="w",
             justify="left", wraplength=px(440)).grid(row=0, column=1, sticky="w")
    if texto:
        tk.Label(corpo, text=texto, bg=COR["fundo"], fg=COR["texto"], font=fonte("corpo"), anchor="w", justify="left",
                 wraplength=px(440)).grid(row=1, column=1, sticky="w", pady=(esp(2), 0))
    if lista:
        caixa = Cartao(corpo)
        caixa.grid(row=2, column=1, sticky="ew", pady=(esp(3), 0))
        itens = list(lista)
        mostrar = itens if len(itens) <= 12 else itens[:11] + [f"e mais {len(itens) - 11}"]
        for indice, item in enumerate(mostrar):
            tk.Label(caixa, text=f"•  {item}", bg=COR["cartao"], fg=COR["texto"], font=fonte("pequeno"), anchor="w",
                     justify="left", wraplength=px(410)).pack(
                fill="x", padx=esp(3), pady=(esp(2) if indice == 0 else 0, esp(2) if indice == len(mostrar) - 1 else 0))
    fila = tk.Frame(janela, bg=COR["fundo"])
    fila.pack(fill="x", padx=margem, pady=margem)

    def responder(valor):
        resposta["valor"] = valor
        janela.destroy()
    for indice, (rotulo, valor, estilo) in enumerate(reversed(botoes)):
        Botao(fila, text=rotulo, style=estilo, command=lambda v=valor: responder(v)).pack(
            side="right", padx=(0, esp(2) if indice else 0))
    if botoes:
        janela.bind("<Return>", lambda _e: responder(botoes[padrao][1]))
    janela.bind("<Escape>", lambda _e: responder(cancelar))
    janela.protocol("WM_DELETE_WINDOW", lambda: responder(cancelar))
    janela.update_idletasks()
    if raiz.winfo_viewable():
        x = raiz.winfo_rootx() + max((raiz.winfo_width() - janela.winfo_reqwidth()) // 2, 0)
        y = raiz.winfo_rooty() + max((raiz.winfo_height() - janela.winfo_reqheight()) // 3, 0)
    else:
        x = (janela.winfo_screenwidth() - janela.winfo_reqwidth()) // 2
        y = (janela.winfo_screenheight() - janela.winfo_reqheight()) // 3
    janela.geometry(f"+{x}+{y}")
    janela.deiconify()
    barra_titulo(janela)
    janela.lift()
    janela.focus_force()
    try:
        janela.grab_set()
    except tk.TclError:  # outra janela ja tem o grab (raro): segue sem
        pass
    raiz.wait_window(janela)
    return resposta["valor"]


def retangulo_redondo(tela, x1, y1, x2, y2, raio, **opcoes):
    """Retangulo de cantos redondos num Canvas (aceita fill, outline, width, dash, tags). Cada canto = arco de 9
    pontos: o smooth=True do Tk so chanfrava o canto."""
    pontos = []
    for cx, cy, inicio in ((x2 - raio, y1 + raio, -90), (x2 - raio, y2 - raio, 0), (x1 + raio, y2 - raio, 90),
                           (x1 + raio, y1 + raio, 180)):
        for passo in range(9):
            angulo = math.radians(inicio + passo * 90 / 8)
            pontos += (cx + raio * math.cos(angulo), cy + raio * math.sin(angulo))
    return tela.create_polygon(pontos, **opcoes)
