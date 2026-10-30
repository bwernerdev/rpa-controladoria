"""Janela do robo do SRP: arrastar a planilha, simular, preencher, retomar e recarregar."""
import json
import os
import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import ImageTk

from core import (caminho_tesseract, citrix_aberto, descrever, executar, itens_repetidos, ler_planilha, pasta_base,
                  pasta_logs, validar)
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
    },
    "escuro": {
        "fundo": "#17191e", "cartao": "#1f2229", "borda": "#2e333c", "borda_forte": "#454c58", "texto": "#e6e8eb",
        "suave": "#9aa1ab", "destaque": "#6b93f0", "destaque_hover": "#86a7f3", "no_destaque": "#0f1115",
        "ok": "#57c27a", "ok_fundo": "#1c2a21", "erro": "#f07a7a", "erro_fundo": "#2e1f21", "aviso": "#d9a441",
        "azul_fundo": "#222c40", "cinza_fundo": "#262a32", "botao": "#2b3039", "botao_hover": "#353b46", "console": "#1b1e24", "console_texto": "#b8bec7",
    },
}
COR = dict(PALETAS["claro"])
FONTE = "Segoe UI"
FONTE_FORTE = "Segoe UI Semibold"
ESTADOS = {"pendente": "Pendente", "executando": "Em andamento", "ok": "Concluída", "erro": "Erro"}
ARQUIVO_CONFIG = pasta_base() / "config.json"


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
        super().__init__(pai, bg=COR["borda"], height=4)
        self.preenchido = tk.Frame(self, bg=COR["destaque"], height=4)
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
        super().__init__()
        self.config_usuario = ler_config()
        self.tema = self.config_usuario.get("tema", "claro") if self.config_usuario.get("tema") in PALETAS else "claro"
        COR.update(PALETAS[self.tema])
        self.title(f"{NOME} v{VERSAO}")
        self.geometry("880x700")
        self.minsize(800, 600)
        self._icone = ImageTk.PhotoImage(desenhar(64))
        self.iconphoto(True, self._icone)
        self.registros, self.iids, self.estados = [], {}, {}
        self.caminho_atual = None
        self.historico = []
        self.resumo = None  # (linhas, com_conta, erros)
        self.status_atual = ("Aguardando planilha", "suave")
        self.citrix = None
        self.ocupado = False
        self.fila = queue.Queue()
        self.parar = threading.Event()
        self.thread = None
        self.concluidas = 0
        self.limite = tk.StringVar(value="0")
        self.pausa = tk.StringVar(value="0.25")
        self.conferir = tk.BooleanVar(value=caminho_tesseract() is not None)
        if DND_FILES:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._soltou)
        self._construir()
        threading.Thread(target=self._vigiar_citrix, daemon=True).start()
        self.after(100, self._drenar)

    # ---------- visual ----------
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
        estilo.configure(".", font=(FONTE, 10), background=COR["fundo"], foreground=COR["texto"])
        base = dict(padding=(14, 6), borderwidth=1, relief="flat", focusthickness=0, lightcolor=COR["cartao"],
                    darkcolor=COR["cartao"])
        estilo.configure("Primario.TButton", **{**base, "lightcolor": COR["destaque"], "darkcolor": COR["destaque"]},
                         font=(FONTE_FORTE, 10), background=COR["destaque"], foreground=COR["no_destaque"],
                         bordercolor=COR["destaque"])
        estilo.map("Primario.TButton", background=[("disabled", COR["cinza_fundo"]), ("active", COR["destaque_hover"])],
                   foreground=[("disabled", COR["suave"])], bordercolor=[("disabled", COR["borda"]), ("active", COR["destaque_hover"])],
                   lightcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["destaque_hover"])],
                   darkcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["destaque_hover"])])
        for nome, cor_texto in (("Secundario", COR["texto"]), ("Perigo", COR["erro"])):
            estilo.configure(f"{nome}.TButton", **{**base, "lightcolor": COR["botao"], "darkcolor": COR["botao"]},
                             font=(FONTE, 10), background=COR["botao"], foreground=cor_texto, bordercolor=COR["botao"])
            estilo.map(f"{nome}.TButton", background=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])],
                       foreground=[("disabled", COR["borda_forte"])],
                       bordercolor=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])],
                       lightcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])],
                       darkcolor=[("disabled", COR["cinza_fundo"]), ("active", COR["botao_hover"])])
        estilo.configure("Link.TButton", font=(FONTE, 9), padding=(6, 2), background=COR["fundo"], foreground=COR["destaque"],
                         borderwidth=0, relief="flat", focusthickness=0)
        estilo.map("Link.TButton", foreground=[("active", COR["destaque_hover"]), ("disabled", COR["suave"])],
                   background=[("active", COR["fundo"])])
        estilo.configure("Treeview", font=(FONTE, 10), rowheight=26, background=COR["cartao"], fieldbackground=COR["cartao"],
                         foreground=COR["texto"], borderwidth=0, lightcolor=COR["cartao"], darkcolor=COR["cartao"],
                         bordercolor=COR["cartao"])
        estilo.configure("Treeview.Heading", font=(FONTE, 9), background=COR["cartao"], foreground=COR["suave"],
                         relief="flat", padding=(8, 6), borderwidth=0, lightcolor=COR["cartao"], darkcolor=COR["borda"],
                         bordercolor=COR["borda"])
        estilo.map("Treeview.Heading", background=[("active", COR["cinza_fundo"])])
        estilo.map("Treeview", background=[("selected", COR["azul_fundo"])], foreground=[("selected", COR["texto"])])
        estilo.configure("Barra.Horizontal.TProgressbar", troughcolor=COR["borda"], background=COR["destaque"],
                         borderwidth=0, thickness=4, lightcolor=COR["destaque"], darkcolor=COR["destaque"])
        estilo.configure("Vertical.TScrollbar", background=COR["cinza_fundo"], troughcolor=COR["cartao"],
                         bordercolor=COR["cartao"], arrowcolor=COR["suave"], relief="flat", lightcolor=COR["cinza_fundo"],
                         darkcolor=COR["cinza_fundo"], gripcount=0)
        estilo.map("Vertical.TScrollbar", background=[("active", COR["borda_forte"])])
        estilo.configure("TEntry", padding=(6, 3), fieldbackground=COR["cartao"], foreground=COR["texto"],
                         bordercolor=COR["borda_forte"], lightcolor=COR["cartao"], darkcolor=COR["cartao"],
                         insertcolor=COR["texto"])
        estilo.map("TEntry", bordercolor=[("focus", COR["destaque"])], lightcolor=[("focus", COR["destaque"])])
        estilo.configure("TCheckbutton", background=COR["fundo"], foreground=COR["texto"], font=(FONTE, 9),
                         indicatorbackground=COR["cartao"], indicatorforeground=COR["destaque"],
                         upperbordercolor=COR["borda_forte"], lowerbordercolor=COR["borda_forte"], focusthickness=0)
        estilo.map("TCheckbutton", background=[("active", COR["fundo"])],
                   indicatorbackground=[("selected", COR["cartao"]), ("active", COR["cinza_fundo"])])

    def _cartao(self, pai):
        return tk.Frame(pai, bg=COR["cartao"], highlightbackground=COR["borda"], highlightcolor=COR["borda"],
                        highlightthickness=1)

    def _chip(self, pai):
        """Indicador de resumo: so texto colorido (sem pilula)."""
        rotulo = tk.Label(pai, text="", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9))
        rotulo.pack(side="left", padx=(0, 14))
        return rotulo

    def _montar(self):
        margem = 20
        topo = tk.Frame(self, bg=COR["fundo"])
        topo.pack(fill="x", padx=margem, pady=(12, 10))
        lado = tk.Frame(topo, bg=COR["fundo"])
        lado.pack(side="right", anchor="n")
        titulos = tk.Frame(topo, bg=COR["fundo"])
        titulos.pack(side="left", fill="x", expand=True)
        tk.Label(titulos, text=NOME, bg=COR["fundo"], fg=COR["texto"], font=(FONTE_FORTE, 13), anchor="w").pack(fill="x")
        tk.Label(titulos, text="Preenchimento de requisições do SRP a partir da planilha", bg=COR["fundo"],
                 fg=COR["suave"], font=(FONTE, 9), anchor="w").pack(fill="x")
        ttk.Button(lado, text="Tema claro" if self.tema == "escuro" else "Tema escuro", style="Link.TButton",
                   command=self.alternar_tema).pack(side="right")
        self.chip_citrix = tk.Label(lado, text="", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9))
        self.chip_citrix.pack(side="right", padx=(0, 10))
        tk.Frame(self, bg=COR["borda"], height=1).pack(fill="x")

        corpo = tk.Frame(self, bg=COR["fundo"])
        corpo.pack(fill="both", expand=True, padx=margem, pady=(14, 12))

        self.zona = tk.Frame(corpo, bg=COR["cartao"], highlightbackground=COR["borda_forte"], highlightcolor=COR["borda_forte"],
                             highlightthickness=1, cursor="hand2")
        self.zona.pack(fill="x")
        self.titulo_zona = tk.Label(self.zona, text="Arraste a planilha aqui", bg=COR["cartao"], fg=COR["texto"],
                                    font=(FONTE_FORTE, 11), cursor="hand2")
        self.titulo_zona.pack(pady=(9, 0))
        self.sub_zona = tk.Label(self.zona, text="ou clique para escolher o arquivo (.xlsx)", bg=COR["cartao"],
                                 fg=COR["suave"], font=(FONTE, 9), cursor="hand2")
        self.sub_zona.pack(pady=(0, 9))
        for parte in (self.zona, self.titulo_zona, self.sub_zona):
            parte.bind("<Button-1>", lambda _e: self.escolher())
            parte.bind("<Enter>", lambda _e: self.zona.configure(highlightbackground=COR["destaque"]))
            parte.bind("<Leave>", lambda _e: self.zona.configure(highlightbackground=COR["borda_forte"]))
        if DND_FILES:
            self.zona.drop_target_register(DND_FILES)
            self.zona.dnd_bind("<<Drop>>", self._soltou)

        chips = tk.Frame(corpo, bg=COR["fundo"])
        chips.pack(fill="x", pady=(10, 8))
        opcoes = tk.Frame(chips, bg=COR["fundo"])
        opcoes.pack(side="right")
        ttk.Checkbutton(opcoes, text="Conferir quantidade", variable=self.conferir).pack(side="left", padx=(0, 14))
        tk.Label(opcoes, text="Limite", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9)).pack(side="left")
        ttk.Entry(opcoes, textvariable=self.limite, width=5, justify="center").pack(side="left", padx=(6, 12))
        tk.Label(opcoes, text="Pausa (s)", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9)).pack(side="left")
        ttk.Entry(opcoes, textvariable=self.pausa, width=5, justify="center").pack(side="left", padx=(6, 0))
        self.chip_linhas, self.chip_conta, self.chip_erros = self._chip(chips), self._chip(chips), self._chip(chips)

        rodape = tk.Frame(corpo, bg=COR["fundo"])
        rodape.pack(side="bottom", fill="x", pady=(6, 0))
        self.b_recarregar = ttk.Button(rodape, text="Recarregar planilha", style="Link.TButton", command=self.recarregar)
        self.b_recarregar.pack(side="left")
        ttk.Button(rodape, text="Abrir pasta de logs", style="Link.TButton", command=lambda: os.startfile(pasta_logs())).pack(side="left")
        tk.Label(rodape, text=f"v{VERSAO}", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 8)).pack(side="right")
        caixa_saida = self._cartao(corpo)
        caixa_saida.pack(side="bottom", fill="x")
        self.saida = tk.Text(caixa_saida, height=3, state="disabled", wrap="word", bg=COR["console"], fg=COR["console_texto"],
                             font=("Consolas", 9), relief="flat", borderwidth=0, padx=10, pady=6,
                             insertbackground=COR["texto"], highlightthickness=0)
        self.saida.pack(fill="x")

        botoes = tk.Frame(corpo, bg=COR["fundo"])
        botoes.pack(side="bottom", fill="x", pady=(12, 10))
        self.b_iniciar = ttk.Button(botoes, text="Iniciar preenchimento", style="Primario.TButton", command=self.iniciar)
        self.b_iniciar.pack(side="left")
        self.b_continuar = ttk.Button(botoes, text="Continuar", style="Secundario.TButton", command=lambda: self.iniciar(True))
        self.b_continuar.pack(side="left", padx=(8, 0))
        self.b_simular = ttk.Button(botoes, text="Simular", style="Secundario.TButton", command=self.simular)
        self.b_simular.pack(side="left", padx=(8, 0))
        self.b_parar = ttk.Button(botoes, text="Parar", style="Perigo.TButton", command=self.parar.set)
        self.b_parar.pack(side="right")

        progresso = tk.Frame(corpo, bg=COR["fundo"])
        progresso.pack(side="bottom", fill="x", pady=(12, 0))
        self.status = tk.Label(progresso, text="", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9), anchor="w")
        self.status.pack(fill="x")
        self.barra = BarraFina(progresso)
        self.barra.pack(fill="x", pady=(4, 0))

        cartao = self._cartao(corpo)
        cartao.pack(fill="both", expand=True)
        colunas = ("linha", "item", "custo", "conta", "quantidade", "status")
        self.tabela = ttk.Treeview(cartao, columns=colunas, show="headings", height=3)
        for coluna, titulo, largura, ancora in zip(colunas, ("Linha", "Item", "C/Custo", "Conta", "Quantidade", "Status"),
                                                   (56, 120, 100, 90, 100, 130), ("center", "w", "w", "w", "center", "w")):
            self.tabela.heading(coluna, text=titulo, anchor=ancora)
            self.tabela.column(coluna, width=largura, minwidth=50, anchor=ancora, stretch=True)
        rolagem = ttk.Scrollbar(cartao, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.tabela.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
        self.tabela.tag_configure("ok", foreground=COR["ok"])
        self.tabela.tag_configure("erro", background=COR["erro_fundo"], foreground=COR["erro"])
        self.tabela.tag_configure("executando", background=COR["azul_fundo"], foreground=COR["texto"])
        self.tabela.tag_configure("pendente", foreground=COR["texto"])

    def _repovoar(self):
        """Redesenha o estado atual (usado ao trocar de tema)."""
        for r in self.registros:
            iid = self.tabela.insert("", "end", values=(r["linha"], r["item"], r["custo"], r["conta"] or "—", r["quantidade"],
                                                         ESTADOS[self.estados.get(r["linha"], "pendente")]),
                                     tags=(self.estados.get(r["linha"], "pendente"),))
            self.iids[r["linha"]] = iid
        if self.caminho_atual:
            self.titulo_zona.configure(text=Path(self.caminho_atual).name)
            self.sub_zona.configure(text="clique ou arraste outra planilha para trocar")
        if self.resumo:
            linhas, com_conta, erros = self.resumo
            self._pintar(self.chip_linhas, f"{linhas} linhas", "azul")
            self._pintar(self.chip_conta, f"{com_conta} com conta", "azul")
            self._pintar(self.chip_erros, f"{erros} erros", "vermelho" if erros else "verde")
        else:
            for chip, texto in ((self.chip_linhas, "0 linhas"), (self.chip_conta, "0 com conta"), (self.chip_erros, "0 erros")):
                self._pintar(chip, texto, "neutro")
        self.barra.configure(maximum=max(len(self.registros), 1), value=self.concluidas)
        self._status(*self.status_atual)
        self._pintar_citrix()
        self.saida.configure(state="normal")
        self.saida.insert("end", "\n".join(self.historico[-200:]) + ("\n" if self.historico else ""))
        self.saida.see("end")
        self.saida.configure(state="disabled")
        self._atualizar_botoes()

    def _pintar(self, chip, texto, tipo):
        cor = {"azul": COR["texto"], "verde": COR["suave"], "vermelho": COR["erro"], "neutro": COR["suave"]}[tipo]
        chip.configure(text=texto, fg=cor)

    def _pintar_citrix(self):
        if self.citrix is None:
            self.chip_citrix.configure(text="Citrix: verificando…", fg=COR["suave"])
        elif self.citrix:
            self.chip_citrix.configure(text="● Citrix conectado", fg=COR["ok"])
        else:
            self.chip_citrix.configure(text="● Citrix não encontrado", fg=COR["erro"])

    def alternar_tema(self):
        self.tema = "claro" if self.tema == "escuro" else "escuro"
        COR.update(PALETAS[self.tema])
        self.config_usuario["tema"] = self.tema
        gravar_config(self.config_usuario)
        self.iids = {}
        self._construir()

    # ---------- comportamento ----------
    def escrever(self, texto):
        self.historico.append(texto)
        self.saida.configure(state="normal")
        self.saida.insert("end", texto + "\n")
        self.saida.see("end")
        self.saida.configure(state="disabled")

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
                    self.ocupado = False
                    self._atualizar_botoes()
                    self._fim_da_execucao()
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
        iid = self.iids.get(linha)
        if not iid:
            return
        self.tabela.set(iid, "status", ESTADOS[nome])
        self.tabela.item(iid, tags=(nome,))
        if nome == "executando":
            self.tabela.see(iid)
            self._status(f"Preenchendo linha {linha}...", "destaque")
        elif nome == "ok":
            self.concluidas += 1
            self.barra["value"] = self.concluidas
            self._status(f"{self.concluidas} de {len(self.registros)} concluídas", "ok")
        elif nome == "erro":
            self._status(f"Interrompido na linha {linha}", "erro")

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
        self.b_continuar.configure(state="normal" if retomar else "disabled",
                                   text=f"Continuar da linha {self._proxima_linha()}" if retomar else "Continuar")
        self.b_recarregar.configure(state="normal" if self.caminho_atual and not self.ocupado else "disabled")

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
        self.carregar(self.caminho_atual)

    def carregar(self, caminho):
        if self.ocupado:
            return
        self.tabela.delete(*self.tabela.get_children())
        self.registros, self.iids, self.estados, self.concluidas, self.resumo = [], {}, {}, 0, None
        self.caminho_atual = caminho
        self.barra.configure(value=0, maximum=1)
        try:
            registros = ler_planilha(caminho)
        except Exception as erro:
            messagebox.showerror("Planilha", str(erro))
            self._atualizar_botoes()
            return
        erros = validar(registros)
        for r in registros:
            self.iids[r["linha"]] = self.tabela.insert(
                "", "end", values=(r["linha"], r["item"], r["custo"], r["conta"] or "—", r["quantidade"], ESTADOS["pendente"]),
                tags=("pendente",))
        self.titulo_zona.configure(text=Path(caminho).name)
        self.sub_zona.configure(text="clique ou arraste outra planilha para trocar")
        self.resumo = (len(registros), sum(1 for r in registros if r["conta"]), len(erros))
        self._pintar(self.chip_linhas, f"{self.resumo[0]} linhas", "azul")
        self._pintar(self.chip_conta, f"{self.resumo[1]} com conta", "azul")
        self._pintar(self.chip_erros, f"{self.resumo[2]} erros", "vermelho" if erros else "verde")
        self.escrever(f"{len(registros)} linha(s) lidas de {caminho}")
        if erros:
            self._status("Corrija a planilha e clique em Recarregar", "erro")
            self.escrever("Planilha invalida, corrija no Excel, salve e clique em Recarregar:\n  " + "\n  ".join(erros))
            messagebox.showerror("Planilha inválida", "\n".join(erros[:10]) + "\n\nCorrija no Excel, salve e clique em 'Recarregar planilha'.")
        elif not registros:
            self._status("Planilha vazia", "aviso")
        else:
            self.registros = registros
            self.barra.configure(maximum=len(registros))
            self._status("Pronta para preencher", "ok")
            repetidos = itens_repetidos(registros)
            if repetidos:
                lista = "; ".join(f"{item} (linhas {', '.join(map(str, linhas))})" for item, linhas in list(repetidos.items())[:5])
                self.escrever(f"AVISO: {len(repetidos)} item(ns) repetido(s): {lista}{' ...' if len(repetidos) > 5 else ''}. "
                              "O SRP soma as quantidades de itens iguais na mesma linha, entao a conferencia da quantidade "
                              "vai interromper. Para esta planilha, desmarque 'Conferir quantidade na tela'.")
                self._status("Itens repetidos: veja o aviso", "aviso")
        self._atualizar_botoes()

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
        if self.conferir.get() and itens_repetidos(self.registros):
            aviso += ("\n\nATENÇÃO: a planilha tem itens repetidos e a conferência da quantidade está ligada. "
                      "Como o SRP soma itens iguais, a conferência vai interromper nesses itens.")
        if not messagebox.askokcancel("Iniciar", aviso):
            return
        self.parar.clear()
        if not retomar:
            self.concluidas = 0
            self.estados = {}
            for iid in self.iids.values():
                self.tabela.set(iid, "status", ESTADOS["pendente"])
                self.tabela.item(iid, tags=("pendente",))
        else:
            for r in registros:
                self._estado_linha(r["linha"], "pendente")
        self.barra.configure(value=self.concluidas, maximum=len(self.registros))
        self._status("Iniciando...", "destaque")
        self.ocupado = True
        self._atualizar_botoes()
        self.thread = threading.Thread(target=self._rodar, args=(registros, pausa, self.conferir.get()), daemon=True)
        self.thread.start()

    def _rodar(self, registros, pausa, conferir):
        try:
            executar(registros, "Desktop Viewer", pausa, 4.0, 5, self.parar, self.fila.put,
                     lambda linha, nome: self.fila.put((linha, nome)), conferir)
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
