"""Janela do robo do SRP: arrastar a planilha, simular e preencher."""
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core import caminho_tesseract, descrever, executar, ler_planilha, pasta_logs, validar

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    RaizTk = TkinterDnD.Tk
except Exception:  # sem arrastar e soltar: o botao de escolher continua funcionando
    DND_FILES, RaizTk = None, tk.Tk

COR = {
    "fundo": "#f3f5f9", "cartao": "#ffffff", "borda": "#dfe4ec", "texto": "#1b2433", "suave": "#667085",
    "topo": "#0f172a", "destaque": "#2563eb", "destaque_hover": "#1d4ed8", "ok": "#15803d", "ok_fundo": "#dcfce7",
    "erro": "#b91c1c", "erro_fundo": "#fee2e2", "aviso": "#b45309", "aviso_fundo": "#fef3c7",
    "azul_fundo": "#dbeafe", "cinza_fundo": "#eef1f6", "console": "#0f172a", "console_texto": "#e2e8f0",
}
FONTE = "Segoe UI"
ESTADOS = {"pendente": "Pendente", "executando": "Em andamento", "ok": "Concluída", "erro": "Erro"}


class App(RaizTk):
    def __init__(self):
        super().__init__()
        self.title("Robô SRP")
        self.geometry("900x760")
        self.minsize(760, 620)
        self.configure(bg=COR["fundo"])
        self.registros = []
        self.iids = {}
        self.fila = queue.Queue()
        self.parar = threading.Event()
        self.thread = None
        self.concluidas = 0
        self.limite = tk.StringVar(value="0")
        self.pausa = tk.StringVar(value="0.25")
        self.conferir = tk.BooleanVar(value=caminho_tesseract() is not None)
        self._estilos()
        self._montar()
        if DND_FILES:
            for alvo in (self, self.zona):
                alvo.drop_target_register(DND_FILES)
                alvo.dnd_bind("<<Drop>>", self._soltou)
        self.after(100, self._drenar)

    # ---------- visual ----------
    def _estilos(self):
        estilo = ttk.Style(self)
        estilo.theme_use("clam")
        estilo.configure(".", font=(FONTE, 10), background=COR["fundo"], foreground=COR["texto"])
        for nome, fundo, hover, texto in (
            ("Primario", COR["destaque"], COR["destaque_hover"], "#ffffff"),
            ("Perigo", "#dc2626", "#b91c1c", "#ffffff"),
            ("Secundario", COR["cartao"], COR["cinza_fundo"], COR["texto"]),
        ):
            estilo.configure(f"{nome}.TButton", font=(FONTE, 10, "bold"), padding=(18, 9), background=fundo,
                             foreground=texto, borderwidth=1, bordercolor=COR["borda"] if nome == "Secundario" else fundo,
                             focusthickness=0, relief="flat")
            estilo.map(f"{nome}.TButton", background=[("disabled", COR["cinza_fundo"]), ("active", hover)],
                       foreground=[("disabled", "#9aa4b5")], bordercolor=[("disabled", COR["borda"])])
        estilo.configure("Link.TButton", font=(FONTE, 9), padding=(8, 4), background=COR["fundo"],
                         foreground=COR["destaque"], borderwidth=0, relief="flat")
        estilo.map("Link.TButton", foreground=[("active", COR["destaque_hover"])], background=[("active", COR["fundo"])])
        estilo.configure("Treeview", font=(FONTE, 10), rowheight=28, background=COR["cartao"], fieldbackground=COR["cartao"],
                         foreground=COR["texto"], borderwidth=0)
        estilo.configure("Treeview.Heading", font=(FONTE, 9, "bold"), background=COR["cinza_fundo"], foreground=COR["suave"],
                         relief="flat", padding=(8, 7))
        estilo.map("Treeview", background=[("selected", COR["azul_fundo"])], foreground=[("selected", COR["texto"])])
        estilo.configure("Barra.Horizontal.TProgressbar", troughcolor=COR["borda"], background=COR["destaque"],
                         borderwidth=0, thickness=8)
        estilo.configure("Vertical.TScrollbar", background=COR["borda"], troughcolor=COR["cartao"], bordercolor=COR["cartao"],
                         arrowcolor=COR["suave"], relief="flat")
        estilo.configure("TEntry", padding=5, fieldbackground=COR["cartao"], bordercolor=COR["borda"])

    def _cartao(self, pai, **opcoes):
        return tk.Frame(pai, bg=COR["cartao"], highlightbackground=COR["borda"], highlightthickness=1, **opcoes)

    def _chip(self, pai, texto, fundo, cor):
        rotulo = tk.Label(pai, text=texto, bg=fundo, fg=cor, font=(FONTE, 9, "bold"), padx=10, pady=3)
        rotulo.pack(side="left", padx=(0, 8))
        return rotulo

    def _montar(self):
        topo = tk.Frame(self, bg=COR["topo"])
        topo.pack(fill="x")
        tk.Label(topo, text="Robô SRP", bg=COR["topo"], fg="#ffffff", font=(FONTE, 17, "bold")).pack(anchor="w", padx=22, pady=(14, 0))
        tk.Label(topo, text="Preenchimento automático de requisições a partir da planilha", bg=COR["topo"],
                 fg="#94a3b8", font=(FONTE, 10)).pack(anchor="w", padx=22, pady=(0, 14))

        corpo = tk.Frame(self, bg=COR["fundo"])
        corpo.pack(fill="both", expand=True, padx=22, pady=16)

        self.zona = tk.Frame(corpo, bg=COR["cartao"], highlightbackground=COR["destaque"], highlightthickness=1, cursor="hand2")
        self.zona.pack(fill="x")
        self.titulo_zona = tk.Label(self.zona, text="Arraste a planilha aqui", bg=COR["cartao"], fg=COR["texto"],
                                    font=(FONTE, 13, "bold"), cursor="hand2")
        self.titulo_zona.pack(pady=(16, 0))
        self.sub_zona = tk.Label(self.zona, text="ou clique para escolher o arquivo (.xlsx)", bg=COR["cartao"],
                                 fg=COR["suave"], font=(FONTE, 10), cursor="hand2")
        self.sub_zona.pack(pady=(2, 16))
        for parte in (self.zona, self.titulo_zona, self.sub_zona):
            parte.bind("<Button-1>", lambda _e: self.escolher())

        chips = tk.Frame(corpo, bg=COR["fundo"])
        chips.pack(fill="x", pady=(12, 8))
        self.chip_linhas = self._chip(chips, "0 linhas", COR["cinza_fundo"], COR["suave"])
        self.chip_conta = self._chip(chips, "0 com conta", COR["cinza_fundo"], COR["suave"])
        self.chip_erros = self._chip(chips, "0 erros", COR["cinza_fundo"], COR["suave"])
        opcoes = tk.Frame(chips, bg=COR["fundo"])
        opcoes.pack(side="right")
        tk.Label(opcoes, text="Limite de linhas", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9)).pack(side="left")
        ttk.Entry(opcoes, textvariable=self.limite, width=5, justify="center").pack(side="left", padx=(6, 14))
        tk.Label(opcoes, text="Pausa (s)", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9)).pack(side="left")
        ttk.Entry(opcoes, textvariable=self.pausa, width=5, justify="center").pack(side="left", padx=(6, 14))
        ttk.Checkbutton(opcoes, text="Conferir quantidade na tela", variable=self.conferir).pack(side="left")

        cartao = self._cartao(corpo)
        cartao.pack(fill="both", expand=True)
        colunas = ("linha", "item", "custo", "conta", "quantidade", "status")
        self.tabela = ttk.Treeview(cartao, columns=colunas, show="headings", height=6)
        for coluna, titulo, largura in zip(colunas, ("Linha", "Item", "C/Custo", "Conta", "Quantidade", "Status"),
                                           (60, 130, 110, 110, 110, 140)):
            self.tabela.heading(coluna, text=titulo)
            self.tabela.column(coluna, width=largura, anchor="center")
        rolagem = ttk.Scrollbar(cartao, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.tabela.pack(side="left", fill="both", expand=True)
        self.tabela.tag_configure("ok", background=COR["ok_fundo"], foreground=COR["ok"])
        self.tabela.tag_configure("erro", background=COR["erro_fundo"], foreground=COR["erro"])
        self.tabela.tag_configure("executando", background=COR["azul_fundo"], foreground=COR["destaque_hover"])
        self.tabela.tag_configure("pendente", foreground=COR["texto"])

        progresso = tk.Frame(corpo, bg=COR["fundo"])
        progresso.pack(fill="x", pady=(12, 0))
        self.barra = ttk.Progressbar(progresso, style="Barra.Horizontal.TProgressbar", mode="determinate")
        self.barra.pack(side="left", fill="x", expand=True)
        self.status = tk.Label(progresso, text="Aguardando planilha", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 10, "bold"))
        self.status.pack(side="left", padx=(12, 0))

        botoes = tk.Frame(corpo, bg=COR["fundo"])
        botoes.pack(fill="x", pady=12)
        self.b_iniciar = ttk.Button(botoes, text="Iniciar preenchimento", style="Primario.TButton", command=self.iniciar, state="disabled")
        self.b_iniciar.pack(side="left")
        self.b_simular = ttk.Button(botoes, text="Simular", style="Secundario.TButton", command=self.simular, state="disabled")
        self.b_simular.pack(side="left", padx=8)
        self.b_parar = ttk.Button(botoes, text="Parar", style="Perigo.TButton", command=self.parar.set, state="disabled")
        self.b_parar.pack(side="left")
        ttk.Button(botoes, text="Abrir pasta de logs", style="Link.TButton", command=lambda: os.startfile(pasta_logs())).pack(side="right")

        self.saida = tk.Text(corpo, height=6, state="disabled", wrap="word", bg=COR["console"], fg=COR["console_texto"],
                             font=("Consolas", 9), relief="flat", padx=12, pady=10, insertbackground="#ffffff")
        self.saida.pack(fill="x")

    # ---------- comportamento ----------
    def escrever(self, texto):
        self.saida.configure(state="normal")
        self.saida.insert("end", texto + "\n")
        self.saida.see("end")
        self.saida.configure(state="disabled")

    def _status(self, texto, cor="suave"):
        self.status.configure(text=texto, fg=COR[cor])

    def _drenar(self):
        try:
            while True:
                item = self.fila.get_nowait()
                if item is None:
                    self._ocupado(False)
                elif isinstance(item, tuple):
                    self._estado_linha(*item)
                else:
                    self.escrever(item)
        except queue.Empty:
            pass
        self.after(100, self._drenar)

    def _estado_linha(self, linha, nome):
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
            self._status(f"{self.concluidas} de {int(self.barra['maximum'])} concluídas", "ok")
        elif nome == "erro":
            self._status(f"Interrompido na linha {linha}", "erro")

    def _ocupado(self, ocupado):
        pronto = bool(self.registros) and not ocupado
        self.b_iniciar.configure(state="normal" if pronto else "disabled")
        self.b_simular.configure(state="normal" if pronto else "disabled")
        self.b_parar.configure(state="normal" if ocupado else "disabled")
        if not ocupado and self.concluidas and self.concluidas == int(self.barra["maximum"]):
            self._status("Tudo concluído", "ok")

    def _soltou(self, evento):
        arquivos = self.tk.splitlist(evento.data)
        if arquivos:
            self.carregar(arquivos[0])

    def escolher(self):
        caminho = filedialog.askopenfilename(title="Planilha do SRP", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if caminho:
            self.carregar(caminho)

    def carregar(self, caminho):
        if self.thread and self.thread.is_alive():
            return
        self.tabela.delete(*self.tabela.get_children())
        self.registros, self.iids, self.concluidas = [], {}, 0
        self.barra.configure(value=0, maximum=1)
        try:
            registros = ler_planilha(caminho)
        except Exception as erro:
            messagebox.showerror("Planilha", str(erro))
            self._ocupado(False)
            return
        erros = validar(registros)
        for r in registros:
            self.iids[r["linha"]] = self.tabela.insert(
                "", "end", values=(r["linha"], r["item"], r["custo"], r["conta"] or "—", r["quantidade"], ESTADOS["pendente"]), tags=("pendente",))
        self.titulo_zona.configure(text=Path(caminho).name)
        self.sub_zona.configure(text="clique ou arraste outra planilha para trocar")
        self.chip_linhas.configure(text=f"{len(registros)} linhas", bg=COR["azul_fundo"], fg=COR["destaque_hover"])
        com_conta = sum(1 for r in registros if r["conta"])
        self.chip_conta.configure(text=f"{com_conta} com conta", bg=COR["azul_fundo"], fg=COR["destaque_hover"])
        self.chip_erros.configure(text=f"{len(erros)} erros", bg=COR["erro_fundo"] if erros else COR["ok_fundo"],
                                  fg=COR["erro"] if erros else COR["ok"])
        self.escrever(f"{len(registros)} linha(s) lidas de {caminho}")
        if erros:
            self._status("Corrija a planilha", "erro")
            self.escrever("Planilha invalida, corrija antes de iniciar:\n  " + "\n  ".join(erros))
            messagebox.showerror("Planilha inválida", "\n".join(erros[:10]))
        elif not registros:
            self._status("Planilha vazia", "aviso")
        else:
            self.registros = registros
            self.barra.configure(maximum=len(registros))
            self._status("Pronta para preencher", "ok")
        self._ocupado(False)

    def selecionadas(self):
        try:
            limite = int(self.limite.get() or 0)
            pausa = float(self.pausa.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Opções", "O limite deve ser um número inteiro e a pausa deve ser um número.")
            return None
        return (self.registros[:limite] if limite > 0 else self.registros), pausa

    def simular(self):
        escolha = self.selecionadas()
        if escolha:
            for r in escolha[0]:
                self.escrever(f"linha {r['linha']}: {descrever(r)}")
            self.escrever("Simulacao: nada foi enviado. A Conta so e digitada se estiver editavel na tela.")

    def iniciar(self):
        escolha = self.selecionadas()
        if not escolha:
            return
        registros, pausa = escolha
        if not messagebox.askokcancel("Iniciar", f"Preencher {len(registros)} linha(s) no SRP?\n\nDepois do OK você terá 5 segundos para clicar no campo Item da primeira linha livre."):
            return
        self.parar.clear()
        self.concluidas = 0
        self.barra.configure(value=0, maximum=len(registros))
        for iid in self.iids.values():
            self.tabela.set(iid, "status", ESTADOS["pendente"])
            self.tabela.item(iid, tags=("pendente",))
        self._status("Iniciando...", "destaque")
        self._ocupado(True)
        self.thread = threading.Thread(target=self._rodar, args=(registros, pausa, self.conferir.get()), daemon=True)
        self.thread.start()

    def _rodar(self, registros, pausa, conferir):
        try:
            executar(registros, "Desktop Viewer", pausa, 4.0, 5, self.parar, self.fila.put,
                     lambda linha, nome: self.fila.put((linha, nome)), conferir)
        finally:
            self.fila.put(None)


if __name__ == "__main__":
    App().mainloop()
