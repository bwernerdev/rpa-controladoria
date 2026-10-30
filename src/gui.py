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
        "fundo": "#f3f5f9", "cartao": "#ffffff", "borda": "#dfe4ec", "texto": "#1b2433", "suave": "#667085",
        "topo": "#0f172a", "destaque": "#2563eb", "destaque_hover": "#1d4ed8", "ok": "#15803d", "ok_fundo": "#dcfce7",
        "erro": "#b91c1c", "erro_fundo": "#fee2e2", "aviso": "#b45309", "aviso_fundo": "#fef3c7",
        "azul_fundo": "#dbeafe", "cinza_fundo": "#eef1f6", "console": "#0f172a", "console_texto": "#e2e8f0",
    },
    "escuro": {
        "fundo": "#0b1220", "cartao": "#111a2e", "borda": "#223049", "texto": "#e5e9f2", "suave": "#8b97ad",
        "topo": "#050a14", "destaque": "#3b82f6", "destaque_hover": "#2563eb", "ok": "#4ade80", "ok_fundo": "#10331f",
        "erro": "#f87171", "erro_fundo": "#3a1517", "aviso": "#fbbf24", "aviso_fundo": "#3a2a0a",
        "azul_fundo": "#14264a", "cinza_fundo": "#18233a", "console": "#050a14", "console_texto": "#cbd5e1",
    },
}
COR = dict(PALETAS["claro"])
FONTE = "Segoe UI"
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


class App(RaizTk):
    def __init__(self):
        super().__init__()
        self.config_usuario = ler_config()
        self.tema = self.config_usuario.get("tema", "claro") if self.config_usuario.get("tema") in PALETAS else "claro"
        COR.update(PALETAS[self.tema])
        self.title(f"{NOME} v{VERSAO}")
        self.geometry("900x800")
        self.minsize(780, 680)
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
        for nome, fundo, hover, texto in (
            ("Primario", COR["destaque"], COR["destaque_hover"], "#ffffff"),
            ("Perigo", "#dc2626", "#b91c1c", "#ffffff"),
            ("Secundario", COR["cartao"], COR["cinza_fundo"], COR["texto"]),
        ):
            estilo.configure(f"{nome}.TButton", font=(FONTE, 10, "bold"), padding=(16, 9), background=fundo,
                             foreground=texto, borderwidth=1, bordercolor=COR["borda"] if nome == "Secundario" else fundo,
                             focusthickness=0, relief="flat")
            estilo.map(f"{nome}.TButton", background=[("disabled", COR["cinza_fundo"]), ("active", hover)],
                       foreground=[("disabled", COR["suave"])], bordercolor=[("disabled", COR["borda"])])
        estilo.configure("Link.TButton", font=(FONTE, 9), padding=(8, 4), background=COR["fundo"],
                         foreground=COR["destaque"], borderwidth=0, relief="flat")
        estilo.map("Link.TButton", foreground=[("active", COR["destaque_hover"]), ("disabled", COR["suave"])],
                   background=[("active", COR["fundo"])])
        estilo.configure("Topo.TButton", font=(FONTE, 9), padding=(8, 4), background=COR["topo"], foreground="#93c5fd",
                         borderwidth=0, relief="flat", focusthickness=0)
        estilo.map("Topo.TButton", foreground=[("active", "#ffffff")], background=[("active", COR["topo"])])
        estilo.configure("Treeview", font=(FONTE, 10), rowheight=28, background=COR["cartao"], fieldbackground=COR["cartao"],
                         foreground=COR["texto"], borderwidth=0, lightcolor=COR["cartao"], darkcolor=COR["cartao"],
                         bordercolor=COR["cartao"])
        estilo.configure("Treeview.Heading", font=(FONTE, 9, "bold"), background=COR["cinza_fundo"], foreground=COR["suave"],
                         relief="flat", padding=(8, 7))
        estilo.map("Treeview", background=[("selected", COR["azul_fundo"])], foreground=[("selected", COR["texto"])])
        estilo.configure("Barra.Horizontal.TProgressbar", troughcolor=COR["borda"], background=COR["destaque"],
                         borderwidth=0, thickness=8)
        estilo.configure("Vertical.TScrollbar", background=COR["borda"], troughcolor=COR["cartao"], bordercolor=COR["cartao"],
                         arrowcolor=COR["suave"], relief="flat")
        estilo.configure("TEntry", padding=5, fieldbackground=COR["cartao"], foreground=COR["texto"], bordercolor=COR["borda"],
                         insertcolor=COR["texto"])
        estilo.configure("TCheckbutton", background=COR["fundo"], foreground=COR["texto"], font=(FONTE, 9))
        estilo.map("TCheckbutton", background=[("active", COR["fundo"])])

    def _cartao(self, pai):
        return tk.Frame(pai, bg=COR["cartao"], highlightbackground=COR["borda"], highlightthickness=1)

    def _chip(self, pai):
        rotulo = tk.Label(pai, text="", bg=COR["cinza_fundo"], fg=COR["suave"], font=(FONTE, 9, "bold"), padx=10, pady=3)
        rotulo.pack(side="left", padx=(0, 8))
        return rotulo

    def _montar(self):
        topo = tk.Frame(self, bg=COR["topo"])
        topo.pack(fill="x")
        titulos = tk.Frame(topo, bg=COR["topo"])
        titulos.pack(side="left")
        tk.Label(titulos, text=NOME, bg=COR["topo"], fg="#ffffff", font=(FONTE, 17, "bold")).pack(anchor="w", padx=22, pady=(14, 0))
        tk.Label(titulos, text="Preenchimento automático de requisições a partir da planilha", bg=COR["topo"],
                 fg="#94a3b8", font=(FONTE, 10)).pack(anchor="w", padx=22, pady=(0, 14))
        lado = tk.Frame(topo, bg=COR["topo"])
        lado.pack(side="right", padx=22)
        self.chip_citrix = tk.Label(lado, text="Citrix: verificando...", bg=COR["topo"], fg="#94a3b8", font=(FONTE, 9, "bold"))
        self.chip_citrix.pack(anchor="e")
        ttk.Button(lado, text="Tema claro" if self.tema == "escuro" else "Tema escuro", style="Topo.TButton",
                   command=self.alternar_tema).pack(anchor="e", pady=(2, 0))

        corpo = tk.Frame(self, bg=COR["fundo"])
        corpo.pack(fill="both", expand=True, padx=22, pady=14)

        self.zona = tk.Frame(corpo, bg=COR["cartao"], highlightbackground=COR["destaque"], highlightthickness=1, cursor="hand2")
        self.zona.pack(fill="x")
        self.titulo_zona = tk.Label(self.zona, text="Arraste a planilha aqui", bg=COR["cartao"], fg=COR["texto"],
                                    font=(FONTE, 13, "bold"), cursor="hand2")
        self.titulo_zona.pack(pady=(14, 0))
        self.sub_zona = tk.Label(self.zona, text="ou clique para escolher o arquivo (.xlsx)", bg=COR["cartao"],
                                 fg=COR["suave"], font=(FONTE, 10), cursor="hand2")
        self.sub_zona.pack(pady=(2, 14))
        for parte in (self.zona, self.titulo_zona, self.sub_zona):
            parte.bind("<Button-1>", lambda _e: self.escolher())
        if DND_FILES:
            self.zona.drop_target_register(DND_FILES)
            self.zona.dnd_bind("<<Drop>>", self._soltou)

        chips = tk.Frame(corpo, bg=COR["fundo"])
        chips.pack(fill="x", pady=(12, 8))
        self.chip_linhas, self.chip_conta, self.chip_erros = self._chip(chips), self._chip(chips), self._chip(chips)
        ttk.Button(chips, text="Abrir pasta de logs", style="Link.TButton", command=lambda: os.startfile(pasta_logs())).pack(side="right")
        self.b_recarregar = ttk.Button(chips, text="Recarregar planilha", style="Link.TButton", command=self.recarregar)
        self.b_recarregar.pack(side="right")
        opcoes = tk.Frame(corpo, bg=COR["fundo"])
        opcoes.pack(fill="x", pady=(0, 8))
        ttk.Checkbutton(opcoes, text="Conferir quantidade na tela", variable=self.conferir).pack(side="left", padx=(0, 18))
        tk.Label(opcoes, text="Limite de linhas", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9)).pack(side="left")
        ttk.Entry(opcoes, textvariable=self.limite, width=5, justify="center").pack(side="left", padx=(6, 14))
        tk.Label(opcoes, text="Pausa (s)", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9)).pack(side="left")
        ttk.Entry(opcoes, textvariable=self.pausa, width=5, justify="center").pack(side="left", padx=6)

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
        self.tabela.tag_configure("executando", background=COR["azul_fundo"], foreground=COR["destaque"])
        self.tabela.tag_configure("pendente", foreground=COR["texto"])

        progresso = tk.Frame(corpo, bg=COR["fundo"])
        progresso.pack(fill="x", pady=(12, 0))
        self.barra = ttk.Progressbar(progresso, style="Barra.Horizontal.TProgressbar", mode="determinate")
        self.barra.pack(side="left", fill="x", expand=True)
        self.status = tk.Label(progresso, text="", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 10, "bold"))
        self.status.pack(side="left", padx=(12, 0))

        botoes = tk.Frame(corpo, bg=COR["fundo"])
        botoes.pack(fill="x", pady=12)
        self.b_iniciar = ttk.Button(botoes, text="Iniciar preenchimento", style="Primario.TButton", command=self.iniciar)
        self.b_iniciar.pack(side="left")
        self.b_continuar = ttk.Button(botoes, text="Continuar", style="Secundario.TButton", command=lambda: self.iniciar(True))
        self.b_continuar.pack(side="left", padx=8)
        self.b_simular = ttk.Button(botoes, text="Simular", style="Secundario.TButton", command=self.simular)
        self.b_simular.pack(side="left")
        self.b_parar = ttk.Button(botoes, text="Parar", style="Perigo.TButton", command=self.parar.set)
        self.b_parar.pack(side="left", padx=8)

        self.saida = tk.Text(corpo, height=6, state="disabled", wrap="word", bg=COR["console"], fg=COR["console_texto"],
                             font=("Consolas", 9), relief="flat", padx=12, pady=10, insertbackground="#ffffff")
        self.saida.pack(fill="x")
        tk.Label(corpo, text=f"{NOME} v{VERSAO}", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 8)).pack(anchor="e", pady=(6, 0))

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
        fundo, cor = {"azul": (COR["azul_fundo"], COR["destaque"]), "verde": (COR["ok_fundo"], COR["ok"]),
                      "vermelho": (COR["erro_fundo"], COR["erro"]), "neutro": (COR["cinza_fundo"], COR["suave"])}[tipo]
        chip.configure(text=texto, bg=fundo, fg=cor)

    def _pintar_citrix(self):
        if self.citrix is None:
            self.chip_citrix.configure(text="Citrix: verificando...", fg="#94a3b8")
        elif self.citrix:
            self.chip_citrix.configure(text="● Citrix conectado", fg="#4ade80")
        else:
            self.chip_citrix.configure(text="● Citrix não encontrado - abra o Desktop Viewer", fg="#f87171")

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


if __name__ == "__main__":
    App().mainloop()
