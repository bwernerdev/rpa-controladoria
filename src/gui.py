"""Janela do robo do SRP: escolher planilha, simular e preencher."""
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import descrever, executar, ler_planilha, validar


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Robo SRP")
        self.geometry("820x560")
        self.minsize(640, 440)
        self.registros = []
        self.fila = queue.Queue()
        self.parar = threading.Event()
        self.thread = None
        self.caminho = tk.StringVar()
        self.limite = tk.StringVar(value="0")
        self.pausa = tk.StringVar(value="0.25")
        self._montar()
        self.after(100, self._drenar)

    def _montar(self):
        topo = ttk.Frame(self, padding=8)
        topo.pack(fill="x")
        ttk.Button(topo, text="Escolher planilha...", command=self.escolher).pack(side="left")
        ttk.Label(topo, textvariable=self.caminho).pack(side="left", padx=8)

        opcoes = ttk.Frame(self, padding=(8, 0))
        opcoes.pack(fill="x")
        ttk.Label(opcoes, text="Limite de linhas (0 = todas):").pack(side="left")
        ttk.Entry(opcoes, textvariable=self.limite, width=6).pack(side="left", padx=(4, 16))
        ttk.Label(opcoes, text="Pausa entre teclas (s):").pack(side="left")
        ttk.Entry(opcoes, textvariable=self.pausa, width=6).pack(side="left", padx=4)

        colunas = ("linha", "item", "custo", "conta", "quantidade")
        self.tabela = ttk.Treeview(self, columns=colunas, show="headings", height=9)
        for coluna, titulo, largura in zip(colunas, ("Linha", "Item", "C/Custo", "Conta", "Quantidade"), (60, 140, 120, 120, 120)):
            self.tabela.heading(coluna, text=titulo)
            self.tabela.column(coluna, width=largura, anchor="center")
        self.tabela.pack(fill="x", padx=8, pady=8)

        botoes = ttk.Frame(self, padding=(8, 0))
        botoes.pack(fill="x")
        self.b_simular = ttk.Button(botoes, text="Simular", command=self.simular, state="disabled")
        self.b_simular.pack(side="left")
        self.b_iniciar = ttk.Button(botoes, text="Iniciar", command=self.iniciar, state="disabled")
        self.b_iniciar.pack(side="left", padx=8)
        self.b_parar = ttk.Button(botoes, text="Parar", command=self.parar.set, state="disabled")
        self.b_parar.pack(side="left")

        self.saida = tk.Text(self, height=12, state="disabled", wrap="word")
        self.saida.pack(fill="both", expand=True, padx=8, pady=8)

    def escrever(self, texto):
        self.saida.configure(state="normal")
        self.saida.insert("end", texto + "\n")
        self.saida.see("end")
        self.saida.configure(state="disabled")

    def _drenar(self):
        try:
            while True:
                item = self.fila.get_nowait()
                if item is None:
                    self._ocupado(False)
                else:
                    self.escrever(item)
        except queue.Empty:
            pass
        self.after(100, self._drenar)

    def _ocupado(self, ocupado):
        self.b_iniciar.configure(state="disabled" if ocupado or not self.registros else "normal")
        self.b_simular.configure(state="disabled" if ocupado or not self.registros else "normal")
        self.b_parar.configure(state="normal" if ocupado else "disabled")

    def escolher(self):
        caminho = filedialog.askopenfilename(title="Planilha do SRP", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not caminho:
            return
        self.caminho.set(caminho)
        self.tabela.delete(*self.tabela.get_children())
        self.registros = []
        try:
            registros = ler_planilha(caminho)
        except Exception as erro:
            messagebox.showerror("Planilha", str(erro))
            self._ocupado(False)
            return
        erros = validar(registros)
        for r in registros:
            self.tabela.insert("", "end", values=(r["linha"], r["item"], r["custo"], r["conta"] or "(vazia)", r["quantidade"]))
        self.escrever(f"{len(registros)} linha(s) lidas de {caminho}")
        if erros:
            self.escrever("Planilha invalida, corrija antes de iniciar:\n  " + "\n  ".join(erros))
            messagebox.showerror("Planilha invalida", "\n".join(erros[:10]))
        elif not registros:
            self.escrever("Nenhuma linha para processar.")
        else:
            self.registros = registros
        self._ocupado(False)

    def selecionadas(self):
        try:
            limite = int(self.limite.get() or 0)
            pausa = float(self.pausa.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Opcoes", "Limite deve ser inteiro e pausa deve ser numero.")
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
        if not messagebox.askokcancel("Iniciar", f"Preencher {len(registros)} linha(s) no SRP?\nApos clicar em OK voce tera 5 segundos para clicar no campo Item da primeira linha livre."):
            return
        self.parar.clear()
        self._ocupado(True)
        self.thread = threading.Thread(
            target=self._rodar, args=(registros, pausa), daemon=True)
        self.thread.start()

    def _rodar(self, registros, pausa):
        try:
            executar(registros, "Desktop Viewer", pausa, 4.0, 5, self.parar, self.fila.put)
        finally:
            self.fila.put(None)


if __name__ == "__main__":
    App().mainloop()
