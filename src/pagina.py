"""Base comum das abas (v1.14.0): o que a SNAC0234 (o proprio App) e a SNAC0236 (Pagina0236) faziam igual, cada uma
com a sua copia. Cada aba define `TELA` (nome nas falhas registradas), `CHAVE_CONSOLE` (preferencia dos Detalhes no
config.json) e `raiz` (a janela Tk: o App e a raiz dele mesmo) e tem `fila`, `saida`, `caixa_saida`, `b_console`,
`status`, `historico`, `console_aberto`, `dica`, `status_atual` e o `_tratar(item)` das mensagens do robo."""
import queue
import tkinter as tk
from tkinter import ttk

import visual
from core import registrar_falha_tela
from visual import COR, esp, fonte, px

# o robo procura a janela do Citrix por este trecho do titulo; render = espera maxima (s) pela tela do SRP depois de
# uma tecla; espera = contagem (s) antes de digitar, para o usuario clicar no campo
JANELA_SRP = "Desktop Viewer"
RENDER = 4.0
ESPERA = 5


class Pagina:
    TELA = ""
    CHAVE_CONSOLE = ""
    UNIDADE = ""  # "linha"/"nota": o que o F9 espera terminar
    raiz = None

    # ---------- fila do robo ----------
    def _drenar(self):
        """Trata as mensagens da thread do robo. Cada mensagem tem o seu try e o reagendamento fica no finally (ate a
        v1.13.4 uma excecao aqui parava o laco de vez, ou descartava a mensagem do fim e a aba ficava ocupada)."""
        try:
            while True:
                try:
                    item = self.fila.get_nowait()
                except queue.Empty:
                    break
                try:
                    self._tratar(item)
                except Exception:
                    registrar_falha_tela(f"mensagem da {self.TELA}: {item!r}"[:300])
        finally:
            self.raiz.after(100, self._drenar)

    def _tratar(self, item):
        """Mensagem da fila: texto (vai para os Detalhes) ou evento com nome, (nome, *dados), tratado por
        `_ao_<nome>(*dados)`: 'estado' (linha, estado), 'fim', 'parar'/'pausa' (F8/F9) e os proprios de cada aba.
        Evento desconhecido = AttributeError, registrado pelo `_drenar` (o laco segue)."""
        if isinstance(item, str):
            self.escrever(item)
            return
        nome, *dados = item
        getattr(self, f"_ao_{nome}")(*dados)

    def _ao_estado(self, linha, nome):
        self._estado(linha, nome)

    def _avisar_estado(self, linha, nome):
        """Callback `estado` do executar (roda na thread do robo): vira o evento 'estado' na fila."""
        self.fila.put(("estado", linha, nome))

    def _ao_parar(self):
        """F8 durante a execucao."""
        if self.ocupado:
            self.escrever("F8: parando.")
            self.parar.set()

    def _ao_pausa(self):
        """F9 durante a execucao: pausa ao fim da linha/nota atual ou, na pausa, continua."""
        if not self.ocupado:
            return
        if self.pausado:
            self.continuar()
        elif not self.pausar.is_set():
            self.pausar.set()
            self.escrever(f"F9: o robô pausa ao terminar a {self.UNIDADE} atual.")
            self._status(f"Pausando ao fim da {self.UNIDADE}...", "aviso")

    # ---------- status e Detalhes ----------
    def _status(self, texto, cor="suave"):
        self.status_atual = (texto, cor)
        self.status.configure(text=texto, fg=COR[cor])
        self._arrumar_andamento()

    def _escrever_console(self, texto):
        """Guarda no historico e escreve nos Detalhes com a cor do tipo; devolve o tipo (erro/aviso/ok/...)."""
        self.historico.append(texto)
        tipo = visual.tipo_mensagem(texto)
        self.saida.configure(state="normal")
        self.saida.insert("end", texto + "\n", tipo or ())
        self.saida.see("end")
        self.saida.configure(state="disabled")
        return tipo

    def abrir_console(self):
        """Abre os Detalhes sem mudar a preferencia salva (o Simular e os erros abrem sozinhos)."""
        if not self.console_aberto:
            self.console_aberto = True
            self._mostrar_console()

    def alternar_console(self):
        self.console_aberto = not self.console_aberto
        self.raiz.config_usuario[self.CHAVE_CONSOLE] = self.console_aberto
        self.raiz.salvar_config()
        self._mostrar_console()

    def _mostrar_console(self):
        """Mostra ou esconde a caixa dos Detalhes logo acima de `antes_do_console` (a barra dos botoes da aba)."""
        if self.console_aberto:
            if not self.caixa_saida.winfo_manager():
                self.caixa_saida.pack(side="bottom", fill="x", before=self.antes_do_console)
            self.saida.see("end")
        else:
            self.caixa_saida.pack_forget()
        self.b_console.configure(text="Detalhes ▾" if self.console_aberto else "Detalhes ▸")

    def copiar(self):
        self.raiz.clipboard_clear()
        self.raiz.clipboard_append("\n".join(self.historico[-2000:]))
        self._status("Detalhes copiados", "ok")

    # ---------- tabela ----------
    def _montar_tabela(self, cartao, colunas, estica=None):
        """Treeview das linhas da planilha dentro do cartao, com a barra de rolagem fina. `colunas` = (chave, titulo,
        largura em px a 96 dpi, ancora); `estica` = chave da unica coluna que estica (None = todas esticam)."""
        chaves = tuple(c[0] for c in colunas)
        tabela = ttk.Treeview(cartao, columns=chaves, show="headings", height=3)
        for chave, titulo, largura, ancora in colunas:
            tabela.heading(chave, text=titulo, anchor=ancora)
            tabela.column(chave, width=px(largura), minwidth=px(50), anchor=ancora, stretch=estica in (None, chave))
        rolagem = self.raiz._rolagem(cartao, tabela.yview)
        tabela.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y", padx=(0, esp(1)), pady=esp(2))
        tabela.pack(side="left", fill="both", expand=True, padx=(esp(1), 0), pady=esp(1))  # margem: cantos redondos
        return tabela

    # ---------- dica ao passar o mouse ----------
    def _mostrar_dica(self, evento, texto):
        """Dica (tooltip) perto do mouse; com o mesmo texto ja aberto, so acompanha o mouse."""
        if self.dica and getattr(self.dica, "texto", None) == texto:
            self.dica.geometry(f"+{evento.x_root + 14}+{evento.y_root + 12}")
            return
        self._esconder_dica()
        self.dica = tk.Toplevel(self.raiz)
        self.dica.texto = texto
        self.dica.overrideredirect(True)
        tk.Label(self.dica, text=texto, bg=COR["texto"], fg=COR["fundo"], font=fonte("pequeno"), padx=esp(2),
                 pady=esp(1), justify="left").pack()
        self.dica.geometry(f"+{evento.x_root + 14}+{evento.y_root + 12}")

    def _esconder_dica(self):
        if getattr(self, "dica", None):
            self.dica.destroy()
            self.dica = None

    def _linha_sob_o_mouse(self, evento):
        iid = self.tabela.identify_row(evento.y)
        return next((linha for linha, i in self.iids.items() if i == iid), None) if iid else None
