"""Aba SNAC0236 (Digitacao Documento Recebimento) da janela. O estado fica neste objeto (sobrevive quando a janela e
redesenhada por tema ou monitor); `montar(pai)` refaz so os widgets. `ui` e o modulo gui (COR, fonte, px, ESTADOS)."""
import queue
import threading
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import snac0236
from core import (VERDE_CONCLUIDA, VERDE_SERIE_PONTO, VK_F8, atalhos_globais, citrix_aberto, fechar_no_excel,
                  pasta_logs, pintar_linhas)


class Pagina0236:
    def __init__(self, app, ui):
        self.app, self.ui = app, ui
        self.caminho, self.registros, self.notas, self.erros = None, [], [], []
        self.estados, self.historico = {}, []
        self.status_atual = ("Aguardando planilha", "suave")
        self.ocupado = self.pausado = False
        self.fila = queue.Queue()
        self.parar, self.retomar = threading.Event(), threading.Event()
        self.soltar_atalhos = None
        self.iids = {}
        # passo a passo: 0,5 s entre as acoes, narradas no console (para acompanhar os testes no SRP real)
        self.passo_a_passo = tk.BooleanVar(app, value=bool(app.config_usuario.get("passo_a_passo_0236", True)))
        self.passo_a_passo.trace_add("write", lambda *_: self._gravar_passo())
        self.execucao = None  # andamento da execucao em curso, para o tempo previsto
        # quantas NFs lancar nesta execucao (usuario, 06/11/2026): vazio = todas as pendentes; so numeros
        self.qtd_nfs = tk.StringVar(app, value="")
        self.qtd_nfs.trace_add("write", lambda *_: self._mudou_qtd())
        # linhas das notas concluidas pintadas de verde na planilha (usuario, 06/11/2026), numa thread so
        self.fila_pintar, self.pintadas, self.avisou_excel = queue.Queue(), set(), False
        # linhas das notas lancadas com "." na Serie (janela "ja existe"; usuario, 09/11/2026): outro verde e comentario
        self.serie_ponto = set()
        self.trava_pintura = threading.Lock()  # o pintor (thread) e o Conferir/Iniciar nao gravam ao mesmo tempo
        threading.Thread(target=self._pintor, daemon=True).start()
        app.after(100, self._drenar)
        app.after(1000, self._tique)

    def _gravar_passo(self):
        self.app.config_usuario["passo_a_passo_0236"] = self.passo_a_passo.get()
        self.ui.gravar_config(self.app.config_usuario)

    # ---------- montagem ----------
    def montar(self, pai):
        ui, COR, px, fonte = self.ui, self.ui.COR, self.ui.px, self.ui.fonte
        self.zona = tk.Frame(pai, bg=COR["cartao"], highlightbackground=COR["borda_forte"], highlightcolor=COR["borda_forte"],
                             highlightthickness=1)
        self.zona.pack(fill="x")
        if ui.DND_FILES:
            self.zona.drop_target_register(ui.DND_FILES)
            self.zona.dnd_bind("<<Drop>>", self.app._soltou)

        rodape = tk.Frame(pai, bg=COR["fundo"])
        rodape.pack(side="bottom", fill="x", pady=(6, 0))
        self.app._link(rodape, "Copiar detalhes", self.copiar).pack(side="left")
        self.app._link(rodape, "Abrir pasta de logs", lambda: __import__("os").startfile(pasta_logs())).pack(
            side="left", padx=(px(14), 0))
        self.app._link(rodape, "Conferir com os logs", self.conferir).pack(side="left", padx=(px(14), 0))
        tk.Label(rodape, text=f"v{ui.VERSAO}", bg=COR["fundo"], fg=COR["suave"], font=fonte(8)).pack(side="right")
        caixa = self.app._cartao(pai)
        caixa.pack(side="bottom", fill="x")
        self.saida = tk.Text(caixa, height=6, state="disabled", wrap="word", bg=COR["console"], fg=COR["console_texto"],
                             font=fonte(9, mono=True), relief="flat", borderwidth=0, padx=px(10), pady=px(6),
                             highlightthickness=0)
        self.saida.pack(fill="x")
        for tipo, cor in (("erro", "erro"), ("aviso", "aviso"), ("srp", "destaque")):
            self.saida.tag_configure(tipo, foreground=COR[cor])

        # barra de baixo responsiva (usuario, 06/11/2026): botoes a esquerda, Parar a direita e as opcoes (Passo a
        # passo, NFs a lancar) no meio; se nao couberem na largura (janela estreita), descem para uma linha propria
        barra = tk.Frame(pai, bg=COR["fundo"])
        barra.pack(side="bottom", fill="x", pady=(12, 10))
        barra.columnconfigure(2, weight=1)
        botoes = tk.Frame(barra, bg=COR["fundo"])
        botoes.grid(row=0, column=0, sticky="w")
        self.b_iniciar = ttk.Button(botoes, text="Iniciar preenchimento", style="Primario.TButton", command=self.iniciar)
        self.b_iniciar.pack(side="left")
        self.b_continuar = ttk.Button(botoes, text="Continuar", style="Secundario.TButton", command=self.continuar)
        self.b_continuar.pack(side="left", padx=(8, 0))
        self.b_simular = ttk.Button(botoes, text="Simular", style="Secundario.TButton", command=self.simular)
        self.b_simular.pack(side="left", padx=(8, 0))
        opcoes = tk.Frame(barra, bg=COR["fundo"])
        self.c_passo = ttk.Checkbutton(opcoes, text="Passo a passo (0,5 s entre ações)", variable=self.passo_a_passo)
        self.c_passo.pack(side="left")
        # quantidade de NFs a lancar: so digitos (ate 4); vazio = todas. A linha do tempo previsto mostra
        # "para N de M nota(s)"
        tk.Label(opcoes, text="NFs a lançar:", bg=COR["fundo"], fg=COR["texto"], font=fonte(9)).pack(
            side="left", padx=(px(14), px(6)))
        so_numeros = (pai.register(lambda novo: novo == "" or (novo.isdigit() and len(novo) <= 4)), "%P")
        self.e_qtd = ttk.Entry(opcoes, textvariable=self.qtd_nfs, width=5, justify="center", font=fonte(10),
                               validate="key", validatecommand=so_numeros)
        self.e_qtd.pack(side="left")
        self.b_parar = ttk.Button(barra, text="Parar", style="Perigo.TButton", command=self.parar.set)
        self.b_parar.grid(row=0, column=3, sticky="e", padx=(px(14), 0))
        embaixo = [None]

        def arrumar(_evento=None):
            largura = barra.winfo_width()
            if largura <= 1:
                return
            precisa = (botoes.winfo_reqwidth() + opcoes.winfo_reqwidth() + self.b_parar.winfo_reqwidth()
                       + 2 * px(14) + px(8))
            descer = largura < precisa
            if descer == embaixo[0]:
                return
            embaixo[0] = descer
            if descer:
                opcoes.grid(row=1, column=0, columnspan=4, sticky="w", pady=(px(10), 0))
            else:
                opcoes.grid(row=0, column=1, sticky="w", padx=(px(14), 0), pady=0)
        opcoes.grid(row=0, column=1, sticky="w", padx=(px(14), 0))
        barra.bind("<Configure>", arrumar)
        botoes.bind("<Configure>", arrumar)  # o Iniciar muda de texto ("Continuar da linha N") e de largura

        progresso = tk.Frame(pai, bg=COR["fundo"])
        progresso.pack(side="bottom", fill="x", pady=(12, 0))
        self.status = tk.Label(progresso, text="", bg=COR["fundo"], fg=COR["suave"], font=fonte(9), anchor="w",
                               justify="left", wraplength=px(800))
        self.status.pack(fill="x")
        self.barra = ui.BarraFina(progresso)
        self.barra.pack(fill="x", pady=(4, 0))
        # tempo previsto para a conclusao (so durante a execucao; `_atualizar_previsao`, a cada segundo)
        self.previsao = tk.Label(progresso, text="", bg=COR["fundo"], fg=COR["texto"], font=fonte(9), anchor="w")
        self._atualizar_previsao()

        cartao = self.app._cartao(pai)
        cartao.pack(fill="both", expand=True, pady=(px(12), 0))
        colunas = ("linha", "nota", "item", "oper", "qtd", "unit", "status")
        self.tabela = ttk.Treeview(cartao, columns=colunas, show="headings", height=3)
        for coluna, titulo, largura, ancora in zip(
                colunas, ("Linha", "Nota (chave)", "Item", "Oper SPM", "Qtd", "Valor Unit", "Status"),
                tuple(px(v) for v in (56, 120, 100, 80, 80, 100, 140)), ("center", "w", "w", "w", "e", "e", "w")):
            self.tabela.heading(coluna, text=titulo, anchor=ancora)
            self.tabela.column(coluna, width=largura, minwidth=px(50), anchor=ancora, stretch=True)
        rolagem = ttk.Scrollbar(cartao, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        self.tabela.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
        self.tabela.tag_configure("nota_par", background=COR["zebra"])
        self.tabela.tag_configure("nota_impar", background=COR["cartao"])
        self.tabela.tag_configure("ok", background=COR["ok_fundo"], foreground=COR["ok"])
        self.tabela.tag_configure("erro", background=COR["erro_fundo"], foreground=COR["erro"])
        self.tabela.tag_configure("executando", background=COR["azul_fundo"], foreground=COR["texto"])
        self.tabela.tag_configure("pausado", background=COR["aviso_fundo"], foreground=COR["aviso"])
        self._repovoar()

    def _repovoar(self):
        self.iids = {}
        for indice, nota in enumerate(self.notas):
            zebra = "nota_par" if indice % 2 else "nota_impar"  # zebra por nota: os itens da mesma nota ficam juntos
            for posicao, r in enumerate(nota["itens"]):
                self.iids[r["linha"]] = self.tabela.insert("", "end", values=self._valores(nota, r, posicao),
                                                           tags=self._tags(r["linha"], zebra))
        self._pintar_zona()
        self.barra.configure(maximum=max(len(self.registros), 1), value=self._feitas())
        self._status(*self.status_atual)
        self.saida.configure(state="normal")
        self.saida.delete("1.0", "end")
        for texto in self.historico[-200:]:
            self.saida.insert("end", texto + "\n", self.ui.tipo_mensagem(texto) or ())
        self.saida.see("end")
        self.saida.configure(state="disabled")
        self._botoes()
        self._atualizar_previsao()

    def _valores(self, nota, r, posicao):
        chave = f"…{nota['chave'][-6:]}" if posicao == 0 else "  ↳"
        return (r["linha"], chave, r["item"], r["oper"], r["qtd"], r["unit"],
                self.ui.ESTADOS[self.estados.get(r["linha"], "pendente")])

    def _tags(self, linha, zebra):
        nome = self.estados.get(linha, "pendente")
        # concluida = fundo verde (usuario, 06/11/2026); a zebra so nas linhas sem fundo proprio (ela venceria)
        return (nome,) if nome in ("erro", "executando", "pausado", "ok") else (zebra, nome)

    def _pintar_zona(self):
        COR, px, fonte = self.ui.COR, self.ui.px, self.ui.fonte
        for filho in self.zona.winfo_children():
            filho.destroy()
        if not self.caminho:
            self.zona.configure(cursor="hand2", highlightbackground=COR["borda_forte"])
            titulo = tk.Label(self.zona, text="Arraste a planilha da SNAC0236 aqui", bg=COR["cartao"], fg=COR["texto"],
                              font=fonte(11, forte=True), cursor="hand2")
            titulo.pack(pady=(px(16), 0))
            sub = tk.Label(self.zona, text="ou clique para escolher (.xlsx com CHAVE NF, FILIAL FISCAL, ... ICMS TERCEIRO)",
                           bg=COR["cartao"], fg=COR["suave"], font=fonte(9), cursor="hand2")
            sub.pack(pady=(0, px(16)))
            for parte in (self.zona, titulo, sub):
                parte.bind("<Button-1>", lambda _e: self.escolher())
                parte.bind("<Enter>", lambda _e: self.zona.configure(highlightbackground=COR["destaque"]))
                parte.bind("<Leave>", lambda _e: self.zona.configure(highlightbackground=COR["borda_forte"]))
            return
        self.zona.configure(cursor="", highlightbackground=COR["borda"])
        for evento in ("<Button-1>", "<Enter>", "<Leave>"):
            self.zona.unbind(evento)
        acoes = tk.Frame(self.zona, bg=COR["cartao"])
        acoes.pack(side="right", padx=px(14))
        self.app._link(acoes, "Recarregar", self.recarregar, bg=COR["cartao"]).pack(side="right")
        self.app._link(acoes, "Trocar planilha", self.escolher, bg=COR["cartao"]).pack(side="right", padx=(0, px(14)))
        tk.Label(self.zona, text="▦", bg=COR["cartao"], fg=COR["erro"] if self.erros else COR["ok"],
                 font=fonte(20)).pack(side="left", padx=(px(14), px(8)))
        textos = tk.Frame(self.zona, bg=COR["cartao"])
        textos.pack(side="left", fill="x", expand=True, pady=px(8))
        tk.Label(textos, text=Path(self.caminho).name, bg=COR["cartao"], fg=COR["texto"], font=fonte(10, forte=True),
                 anchor="w").pack(fill="x")
        resumo = f"{len(self.notas)} nota(s)  ·  {len(self.registros)} item(ns)"
        if self.erros:
            resumo += f"  ·  {len(self.erros)} erro(s) na planilha"
        tk.Label(textos, text=resumo, bg=COR["cartao"], fg=COR["erro"] if self.erros else COR["suave"], font=fonte(9),
                 anchor="w").pack(fill="x")

    # ---------- planilha ----------
    def escolher(self):
        caminho = filedialog.askopenfilename(title="Planilha da SNAC0236", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if caminho:
            self.carregar(caminho)

    def recarregar(self):
        if self.caminho and not self.ocupado:
            self.carregar(self.caminho)

    def carregar(self, caminho):
        if self.ocupado:
            return
        try:
            registros, notas, erros = snac0236.carregar(caminho)
        except Exception as erro:
            messagebox.showerror("Planilha", f"Não consegui ler a planilha:\n{erro}")
            return
        self.caminho, self.registros, self.notas, self.erros, self.estados = caminho, registros, notas, erros, {}
        self.pintadas, self.serie_ponto = set(), set()
        self.escrever(f"Planilha {Path(caminho).name}: {len(notas)} nota(s), {len(registros)} item(ns).")
        # notas ja pintadas de verde na planilha = ja feitas: o robo pula e segue pelas outras (usuario, 06/11/2026)
        verdes, partidas = [], []
        for nota in notas:
            cores = [r.get("verde") for r in nota["itens"]]
            if all(cores):
                verdes.append(nota)
                for r in nota["itens"]:
                    self.estados[r["linha"]] = "ok"
                    self.pintadas.add(r["linha"])
            elif any(cores):
                partidas.append(nota)
        if verdes:
            pendentes = [n for n in notas if n not in verdes]
            self.escrever(f"{len(verdes)} nota(s) já pintada(s) de verde na planilha: o robô pula essas"
                          + (f" e começa pela linha {pendentes[0]['itens'][0]['linha']}." if pendentes else
                             " (todas as notas)."))
        for nota in partidas:
            linhas = ", ".join(str(r["linha"]) for r in nota["itens"] if r.get("verde"))
            self.escrever(f"Aviso: a nota …{nota['chave'][-6:]} tem só parte das linhas em verde ({linhas}); "
                          "o robô digita a nota inteira. Se ela já foi integrada, pinte todas as linhas dela de verde.")
        for erro in erros[:30]:
            self.escrever(f"Erro: {erro}")
        if len(erros) > 30:
            self.escrever(f"Erro: ... e mais {len(erros) - 30}.")
        if erros:
            self._status(f"{len(erros)} erro(s) na planilha: corrija no Excel e clique em Recarregar", "erro")
        elif not notas:
            self._status("A planilha não tem linhas preenchidas", "aviso")
        elif len(verdes) == len(notas):
            self._status("Todas as notas já estão em verde na planilha (concluídas)", "ok")
        else:
            self._status(f"Pronta: {len(notas) - len(verdes)} nota(s) a fazer. Clique em Simular para ver as teclas ou "
                         "Iniciar", "suave")
        for iid in self.tabela.get_children():
            self.tabela.delete(iid)
        self._repovoar()

    # ---------- execucao ----------
    def _feitas(self):
        return sum(1 for r in self.registros if self.estados.get(r["linha"]) == "ok")

    def _pendentes(self):
        return [n for n in self.notas if any(self.estados.get(r["linha"]) != "ok" for r in n["itens"])]

    def _limite_nfs(self):
        """Quantidade de NFs digitada (None = vazio ou 0 = todas)."""
        texto = self.qtd_nfs.get().strip()
        return int(texto) if texto.isdigit() and int(texto) > 0 else None

    def _lote(self):
        """Notas que o Iniciar vai lancar: as pendentes, so as primeiras `_limite_nfs` se houver numero."""
        pendentes = self._pendentes()
        limite = self._limite_nfs()
        return pendentes[:limite] if limite else pendentes

    def _mudou_qtd(self):
        try:
            self._atualizar_previsao()
        except (tk.TclError, AttributeError):
            pass  # aba ainda nao montada

    def _botoes(self):
        pronto = bool(self.notas) and not self.erros and not self.ocupado and not self.app.ocupado
        pendentes = self._pendentes()
        retomar = pronto and pendentes and len(pendentes) < len(self.notas)
        self.b_iniciar.configure(state="normal" if pronto and pendentes else "disabled",
                                 text=f"Continuar da linha {pendentes[0]['itens'][0]['linha']}" if retomar else "Iniciar preenchimento")
        self.b_simular.configure(state="normal" if pronto else "disabled")
        self.b_parar.configure(state="normal" if self.ocupado else "disabled")
        self.b_continuar.configure(state="normal" if self.ocupado and self.pausado else "disabled")
        self.c_passo.configure(state="disabled" if self.ocupado else "normal")
        self.e_qtd.configure(state="disabled" if self.ocupado else "normal")

    def simular(self):
        for nota in self.notas:
            self.escrever(f"Nota …{nota['chave'][-6:]}: {snac0236.descrever(nota)}")
        self._status("Simulação no console (nada foi digitado)", "suave")

    def iniciar(self):
        if not self._lote() or not self._pular_ja_integradas():
            return
        notas = self._lote()
        if not notas:
            return
        faltam = len(self._pendentes())
        if not citrix_aberto():
            messagebox.showwarning("Citrix não encontrado", "Não achei a janela do Citrix (Desktop Viewer).\nAbra a SNAC0236 e tente de novo.")
            return
        itens = sum(len(n["itens"]) for n in notas)
        if not messagebox.askokcancel(
                "Iniciar", f"Preencher {len(notas)} nota(s) ({itens} item(ns)) na SNAC0236, a partir da linha "
                           f"{notas[0]['itens'][0]['linha']}?"
                           + (f" Das {faltam} pendentes, ficam {faltam - len(notas)} para depois (NFs a lançar: "
                              f"{len(notas)})." if len(notas) < faltam else "") + "\n\nTempo previsto: "
                           f"{snac0236.texto_duracao(snac0236.custo_previsto(notas))} (sem contar as pausas; "
                           "a previsão se ajusta depois da 1ª nota).\n\nDepois do OK você terá 5 segundos para clicar no campo "
                           "Chave Acesso NF-e (vazio).\n\nO robô clica em Integrar no fim de cada nota e pausa para você "
                           "conferir; para a próxima, deixe o cursor na Chave e clique em Continuar.\n\nSe a planilha "
                           f"{Path(self.caminho).name} estiver aberta no Excel, ela será salva e fechada (só ela)."):
            return
        if not self._fechar_planilha_no_excel():
            return
        for n in notas:
            for r in n["itens"]:
                self._estado(r["linha"], "pendente")
        self.parar.clear()
        self.retomar.clear()
        self.ocupado = self.app.ocupado = True
        self.avisou_excel = False
        self.execucao = {"inicio": time.time(), "notas": notas, "feitas": 0, "gasto": 0.0, "tempo_pausado": 0.0,
                         "pausado_desde": None}
        self.app._atualizar_botoes()
        self._botoes()
        self._status("Iniciando...", "destaque")
        self._atualizar_previsao()
        self.soltar_atalhos, falhas = atalhos_globais({VK_F8: lambda: self.fila.put(("parar",))})
        self.escrever("Aviso: F8 ja esta em uso por outro programa; use o botao Parar ou o mouse no canto." if falhas
                      else "Atalho: F8 para.")
        intervalo = 0.5 if self.passo_a_passo.get() else 0
        if intervalo:
            self.escrever("Passo a passo ligado: 0,5 s antes de cada acao; a proxima acao aparece aqui e na linha de status.")
        threading.Thread(target=self._rodar, args=(notas, intervalo), daemon=True).start()

    def _fechar_planilha_no_excel(self):
        """Salva e fecha no Excel SO a planilha carregada no robo (usuario, 06/11/2026): assim as notas concluidas sao
        pintadas direto no arquivo. Se ela mudou no Excel depois de carregada, recarrega e nao inicia (o usuario
        confere). Devolve False para nao iniciar."""
        try:
            fechou = fechar_no_excel(self.caminho)
        except Exception as erro:  # nao impede o preenchimento: pinta no Excel aberto
            self.escrever(f"Aviso: não salvei nem fechei a planilha no Excel ({erro}). As notas concluídas serão "
                          "pintadas no Excel aberto.")
            return True
        if not fechou:
            return True
        self.escrever(f"Planilha {Path(self.caminho).name} salva e fechada no Excel.")
        try:
            registros, _, erros = snac0236.carregar(self.caminho)
        except Exception:
            registros, erros = None, True
        if registros == self.registros and not erros:
            return True
        self.carregar(self.caminho)
        # registros com a cor: linha pintada de verde no Excel antes do Iniciar tambem conta como alteracao
        messagebox.showwarning("Planilha alterada", "A planilha tinha alterações no Excel, que foram salvas. Recarreguei "
                                                    "a planilha: confira as linhas e clique em Iniciar de novo.")
        return False

    def _rodar(self, notas, intervalo):
        try:
            snac0236.executar(notas, "Desktop Viewer", 0.25, 4.0, 5, self.parar, self.fila.put,
                              lambda linha, nome: self.fila.put((linha, nome)), self.retomar, intervalo,
                              serie_ponto=lambda linhas: self.fila.put(("serie_ponto", tuple(linhas))))
        finally:
            self.fila.put(None)

    def continuar(self):
        if self.ocupado and self.pausado:
            self.escrever("Continuando: o robo volta ao Citrix e confere onde esta o cursor.")
            self.pausado = False
            self._botoes()
            self.retomar.set()

    def _drenar(self):
        try:
            while True:
                item = self.fila.get_nowait()
                if item is None:
                    if self.soltar_atalhos:
                        self.soltar_atalhos()
                        self.soltar_atalhos = None
                    self.ocupado = self.pausado = False
                    self.app.ocupado = False
                    self.app._atualizar_botoes()
                    self._fim()
                elif item == ("parar",):
                    self.escrever("F8: parando.")
                    self.parar.set()
                elif isinstance(item, tuple) and item[0] == "serie_ponto":
                    self.serie_ponto.update(item[1])
                elif isinstance(item, tuple):
                    self._estado(*item)
                else:
                    self.escrever(item)
        except queue.Empty:
            pass
        except tk.TclError:
            pass  # widgets sendo refeitos (tema/monitor): o resto sai na proxima volta
        self.app.after(100, self._drenar)

    # ---------- tempo previsto ----------
    def _tique(self):
        try:
            self._atualizar_previsao()
        except (tk.TclError, AttributeError):
            pass  # widgets sendo refeitos (tema/monitor) ou a aba ainda nao montada
        finally:
            self.app.after(1000, self._tique)

    def _tempo_util(self, agora=None):
        """Segundos desde o Iniciar, sem o tempo pausado (a previsao nao conta o tempo parado esperando o usuario)."""
        e, agora = self.execucao, agora or time.time()
        pausado = e["tempo_pausado"] + (agora - e["pausado_desde"] if e["pausado_desde"] else 0)
        return agora - e["inicio"] - pausado

    def _contar_execucao(self, nome):
        e = self.execucao
        if not e:
            return
        agora = time.time()
        if nome == "pausado" and not e["pausado_desde"]:
            e["pausado_desde"] = agora
        elif nome != "pausado" and e["pausado_desde"]:
            e["tempo_pausado"] += agora - e["pausado_desde"]
            e["pausado_desde"] = None
        feitas = sum(1 for n in e["notas"] if all(self.estados.get(r["linha"]) == "ok" for r in n["itens"]))
        if feitas > e["feitas"]:
            e["feitas"], e["gasto"] = feitas, self._tempo_util(agora)

    def _atualizar_previsao(self):
        e = self.execucao
        if not (e and self.ocupado):
            # antes de iniciar (planilha importada sem erros): previsao das notas que faltam (usuario, 06/11/2026)
            # com "NFs a lancar" preenchido: so as primeiras N pendentes (usuario, 06/11/2026)
            pendentes = self._lote() if self.notas and not self.erros and not self.ocupado else []
            if not pendentes:
                self.previsao.pack_forget()
                return
            itens = sum(len(n["itens"]) for n in pendentes)
            faltam = len(self._pendentes())
            termino = time.strftime("%H:%M", time.localtime(time.time() + snac0236.custo_previsto(pendentes)))
            self.previsao.configure(text=f"Tempo previsto: {snac0236.texto_duracao(snac0236.custo_previsto(pendentes))} "
                                         f"para {len(pendentes)}" + (f" de {faltam}" if len(pendentes) < faltam else "")
                                         + f" nota(s), {itens} item(ns)  ·  se iniciar agora, término "
                                         f"~{termino}  ·  sem contar as pausas")
            if not self.previsao.winfo_ismapped():
                self.previsao.pack(fill="x", pady=(4, 0))
            return
        notas, feitas = e["notas"], e["feitas"]
        util = self._tempo_util()
        texto = f"Nota {min(feitas + 1, len(notas))} de {len(notas)}  ·  decorrido {self.ui.duracao(time.time() - e['inicio'])}"
        if feitas >= len(notas):
            texto += "  ·  terminando…"
        else:
            # previsto para as que faltam menos o que a nota atual ja levou (a contagem desce a cada segundo)
            resta = snac0236.tempo_previsto(notas[:feitas], notas[feitas:], e["gasto"]) - (util - e["gasto"])
            resta = max(resta, snac0236.SEG_NOTA * 0.1 * (len(notas) - feitas))
            termino = time.strftime("%H:%M", time.localtime(time.time() + resta))
            texto += f"  ·  previsto: {snac0236.texto_duracao(resta)} restantes (término ~{termino})"
            if e["pausado_desde"]:
                texto += "  ·  pausado: o tempo parado não entra na conta"
            elif not feitas:
                texto += "  ·  estimativa inicial"
        self.previsao.configure(text=texto)
        if not self.previsao.winfo_ismapped():
            self.previsao.pack(fill="x", pady=(4, 0))

    def _fim(self):
        if self.execucao:
            e, self.execucao = self.execucao, None
            self.escrever(f"Tempo total: {self.ui.duracao(time.time() - e['inicio'])} ({e['feitas']} de "
                          f"{len(e['notas'])} nota(s) concluída(s)).")
        self._atualizar_previsao()
        feitas, total = self._feitas(), len(self.registros)
        if total and feitas == total:
            self._status("Tudo concluído", "ok")
        elif not self.status_atual[0].startswith(("Erro", "Interrompido", "Nao iniciei")):
            self._status(f"Parado: {feitas} de {total} item(ns) concluído(s)", "aviso")
        self._botoes()

    def _estado(self, linha, nome):
        self.estados[linha] = nome
        if nome == "pausado":
            self.pausado = True
        elif nome == "executando":
            self.pausado = False
        self._contar_execucao(nome)
        if nome == "ok" and self.ocupado:
            self._pintar_nota(linha)
        iid = self.iids.get(linha)
        if iid:
            indice = next((i for i, n in enumerate(self.notas) if any(r["linha"] == linha for r in n["itens"])), 0)
            self.tabela.set(iid, "status", self.ui.ESTADOS[nome])
            self.tabela.item(iid, tags=self._tags(linha, "nota_par" if indice % 2 else "nota_impar"))
            if nome in ("executando", "pausado"):
                self.tabela.see(iid)
        self.barra.configure(value=self._feitas())
        self._botoes()

    # ---------- linhas concluidas pintadas na planilha ----------
    def _pintar_nota(self, linha):
        """Nota da `linha` toda concluida: manda as linhas dela para o pintor (uma vez por nota)."""
        nota = next((n for n in self.notas if any(r["linha"] == linha for r in n["itens"])), None)
        if not nota or not self.caminho:
            return
        linhas = [r["linha"] for r in nota["itens"]]
        if all(self.estados.get(l) == "ok" for l in linhas) and not set(linhas) <= self.pintadas:
            self.pintadas.update(linhas)
            self.fila_pintar.put(self._pedido_pintura(nota))

    def _pedido_pintura(self, nota, quando=None):
        """(planilha, linhas, cor, comentarios) para o pintor. Nota com "." na Serie: verde mais escuro e comentario
        na CHAVE NF da 1a linha (usuario, 09/11/2026). `quando`: data do log em que isso aconteceu (texto)."""
        linhas = [r["linha"] for r in nota["itens"]]
        if not set(linhas) & self.serie_ponto:
            return self.caminho, linhas, VERDE_CONCLUIDA, {}
        texto = (f"Robô SRP ({quando or f'{datetime.now():%d/%m/%Y %H:%M}'}): o SRP disse que o Documento "
                 "Recebimento já existia; o robô apagou a Chave e pôs '.' no fim da Série.")
        return self.caminho, linhas, VERDE_SERIE_PONTO, {linhas[0]: texto}

    def _pintor(self):
        """Thread unica: uma gravacao da planilha por vez; junta as notas que terminaram enquanto pintava."""
        while True:
            pedidos = [self.fila_pintar.get()]
            while True:
                try:
                    pedidos.append(self.fila_pintar.get_nowait())
                except queue.Empty:
                    break
            self._pintar(pedidos)

    def _pintar(self, pedidos):
        """Grava os pedidos (planilha, linhas, cor, comentarios), juntando os da mesma planilha e cor."""
        with self.trava_pintura:
            self._gravar_pedidos(pedidos)

    def _gravar_pedidos(self, pedidos):
        juntos = {}  # (planilha, cor) -> (linhas, comentarios)
        for caminho, linhas, cor, comentarios in pedidos:
            grupo = juntos.setdefault((caminho, cor), ([], {}))
            grupo[0].extend(linhas)
            grupo[1].update(comentarios)
        for (caminho, cor), (linhas, comentarios) in juntos.items():
            lista = ", ".join(str(l) for l in sorted(linhas))
            serie = " (verde mais escuro e comentário na CHAVE NF: '.' na Série)" if cor == VERDE_SERIE_PONTO else ""
            try:
                como = pintar_linhas(caminho, linhas, cor, comentarios)
            except Exception as erro:  # pintar nunca para o robo
                self._avisar_pintura(f"Aviso: não pintei de verde a(s) linha(s) {lista} na planilha "
                                     f"{Path(caminho).name} ({erro}).")
                continue
            if como == "excel" and not self.avisou_excel:
                self.avisou_excel = True
                self._avisar_pintura("Aviso: a planilha está aberta no Excel: pinto as linhas concluídas lá; salve "
                                     "a planilha no Excel para manter as cores.")
            if caminho == self.caminho:  # a planilha carregada passa a ter essas linhas verdes (o Iniciar compara)
                for r in self.registros:
                    if r["linha"] in linhas:
                        r["verde"] = True
            self._avisar_pintura(f"  Planilha: linha(s) {lista} pintada(s) de verde{serie}"
                                 + (" (no Excel aberto, falta salvar)" if como == "excel" else " (arquivo gravado)"))

    def _avisar_pintura(self, texto):
        """Mensagem da pintura no console e no log da execucao (usuario, 09/11/2026: antes so ia para o console)."""
        self.fila.put(texto)
        snac0236.registrar_no_log("  pintura: " + texto.strip())

    # ---------- conferir a planilha com os logs (usuario, 09/11/2026) ----------
    def _ler_conferencia(self):
        """(notas da planilha relidas, com a cor atual de cada linha; resultado de conferir_com_logs; logs) ou None."""
        try:
            _, notas, _ = snac0236.carregar(self.caminho)
            logs = snac0236.notas_nos_logs(pasta_logs())
        except Exception as erro:
            messagebox.showerror("Conferir com os logs", f"Não consegui ler a planilha ou os logs:\n{erro}")
            return None
        return notas, snac0236.conferir_com_logs(notas, logs), logs

    def _linhas(self, nota):
        linhas = [r["linha"] for r in nota["itens"]]
        return str(linhas[0]) if len(linhas) == 1 else f"{linhas[0]}-{linhas[-1]}"

    def _pintar_lancadas(self, notas, logs):
        """Pinta na hora (nao pelo pintor: o Iniciar confere a planilha logo depois) as notas que o robo integrou e que
        nao estao verdes, e marca como concluidas no app."""
        pedidos = []
        for nota in notas:
            info = logs.get(nota["chave"][-6:], {})
            linhas = [r["linha"] for r in nota["itens"]]
            if info.get("serie_ponto"):
                self.serie_ponto.update(linhas)
            for linha in linhas:
                self._estado(linha, "ok")
            self.pintadas.update(linhas)
            pedidos.append(self._pedido_pintura(nota, f"log de {info.get('quando', '?')}"))
        self._pintar(pedidos)

    def conferir(self):
        """Confere a planilha carregada com os logs da SNAC0236 (ultimos 30 dias) e mostra no console: notas lancadas
        e nao pintadas (oferece pintar), verdes sem o 'Integrado com Sucesso' nos logs, verdes fora dos logs e as
        lancadas com '.' na Serie."""
        if not self.caminho:
            messagebox.showinfo("Conferir com os logs", "Carregue a planilha da SNAC0236 primeiro.")
            return
        if self.ocupado:
            return
        lido = self._ler_conferencia()
        if not lido:
            return
        notas, achado, logs = lido
        nome = Path(self.caminho).name
        self.escrever(f"Conferência de {nome} com os logs ({len(logs)} nota(s) nos logs dos últimos 30 dias; os logs "
                      "guardam só os 6 últimos dígitos da chave):")

        def lista(notas_, extra=lambda n: ""):
            return "; ".join(f"linha {self._linhas(n)} (…{n['chave'][-6:]}{extra(n)})" for n in notas_)

        quando = lambda n: f", {logs[n['chave'][-6:]]['quando']}"  # noqa: E731
        nao = achado["nao_pintadas"]
        self.escrever(f"Aviso: {len(nao)} nota(s) integrada(s) pelo robô e sem verde na planilha: {lista(nao, quando)}."
                      if nao else "  Nenhuma nota integrada pelo robô ficou sem verde.")
        sem = achado["sem_confirmacao"]
        if sem:
            motivo = lambda n: (", terminada à mão" if logs[n['chave'][-6:]]["a_mao"]  # noqa: E731
                                else ", sem 'Integrado com Sucesso'")
            self.escrever(f"Aviso: {len(sem)} nota(s) verde(s) sem a confirmação do SRP nos logs (confira no SRP): "
                          f"{lista(sem, motivo)}.")
        fora = achado["fora_dos_logs"]
        if fora:
            self.escrever(f"  {len(fora)} nota(s) verde(s) que não aparecem nos logs (pintadas à mão ou lançadas há "
                          f"mais de 30 dias): {lista(fora)}.")
        serie = achado["serie_ponto"]
        if serie:
            self.escrever(f"  {len(serie)} nota(s) lançada(s) com '.' na Série (o SRP disse que já existia): "
                          f"{lista(serie, quando)}.")
        if not (nao or sem):
            self._status(f"Conferência: a planilha bate com os logs ({nome})", "ok")
            return
        self._status(f"Conferência: {len(nao)} sem verde, {len(sem)} sem confirmação (veja os detalhes)", "aviso")
        if nao and messagebox.askyesno(
                "Conferir com os logs", f"{len(nao)} nota(s) foram integradas pelo robô (o SRP disse 'Integrado com "
                                        f"Sucesso') e não estão verdes na planilha:\n\n"
                                        + "\n".join(f"linha {self._linhas(n)}  (…{n['chave'][-6:]})" for n in nao[:15])
                                        + ("\n..." if len(nao) > 15 else "") + "\n\nPintar de verde agora?"):
            self._pintar_lancadas([n for n in self.notas if n["chave"] in {x["chave"] for x in nao}], logs)
            self._status(f"{len(nao)} nota(s) pintada(s) de verde (veja os detalhes)", "ok")

    def _pular_ja_integradas(self):
        """Antes do Iniciar: notas do lote que os logs dizem integradas com sucesso (e nao estao verdes). Lancar de
        novo daria a janela "ja existe" e o robo poria '.' na Serie, duplicando o documento. Pergunta: pintar e pular
        (Sim), lancar assim mesmo (Nao) ou cancelar. Devolve False para nao iniciar."""
        try:
            logs = snac0236.notas_nos_logs(pasta_logs())
        except Exception:
            return True
        lote = self._lote()
        ja = [n for n in lote if logs.get(n["chave"][-6:], {}).get("sucesso")]
        if not ja:
            return True
        resposta = messagebox.askyesnocancel(
            "Notas já integradas", f"Segundo os logs, {len(ja)} nota(s) deste lote já foram integradas pelo robô (o "
                                   "SRP disse 'Integrado com Sucesso'), mas não estão verdes na planilha:\n\n"
                                   + "\n".join(f"linha {self._linhas(n)}  (…{n['chave'][-6:]}, "
                                                f"{logs[n['chave'][-6:]]['quando']})" for n in ja[:15])
                                   + ("\n..." if len(ja) > 15 else "")
                                   + "\n\nLançar de novo faria o SRP dizer que o documento já existe e o robô poria "
                                     "'.' na Série (documento duplicado).\n\nSim: pintar de verde e pular essas.\n"
                                     "Não: lançar assim mesmo.\nCancelar: não iniciar.")
        if resposta is None:
            return False
        if resposta:
            self.escrever(f"Aviso: {len(ja)} nota(s) já integrada(s) segundo os logs: pintadas de verde e puladas "
                          f"(linhas {', '.join(self._linhas(n) for n in ja)}).")
            self._pintar_lancadas(ja, logs)
        return True

    # ---------- mensagens ----------
    def escrever(self, texto):
        self.historico.append(texto)
        tipo = self.ui.tipo_mensagem(texto)
        self.saida.configure(state="normal")
        self.saida.insert("end", texto + "\n", tipo or ())
        self.saida.see("end")
        self.saida.configure(state="disabled")
        if self.ocupado and not texto.startswith("  "):
            cor = "erro" if tipo == "erro" else "aviso" if "PAUSADO" in texto or tipo == "aviso" else "destaque"
            self._status(texto if len(texto) < 220 else texto[:217] + "...", cor)

    def _status(self, texto, cor="suave"):
        self.status_atual = (texto, cor)
        self.status.configure(text=texto, fg=self.ui.COR[cor])

    def copiar(self):
        self.app.clipboard_clear()
        self.app.clipboard_append("\n".join(self.historico[-500:]))
        self._status("Detalhes copiados", "suave")
