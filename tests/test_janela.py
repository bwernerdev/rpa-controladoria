"""Testes da janela (gui.py e gui_snac0236.py) sem o SRP: o App e montado com config, progresso, dialogos e Citrix
falsos (nada toca no config.json, no progresso.json nem nos logs reais). Cobrem o que a v1.14.0 mudou na tela: fila
que sobrevive a erro, X com o robo rodando, Simular abrindo os Detalhes, contagem do inicio, Opcoes normalizadas,
dialogo proprio (Enter/Esc), F9 da SNAC0236 e troca de tema. Precisam do customtkinter e do tkinterdnd2 (pulam sem)."""
import importlib.util
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
TEM_TELA = all(importlib.util.find_spec(m) for m in ("customtkinter", "tkinterdnd2"))
if TEM_TELA:
    import gui
    import gui_snac0236
    import pagina


@unittest.skipUnless(TEM_TELA, "sem customtkinter/tkinterdnd2")
class Janela(unittest.TestCase):
    def setUp(self):
        self.config = {"tema": "claro", "aba": "snac0234"}
        self.gravados, self.falhas, self.dialogos = [], [], []
        self.respostas = []  # respostas dos perguntar, na ordem (vazio = True)
        trocas = {
            (gui, "ler_config"): lambda: dict(self.config),
            (gui, "gravar_config"): lambda dados: self.gravados.append(dict(dados)),
            (gui, "gravar_progresso"): lambda *a, **k: None,
            (gui, "ler_progresso"): lambda *a, **k: (set(), None),
            (gui, "contas_suspeitas"): lambda *a, **k: [],
            (gui, "citrix_aberto"): lambda: False,
            (gui, "registrar_falha_tela"): lambda contexto: self.falhas.append(contexto),
            (pagina, "registrar_falha_tela"): lambda contexto: self.falhas.append(contexto),
            (gui.App, "avisar"): lambda _app, *a, **k: self.dialogos.append(("avisar",) + a),
            (gui.App, "perguntar"): lambda _app, *a, **k: (self.dialogos.append(("perguntar",) + a),
                                                            self.respostas.pop(0) if self.respostas else True)[1],
        }
        self.originais = {chave: getattr(*chave) for chave in trocas}
        for (alvo, nome), valor in trocas.items():
            setattr(alvo, nome, valor)
        self.app = gui.App()
        self.rodar()

    def tearDown(self):
        try:
            if self._existe():
                self.app.destroy()  # erro aqui = teste falha (a v1.14.0 ja teve um destroy quebrado)
        finally:
            self._restaurar()

    def _restaurar(self):
        for (alvo, nome), valor in self.originais.items():
            setattr(alvo, nome, valor)

    def rodar(self, segundos=0.3):
        fim = time.time() + segundos
        while time.time() < fim:
            try:
                self.app.update()
            except Exception:  # janela fechada no meio
                return
            time.sleep(0.02)

    def planilha(self, linhas=3):
        from openpyxl import Workbook
        pasta = Path(tempfile.mkdtemp())
        livro = Workbook()
        aba = livro.active
        aba.append(["Item", "C/Custo", "Conta", "Quantidade"])
        for i in range(linhas):
            aba.append([str(900135 + i), "7300", "8712", 2])
        caminho = pasta / "REQUISICAO FICTICIA.xlsx"
        livro.save(caminho)
        return str(caminho)

    def test_carregar_e_simular_abre_os_detalhes(self):
        self.assertFalse(self.app.console_aberto)
        self.app.carregar(self.planilha())
        self.assertEqual(len(self.app.registros), 3)
        self.app.simular()
        self.assertTrue(self.app.console_aberto)
        self.assertIn("Simulação de 3 linhas", self.app.status_atual[0])
        self.assertFalse(any(g.get("console_aberto") for g in self.gravados))  # abrir sozinho nao muda a preferencia

    def test_fila_sobrevive_a_mensagem_com_erro(self):
        self.app.fila.put(("estado", 1))  # sem o nome do estado: TypeError
        self.app.fila.put("depois do erro")
        self.rodar()
        self.assertTrue(self.falhas)
        self.assertIn("depois do erro", self.app.historico)
        self.app.fila.put("e o laco continua")
        self.rodar()
        self.assertIn("e o laco continua", self.app.historico)

    def test_fila_da_0236_sobrevive_a_mensagem_com_erro(self):
        p = self.app.p0236
        p.fila.put(("evento_que_nao_existe",))  # AttributeError
        p.fila.put("depois do erro")
        self.rodar()
        self.assertTrue(self.falhas)
        self.assertIn("depois do erro", p.historico)

    def test_fechar_parado_fecha(self):
        self.app.fechar()
        self.rodar()
        self.assertFalse(self._existe())

    def _existe(self):
        try:
            return bool(self.app.winfo_exists())
        except Exception:
            return False

    def test_fechar_com_o_robo_rodando_pergunta_e_para(self):
        self.app.ocupado = True
        self.app.thread = threading.Thread(target=lambda: self.app.parar.wait(10), daemon=True)
        self.app.thread.start()
        self.respostas = [False]
        self.app.fechar()
        self.rodar()
        self.assertTrue(self._existe())  # "Continuar preenchendo"
        self.assertFalse(self.app.parar.is_set())
        self.app.fechar()  # agora "Parar e fechar"
        self.assertTrue(self.app.parar.is_set())
        self.rodar(1.0)
        self.assertFalse(self._existe())
        self.assertEqual(sum(1 for d in self.dialogos if d[0] == "perguntar"), 2)

    def test_contagem_do_inicio(self):
        self.app.ocupado = True
        self.app._status("Iniciando...", "destaque")
        self.app.contagem(self.app, "Clique no campo Item no SRP", 2)
        self.rodar(0.2)
        self.assertEqual(self.app.status_atual[0], "Clique no campo Item no SRP: começa em 2 s")
        self.rodar(1.1)
        self.assertEqual(self.app.status_atual[0], "Clique no campo Item no SRP: começa em 1 s")
        self.app._status("Preenchendo linha 2...", "destaque")  # o robo escreveu: a contagem para
        self.rodar(1.2)
        self.assertEqual(self.app.status_atual[0], "Preenchendo linha 2...")

    def test_opcoes_normalizam_e_gravam(self):
        self.app.pausa.set("9")
        self.app.limite.set("")
        self.app._gravar_opcoes()
        self.assertEqual(self.app.pausa.get(), "0.25")
        self.assertEqual(self.app.limite.get(), "0")
        self.app.pausa.set("0,5")
        self.app._gravar_opcoes()
        self.assertEqual(self.app.pausa.get(), "0.5")
        self.assertEqual(self.gravados[-1]["pausa_0234"], 0.5)

    def test_opcoes_abrem_e_fecham_com_esc(self):
        self.app.abrir_opcoes()
        self.rodar(0.2)
        janelas = [w for w in self.app.winfo_children() if isinstance(w, gui.tk.Toplevel)]
        self.assertEqual(len(janelas), 1)
        janelas[0].event_generate("<Escape>")
        self.rodar(0.2)
        self.assertFalse(janelas[0].winfo_exists())

    def test_dialogo_enter_e_esc(self):
        for tecla, esperado in (("<Return>", "sim"), ("<Escape>", "cancelado")):
            def apertar(t=tecla):
                janela = [w for w in self.app.winfo_children() if isinstance(w, gui.tk.Toplevel)][-1]
                janela.event_generate(t)
            self.app.after(300, apertar)
            valor = gui.dialogo(self.app, "Título", "Texto", [("Não", "nao", "Secundario.TButton"),
                                                              ("Sim", "sim", "Primario.TButton")],
                                lista=["um", "dois"], padrao=1, cancelar="cancelado")
            self.assertEqual(valor, esperado)

    def test_f9_da_0236_pede_pausa_e_continua(self):
        p = self.app.p0236
        p.ocupado = True
        p.fila.put(("pausa",))
        self.rodar()
        self.assertTrue(p.pausar.is_set())
        self.assertEqual(p.status_atual[0], "Pausando ao fim da nota...")
        p.pausado = True
        p.fila.put(("pausa",))  # na pausa, F9 continua
        self.rodar()
        self.assertTrue(p.retomar.is_set())

    def test_atalho_escolhe_na_aba_aberta(self):
        chamados = []
        self.app.escolher = lambda: chamados.append("0234")
        self.app.p0236.escolher = lambda: chamados.append("0236")
        self.app._atalho("escolher")
        self.app.trocar_aba("snac0236")
        self.app._atalho("escolher")
        self.app.ocupado = True
        self.app._atalho("escolher")  # durante um preenchimento, nada
        self.assertEqual(chamados, ["0234", "0236"])

    def test_detalhes_e_dica_das_duas_abas(self):
        """Pagina (base das abas): cada aba grava a propria preferencia dos Detalhes, copia o historico e mostra dica."""
        for pagina_, chave in ((self.app, "console_aberto"), (self.app.p0236, "console_0236")):
            pagina_.alternar_console()
            self.assertTrue(pagina_.console_aberto)
            self.assertIs(self.gravados[-1][chave], True)
            pagina_.escrever("linha de teste")
            pagina_.copiar()
            self.assertIn("linha de teste", self.app.clipboard_get())
            pagina_.alternar_console()
            self.assertIs(self.gravados[-1][chave], False)
            evento = type("Evento", (), {"x_root": 10, "y_root": 10})()
            pagina_._mostrar_dica(evento, "dica")
            self.assertTrue(pagina_.dica.winfo_exists())
            pagina_._esconder_dica()
            self.assertIsNone(pagina_.dica)

    def test_cores_das_linhas_concluidas(self):
        """Concluida = verde comum; Serie com '.' = fundo mais escuro com o texto verde escuro (contraste 5:1)."""
        for tema in ("claro", "escuro"):
            self.app.escolher_tema(tema)
            self.rodar(0.2)
            tabela = self.app.p0236.tabela
            self.assertEqual(str(tabela.tag_configure("ok", "foreground")), gui.COR["ok"])
            self.assertEqual(str(tabela.tag_configure("ok_serie", "foreground")), gui.COR["ok_serie_texto"])
            self.assertEqual(str(self.app.tabela.tag_configure("ok", "foreground")), gui.COR["ok"])
            self.assertEqual([str(tabela.column(c, "stretch")) for c in ("linha", "status")], ["1", "1"])
            self.assertEqual([str(self.app.tabela.column(c, "stretch")) for c in ("linha", "status")], ["0", "1"])

    def _trocar(self, alvo, nome, valor):
        original = getattr(alvo, nome)
        setattr(alvo, nome, valor)
        self.addCleanup(setattr, alvo, nome, original)

    def test_iniciar_0234_com_robo_falso(self):
        """Iniciar -> thread -> eventos 'estado'/'fim' na fila -> tabela e botoes; executar recebe argumentos nomeados."""
        chamadas = []

        def executar(registros, **opcoes):
            chamadas.append(opcoes)
            for r in registros:
                opcoes["estado"](r["linha"], "executando")
                opcoes["aviso"](f"Linha {r['linha']}: ok")
                opcoes["estado"](r["linha"], "ok")
            return 0
        self._trocar(gui, "executar", executar)
        self._trocar(gui, "citrix_aberto", lambda: True)
        self._trocar(gui, "atalhos_globais", lambda _teclas: (lambda: None, []))
        self.app.carregar(self.planilha())
        self.app.iniciar()
        self.assertTrue(self.app.ocupado)
        self.rodar(1.0)  # o robo falso nao espera a contagem de 5 s (ela corre no executar de verdade)
        self.assertFalse(self.app.ocupado)
        self.assertEqual(set(self.app.estados.values()), {"ok"})
        self.assertEqual(len(chamadas), 1)
        self.assertEqual((chamadas[0]["janela_trecho"], chamadas[0]["render"], chamadas[0]["espera"]),
                         ("Desktop Viewer", 4.0, 5))
        self.assertIn("Linha 2: ok", self.app.historico)

    def test_rodar_0236_manda_argumentos_nomeados_e_fim(self):
        chamadas = []
        self._trocar(gui_snac0236.snac0236, "executar", lambda notas, **opcoes: chamadas.append(opcoes))
        p = self.app.p0236
        p.ocupado = self.app.ocupado = True
        p.execucao = {"inicio": time.time(), "notas": [], "feitas": 0, "gasto": 0.0, "tempo_pausado": 0.0,
                      "pausado_desde": None}
        p._rodar([], 0.5)
        self.rodar()
        self.assertFalse(p.ocupado or self.app.ocupado)
        self.assertEqual(chamadas[0]["intervalo"], 0.5)
        self.assertEqual(chamadas[0]["pausa"], gui_snac0236.PAUSA_TECLAS)
        chamadas[0]["serie_ponto"]([7])
        chamadas[0]["estado"](7, "ok")
        self.assertEqual(p.fila.get_nowait(), ("serie_ponto", (7,)))
        self.assertEqual(p.fila.get_nowait(), ("estado", 7, "ok"))

    def test_0236_cria_a_aba_lancados_e_copia_a_nota_integrada(self):
        """Iniciar cria a aba "Lançados"; nota concluida vai pelo pintor: verde na aba dos dados e copia com a data."""
        from openpyxl import Workbook, load_workbook
        caminho = Path(tempfile.mkdtemp()) / "NOTAS FICTICIAS.xlsx"
        livro = Workbook()
        livro.active.title = "Notas"
        livro.active.append(["CHAVE NF", "ITEM", "QTD FISCAL"])
        livro.active.append(["1" * 44, "900135", 2])
        livro.active.append([None, "900136", 3])
        livro.save(caminho)
        p = self.app.p0236
        p.caminho, p.registros = str(caminho), []
        p._criar_aba_lancados()
        p._criar_aba_lancados()  # a 2a vez nao faz nada (a aba ja existe)
        nota = {"chave": "1" * 44, "itens": [{"linha": 2}, {"linha": 3}]}
        p._pintar([p._pedido_pintura(nota)])
        mensagens = [m for m in p.fila.queue if isinstance(m, str)]
        self.assertEqual(sum("aba Lançados criada" in m for m in mensagens), 1, mensagens)
        self.assertTrue(any("pintada(s) de verde e copiada(s) para a aba Lançados" in m for m in mensagens), mensagens)
        livro = load_workbook(caminho)
        self.assertEqual(livro.active.title, "Notas")
        copia = [[c.value for c in linha] for linha in livro["Lançados"].iter_rows(min_row=2)]
        self.assertEqual([linha[:3] for linha in copia], [["1" * 44, "900135", 2], [None, "900136", 3]])
        self.assertTrue(all(isinstance(linha[3], gui_snac0236.datetime) for linha in copia))
        self.assertTrue(all(gui_snac0236.snac0236.core.fundo_verde(livro["Notas"].cell(row=n, column=1).fill) for n in (2, 3)))

    def test_trocar_tema_refaz_a_janela(self):
        self.app.carregar(self.planilha())
        self.app.escolher_tema("escuro")
        self.rodar()
        self.assertEqual(gui.COR["fundo"], gui.PALETAS["escuro"]["fundo"])
        self.assertEqual(len(self.app.tabela.get_children()), 3)
        self.app.escolher_tema("claro")
        self.rodar()
        self.assertEqual(gui.COR["fundo"], gui.PALETAS["claro"]["fundo"])


if __name__ == "__main__":
    unittest.main()
