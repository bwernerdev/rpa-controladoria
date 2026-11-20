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
            (gui_snac0236, "registrar_falha_tela"): lambda contexto: self.falhas.append(contexto),
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
        self.app.fila.put(("linha_sem_estado",))  # _estado_linha com argumento faltando: TypeError
        self.app.fila.put("depois do erro")
        self.rodar()
        self.assertTrue(self.falhas)
        self.assertIn("depois do erro", self.app.historico)
        self.app.fila.put("e o laco continua")
        self.rodar()
        self.assertIn("e o laco continua", self.app.historico)

    def test_fila_da_0236_sobrevive_a_mensagem_com_erro(self):
        p = self.app.p0236
        p.fila.put(("linha_sem_estado",))
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
