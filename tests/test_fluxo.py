"""Fluxo de preenchimento contra o SRP simulado: campos travados, pausa por aviso, pausa por quantidade divergente."""
import json
import threading
import time
import unittest

from PIL import Image

from simulador import SRP, ambiente, core, executar, registros

LIVRE = {"custo": False, "conta": False}


class CamposTravados(unittest.TestCase):
    def rodar(self, travas, com_cabecalho=True):
        srp = SRP(travas)
        with ambiente(srp, com_cabecalho):
            codigo, avisos, _ = executar(srp, registros(len(travas)))
        return codigo, srp.valores[:len(travas)], avisos

    def test_quatro_combinacoes(self):
        for custo in (False, True):
            for conta in (False, True):
                with self.subTest(custo_travado=custo, conta_travada=conta):
                    codigo, valores, _ = self.rodar([{"custo": custo, "conta": conta}] * 2)
                    esperado = {"item": "100", "quantidade": "1"}
                    if not custo:
                        esperado["custo"] = "7300"
                    if not conta:
                        esperado["conta"] = "301"
                    self.assertEqual(codigo, 0)
                    self.assertEqual(valores[0], esperado)

    def test_linhas_com_travas_diferentes(self):
        codigo, valores, _ = self.rodar([LIVRE, {"custo": True, "conta": False}, {"custo": True, "conta": True}])
        self.assertEqual(codigo, 0)
        self.assertEqual([sorted(v) for v in valores],
                         [["conta", "custo", "item", "quantidade"], ["conta", "item", "quantidade"], ["item", "quantidade"]])

    def test_sem_cabecalho_avisa_e_digita_c_custo(self):
        codigo, valores, avisos = self.rodar([{"custo": False, "conta": True}], com_cabecalho=False)
        self.assertEqual(codigo, 0)
        self.assertEqual(valores[0], {"item": "100", "custo": "7300", "quantidade": "1"})
        self.assertTrue(any("cabecalho" in a for a in avisos))


class PausaPorAviso(unittest.TestCase):
    """O SRP recusa a Conta da linha 3 (indice 1); o usuario corrige no SRP e clica em Continuar."""

    def rodar(self, reacao, retomar=True, extras=None):
        srp = SRP([LIVRE] * 3, recusar={(1, "conta")})
        evento = threading.Event() if retomar else None
        with ambiente(srp, extras=extras) as pasta:
            codigo, avisos, estados = executar(srp, registros(3), evento, reacao)
        return codigo, srp.valores[:3], avisos, estados, pasta

    @staticmethod
    def corrige_e_tab(srp, retomar, _linha):
        time.sleep(0.2)
        with srp.lock:
            srp.recusar.clear()
            srp.valores[srp.linha]["conta"] = "8712"
            srp.campo = "quantidade"
        retomar.set()

    def test_corrige_tab_e_continua(self):
        codigo, valores, _, estados, _ = self.rodar(self.corrige_e_tab)
        self.assertEqual(codigo, 0)
        self.assertEqual(valores[1], {"item": "101", "custo": "7300", "conta": "8712", "quantidade": "2"})
        self.assertEqual(valores[2]["quantidade"], "3")
        self.assertIn((3, "pausado"), estados)

    def test_pausa_na_hora_que_o_aviso_aparece(self):
        """Antes esperava o tempo maximo (render) para concluir que o campo nao mudou; agora pausa ao ver o aviso."""
        srp = SRP([LIVRE] * 3, recusar={(1, "conta")})
        momentos = {}

        def reacao(srp, retomar, _linha):
            momentos["pausa"] = time.time()
            self.corrige_e_tab(srp, retomar, _linha)
        with ambiente(srp):
            inicio = time.time()
            codigo, _, _ = executar(srp, registros(3), threading.Event(), reacao, render=5.0)
        self.assertEqual(codigo, 0)
        self.assertLess(momentos["pausa"] - inicio, 2.0)

    def test_esquece_o_tab_continua_pausado(self):
        tentativas = []

        def reacao(srp, retomar, _linha):
            time.sleep(0.2)
            tentativas.append(1)
            with srp.lock:
                srp.recusar.clear()
                srp.valores[srp.linha]["conta"] = "8712"
                if len(tentativas) >= 2:
                    srp.campo = "quantidade"
            retomar.set()
        codigo, valores, avisos, _, _ = self.rodar(reacao)
        self.assertEqual(codigo, 0)
        self.assertEqual(valores[1]["quantidade"], "2")
        self.assertTrue(any("Ainda pausado" in a for a in avisos))

    def test_linha_completada_a_mao(self):
        def reacao(srp, retomar, _linha):
            time.sleep(0.2)
            with srp.lock:
                srp.recusar.clear()
                srp.valores[srp.linha].update({"conta": "8712", "quantidade": "5"})
                srp.linha += 1
                srp.campo = "item"
                srp.valores.append({})
            retomar.set()
        codigo, valores, _, _, _ = self.rodar(reacao)
        self.assertEqual(codigo, 0)
        self.assertEqual(valores[1]["quantidade"], "5")
        self.assertEqual(valores[2], {"item": "102", "custo": "7300", "conta": "301", "quantidade": "3"})

    def test_sem_continuar_interrompe(self):
        codigo, _, avisos, _, _ = self.rodar(None, retomar=False)
        self.assertEqual(codigo, 1)
        self.assertTrue(any("Interrompido" in a for a in avisos))

    def test_aviso_lido_e_conta_aprendida(self):
        info = {"texto": "Conta Contabil invalida para o item 101 conta 301 c.custo 7300. Conta correta=8712",
                "item": "101", "custo": "7300", "conta_correta": "8712"}
        codigo, _, avisos, _, pasta = self.rodar(self.corrige_e_tab, extras={"ler_aviso": lambda imagem, t: info})
        self.assertEqual(codigo, 0)
        self.assertTrue(any("SRP:" in a and "8712" in a for a in avisos))
        aprendido = json.loads((pasta / "aprendizado.json").read_text(encoding="utf-8"))
        self.assertEqual(aprendido["contas"]["101|7300"]["conta"], "8712")

    def test_aviso_de_outro_item_nao_ensina(self):
        info = {"texto": "...", "item": "999", "custo": "7300", "conta_correta": "8712"}
        _, _, _, _, pasta = self.rodar(self.corrige_e_tab, extras={"ler_aviso": lambda imagem, t: info})
        self.assertFalse((pasta / "aprendizado.json").exists())


class PausaPorF9(unittest.TestCase):
    def test_pausa_no_fim_da_linha_e_continua(self):
        srp = SRP([LIVRE] * 3)
        pausar, retomar = threading.Event(), threading.Event()
        pausas = []

        def estado(linha, nome):
            if nome == "ok" and linha == 2:
                pausar.set()  # F9 durante a linha 2: so pausa antes da linha 3
            if nome == "pausado":
                pausas.append(linha)
                threading.Timer(0.2, retomar.set).start()  # usuario clica em Continuar (cursor ja no Item)
        avisos = []
        with ambiente(srp):
            codigo = core.executar(registros(3), "Desktop Viewer", 0, 0.3, 0, threading.Event(), avisos.append, estado,
                                   False, retomar, pausar)
        self.assertEqual(codigo, 0)
        self.assertEqual(pausas, [3])
        self.assertEqual([v.get("quantidade") for v in srp.valores[:3]], ["1", "2", "3"])
        self.assertTrue(any("F9" in a for a in avisos))
        self.assertFalse(pausar.is_set())


class PausaPorQuantidade(unittest.TestCase):
    """O OCR (simulado) le na linha 3 uma quantidade diferente da planilha."""

    def rodar(self, reacao, retomar=True):
        srp = SRP([LIVRE] * 4)
        extras = {"recortar_celula": lambda janela, celula: Image.new("RGB", (90, 19), (255, 255, 255)),
                  "conferir_recorte": lambda recorte, t, esperada: (esperada != 2.0, [esperada if esperada != 2.0 else 20.0])}
        evento = threading.Event() if retomar else None
        with ambiente(srp, extras=extras):
            return executar(srp, registros(4), evento, reacao, conferir=True)

    def test_pausa_e_continua_quando_confere_no_srp(self):
        def reacao(srp, retomar, _linha):
            time.sleep(0.2)
            retomar.set()  # usuario confere: esta certo; cursor segue no Item da proxima linha
        codigo, avisos, estados = self.rodar(reacao)
        self.assertEqual(codigo, 0)
        self.assertIn((3, "pausado"), estados)
        self.assertIn((3, "ok"), estados[estados.index((3, "pausado")):])
        self.assertTrue(any("difere da planilha" in a and "PAUSADO" in a for a in avisos))

    def test_sem_continuar_interrompe(self):
        codigo, avisos, estados = self.rodar(None, retomar=False)
        self.assertEqual(codigo, 1)
        self.assertIn((3, "erro"), estados)


class ContasSuspeitas(unittest.TestCase):
    def test_avisa_conta_ja_recusada(self):
        srp = SRP([LIVRE])
        with ambiente(srp):
            core.aprender_conta("100", "7300", "8712")
            regs = registros(2)
            self.assertEqual(core.contas_suspeitas(regs), [(2, "100", "301", "8712")])
            regs[0]["conta"] = "8712"
            self.assertEqual(core.contas_suspeitas(regs), [])


if __name__ == "__main__":
    unittest.main()
