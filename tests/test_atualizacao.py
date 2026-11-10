"""Autoatualizacao: comparacao de versoes, leitura do release (API simulada) e download com conferencia do SHA-256."""
import hashlib
import io
import json
import pathlib
import tempfile
import unittest
import zipfile
from unittest import mock

from simulador import core


def resposta_api(dados):
    contexto = mock.MagicMock()
    contexto.__enter__.return_value = io.BytesIO(json.dumps(dados).encode("utf-8"))
    return contexto


RELEASE = {"tag_name": "v1.10.0", "body": "### Novidades\n- algo",
           "assets": [{"name": "RPA-SRP-v1.10.0.zip", "size": 10, "digest": "sha256:abc",
                       "browser_download_url": "https://exemplo/RPA-SRP-v1.10.0.zip"}]}


class Versoes(unittest.TestCase):
    def test_compara_numericamente(self):
        self.assertGreater(core.versao_tupla("v1.10.0"), core.versao_tupla("1.9.3"))
        self.assertEqual(core.versao_tupla("1.8.0"), (1, 8, 0))
        self.assertEqual(core.versao_tupla("lixo"), (0,))


class Buscar(unittest.TestCase):
    def buscar(self, dados, atual="1.9.0"):
        with mock.patch("urllib.request.urlopen", return_value=resposta_api(dados)):
            return core.buscar_atualizacao(atual)

    def test_versao_nova(self):
        info = self.buscar(RELEASE)
        self.assertEqual((info["versao"], info["sha256"], info["url"]), ("1.10.0", "abc", "https://exemplo/RPA-SRP-v1.10.0.zip"))

    def test_mesma_versao_ou_sem_zip(self):
        self.assertIsNone(self.buscar(RELEASE, atual="1.10.0"))
        self.assertIsNone(self.buscar({**RELEASE, "assets": []}))

    def test_sem_internet_nao_levanta_erro(self):
        with mock.patch("urllib.request.urlopen", side_effect=OSError("sem rede")),                 mock.patch.object(core, "_consulta_powershell", side_effect=OSError("proxy fora")):
            info, texto = core.verificar_atualizacao("1.0.0")
        self.assertIsNone(info)
        self.assertIn("sem rede", texto)
        self.assertIn("proxy fora", texto)

    def test_direto_barrado_vai_pelo_proxy_do_windows(self):
        """Rede com PAC: o urllib nao ve o proxy e e barrado; o PowerShell usa o proxy do Windows."""
        with mock.patch("urllib.request.urlopen", side_effect=OSError("barrado")),                 mock.patch.object(core, "_consulta_powershell", return_value=RELEASE) as ps:
            info, texto = core.verificar_atualizacao("1.9.0")
        ps.assert_called_once()
        self.assertEqual((info["versao"], info["proxy"]), ("1.10.0", True))
        self.assertIn("proxy do Windows", texto)

    def test_texto_sem_versao_nova(self):
        with mock.patch("urllib.request.urlopen", return_value=resposta_api(RELEASE)):
            info, texto = core.verificar_atualizacao("1.10.0")
        self.assertIsNone(info)
        self.assertIn("mais nova", texto)

    def test_registra_no_log(self):
        pasta = pathlib.Path(tempfile.mkdtemp())
        with mock.patch.object(core, "pasta_logs", return_value=pasta):
            core.registrar_atualizacao("Atualizacoes: teste")
        self.assertIn("Atualizacoes: teste", (pasta / "atualizacao.log").read_text(encoding="utf-8"))


class Baixar(unittest.TestCase):
    def pacote(self):
        pasta = pathlib.Path(tempfile.mkdtemp())
        arquivo = pasta / "pacote.zip"
        with zipfile.ZipFile(arquivo, "w") as z:
            z.writestr("RPA-SRP.exe", b"exe")
            z.writestr("instalar.ps1", "# instalar")
            z.writestr("tesseract/tesseract.exe", b"t")
        return arquivo, hashlib.sha256(arquivo.read_bytes()).hexdigest()

    def test_baixa_confere_e_extrai(self):
        arquivo, soma = self.pacote()
        passos = []
        destino = core.baixar_atualizacao({"url": arquivo.as_uri(), "sha256": soma, "tamanho": arquivo.stat().st_size},
                                          lambda feito, total: passos.append((feito, total)))
        self.assertTrue((destino / "instalar.ps1").exists() and (destino / "tesseract" / "tesseract.exe").exists())
        self.assertEqual(passos[-1][0], arquivo.stat().st_size)

    def test_download_pelo_proxy_do_windows(self):
        """Consulta que so passou pelo proxy: baixa pelo PowerShell e confere o SHA-256 do mesmo jeito."""
        origem, soma = self.pacote()

        def falso(url, arquivo, total, progresso):
            arquivo.write_bytes(origem.read_bytes())
        with mock.patch.object(core, "_baixar_powershell", side_effect=falso) as ps,                 mock.patch("urllib.request.urlopen", side_effect=AssertionError("nao devia ir direto")):
            destino = core.baixar_atualizacao({"url": "https://exemplo/x.zip", "sha256": soma, "tamanho": None,
                                               "proxy": True})
        ps.assert_called_once()
        self.assertTrue((destino / "RPA-SRP.exe").exists())

    def test_download_direto_barrado_tenta_o_proxy(self):
        origem, _ = self.pacote()
        with mock.patch.object(core, "_baixar_powershell",
                               side_effect=lambda url, arquivo, *_: arquivo.write_bytes(origem.read_bytes())) as ps,                 mock.patch("urllib.request.urlopen", side_effect=OSError("barrado")):
            core.baixar_atualizacao({"url": "https://exemplo/x.zip", "sha256": None, "tamanho": None})
        ps.assert_called_once()

    def test_sha256_diferente_cancela(self):
        arquivo, _ = self.pacote()
        with self.assertRaises(ValueError):
            core.baixar_atualizacao({"url": arquivo.as_uri(), "sha256": "0" * 64, "tamanho": None})


if __name__ == "__main__":
    unittest.main()
