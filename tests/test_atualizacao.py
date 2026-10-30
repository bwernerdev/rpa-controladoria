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
        with mock.patch("urllib.request.urlopen", side_effect=OSError("sem rede")):
            self.assertIsNone(core.buscar_atualizacao("1.0.0"))


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

    def test_sha256_diferente_cancela(self):
        arquivo, _ = self.pacote()
        with self.assertRaises(ValueError):
            core.baixar_atualizacao({"url": arquivo.as_uri(), "sha256": "0" * 64, "tamanho": None})


if __name__ == "__main__":
    unittest.main()
