"""Remove a informacao de depuracao (secoes DWARF "/N" e tabela de simbolos COFF) de DLLs/EXEs do MinGW, como `strip`.

Uso: python scripts\\enxugar_dll.py ARQUIVO [ARQUIVO ...]
O codigo e os dados carregados nao mudam; so some o que fica depois da ultima secao usada em execucao.
"""
import struct
import sys
from pathlib import Path


def alinhar(valor, passo):
    return (valor + passo - 1) // passo * passo


def soma_pe(dados, pos_soma):
    """Checksum do cabecalho PE (mesmo algoritmo do CheckSumMappedFile)."""
    total = 0
    tamanho = len(dados)
    preenchido = dados + b"\0" * (tamanho % 2)
    for i in range(0, len(preenchido), 2):
        if i == pos_soma or i == pos_soma + 2:
            continue
        total += struct.unpack_from("<H", preenchido, i)[0]
        total = (total & 0xFFFF) + (total >> 16)
    total = (total & 0xFFFF) + (total >> 16)
    return (total + tamanho) & 0xFFFFFFFF


def enxugar(caminho):
    dados = bytearray(Path(caminho).read_bytes())
    pe = struct.unpack_from("<I", dados, 0x3C)[0]
    if dados[pe:pe + 4] != b"PE\0\0":
        return None
    n = struct.unpack_from("<H", dados, pe + 6)[0]
    opt = pe + 24
    tam_opt = struct.unpack_from("<H", dados, pe + 20)[0]
    alinh_secao = struct.unpack_from("<I", dados, opt + 32)[0]
    cab = opt + tam_opt
    secoes = [struct.unpack_from("<8sIIII", dados, cab + 40 * i) for i in range(n)]
    manter = [s for s in secoes if not s[0].startswith(b"/") and not s[0].startswith(b".debug")]
    if len(manter) == n and struct.unpack_from("<I", dados, pe + 12)[0] == 0:
        return None
    if secoes[:len(manter)] != manter:
        raise ValueError(f"{caminho}: secoes de depuracao no meio do arquivo; nada alterado")
    fim = max(s[4] + s[3] for s in manter)
    if any(s[4] < fim and s[3] for s in secoes[len(manter):]):
        raise ValueError(f"{caminho}: secao de depuracao antes do fim do codigo; nada alterado")
    for i in range(len(manter), n):
        dados[cab + 40 * i:cab + 40 * (i + 1)] = b"\0" * 40
    struct.pack_into("<H", dados, pe + 6, len(manter))
    struct.pack_into("<II", dados, pe + 12, 0, 0)  # tabela de simbolos COFF
    struct.pack_into("<I", dados, opt + 56, alinhar(max(s[2] + s[1] for s in manter), alinh_secao))  # SizeOfImage
    dados = dados[:fim]
    struct.pack_into("<I", dados, opt + 64, soma_pe(bytes(dados), opt + 64))
    antes = Path(caminho).stat().st_size
    Path(caminho).write_bytes(dados)
    return antes, len(dados)


if __name__ == "__main__":
    for arquivo in sys.argv[1:]:
        resultado = enxugar(arquivo)
        if resultado:
            print(f"{arquivo}: {resultado[0] / 2**20:.1f} MB -> {resultado[1] / 2**20:.1f} MB")
