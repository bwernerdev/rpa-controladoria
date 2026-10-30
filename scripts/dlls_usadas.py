"""Lista quais DLLs da pasta o tesseract.exe carrega (tabelas de importacao normal e delay-load, recursivo).

Uso: python scripts\\dlls_usadas.py vendor\\tesseract
"USA" = necessaria; "sobra" = so serve as ferramentas de treino do Tesseract e pode sair da pasta (rode os testes depois).
"""
import struct
import sys
from pathlib import Path

pasta = Path(sys.argv[1])
existentes = {p.name.lower(): p for p in pasta.glob("*.dll")}


def importacoes(caminho):
    d = caminho.read_bytes()
    pe = struct.unpack_from("<I", d, 0x3C)[0]
    n = struct.unpack_from("<H", d, pe + 6)[0]
    opt = pe + 24
    magic = struct.unpack_from("<H", d, opt)[0]
    dirs = opt + (112 if magic == 0x20B else 96)
    tam_opt = struct.unpack_from("<H", d, pe + 20)[0]
    cab = opt + tam_opt
    secoes = [struct.unpack_from("<8sIIII", d, cab + 40 * i) for i in range(n)]

    def off(rva):
        for _, vs, va, rs, rp in secoes:
            if va <= rva < va + max(vs, rs):
                return rva - va + rp
        return None

    nomes = []
    for indice, passo, pos_nome in ((1, 20, 12), (13, 32, 4)):  # import normal e delay-load
        rva = struct.unpack_from("<I", d, dirs + 8 * indice)[0]
        if not rva:
            continue
        o = off(rva)
        while o is not None:
            nome_rva = struct.unpack_from("<I", d, o + pos_nome)[0]
            if not nome_rva:
                break
            p = off(nome_rva)
            nomes.append(d[p:d.index(b"\0", p)].decode().lower())
            o += passo
    return nomes


usadas, fila = set(), ["tesseract.exe"]
while fila:
    atual = fila.pop()
    caminho = pasta / atual if atual == "tesseract.exe" else existentes.get(atual)
    if not caminho:
        continue
    for nome in importacoes(caminho):
        if nome in existentes and nome not in usadas:
            usadas.add(nome)
            fila.append(nome)
for nome in sorted(existentes):
    print(("USA  " if nome in usadas else "sobra") + f" {existentes[nome].stat().st_size / 2**20:6.1f} MB {nome}")
