"""Logica do robo do SRP (via Citrix): leitura da planilha, validacao e envio de teclas.

Sequencia por linha: Item, Tab, C/Custo, Tab, [Conta, Tab se o campo estiver editavel], Quantidade, Tab.
A tela do Citrix e so imagem: o campo em foco e o bloco ciano; a Conta e editavel se o foco cai logo apos o C/Custo.
"""
import ctypes
import re
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

sys.coinit_flags = 2  # COM em modo STA, necessario para arrastar e soltar (tkdnd/OLE) e pywinauto coexistirem

from openpyxl import load_workbook
from PIL import Image, ImageChops, ImageFilter, ImageGrab, ImageOps
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

COLUNAS = {"item": "item", "c/custo": "custo", "conta": "conta", "quantidade": "quantidade"}
OBRIGATORIAS = ("item", "custo", "quantidade")


def texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def ler_planilha(caminho):
    ws = load_workbook(caminho, read_only=True, data_only=True).active
    linhas = ws.iter_rows(values_only=True)
    cabecalho = next(linhas, None) or ()
    indices = {}
    for posicao, nome in enumerate(cabecalho):
        chave = COLUNAS.get(texto(nome).lower())
        if chave:
            indices[chave] = posicao
    faltando = [nome for nome, chave in COLUNAS.items() if chave not in indices]
    if faltando:
        raise ValueError(f"Colunas ausentes no cabecalho: {', '.join(faltando)}")
    registros = []
    for numero, linha in enumerate(linhas, start=2):
        registro = {chave: texto(linha[pos]) if pos < len(linha) else "" for chave, pos in indices.items()}
        if any(registro.values()):
            registro["linha"] = numero
            registros.append(registro)
    return registros


def validar(registros):
    erros = []
    for r in registros:
        for campo in OBRIGATORIAS:
            if not r[campo]:
                erros.append(f"linha {r['linha']}: {campo} em branco")
        for campo in ("item", "custo", "conta"):
            if r[campo] and not re.fullmatch(r"[0-9A-Za-z.\-]+", r[campo]):
                erros.append(f"linha {r['linha']}: {campo} com caracteres invalidos: {r[campo]!r}")
        if r["quantidade"] and not re.fullmatch(r"\d+(?:[.,]\d+)?", r["quantidade"]):
            erros.append(f"linha {r['linha']}: quantidade invalida: {r['quantidade']!r}")
    return erros


def passos(r, conta_editavel=None):
    """conta_editavel None = decide pela planilha (simulacao); True/False = detectado na tela."""
    usar_conta = bool(r["conta"]) if conta_editavel is None else conta_editavel
    seq = [("digitar", r["item"]), ("tab", None), ("digitar", r["custo"]), ("tab", None)]
    if usar_conta:
        seq += [("digitar", r["conta"]), ("tab", None)]
    seq += [("digitar", r["quantidade"].replace(".", ",")), ("tab", None)]
    return seq


def descrever(r):
    return " ".join("TAB" if v is None else f"'{v}'" for _, v in passos(r))


def maior_caixa_ciano(imagem):
    """Retângulo (esq, topo, dir, base) do maior bloco ciano puro; ignora pixels ciano soltos."""
    r, g, b = imagem.convert("RGB").split()
    mascara = ImageChops.multiply(ImageChops.multiply(r.point(lambda v: 255 if v < 40 else 0),
                                                      g.point(lambda v: 255 if v > 215 else 0)),
                                  b.point(lambda v: 255 if v > 215 else 0))
    mascara = mascara.filter(ImageFilter.MinFilter(5))
    largura, altura = mascara.size
    dados = mascara.load()
    visto = set()
    melhor = None
    for y0 in range(altura):
        for x0 in range(largura):
            if dados[x0, y0] and (x0, y0) not in visto:
                pilha, area = [(x0, y0)], 0
                esq, topo, dir_, base = x0, y0, x0, y0
                visto.add((x0, y0))
                while pilha:
                    x, y = pilha.pop()
                    area += 1
                    esq, topo, dir_, base = min(esq, x), min(topo, y), max(dir_, x), max(base, y)
                    for vx, vy in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
                        if 0 <= vx < largura and 0 <= vy < altura and dados[vx, vy] and (vx, vy) not in visto:
                            visto.add((vx, vy))
                            pilha.append((vx, vy))
                if melhor is None or area > melhor[0]:
                    melhor = (area, (esq, topo, dir_ + 1, base + 1))
    return melhor[1] if melhor else None


def capturar(janela):
    ret = janela.rectangle()
    return ImageGrab.grab(bbox=(ret.left, ret.top, ret.right, ret.bottom), all_screens=True)


def caixa_foco(janela):
    """Retângulo do campo em foco (ciano) na janela do Citrix."""
    return maior_caixa_ciano(capturar(janela))


def caminho_tesseract():
    """Tesseract portatil: pasta `tesseract` ao lado do exe, ou `vendor/tesseract` no codigo-fonte."""
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent / "vendor"
    exe = base / "tesseract" / "tesseract.exe"
    return exe if exe.exists() else None


def ocr_numero(imagem, tesseract):
    """Le um numero de uma imagem recortada com o Tesseract (somente digitos, ponto, virgula e sinal)."""
    imagem = ImageOps.autocontrast(imagem.convert("L").resize((imagem.width * 4, imagem.height * 4), Image.LANCZOS))
    imagem = ImageOps.expand(imagem, border=20, fill=255)
    with tempfile.TemporaryDirectory() as pasta:
        arquivo = Path(pasta) / "campo.png"
        imagem.save(arquivo)
        saida = subprocess.run(
            [str(tesseract), str(arquivo), "stdout", "--tessdata-dir", str(tesseract.parent / "tessdata"), "--psm", "7",
             "-c", "tessedit_char_whitelist=0123456789.,-"],
            capture_output=True, text=True, timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    limpo = saida.stdout.strip().replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return None


def ler_quantidade(janela, tesseract):
    """Le o numero do campo em foco (Quantidade) na tela do Citrix."""
    imagem = capturar(janela)
    caixa = maior_caixa_ciano(imagem)
    return ocr_numero(imagem.crop(caixa), tesseract) if caixa else None


def aguardar_mudanca(janela, referencia, limite):
    """Espera o foco (ciano) sair de `referencia`; devolve a nova caixa ou None no tempo esgotado."""
    prazo = time.time() + limite
    while time.time() < prazo:
        atual = caixa_foco(janela)
        if atual and referencia and (abs(atual[0] - referencia[0]) > 5 or abs(atual[1] - referencia[1]) > 5):
            time.sleep(0.15)
            return caixa_foco(janela)
        time.sleep(0.1)
    return None


def conta_editavel(antes, depois):
    """Conta editavel se o novo foco comeca logo apos o C/Custo (vao menor que 1,5 largura)."""
    largura = antes[2] - antes[0]
    return (depois[0] - antes[2]) < largura * 1.5


def mouse_no_canto():
    class Ponto(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
    ponto = Ponto()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(ponto))
    return ponto.x <= 2 and ponto.y <= 2


def janela_citrix(trecho):
    for janela in Desktop(backend="uia").windows():
        if trecho.lower() in janela.window_text().lower():
            return janela
    raise LookupError(f"Janela com '{trecho}' no titulo nao encontrada.")


def pasta_logs():
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent
    pasta = base / "logs"
    pasta.mkdir(exist_ok=True)
    return pasta


def executar(registros, janela_trecho, pausa, render, espera, parar, aviso, estado=None, conferir=False):
    """Preenche o SRP. `parar` e um threading.Event; `aviso(texto)` recebe o andamento;
    `estado(linha, nome)` (opcional) recebe 'executando', 'ok' ou 'erro' por linha;
    `conferir` le a Quantidade na tela (OCR) antes de confirmar cada linha. Devolve 0 se concluiu."""
    tesseract = caminho_tesseract() if conferir else None
    log = open(pasta_logs() / f"{datetime.now():%Y%m%d-%H%M%S}.log", "w", encoding="utf-8")
    atual = None

    def marcar(r, nome):
        nonlocal atual
        atual = r["linha"] if nome == "executando" else None
        if estado:
            estado(r["linha"], nome)

    def registrar(texto, mostrar=True):
        log.write(texto + "\n")
        log.flush()
        if mostrar:
            aviso(texto)

    def teclar(acao, valor):
        if parar.is_set() or mouse_no_canto():
            raise KeyboardInterrupt
        send_keys("{TAB}" if acao == "tab" else valor.replace(",", "{,}"), pause=0.05)
        time.sleep(pausa)

    try:
        if conferir and not tesseract:
            aviso("ATENCAO: pasta 'tesseract' nao encontrada ao lado do programa. Conferencia da quantidade DESATIVADA nesta execucao.")
        janela = janela_citrix(janela_trecho)
        janela.set_focus()
        aviso(f"Clique no campo Item da primeira linha. Iniciando em {espera}s (mouse no canto superior esquerdo aborta).")
        for _ in range(espera * 10):
            if parar.is_set():
                registrar("Interrompido antes de iniciar.")
                return 1
            time.sleep(0.1)
        for r in registros:
            if janela_trecho.lower() not in Desktop(backend="uia").window(active_only=True).window_text().lower():
                registrar(f"Foco saiu do Citrix antes da linha {r['linha']}. Interrompido.")
                return 1
            marcar(r, "executando")
            inicial = caixa_foco(janela)
            teclar("digitar", r["item"]); teclar("tab", None)
            antes = aguardar_mudanca(janela, inicial, render)
            teclar("digitar", r["custo"]); teclar("tab", None)
            depois = aguardar_mudanca(janela, antes, render)
            if not inicial or not antes or not depois:
                registrar(f"Linha {r['linha']}: campo em foco (ciano) nao detectado. Interrompido.")
                return 1
            editavel = conta_editavel(antes, depois)
            registrar(f"  linha {r['linha']}: foco C/Custo={antes} depois={depois} editavel={editavel}", mostrar=False)
            if editavel:
                if not r["conta"]:
                    registrar(f"Linha {r['linha']}: Conta editavel mas em branco na planilha. Interrompido.")
                    return 1
                teclar("digitar", r["conta"]); teclar("tab", None)
            teclar("digitar", r["quantidade"].replace(".", ","))
            if tesseract:
                time.sleep(0.3)
                lida = ler_quantidade(janela, tesseract)
                esperada = float(r["quantidade"].replace(",", "."))
                registrar(f"  linha {r['linha']}: quantidade lida={lida} esperada={esperada}", mostrar=False)
                if lida is None or abs(lida - esperada) > 1e-6:
                    registrar(f"Linha {r['linha']}: quantidade na tela ({lida}) difere da planilha ({esperada}). Interrompido.")
                    return 1
            teclar("tab", None)
            conta = "digitada " + r["conta"] if editavel else "travada"
            registrar(f"OK linha {r['linha']}: {r['item']} {r['custo']} conta={conta} {r['quantidade']}")
            marcar(r, "ok")
        registrar("Concluido.")
        return 0
    except KeyboardInterrupt:
        registrar("Interrompido pelo usuario.")
        return 1
    except Exception as erro:
        registrar(f"Erro: {erro}")
        return 1
    finally:
        if atual is not None and estado:
            estado(atual, "erro")
        log.close()
