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
from concurrent.futures import ThreadPoolExecutor
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


def itens_repetidos(registros):
    """{item: [linhas]} dos codigos de Item que aparecem mais de uma vez (o SRP soma itens iguais na mesma linha)."""
    linhas = {}
    for r in registros:
        linhas.setdefault(r["item"], []).append(r["linha"])
    return {item: lista for item, lista in linhas.items() if len(lista) > 1}


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


def _mascara_ciano(imagem):
    r, g, b = imagem.convert("RGB").split()
    mascara = ImageChops.multiply(ImageChops.multiply(r.point(lambda v: 255 if v < 40 else 0),
                                                      g.point(lambda v: 255 if v > 215 else 0)),
                                  b.point(lambda v: 255 if v > 215 else 0))
    return mascara.filter(ImageFilter.MinFilter(5))


def _trechos(valores):
    """Intervalos [ini, fim) onde a sequencia e diferente de zero."""
    trechos, inicio = [], None
    for i, v in enumerate(valores):
        if v and inicio is None:
            inicio = i
        elif not v and inicio is not None:
            trechos.append((inicio, i))
            inicio = None
    if inicio is not None:
        trechos.append((inicio, len(valores)))
    return trechos


def maior_caixa_ciano(imagem):
    """Retangulo (esq, topo, dir, base) do maior bloco ciano puro; ignora pixels ciano soltos.
    Trabalha por faixas (getbbox/resize em C): ~100x mais rapido que percorrer pixel a pixel, e nao trava a janela."""
    mascara = _mascara_ciano(imagem)
    caixa_total = mascara.getbbox()
    if not caixa_total:
        return None
    esq0, topo0, dir0, base0 = caixa_total
    regiao = mascara.crop(caixa_total)
    largura, altura = regiao.size
    linhas = [regiao.crop((0, y, largura, y + 1)).getbbox() is not None for y in range(altura)]
    melhor = None
    for y_ini, y_fim in _trechos(linhas):
        faixa = regiao.crop((0, y_ini, largura, y_fim))
        colunas = list(faixa.resize((largura, 1), Image.BOX).getdata())
        for x_ini, x_fim in _trechos(colunas):
            bloco = faixa.crop((x_ini, 0, x_fim, y_fim - y_ini)).getbbox()
            if not bloco:
                continue
            area = (bloco[2] - bloco[0]) * (bloco[3] - bloco[1])
            if melhor is None or area > melhor[0]:
                melhor = (area, (esq0 + x_ini + bloco[0], topo0 + y_ini + bloco[1], esq0 + x_ini + bloco[2], topo0 + y_ini + bloco[3]))
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


def preparar_ocr(imagem, escala=6, limiar=160):
    """Preto e branco (descarta a moldura clara da celula e o fundo ciano), ampliacao e margem branca para o Tesseract."""
    pb = imagem.convert("L").point(lambda v: 0 if v < limiar else 255)
    pb = pb.resize((pb.width * escala, pb.height * escala), Image.LANCZOS)
    return ImageOps.expand(pb, border=30, fill=255)


def ocr_numero(imagem, tesseract, psm=7, escala=6):
    """Le um numero de uma imagem recortada com o Tesseract (somente digitos, ponto, virgula e sinal)."""
    with tempfile.TemporaryDirectory() as pasta:
        arquivo = Path(pasta) / "campo.png"
        preparar_ocr(imagem, escala).save(arquivo)
        saida = subprocess.run(
            [str(tesseract), str(arquivo), "stdout", "--tessdata-dir", str(tesseract.parent / "tessdata"), "--psm", str(psm),
             "-c", "tessedit_char_whitelist=0123456789.,-"],
            capture_output=True, text=True, timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    limpo = saida.stdout.strip().replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return None


def recortar_celula(janela, caixa):
    """Captura a tela e recorta `caixa` (coordenadas da janela) com margem."""
    imagem = capturar(janela)
    esq, topo, dir_, base = caixa
    # a caixa vem encolhida ~2 px por lado (filtro de pixels soltos): margem maior, sobretudo a direita, onde o numero encosta
    return imagem.crop((max(esq - 4, 0), max(topo - 4, 0), min(dir_ + 8, imagem.width), min(base + 4, imagem.height)))


def conferir_recorte(recorte, tesseract, esperada):
    """Le o numero do recorte; tenta outras ampliacoes so se a primeira nao conferir. Devolve (confere, leituras).
    Funciona com o valor ja formatado (ex.: 1.0000); um digito isolado o Tesseract nao le de forma confiavel."""
    leituras = []
    for escala in (6, 4, 8):
        leituras.append(ocr_numero(recorte, tesseract, 7, escala))
        if leituras[-1] is not None and abs(leituras[-1] - esperada) < 1e-6:
            return True, leituras
    return False, leituras


def aguardar_mudanca(janela, referencia, limite):
    """Espera o foco (ciano) sair de `referencia`; devolve a nova caixa ou None no tempo esgotado."""
    prazo = time.time() + limite
    while time.time() < prazo:
        atual = caixa_foco(janela)
        if atual and referencia and (abs(atual[0] - referencia[0]) > 5 or abs(atual[1] - referencia[1]) > 5):
            return atual
        time.sleep(0.05)
    return None


def titulo_janela_ativa():
    """Titulo da janela em primeiro plano (chamada direta ao Windows; a busca por UIA leva ~1 s)."""
    alvo = ctypes.windll.user32.GetForegroundWindow()
    tamanho = ctypes.windll.user32.GetWindowTextLengthW(alvo)
    texto = ctypes.create_unicode_buffer(tamanho + 1)
    ctypes.windll.user32.GetWindowTextW(alvo, texto, tamanho + 1)
    return texto.value


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


def pasta_base():
    """Pasta do .exe quando empacotado; raiz do projeto no codigo-fonte."""
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent


def citrix_aberto(trecho="Desktop Viewer"):
    """Ha janela visivel com `trecho` no titulo? Usa EnumWindows (rapido; nao disputa a janela do app)."""
    user32 = ctypes.windll.user32
    achou = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def visitar(alvo, _):
        if user32.IsWindowVisible(alvo):
            tamanho = user32.GetWindowTextLengthW(alvo)
            if tamanho:
                texto = ctypes.create_unicode_buffer(tamanho + 1)
                user32.GetWindowTextW(alvo, texto, tamanho + 1)
                if trecho.lower() in texto.value.lower():
                    achou.append(True)
                    return False
        return True

    user32.EnumWindows(visitar, None)
    return bool(achou)


def pasta_logs():
    pasta = pasta_base() / "logs"
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

    pool = ThreadPoolExecutor(max_workers=1)
    pendentes = []  # (linha, esperada, futuro, recorte): conferencias de OCR em andamento, em ordem

    def revisar(aguardar=False):
        """Trata os resultados de OCR ja prontos (todos, com aguardar=True). Devolve a linha que nao conferiu, se houver."""
        restantes = []
        falha = None
        for linha, esperada, futuro, recorte in pendentes:
            if falha is None and (aguardar or futuro.done()):
                confere, lidas = futuro.result()
                registrar(f"  linha {linha}: quantidade lidas={lidas} esperada={esperada}", mostrar=False)
                if not confere:
                    pasta_ocr = pasta_logs() / "ocr"
                    pasta_ocr.mkdir(exist_ok=True)
                    recorte.save(pasta_ocr / f"{datetime.now():%H%M%S}-linha{linha}.png")
                    registrar(f"Linha {linha}: quantidade na tela {lidas} difere da planilha ({esperada}). Interrompido. "
                              "A linha ja foi confirmada no SRP: corrija-a (as seguintes, ate aqui, tambem podem estar afetadas). "
                              "Imagem lida em logs/ocr.")
                    falha = linha
            else:
                restantes.append((linha, esperada, futuro, recorte))
        pendentes[:] = restantes
        return falha

    def teclar(acao, valor, esperar=True):
        """`esperar=False` em Tab seguido de aguardar_mudanca (a propria espera do foco sincroniza)."""
        if parar.is_set() or mouse_no_canto():
            raise KeyboardInterrupt
        send_keys("{TAB}" if acao == "tab" else valor.replace(",", "{,}"), pause=0.05)
        if acao == "digitar":
            time.sleep(min(pausa, 0.1))
        elif esperar:
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
        passo = 19  # distancia vertical entre linhas do formulario (px); recalculada quando observavel
        coluna_qtd = None  # (esq, dir, altura) da coluna Quantidade, aprendida uma vez
        proximo = None  # caixa do Item da proxima linha (ja conhecida ao fim da anterior)
        for r in registros:
            if janela_trecho.lower() not in titulo_janela_ativa().lower():
                registrar(f"Foco saiu do Citrix antes da linha {r['linha']}. Interrompido.")
                return 1
            falha = revisar()
            if falha:
                if estado:
                    estado(falha, "erro")
                return 1
            marcar(r, "executando")
            t0 = time.time()
            inicial = proximo or caixa_foco(janela)
            teclar("digitar", r["item"]); teclar("tab", None, esperar=False)
            antes = aguardar_mudanca(janela, inicial, render)
            t1 = time.time()
            teclar("digitar", r["custo"]); teclar("tab", None, esperar=False)
            depois = aguardar_mudanca(janela, antes, render)
            t2 = time.time()
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
                if coluna_qtd is None:
                    foco = caixa_foco(janela)
                    if foco:
                        coluna_qtd = (foco[0], foco[2], foco[3] - foco[1])
            elif coluna_qtd is None:
                coluna_qtd = (depois[0], depois[2], depois[3] - depois[1])
            teclar("digitar", r["quantidade"].replace(".", ","))
            referencia = (coluna_qtd[0], inicial[1], coluna_qtd[1], inicial[1] + coluna_qtd[2]) if coluna_qtd else None
            teclar("tab", None, esperar=False)
            t3 = time.time()
            proximo = aguardar_mudanca(janela, referencia, render) if referencia else None
            t4 = t5 = time.time()
            if not proximo and not tesseract:
                time.sleep(pausa)  # sem deteccao nem conferencia: sincroniza por tempo
            if tesseract:
                # A celula so e lida depois do Tab (valor formatado, ex.: 1.0000): fica na linha acima do novo foco.
                if not proximo:
                    registrar(f"Linha {r['linha']}: nao consegui conferir a quantidade (foco nao mudou ou nao foi detectado). "
                              "Confira no SRP se ha aviso na tela. Interrompido.")
                    return 1
                if proximo[1] - inicial[1] > 8:
                    passo = proximo[1] - inicial[1]
                topo = proximo[1] - passo
                celula = (coluna_qtd[0], topo, coluna_qtd[1], topo + coluna_qtd[2])
                esperada = float(r["quantidade"].replace(",", "."))
                recorte = recortar_celula(janela, celula)
                pendentes.append((r["linha"], esperada, pool.submit(conferir_recorte, recorte, tesseract, esperada), recorte))
                t5 = time.time()
            registrar(f"  linha {r['linha']}: tempos (s) item_tab={t1 - t0:.2f} custo_tab={t2 - t1:.2f} "
                      f"digitar_conta_qtd={t3 - t2:.2f} resposta_tab_qtd={t4 - t3:.2f} ocr={t5 - t4:.2f}", mostrar=False)
            conta = "digitada " + r["conta"] if editavel else "travada"
            registrar(f"OK linha {r['linha']}: {r['item']} {r['custo']} conta={conta} {r['quantidade']}")
            marcar(r, "ok")
        falha = revisar(aguardar=True)
        if falha:
            if estado:
                estado(falha, "erro")
            return 1
        registrar("Concluido.")
        return 0
    except KeyboardInterrupt:
        registrar("Interrompido pelo usuario.")
        return 1
    except Exception as erro:
        registrar(f"Erro: {erro}")
        return 1
    finally:
        pool.shutdown(wait=False, cancel_futures=True)
        if atual is not None and estado:
            estado(atual, "erro")
        log.close()
