"""Logica do robo do SRP (via Citrix): leitura da planilha, validacao e envio de teclas.

Sequencia por linha: Item, Tab, C/Custo, Tab, [Conta, Tab se o campo estiver editavel], Quantidade, Tab.
A tela do Citrix e so imagem: o campo em foco e o bloco ciano; a Conta e editavel se o foco cai logo apos o C/Custo.
"""
import ctypes
import hashlib
import json
import os
import re
from collections import Counter
import subprocess
import sys
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

sys.coinit_flags = 2  # COM em modo STA, necessario para arrastar e soltar (tkdnd/OLE) e pywinauto coexistirem

from PIL import Image, ImageChops, ImageFilter, ImageGrab, ImageOps


def send_keys(*args, **kwargs):
    """pywinauto.keyboard.send_keys importado so na primeira tecla: o pywinauto leva ~1,3 s para carregar."""
    from pywinauto.keyboard import send_keys as enviar
    return enviar(*args, **kwargs)

COLUNAS = {"item": "item", "c/custo": "custo", "conta": "conta", "quantidade": "quantidade"}
OBRIGATORIAS = ("item", "custo", "quantidade")


def texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def ler_planilha(caminho):
    from openpyxl import load_workbook  # importado so ao ler a planilha (abre o app mais rapido)

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


CORES = {
    "ciano": (lambda v: v < 40, lambda v: v > 215, lambda v: v > 215),     # campo em foco
    "vermelho": (lambda v: v > 215, lambda v: v < 40, lambda v: v < 40),  # icone dos avisos do SRP
    # texto selecionado no campo em foco (SNAC0236: a Data Emissao que o SRP preenche fica branca com a data selecionada,
    # sem ciano); medido (72, 100, 141) na tela 4K em 04/11/2026
    "selecao": (lambda v: 60 <= v <= 85, lambda v: 88 <= v <= 112, lambda v: 128 <= v <= 155),
    # barra de titulo das janelas do SRP (a principal e as que abrem por cima: avisos, Resumo dos Tributos, Frete LTCC);
    # medido (45, 132, 197) uniforme nos prints e na tela 4K em 04/11/2026
    "titulo_srp": (lambda v: 40 <= v <= 50, lambda v: 127 <= v <= 137, lambda v: 192 <= v <= 202),
}


def _mascara_cor(imagem, cor="ciano"):
    """(mascara erodida, deslocamento x, y) so da area onde a cor aparece: o filtro de erosao na tela inteira custava
    ~50 ms por consulta; sem a cor na tela (o vermelho do aviso, quase sempre) devolve (None, 0, 0) na hora."""
    r, g, b = imagem.convert("RGB").split()
    tr, tg, tb = CORES[cor]
    mascara = ImageChops.multiply(ImageChops.multiply(r.point(lambda v: 255 if tr(v) else 0),
                                                      g.point(lambda v: 255 if tg(v) else 0)),
                                  b.point(lambda v: 255 if tb(v) else 0))
    caixa = mascara.getbbox()
    if not caixa:
        return None, 0, 0
    esq, topo = max(caixa[0] - 3, 0), max(caixa[1] - 3, 0)
    recorte = mascara.crop((esq, topo, min(caixa[2] + 3, mascara.width), min(caixa[3] + 3, mascara.height)))
    return recorte.filter(ImageFilter.MinFilter(5)), esq, topo


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
    return maior_caixa_cor(imagem, "ciano")


def caixas_cor(imagem, cor):
    """Retangulos (esq, topo, dir, base) dos blocos da cor pura; ignora pixels soltos.
    Trabalha por faixas (getbbox/resize em C): ~100x mais rapido que percorrer pixel a pixel, e nao trava a janela."""
    mascara, dx, dy = _mascara_cor(imagem, cor)
    caixa_total = mascara.getbbox() if mascara else None
    if not caixa_total:
        return []
    esq0, topo0 = caixa_total[0] + dx, caixa_total[1] + dy
    regiao = mascara.crop(caixa_total)
    largura, altura = regiao.size
    linhas = [regiao.crop((0, y, largura, y + 1)).getbbox() is not None for y in range(altura)]
    caixas = []
    for y_ini, y_fim in _trechos(linhas):
        faixa = regiao.crop((0, y_ini, largura, y_fim))
        colunas = list(faixa.resize((largura, 1), Image.BOX).tobytes())
        for x_ini, x_fim in _trechos(colunas):
            bloco = faixa.crop((x_ini, 0, x_fim, y_fim - y_ini)).getbbox()
            if bloco:
                caixas.append((esq0 + x_ini + bloco[0], topo0 + y_ini + bloco[1], esq0 + x_ini + bloco[2], topo0 + y_ini + bloco[3]))
    return caixas


def maior_caixa_cor(imagem, cor):
    """Maior bloco da cor pura (ex.: o campo em foco, ciano)."""
    caixas = caixas_cor(imagem, cor)
    return max(caixas, key=lambda c: (c[2] - c[0]) * (c[3] - c[1])) if caixas else None


def _sobrepoe(caixa, outras, folga=4):
    esq, topo, dir_, base = caixa
    return any(esq < o[2] + folga and dir_ > o[0] - folga and topo < o[3] + folga and base > o[1] - folga for o in outras)


def icone_aviso(imagem, fixos=()):
    """Icone vermelho do aviso do SRP: bloco vermelho NOVO, quase quadrado, de 8 a 60 px (ja erodido pelo filtro).
    `fixos` = blocos vermelhos que ja estavam na tela no inicio (`caixas_cor` da calibracao): o cabecalho do SRP tem
    uma faixa vermelha que o logo corta, e o pedaco da esquerda (~17x21 px, na borda) parecia um icone; na v1.8.0 isso
    fazia o robo pausar em todo campo. Tambem descarta blocos encostados na borda da janela."""
    largura, altura = imagem.size
    caixas = caixas_cor(imagem, "vermelho")
    # faixa do titulo do SRP (bloco vermelho largo): os pedacos picotados pelo logo/letras ao lado dela nao sao icone.
    # Visto em 04/11/2026 na SNAC0236 em monitor 4K: um pedaco 8x8 entre o logo e a faixa barrava o inicio.
    faixas = [c for c in caixas if c[2] - c[0] >= 200]
    candidatos = [c for c in caixas
                  if 8 <= c[2] - c[0] <= 60 and 8 <= c[3] - c[1] <= 60 and 0.6 <= (c[2] - c[0]) / (c[3] - c[1]) <= 1.6
                  and c[0] > 3 and c[1] > 3 and c[2] < largura - 3 and c[3] < altura - 3 and not _sobrepoe(c, fixos)
                  and not _sobrepoe(c, faixas, folga=60)]
    return max(candidatos, key=lambda c: (c[2] - c[0]) * (c[3] - c[1])) if candidatos else None


def problema_tela(retangulo, imagem):
    """Checagem antes de digitar: motivo (texto para o usuario) se a tela do Citrix nao serve para comecar, senao None.
    `retangulo` = (esq, topo, dir, base) da janela; `imagem` = captura dela."""
    esq, topo, dir_, base = retangulo
    if esq <= -30000 or dir_ - esq < 400 or base - topo < 300:
        return "a janela do Citrix esta minimizada ou pequena demais. Maximize o Citrix com o SRP aberto"
    menor, maior = imagem.convert("L").getextrema()
    if maior - menor < 40:
        return "a tela do Citrix esta vazia (sessao bloqueada, desconectada ou carregando). Desbloqueie o Citrix e abra o SRP"
    if icone_aviso(imagem):
        return "ja ha um aviso aberto no SRP. Feche o aviso (OK) e deixe o cursor no campo Item"
    return None


def capturar(janela):
    ret = janela.rectangle()
    return ImageGrab.grab(bbox=(ret.left, ret.top, ret.right, ret.bottom), all_screens=True)


def caixa_foco(janela):
    """Retangulo do campo em foco (ciano) na janela do Citrix."""
    return maior_caixa_ciano(capturar(janela))


def estado_tela(janela, fixos=()):
    """(caixa do foco ciano, ha aviso do SRP na tela) a partir de uma unica captura; `fixos` como em icone_aviso."""
    imagem = capturar(janela)
    return maior_caixa_ciano(imagem), icone_aviso(imagem, fixos) is not None


def caminho_tesseract():
    """Tesseract portatil: pasta `tesseract` ou `vendor/tesseract` ao lado do exe (ou da raiz do projeto, no codigo-fonte)."""
    for pasta in (pasta_base() / "tesseract", pasta_base() / "vendor" / "tesseract"):
        if (pasta / "tesseract.exe").exists():
            return pasta / "tesseract.exe"
    return None


def preparar_ocr(imagem, escala=6, limiar=120):
    """Preto e branco (descarta a moldura clara da celula e o fundo ciano), ampliacao e margem branca para o Tesseract.
    Usa o canal de cor mais escuro de cada pixel, nao a luminancia: o ClearType desenha os digitos com franjas
    coloridas e, pela luminancia, parte do "5" sumia (45.0000 era lido como 4.0)."""
    r, g, b = imagem.convert("RGB").split()
    pb = ImageChops.darker(ImageChops.darker(r, g), b).point(lambda v: 0 if v < limiar else 255)
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
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
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


def aguardar_mudanca(janela, referencia, limite, ver_aviso=False, fixos=()):
    """Espera o foco (ciano) sair de `referencia`; devolve a nova caixa, None no tempo esgotado ou, com ver_aviso,
    a palavra "aviso" assim que o aviso do SRP (icone vermelho novo, fora de `fixos`) aparecer."""
    prazo = time.time() + limite
    while time.time() < prazo:
        tem_aviso = False
        if ver_aviso:
            atual, tem_aviso = estado_tela(janela, fixos)
        else:
            atual = caixa_foco(janela)
        if atual and referencia and (abs(atual[0] - referencia[0]) > 5 or abs(atual[1] - referencia[1]) > 5):
            return atual  # o foco andou: o SRP aceitou o campo (um aviso de verdade prende o foco)
        if tem_aviso:
            return "aviso"
        time.sleep(0.02)  # cada volta ja custa ~60 ms (captura + busca das cores)
    return None


def titulo_janela_ativa():
    """Titulo da janela em primeiro plano (chamada direta ao Windows; a busca por UIA leva ~1 s)."""
    alvo = ctypes.windll.user32.GetForegroundWindow()
    tamanho = ctypes.windll.user32.GetWindowTextLengthW(alvo)
    texto = ctypes.create_unicode_buffer(tamanho + 1)
    ctypes.windll.user32.GetWindowTextW(alvo, texto, tamanho + 1)
    return texto.value


def palavras_tela(imagem, tesseract, escala=2):
    """[(texto, esq, topo, larg, alt)] das palavras da imagem (Tesseract em modo de texto esparso)."""
    cinza = imagem.convert("L").resize((imagem.width * escala, imagem.height * escala), Image.LANCZOS)
    with tempfile.TemporaryDirectory() as pasta:
        arquivo = Path(pasta) / "tela.png"
        cinza.save(arquivo)
        saida = subprocess.run(
            [str(tesseract), str(arquivo), "stdout", "--tessdata-dir", str(tesseract.parent / "tessdata"), "--psm", "11", "tsv"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    palavras = []
    for linha in saida.stdout.splitlines()[1:]:
        campos = linha.split("\t")
        if len(campos) == 12 and campos[11].strip():
            esq, topo, larg, alt = (int(v) // escala for v in campos[6:10])
            palavras.append((campos[11].strip(), esq, topo, larg, alt))
    return palavras


def ler_aviso(imagem, tesseract, fixos=()):
    """Le a janela de aviso do SRP (localizada pelo icone vermelho). Devolve {'texto', 'item', 'custo', 'conta_correta'}
    ou None. O OCR erra digitos de vez em quando (8712 -> 3712): le de 5 jeitos e so aceita a conta da maioria."""
    icone = icone_aviso(imagem, fixos)
    if not icone:
        return None
    esq, topo, dir_, base = icone
    f = min(max(dir_ - esq, 10) / 19, 2.5)  # o icone detectado tem ~19 px a 100% de escala; a janela de aviso tem ~380 x 110
    recorte = imagem.crop((max(0, esq - int(40 * f)), max(0, topo - int(60 * f)),
                           min(imagem.width, esq + int(420 * f)), min(imagem.height, base + int(60 * f))))
    r, g, b = recorte.convert("RGB").split()
    escuro = ImageChops.darker(ImageChops.darker(r, g), b).point(lambda v: 0 if v < 120 else 255).convert("RGB")
    def ler(img, escala):
        linhas = {}
        for texto, x, y, _, _ in palavras_tela(img, tesseract, escala):
            linhas.setdefault(round(y / 6), []).append((x, texto))
        return " ".join(" ".join(t for _, t in sorted(l)) for _, l in sorted(linhas.items()))

    variantes = ((recorte, 2), (recorte, 4), (recorte, 5), (escuro, 4), (escuro, 5))
    with ThreadPoolExecutor(max_workers=len(variantes)) as paralelo:  # cada leitura e um processo do Tesseract
        leituras = list(paralelo.map(lambda v: ler(*v), variantes))
    def votar(padrao):
        achados = [m.group(1) for m in (re.search(padrao, t, re.I | re.S) for t in leituras) if m]
        if not achados:
            return None
        valor, votos = Counter(achados).most_common(1)[0]
        return valor if votos >= 3 else None
    texto = max(leituras, key=len)
    texto = re.sub(r"^.*?(?=conta|o saldo|saldo)", "", texto, count=1, flags=re.I | re.S) or texto
    palavras_texto = texto.split()
    com_numero = [i for i, w in enumerate(palavras_texto) if re.search(r"\d", w)]
    if com_numero:  # depois do ultimo numero vem o botao OK lido como letras soltas
        texto = " ".join(palavras_texto[:com_numero[-1] + 1])
    return {"texto": texto.strip(), "item": votar(r"item\s*(\d{3,})"), "custo": votar(r"c\W?\s*custo\s*(\d{2,})"),
            "conta_correta": votar(r"correta\s*=?\s*(\d{2,})")}


def arquivo_aprendizado():
    return pasta_base() / "aprendizado.json"


def ler_aprendizado():
    try:
        return json.loads(arquivo_aprendizado().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"contas": {}}


def aprender_conta(item, custo, conta):
    """Guarda a conta que o SRP informou como correta para item + C/Custo (so um aviso futuro; nada e corrigido sozinho)."""
    dados = ler_aprendizado()
    dados.setdefault("contas", {})[f"{item}|{custo}"] = {"conta": conta, "em": datetime.now().strftime("%Y-%m-%d %H:%M")}
    try:
        arquivo_aprendizado().write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def contas_suspeitas(registros):
    """[(linha, item, conta_da_planilha, conta_correta)] das linhas cuja conta o SRP ja recusou para o mesmo item + C/Custo."""
    contas = ler_aprendizado().get("contas", {})
    suspeitas = []
    for r in registros:
        certo = contas.get(f"{r['item']}|{r['custo']}")
        if certo and r["conta"] and r["conta"] != certo["conta"]:
            suspeitas.append((r["linha"], r["item"], r["conta"], certo["conta"]))
    return suspeitas


def aplicar_sugestoes(registros, linhas):
    """Troca a Conta das `linhas` que o USUARIO aceitou pela que o SRP informou como correta (ver contas_suspeitas).
    So na memoria: a planilha no disco nao muda. Devolve [(linha, conta_antes, conta_nova)]."""
    sugeridas = {linha: correta for linha, _, _, correta in contas_suspeitas(registros)}
    trocas = []
    for r in registros:
        if r["linha"] in linhas and r["linha"] in sugeridas:
            trocas.append((r["linha"], r["conta"], sugeridas[r["linha"]]))
            r["conta"] = sugeridas[r["linha"]]
    return trocas


def localizar_colunas(imagem, tesseract):
    """Centro horizontal das colunas C/Custo, Conta e Quantidade (e Item, se lido) pelo cabecalho da grade.
    Devolve {'custo': x, 'conta': x, 'quantidade': x[, 'item': x]} ou None se o cabecalho nao for achado."""
    palavras = palavras_tela(imagem, tesseract)
    quant = [p for p in palavras if p[0].lower().startswith("quant")]
    for _, qx, qy, qw, _ in quant:
        mesma_linha = [p for p in palavras if abs(p[2] - qy) <= 6]
        colunas = {"quantidade": qx + qw / 2}
        for texto, x, _, w, _ in mesma_linha:
            t = texto.lower()
            if "custo" in t:
                colunas["custo"] = x + w / 2
            elif t == "conta":
                colunas["conta"] = x + w / 2
            elif t == "item":
                colunas["item"] = x + w / 2
        if "custo" in colunas and "conta" in colunas:
            return colunas
    return None


def qual_coluna(caixa, colunas):
    """Nome da coluna cujo centro do cabecalho fica mais perto do centro da caixa em foco (None se nenhuma servir)."""
    centro = (caixa[0] + caixa[2]) / 2
    nome, x = min(colunas.items(), key=lambda item: abs(item[1] - centro))
    return nome if abs(x - centro) <= max(40, caixa[2] - caixa[0]) else None


def conta_editavel(antes, depois):
    """Conta editavel se o novo foco comeca logo apos o C/Custo (vao menor que 1,5 largura)."""
    largura = antes[2] - antes[0]
    return (depois[0] - antes[2]) < largura * 1.5


VK_F8, VK_F9 = 0x77, 0x78


def atalhos_globais(acoes):
    """Teclas globais do Windows (funcionam com o foco em outra janela): `acoes` = {codigo_vk: funcao}.
    As funcoes rodam na thread das teclas. Devolve (soltar, teclas_que_falharam); chamar soltar() libera as teclas.
    Se o Citrix capturar o teclado (tela cheia), as teclas podem nao chegar: o mouse no canto continua valendo."""
    import ctypes.wintypes
    user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
    pronto, info = threading.Event(), {"falhas": list(acoes)}
    mapa = dict(enumerate(acoes, 1))

    def laco():
        info["thread"] = kernel32.GetCurrentThreadId()
        registrados = [i for i, vk in mapa.items() if user32.RegisterHotKey(None, i, 0x4000, vk)]  # 0x4000 = sem repeticao
        info["falhas"] = [vk for i, vk in mapa.items() if i not in registrados]
        pronto.set()
        msg = ctypes.wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == 0x0312 and msg.wParam in mapa:  # WM_HOTKEY
                acoes[mapa[msg.wParam]]()
        for i in registrados:
            user32.UnregisterHotKey(None, i)

    threading.Thread(target=laco, daemon=True).start()
    pronto.wait(2)

    def soltar():
        if "thread" in info:
            user32.PostThreadMessageW(info["thread"], 0x0012, 0, 0)  # WM_QUIT encerra o laco e libera as teclas

    return soltar, info["falhas"]


def mouse_no_canto():
    class Ponto(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
    ponto = Ponto()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(ponto))
    return ponto.x <= 2 and ponto.y <= 2


def janela_citrix(trecho):
    # Acha pelo EnumWindows (o mesmo do chip "Citrix aberto") e so depois embrulha no pywinauto:
    # a lista de janelas do UIA (Desktop.windows()) chegou a vir sem o Citrix com ele aberto (v1.9.2, 30/10 17:20).
    from pywinauto import Desktop  # importado so ao iniciar (abre o app mais rapido)

    for alvo in _janelas_com(trecho):
        for backend in ("uia", "win32"):
            try:
                return Desktop(backend=backend).window(handle=alvo).wrapper_object()
            except Exception:
                continue
    raise LookupError(f"Janela com '{trecho}' no titulo nao encontrada.")


def pasta_base():
    """Pasta do .exe quando empacotado; raiz do projeto no codigo-fonte."""
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent


def _janelas_com(trecho):
    """Handles das janelas visiveis com `trecho` no titulo (EnumWindows: rapido; nao disputa a janela do app)."""
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
                    achou.append(alvo)
        return True

    user32.EnumWindows(visitar, None)
    return achou


def citrix_aberto(trecho="Desktop Viewer"):
    """Ha janela visivel com `trecho` no titulo?"""
    return bool(_janelas_com(trecho))


def pasta_logs():
    pasta = pasta_base() / "logs"
    pasta.mkdir(exist_ok=True)
    return pasta


def arquivo_progresso():
    return pasta_base() / "progresso.json"


def assinatura_planilha(registros):
    """Identifica a planilha pelo conteudo (nao pelo caminho): mudou uma celula, o andamento salvo nao vale mais."""
    texto = json.dumps([[r["linha"], r["item"], r["custo"], r["conta"], r["quantidade"]] for r in registros])
    return hashlib.sha1(texto.encode("utf-8")).hexdigest()


def gravar_progresso(registros, linhas_ok):
    """Guarda as linhas concluidas da planilha atual (so a ultima planilha; sem linhas, apaga o andamento)."""
    try:
        if linhas_ok:
            dados = {"assinatura": assinatura_planilha(registros), "ok": sorted(linhas_ok),
                     "em": datetime.now().strftime("%d/%m/%Y %H:%M")}
            arquivo_progresso().write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            arquivo_progresso().unlink(missing_ok=True)
    except OSError:
        pass


def ler_progresso(registros):
    """(linhas concluidas, quando) salvas para esta mesma planilha, ou (set(), None)."""
    try:
        dados = json.loads(arquivo_progresso().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return set(), None
    if dados.get("assinatura") != assinatura_planilha(registros):
        return set(), None
    validas = {r["linha"] for r in registros}
    return {l for l in dados.get("ok", []) if l in validas}, dados.get("em")


REPOSITORIO = "bwernerdev/rpa-controladoria"  # publico: os releases sao lidos sem token


def versao_tupla(texto):
    """'v1.10.2' -> (1, 10, 2); texto invalido -> (0,)."""
    try:
        return tuple(int(p) for p in texto.strip().lstrip("vV").split("."))
    except ValueError:
        return (0,)


def buscar_atualizacao(atual, tempo_limite=5):
    """Release mais nova que `atual` no GitHub: {'versao', 'url', 'sha256', 'tamanho', 'notas'}; None se nao houver,
    se nao tiver o zip ou sem internet (nunca levanta erro: a checagem nao pode atrapalhar o app)."""
    import urllib.request
    pedido = urllib.request.Request(f"https://api.github.com/repos/{REPOSITORIO}/releases/latest",
                                    headers={"Accept": "application/vnd.github+json", "User-Agent": "RPA-SRP"})
    try:
        with urllib.request.urlopen(pedido, timeout=tempo_limite) as resposta:
            dados = json.load(resposta)
    except Exception:
        return None
    versao = str(dados.get("tag_name", "")).lstrip("vV")
    if versao_tupla(versao) <= versao_tupla(atual):
        return None
    zips = [a for a in dados.get("assets", []) if str(a.get("name", "")).lower().endswith(".zip")]
    if not zips:
        return None
    digest = str(zips[0].get("digest") or "")
    return {"versao": versao, "url": zips[0]["browser_download_url"], "tamanho": zips[0].get("size"),
            "sha256": digest.split(":", 1)[1] if digest.startswith("sha256:") else None, "notas": dados.get("body") or ""}


def baixar_atualizacao(info, progresso=None):
    """Baixa e extrai o zip do release numa pasta temporaria; confere o SHA-256 publicado pelo GitHub.
    `progresso(baixados, total)` e opcional. Devolve a pasta com instalar.ps1 e RPA-SRP.exe."""
    import urllib.request
    import zipfile
    pasta = Path(tempfile.mkdtemp(prefix="rpa-srp-"))
    arquivo = pasta / "pacote.zip"
    soma = hashlib.sha256()
    pedido = urllib.request.Request(info["url"], headers={"User-Agent": "RPA-SRP"})
    with urllib.request.urlopen(pedido, timeout=30) as resposta, open(arquivo, "wb") as saida:
        total, baixados = info.get("tamanho") or int(resposta.headers.get("Content-Length") or 0), 0
        while True:
            bloco = resposta.read(256 * 1024)
            if not bloco:
                break
            saida.write(bloco)
            soma.update(bloco)
            baixados += len(bloco)
            if progresso:
                progresso(baixados, total)
    if info.get("sha256") and soma.hexdigest() != info["sha256"].lower():
        raise ValueError("o arquivo baixado nao confere com o publicado (SHA-256 diferente); atualizacao cancelada")
    destino = pasta / "pacote"
    with zipfile.ZipFile(arquivo) as zip_:
        zip_.extractall(destino)  # extractall descarta caminhos absolutos e ".."
    if not (destino / "instalar.ps1").exists() or not (destino / "RPA-SRP.exe").exists():
        raise ValueError("o pacote baixado nao tem instalar.ps1 e RPA-SRP.exe")
    return destino


def aplicar_atualizacao(pacote, destino):
    """Roda o instalar.ps1 do pacote num processo separado: ele espera este app fechar, troca exe e Tesseract em
    `destino` (mantendo config, contas aprendidas, andamento e logs) e abre o app de novo. Chamar e fechar o app."""
    log = Path(destino) / "logs" / "atualizacao.log"
    log.parent.mkdir(parents=True, exist_ok=True)
    # o exe novo, aberto no fim pelo script, nao pode herdar as variaveis do PyInstaller deste app
    # (senao procura o python314.dll na pasta _MEI deste, apagada quando ele fecha)
    ambiente = {k: v for k, v in os.environ.items() if not k.upper().startswith(("_PYI_", "_MEIPASS"))}
    ambiente["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    subprocess.Popen(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File",
         str(Path(pacote) / "instalar.ps1"), "-Destino", str(destino), "-SemAtalhos", "-Esperar", "60", "-Abrir",
         "-Log", str(log)],
        # console oculto (CREATE_NO_WINDOW); DETACHED_PROCESS faz o PowerShell 5.1 sair sem rodar nada
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
        close_fds=True, env=ambiente)


def limpar_logs(dias=30):
    """Apaga logs (e recortes de OCR em logs/ocr) com mais de `dias` dias; devolve quantos apagou."""
    limite = time.time() - dias * 86400
    apagados = 0
    for arquivo in [*pasta_logs().glob("*.log"), *pasta_logs().glob("ocr/*.png")]:
        try:
            if arquivo.stat().st_mtime < limite:
                arquivo.unlink()
                apagados += 1
        except OSError:
            pass
    return apagados


ORDEM = ("item", "custo", "conta", "quantidade")
NOME_CAMPO = {"item": "Item", "custo": "C/Custo", "conta": "Conta", "quantidade": "Quantidade"}


def proximo_campo(atual, caixa_antes, caixa_depois, colunas):
    """Campo onde o foco caiu depois do Tab em `atual`: 'custo'/'conta'/'quantidade', 'fim' (Item da linha seguinte)
    ou None se o foco foi para um lugar inesperado. Sem cabecalho (`colunas` None) usa a regra antiga."""
    if colunas:
        onde = qual_coluna(caixa_depois, colunas)
        if atual == "quantidade":
            return "fim" if onde == "item" else None
        return onde if onde in ORDEM and ORDEM.index(onde) > ORDEM.index(atual) else None
    if atual == "item":
        return "custo"
    if atual == "custo":
        return "conta" if conta_editavel(caixa_antes, caixa_depois) else "quantidade"
    return "quantidade" if atual == "conta" else "fim"


def executar(registros, janela_trecho, pausa, render, espera, parar, aviso, estado=None, conferir=False, retomar=None,
             pausar=None):
    """Preenche o SRP. `parar` e um threading.Event; `aviso(texto)` recebe o andamento;
    `estado(linha, nome)` (opcional) recebe 'executando', 'pausado', 'ok' ou 'erro' por linha;
    `conferir` le a Quantidade na tela (OCR) antes de confirmar cada linha;
    `retomar` (threading.Event, opcional): se o SRP nao avancar (aviso na tela), o robo PAUSA e espera esse
    evento (botao Continuar) em vez de interromper. Sem ele, interrompe como antes.
    `pausar` (threading.Event, opcional; F9): pausa ao terminar a linha atual e espera `retomar`. Devolve 0 se concluiu."""
    tesseract = caminho_tesseract() if conferir else None
    limpar_logs()
    log = open(pasta_logs() / f"{datetime.now():%Y%m%d-%H%M%S}.log", "w", encoding="utf-8")
    atual = None

    def marcar(r, nome):
        nonlocal atual
        atual = r["linha"] if nome == "executando" else None
        if estado:
            estado(r["linha"], nome)

    trava_log = threading.Lock()

    def registrar(texto, mostrar=True):
        with trava_log:
            if log.closed:
                return
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
                    falha = (linha, lidas, esperada)
            else:
                restantes.append((linha, esperada, futuro, recorte))
        pendentes[:] = restantes
        return falha

    def teclar(acao, valor, esperar=True):
        """`esperar=False` em Tab seguido de aguardar_mudanca (a propria espera do foco sincroniza)."""
        if parar.is_set() or mouse_no_canto():
            raise KeyboardInterrupt
        # 0,03 s por tecla (era 0,05): o Citrix recebe as teclas em ordem; a sincronizacao real e o foco mudar apos o Tab
        send_keys("{TAB}" if acao == "tab" else valor.replace(",", "{,}"), pause=0.03)
        if acao == "digitar":
            time.sleep(min(pausa, 0.05))
        elif esperar:
            time.sleep(pausa)

    def esperar_continuar():
        retomar.clear()
        while not retomar.is_set():
            if parar.is_set():
                raise KeyboardInterrupt
            time.sleep(0.2)

    def tratar_divergencia(falha, no_fim=False):
        """Quantidade lida diferente da planilha: pausa (se houver Continuar) ou interrompe. True = seguir."""
        linha, lidas, esperada = falha
        if retomar is None:
            registrar(f"Linha {linha}: quantidade na tela {lidas} difere da planilha ({esperada}). Interrompido. "
                      "A linha ja foi confirmada no SRP: corrija-a (as seguintes, ate aqui, tambem podem estar afetadas). "
                      "Imagem lida em logs/ocr.")
            if estado:
                estado(linha, "erro")
            return False
        registrar(f"Linha {linha}: a quantidade lida na tela {lidas} difere da planilha ({esperada}). PAUSADO. "
                  f"Confira a linha {linha} no SRP (imagem lida em logs/ocr). Se estiver certa, clique em Continuar. "
                  "Se nao, corrija-a" + ("." if no_fim else ", volte o cursor para o Item da proxima linha livre") +
                  " e clique em Continuar.")
        if estado:
            estado(linha, "pausado")
        while True:
            esperar_continuar()
            if no_fim:
                break
            janela.set_focus()
            aviso("Continuando em 2 s...")
            for _ in range(20):
                if parar.is_set():
                    raise KeyboardInterrupt
                time.sleep(0.1)
            caixa = caixa_foco(janela)
            if caixa and (not colunas or qual_coluna(caixa, colunas) == "item"):
                break
            registrar("Ainda pausado: deixe o cursor no Item da proxima linha livre e clique em Continuar.")
            if estado:
                estado(linha, "pausado")
        registrar(f"  linha {linha}: quantidade conferida pelo usuario apos a pausa", mostrar=False)
        if estado:
            estado(linha, "ok")
        return True

    def pausa_pedida(r):
        """F9: pausa antes de comecar a linha `r`; volta quando o cursor estiver no Item e o usuario continuar."""
        pausar.clear()
        registrar(f"Pausado antes da linha {r['linha']} (F9). Para continuar: deixe o cursor no Item da linha "
                  f"{r['linha']} no SRP e clique em Continuar (ou F9).")
        if estado:
            estado(r["linha"], "pausado")
        while True:
            esperar_continuar()
            pausar.clear()
            janela.set_focus()
            aviso("Continuando em 2 s...")
            for _ in range(20):
                if parar.is_set():
                    raise KeyboardInterrupt
                time.sleep(0.1)
            caixa = caixa_foco(janela)
            if caixa and (not colunas or qual_coluna(caixa, colunas) == "item"):
                return caixa
            registrar(f"Ainda pausado: deixe o cursor no Item da linha {r['linha']} e clique em Continuar (ou F9).")
            if estado:
                estado(r["linha"], "pausado")

    def esperar_correcao(r, campo_erro):
        """Pausa ate o usuario corrigir no SRP e clicar em Continuar. Devolve (caixa do foco, campo onde retomar);
        'fim' quando o cursor ja esta no Item (linha completada a mao)."""
        if exe_tesseract:
            tela_do_aviso = capturar(janela)  # captura ja: o usuario pode fechar o aviso a qualquer momento

            def ler_em_segundo_plano():
                try:
                    info = ler_aviso(tela_do_aviso, exe_tesseract, fixos)
                except Exception as erro:
                    registrar(f"  leitura do aviso falhou: {erro}", mostrar=False)
                    return
                if not info:
                    return
                registrar(f'Linha {r["linha"]}: SRP: "{info["texto"]}"')
                if info["conta_correta"] and info["item"] == r["item"] and (info["custo"] in (None, r["custo"])):
                    aprender_conta(r["item"], r["custo"], info["conta_correta"])
                    registrar(f"Linha {r['linha']}: o SRP informou a Conta correta {info['conta_correta']} para o item "
                              f"{r['item']} (C/Custo {r['custo']}); guardada para avisar nas proximas planilhas.")
            threading.Thread(target=ler_em_segundo_plano, daemon=True).start()
        registrar(f"Linha {r['linha']}: o SRP nao avancou depois do {NOME_CAMPO[campo_erro]}. PAUSADO. "
                  f"No SRP: feche o aviso, corrija o {NOME_CAMPO[campo_erro]}, aperte Tab para ir ao campo seguinte "
                  "e clique em Continuar. Se completou a linha a mao, deixe o cursor no Item da linha seguinte.")
        if estado:
            estado(r["linha"], "pausado")
        while True:
            esperar_continuar()
            janela.set_focus()
            aviso("Continuando em 2 s...")
            for _ in range(20):
                if parar.is_set():
                    raise KeyboardInterrupt
                time.sleep(0.1)
            caixa = caixa_foco(janela)
            onde = qual_coluna(caixa, colunas) if caixa else None
            if onde == "item":
                registrar(f"  linha {r['linha']}: retomada com o cursor no Item (linha completada a mao)", mostrar=False)
                if estado:
                    estado(r["linha"], "executando")
                return caixa, "fim"
            if onde in ORDEM and ORDEM.index(onde) > ORDEM.index(campo_erro):
                registrar(f"  linha {r['linha']}: retomada no campo {NOME_CAMPO[onde]}", mostrar=False)
                if estado:
                    estado(r["linha"], "executando")
                return caixa, onde
            registrar(f"Ainda pausado: o cursor esta em '{NOME_CAMPO.get(onde, 'campo nao identificado')}'. "
                      f"Deixe-o no campo depois do {NOME_CAMPO[campo_erro]} (ou no Item da linha seguinte) e clique em Continuar.")
            if estado:
                estado(r["linha"], "pausado")

    try:
        if conferir and not tesseract:
            aviso("ATENCAO: pasta 'tesseract' nao encontrada ao lado do programa. Conferencia da quantidade DESATIVADA nesta execucao.")
        janela = janela_citrix(janela_trecho)
        janela.set_focus()
        aviso(f"Clique no campo Item da primeira linha. Iniciando em {espera}s (mouse no canto superior esquerdo aborta).")
        inicio_contagem = time.time()
        colunas = None
        exe_tesseract = caminho_tesseract()
        tela_inicial = capturar(janela)
        ret = janela.rectangle()
        problema = problema_tela((ret.left, ret.top, ret.right, ret.bottom), tela_inicial)
        if problema:
            registrar(f"Nao iniciei: {problema}. Nada foi digitado.")
            return 1
        try:  # vermelhos que ja estao na tela (faixa do cabecalho do SRP): nao sao aviso
            fixos = caixas_cor(tela_inicial, "vermelho")
        except Exception:
            fixos = []
        registrar(f"  vermelhos fixos na tela: {fixos}", mostrar=False)
        if exe_tesseract:
            try:
                colunas = localizar_colunas(tela_inicial, exe_tesseract)
            except Exception as erro:  # sem cabecalho, segue no modo antigo
                registrar(f"  leitura do cabecalho falhou: {erro}", mostrar=False)
        if colunas:
            registrar(f"  colunas pelo cabecalho: {colunas}", mostrar=False)
        else:
            aviso("ATENCAO: nao localizei o cabecalho da grade (C/Custo, Conta, Quantidade). "
                  "Sem ele o robo nao detecta C/Custo travado: o C/Custo sera sempre digitado.")
        while time.time() - inicio_contagem < espera:
            if parar.is_set():
                registrar("Interrompido antes de iniciar.")
                return 1
            time.sleep(0.1)
        # fim da contagem: o cursor tem que estar no Item, senao nada e digitado
        proximo = caixa_foco(janela)  # caixa do Item da proxima linha (ja conhecida ao fim da anterior)
        if not proximo:
            registrar("Nao iniciei: nenhum campo do SRP em foco (ciano). Clique no campo Item da primeira linha livre "
                      "e inicie de novo. Nada foi digitado.")
            return 1
        if colunas and "item" in colunas and qual_coluna(proximo, colunas) != "item":
            onde = qual_coluna(proximo, colunas)
            registrar(f"Nao iniciei: o cursor esta {f'na coluna {NOME_CAMPO[onde]}' if onde else 'fora da grade'}, "
                      "nao no Item. Clique no campo Item da primeira linha livre e inicie de novo. Nada foi digitado.")
            return 1
        passo = 19  # distancia vertical entre linhas do formulario (px); recalculada quando observavel
        coluna_qtd = None  # (esq, dir, altura) da coluna Quantidade, aprendida uma vez
        for r in registros:
            if janela_trecho.lower() not in titulo_janela_ativa().lower():
                registrar(f"Foco saiu do Citrix antes da linha {r['linha']}. Interrompido.")
                return 1
            falha = revisar()
            while falha:
                if not tratar_divergencia(falha):
                    return 1
                falha = revisar()
            if pausar is not None and pausar.is_set() and retomar is not None:
                proximo = pausa_pedida(r)
            marcar(r, "executando")
            t0 = time.time()
            inicial = proximo or caixa_foco(janela)
            if not inicial:
                registrar(f"Linha {r['linha']}: campo em foco (ciano) nao detectado. Interrompido.")
                return 1
            foco, campo = inicial, "item"
            digitados, manual, proximo = [], False, None
            while campo != "fim":
                if campo == "conta" and not r["conta"]:
                    registrar(f"Linha {r['linha']}: Conta editavel mas em branco na planilha. Interrompido.")
                    return 1
                if campo == "quantidade" and coluna_qtd is None:
                    coluna_qtd = (foco[0], foco[2], foco[3] - foco[1])
                valor = r["quantidade"].replace(".", ",") if campo == "quantidade" else r[campo]
                teclar("digitar", valor); teclar("tab", None, esperar=False)
                digitados.append(campo)
                novo = aguardar_mudanca(janela, foco, render, ver_aviso=bool(colunas), fixos=fixos)
                if novo == "aviso":
                    novo = None
                seguinte = proximo_campo(campo, foco, novo, colunas) if novo else None
                if seguinte is None and campo == "quantidade" and not tesseract and not colunas:
                    time.sleep(pausa)  # sem deteccao nem conferencia: sincroniza por tempo (modo antigo)
                    break
                if seguinte is None:
                    if retomar is None or not colunas:
                        registrar(f"Linha {r['linha']}: o SRP nao avancou depois do {NOME_CAMPO[campo]} "
                                  "(aviso na tela?). Interrompido.")
                        return 1
                    novo, seguinte = esperar_correcao(r, campo)
                    manual = True
                foco, campo = novo, seguinte
                if campo == "fim":
                    proximo = novo
            registrar(f"  linha {r['linha']}: digitados={digitados} manual={manual} tempo={time.time() - t0:.2f}s", mostrar=False)
            if tesseract and proximo and not manual and coluna_qtd:
                # A celula so e lida depois do Tab (valor formatado, ex.: 1.0000): fica na linha acima do novo foco.
                if proximo[1] - inicial[1] > 8:
                    passo = proximo[1] - inicial[1]
                topo = proximo[1] - passo
                celula = (coluna_qtd[0], topo, coluna_qtd[1], topo + coluna_qtd[2])
                esperada = float(r["quantidade"].replace(",", "."))
                recorte = recortar_celula(janela, celula)
                pendentes.append((r["linha"], esperada, pool.submit(conferir_recorte, recorte, tesseract, esperada), recorte))
            elif tesseract and manual:
                registrar(f"  linha {r['linha']}: conferencia da quantidade pulada (linha ajustada a mao)", mostrar=False)
            conta = "digitada " + r["conta"] if "conta" in digitados else "travada"
            custo = r["custo"] if "custo" in digitados else "travado"
            registrar(f"OK linha {r['linha']}: {r['item']} c/custo={custo} conta={conta} {r['quantidade']}"
                      + (" (ajustada a mao)" if manual else ""))
            marcar(r, "ok")
        falha = revisar(aguardar=True)
        while falha:
            if not tratar_divergencia(falha, no_fim=True):
                return 1
            falha = revisar(aguardar=True)
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
