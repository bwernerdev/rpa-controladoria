"""Logica do robo do SRP (via Citrix): leitura da planilha, validacao e envio de teclas.

Sequencia por linha: Item, Tab, C/Custo, Tab, [Conta, Tab se o campo estiver editavel], Quantidade, Tab.
A tela do Citrix e so imagem: o campo em foco e o bloco ciano; a Conta e editavel se o foco cai logo apos o C/Custo.
"""
import ctypes
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
import unicodedata
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

sys.coinit_flags = 2  # COM em modo STA, necessario para arrastar e soltar (tkdnd/OLE) e pywinauto coexistirem

from PIL import Image, ImageChops, ImageFilter, ImageGrab, ImageOps  # noqa: E402 (depois do coinit_flags)


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

    # read_only deixa o arquivo aberto ate o close(): sem ele a planilha ficava travada e o Excel nao conseguia salvar
    # por cima (06/11/2026: "salvar e fechar" no Iniciar travou o Excel numa janela de erro)
    livro = load_workbook(caminho, read_only=True, data_only=True)
    try:
        linhas = iter(list(livro.active.iter_rows(values_only=True)))
    finally:
        livro.close()
    cabecalho = next(linhas, None) or ()
    indices = {}
    for posicao, nome in enumerate(cabecalho):
        chave = COLUNAS.get(texto(nome).lower())
        if chave:
            indices[chave] = posicao
    faltando = [nome for nome, chave in COLUNAS.items() if chave not in indices]
    if faltando:
        raise ValueError(f"Colunas ausentes no cabeçalho: {', '.join(faltando)}")
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
            erros.append(f"linha {r['linha']}: quantidade inválida: {r['quantidade']!r}")
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
# cores com texto branco por cima: fecha os buracos do texto (dilata 2 px) antes da erosao. Com o Citrix em 1366x768
# os tracos de 1 px da data picotavam o azul da selecao e a erosao apagava tudo; o robo passou da Data Emissao
# (05/11/2026 14:03). Linhas de ate 4 px (bordas da janela no mesmo azul) continuam sumindo.
FECHAR = ("selecao",)


_TABELAS_COR = {}


def _mascara_crua(imagem, cor):
    """Mascara (L, 255 = pixel da cor) sem erosao. Uma tabela de 768 posicoes no `point` da imagem RGB inteira e o
    menor (`darker`) das 3 bandas: ~10 ms em 1920x1080 contra ~15 ms das 3 bandas separadas (mesmo resultado, 24/11/2026)."""
    tabela = _TABELAS_COR.get(cor)
    if tabela is None:
        tabela = _TABELAS_COR[cor] = [255 if teste(v) else 0 for teste in CORES[cor] for v in range(256)]
    r, g, b = (imagem if imagem.mode == "RGB" else imagem.convert("RGB")).point(tabela).split()
    return ImageChops.darker(ImageChops.darker(r, g), b)


def _mascara_cor(imagem, cor="ciano"):
    """(mascara erodida, deslocamento x, y) so da area onde a cor aparece: o filtro de erosao na tela inteira custava
    ~50 ms por consulta; sem a cor na tela (o vermelho do aviso, quase sempre) devolve (None, 0, 0) na hora."""
    mascara = _mascara_crua(imagem, cor)
    caixa = mascara.getbbox()
    if not caixa:
        return None, 0, 0
    folga = 5 if cor in FECHAR else 3
    esq, topo = max(caixa[0] - folga, 0), max(caixa[1] - folga, 0)
    recorte = mascara.crop((esq, topo, min(caixa[2] + folga, mascara.width), min(caixa[3] + folga, mascara.height)))
    if cor in FECHAR:  # fechamento (dilata 2 px) + a mesma erosao: o bloco sai do mesmo tamanho que sem texto
        return recorte.filter(ImageFilter.MaxFilter(5)).filter(ImageFilter.MinFilter(9)), esq, topo
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
    # A faixa vermelha do titulo do SRP faz a erosao cobrir quase a tela toda (~60 ms por consulta, 1920x1080). Se o
    # vermelho da tela e o mesmo da ultima consulta (mesmos pixels, mesmos fixos), a resposta tambem e: ~12 ms.
    crua = _mascara_crua(imagem, "vermelho")
    ultima = _ULTIMO_AVISO[0]
    if (ultima and ultima[0] == tuple(fixos) and ultima[1].size == crua.size
            and not ImageChops.difference(ultima[1], crua).getbbox()):
        return ultima[2]
    resultado = _icone_aviso(imagem, fixos)
    _ULTIMO_AVISO[0] = (tuple(fixos), crua, resultado)
    return resultado


_ULTIMO_AVISO = [None]  # (fixos, mascara crua do vermelho, resposta) da ultima consulta do icone_aviso


def _icone_aviso(imagem, fixos):
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
        return "já há um aviso aberto no SRP. Feche o aviso (OK) e deixe o cursor no campo Item"
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
        return " ".join(" ".join(t for _, t in sorted(palavras)) for _, palavras in sorted(linhas.items()))

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
        gravar_json(arquivo_aprendizado(), dados, indent=2)
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
    tolerancia = max(40, caixa[2] - caixa[0])
    nome, x = min(colunas.items(), key=lambda item: abs(item[1] - centro))
    if abs(x - centro) <= tolerancia:
        return nome
    # "Item" nem sempre sai no OCR (25/11/2026); e a 1a coluna da grade, entao foco bem a esquerda do C/Custo e o Item
    if "item" not in colunas and "custo" in colunas and centro < colunas["custo"] - tolerancia:
        return "item"
    return None


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
    raise LookupError(f"Janela com '{trecho}' no título não encontrada.")


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


def gravar_json(arquivo, dados, **opcoes):
    """Grava JSON por um arquivo temporario ao lado, trocado de uma vez (os.replace): o app fechado ou a energia caindo
    no meio da gravacao nao deixa o arquivo pela metade (progresso, aprendizado, posicoes e config). OSError sobe."""
    arquivo = Path(arquivo)
    temporario = arquivo.with_name(arquivo.name + ".tmp")
    temporario.write_text(json.dumps(dados, ensure_ascii=False, **opcoes), encoding="utf-8")
    os.replace(temporario, arquivo)


def arquivo_progresso():
    return pasta_base() / "progresso.json"


# ---------- linhas concluidas pintadas na planilha (usuario, 06/11/2026) ----------
VERDE_CONCLUIDA = "C6EFCE"  # o verde do estilo "Bom" do Excel
# nota lancada com "." na Serie (o SRP disse que o documento ja existia; usuario, 09/11/2026): verde mais escuro
# ("Verde, Enfase 6, 40% mais claro" do Excel; ainda conta como verde = feita ao recarregar) e comentario na CHAVE NF
VERDE_SERIE_PONTO = "A9D08E"


def cores_do_tema(xml):
    """Cores do tema da planilha na ordem dos indices de cor do Excel (0 lt1, 1 dk1, 2 lt2, 3 dk2, 4-9 accent1-6, 10
    hlink, 11 folHlink), em hex 'RRGGBB'; [] sem tema. As celulas pintadas com as cores do tema guardam so o indice."""
    if not xml:
        return []
    import xml.etree.ElementTree as ET
    try:
        raiz = ET.fromstring(xml)
    except ET.ParseError:
        return []
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    esquema = raiz.find(f".//{a}clrScheme")
    if esquema is None:
        return []
    cores = {}
    for filho in esquema:
        nome = filho.tag.replace(a, "")
        cor = next(iter(filho), None)
        if cor is not None:
            cores[nome] = (cor.get("val") if cor.tag == f"{a}srgbClr" else cor.get("lastClr")) or ""
    ordem = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
    return [cores.get(nome, "") for nome in ordem]


def fundo_verde(fill, tema=()):
    """Fundo solido verde (o VERDE_CONCLUIDA do robo ou outro verde pintado a mao, do verde-claro "80% mais claro" do
    Excel, E2EFDA, ao escuro): matiz entre 75 e 165 graus e alguma saturacao. Cor do tema: pelo indice em `tema`
    (cores_do_tema), com o clareamento/escurecimento (tint) aplicado."""
    if fill is None or fill.fill_type != "solid":
        return False
    cor = fill.fgColor
    if cor.type == "rgb" and isinstance(cor.rgb, str):
        hexa = cor.rgb[-6:]
    elif cor.type == "theme" and isinstance(cor.theme, int) and cor.theme < len(tema):
        hexa = tema[cor.theme]
    else:
        return False
    try:
        r, g, b = (int(hexa[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return False
    tint = cor.tint or 0
    r, g, b = ((c + (255 - c) * tint) if tint > 0 else c * (1 + tint) for c in (r, g, b))
    import colorsys
    matiz, saturacao, valor = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
    return 75 <= matiz * 360 <= 165 and saturacao >= 0.07 and valor >= 0.25


# aba com a copia das linhas das notas integradas (SNAC0236; usuario, 20/11/2026): as linhas continuam verdes na aba
# dos dados e a copia ganha a coluna LANCADO EM (data e hora da integracao)
ABA_LANCADOS = "Lançados"
TITULO_LANCADO_EM = "LANÇADO EM"
FORMATO_LANCADO_EM = "dd/mm/yyyy hh:mm"


def e_aba_lancados(nome):
    return str(nome).strip().casefold() == ABA_LANCADOS.casefold()


def aba_de_dados(livro):
    """A aba que o robo le e pinta: a ativa, menos a "Lançados" (se o usuario salvou a planilha com ela aberta, vale a
    primeira outra aba)."""
    ativa = livro.active
    if not e_aba_lancados(ativa.title):
        return ativa
    return next((aba for aba in livro.worksheets if not e_aba_lancados(aba.title)), ativa)


def _texto_chave(valor):
    return "" if valor is None else str(valor).strip()


def _chaves_por_linha(valores):
    """[(linha, chave)] de uma sequencia [(linha, valor da CHAVE)]: chave em branco = item da nota de cima."""
    saida, atual = [], ""
    for linha, valor in valores:
        atual = _texto_chave(valor) or atual
        saida.append((linha, atual))
    return saida


def pintar_linhas(caminho, linhas, cor=VERDE_CONCLUIDA, comentarios=None, lancados=None, avisos=None):
    """Fundo `cor` nas `linhas` da aba dos dados (`aba_de_dados`), nas colunas do cabecalho; `comentarios` {linha: texto}
    vao na celula da coluna CHAVE (sem ela, na coluna A), no lugar de um comentario que ja houver ali.
    `lancados` {linha: datetime} (SNAC0236): depois de pintar, copia essas linhas (valores, formato, cor e comentario)
    para o fim da aba "Lançados", com a data e hora na coluna LANCADO EM; cria a aba se faltar (com `lancados` vazio, so
    cria). Nota cuja CHAVE ja esta na "Lançados" nao e copiada de novo. Aba "Lançados" com outras colunas = nao copia e
    poe o motivo em `avisos` (lista), sem impedir a pintura.
    Planilha fechada: grava o arquivo (openpyxl; arquivo temporario e troca, para nao deixar a planilha pela metade).
    Planilha com formula: o openpyxl gravaria as formulas sem o valor calculado (o robo e qualquer leitura sem o Excel
    passariam a ver as celulas vazias; 24/11/2026). Entao abre num Excel proprio, escondido, faz o mesmo pelo COM, salva
    e fecha (`_pintar_com_excel`).
    Aberta no Excel (o Excel trava o arquivo): faz o mesmo na janela do Excel aberta, sem salvar (o usuario salva).
    Devolve 'arquivo', 'excel' ou None (nada a fazer); falha = excecao (quem chama avisa e segue: pintar nunca para o
    robo)."""
    from openpyxl import load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill

    caminho = Path(caminho)
    comentarios = comentarios or {}
    linhas = sorted(set(linhas))
    if not linhas:
        if lancados is None:
            return None
        leve = load_workbook(caminho, read_only=True)  # so os nomes das abas: nao regrava a planilha a toa
        try:
            if any(e_aba_lancados(nome) for nome in leve.sheetnames):
                return None
        finally:
            leve.close()
    livro = load_workbook(caminho, keep_vba=caminho.suffix.lower() == ".xlsm")
    aba = aba_de_dados(livro)
    colunas = max((c.column for c in aba[1] if c.value not in (None, "")), default=1)
    coluna_chave = next((c.column for c in aba[1] if str(c.value or "").strip().lower().startswith("chave")), 1)
    if _tem_formula(livro):
        return _pintar_com_excel(caminho, aba.title, linhas, colunas, cor, comentarios, coluna_chave, lancados, avisos)
    fundo = PatternFill("solid", fgColor=cor)
    for linha in linhas:
        for coluna in range(1, colunas + 1):
            aba.cell(row=linha, column=coluna).fill = fundo
    for linha, texto in comentarios.items():
        comentario = Comment(texto, "Robô SRP")
        comentario.width, comentario.height = 300, 80
        aba.cell(row=linha, column=coluna_chave).comment = comentario
    if lancados is not None:
        _copiar_lancados(caminho, livro, aba, colunas, coluna_chave, lancados, avisos)
    temporario = caminho.with_name(f"~rpa-{caminho.name}")
    try:
        livro.save(temporario)
        os.replace(temporario, caminho)
        return "arquivo"
    except PermissionError:
        pass  # planilha aberta no Excel
    finally:
        if temporario.exists():
            try:
                temporario.unlink()
            except OSError:
                pass
    _pintar_no_excel(caminho, aba.title, linhas, colunas, cor, comentarios, coluna_chave, lancados, avisos)
    return "excel"


def _copiar_lancados(caminho, livro, aba, colunas, coluna_chave, lancados, avisos):
    """Copia as linhas de `lancados` {linha: datetime} para a aba "Lançados" do `livro` (openpyxl, ainda nao salvo)."""
    from copy import copy

    from openpyxl.comments import Comment
    from openpyxl.utils import get_column_letter

    destino = next((a for a in livro.worksheets if e_aba_lancados(a.title)), None)
    titulos = [aba.cell(row=1, column=c).value for c in range(1, colunas + 1)]
    if destino is None:
        destino = livro.create_sheet(ABA_LANCADOS)  # no fim; a aba ativa continua a dos dados
        for c in range(1, colunas + 2):
            origem = aba.cell(row=1, column=min(c, colunas))
            celula = destino.cell(row=1, column=c, value=titulos[c - 1] if c <= colunas else TITULO_LANCADO_EM)
            celula._style = copy(origem._style)
            largura = aba.column_dimensions[get_column_letter(c)].width if c <= colunas else 18
            if largura:
                destino.column_dimensions[get_column_letter(c)].width = largura
        destino.freeze_panes = "A2"
    elif [destino.cell(row=1, column=c).value for c in range(1, colunas + 1)] != titulos:
        if avisos is not None:
            avisos.append(f"a aba {ABA_LANCADOS} tem colunas diferentes da planilha: não copiei as notas para lá")
        return
    if not lancados:
        return
    # chave de cada linha (em branco = item da nota de cima): nota ja copiada antes nao entra de novo
    ja = {chave for _, chave in _chaves_por_linha(
        (n, destino.cell(row=n, column=coluna_chave).value) for n in range(2, destino.max_row + 1)) if chave}
    pedidas = sorted(lancados)
    chave_de = dict(_chaves_por_linha((n, aba.cell(row=n, column=coluna_chave).value) for n in range(2, max(pedidas) + 1)))
    copiar = [n for n in pedidas if chave_de.get(n) and chave_de[n] not in ja]
    if not copiar:
        return
    calculados = _valores_calculados(caminho, aba.title, copiar, colunas)
    proxima = destino.max_row + 1
    while proxima > 2 and all(destino.cell(row=proxima - 1, column=c).value in (None, "") for c in range(1, colunas + 2)):
        proxima -= 1  # linhas vazias com formato no fim da aba nao contam
    for n in copiar:
        for c in range(1, colunas + 1):
            origem = aba.cell(row=n, column=c)
            valor = calculados.get((n, c), origem.value) if origem.data_type == "f" else origem.value
            celula = destino.cell(row=proxima, column=c, value=valor)
            celula._style = copy(origem._style)
            if origem.comment is not None:
                comentario = Comment(origem.comment.text, origem.comment.author)
                comentario.width, comentario.height = origem.comment.width, origem.comment.height
                celula.comment = comentario
        celula = destino.cell(row=proxima, column=colunas + 1, value=lancados[n])
        celula.number_format = FORMATO_LANCADO_EM
        proxima += 1


def _tem_formula(livro):
    """Alguma celula com formula em alguma aba (pasta lida pelo openpyxl sem `data_only`)?"""
    return any(celula.data_type == "f" for aba in livro.worksheets for linha in aba.iter_rows() for celula in linha)


def formulas_sem_valor(caminho, nome_aba, valores):
    """Celulas ("A2"...) da aba com formula e sem o valor calculado guardado no arquivo (`valores`: as linhas lidas com
    `data_only`). Acontece quando a planilha e gravada por um programa que nao calcula (o robo ate a v1.16.2, ao pintar e
    criar a aba Lançados com o openpyxl; 24/11/2026): o Excel recalcula ao abrir, mas o arquivo fica sem os valores."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    livro = load_workbook(caminho, read_only=True)
    try:
        faltam = []
        for n, linha in enumerate(livro[nome_aba].iter_rows(max_row=len(valores), values_only=True), start=1):
            lidos = valores[n - 1]
            faltam += [f"{get_column_letter(c + 1)}{n}" for c, formula in enumerate(linha)
                       if isinstance(formula, str) and formula.startswith("=") and (c >= len(lidos) or lidos[c] is None)]
        return faltam
    finally:
        livro.close()


def _valores_calculados(caminho, nome_aba, linhas, colunas):
    """{(linha, coluna): valor} do ultimo calculo do Excel: a copia leva o valor, nao a formula (a formula apontaria
    para outras linhas na aba Lançados)."""
    from openpyxl import load_workbook

    livro = load_workbook(caminho, read_only=True, data_only=True)
    try:
        aba = livro[nome_aba]
        querer = set(linhas)
        saida = {}
        for n, linha in enumerate(aba.iter_rows(min_row=min(querer), max_row=max(querer), max_col=colunas,
                                                values_only=True), start=min(querer)):
            if n in querer:
                saida.update({(n, c): valor for c, valor in enumerate(linha, start=1)})
        return saida
    finally:
        livro.close()


def _pintar_no_excel(caminho, nome_aba, linhas, colunas, cor, comentarios=None, coluna_chave=1, lancados=None,
                     avisos=None):
    """Pinta (e copia para a aba Lançados) pela planilha aberta no Excel (COM). So se ela estiver aberta: nunca abre o
    arquivo nem salva."""
    import pythoncom

    pythoncom.CoInitialize()
    try:
        livro = _pasta_aberta(caminho)
        if livro is None:
            raise RuntimeError("a planilha está travada por outro programa e não achei ela aberta no Excel")
        _pintar_livro_excel(livro, nome_aba, linhas, colunas, cor, comentarios, coluna_chave, lancados, avisos)
    finally:
        pythoncom.CoUninitialize()


def _pintar_livro_excel(livro, nome_aba, linhas, colunas, cor, comentarios, coluna_chave, lancados, avisos):
    """Pintura, comentarios e copia para a aba Lançados numa pasta do Excel (objeto COM), sem salvar."""
    from openpyxl.utils import get_column_letter

    r, g, b = (int(cor[i:i + 2], 16) for i in (0, 2, 4))
    aba = livro.Worksheets(nome_aba)
    ultima = get_column_letter(colunas)
    for linha in linhas:
        aba.Range(f"A{linha}:{ultima}{linha}").Interior.Color = r + g * 256 + b * 65536
    for linha, texto in (comentarios or {}).items():
        celula = aba.Cells(linha, coluna_chave)
        celula.ClearComments()
        celula.AddComment(texto)
    if lancados is not None:
        _copiar_lancados_excel(livro, aba, colunas, coluna_chave, lancados, avisos)


def _pintar_com_excel(caminho, nome_aba, linhas, colunas, cor, comentarios=None, coluna_chave=1, lancados=None,
                      avisos=None):
    """Planilha com formula: pinta pelo Excel, que grava as formulas com o valor calculado. Aberta no Excel do usuario =
    na janela dele, sem salvar ('excel'); fechada = num Excel proprio e escondido (`DispatchEx`: nao mexe nas janelas
    do usuario), que abre, faz, salva, fecha e sai ('arquivo')."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        livro = _pasta_aberta(caminho)
        if livro is not None:
            _pintar_livro_excel(livro, nome_aba, linhas, colunas, cor, comentarios, coluna_chave, lancados, avisos)
            return "excel"
        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except pythoncom.com_error as erro:
            raise RuntimeError("a planilha tem fórmulas e não consegui abrir o Excel para gravar com os valores "
                               f"calculados: {erro.strerror or erro}") from None
        try:
            app.Visible, app.DisplayAlerts, app.ScreenUpdating = False, False, False
            livro = app.Workbooks.Open(str(Path(caminho).resolve()), UpdateLinks=0)
            try:
                if livro.ReadOnly:  # travada por outro programa: salvar abriria o "Salvar como"
                    raise RuntimeError("a planilha está travada por outro programa (o Excel abriu como somente leitura)")
                _pintar_livro_excel(livro, nome_aba, linhas, colunas, cor, comentarios, coluna_chave, lancados, avisos)
                livro.Save()
            finally:
                livro.Close(SaveChanges=False)
        finally:
            app.Quit()
        return "arquivo"
    finally:
        pythoncom.CoUninitialize()


def _coluna_excel(aba, coluna, de, ate):
    """Valores de uma coluna (linhas `de` a `ate`) pelo COM; uma celula so vem como valor solto, nao como tupla."""
    if ate < de:
        return []
    valores = aba.Range(aba.Cells(de, coluna), aba.Cells(ate, coluna)).Value
    return [linha[0] for linha in valores] if isinstance(valores, tuple) else [valores]


def _copiar_lancados_excel(livro, aba, colunas, coluna_chave, lancados, avisos):
    """O mesmo que `_copiar_lancados`, na pasta aberta no Excel (COM, sem salvar)."""
    from openpyxl.utils import get_column_letter

    ultima = get_column_letter(colunas)
    data = get_column_letter(colunas + 1)
    destino = next((livro.Worksheets(i) for i in range(1, livro.Worksheets.Count + 1)
                    if e_aba_lancados(livro.Worksheets(i).Name)), None)
    titulos = list(aba.Range(f"A1:{ultima}1").Value[0])
    if destino is None:
        ativa = livro.ActiveSheet
        destino = livro.Worksheets.Add(None, livro.Worksheets(livro.Worksheets.Count))  # (Before, After): no fim
        destino.Name = ABA_LANCADOS
        aba.Range(f"A1:{ultima}1").Copy(destino.Range("A1"))
        aba.Range(f"{ultima}1").Copy(destino.Range(f"{data}1"))
        destino.Range(f"{data}1").Value = TITULO_LANCADO_EM
        for c in range(1, colunas + 1):
            destino.Columns(c).ColumnWidth = aba.Columns(c).ColumnWidth
        destino.Columns(colunas + 1).ColumnWidth = 18
        ativa.Activate()  # o Add deixa a aba nova ativa: volta para a que o usuario via
    elif list(destino.Range(f"A1:{ultima}1").Value[0]) != titulos:
        if avisos is not None:
            avisos.append(f"a aba {ABA_LANCADOS} tem colunas diferentes da planilha: não copiei as notas para lá")
        return
    if not lancados:
        return
    usada = destino.UsedRange
    fim = usada.Row + usada.Rows.Count - 1
    while fim > 1 and all(v in (None, "") for v in destino.Range(f"A{fim}:{data}{fim}").Value[0]):
        fim -= 1
    ja = set()
    if fim >= 2:
        valores = _coluna_excel(destino, coluna_chave, 2, fim)
        ja = {chave for _, chave in _chaves_por_linha(enumerate(valores, start=2)) if chave}
    pedidas = sorted(lancados)
    chave_de = dict(_chaves_por_linha(enumerate(_coluna_excel(aba, coluna_chave, 2, max(pedidas)), start=2)))
    proxima = fim + 1
    formato = _formato_data_hora_excel(livro.Application)
    for n in pedidas:
        if not chave_de.get(n) or chave_de[n] in ja:
            continue
        de, para = aba.Range(f"A{n}:{ultima}{n}"), destino.Range(f"A{proxima}:{ultima}{proxima}")
        de.Copy(para)  # valores, formato, cor e comentario (texto continua texto: a CHAVE nao vira numero)
        for c in range(1, colunas + 1):
            if de.Cells(1, c).HasFormula:  # a formula apontaria para outras linhas: fica o valor calculado
                valor = de.Cells(1, c).Value
                para.Cells(1, c).Value = f"'{valor}" if isinstance(valor, str) else valor
        celula = destino.Range(f"{data}{proxima}")
        # numero de serie do Excel (dias desde 30/12/1899): o datetime pelo COM chegava 3 h adiantado (fuso)
        celula.Value = (lancados[n] - datetime(1899, 12, 30)).total_seconds() / 86400
        celula.NumberFormatLocal = formato
        proxima += 1


def _formato_data_hora_excel(app):
    """dd/mm/aaaa hh:mm com as letras do idioma do Excel (no Excel em portugues o ano e 'a': com 'yyyy' a celula
    mostrava '20/11/yyyy'). Codigos do Application.International: 19 ano, 20 mes, 21 dia, 22 hora, 23 minuto."""
    try:
        codigos = app.International
        ano, mes, dia, hora, minuto = (codigos[i - 1] for i in (19, 20, 21, 22, 23))
    except Exception:
        ano, mes, dia, hora, minuto = "y", "m", "d", "h", "m"
    return f"{dia * 2}/{mes * 2}/{ano * 4} {hora * 2}:{minuto * 2}"


def _pasta_aberta(caminho):
    """A pasta do Excel aberta com este arquivo (objeto COM) ou None. O Excel registra cada pasta aberta na tabela de
    objetos em execucao (ROT) pelo caminho: acha a planilha em qualquer janela do Excel, sem abrir nada
    (GetObject(caminho) abriria o arquivo se ele nao estivesse aberto). Chamar com o COM ja iniciado."""
    import pythoncom
    import win32com.client

    alvo = os.path.normcase(os.path.abspath(caminho))
    tabela, contexto = pythoncom.GetRunningObjectTable(), pythoncom.CreateBindCtx(0)
    for moniker in tabela:
        try:
            if os.path.normcase(moniker.GetDisplayName(contexto, None)) == alvo:
                return win32com.client.Dispatch(tabela.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
        except pythoncom.com_error:
            continue
    return None


def fechar_no_excel(caminho, tentativas=5):
    """Salva e fecha no Excel SO a pasta deste arquivo (a planilha carregada no robo; usuario, 06/11/2026): as outras
    pastas e o proprio Excel ficam abertos. Devolve True se estava aberta e foi salva e fechada, False se nao estava
    aberta. Falha (ex.: celula em edicao no Excel, que recusa as chamadas) = RuntimeError com o motivo."""
    import pythoncom

    pythoncom.CoInitialize()
    try:
        livro = _pasta_aberta(caminho)
        if livro is None:
            return False
        for tentativa in range(tentativas):
            try:
                salvar = not livro.Saved
                if salvar and livro.ReadOnly:  # salvar abriria o "Salvar como" do Excel e travaria aqui
                    raise RuntimeError("a planilha está aberta no Excel como somente leitura e tem alterações; "
                                       "feche-a no Excel")
                livro.Close(SaveChanges=salvar)
                return True
            except pythoncom.com_error as erro:
                if tentativa == tentativas - 1:
                    raise RuntimeError("o Excel não aceitou salvar e fechar a planilha (alguma célula em edição ou "
                                       f"janela aberta no Excel?): {erro.strerror or erro}") from None
                time.sleep(0.5)
    finally:
        pythoncom.CoUninitialize()


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
            gravar_json(arquivo_progresso(), dados, indent=2)
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
    return {n for n in dados.get("ok", []) if n in validas}, dados.get("em")


def sem_acentos(texto):
    """Texto sem acentos ('Não iniciei' -> 'Nao iniciei'). As mensagens do robo tem acento desde a v1.14.0 (antes eram
    escritas sem e o `com_acentos` os punha so na tela); quem le mensagens ou logs compara pelo texto sem acento, que
    vale para os logs antigos e os novos."""
    return "".join(c for c in unicodedata.normalize("NFD", texto) if not unicodedata.combining(c))


def qtd(n, um, varios):
    """Quantidade com o plural certo para os textos da tela: qtd(1, "nota", "notas") -> '1 nota'."""
    return f"{n} {um if n == 1 else varios}"


REPOSITORIO = "bwernerdev/rpa-controladoria"  # publico: os releases sao lidos sem token


def versao_tupla(texto):
    """'v1.10.2' -> (1, 10, 2); texto invalido -> (0,)."""
    try:
        return tuple(int(p) for p in texto.strip().lstrip("vV").split("."))
    except ValueError:
        return (0,)


def _powershell_web(script, tempo_limite):
    """Roda um trecho do PowerShell que usa o proxy do Windows (PAC/autodetect da rede da empresa, com a credencial do
    usuario logado). O urllib do Python so enxerga proxy fixo: numa rede com PAC ele tenta ir direto e e barrado."""
    prefixo = ("$ProgressPreference='SilentlyContinue'; $ErrorActionPreference='Stop'; "
               "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; "
               "$p=[Net.WebRequest]::DefaultWebProxy; if ($p) { $p.Credentials=[Net.CredentialCache]::DefaultCredentials }; "
               "[Console]::OutputEncoding=[Text.Encoding]::UTF8; ")
    resultado = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", prefixo + script],
                               capture_output=True, encoding="utf-8", errors="replace", timeout=tempo_limite,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if resultado.returncode != 0:
        erro = " ".join((resultado.stderr or "").split())
        raise OSError(erro[:300] or f"PowerShell saiu com {resultado.returncode}")
    return resultado.stdout


def _texto_ps(texto):
    return "'" + str(texto).replace("'", "''") + "'"


def _consulta_powershell(url, tempo_limite=40):
    """GET de uma URL da API do GitHub pelo proxy do Windows; devolve o JSON."""
    saida = _powershell_web(f"(Invoke-WebRequest -UseBasicParsing -Uri {_texto_ps(url)} -UserAgent 'RPA-SRP' "
                            "-Headers @{Accept='application/vnd.github+json'} -TimeoutSec 30).Content", tempo_limite)
    return json.loads(saida)


def _erro_curto(erro):
    texto = " ".join(str(erro).split()) or type(erro).__name__
    return texto[:200]


def verificar_atualizacao(atual, tempo_limite=5):
    """Consulta o release mais novo do GitHub. Devolve (info, texto): `info` como em `buscar_atualizacao` (ou None) e
    `texto` = o que aconteceu, para o console e o logs\atualizacao.log. Tenta direto (urllib) e, se falhar, pelo proxy do
    Windows (PowerShell). Nunca levanta erro: a checagem nao pode atrapalhar o app."""
    import urllib.request
    url = f"https://api.github.com/repos/{REPOSITORIO}/releases/latest"
    caminho = "direta"
    try:
        pedido = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json", "User-Agent": "RPA-SRP"})
        with urllib.request.urlopen(pedido, timeout=tempo_limite) as resposta:
            dados = json.load(resposta)
    except Exception as erro_direto:
        caminho = "pelo proxy do Windows"
        try:
            dados = _consulta_powershell(url)
        except Exception as erro_proxy:
            return None, (f"Não consegui consultar atualizações (direto: {_erro_curto(erro_direto)}; "
                          f"pelo proxy do Windows: {_erro_curto(erro_proxy)})")
    try:
        versao = str(dados.get("tag_name", "")).lstrip("vV")
        if versao_tupla(versao) <= versao_tupla(atual):
            return None, f"Atualizações: a v{atual} é a mais nova (última publicada: v{versao}; consulta {caminho})"
        zips = [a for a in dados.get("assets", []) if str(a.get("name", "")).lower().endswith(".zip")]
        if not zips:
            return None, f"Atualizações: v{versao} publicada sem o zip; nada a fazer (consulta {caminho})"
        digest = str(zips[0].get("digest") or "")
        info = {"versao": versao, "url": zips[0]["browser_download_url"], "tamanho": zips[0].get("size"),
                "sha256": digest.split(":", 1)[1] if digest.startswith("sha256:") else None,
                "notas": dados.get("body") or "", "proxy": caminho != "direta"}
    except Exception as erro:
        return None, f"Resposta do GitHub inesperada: {_erro_curto(erro)} (consulta {caminho})"
    return info, f"Atualizações: v{versao} disponível (você tem a v{atual}; consulta {caminho})"


def buscar_atualizacao(atual, tempo_limite=5):
    """Release mais nova que `atual` no GitHub: {'versao', 'url', 'sha256', 'tamanho', 'notas', 'proxy'}; None se nao
    houver, se nao tiver o zip ou sem internet (nunca levanta erro: a checagem nao pode atrapalhar o app)."""
    return verificar_atualizacao(atual, tempo_limite)[0]


def registrar_atualizacao(texto):
    """Acrescenta uma linha com data e hora no logs\atualizacao.log (o mesmo do instalar.ps1). Falha = ignora."""
    try:
        arquivo = pasta_logs() / "atualizacao.log"
        arquivo.parent.mkdir(parents=True, exist_ok=True)
        with open(arquivo, "a", encoding="utf-8") as saida:
            saida.write(f"{time.strftime('%d/%m/%Y %H:%M:%S')} {texto}\n")
    except OSError:
        pass


def registrar_falha_tela(contexto):
    """Guarda o traceback da excecao em curso em logs\\app-erros.log (com data, hora e `contexto`). O exe roda sem
    console: sem isso, um erro dentro de um callback da janela sumia sem rastro. Falha ao gravar = ignora."""
    import traceback
    try:
        arquivo = pasta_logs() / "app-erros.log"
        arquivo.parent.mkdir(parents=True, exist_ok=True)
        with open(arquivo, "a", encoding="utf-8") as saida:
            saida.write(f"{time.strftime('%d/%m/%Y %H:%M:%S')} {contexto}\n{traceback.format_exc()}\n")
    except OSError:
        pass


def _baixar_powershell(url, arquivo, total, progresso, tempo_limite=900):
    """Baixa `url` para `arquivo` pelo proxy do Windows, informando o tamanho do arquivo enquanto cresce."""
    import threading
    feito = threading.Event()
    falha = []

    def rodar():
        try:
            _powershell_web(f"Invoke-WebRequest -UseBasicParsing -Uri {_texto_ps(url)} -UserAgent 'RPA-SRP' "
                            f"-OutFile {_texto_ps(arquivo)} -TimeoutSec 600", tempo_limite)
        except Exception as erro:
            falha.append(erro)
        finally:
            feito.set()
    threading.Thread(target=rodar, daemon=True).start()
    while not feito.wait(0.5):
        if progresso and arquivo.exists():
            progresso(arquivo.stat().st_size, total)
    if falha:
        raise OSError(f"download pelo proxy do Windows falhou: {_erro_curto(falha[0])}")
    if progresso and arquivo.exists():
        progresso(arquivo.stat().st_size, total)


def baixar_atualizacao(info, progresso=None):
    """Baixa e extrai o zip do release numa pasta temporaria; confere o SHA-256 publicado pelo GitHub.
    `progresso(baixados, total)` e opcional. Devolve a pasta com instalar.ps1 e RPA-SRP.exe."""
    import urllib.request
    import zipfile
    pasta = Path(tempfile.mkdtemp(prefix="rpa-srp-"))
    arquivo = pasta / "pacote.zip"
    soma = hashlib.sha256()
    pedido = urllib.request.Request(info["url"], headers={"User-Agent": "RPA-SRP"})
    try:
        if info.get("proxy"):  # a consulta so passou pelo proxy do Windows: o download direto tambem seria barrado
            raise OSError("consulta feita pelo proxy do Windows")
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
    except OSError:
        arquivo.unlink(missing_ok=True)
        _baixar_powershell(info["url"], arquivo, info.get("tamanho") or 0, progresso)
        soma = hashlib.sha256(arquivo.read_bytes())
    if info.get("sha256") and soma.hexdigest() != info["sha256"].lower():
        raise ValueError("o arquivo baixado não confere com o publicado (SHA-256 diferente); atualização cancelada")
    destino = pasta / "pacote"
    with zipfile.ZipFile(arquivo) as zip_:
        zip_.extractall(destino)  # extractall descarta caminhos absolutos e ".."
    if not (destino / "instalar.ps1").exists() or not (destino / "RPA-SRP.exe").exists():
        raise ValueError("o pacote baixado não tem instalar.ps1 e RPA-SRP.exe")
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


COMPACTAR_DEPOIS = 2  # dias: logs de execucao mais velhos que isso vao para logs/arquivo-AAAAMM.zip
DIAS_TELAS = 7  # prints de logs/telas (cursor parado): so servem para analisar uma pausa recente


def _log_de_execucao(arquivo):
    """Log de uma execucao do robo ('AAAAMMDD-HHMMSS[...].log'), nao o atualizacao.log."""
    return arquivo.stem[:8].isdigit() and arquivo.stem[8:9] == "-"


def compactar_logs(dias=COMPACTAR_DEPOIS):
    """Move os logs de execucao com mais de `dias` dias para um zip por mes (`logs/arquivo-AAAAMM.zip`, pelo nome do log;
    texto de log compacta ~90%). Devolve quantos moveu. Falha num arquivo = ele fica como esta (tenta na proxima)."""
    import zipfile
    limite = time.time() - dias * 86400
    movidos = 0
    for arquivo in sorted(pasta_logs().glob("*.log")):
        try:
            if not _log_de_execucao(arquivo) or arquivo.stat().st_mtime >= limite:
                continue
            with zipfile.ZipFile(pasta_logs() / f"arquivo-{arquivo.stem[:6]}.zip", "a", zipfile.ZIP_DEFLATED) as pacote:
                if arquivo.name not in pacote.namelist():
                    pacote.write(arquivo, arquivo.name)
            arquivo.unlink()
            movidos += 1
        except (OSError, zipfile.BadZipFile):
            pass
    return movidos


def ler_logs(padrao="*.log", pasta=None):
    """[(nome, texto)] dos logs de execucao que casam com `padrao`, soltos e dentro dos zips mensais, em ordem de nome
    (= ordem de hora)."""
    import fnmatch
    import zipfile
    pasta = pasta or pasta_logs()
    lidos = {}
    for pacote in sorted(pasta.glob("arquivo-*.zip")):
        try:
            with zipfile.ZipFile(pacote) as zipado:
                for nome in zipado.namelist():
                    if fnmatch.fnmatch(nome, padrao):
                        lidos[nome] = zipado.read(nome).decode("utf-8", errors="replace")
        except (OSError, zipfile.BadZipFile):
            continue
    for arquivo in pasta.glob(padrao):
        try:
            lidos[arquivo.name] = arquivo.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
    return sorted(lidos.items())


def limpar_logs(dias=30, dias_telas=DIAS_TELAS):
    """Compacta os logs de execucao com mais de 2 dias (`compactar_logs`) e apaga o que passou do prazo: logs soltos,
    zips mensais (30 dias depois do ultimo log que entrou neles) e recortes de OCR (logs/ocr) com mais de `dias` dias;
    prints de logs/telas com mais de `dias_telas` dias (usuario, 09/11/2026: logs ocupando pouco espaco). Devolve
    quantos apagou."""
    compactar_logs()
    limite, limite_telas = time.time() - dias * 86400, time.time() - dias_telas * 86400
    apagados = 0
    pasta = pasta_logs()
    for arquivo, prazo in [*((a, limite) for a in pasta.glob("*.log")), *((a, limite) for a in pasta.glob("arquivo-*.zip")),
                           *((a, limite) for a in pasta.glob("ocr/*.png")),
                           *((a, limite_telas) for a in pasta.glob("telas/*.png"))]:
        try:
            if arquivo.stat().st_mtime < prazo:
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


class Execucao0234:
    """Uma execucao do `executar` (v1.14.0: antes uma funcao de 307 linhas com 8 funcoes internas).
    Os parametros e o estado que as etapas dividem ficam em self; cada etapa e um metodo; `rodar` e o laco."""

    def __init__(self, registros, janela_trecho, pausa, render, espera, parar, aviso, estado=None, conferir=False, retomar=None,
             pausar=None):
        self.registros = registros
        self.janela_trecho = janela_trecho
        self.pausa = pausa
        self.render = render
        self.espera = espera
        self.parar = parar
        self.aviso = aviso
        self.estado = estado
        self.conferir = conferir
        self.retomar = retomar
        self.pausar = pausar

    def marcar(self, r, nome):
        self.atual = r["linha"] if nome == "executando" else None
        if self.estado:
            self.estado(r["linha"], nome)

    def registrar(self, texto, mostrar=True):
        with self.trava_log:
            if self.log.closed:
                return
            self.log.write(texto + "\n")
            self.log.flush()
        if mostrar:
            self.aviso(texto)

    def revisar(self, aguardar=False):
        """Trata os resultados de OCR ja prontos (todos, com aguardar=True). Devolve a linha que nao conferiu, se houver."""
        restantes = []
        falha = None
        for linha, esperada, futuro, recorte in self.pendentes:
            if falha is None and (aguardar or futuro.done()):
                confere, lidas = futuro.result()
                self.registrar(f"  linha {linha}: quantidade lidas={lidas} esperada={esperada}", mostrar=False)
                if not confere:
                    pasta_ocr = pasta_logs() / "ocr"
                    pasta_ocr.mkdir(exist_ok=True)
                    recorte.save(pasta_ocr / f"{datetime.now():%H%M%S}-linha{linha}.png")
                    falha = (linha, lidas, esperada)
            else:
                restantes.append((linha, esperada, futuro, recorte))
        self.pendentes[:] = restantes
        return falha

    def teclar(self, acao, valor, esperar=True):
        """`esperar=False` em Tab seguido de aguardar_mudanca (a propria espera do foco sincroniza)."""
        if self.parar.is_set() or mouse_no_canto():
            raise KeyboardInterrupt
        # 0,03 s por tecla (era 0,05): o Citrix recebe as teclas em ordem; a sincronizacao real e o foco mudar apos o Tab
        send_keys("{TAB}" if acao == "tab" else valor.replace(",", "{,}"), pause=0.03)
        if acao == "digitar":
            time.sleep(min(self.pausa, 0.05))
        elif esperar:
            time.sleep(self.pausa)

    def esperar_continuar(self):
        self.retomar.clear()
        while not self.retomar.is_set():
            if self.parar.is_set():
                raise KeyboardInterrupt
            time.sleep(0.2)

    def tratar_divergencia(self, falha, no_fim=False):
        """Quantidade lida diferente da planilha: pausa (se houver Continuar) ou interrompe. True = seguir."""
        linha, lidas, esperada = falha
        if self.retomar is None:
            self.registrar(f"Linha {linha}: quantidade na tela {lidas} difere da planilha ({esperada}). Interrompido. "
                      "A linha já foi confirmada no SRP: corrija-a (as seguintes, até aqui, também podem estar afetadas). "
                      "Imagem lida em logs/ocr.")
            if self.estado:
                self.estado(linha, "erro")
            return False
        self.registrar(f"Linha {linha}: a quantidade lida na tela {lidas} difere da planilha ({esperada}). PAUSADO. "
                  f"Confira a linha {linha} no SRP (imagem lida em logs/ocr). Se estiver certa, clique em Continuar. "
                  "Se não, corrija-a" + ("." if no_fim else ", volte o cursor para o Item da próxima linha livre") +
                  " e clique em Continuar.")
        if self.estado:
            self.estado(linha, "pausado")
        while True:
            self.esperar_continuar()
            if no_fim:
                break
            self.janela.set_focus()
            self.aviso("Continuando em 2 s...")
            for _ in range(20):
                if self.parar.is_set():
                    raise KeyboardInterrupt
                time.sleep(0.1)
            caixa = caixa_foco(self.janela)
            if caixa and (not self.colunas or qual_coluna(caixa, self.colunas) == "item"):
                break
            self.registrar("Ainda pausado: deixe o cursor no Item da próxima linha livre e clique em Continuar.")
            if self.estado:
                self.estado(linha, "pausado")
        self.registrar(f"  linha {linha}: quantidade conferida pelo usuário após a pausa", mostrar=False)
        if self.estado:
            self.estado(linha, "ok")
        return True

    def pausa_pedida(self, r):
        """F9: pausa antes de comecar a linha `r`; volta quando o cursor estiver no Item e o usuario continuar."""
        self.pausar.clear()
        self.registrar(f"Pausado antes da linha {r['linha']} (F9). Para continuar: deixe o cursor no Item da linha "
                  f"{r['linha']} no SRP e clique em Continuar (ou F9).")
        if self.estado:
            self.estado(r["linha"], "pausado")
        while True:
            self.esperar_continuar()
            self.pausar.clear()
            self.janela.set_focus()
            self.aviso("Continuando em 2 s...")
            for _ in range(20):
                if self.parar.is_set():
                    raise KeyboardInterrupt
                time.sleep(0.1)
            caixa = caixa_foco(self.janela)
            if caixa and (not self.colunas or qual_coluna(caixa, self.colunas) == "item"):
                return caixa
            self.registrar(f"Ainda pausado: deixe o cursor no Item da linha {r['linha']} e clique em Continuar (ou F9).")
            if self.estado:
                self.estado(r["linha"], "pausado")

    def esperar_correcao(self, r, campo_erro):
        """Pausa ate o usuario corrigir no SRP e clicar em Continuar. Devolve (caixa do foco, campo onde retomar);
        'fim' quando o cursor ja esta no Item (linha completada a mao)."""
        if self.exe_tesseract:
            tela_do_aviso = capturar(self.janela)  # captura ja: o usuario pode fechar o aviso a qualquer momento

            def ler_em_segundo_plano():
                try:
                    info = ler_aviso(tela_do_aviso, self.exe_tesseract, self.fixos)
                except Exception as erro:
                    self.registrar(f"  leitura do aviso falhou: {erro}", mostrar=False)
                    return
                if not info:
                    return
                self.registrar(f'Linha {r["linha"]}: SRP: "{info["texto"]}"')
                if info["conta_correta"] and info["item"] == r["item"] and (info["custo"] in (None, r["custo"])):
                    aprender_conta(r["item"], r["custo"], info["conta_correta"])
                    self.registrar(f"Linha {r['linha']}: o SRP informou a Conta correta {info['conta_correta']} para o item "
                              f"{r['item']} (C/Custo {r['custo']}); guardada para avisar nas próximas planilhas.")
            threading.Thread(target=ler_em_segundo_plano, daemon=True).start()
        self.registrar(f"Linha {r['linha']}: o SRP não avançou depois do {NOME_CAMPO[campo_erro]}. PAUSADO. "
                  f"No SRP: feche o aviso, corrija o {NOME_CAMPO[campo_erro]}, aperte Tab para ir ao campo seguinte "
                  "e clique em Continuar. Se completou a linha a mão, deixe o cursor no Item da linha seguinte.")
        if self.estado:
            self.estado(r["linha"], "pausado")
        while True:
            self.esperar_continuar()
            self.janela.set_focus()
            self.aviso("Continuando em 2 s...")
            for _ in range(20):
                if self.parar.is_set():
                    raise KeyboardInterrupt
                time.sleep(0.1)
            caixa = caixa_foco(self.janela)
            onde = qual_coluna(caixa, self.colunas) if caixa else None
            if onde == "item":
                self.registrar(f"  linha {r['linha']}: retomada com o cursor no Item (linha completada a mão)", mostrar=False)
                if self.estado:
                    self.estado(r["linha"], "executando")
                return caixa, "fim"
            if onde in ORDEM and ORDEM.index(onde) > ORDEM.index(campo_erro):
                self.registrar(f"  linha {r['linha']}: retomada no campo {NOME_CAMPO[onde]}", mostrar=False)
                if self.estado:
                    self.estado(r["linha"], "executando")
                return caixa, onde
            self.registrar(f"Ainda pausado: o cursor está em '{NOME_CAMPO.get(onde, 'campo não identificado')}'. "
                      f"Deixe-o no campo depois do {NOME_CAMPO[campo_erro]} (ou no Item da linha seguinte) e clique em Continuar.")
            if self.estado:
                self.estado(r["linha"], "pausado")

    def rodar(self):
        tesseract = caminho_tesseract() if self.conferir else None
        limpar_logs()
        # aberto a execucao inteira (varios metodos escrevem); fechado no fim do `rodar`
        self.log = open(pasta_logs() / f"{datetime.now():%Y%m%d-%H%M%S}.log", "w", encoding="utf-8")  # noqa: SIM115
        self.atual = None

        self.trava_log = threading.Lock()

        pool = ThreadPoolExecutor(max_workers=1)
        self.pendentes = []  # (linha, esperada, futuro, recorte): conferencias de OCR em andamento, em ordem

        try:
            if self.conferir and not tesseract:
                self.aviso("ATENÇÃO: pasta 'tesseract' não encontrada ao lado do programa. Conferência da quantidade DESATIVADA nesta execução.")
            self.janela = janela_citrix(self.janela_trecho)
            self.janela.set_focus()
            self.aviso(f"Clique no campo Item da primeira linha. Iniciando em {self.espera}s (mouse no canto superior esquerdo aborta).")
            inicio_contagem = time.time()
            self.colunas = None
            self.exe_tesseract = caminho_tesseract()
            tela_inicial = capturar(self.janela)
            ret = self.janela.rectangle()
            problema = problema_tela((ret.left, ret.top, ret.right, ret.bottom), tela_inicial)
            if problema:
                self.registrar(f"Não iniciei: {problema}. Nada foi digitado.")
                return 1
            try:  # vermelhos que ja estao na tela (faixa do cabecalho do SRP): nao sao aviso
                self.fixos = caixas_cor(tela_inicial, "vermelho")
            except Exception:
                self.fixos = []
            self.registrar(f"  vermelhos fixos na tela: {self.fixos}", mostrar=False)
            if self.exe_tesseract:
                try:
                    self.colunas = localizar_colunas(tela_inicial, self.exe_tesseract)
                except Exception as erro:  # sem cabecalho, segue no modo antigo
                    self.registrar(f"  leitura do cabeçalho falhou: {erro}", mostrar=False)
            if self.colunas:
                self.registrar(f"  colunas pelo cabeçalho: {self.colunas}", mostrar=False)
            else:
                self.aviso("ATENÇÃO: não localizei o cabeçalho da grade (C/Custo, Conta, Quantidade). "
                      "Sem ele o robô não detecta C/Custo travado: o C/Custo será sempre digitado.")
            while time.time() - inicio_contagem < self.espera:
                if self.parar.is_set():
                    self.registrar("Interrompido antes de iniciar.")
                    return 1
                time.sleep(0.1)
            # fim da contagem: o cursor tem que estar no Item, senao nada e digitado
            proximo = caixa_foco(self.janela)  # caixa do Item da proxima linha (ja conhecida ao fim da anterior)
            if not proximo:
                self.registrar("Não iniciei: nenhum campo do SRP em foco (ciano). Clique no campo Item da primeira linha livre "
                          "e inicie de novo. Nada foi digitado.")
                return 1
            if self.colunas and "item" in self.colunas and qual_coluna(proximo, self.colunas) != "item":
                onde = qual_coluna(proximo, self.colunas)
                self.registrar(f"Não iniciei: o cursor está {f'na coluna {NOME_CAMPO[onde]}' if onde else 'fora da grade'}, "
                          "não no Item. Clique no campo Item da primeira linha livre e inicie de novo. Nada foi digitado.")
                return 1
            passo = 19  # distancia vertical entre linhas do formulario (px); recalculada quando observavel
            coluna_qtd = None  # (esq, dir, altura) da coluna Quantidade, aprendida uma vez
            for r in self.registros:
                if self.janela_trecho.lower() not in titulo_janela_ativa().lower():
                    self.registrar(f"Foco saiu do Citrix antes da linha {r['linha']}. Interrompido.")
                    return 1
                falha = self.revisar()
                while falha:
                    if not self.tratar_divergencia(falha):
                        return 1
                    falha = self.revisar()
                if self.pausar is not None and self.pausar.is_set() and self.retomar is not None:
                    proximo = self.pausa_pedida(r)
                self.marcar(r, "executando")
                t0 = time.time()
                inicial = proximo or caixa_foco(self.janela)
                if not inicial:
                    self.registrar(f"Linha {r['linha']}: campo em foco (ciano) não detectado. Interrompido.")
                    return 1
                foco, campo = inicial, "item"
                digitados, manual, proximo = [], False, None
                while campo != "fim":
                    if campo == "conta" and not r["conta"]:
                        self.registrar(f"Linha {r['linha']}: Conta editável mas em branco na planilha. Interrompido.")
                        return 1
                    if campo == "quantidade" and coluna_qtd is None:
                        coluna_qtd = (foco[0], foco[2], foco[3] - foco[1])
                    valor = r["quantidade"].replace(".", ",") if campo == "quantidade" else r[campo]
                    self.teclar("digitar", valor)
                    self.teclar("tab", None, esperar=False)
                    digitados.append(campo)
                    novo = aguardar_mudanca(self.janela, foco, self.render, ver_aviso=bool(self.colunas), fixos=self.fixos)
                    if novo == "aviso":
                        novo = None
                    seguinte = proximo_campo(campo, foco, novo, self.colunas) if novo else None
                    if seguinte is None and campo == "quantidade" and not tesseract and not self.colunas:
                        time.sleep(self.pausa)  # sem deteccao nem conferencia: sincroniza por tempo (modo antigo)
                        break
                    if seguinte is None:
                        if self.retomar is None or not self.colunas:
                            self.registrar(f"Linha {r['linha']}: o SRP não avançou depois do {NOME_CAMPO[campo]} "
                                      "(aviso na tela?). Interrompido.")
                            return 1
                        novo, seguinte = self.esperar_correcao(r, campo)
                        manual = True
                    foco, campo = novo, seguinte
                    if campo == "fim":
                        proximo = novo
                self.registrar(f"  linha {r['linha']}: digitados={digitados} manual={manual} tempo={time.time() - t0:.2f}s", mostrar=False)
                if tesseract and proximo and not manual and coluna_qtd:
                    # A celula so e lida depois do Tab (valor formatado, ex.: 1.0000): fica na linha acima do novo foco.
                    if proximo[1] - inicial[1] > 8:
                        passo = proximo[1] - inicial[1]
                    topo = proximo[1] - passo
                    celula = (coluna_qtd[0], topo, coluna_qtd[1], topo + coluna_qtd[2])
                    esperada = float(r["quantidade"].replace(",", "."))
                    recorte = recortar_celula(self.janela, celula)
                    self.pendentes.append((r["linha"], esperada, pool.submit(conferir_recorte, recorte, tesseract, esperada), recorte))
                elif tesseract and manual:
                    self.registrar(f"  linha {r['linha']}: conferência da quantidade pulada (linha ajustada a mão)", mostrar=False)
                conta = "digitada " + r["conta"] if "conta" in digitados else "travada"
                custo = r["custo"] if "custo" in digitados else "travado"
                self.registrar(f"OK linha {r['linha']}: {r['item']} c/custo={custo} conta={conta} {r['quantidade']}"
                          + (" (ajustada a mão)" if manual else ""))
                self.marcar(r, "ok")
            falha = self.revisar(aguardar=True)
            while falha:
                if not self.tratar_divergencia(falha, no_fim=True):
                    return 1
                falha = self.revisar(aguardar=True)
            self.registrar("Concluído.")
            return 0
        except KeyboardInterrupt:
            self.registrar("Interrompido pelo usuário.")
            return 1
        except Exception as erro:
            self.registrar(f"Erro: {erro}")
            return 1
        finally:
            pool.shutdown(wait=False, cancel_futures=True)
            if self.atual is not None and self.estado:
                self.estado(self.atual, "erro")
            self.log.close()


def executar(registros, janela_trecho, pausa, render, espera, parar, aviso, estado=None, conferir=False, retomar=None,
             pausar=None):
    """Preenche o SRP. `parar` e um threading.Event; `aviso(texto)` recebe o andamento;
    `estado(linha, nome)` (opcional) recebe 'executando', 'pausado', 'ok' ou 'erro' por linha;
    `conferir` le a Quantidade na tela (OCR) antes de confirmar cada linha;
    `retomar` (threading.Event, opcional): se o SRP nao avancar (aviso na tela), o robo PAUSA e espera esse
    evento (botao Continuar) em vez de interromper. Sem ele, interrompe como antes.
    `pausar` (threading.Event, opcional; F9): pausa ao terminar a linha atual e espera `retomar`. Devolve 0 se concluiu."""
    return Execucao0234(registros=registros, janela_trecho=janela_trecho, pausa=pausa, render=render, espera=espera,
                        parar=parar, aviso=aviso, estado=estado, conferir=conferir, retomar=retomar,
                        pausar=pausar).rodar()
