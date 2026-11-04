"""SNAC0236 (Digitacao Documento Recebimento): planilha, validacao e preenchimento pelo Citrix.

O SRP anda entre os campos com Enter. Sequencia por nota (combinada com o usuario em 04/11/2026):
  Chave NF + 5 Enter, Filial (nos campos Filial Fisica; 2 partes: 1a + 1 Enter, 2a + 2 Enter), Data Emissao + 7 Enter,
  Valor Total Documento + 1 Enter, Frete + 10 Enter (chega no Item da grade); por item:
  Item + 1 Enter (abre o Rateio), Oper SPM + 1 Enter, Qtd Fiscal + 2 Enter, Valor Unit + 1 Enter,
  Valor do ICMS + 7 Enter, ICMS Terceiro + 1 Enter; no fim da nota, clique em Integrar.
Linhas com a mesma CHAVE NF sao itens da mesma nota (cabecalho digitado uma vez).
Cada Enter so conta se o foco (ciano) andar; antes dos campos do cabecalho e do Item o robo confere, pelo OCR feito
no inicio, se o foco esta ao lado do rotulo certo. Qualquer duvida pausa (ou interrompe) sem digitar.
"""
import re
import threading
import time
import unicodedata
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import core

# coluna da planilha (normalizada) -> chave
COLUNAS = {"chave nf": "chave", "filial fiscal": "filial", "data emissao": "data", "valor total documento": "valor_total",
           "frete": "frete", "item": "item", "oper spm": "oper", "qtd fiscal": "qtd", "valor unit": "unit",
           "valor do icms": "icms", "icms terceiro": "icms_terceiro"}
TITULOS = {"chave": "CHAVE NF", "filial": "FILIAL FISCAL", "data": "DATA EMISSÃO", "valor_total": "VALOR TOTAL DOCUMENTO",
           "frete": "FRETE", "item": "ITEM", "oper": "OPER SPM", "qtd": "QTD FISCAL", "unit": "VALOR UNIT",
           "icms": "VALOR DO ICMS", "icms_terceiro": "ICMS TERCEIRO"}
CABECALHO = ("chave", "filial", "data", "valor_total", "frete")  # campos da nota (repetidos ou so na 1a linha)
OBRIGATORIOS = ("chave", "filial", "data", "valor_total", "item", "oper", "qtd", "unit")
FRETES = ("CIF", "FOB")  # o campo Frete do SRP recebe o tipo de frete, nao valor
VALORES = ("valor_total", "qtd", "unit", "icms", "icms_terceiro")

# (campo, quantos Enter depois de digitar). "filial1"/"filial2" = as duas partes da FILIAL FISCAL.
SEQ_CABECALHO = (("chave", 5), ("filial1", 1), ("filial2", 2), ("data", 7), ("valor_total", 1), ("frete", 10))
SEQ_ITEM = (("item", 1), ("oper", 1), ("qtd", 2), ("unit", 1), ("icms", 7), ("icms_terceiro", 1))
# a FILIAL FISCAL da planilha vai nos campos "Filial Fisica" da tela: depois da Chave + 5 Enter o SRP preenche CNPJ,
# Serie e Docto e o cursor cai na Filial Fisica (visto na tela real em 04/11/2026, monitor 4K)
NOME = {"chave": "Chave Acesso NF-e", "filial1": "Filial Fisica (1a parte)", "filial2": "Filial Fisica (2a parte)",
        "data": "Data Emissao", "valor_total": "Valor Total Documento", "frete": "Frete", "item": "Item",
        "oper": "Oper SPM", "qtd": "Qtd Fiscal", "unit": "Valor Unit", "icms": "Valor do ICMS",
        "icms_terceiro": "ICMS Terceiro"}


def normalizar(texto):
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", sem_acento).strip().strip(":").strip().lower()


def chave_coluna(titulo):
    nome = normalizar(titulo)
    if nome.startswith("icms terceiro"):  # "ICMS TERCEIRO (BASE DE CALCULO NF)"
        return "icms_terceiro"
    return COLUNAS.get(nome)


def formatar_valor(valor):
    """Numero como o SRP quer: ponto decimal, sem milhar (47407.28). Devolve (texto, erro)."""
    if valor is None or valor == "":
        return "", None
    if isinstance(valor, bool):
        return "", "valor invalido"
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and valor.is_integer():
            valor = int(valor)
        texto = format(Decimal(repr(valor)), "f") if isinstance(valor, float) else str(valor)
    else:
        texto = str(valor).strip().replace(" ", "")
        if texto.count(",") == 1 and "." not in texto:
            texto = texto.replace(",", ".")  # 47407,28 digitado como texto
    if not re.fullmatch(r"\d+(?:\.\d+)?", texto):
        return "", f"valor invalido: {valor!r} (use so numeros, com ponto ou virgula decimal, sem separador de milhar)"
    try:
        Decimal(texto)
    except InvalidOperation:
        return "", f"valor invalido: {valor!r}"
    return texto, None


def formatar_data(valor):
    if valor is None or valor == "":
        return "", None
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%Y"), None
    texto = str(valor).strip()
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", texto):
        try:
            datetime.strptime(texto, "%d/%m/%Y")
            return texto, None
        except ValueError:
            pass
    return "", f"data invalida: {valor!r} (use dd/mm/aaaa)"


def formatar_chave(valor):
    if valor is None or valor == "":
        return "", None
    if not isinstance(valor, str):
        # o Excel guarda numero com 15 digitos de precisao: os 44 da chave ja chegam errados
        return "", "a CHAVE NF esta como numero no Excel e perde digitos: formate a coluna como Texto e cole a chave de novo"
    texto = re.sub(r"\s", "", valor)
    if not re.fullmatch(r"\d{44}", texto):
        return "", f"a CHAVE NF precisa ter 44 digitos (tem {len(re.sub(r'[^0-9]', '', texto))})"
    return texto, None


def formatar_filial(valor):
    """'30 46' -> ('30', '46'). Devolve ((1a, 2a), erro)."""
    if valor is None or valor == "":
        return ("", ""), None
    partes = [p for p in re.split(r"[\s/\-]+", core.texto(valor)) if p]
    if len(partes) != 2 or not all(p.isdigit() for p in partes):
        return ("", ""), f"FILIAL FISCAL precisa de 2 partes separadas por espaco (ex.: 30 46), veio {valor!r}"
    return tuple(partes), None


def ler_planilha(caminho):
    """Linhas da planilha ja formatadas para o SRP. Devolve (registros, erros)."""
    from openpyxl import load_workbook

    ws = load_workbook(caminho, read_only=True, data_only=True).active
    linhas = ws.iter_rows(values_only=True)
    cabecalho = next(linhas, None) or ()
    indices = {}
    for posicao, titulo in enumerate(cabecalho):
        chave = chave_coluna(titulo) if titulo is not None else None
        if chave and chave not in indices:
            indices[chave] = posicao
    faltando = [TITULOS[c] for c in TITULOS if c not in indices]
    if faltando:
        raise ValueError(f"Colunas ausentes no cabecalho: {', '.join(faltando)}")
    registros, erros = [], []
    for numero, linha in enumerate(linhas, start=2):
        bruto = {c: (linha[p] if p < len(linha) else None) for c, p in indices.items()}
        if all(v is None or str(v).strip() == "" for v in bruto.values()):
            continue
        r = {"linha": numero}
        problemas = []
        r["chave"], erro = formatar_chave(bruto["chave"])
        problemas.append(erro)
        (r["filial1"], r["filial2"]), erro = formatar_filial(bruto["filial"])
        r["filial"] = f"{r['filial1']} {r['filial2']}".strip()
        problemas.append(erro)
        r["data"], erro = formatar_data(bruto["data"])
        problemas.append(erro)
        r["frete"] = core.texto(bruto["frete"]).upper()
        if r["frete"] and r["frete"] not in FRETES:
            problemas.append(f"FRETE invalido: {bruto['frete']!r} (use CIF ou FOB)")
        for campo in VALORES:
            r[campo], erro = formatar_valor(bruto[campo])
            problemas.append(f"{TITULOS[campo]}: {erro}" if erro else None)
        for campo in ("item", "oper"):
            r[campo] = core.texto(bruto[campo])
            if r[campo] and not re.fullmatch(r"[0-9A-Za-z.\-]+", r[campo]):
                problemas.append(f"{TITULOS[campo]} com caracteres invalidos: {r[campo]!r}")
        erros += [f"linha {numero}: {p}" for p in problemas if p]
        registros.append(r)
    return registros, erros


def agrupar_notas(registros):
    """Notas na ordem da planilha: {'chave', 'filial1', 'filial2', 'data', 'valor_total', 'frete', 'itens': [registros]}.
    O cabecalho vem da 1a linha da chave; nas outras pode estar em branco ou igual. Devolve (notas, erros)."""
    notas, erros, vistas = [], [], {}
    for r in registros:
        if r["chave"] and notas and r["chave"] != notas[-1]["chave"] and r["chave"] in vistas:
            erros.append(f"linha {r['linha']}: a CHAVE NF ...{r['chave'][-6:]} ja apareceu antes (linha {vistas[r['chave']]}); "
                         "deixe os itens da mesma nota em linhas seguidas")
        if not r["chave"] or not notas or r["chave"] != notas[-1]["chave"]:
            if not r["chave"] and notas:
                erros.append(f"linha {r['linha']}: CHAVE NF em branco (repita a chave em todos os itens da nota)")
                continue
            vistas.setdefault(r["chave"], r["linha"])
            notas.append({c: r[c] for c in ("chave", "filial1", "filial2", "data", "valor_total", "frete")} | {"itens": [r]})
            continue
        nota = notas[-1]
        for campo in ("filial1", "filial2", "data", "valor_total", "frete"):
            if r[campo] and r[campo] != nota[campo]:
                titulo = TITULOS["filial" if campo.startswith("filial") else campo]
                erros.append(f"linha {r['linha']}: {titulo} diferente da 1a linha da mesma nota ({r[campo]} x {nota[campo]})")
        nota["itens"].append(r)
    return notas, erros


def validar(registros, notas):
    erros = []
    for nota in notas:
        primeira = nota["itens"][0]["linha"]
        for campo in ("chave", "data", "valor_total", "frete"):
            if not nota[campo]:
                erros.append(f"linha {primeira}: {TITULOS[campo]} em branco")
        if not nota["filial1"]:
            erros.append(f"linha {primeira}: FILIAL FISCAL em branco")
        for r in nota["itens"]:
            for campo in ("item", "oper", "qtd", "unit"):
                if not r[campo]:
                    erros.append(f"linha {r['linha']}: {TITULOS[campo]} em branco")
    return erros


def carregar(caminho):
    """(registros, notas, erros) prontos para a tela."""
    registros, erros = ler_planilha(caminho)
    notas, erros_notas = agrupar_notas(registros)
    return registros, notas, erros + erros_notas + validar(registros, notas)


def valor_campo(nota, r, campo):
    return nota[campo] if campo in ("chave", "filial1", "filial2", "data", "valor_total", "frete") else r[campo]


def descrever(nota):
    """Sequencia de teclas da nota, para o Simular."""
    partes = [f"'{nota[c]}' {n}xENTER" for c, n in SEQ_CABECALHO]
    for r in nota["itens"]:
        partes.append(f"| linha {r['linha']}: " + " ".join(f"'{r[c]}' {n}xENTER" for c, n in SEQ_ITEM))
    return " ".join(partes) + " | clicar Integrar"


# ---------- tela ----------
def frases(palavras, folga=15):
    """Junta palavras vizinhas da mesma linha em frases (rotulos): [(texto normalizado, (esq, topo, dir, base))]."""
    restantes = sorted(palavras, key=lambda p: (p[2], p[1]))
    grupos = []
    for texto, esq, topo, larg, alt in sorted(restantes, key=lambda p: p[1]):
        centro = topo + alt / 2
        for g in grupos:
            e, t, d, b = g[1]
            if abs((t + b) / 2 - centro) <= max(6, alt / 2) and 0 <= esq - d <= folga:
                g[0].append(texto)
                g[1] = (e, min(t, topo), esq + larg, max(b, topo + alt))
                break
        else:
            grupos.append([[texto], (esq, topo, esq + larg, topo + alt)])
    return [(normalizar(" ".join(g[0])), tuple(g[1])) for g in grupos]


def localizar_rotulos(palavras):
    """Retangulos dos rotulos da tela SNAC0236 lidos no OCR: chave, filial, data, valor_total, frete, item (cabecalho da
    coluna da grade) e integrar (botao). Os que nao forem achados ficam de fora (o robo nao confere aquele campo)."""
    lista = frases(palavras)
    achados = {}

    def primeiro(pred):
        return next((caixa for texto, caixa in lista if pred(texto)), None)

    regras = {
        "chave": lambda t: t.startswith("chave"),
        "filial": lambda t: t.startswith("filial fisica"),
        "data": lambda t: t.startswith("data emiss"),
        "valor_total": lambda t: t.startswith("valor total"),
        "frete": lambda t: t == "frete",
        "integrar": lambda t: t == "integrar",
    }
    for campo, pred in regras.items():
        caixa = primeiro(pred)
        if caixa:
            achados[campo] = caixa
    descricoes = [caixa for texto, caixa in lista if texto.startswith("descri")]
    for texto, caixa in lista:
        # em 4K o Tesseract le o "Item" do cabecalho da grade como "tem"
        if texto in ("item", "tem", "ltem", "1tem") and any(abs((caixa[1] + caixa[3]) / 2 - (d[1] + d[3]) / 2) <= 14 for d in descricoes):
            achados["item"] = caixa
            break
    return achados


def binarizar(imagem, limiar=80):
    """Texto escuro em preto, o resto em branco: os botoes do SRP (texto preto em fundo cinza) so sao lidos assim."""
    return imagem.convert("L").point(lambda v: 0 if v < limiar else 255).convert("RGB")


def ler_rotulos(imagem, tesseract, guardar=None):
    """localizar_rotulos com uma segunda leitura (imagem binarizada) para os rotulos que faltarem (ex.: Integrar).
    `guardar` (lista, opcional) recebe as frases lidas, para dizer onde o cursor esta quando algo nao confere."""
    palavras = core.palavras_tela(imagem, tesseract)
    rotulos = localizar_rotulos(palavras)
    if guardar is not None:
        guardar.extend(frases(palavras))
    if any(c not in rotulos for c in ("chave", "filial", "data", "valor_total", "frete", "item", "integrar")):
        palavras = core.palavras_tela(binarizar(imagem), tesseract)
        for campo, caixa in localizar_rotulos(palavras).items():
            rotulos.setdefault(campo, caixa)
    return rotulos


def onde_esta(caixa, lista):
    """Rotulo mais proximo do campo em foco: a frase a esquerda na mesma linha ou, na grade, o cabecalho acima."""
    if not caixa or not lista:
        return None
    centro_y, centro_x = (caixa[1] + caixa[3]) / 2, (caixa[0] + caixa[2]) / 2
    esquerda = [(caixa[0] - c[2], t) for t, c in lista
                if abs((c[1] + c[3]) / 2 - centro_y) <= max(8, (caixa[3] - caixa[1]) / 2 + 4) and c[2] <= caixa[0] + 5]
    if esquerda:
        return min(esquerda)[1]
    acima = [(caixa[1] - c[3], t) for t, c in lista
             if c[3] <= caixa[1] + 5 and abs((c[0] + c[2]) / 2 - centro_x) <= max(40, caixa[2] - caixa[0])]
    return min(acima)[1] if acima else None


def ao_lado(caixa, rotulo, distancia=60):
    """O campo em foco esta logo a direita do rotulo, na mesma linha?"""
    centro, centro_rotulo = (caixa[1] + caixa[3]) / 2, (rotulo[1] + rotulo[3]) / 2
    return abs(centro - centro_rotulo) <= max(8, (caixa[3] - caixa[1]) / 2 + 4) and 0 <= caixa[0] - rotulo[2] <= distancia


def na_coluna(caixa, rotulo):
    """O campo em foco esta na coluna do rotulo (cabecalho da grade), abaixo dele?"""
    centro, centro_rotulo = (caixa[0] + caixa[2]) / 2, (rotulo[0] + rotulo[2]) / 2
    return abs(centro - centro_rotulo) <= max(40, caixa[2] - caixa[0]) and caixa[1] > rotulo[1]


def posicao_ok(campo, caixa, rotulos, anterior=None):
    """True/False se da para conferir o campo pela posicao; None se o rotulo nao foi lido (nao confere)."""
    if campo == "filial2":
        rotulo = rotulos.get("filial")
        if not rotulo or not anterior:
            return None
        return abs((caixa[1] + caixa[3]) / 2 - (anterior[1] + anterior[3]) / 2) <= 6 and caixa[0] > anterior[2]
    if campo == "item":
        return na_coluna(caixa, rotulos["item"]) if "item" in rotulos else None
    chave = "filial" if campo == "filial1" else campo
    if chave in rotulos:
        return ao_lado(caixa, rotulos[chave])
    return None


def clicar(x, y):
    """Clique na tela (coordenadas fisicas)."""
    from pywinauto import mouse
    mouse.click(button="left", coords=(int(x), int(y)))


class Travou(Exception):
    def __init__(self, campo, motivo):
        super().__init__(motivo)
        self.campo, self.motivo = campo, motivo


def executar(notas, janela_trecho, pausa, render, espera, parar, aviso, estado=None, retomar=None, intervalo=0):
    """Preenche as notas no SRP. `intervalo` (s): espera extra antes de cada acao (digitar, Enter, Integrar) e narra a
    acao no console, para o usuario acompanhar o percurso nos testes (0 = velocidade normal). `estado(linha, nome)` por linha da planilha; `retomar` (Event, botao Continuar) liga
    as pausas: depois de cada Integrar e quando algo nao confere. Sem ele, qualquer problema interrompe. Devolve 0 se concluiu."""
    core.limpar_logs()
    log = open(core.pasta_logs() / f"{datetime.now():%Y%m%d-%H%M%S}-snac0236.log", "w", encoding="utf-8")
    trava_log = threading.Lock()
    atual = []  # linhas marcadas 'executando'

    def registrar(texto, mostrar=True):
        with trava_log:
            if not log.closed:
                log.write(texto + "\n")
                log.flush()
        if mostrar:
            aviso(texto)

    def marcar(linhas, nome):
        for linha in linhas:
            if estado:
                estado(linha, nome)
        atual[:] = list(linhas) if nome == "executando" else [l for l in atual if l not in linhas]

    def checar_parada():
        if parar.is_set() or core.mouse_no_canto():
            raise KeyboardInterrupt

    def devagar(texto):
        """Modo passo a passo: anuncia a proxima acao e espera `intervalo` segundos (F8/Parar/mouse no canto valem)."""
        if intervalo <= 0:
            return
        aviso(f"Passo a passo: {texto}")
        fim = time.time() + intervalo
        while time.time() < fim:
            checar_parada()
            time.sleep(0.1)

    def digitar(texto, campo):
        checar_parada()
        devagar(f"digitar '{texto}' em {NOME[campo]}" if texto else f"{NOME[campo]} fica em branco")
        if texto:
            core.send_keys(texto.replace(",", "{,}"), pause=0.03)
            time.sleep(min(pausa, 0.05))

    def enters(campo, quantos, foco):
        """Enter `quantos` vezes; cada um precisa mover o foco (senao: aviso do SRP ou campo preso)."""
        for vez in range(quantos):
            checar_parada()
            devagar(f"Enter {vez + 1} de {quantos} depois do {NOME[campo]}")
            core.send_keys("{ENTER}", pause=0.03)
            novo = core.aguardar_mudanca(janela, foco, render, ver_aviso=True, fixos=fixos)
            if novo in (None, "aviso"):
                motivo = "apareceu um aviso do SRP" if novo == "aviso" else "o cursor nao saiu do lugar"
                raise Travou(campo, f"{motivo} no Enter {vez + 1} de {quantos} depois do {NOME[campo]}")
            foco = novo
        return foco

    def local(caixa):
        texto = onde_esta(caixa, lidas)
        return f" (o cursor esta em '{texto}')" if texto else ""

    def ir_ate(campo, alvo, quantos, foco, anterior):
        """Depois de digitar `campo`, Enter ate o foco chegar em `alvo` (conferido pelo rotulo lido no inicio), no maximo
        `quantos` + 4 vezes: o SRP pula ou nao campos que ele mesmo preenche (CNPJ, Filial Fiscal...). Sem rotulo para
        conferir, so conta os `quantos` Enter combinados."""
        if posicao_ok(alvo, foco, rotulos, anterior) is None:
            return enters(campo, quantos, foco)
        limite = quantos + 4
        for vez in range(limite):
            checar_parada()
            devagar(f"Enter {vez + 1} depois do {NOME[campo]} (procurando {NOME[alvo]}; previstos {quantos})")
            core.send_keys("{ENTER}", pause=0.03)
            novo = core.aguardar_mudanca(janela, foco, render, ver_aviso=True, fixos=fixos)
            if novo in (None, "aviso"):
                motivo = "apareceu um aviso do SRP" if novo == "aviso" else "o cursor nao saiu do lugar"
                raise Travou(campo, f"{motivo} no Enter {vez + 1} depois do {NOME[campo]}")
            foco = novo
            if posicao_ok(alvo, foco, rotulos, anterior):
                if vez + 1 != quantos:
                    registrar(f"  {NOME[alvo]}: cheguei com {vez + 1} Enter (previstos {quantos})", mostrar=False)
                return foco
        raise Travou(alvo, f"nao cheguei no campo {NOME[alvo]} com {limite} Enter depois do {NOME[campo]}{local(foco)}")

    def esperar_continuar():
        retomar.clear()
        while not retomar.is_set():
            if parar.is_set():
                raise KeyboardInterrupt
            time.sleep(0.2)

    def foco_depois_de_continuar():
        janela.set_focus()
        aviso("Continuando em 2 s...")
        for _ in range(20):
            checar_parada()
            time.sleep(0.1)
        return core.caixa_foco(janela)

    def esperar_na_chave(texto, linhas):
        """Pausa ate o cursor estar na Chave Acesso NF-e (proxima nota) e o usuario clicar em Continuar."""
        registrar(texto)
        marcar(linhas, "pausado")
        while True:
            esperar_continuar()
            caixa = foco_depois_de_continuar()
            if caixa and posicao_ok("chave", caixa, rotulos) is not False:
                return caixa
            registrar("Ainda pausado: deixe o cursor no campo Chave Acesso NF-e (vazio) e clique em Continuar.")
            marcar(linhas, "pausado")

    def esperar_no_item(r):
        registrar(f"Linha {r['linha']}: o cursor nao esta no Item da grade. PAUSADO. Deixe o cursor no Item da proxima "
                  "linha livre da grade e clique em Continuar.")
        marcar([r["linha"]], "pausado")
        while True:
            esperar_continuar()
            caixa = foco_depois_de_continuar()
            if caixa and posicao_ok("item", caixa, rotulos) is not False:
                marcar([r["linha"]], "executando")
                return caixa
            registrar("Ainda pausado: deixe o cursor no Item da proxima linha livre da grade e clique em Continuar.")
            marcar([r["linha"]], "pausado")

    def localizar_integrar():
        if not tesseract:
            return None
        try:
            # le so a faixa em volta de onde o botao estava no inicio (mais rapido); sem posicao inicial, a tela toda
            tela_agora = core.capturar(janela)
            antes = rotulos.get("integrar")
            if antes:
                esq, topo = max(antes[0] - 300, 0), max(antes[1] - 80, 0)
                recorte = tela_agora.crop((esq, topo, min(antes[2] + 300, tela_agora.width), min(antes[3] + 80, tela_agora.height)))
                caixa = localizar_rotulos(core.palavras_tela(binarizar(recorte), tesseract)).get("integrar")
                achado = (caixa[0] + esq, caixa[1] + topo, caixa[2] + esq, caixa[3] + topo) if caixa else None
            else:
                achado = ler_rotulos(tela_agora, tesseract).get("integrar")
        except Exception as erro:
            registrar(f"  leitura do botao Integrar falhou: {erro}", mostrar=False)
            return None
        return achado

    def integrar(nota):
        """Clica em Integrar se o botao estiver a vista (confere de novo pelo OCR antes de clicar). True = clicou."""
        botao = localizar_integrar()
        if not botao:
            return False
        ret = janela.rectangle()
        x, y = ret.left + (botao[0] + botao[2]) / 2, ret.top + (botao[1] + botao[3]) / 2
        checar_parada()
        devagar("clicar em Integrar")
        clicar(x, y)
        registrar(f"  nota ...{nota['chave'][-6:]}: clique em Integrar em ({x:.0f}, {y:.0f})", mostrar=False)
        return True

    try:
        tesseract = core.caminho_tesseract()
        janela = core.janela_citrix(janela_trecho)
        janela.set_focus()
        aviso(f"Clique no campo Chave Acesso NF-e. Iniciando em {espera}s (mouse no canto superior esquerdo aborta).")
        inicio = time.time()
        tela = core.capturar(janela)
        ret = janela.rectangle()
        problema = core.problema_tela((ret.left, ret.top, ret.right, ret.bottom), tela)
        if problema:
            registrar(f"Nao iniciei: {problema.replace('no campo Item', 'no campo Chave Acesso NF-e')}. Nada foi digitado.")
            return 1
        try:
            fixos = core.caixas_cor(tela, "vermelho")
        except Exception:
            fixos = []
        rotulos, lidas = {}, []
        if tesseract:
            try:
                rotulos = ler_rotulos(tela, tesseract, lidas)
            except Exception as erro:
                registrar(f"  leitura dos rotulos falhou: {erro}", mostrar=False)
        registrar(f"  rotulos lidos: {rotulos}", mostrar=False)
        faltam = [NOME.get(c, c) for c in ("chave", "filial", "data", "valor_total", "frete", "item") if c not in rotulos]
        if faltam:
            aviso(f"ATENCAO: nao li na tela os rotulos {', '.join(faltam)}: nesses campos o robo so conta os Enter, sem conferir a posicao.")
        if "integrar" not in rotulos:
            aviso("ATENCAO: nao achei o botao Integrar na tela: no fim de cada nota o robo pausa para voce clicar.")
        while time.time() - inicio < espera:
            if parar.is_set():
                registrar("Interrompido antes de iniciar.")
                return 1
            time.sleep(0.1)
        foco = core.caixa_foco(janela)
        if not foco:
            registrar("Nao iniciei: nenhum campo do SRP em foco (ciano). Clique no campo Chave Acesso NF-e e inicie de novo. "
                      "Nada foi digitado.")
            return 1
        if posicao_ok("chave", foco, rotulos) is False:
            registrar("Nao iniciei: o cursor nao esta no campo Chave Acesso NF-e. Clique nele e inicie de novo. Nada foi digitado.")
            return 1

        for numero, nota in enumerate(notas, start=1):
            linhas = [r["linha"] for r in nota["itens"]]
            ultima = numero == len(notas)
            if core.titulo_janela_ativa() and janela_trecho.lower() not in core.titulo_janela_ativa().lower():
                registrar(f"Foco saiu do Citrix antes da nota da linha {linhas[0]}. Interrompido.")
                return 1
            t0 = time.time()
            etiqueta = f"Nota ...{nota['chave'][-6:]} (linha{'s' if len(linhas) > 1 else ''} {linhas[0]}" + \
                       (f"-{linhas[-1]})" if len(linhas) > 1 else ")")
            try:
                marcar([linhas[0]], "executando")
                if posicao_ok("chave", foco, rotulos) is False:
                    raise Travou("chave", f"o cursor nao esta no campo {NOME['chave']}{local(foco)}")
                for indice, (campo, quantos) in enumerate(SEQ_CABECALHO):
                    digitar(nota[campo], campo)
                    alvo = SEQ_CABECALHO[indice + 1][0] if indice + 1 < len(SEQ_CABECALHO) else "item"
                    foco = ir_ate(campo, alvo, quantos, foco, anterior=foco)
                for indice, r in enumerate(nota["itens"]):
                    marcar([r["linha"]], "executando")
                    if posicao_ok("item", foco, rotulos) is False:
                        if retomar is None:
                            raise Travou("item", f"o cursor nao esta no Item da grade{local(foco)}")
                        foco = esperar_no_item(r)
                    for campo, quantos in SEQ_ITEM:
                        digitar(r[campo], campo)
                        foco = enters(campo, quantos, foco)
                    registrar(f"  linha {r['linha']}: item {r['item']} digitado", mostrar=False)
                    if indice < len(nota["itens"]) - 1:
                        marcar([r["linha"]], "ok")
            except Travou as travou:
                if retomar is None:
                    marcar(linhas, "erro")
                    registrar(f"{etiqueta}: {travou.motivo}. Interrompido. A nota ficou pela metade no SRP: termine ou "
                              "cancele a mao.")
                    return 1
                if ultima:
                    registrar(f"{etiqueta}: {travou.motivo}. PAUSADO. Termine esta nota a mao no SRP (inclusive Integrar) "
                              "e clique em Continuar.")
                    marcar(linhas, "pausado")
                    esperar_continuar()
                else:
                    foco = esperar_na_chave(
                        f"{etiqueta}: {travou.motivo}. PAUSADO. Termine esta nota a mao no SRP (inclusive Integrar), "
                        "deixe o cursor na Chave Acesso NF-e da proxima nota e clique em Continuar.", linhas)
                registrar(f"OK {etiqueta}: terminada a mao.")
                marcar(linhas, "ok")
                continue
            clicou = integrar(nota)
            registrar(f"OK {etiqueta}: {len(linhas)} item(ns) em {time.time() - t0:.1f}s" +
                      ("; cliquei em Integrar." if clicou else "."))
            if not clicou:
                registrar(f"{etiqueta}: nao achei o botao Integrar na tela. Clique em Integrar no SRP.")
            marcar(linhas, "ok")
            if ultima:
                if not clicou and retomar is not None:
                    registrar("Depois de integrar, clique em Continuar para encerrar.")
                    marcar(linhas, "pausado")
                    esperar_continuar()
                    marcar(linhas, "ok")
                break
            if retomar is None:
                registrar("Sem o botao Continuar o robo para aqui: confira a integracao e inicie de novo pela proxima nota.")
                return 1
            foco = esperar_na_chave(
                "PAUSADO: confira a integracao no SRP. Para a proxima nota, deixe o cursor na Chave Acesso NF-e e clique "
                "em Continuar.", [notas[numero]["itens"][0]["linha"]])
        registrar("Concluido. Confira a integracao da ultima nota no SRP.")
        return 0
    except KeyboardInterrupt:
        registrar("Interrompido pelo usuario.")
        return 1
    except Exception as erro:
        registrar(f"Erro: {erro}")
        return 1
    finally:
        if atual and estado:
            for linha in atual:
                estado(linha, "erro")
        log.close()
