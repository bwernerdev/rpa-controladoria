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
import json
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
           "valor do icms": "icms", "icms terceiro": "icms_terceiro", "aliquota": "aliquota_icms"}
TITULOS = {"chave": "CHAVE NF", "filial": "FILIAL FISCAL", "data": "DATA EMISSÃO", "valor_total": "VALOR TOTAL DOCUMENTO",
           "frete": "FRETE", "item": "ITEM", "oper": "OPER SPM", "qtd": "QTD FISCAL", "unit": "VALOR UNIT",
           "icms": "VALOR DO ICMS", "icms_terceiro": "BASE CALC. ICMS", "aliquota_icms": "ALIQUOTA"}
CABECALHO = ("chave", "filial", "data", "valor_total", "frete")  # campos da nota (repetidos ou so na 1a linha)
OBRIGATORIOS = ("chave", "filial", "data", "valor_total", "item", "oper", "qtd", "unit")
FRETES = ("CIF", "FOB")  # o campo Frete do SRP recebe o tipo de frete, nao valor
VALORES = ("valor_total", "qtd", "unit", "icms", "icms_terceiro", "aliquota_icms")

# (campo, quantos Enter depois de digitar). "filial1"/"filial2" = as duas partes da FILIAL FISCAL.
ENTERS_FRETE_ITEM = 10  # Enter previstos do Frete ao Item da grade (Marcas, Observacoes, Cond. Pagamento, lista sem cor...)
SEQ_CABECALHO = (("chave", 5), ("filial1", 1), ("filial2", 2), ("data", 7), ("valor_total", 1), ("frete", 10))
# depois do Oper SPM o 1o Enter cai no Nat Oper (ja preenchido; nao mexer) e o 2o no Qtd Fiscal (04/11/2026)
# depois do Valor do ICMS: Enter ate a linha ICMS TERCEIROS do Rateio (`preencher_rateio`, sem contagem fixa)
SEQ_ITEM = (("item", 1), ("oper", 2), ("qtd", 2), ("unit", 1), ("icms", 0))
# a FILIAL FISCAL da planilha vai nos campos "Filial Fisica" da tela: depois da Chave + 5 Enter o SRP preenche CNPJ,
# Serie e Docto e o cursor cai na Filial Fisica (visto na tela real em 04/11/2026, monitor 4K)
NOME = {"chave": "Chave Acesso NF-e", "filial1": "Filial Fisica (1a parte)", "filial2": "Filial Fisica (2a parte)",
        "data": "Data Emissao", "valor_total": "Valor Total Documento", "frete": "Frete", "item": "Item",
        "oper": "Oper SPM", "qtd": "Qtd Fiscal", "unit": "Valor Unit", "icms": "Valor do ICMS",
        "icms_terceiro": "ICMS Terceiro", "rateio": "Rateio", "ident": "Ident Tributo (ICMS TERCEIROS)",
        "base": "Base Calculo (ICMS TERCEIROS)", "aliquota": "Aliquota (ICMS TERCEIROS)",
        "lancamento": "Lancamento (ICMS TERCEIROS)", "serie": "Serie"}
# depois da Chave + Enter o cursor passa por CNPJ (2 campos), Serie, Docto e cai na Filial Fisica (log de 09/11/2026)
ENTERS_CHAVE_SERIE = 3


def normalizar(texto):
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", sem_acento).strip().strip(":").strip().lower()


def chave_coluna(titulo):
    nome = normalizar(titulo)
    # base de calculo da linha ICMS TERCEIROS do Rateio: "BASE CALC. ICMS" desde 05/11/2026 (antes
    # "ICMS TERCEIROS (BASE DE CALCULO NF)"; o nome antigo continua aceito)
    if nome.startswith("icms terceiro") or nome.startswith("base calc"):
        return "icms_terceiro"
    if nome.startswith("aliquota"):  # "ALIQUOTA" (da linha ICMS TERCEIROS do Rateio; era BASE CALC. ICMS ate 05/11/2026)
        return "aliquota_icms"
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

    # read_only deixa o arquivo aberto ate o close(): sem ele a planilha ficava travada e o Excel nao conseguia salvar
    # por cima (06/11/2026: "salvar e fechar" no Iniciar travou o Excel numa janela de erro)
    livro = load_workbook(caminho, read_only=True, data_only=True)
    try:
        celulas = list(livro.active.iter_rows())
        tema = core.cores_do_tema(getattr(livro, "loaded_theme", None))
    finally:
        livro.close()
    linhas = iter([tuple(c.value for c in linha) for linha in celulas])
    cabecalho = next(linhas, None) or ()
    indices, repetidas = {}, []
    for posicao, titulo in enumerate(cabecalho):
        chave = chave_coluna(titulo) if titulo is not None else None
        if chave and chave not in indices:
            indices[chave] = posicao
        elif chave == "icms_terceiro":  # nome novo e antigo juntos: nao adivinhar qual e a base
            repetidas.append(str(titulo))
    if repetidas:
        raise ValueError(f"Duas colunas para a base do ICMS TERCEIROS ({cabecalho[indices['icms_terceiro']]} e "
                         f"{', '.join(repetidas)}): deixe so a {TITULOS['icms_terceiro']}")
    faltando = [TITULOS[c] for c in TITULOS if c not in indices]
    if faltando:
        raise ValueError(f"Colunas ausentes no cabecalho: {', '.join(faltando)}")
    registros, erros = [], []
    for numero, linha in enumerate(linhas, start=2):
        bruto = {c: (linha[p] if p < len(linha) else None) for c, p in indices.items()}
        if all(v is None or str(v).strip() == "" for v in bruto.values()):
            continue
        # linha pintada de verde (pelo robo ao concluir a nota, ou a mao): o app trata como ja feita (usuario, 06/11/2026).
        # Verde = todas as celulas preenchidas das colunas lidas com fundo verde.
        preenchidas = [celulas[numero - 1][p] for c, p in indices.items()
                       if p < len(celulas[numero - 1]) and not (bruto[c] is None or str(bruto[c]).strip() == "")]
        r = {"linha": numero, "verde": all(core.fundo_verde(getattr(c, "fill", None), tema) for c in preenchidas)}
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


CABECALHO_SEM_CHAVE = ("filial1", "filial2", "data", "valor_total", "frete")


def agrupar_notas(registros):
    """Notas na ordem da planilha: {'chave', 'filial1', 'filial2', 'data', 'valor_total', 'frete', 'itens': [registros]}.
    O cabecalho vem da 1a linha da chave; nas outras pode estar em branco ou igual. Linha com a CHAVE NF e todo o
    cabecalho em branco = mais um item da nota de cima (como o usuario preenche, 05/11/2026). Devolve (notas, erros)."""
    notas, erros, vistas = [], [], {}
    for r in registros:
        if r["chave"] and notas and r["chave"] != notas[-1]["chave"] and r["chave"] in vistas:
            erros.append(f"linha {r['linha']}: a CHAVE NF ...{r['chave'][-6:]} ja apareceu antes (linha {vistas[r['chave']]}); "
                         "deixe os itens da mesma nota em linhas seguidas")
        if not r["chave"] or not notas or r["chave"] != notas[-1]["chave"]:
            if not r["chave"] and notas:
                if any(r[c] for c in CABECALHO_SEM_CHAVE):  # cabecalho de outra nota sem a chave: nao adivinhar
                    erros.append(f"linha {r['linha']}: CHAVE NF em branco com o cabecalho preenchido (para mais um item "
                                 "da nota de cima, deixe a chave e o cabecalho em branco ou repita a chave)")
                else:
                    notas[-1]["itens"].append(r)
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
            terceiro, aliquota = (Decimal(r.get(c) or "0") for c in ("icms_terceiro", "aliquota_icms"))
            if (terceiro != 0) != (aliquota != 0):
                erros.append(f"linha {r['linha']}: {TITULOS['icms_terceiro']} e {TITULOS['aliquota_icms']} precisam ser os dois 0 "
                             "(sem imposto) ou os dois diferentes de 0 (com imposto)")
            elif 0 < aliquota < 1:  # celula em porcentagem no Excel: 12% chega como 0,12 e o SRP receberia 0.12
                erros.append(f"linha {r['linha']}: ALIQUOTA {r['aliquota_icms']} parece porcentagem do Excel (12% vira 0,12): "
                             "digite o numero da aliquota, ex.: 12")
            elif aliquota < 0 or aliquota > 100:
                erros.append(f"linha {r['linha']}: ALIQUOTA invalida: {r['aliquota_icms']}")
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
        partes.append(f"| linha {r['linha']}: " + " ".join(f"'{r[c]}' {n}xENTER" for c, n in SEQ_ITEM)
                      + " ENTER ate ICMS TERCEIROS no Rateio: "
                      + " ".join(f"{NOME_RATEIO[c]}='{v}'" for c, v in valores_rateio(r).items()) + " | Tributos OK")
    return " ".join(partes) + " | clicar Integrar"


# tempo previsto: nota de 1 item ~37-43 s no SRP real (log de 05/11/2026 18:21, Citrix 1366x768); cada item a mais
# repete a grade e o Rateio (~25 s, chute ate medir). Depois da 1a nota, a medida corrige a conta (`tempo_previsto`).
SEG_NOTA, SEG_ITEM_EXTRA = 40.0, 25.0


def custo_previsto(notas):
    """Segundos previstos para as `notas` pelos tempos de referencia (antes de medir)."""
    return sum(SEG_NOTA + SEG_ITEM_EXTRA * (len(n["itens"]) - 1) for n in notas)


def tempo_previsto(feitas, faltam, gasto):
    """Segundos que faltam para as notas `faltam`, depois de `feitas` terem levado `gasto` segundos (sem as pausas).
    Sem nota feita = so os tempos de referencia; com notas feitas = a referencia corrigida pelo ritmo medido (o fator
    respeita a mistura de notas de 1 e de varios itens)."""
    restante = custo_previsto(faltam)
    referencia = custo_previsto(feitas)
    if not feitas or referencia <= 0 or gasto <= 0:
        return restante
    return restante * gasto / referencia


def texto_duracao(segundos):
    """'~1 h 52 min', '~14 min', '~40 s' (aproximado, para o usuario)."""
    segundos = max(0, round(segundos))
    if segundos >= 3600:
        horas, minutos = divmod(round(segundos / 60), 60)
        return f"~{horas} h {minutos:02d} min"
    if segundos >= 90:
        return f"~{round(segundos / 60)} min"
    return f"~{max(5, round(segundos / 5) * 5)} s"


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

    def na_linha_da_grade(caixa):
        return any(abs((caixa[1] + caixa[3]) / 2 - (d[1] + d[3]) / 2) <= 14 for d in descricoes)

    # cabecalhos da grade (linha de baixo dos titulos em 2 linhas), como o Tesseract le em 4K:
    # "tem" (Item), "spm" (Oper Spm), "(ltd. fiscal" (Qtd.Fiscal), "unitario fiscal" (Valor Unitario Fiscal)
    grade = {
        "item": lambda t: t in ("item", "tem", "ltem", "1tem"),
        "oper": lambda t: t == "spm" or t.endswith(" spm"),
        "qtd": lambda t: "fiscal" in t and "td" in t and "unit" not in t,
        "unit": lambda t: "unitario" in t,
    }
    for campo, pred in grade.items():
        caixa = next((c for texto, c in lista if pred(texto) and na_linha_da_grade(c)), None)
        if caixa:
            achados[campo] = caixa
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
        if guardar is not None:
            guardar.extend(frases(palavras))
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


def foco_unido(imagem):
    """Campo em foco (ciano; sem ciano, o texto selecionado) inteiro: une os pedacos da mesma linha. Com texto selecionado no campo (ex.: a Data
    Emissao que o SRP preenche), o ciano se parte e o maior pedaco podia ser so a sobra da direita. Ignora blocos
    com menos de 8 px de altura (ha um traco ciano de 23x2 no alto da tela do Citrix)."""
    caixas = [c for c in core.caixas_cor(imagem, "ciano") if c[3] - c[1] >= 8]  # altura: pedacos estreitos valem
    if not caixas:
        # campo ja preenchido pelo SRP (Data Emissao): fundo branco e o texto todo selecionado em azul escuro, sem ciano.
        # Bordas de 1 px da janela tem o mesmo azul: so blocos de 8 px ou mais nos dois sentidos.
        # Resolucao inteira: em meia escala o texto branco da data partia o bloco azul e sobrava so uma faixa de 8 px
        # no alto do campo; o robo nao reconheceu a Data Emissao e passou dela (04/11/2026 18:51).
        # Com o fechamento do texto (`core.FECHAR`) bordas de 1-4 px ja somem: basta 4 px de altura (em 1366x768 o
        # campo tem ~14 px e a selecao sai com ~9; 05/11/2026).
        caixas = [c for c in core.caixas_cor(imagem, "selecao") if c[3] - c[1] >= 4 and c[2] - c[0] >= 8]
    if not caixas:
        return None
    esq, topo, dir_, base = max(caixas, key=lambda c: (c[2] - c[0]) * (c[3] - c[1]))
    # texto selecionado so na faixa da linha do campo (procurar na tela 4K inteira levava 0,27 s a cada consulta)
    faixa_topo = max(topo - 6, 0)
    faixa = imagem.crop((0, faixa_topo, imagem.width, min(base + 6, imagem.height)))
    caixas = caixas + [(c[0], c[1] + faixa_topo, c[2], c[3] + faixa_topo) for c in core.caixas_cor(faixa, "selecao")
                       if c[3] - c[1] >= 4 and c[2] - c[0] >= 8]
    folga = max(300 * imagem.width // 3840, 8 * (base - topo))  # proporcional a tela (1366x768: ~107 px)  # a selecao do texto da Data deixava ~170 px sem ciano (4K)
    mesma_linha = sorted(c for c in caixas if abs(c[1] - topo) <= 4 and abs(c[3] - base) <= 4)
    for c in mesma_linha:  # da esquerda para a direita, colando os pedacos vizinhos a direita
        if c[0] > dir_ and c[0] - dir_ <= folga:
            dir_ = c[2]
    for c in reversed(mesma_linha):  # e os da esquerda
        if c[2] < esq and esq - c[2] <= folga:
            esq = c[0]
    return (esq, topo, dir_, base)


def foco_tela(janela):
    return foco_unido(core.capturar(janela))


def ler_estado(janela, fixos=(), imagem=None):
    """(foco unido, ha aviso do SRP) numa captura so (`imagem`: captura ja feita)."""
    imagem = core.capturar(janela) if imagem is None else imagem
    return foco_unido(imagem), core.icone_aviso(imagem, fixos) is not None


# 1,2 s sem ciano = campo sem cor (lista Lemg...). Ja foi 0,8 s (04/11/2026: na rolagem da grade o robo achou que era
# campo sem cor e digitou o Valor Unit no Vlr ICMS) e 2 s; hoje o Valor Unit tem entrada propria (entrar_no_valor_unit)
# e a grade rolada e conferida pela faixa acima do foco. Cada campo sem cor custava 2 s parado.
SEM_COR = 1.2
# onde o foco ja ficou sem cor nesta execucao (mesmo campo, mesmo Enter): 0,4 s. Nos logs de 09/11/2026 sao sempre os
# mesmos 4 pontos por nota (Enter 3 e 5 depois da Data, Enter 8 e 9 depois do Valor do ICMS), ~1,8 s cada. O robo so
# aperta Enter nesses campos (nunca digita sem ver o campo); campo lento continua com o ciano velho e nao conta.
SEM_COR_CONHECIDO = 0.4


def aguardar(janela, referencia, limite, fixos=(), sem_cor=SEM_COR, antes=None):
    """Depois de um Enter: a nova caixa do foco; "aviso" se o aviso do SRP aparecer; "sem_cor" se nenhum campo ficar
    ciano por `sem_cor` s (listas como o Lemg so ganham borda pontilhada); None se o foco nao andar em `limite` s.
    Com `referencia` None (vindo de um campo sem cor), qualquer ciano que aparecer conta. Com `antes` (captura de antes
    do Enter): ciano no mesmo lugar tambem conta se a faixa acima dele mudou (na grade rolada, depois do Valor Unit, o
    campo em foco fica sempre na borda direita e quem anda e a grade; 04/11/2026)."""
    prazo, sem_desde, sem_cor = time.time() + limite, None, min(sem_cor, limite / 2)
    while time.time() < prazo:
        # na grade rolada a mesma captura serve para o foco e para a faixa acima dele (~0,16 s a menos por consulta)
        imagem = core.capturar(janela) if antes is not None else None
        atual, tem_aviso = ler_estado(janela, fixos, imagem)
        if atual and (referencia is None or abs(atual[0] - referencia[0]) > 5 or abs(atual[1] - referencia[1]) > 5):
            return atual
        if atual and antes is not None and mudou_acima(antes, imagem, referencia):
            time.sleep(0.3)  # a grade termina de rolar
            return ler_estado(janela, fixos)[0] or atual
        if tem_aviso:
            return "aviso"
        if atual is None:
            sem_desde = sem_desde or time.time()
            if time.time() - sem_desde >= sem_cor:
                return "sem_cor"
        else:
            sem_desde = None
        time.sleep(0.02)
    return None


def vizinho_esquerda(caixa, lista):
    """(texto, retangulo) da frase lida mais perto, a esquerda do campo, na mesma linha."""
    centro_y = (caixa[1] + caixa[3]) / 2
    esquerda = [(caixa[0] - c[2], t, c) for t, c in lista
                if abs((c[1] + c[3]) / 2 - centro_y) <= max(8, (caixa[3] - caixa[1]) / 2 + 4) and c[2] <= caixa[0] + 5]
    return min(esquerda)[1:] if esquerda else None


def ao_lado(caixa, rotulo, lista=None, distancia=60):
    """O campo em foco esta a direita do rotulo, na mesma linha? Com `lista` (frases lidas no inicio, formulario
    vazio): o rotulo tem que ser a frase mais proxima a esquerda do campo. Sempre: no maximo `distancia` px (ou 3
    alturas do campo) entre o rotulo e o campo; o foco_unido junta os pedacos, entao o campo comeca junto do rotulo."""
    centro, centro_rotulo = (caixa[1] + caixa[3]) / 2, (rotulo[1] + rotulo[3]) / 2
    if abs(centro - centro_rotulo) > max(8, (caixa[3] - caixa[1]) / 2 + 4) or caixa[0] < rotulo[2] - 5:
        return False
    if caixa[0] - rotulo[2] > max(distancia, 3 * (caixa[3] - caixa[1])):
        return False
    vizinho = vizinho_esquerda(caixa, lista) if lista else None
    if vizinho:
        c = vizinho[1]
        return abs((c[0] + c[2]) / 2 - (rotulo[0] + rotulo[2]) / 2) <= 15 and abs((c[1] + c[3]) / 2 - centro_rotulo) <= 15
    return True


def na_coluna(caixa, rotulo):
    """O campo em foco esta na coluna do rotulo (cabecalho da grade), abaixo dele?"""
    centro, centro_rotulo = (caixa[0] + caixa[2]) / 2, (rotulo[0] + rotulo[2]) / 2
    return abs(centro - centro_rotulo) <= max(40, caixa[2] - caixa[0]) and caixa[1] > rotulo[1]


def posicao_ok(campo, caixa, rotulos, anterior=None, lista=None):
    """True/False se da para conferir o campo pela posicao; None se o rotulo nao foi lido (nao confere).
    Foco sem cor (caixa None) nunca confere."""
    rotulo = "filial" if campo in ("filial1", "filial2") else campo
    if caixa is None:
        return False if rotulo in rotulos else None
    if campo == "filial2":
        rotulo = rotulos.get("filial")
        if not rotulo or not anterior:
            return None
        return abs((caixa[1] + caixa[3]) / 2 - (anterior[1] + anterior[3]) / 2) <= 6 and caixa[0] > anterior[2]
    if campo in ("item", "oper", "qtd"):  # colunas da grade visiveis no inicio (depois do Valor Unit a grade rola)
        return na_coluna(caixa, rotulos[campo]) if campo in rotulos else None
    if campo in ("unit", "icms", "icms_terceiro"):  # grade ja rolada: posicao do inicio nao vale (titulo_acima confere)
        return None
    chave = "filial" if campo == "filial1" else campo
    if chave in rotulos:
        return ao_lado(caixa, rotulos[chave], lista)
    return None


# texto esperado no titulo da coluna, acima do campo em foco, nas colunas que so aparecem depois que a grade rola
# (o Valor Unit fica amarelo, sem ciano: nao ha caixa para ler o titulo acima; ver entrar_no_valor_unit)
# em 4K o Tesseract leu "Vlr ICMS Item Doc Forn" como "vir oms tem doc forn" (04/11/2026)
TITULO_COLUNA = {"icms": ("icms", "cms", "oms", "doc forn", "forn")}


def titulo_acima(imagem, caixa, tesseract, topo_linha1=None):
    """Texto (normalizado) do titulo da coluna acima do campo em foco, lido na hora (a grade pode ter rolado). O titulo
    fica acima da 1a linha da grade: `topo_linha1` (topo do mesmo campo no 1o item da nota) estica o recorte ate ele
    nos itens de baixo; no 3o item de uma nota (06/11/2026, 1366x768) as 4 alturas acima do campo so pegavam os valores das linhas
    1 e 2 (')[ 5903.470000 )[ 1292.100000') e o robo pausou com o cursor certo."""
    altura = caixa[3] - caixa[1]
    topo = min(caixa[1], topo_linha1) if topo_linha1 is not None else caixa[1]
    area = (max(caixa[0] - 10, 0), max(topo - 4 * altura, 0), min(caixa[2] + 10, imagem.width), max(caixa[1] - 2, 1))
    if area[2] <= area[0] or area[3] <= area[1]:
        return ""
    recorte = imagem.crop(area)
    return normalizar(" ".join(p[0] for p in core.palavras_tela(recorte, tesseract)))


# campos depois dos quais a grade ja rolou: o foco fica na borda direita e a grade anda por baixo dele
GRADE_ROLADA = ("unit", "icms", "icms_terceiro")


def mudou_acima(antes, agora, caixa):
    """A faixa acima do campo (titulos e linhas de cima da grade) mudou entre as duas capturas? = a grade rolou."""
    altura = caixa[3] - caixa[1]
    area = (max(caixa[0] - 400, 0), max(caixa[1] - 6 * altura, 0), min(caixa[2] + 400, agora.width), max(caixa[1] - 2, 1))
    if area[2] <= area[0] or area[3] <= area[1] or antes.size != agora.size:
        return antes.size != agora.size
    from PIL import ImageChops
    diferenca = ImageChops.difference(antes.crop(area).convert("L"), agora.crop(area).convert("L"))
    return sum(diferenca.point(lambda v: 255 if v > 40 else 0).histogram()[255:]) > 50


# ---------- Rateio (janela "Rateios", abre com Enter depois do Valor do ICMS) ----------
# Grade "Rateio Item Documento Recebimento Tributo": as linhas COFINS e PIS ficam como o SRP preencheu; na linha
# ICMS TERCEIROS (combinado com o usuario em 04/11/2026):
#   sem imposto (ICMS TERCEIROS e ALIQUOTA da planilha = 0): Ident Tributo 0, Base Calculo 0, Aliquota 0
#   (o Lancamento nao recebe foco: o SRP calcula; no SRP real o Enter so para em Ident, Base Calculo e Aliquota)
#   com imposto (os dois diferentes de 0): Ident Tributo 1, Base Calculo = ICMS TERCEIROS (BASE DE CALCULO NF) e
#   Aliquota = ALIQUOTA (usuario, 05/11/2026: na planilha 153241.57 x 12% = 18388.99, o VALOR DO ICMS; o Lancamento o SRP
#   calcula). Ate 05/11 a Base vinha da coluna BASE CALC. ICMS e a Aliquota ficava como o SRP pos.
# Depois clica em "Tributos OK" (fecha o Rateio) e, no fim da nota, em Integrar.
# O usuario contou 7 Enter ate o Rateio abrir ("mas pode ser mais"); no SRP real (log de 04/11) o foco passou pela parte
# de cima do Rateio e chegou na grade de tributos no 10o Enter; cada linha de tributo teve 3 paradas.
MAX_ATE_RATEIO = 20  # Enter depois do Valor do ICMS ate o Ident Tributo da linha ICMS TERCEIROS
LER_RATEIO_DESDE = 8  # a partir deste Enter, le a tela procurando o Rateio (lido no 9o nos testes de 04/11; ~5 s cada)
COLUNAS_RATEIO = ("ident", "base", "aliquota", "lancamento")  # na ordem da tela
NOME_RATEIO = {"ident": "Ident Tributo", "base": "Base Calculo", "aliquota": "Aliquota", "lancamento": "Lancamento"}


def arquivo_posicoes():
    """Posicoes do Rateio lidas no OCR, guardadas entre execucoes (ao lado do exe; fora do git)."""
    return core.pasta_base() / "posicoes_0236.json"


def ler_posicoes(tamanho):
    """Posicoes do Rateio salvas para uma janela do Citrix deste tamanho (outro tamanho/monitor = None)."""
    try:
        dados = json.loads(arquivo_posicoes().read_text(encoding="utf-8"))
        if tuple(dados.get("janela", ())) == tuple(tamanho):
            return {chave: tuple(caixa) for chave, caixa in dados["rateio"].items()}
    except (OSError, ValueError, KeyError, TypeError, AttributeError):
        pass
    return None


def gravar_posicoes(tamanho, rateio):
    try:
        arquivo_posicoes().write_text(json.dumps({"janela": list(tamanho), "rateio": {c: list(v) for c, v in rateio.items()}}),
                                      encoding="utf-8")
    except OSError:
        pass


def tem_imposto(r):
    """Os dois valores da planilha diferentes de 0 = nota com imposto (vazio conta como 0)."""
    return Decimal(r.get("icms_terceiro") or "0") != 0 and Decimal(r.get("aliquota_icms") or "0") != 0


def valores_rateio(r):
    """O que digitar na linha ICMS TERCEIROS, por coluna (na ordem da tela)."""
    if tem_imposto(r):
        return {"ident": "1", "base": r["icms_terceiro"], "aliquota": r["aliquota_icms"]}
    return {"ident": "0", "base": "0", "aliquota": "0"}


def localizar_rateio(palavras):
    """Rotulos do Rateio lidos no OCR: linha ICMS TERCEIROS, colunas da grade de tributos e botao Tributos OK."""
    lista = frases(palavras)
    testes = {
        # em 1366x768 o OCR leu "11s icms erceiros" e "is terceiros" (05/11/2026)
        "terceiros": lambda t: "erceiro" in t or "icms tercei" in t,
        "ident": lambda t: t.startswith("ident"),
        "base": lambda t: t.startswith("base calc") or t == "base",
        "aliquota": lambda t: t.startswith("aliquota") and "zero" not in t,  # nao a linha "COFINS ALIQUOTA ZERO"
        "lancamento": lambda t: t.startswith(("lancamento", "langamento", "lancament")),
        "tributos_ok": lambda t: "tributos" in t,  # botao; "Nome Tributo"/"Ident Tributo" sao no singular
    }
    achados = {}
    for chave, teste in testes.items():
        caixa = next((c for texto, c in lista if teste(texto)), None)
        if caixa:
            achados[chave] = caixa
    if "ident" not in achados:
        # titulo em 2 linhas "Ident / Tributo": com o Citrix em 1366x768 o OCR nao leu o "Ident", mas leu o "Tributo"
        # sozinho (05/11/2026 14:09). O mais a esquerda: o "Clf Tribut" fica a direita da Descricao.
        sozinhos = sorted((c for texto, c in lista if texto == "tributo"), key=lambda c: c[0])
        if sozinhos:
            achados["ident"] = sozinhos[0]
    return achados


def ler_rateio(imagem, tesseract):
    """localizar_rateio na tela; leituras extras (binarizada: texto preto em botao cinza; em tela pequena, ampliada 4x)
    so para o que faltar. Em 1366x768 (05/11/2026) a leitura 2x nao achou Ident nem ICMS TERCEIROS e a 4x achou o Ident
    mas perdeu o ICMS TERCEIROS: junta as leituras."""
    escalas = (2,) if imagem.width >= 2000 else (2, 4)
    achados = {}
    for escala in escalas:
        # binarizada a 120: em 1366x768 a unica que leu a linha ICMS TERCEIROS em todas as ampliacoes
        for img in (imagem, binarizar(imagem)) + ((binarizar(imagem, 120),) if len(escalas) > 1 else ()):
            if all(c in achados for c in ("terceiros", "tributos_ok") + COLUNAS_RATEIO):
                return achados
            for chave, caixa in localizar_rateio(core.palavras_tela(img, tesseract, escala=escala)).items():
                achados.setdefault(chave, caixa)
    return achados


def coluna_rateio(caixa, rateio):
    """Coluna da grade de tributos (pelo titulo mais perto na horizontal) ou None se longe de todos."""
    if caixa is None:
        return None
    centro, largura = (caixa[0] + caixa[2]) / 2, caixa[2] - caixa[0]
    perto = [(abs(centro - (rateio[c][0] + rateio[c][2]) / 2), c) for c in COLUNAS_RATEIO if c in rateio]
    if not perto:
        return None
    distancia, coluna = min(perto)
    titulo = rateio[coluna]
    return coluna if distancia <= max(largura, titulo[2] - titulo[0]) * 0.75 else None


def na_linha(caixa, rotulo):
    """O campo em foco esta na mesma linha do rotulo (ex.: a linha ICMS TERCEIROS)?"""
    return caixa is not None and abs((caixa[1] + caixa[3]) / 2 - (rotulo[1] + rotulo[3]) / 2) <= max(8, (caixa[3] - caixa[1]) / 2 + 4)


def celula_ident(caixa, rateio):
    """Foco na celula Ident Tributo do Rateio: coluna Ident e estreito. Na grade principal, campo largo na mesma altura
    (ex.: a Descricao do 3o item em 1366x768, 207 px contra 27 px da celula Ident) nao conta (06/11/2026)."""
    return (caixa is not None and coluna_rateio(caixa, rateio) == "ident"
            and caixa[2] - caixa[0] <= 2 * (rateio["ident"][2] - rateio["ident"][0]))


# depois do 1o clique em Integrar: 7 Enter e Integrar de novo (usuario, 04/11/2026)
ENTERS_DEPOIS_INTEGRAR = 7
# aviso que pode aparecer depois do 2o Integrar; o robo fecha com Enter (usuario, 04/11/2026). Outros avisos: nao mexe.
# "Atencao: Este documento possui frete com ICMS PAGO no LTCC, por isso foi alterado o Frete para CIF. [OK]"
AVISO_FRETE_CIF = ("icms pago", "alterado o frete", "frete para cif")
LEITURAS_AVISO = 3  # leituras seguidas esperando uma janela ja tratada fechar
PIXELS_JANELA = 2000  # mudanca minima na tela para contar como janela abrindo/fechando (cursor piscando: ~60 px)
MAX_PASSOS_INTEGRAR = 30  # leituras da tela no maximo depois do 1o Integrar


def aviso_frete_cif(texto):
    """O texto (normalizado) lido na tela tem o aviso do frete alterado para CIF?"""
    return sum(t in texto for t in AVISO_FRETE_CIF) >= 2


def resumo_tributos(texto):
    """Janela "Resumo dos Tributos do Documento" (COFINS/PIS ... [Confirmar] [Cancelar]), que vem depois do aviso do frete;
    o robo confirma com Enter (usuario, 04/11/2026). Em 4K o titulo (branco no azul) sai pela metade: "ae dos tributos";
    no recorte da janela saiu "resume dos tributes ... conftirmar ... cancelar"."""
    return "dos tribut" in texto and any(t in texto for t in ("confirmar", "nfirmar", "cancelar"))


def janela_ltcc(texto):
    """Janela "Centro de Custo e Conta Contabil para Frete LTCC" (vem depois do Resumo dos Tributos); o robo clica no
    botao "Frete LTCC" (usuario, 04/11/2026)."""
    return "informe o centro de custo" in texto or ("centro de custo" in texto and "frete ltcc" in texto)


def botao_frete_ltcc(lista):
    """Botao "Frete LTCC" nas frases do OCR binarizado: na linha do "Voltar" e a direita dele (o titulo e os cabecalhos
    da grade tambem dizem "Frete LTCC"). Em 4K sai "frete ltc"."""
    voltar = next((c for t, c in lista if t.startswith("voltar")), None)
    if not voltar:
        return None
    altura = voltar[3] - voltar[1]
    return next((c for t, c in lista if t.startswith("frete ltc") and c[0] > voltar[2]
                 and abs((c[1] + c[3]) / 2 - (voltar[1] + voltar[3]) / 2) <= max(10, altura)), None)


def nf_referencia(texto):
    """Janela "Nota fiscal de referencia do faturamento" ("Escolha a Nota Fiscal Faturamento e Pressione o Botao OK"),
    que o SRP abre no caminho do Frete ao Item e as vezes depois do 1o Integrar (05/11/2026). O robo so da OK, sem ler
    a nota (usuario, 05/11/2026). Sem tratar, os 7 Enter atravessavam a grade dela
    com a tela inteira lida antes de cada um (~16 s, log de 05/11 16:30)."""
    return "referencia emitida" in texto or "nota fiscal faturamento" in texto or (
        "referencia" in texto and "faturamento" in texto)


def botao_ok_referencia(lista):
    """Botao "OK" da janela de NF de referencia nas frases do OCR binarizado: a frase "ok" mais baixa."""
    oks = [c for t, c in lista if t in ("ok", "0k", "ok.")]
    return max(oks, key=lambda c: c[1]) if oks else None


def botao_referencia(imagem, area, tesseract):
    """Botao OK da janela de NF de referencia no recorte `area` (OCR binarizado: texto preto no botao cinza)."""
    return botao_ok_referencia(frases(core.palavras_tela(binarizar(imagem.crop(area)), tesseract)))


def integrado_com_sucesso(texto):
    """Aviso final "Documento Integrado com Sucesso. [OK]": o robo da Enter."""
    return "integrado com sucesso" in texto


# janela "Atencao" do Integrar: "---- Valor Calculado para Documento ---- Total Geral...: X / Total Documento: Y /
# Total Acessoria..: Z / Botao <Ajustar> para integrar este valor / Botao <Cancelar> para usuario conferir"
# [Cancelar] [Ajustar]. Despesa acessoria de ate 0,05: o robo clica em Ajustar e segue (usuario, 06/11/2026); acima
# disso, ou sem ler os 3 valores conferindo entre si, nao clica em nada e pausa. O foco fica no Cancelar: nada de Enter.
LIMITE_ACESSORIA = 0.05
_NUMERO = r"(-?\d*[.,]\d{1,2}(?!\d)|-?\d+(?![\d.,]))"


def ajuste_acessoria(texto):
    """O texto (normalizado) lido tem a janela "Valor Calculado para Documento" (despesa acessoria)?"""
    return "valor calculado" in texto or ("acessoria" in texto and "ajustar" in texto)


def valores_acessoria(texto):
    """(total geral, total documento, total acessoria) lidos no texto normalizado, so se os 3 conferirem entre si
    (|documento - geral| = |acessoria|); senao None."""
    # o ultimo numero entre o rotulo e o rotulo seguinte: os pontilhados ("Total Geral.......:") saem como "2" no OCR
    # do print real (06/11/2026): "total geral 2 236367.42"
    valores = []
    for rotulo, fim in (("total geral", "total documento"), ("total documento", "acessoria"), ("acessoria", r"bot|<|ajust")):
        achado = re.search(re.escape(rotulo) + r"(.*?)(?:" + fim + "|$)", texto)
        numeros = re.findall(_NUMERO, achado.group(1)) if achado else []
        if not numeros:
            return None
        valores.append(float(numeros[-1].replace(",", ".")))
    geral, documento, acessoria = valores
    return tuple(valores) if abs(abs(documento - geral) - abs(acessoria)) < 0.005 else None


def botao_ajustar(lista):
    """Botao "Ajustar" nas frases do OCR: a frase que e so "ajustar" (o texto "Botao <Ajustar> para integrar..." fica
    numa frase maior), a mais baixa."""
    botoes = [c for t, c in lista if t.strip("<>[]|(). ") in ("ajustar", "ajusta")]
    return max(botoes, key=lambda c: c[1]) if botoes else None


def barras_dialogo(imagem):
    """Barras de titulo (azul `titulo_srp`) das janelas que o SRP abre por cima da tela principal; a da tela principal
    (quase a largura toda) fica de fora. ~0,07 s na tela 4K, contra ~2,2 s do OCR da tela inteira."""
    minimo = max(imagem.height // 100, 12)  # 4K: barra de 56 px
    return [c for c in core.caixas_cor(imagem, "titulo_srp")
            if c[3] - c[1] >= minimo and 250 <= c[2] - c[0] < 0.9 * imagem.width]


def area_dialogo(imagem, barras):
    """Recorte que cobre as janelas das `barras`. Nos prints de 04/11/2026 (aviso do frete, Resumo, Frete LTCC,
    Integrado com Sucesso) cada janela tem altura de 0,46 a 0,6 da largura da barra: o recorte vai ate 0,8."""
    base = max(b[1] + max(0.8 * (b[2] - b[0]), 15 * (b[3] - b[1])) for b in barras)
    return (max(min(b[0] for b in barras) - 20, 0), max(min(b[1] for b in barras) - 5, 0),
            min(max(b[2] for b in barras) + 20, imagem.width), min(int(base), imagem.height))


def documento_ja_existe(texto):
    """Janela "Atencao: Documento Recebimento 184598 ja existe, foi integrado no sistema em 15/10/2010 [OK]", que o SRP
    abre no Enter depois da Chave Acesso NF-e (o cursor fica parado; 09/11/2026). O robo da OK, apaga a Chave e poe
    "." no fim da Serie ("1" -> "1."), so quando essa janela aparece (usuario, 09/11/2026)."""
    return "ja existe" in texto or "integrado no sistema" in texto


def janelas_no_texto(texto):
    """Janelas pos-Integrar reconhecidas no texto (normalizado) lido."""
    return [nome for nome, teste in (("referencia", nf_referencia), ("acessoria", ajuste_acessoria),
                                     ("frete", aviso_frete_cif), ("resumo", resumo_tributos),
                                     ("ltcc", janela_ltcc), ("sucesso", integrado_com_sucesso)) if teste(texto)]


def tela_mudou(antes, agora, minimo=1):
    """Mais de `minimo` pixels diferentes entre as capturas (use um minimo alto para nao contar o cursor piscando)."""
    if antes.size != agora.size:
        return True
    from PIL import ImageChops
    diferenca = ImageChops.difference(antes.convert("L"), agora.convert("L")).point(lambda v: 255 if v > 40 else 0)
    return diferenca.histogram()[255] >= minimo


def clicar(x, y):
    """Clique na tela (coordenadas fisicas)."""
    from pywinauto import mouse
    mouse.click(button="left", coords=(int(x), int(y)))


class Travou(Exception):
    def __init__(self, campo, motivo):
        super().__init__(motivo)
        self.campo, self.motivo = campo, motivo


class DocumentoExiste(Travou):
    """Janela "Documento Recebimento ... ja existe" no Enter depois da Chave (motivo = texto lido da janela)."""


# log da execucao atual (ou da ultima): `registrar_no_log` escreve nele de fora do robo, ex.: a pintura da planilha
# (usuario, 09/11/2026: as mensagens da pintura so iam para o console e nao deu para saber por que 2 notas lancadas
# ficaram sem verde)
_LOG = {"caminho": None, "arquivo": None, "ultimo": 0.0}
_TRAVA_LOG = threading.Lock()


def _escrever_log(texto):
    """Linha no log com hora e o tempo desde a linha anterior (+s). Chamar com `_TRAVA_LOG`."""
    arquivo = _LOG["arquivo"]
    fechar = arquivo is None or arquivo.closed
    if fechar:  # a execucao ja terminou: acrescenta no fim do log dela
        if not _LOG["caminho"]:
            return
        try:
            arquivo = open(_LOG["caminho"], "a", encoding="utf-8")
        except OSError:
            return
    agora = time.time()
    arquivo.write(f"{datetime.now():%H:%M:%S} +{agora - _LOG['ultimo']:5.1f}s {texto}\n")
    _LOG["ultimo"] = agora
    arquivo.flush()
    if fechar:
        arquivo.close()


def registrar_no_log(texto):
    """Linha no log da execucao atual (ou da ultima, se ja terminou); sem execucao nenhuma ainda, nada."""
    with _TRAVA_LOG:
        _escrever_log(texto)


_OK_NOTA = re.compile(r"OK Nota \.\.\.(\d{6}) \(linhas? [\d-]+\): (.*)")
_JA_EXISTE = re.compile(r"Nota \.\.\.(\d{6}): o SRP disse que o Documento Recebimento ja existe")


def notas_nos_logs(pasta):
    """O que os logs da SNAC0236 em `pasta` dizem de cada nota, pelos 6 ultimos digitos da chave (os logs so guardam
    esses): {sufixo: {"sucesso", "a_mao", "sem_confirmacao", "serie_ponto": bool, "quando": 'dd/mm hh:mm'}}.
    sucesso = o robo viu "Integrado com Sucesso"; a_mao = o usuario terminou a nota numa pausa; sem_confirmacao = o
    robo integrou mas nao viu a mensagem de sucesso (pausou para conferir). Serve para conferir a planilha (usuario,
    09/11/2026). Os logs ficam 30 dias."""
    notas = {}
    for arquivo in sorted(pasta.glob("*-snac0236.log")):
        nome = arquivo.name
        quando = f"{nome[6:8]}/{nome[4:6]} {nome[9:11]}:{nome[11:13]}"
        try:
            texto = arquivo.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        for linha in texto.splitlines():
            achado = _OK_NOTA.search(linha)
            if achado:
                nota = notas.setdefault(achado.group(1), {"sucesso": False, "a_mao": False, "sem_confirmacao": False,
                                                          "serie_ponto": False, "quando": quando})
                resto = achado.group(2)
                chave = ("sucesso" if "integrada com sucesso" in resto else "a_mao" if "terminada a mao" in resto
                         else "sem_confirmacao")
                nota[chave], nota["quando"] = True, quando
                continue
            achado = _JA_EXISTE.search(linha)
            if achado:
                notas.setdefault(achado.group(1), {"sucesso": False, "a_mao": False, "sem_confirmacao": False,
                                                   "serie_ponto": False, "quando": quando})["serie_ponto"] = True
    return notas


def conferir_com_logs(notas, logs):
    """Compara as notas da planilha (com r["verde"] de cada linha) com `notas_nos_logs`. Devolve listas de notas:
    'nao_pintadas' (o robo viu o sucesso, mas a nota nao esta toda verde), 'sem_confirmacao' (verde, mas nos logs nao
    ha o sucesso: terminada a mao ou sem a mensagem), 'fora_dos_logs' (verde, sem nada nos logs: pintada a mao ou
    ha mais de 30 dias) e 'serie_ponto' (o robo pos "." na Serie)."""
    saida = {"nao_pintadas": [], "sem_confirmacao": [], "fora_dos_logs": [], "serie_ponto": []}
    for nota in notas:
        info = logs.get(nota["chave"][-6:])
        verde = all(r.get("verde") for r in nota["itens"])
        if info and info["serie_ponto"]:
            saida["serie_ponto"].append(nota)
        if info and info["sucesso"] and not verde:
            saida["nao_pintadas"].append(nota)
        elif verde and info and not info["sucesso"]:
            saida["sem_confirmacao"].append(nota)
        elif verde and not info:
            saida["fora_dos_logs"].append(nota)
    return saida


def executar(notas, janela_trecho, pausa, render, espera, parar, aviso, estado=None, retomar=None, intervalo=0,
             serie_ponto=None):
    """Preenche as notas no SRP. `serie_ponto(linhas)`: chamado antes de marcar 'ok' a nota em que o robo pos "." na
    Serie (janela "ja existe"). `intervalo` (s): espera extra antes de cada acao (digitar, Enter, Integrar) e narra a
    acao no console, para o usuario acompanhar o percurso nos testes (0 = velocidade normal). `estado(linha, nome)` por linha da planilha; `retomar` (Event, botao Continuar) liga
    as pausas: depois de cada Integrar e quando algo nao confere. Sem ele, qualquer problema interrompe. Devolve 0 se concluiu."""
    core.limpar_logs()
    caminho_log = core.pasta_logs() / f"{datetime.now():%Y%m%d-%H%M%S}-snac0236.log"
    log = open(caminho_log, "w", encoding="utf-8")
    with _TRAVA_LOG:
        _LOG.update(caminho=caminho_log, arquivo=log, ultimo=time.time())
    rateio_salvo = [None]
    confia_barra = [True]  # a barra azul achou todas as janelas ate agora (senao le sempre a tela inteira)
    atual = []  # linhas marcadas 'executando'
    serie_com_ponto_feita = [False]  # nesta nota o robo ja pos "." na Serie (janela "ja existe")
    sem_cor_visto = set()  # (campo, Enter) em que o foco ficou sem cor nesta execucao: ali a espera e curta

    def registrar(texto, mostrar=True):
        """Linha no log com hora e o tempo desde a linha anterior (+s), para achar o que demora."""
        with _TRAVA_LOG:
            if not log.closed:
                _escrever_log(texto)
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
        registrar(f"  passo a passo: {texto}", mostrar=False)
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

    def limpar_campo(campo, quantos=10):
        """Apaga o conteudo do campo em foco: End + Backspace (serve com o texto selecionado ou com o cursor no meio)."""
        checar_parada()
        devagar(f"apagar o que houver em {NOME[campo]}")
        core.send_keys(f"{{END}}{{BACKSPACE {quantos}}}", pause=0.03)

    def enters(campo, quantos, foco):
        """Enter `quantos` vezes; cada um precisa mover o foco ou, na grade rolada, rolar a grade (senao: aviso do SRP
        ou campo preso)."""
        for vez in range(quantos):
            checar_parada()
            devagar(f"Enter {vez + 1} de {quantos} depois do {NOME[campo]}")
            antes = core.capturar(janela) if campo in GRADE_ROLADA and foco is not None else None
            core.send_keys("{ENTER}", pause=0.03)
            foco = depois_do_enter(campo, vez, foco, antes)
        return foco

    def depois_do_enter(campo, vez, foco, antes=None):
        """Nova caixa do foco (None = campo sem cor, ex.: lista Lemg); aviso ou foco parado = Travou. Onde o foco ja
        ficou sem cor nesta execucao (mesmo campo e mesmo Enter), espera so SEM_COR_CONHECIDO s sem ciano."""
        curto = (campo, vez) in sem_cor_visto
        novo = aguardar(janela, foco, render, fixos, sem_cor=SEM_COR_CONHECIDO if curto else SEM_COR, antes=antes)
        if novo in (None, "aviso") and campo in ("chave", "serie"):
            texto = janela_ja_existe()
            if texto and campo == "chave" and not serie_com_ponto_feita[0]:
                raise DocumentoExiste(campo, texto)
            if texto:
                raise Travou(campo, f"o SRP diz que o Documento Recebimento ja existe mesmo com '.' na Serie ('{texto[:90]}')")
        if novo is None:
            # 09/11/2026 13:04: o Enter depois do Valor do ICMS ficou 4,5 s parado (nas outras notas, ate 0,7 s) e o robo
            # pausou sem saber por que. Antes de pausar espera mais uma vez (SRP lento), sem apertar nada.
            novo = aguardar(janela, foco, render, fixos, antes=antes)
            if novo is not None:
                registrar(f"  Enter {vez + 1} depois do {NOME[campo]}: o SRP demorou mais de {render:.0f} s", mostrar=False)
        if novo in (None, "aviso"):
            motivo = "apareceu um aviso do SRP" if novo == "aviso" else "o cursor nao saiu do lugar"
            raise Travou(campo, f"{motivo} no Enter {vez + 1} depois do {NOME[campo]}{diagnostico(campo)}")
        if novo == "sem_cor" and not curto:
            sem_cor_visto.add((campo, vez))
        novo = None if novo == "sem_cor" else novo
        registrar(f"  Enter {vez + 1} depois do {NOME[campo]}: foco={novo} perto de {onde_esta(novo, lidas)!r}",
                  mostrar=False)
        return novo

    def diagnostico(campo):
        """Quando o cursor prende: texto da janela do SRP por cima (barra azul), se houver, e um print da janela do
        Citrix em logs\\telas (fica so neste computador; apagado com os logs, em 30 dias). Texto para a mensagem."""
        try:
            imagem = core.capturar(janela)
        except Exception:
            return ""
        nome = f"{datetime.now():%Y%m%d-%H%M%S}-{campo}.png"
        try:
            pasta = core.pasta_logs() / "telas"
            pasta.mkdir(exist_ok=True)
            imagem.save(pasta / nome)
            registrar(f"  print da tela: logs\\telas\\{nome}", mostrar=False)
        except Exception as erro:
            registrar(f"  print da tela falhou: {erro}", mostrar=False)
        texto = ""
        if tesseract:
            try:
                barras = barras_dialogo(imagem)
                if barras:
                    texto = normalizar(" ".join(p[0] for p in core.palavras_tela(
                        imagem.crop(area_dialogo(imagem, barras)), tesseract, escala=4 if imagem.width < 2000 else 2)))
            except Exception as erro:
                registrar(f"  leitura da janela por cima falhou: {erro}", mostrar=False)
        registrar(f"  janela por cima: {texto[:200]!r}" if texto else "  nenhuma janela do SRP por cima", mostrar=False)
        return f" (janela do SRP na tela: '{texto[:100]}')" if texto else ""

    def local(caixa):
        if caixa is None:
            return " (o cursor esta num campo sem cor, como uma lista)"
        texto = onde_esta(caixa, lidas)
        return f" (o cursor esta em '{texto}')" if texto else ""

    def ir_ate(campo, alvo, quantos, foco, anterior):
        """Depois de digitar `campo`, Enter ate o foco chegar em `alvo` (conferido pelo rotulo lido no inicio), no maximo
        `quantos` + 4 vezes: o SRP pula ou nao campos que ele mesmo preenche (CNPJ, Filial Fiscal...). Sem rotulo para
        conferir, so conta os `quantos` Enter combinados."""
        if posicao_ok(alvo, foco, rotulos, anterior, lidas) is None:
            return enters(campo, quantos, foco)
        limite = quantos + 4
        for vez in range(limite):
            checar_parada()
            devagar(f"Enter {vez + 1} depois do {NOME[campo]} (procurando {NOME[alvo]}; previstos {quantos})")
            core.send_keys("{ENTER}", pause=0.03)
            foco = depois_do_enter(campo, vez, foco)
            if posicao_ok(alvo, foco, rotulos, anterior, lidas):
                if vez + 1 != quantos:
                    registrar(f"  {NOME[alvo]}: cheguei com {vez + 1} Enter (previstos {quantos})", mostrar=False)
                return foco
        raise Travou(alvo, f"nao cheguei no campo {NOME[alvo]} com {limite} Enter depois do {NOME[campo]}{local(foco)}")

    def fechar_referencia(imagem, area):
        """Janela "Nota fiscal de referencia do faturamento": so OK, sem ler a nota (usuario, 05/11/2026). Clique no
        botao OK; sem o botao lido, Enter ate a janela fechar (o Enter atravessa a grade e chega no OK), no maximo 8."""
        botao = botao_referencia(imagem, area, tesseract)
        if botao:
            devagar("clicar em OK na 'Nota Fiscal de Referencia'")
            ret = janela.rectangle()
            clicar(ret.left + area[0] + (botao[0] + botao[2]) / 2, ret.top + area[1] + (botao[1] + botao[3]) / 2)
            registrar("  NF de referencia: cliquei em OK", mostrar=False)
            return
        registrar("  NF de referencia: nao li o botao OK; Enter ate a janela fechar", mostrar=False)
        for vez in range(8):
            checar_parada()
            devagar(f"Enter {vez + 1} na 'Nota Fiscal de Referencia' (ate o OK)")
            antes = core.capturar(janela)
            core.send_keys("{ENTER}", pause=0.03)
            esperar_tela(antes, 1.0)
            if not barras_dialogo(core.capturar(janela)):
                registrar(f"  NF de referencia: fechou com {vez + 1} Enter", mostrar=False)
                return
        raise Travou("item", "a janela 'Nota Fiscal de Referencia' nao fechou (nao li o botao OK e 8 Enter nao bastaram)")

    def janela_ja_existe():
        """Texto (normalizado) da janela "Documento Recebimento ... ja existe" se ela estiver aberta; senao None. So e
        chamada quando o cursor nao anda depois da Chave ou da Serie (le o recorte da barra azul; sem barra, a tela)."""
        if not tesseract:
            return None
        imagem = core.capturar(janela)
        barras = barras_dialogo(imagem)
        alvo = imagem.crop(area_dialogo(imagem, barras)) if barras else imagem
        for escala in ((2, 4) if barras else (2,)):  # em 1366x768 o texto da janela e pequeno
            texto = normalizar(" ".join(p[0] for p in core.palavras_tela(alvo, tesseract, escala=escala)))
            if documento_ja_existe(texto):
                return texto
        return None

    def serie_com_ponto(nota, foco, texto):
        """Janela "Documento Recebimento ... ja existe, foi integrado no sistema em ..." depois da Chave (usuario,
        09/11/2026): Enter (OK), apaga a Chave, Enter ate a Serie e poe "." no fim dela ("1" -> "1."); depois segue
        com Enter ate a Filial Fisica, como sempre. Devolve o foco na Filial Fisica."""
        registrar(f"  janela 'ja existe': {texto[:120]!r}", mostrar=False)
        checar_parada()
        devagar("Enter no aviso 'Documento Recebimento ja existe' (OK)")
        antes = core.capturar(janela)
        core.send_keys("{ENTER}", pause=0.03)
        esperar_tela(antes, render)
        if janela_ja_existe():
            raise Travou("chave", "a janela 'Documento Recebimento ja existe' nao fechou com Enter")
        foco = foco_tela(janela) or foco
        if posicao_ok("chave", foco, rotulos, None, lidas) is False:
            raise Travou("chave", f"depois do OK na janela 'ja existe' o cursor nao voltou para a Chave{local(foco)}")
        limpar_campo("chave", 60)
        serie_com_ponto_feita[0] = True
        rotulo = next((c for t, c in lidas if t.startswith("serie")), None)
        if not rotulo:
            raise Travou("serie", "o SRP disse que o Documento Recebimento ja existe, mas nao li o rotulo Serie na tela")
        limite = ENTERS_CHAVE_SERIE + 4
        for vez in range(limite):
            checar_parada()
            devagar(f"Enter {vez + 1} depois de apagar a Chave (procurando a Serie)")
            core.send_keys("{ENTER}", pause=0.03)
            foco = depois_do_enter("chave", vez, foco)
            if foco and ao_lado(foco, rotulo, lidas):
                break
        else:
            raise Travou("serie", f"nao cheguei na Serie com {limite} Enter depois de apagar a Chave{local(foco)}")
        checar_parada()
        devagar("ir para o fim da Serie e digitar '.'")
        core.send_keys("{END}.", pause=0.03)
        time.sleep(min(pausa, 0.05))
        registrar(f"Nota ...{nota['chave'][-6:]}: o SRP disse que o Documento Recebimento ja existe; apaguei a Chave e "
                  "pus '.' no fim da Serie.")
        return ir_ate("serie", "filial1", 2, foco, anterior=foco)

    def janela_no_caminho(foco):
        """Depois de cada Enter do Frete ao Item: janela por cima (barra de titulo azul, ~0,07 s)? A "Nota fiscal de
        referencia do faturamento" abre nesse caminho (usuario, 05/11/2026): `fechar_referencia` e segue. Outra janela
        qualquer: pausa (nao da Enter as cegas)."""
        if not tesseract:
            return foco
        imagem = core.capturar(janela)
        barras = barras_dialogo(imagem)
        if not barras:
            return foco
        area = area_dialogo(imagem, barras)
        texto = normalizar(" ".join(p[0] for p in core.palavras_tela(imagem.crop(area), tesseract)))
        if not nf_referencia(texto):
            raise Travou("item", f"apareceu uma janela do SRP entre o Frete e o Item ('{texto[:60]}')")
        fechar_referencia(imagem, area)
        novo = aguardar(janela, foco, render, fixos)
        if novo in (None, "aviso"):
            raise Travou("item", "depois do OK na 'Nota Fiscal de Referencia' "
                         + ("apareceu um aviso do SRP" if novo == "aviso" else "a tela nao mudou"))
        registrar(f"  NF de referencia fechada entre o Frete e o Item: foco={novo}", mostrar=False)
        return novo if isinstance(novo, tuple) else None

    def ir_para_item(foco):
        """Depois do Frete: Enter ate o Item da 1a linha da grade, no maximo ENTERS_FRETE_ITEM + 4 (usuario, 05/11/2026).
        De 04/11 a 05/11 o robo clicava no Item, mas o clique pulava os campos em que o SRP abre a janela "Nota fiscal de
        referencia do faturamento", que entao so aparecia depois (raramente). Depois de cada Enter, `janela_no_caminho`.
        Sem o rotulo Item lido: so a contagem."""
        conferir = posicao_ok("item", foco, rotulos, foco, lidas) is not None
        limite = ENTERS_FRETE_ITEM + 4 if conferir else ENTERS_FRETE_ITEM
        for vez in range(limite):
            checar_parada()
            devagar(f"Enter {vez + 1} depois do Frete (procurando o Item da grade)")
            core.send_keys("{ENTER}", pause=0.03)
            foco = janela_no_caminho(depois_do_enter("frete", vez, foco))
            if conferir and posicao_ok("item", foco, rotulos, None, lidas):
                registrar(f"  Item: cheguei com {vez + 1} Enter depois do Frete", mostrar=False)
                return foco
        if not conferir:
            return foco
        raise Travou("item", f"nao cheguei no Item da grade com {limite} Enter depois do Frete{local(foco)}")

    def descer_para_item(foco, anterior):
        """Item seguinte da mesma nota. Depois do Tributos OK o SRP deixa o foco no Item da MESMA linha (05/11/2026 15:20:
        o robo digitou o 2o item por cima do 1o): seta para baixo leva ao Item da linha de baixo (sugestao do usuario).
        `anterior` = caixa do Item da linha do item anterior."""
        meio = lambda c: (c[1] + c[3]) / 2  # noqa: E731
        if (isinstance(foco, tuple) and posicao_ok("item", foco, rotulos, None, lidas) is not False
                and meio(foco) > anterior[3]):
            registrar(f"  depois do Tributos OK o foco ja esta no Item da linha de baixo: {foco}", mostrar=False)
            return foco
        if not (isinstance(foco, tuple) and posicao_ok("item", foco, rotulos, None, lidas) is not False
                and anterior[1] <= meio(foco) <= anterior[3]):
            raise Travou("item", f"depois do Tributos OK o cursor nao ficou no Item da linha do item anterior{local(foco)}; "
                                 "nao sei em que linha digitar o proximo item")
        checar_parada()
        devagar("seta para baixo: Item da proxima linha da grade")
        core.send_keys("{DOWN}", pause=0.03)
        novo = aguardar(janela, foco, render, fixos)
        if novo == "aviso":
            raise Travou("item", "apareceu um aviso do SRP depois da seta para baixo na grade")
        if (isinstance(novo, tuple) and posicao_ok("item", novo, rotulos, None, lidas) is not False
                and meio(novo) > anterior[3]):
            registrar(f"  seta para baixo: Item da linha de baixo, foco={novo}", mostrar=False)
            return novo
        raise Travou("item", f"a seta para baixo nao levou ao Item da linha de baixo{local(novo if isinstance(novo, tuple) else None)}")

    def entrar_no_valor_unit(foco):
        """2o Enter depois do Qtd Fiscal: sai do Qtd.Fisica e entra no Valor Unit, que fica AMARELO (sem ciano) e faz a
        grade rolar para a direita. Conta como entrou se o ciano andar ou sumir, ou se a faixa acima do campo mudar
        (rolagem); nada disso em `render` s = Travou. Devolve o ciano que sobrar na tela (ou None)."""
        checar_parada()
        devagar("Enter 2 de 2 depois do Qtd Fiscal (entra no Valor Unit, que fica amarelo)")
        antes = core.capturar(janela)
        core.send_keys("{ENTER}", pause=0.03)
        prazo, sumiu = time.time() + render, None
        while time.time() < prazo:
            atual, tem_aviso = ler_estado(janela, fixos)
            if tem_aviso:
                raise Travou("unit", "apareceu um aviso do SRP no Enter 2 depois do Qtd Fiscal")
            andou = atual is not None and foco is not None and (abs(atual[0] - foco[0]) > 5 or abs(atual[1] - foco[1]) > 5)
            sumiu = (sumiu or time.time()) if atual is None else None
            rolou = foco is not None and mudou_acima(antes, core.capturar(janela), foco)
            if andou or rolou or (sumiu and time.time() - sumiu >= min(0.3, render / 2)):
                time.sleep(0.3)  # a grade termina de rolar
                if andou or rolou:  # sem ciano (o normal: Valor Unit amarelo) nao precisa reler
                    atual, _ = ler_estado(janela, fixos)
                registrar(f"  Enter 2 depois do Qtd Fiscal: no Valor Unit (andou={andou}, rolou={rolou}, "
                          f"sem ciano={bool(sumiu)}); ciano agora={atual}", mostrar=False)
                return atual
            time.sleep(0.05)
        raise Travou("unit", "o cursor nao saiu do Qtd.Fisica no Enter 2 depois do Qtd Fiscal (a grade nao rolou)")

    def enter_rateio(campo, vez, foco):
        checar_parada()
        antes = core.capturar(janela) if foco is not None else None
        core.send_keys("{ENTER}", pause=0.03)
        return depois_do_enter(campo, vez, foco, antes)

    def achar_botao(chave, antes):
        """Rele o botao (`chave` de localizar_rateio) na faixa em volta de onde estava; None se nao estiver a vista."""
        tela_agora = core.capturar(janela)
        esq, topo = max(antes[0] - 300, 0), max(antes[1] - 80, 0)
        recorte = tela_agora.crop((esq, topo, min(antes[2] + 300, tela_agora.width), min(antes[3] + 80, tela_agora.height)))
        caixa = localizar_rateio(core.palavras_tela(binarizar(recorte), tesseract)).get(chave)
        return (caixa[0] + esq, caixa[1] + topo, caixa[2] + esq, caixa[3] + topo) if caixa else None

    def confirma_terceiros(rateio):
        """Le so o recorte em volta do rotulo ICMS TERCEIROS (~0,5 s, contra ~5 s da tela inteira)."""
        c = rateio["terceiros"]
        h = c[3] - c[1]
        imagem = core.capturar(janela)
        recorte = imagem.crop((max(c[0] - 2 * h, 0), max(c[1] - h, 0), min(c[2] + 2 * h, imagem.width),
                               min(c[3] + h, imagem.height)))
        escala = 2 if h >= 20 else 4  # em 1366x768 o rotulo tem ~8 px
        for img in (recorte, binarizar(recorte)) + ((binarizar(recorte, 120),) if escala > 2 else ()):
            texto = normalizar(" ".join(p[0] for p in core.palavras_tela(img, tesseract, escala=escala)))
            if "tercei" in texto or "erceiro" in texto:
                return True
        registrar(f"  recorte do ICMS TERCEIROS nao confere: {texto!r}", mostrar=False)
        return False

    def clicar_ident_terceiros(pos, foco, vez):
        """Com o Rateio aberto e o foco no Ident Tributo de uma linha acima (199/198): clique no Ident Tributo da linha
        ICMS TERCEIROS, em vez dos ~6 Enter pelas linhas de cima (usuario, 04/11/2026). Foco que nao anda: segue com
        Enter, como antes."""
        t = pos["terceiros"]
        x, y = (foco[0] + foco[2]) / 2, (t[1] + t[3]) / 2
        ret = janela.rectangle()
        checar_parada()
        devagar("clicar no Ident Tributo da linha ICMS TERCEIROS")
        clicar(ret.left + x, ret.top + y)
        novo = aguardar(janela, foco, render, fixos)
        if novo == "aviso":
            raise Travou("rateio", "apareceu um aviso do SRP depois do clique no Ident Tributo do ICMS TERCEIROS")
        if isinstance(novo, tuple):
            registrar(f"  cliquei no Ident Tributo do ICMS TERCEIROS em ({x:.0f}, {y:.0f}) no Enter {vez + 1}: foco={novo}",
                      mostrar=False)
            return novo
        registrar(f"  cliquei no Ident Tributo do ICMS TERCEIROS, mas o foco nao andou ({novo}): sigo com Enter", mostrar=False)
        return None if novo == "sem_cor" else foco

    def preencher_rateio(r, foco):
        """Depois do Valor do ICMS: Enter ate o Ident Tributo da linha ICMS TERCEIROS (lida no OCR quando o Rateio abre),
        digita os valores da linha (`valores_rateio`), clica em Tributos OK e espera o Rateio fechar."""
        if not tesseract:
            raise Travou("rateio", "sem o Tesseract nao consigo achar a linha ICMS TERCEIROS no Rateio")
        rateio = None
        salvo = rateio_salvo[0]  # posicoes lidas numa nota anterior (a janela do Rateio abre sempre no mesmo lugar)
        clicou = False
        for vez in range(MAX_ATE_RATEIO):
            devagar(f"Enter {vez + 1} depois do Valor do ICMS (procurando a linha ICMS TERCEIROS do Rateio)")
            foco = enter_rateio("icms", vez, foco)
            pos = rateio or salvo
            # foco estreito na coluna Ident, abaixo do titulo e acima da linha ICMS TERCEIROS = Rateio aberto numa linha
            # de cima (na grade principal o campo nessa altura e largo: 635 px contra 107 px da celula Ident, 4K)
            if (pos and not clicou and celula_ident(foco, pos)
                    and foco[1] > pos["ident"][3] and foco[3] < pos["terceiros"][1]):
                clicou = True
                foco = clicar_ident_terceiros(pos, foco, vez)
            if rateio is None and salvo and na_linha(foco, salvo["terceiros"]) and celula_ident(foco, salvo):
                if confirma_terceiros(salvo):
                    rateio = salvo
                    registrar(f"  Rateio: posicoes da nota anterior; ICMS TERCEIROS conferido no recorte (Enter {vez + 1})",
                              mostrar=False)
                else:
                    salvo = None  # le a tela inteira de novo
            if salvo and rateio is None and vez + 1 >= LER_RATEIO_DESDE + 10:
                salvo = None  # nao bateu com as posicoes salvas: le a tela inteira
            if rateio is None and not salvo and vez + 1 >= LER_RATEIO_DESDE:
                t_leitura = time.time()
                lido = ler_rateio(core.capturar(janela), tesseract)
                registrar(f"  leitura da tela procurando o Rateio: {time.time() - t_leitura:.1f}s", mostrar=False)
                if "terceiros" in lido and "tributos_ok" in lido:
                    rateio = lido
                    registrar(f"  Rateio lido no Enter {vez + 1}: {rateio}", mostrar=False)
                    faltam = [NOME_RATEIO[c] for c in COLUNAS_RATEIO if c not in rateio and c in valores_rateio(r) or
                              c == "ident" and c not in rateio]
                    if faltam:
                        raise Travou("rateio", f"nao li no Rateio a(s) coluna(s) {', '.join(faltam)}")
                    rateio_salvo[0] = rateio
                    gravar_posicoes(tamanho, rateio)
            if rateio and na_linha(foco, rateio["terceiros"]) and celula_ident(foco, rateio):
                registrar(f"  ICMS TERCEIROS: cheguei no Ident Tributo com {vez + 1} Enter depois do Valor do ICMS", mostrar=False)
                break
            linha = (rateio or salvo or {}).get("terceiros")
            if linha and foco is not None and foco[1] > linha[3] + (foco[3] - foco[1]):
                raise Travou("rateio", "passei da linha ICMS TERCEIROS no Rateio sem parar no Ident Tributo")
        else:
            raise Travou("rateio", f"nao cheguei no Ident Tributo da linha ICMS TERCEIROS com {MAX_ATE_RATEIO} Enter depois "
                         "do Valor do ICMS" + ("" if rateio else " (nao achei o Rateio na tela)"))
        valores = valores_rateio(r)
        registrar(f"  linha {r['linha']}: {'com' if tem_imposto(r) else 'sem'} imposto: {valores}", mostrar=False)
        pendentes = [c for c in COLUNAS_RATEIO if c in valores]
        # O Ident do ICMS TERCEIROS ja vem com 1 no SRP (05/11/2026). Chegando pelo clique, o cursor fica onde clicou e o
        # texto nao e selecionado: digitar "1" virava "11" e dava erro. Pelo Enter o texto vem selecionado. Apaga antes.
        limpar_campo("ident")
        digitar(valores["ident"], "ident")
        pendentes.remove("ident")
        for vez in range(len(COLUNAS_RATEIO) + 2):
            if not pendentes:
                break
            devagar(f"Enter {vez + 1} na linha ICMS TERCEIROS (procurando {NOME_RATEIO[pendentes[0]]})")
            foco = enter_rateio("rateio", vez, foco)
            if not na_linha(foco, rateio["terceiros"]):
                raise Travou("rateio", f"o cursor saiu da linha ICMS TERCEIROS antes do campo {NOME_RATEIO[pendentes[0]]}"
                                       f"{local(foco)}; confira os campos dessa linha")
            coluna = coluna_rateio(foco, rateio)
            registrar(f"  Enter {vez + 1} na linha ICMS TERCEIROS: foco={foco} coluna={coluna}", mostrar=False)
            if coluna in pendentes:
                if coluna != pendentes[0]:
                    raise Travou("rateio", f"o cursor pulou o campo {NOME_RATEIO[pendentes[0]]} da linha ICMS TERCEIROS")
                digitar(valores[coluna], coluna)
                pendentes.remove(coluna)
        if pendentes:
            raise Travou("rateio", f"nao cheguei no campo {NOME_RATEIO[pendentes[0]]} da linha ICMS TERCEIROS")
        botao = achar_botao("tributos_ok", rateio["tributos_ok"])
        if not botao:
            raise Travou("rateio", "nao achei o botao Tributos OK na tela")
        ret = janela.rectangle()
        checar_parada()
        devagar("clicar em Tributos OK")
        clicar(ret.left + (botao[0] + botao[2]) / 2, ret.top + (botao[1] + botao[3]) / 2)
        registrar(f"  linha {r['linha']}: clique em Tributos OK", mostrar=False)
        prazo = time.time() + 2 * render
        while time.time() < prazo:
            checar_parada()
            time.sleep(0.3)
            if not achar_botao("tributos_ok", rateio["tributos_ok"]):
                time.sleep(0.3)
                return foco_tela(janela)
        raise Travou("rateio", "o Rateio nao fechou depois do Tributos OK (aviso do SRP?)")

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
        return foco_tela(janela)

    def esperar_na_chave(texto, linhas):
        """Pausa ate o cursor estar na Chave Acesso NF-e (proxima nota) e o usuario clicar em Continuar."""
        registrar(texto)
        marcar(linhas, "pausado")
        while True:
            esperar_continuar()
            caixa = foco_depois_de_continuar()
            if caixa and posicao_ok("chave", caixa, rotulos, None, lidas) is not False:
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
            if caixa and posicao_ok("item", caixa, rotulos, None, lidas) is not False:
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

    def clicar_integrar(nota, vez):
        """Clica em Integrar se o botao estiver a vista (confere de novo pelo OCR antes de clicar). True = clicou."""
        botao = localizar_integrar()
        if not botao:
            return False
        ret = janela.rectangle()
        x, y = ret.left + (botao[0] + botao[2]) / 2, ret.top + (botao[1] + botao[3]) / 2
        checar_parada()
        devagar("clicar em Integrar" + (" de novo" if vez > 1 else ""))
        antes = core.capturar(janela)
        clicar(x, y)
        registrar(f"  nota ...{nota['chave'][-6:]}: clique {vez} em Integrar em ({x:.0f}, {y:.0f})", mostrar=False)
        prazo = time.time() + render  # espera a tela reagir ao clique antes do proximo passo
        while time.time() < prazo and not tela_mudou(antes, core.capturar(janela)):
            checar_parada()
            time.sleep(0.1)
        time.sleep(0.3)
        return True

    def esperar_tela(antes, limite):
        """Espera a tela mudar em relacao a `antes` e parar de mudar (janela terminou de abrir/fechar), no maximo
        `limite` s. Evita ler a tela de novo (~2,5 s em 4K) antes de ela mudar."""
        prazo, anterior = time.time() + limite, None
        while time.time() < prazo:
            checar_parada()
            agora = core.capturar(janela)
            if anterior is not None and tela_mudou(antes, agora, PIXELS_JANELA) and not tela_mudou(anterior, agora, PIXELS_JANELA):
                return True
            anterior = agora
            time.sleep(0.15)
        return False

    def ajustar_acessoria(imagem, area):
        """Janela "Valor Calculado para Documento": le os 3 totais (varias leituras do recorte; em 1366x768 o texto e
        pequeno) e, com despesa acessoria de ate LIMITE_ACESSORIA, clica em Ajustar. Devolve None (clicou) ou o motivo
        para nao clicar."""
        recorte = imagem.crop(area)
        valores, botao = None, None
        # no print real (06/11/2026) a leitura 4x leu os 3 totais e a binarizada nao leu nada: fica por ultimo
        for img, escala in ((recorte, 2), (recorte, 4), (binarizar(recorte), 2)):
            palavras = core.palavras_tela(img, tesseract, escala=escala)
            valores = valores or valores_acessoria(normalizar(" ".join(p[0] for p in palavras)))
            botao = botao or botao_ajustar(frases(palavras))
            if valores and botao:
                break
        registrar(f"  despesa acessoria: valores lidos={valores}, botao Ajustar={botao}", mostrar=False)
        if not valores:
            return "nao consegui ler os totais (Geral, Documento e Acessoria conferindo entre si)"
        geral, documento, acessoria = valores
        if abs(acessoria) > LIMITE_ACESSORIA + 1e-9:
            return f"despesa acessoria de {acessoria:.2f} (Geral {geral:.2f}, Documento {documento:.2f}), acima de 0,05"
        if not botao:
            return f"despesa acessoria de {acessoria:.2f}, mas nao achei o botao Ajustar"
        devagar(f"clicar em Ajustar (despesa acessoria de {acessoria:.2f})")
        ret = janela.rectangle()
        checar_parada()
        clicar(ret.left + area[0] + (botao[0] + botao[2]) / 2, ret.top + area[1] + (botao[1] + botao[3]) / 2)
        registrar(f"  despesa acessoria de {acessoria:.2f} (ate 0,05): cliquei em Ajustar", mostrar=False)
        return None

    def integrar(nota):
        """Integrar e o que vem depois (usuario, 04/11/2026), lendo a tela antes de cada passo:
        aviso do frete alterado para CIF -> Enter (OK); "Resumo dos Tributos do Documento" -> Enter (Confirmar);
        janela "Centro de Custo e Conta Contabil para Frete LTCC" -> CLIQUE em "Frete LTCC" (Enter ali roda sem fim);
        "Valor Calculado para Documento" com despesa acessoria de ate 0,05 -> CLIQUE em Ajustar (acima: pausa);
        "Documento Integrado com Sucesso." -> Enter (OK) e acaba. Sem nenhuma dessas janelas: Enter, ate
        ENTERS_DEPOIS_INTEGRAR, e depois Integrar de novo. Janela ja tratada que continua aberta: espera, sem Enter.
        Devolve 0 (nao achei o Integrar), 1 (nao achei o 2o Integrar), 2 (sem "Integrado com Sucesso") ou 3 (sucesso)."""
        if not clicar_integrar(nota, 1):
            return 0
        if not tesseract:
            return 1
        etiqueta = f"nota ...{nota['chave'][-6:]}"
        feitos, enters, cliques, esperas, desde_clique = [], 0, 1, 0, 0
        for _ in range(MAX_PASSOS_INTEGRAR):
            checar_parada()
            imagem = core.capturar(janela)
            t_leitura = time.time()
            # So o recorte das janelas por cima (barra de titulo azul): ~0,6 s em vez de ~2,2 s da tela inteira. Sem
            # barra nem le (so a busca da barra, ~0,07 s), menos na 1a leitura depois de cada clique no Integrar e na
            # ultima espera, que leem a tela inteira para conferir a barra: janela achada ali sem barra = daqui em diante
            # sempre a tela inteira (`confia_barra`). Antes, a tela inteira antes de cada um dos 7 Enter custou ~33 s
            # numa nota (04/11/2026 19:21). Recorte sem janela conhecida: tela inteira.
            barras = barras_dialogo(imagem)
            area, janelas = None, []
            if barras:
                area = area_dialogo(imagem, barras)
                janelas = janelas_no_texto(normalizar(" ".join(p[0] for p in core.palavras_tela(imagem.crop(area), tesseract))))
            if not janelas and (barras or not confia_barra[0] or desde_clique == 0 or esperas >= LEITURAS_AVISO - 1):
                area = (0, 0, imagem.width, imagem.height)
                janelas = janelas_no_texto(normalizar(" ".join(p[0] for p in core.palavras_tela(imagem, tesseract))))
                if janelas and not barras and confia_barra[0]:
                    confia_barra[0] = False
                    registrar(f"  janela {janelas} sem barra de titulo azul: daqui em diante leio sempre a tela inteira",
                              mostrar=False)
            desde_clique += 1
            onde = "nao li (nenhuma barra de janela)" if area is None else (
                "tela inteira" if area == (0, 0, imagem.width, imagem.height) else f"recorte {area}")
            registrar(f"  leitura depois do Integrar: {time.time() - t_leitura:.1f}s, {onde}, janelas={janelas}", mostrar=False)
            nova = next((nome for nome in janelas if nome not in feitos), None)
            if nova == "referencia":
                try:
                    fechar_referencia(imagem, area)
                except Travou as erro:
                    registrar(f"Nota ...{nota['chave'][-6:]}: {erro.motivo}. Termine a integracao no SRP.")
                    return 2
                # no log de 05/11 16:30 a janela fechou no 7o Enter e o 2o Integrar seguiu com frete, Resumo, LTCC e
                # sucesso: depois do OK vai direto ao 2o Integrar, sem Enter as cegas
                enters = ENTERS_DEPOIS_INTEGRAR
            elif nova == "acessoria":
                problema = ajustar_acessoria(imagem, area)
                if problema:
                    registrar(f"Nota ...{nota['chave'][-6:]}: janela 'Valor Calculado para Documento': {problema}. "
                              "Confira e clique em Ajustar ou Cancelar no SRP.")
                    return 2
            elif nova == "ltcc":
                botao = botao_frete_ltcc(frases(core.palavras_tela(binarizar(imagem.crop(area)), tesseract)))
                if not botao:
                    registrar(f"Nota ...{nota['chave'][-6:]}: nao achei o botao Frete LTCC. Clique em Frete LTCC no SRP.")
                    return 2
                devagar("clicar em Frete LTCC")
                ret = janela.rectangle()
                clicar(ret.left + area[0] + (botao[0] + botao[2]) / 2, ret.top + area[1] + (botao[1] + botao[3]) / 2)
            elif nova:
                devagar({"frete": "Enter no aviso 'frete alterado para CIF' (OK)",
                         "resumo": "Enter no 'Resumo dos Tributos do Documento' (Confirmar)",
                         "sucesso": "Enter em 'Documento Integrado com Sucesso' (OK)"}[nova])
                core.send_keys("{ENTER}", pause=0.03)
            if nova:
                registrar(f"  {etiqueta}: {nova} tratado", mostrar=False)
                feitos.append(nova)
                esperas = 0
                if nova == "sucesso":
                    registrar(f"Nota ...{nota['chave'][-6:]}: o SRP disse 'Documento Integrado com Sucesso'.")
                    esperar_tela(imagem, 1.5)
                    return 3
                esperar_tela(imagem, render)
                continue
            # a NF de referencia vem antes do 2o Integrar: depois dela, sem janela, segue para o 2o Integrar
            if janelas or any(f != "referencia" for f in feitos) or (cliques == 2 and enters >= ENTERS_DEPOIS_INTEGRAR):
                # janela ja tratada ainda aberta, entre uma janela e outra, ou ja integrei 2 vezes: so espera a tela
                # mudar (Enter aqui pode cair na janela seguinte; na do Frete LTCC roda sem fim)
                esperas += 1
                if esperas >= LEITURAS_AVISO:
                    break
                esperar_tela(imagem, render)
                continue
            if enters < ENTERS_DEPOIS_INTEGRAR:
                enters += 1
                devagar(f"Enter {enters} de {ENTERS_DEPOIS_INTEGRAR} depois do Integrar")
                core.send_keys("{ENTER}", pause=0.03)
                esperar_tela(imagem, 1.0)
                continue
            if not clicar_integrar(nota, 2):
                registrar(f"Nota ...{nota['chave'][-6:]}: dei {ENTERS_DEPOIS_INTEGRAR} Enter depois do Integrar, mas nao "
                          "achei o Integrar de novo. Termine a integracao no SRP.")
                return 1
            cliques, desde_clique = 2, 0
        registrar(f"Nota ...{nota['chave'][-6:]}: nao vi a mensagem 'Documento Integrado com Sucesso'"
                  + (f" (janela ainda aberta: {', '.join(janelas)})" if janelas else "") + ". Confira no SRP.")
        return 2

    def ir_para_chave():
        """Depois do "Integrado com Sucesso" o SRP volta para a tela vazia com o foco no CNPJ Fornecedor: clica no campo
        Chave Acesso NF-e (linha de cima, a direita do rotulo lido no inicio) e confere que o foco foi para la (usuario,
        04/11/2026). Devolve o foco na Chave ou None (o robo pausa como antes)."""
        rotulo = rotulos.get("chave")
        if not rotulo:
            return None
        prazo, foco = time.time() + render, None
        while time.time() < prazo and foco is None:  # espera a tela voltar (foco no CNPJ)
            checar_parada()
            foco = foco_tela(janela)
            if foco is None:
                time.sleep(0.1)
        altura = rotulo[3] - rotulo[1]
        ret = janela.rectangle()
        x, y = rotulo[2] + max(40, 3 * altura), (rotulo[1] + rotulo[3]) / 2
        checar_parada()
        devagar("clicar no campo Chave Acesso NF-e (proxima nota)")
        clicar(ret.left + x, ret.top + y)
        novo = aguardar(janela, foco, render, fixos)
        if isinstance(novo, tuple) and posicao_ok("chave", novo, rotulos, None, lidas):
            registrar(f"  cliquei na Chave Acesso NF-e em ({ret.left + x:.0f}, {ret.top + y:.0f}): foco={novo}", mostrar=False)
            return novo
        registrar(f"  cliquei na Chave Acesso NF-e, mas o foco nao foi para la (foco={novo})", mostrar=False)
        return None


    try:
        tesseract = core.caminho_tesseract()
        janela = core.janela_citrix(janela_trecho)
        janela.set_focus()
        ret0 = janela.rectangle()
        tamanho = (ret0.right - ret0.left, ret0.bottom - ret0.top)
        rateio_salvo[0] = ler_posicoes(tamanho)  # da execucao anterior: a 1a nota tambem nao le a tela inteira
        if rateio_salvo[0]:
            registrar("  posicoes do Rateio da execucao anterior (mesmo tamanho de janela)", mostrar=False)
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
        faltam = [NOME.get(c, c) for c in ("chave", "filial", "data", "valor_total", "frete", "item", "oper", "qtd", "unit")
                  if c not in rotulos]
        if faltam:
            aviso(f"ATENCAO: nao li na tela os rotulos {', '.join(faltam)}: nesses campos o robo so conta os Enter, sem conferir a posicao.")
        if "integrar" not in rotulos:
            aviso("ATENCAO: nao achei o botao Integrar na tela: no fim de cada nota o robo pausa para voce clicar.")
        while time.time() - inicio < espera:
            if parar.is_set():
                registrar("Interrompido antes de iniciar.")
                return 1
            time.sleep(0.1)
        foco = foco_tela(janela)
        if not foco:
            registrar("Nao iniciei: nenhum campo do SRP em foco (ciano). Clique no campo Chave Acesso NF-e e inicie de novo. "
                      "Nada foi digitado.")
            return 1
        if posicao_ok("chave", foco, rotulos, None, lidas) is False:
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
            serie_com_ponto_feita[0] = False
            try:
                marcar([linhas[0]], "executando")
                if posicao_ok("chave", foco, rotulos, None, lidas) is False:
                    raise Travou("chave", f"o cursor nao esta no campo {NOME['chave']}{local(foco)}")
                for indice, (campo, quantos) in enumerate(SEQ_CABECALHO):
                    if foco is None:
                        raise Travou(campo, f"nao sei em que campo o cursor esta antes do {NOME[campo]}{local(foco)}")
                    digitar(nota[campo], campo)
                    if campo == "frete":  # ultimo do cabecalho: Enter ate o Item da grade (janela de NF de referencia no caminho)
                        foco = ir_para_item(foco)
                        continue
                    try:
                        foco = ir_ate(campo, SEQ_CABECALHO[indice + 1][0], quantos, foco, anterior=foco)
                    except DocumentoExiste as existe:  # so no Enter depois da Chave
                        foco = serie_com_ponto(nota, foco, existe.motivo)
                caixa_item = None  # Item da linha do item anterior
                topos_linha1 = {}  # campo -> topo dele no 1o item (1a linha da grade): o titulo da coluna fica acima
                for indice, r in enumerate(nota["itens"]):
                    marcar([r["linha"]], "executando")
                    if caixa_item is not None:
                        foco = descer_para_item(foco, caixa_item)
                    if posicao_ok("item", foco, rotulos, None, lidas) is False:
                        if retomar is None:
                            raise Travou("item", f"o cursor nao esta no Item da grade{local(foco)}")
                        foco = esperar_no_item(r)
                    caixa_item = foco
                    for posicao, (campo, quantos) in enumerate(SEQ_ITEM):
                        if foco is None and campo != "unit":  # o Valor Unit fica amarelo (conferido na entrada)
                            raise Travou(campo, f"o cursor esta num campo sem cor antes do {NOME[campo]}; nao digito sem ver o campo")
                        if campo != "unit" and posicao_ok(campo, foco, rotulos, None, lidas) is False:
                            raise Travou(campo, f"o cursor nao esta na coluna {NOME[campo]}{local(foco)}")
                        if campo in TITULO_COLUNA and tesseract:
                            if indice == 0:
                                topos_linha1[campo] = foco[1]
                            titulo = titulo_acima(core.capturar(janela), foco, tesseract, topos_linha1.get(campo))
                            registrar(f"  titulo acima do {NOME[campo]}: {titulo!r}", mostrar=False)
                            if not any(t in titulo for t in TITULO_COLUNA[campo]):
                                raise Travou(campo, f"o cursor nao esta na coluna {NOME[campo]} (titulo acima: '{titulo}')")
                        digitar(r[campo], campo)
                        alvo = SEQ_ITEM[posicao + 1][0] if posicao + 1 < len(SEQ_ITEM) else None
                        if alvo in ("oper", "qtd"):  # colunas visiveis no inicio: Enter ate a coluna certa
                            foco = ir_ate(campo, alvo, quantos, foco, anterior=foco)
                        elif alvo is None:  # Valor do ICMS: Enter ate a linha ICMS TERCEIROS do Rateio
                            foco = preencher_rateio(r, foco)
                        elif alvo == "unit":  # Enter no Qtd.Fisica (ciano), Enter no Valor Unit (amarelo, grade rola)
                            foco = entrar_no_valor_unit(enters(campo, quantos - 1, foco))
                        else:  # depois do Valor Unit a grade rola para a direita: so a contagem combinada
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
                if serie_com_ponto_feita[0] and serie_ponto:
                    serie_ponto(linhas)
                marcar(linhas, "ok")
                continue
            cliques = integrar(nota)
            clicou = cliques >= 2
            registrar(f"OK {etiqueta}: {len(linhas)} item(ns) em {time.time() - t0:.1f}s" +
                      ("; integrada com sucesso." if cliques == 3 else "."))
            if cliques == 0:
                registrar(f"{etiqueta}: nao achei o botao Integrar na tela. Clique em Integrar no SRP.")
            if serie_com_ponto_feita[0] and serie_ponto:  # antes do "ok": a pintura da planilha usa a outra cor
                serie_ponto(linhas)
            marcar(linhas, "ok")
            if cliques == 3:  # integrada com sucesso: clica na Chave (tambem na ultima nota, para a proxima planilha)
                foco_chave = ir_para_chave()
            if ultima:
                if not clicou and retomar is not None:
                    registrar("Depois de integrar, clique em Continuar para encerrar.")
                    marcar(linhas, "pausado")
                    esperar_continuar()
                    marcar(linhas, "ok")
                break
            if cliques == 3 and foco_chave:  # integrada com sucesso e cursor na Chave: proxima nota, sem pausa
                foco = foco_chave
                continue
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
        with _TRAVA_LOG:
            log.close()
