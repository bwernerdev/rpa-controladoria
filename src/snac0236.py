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
# Nat Oper do item pela UF do emitente (os 2 primeiros digitos da CHAVE NF = codigo IBGE da UF; 42 = SC):
# dentro de SC 1949, fora 2949 (usuario, 24/11/2026). So na OPER SPM 542: nas outras o robo passa pelo campo sem mexer.
# A coluna NAT OPER da planilha e opcional; preenchida num item da 542, so confere.
UF_SC = "42"
OPER_COM_NAT = "542"
NAT_DENTRO, NAT_FORA = "1949", "2949"
TITULO_NAT_OPER = "NAT OPER"
# coluna CD EMPRESA (opcional; usuario, 26/11/2026): preenchida, a OPER SPM sai dela e a coluna OPER SPM e ignorada
TITULO_CD_EMPRESA = "CD EMPRESA"
OPER_POR_EMPRESA = {"30": "542", "36": "536"}

# (campo, quantos Enter depois de digitar). "filial1"/"filial2" = as duas partes da FILIAL FISCAL.
ENTERS_FRETE_ITEM = 10  # Enter previstos do Frete ao Item da grade (Marcas, Observacoes, Cond. Pagamento, lista sem cor...)
SEQ_CABECALHO = (("chave", 5), ("filial1", 1), ("filial2", 2), ("data", 7), ("valor_total", 1), ("frete", 10))
# depois do Oper SPM o 1o Enter cai no Nat Oper (04/11/2026) e o 2o no Qtd Fiscal. O SRP ja traz um Nat Oper; desde
# 24/11/2026, na OPER SPM 542, o robo apaga e digita o da chave (`nat_oper_da_chave`); nas outras so da o Enter
# depois do Valor do ICMS: Enter ate a linha ICMS TERCEIROS do Rateio (`preencher_rateio`, sem contagem fixa)
SEQ_ITEM = (("item", 1), ("oper", 1), ("nat_oper", 1), ("qtd", 2), ("unit", 1), ("icms", 0))
# a FILIAL FISCAL da planilha vai nos campos "Filial Fisica" da tela: depois da Chave + 5 Enter o SRP preenche CNPJ,
# Serie e Docto e o cursor cai na Filial Fisica (visto na tela real em 04/11/2026, monitor 4K)
NOME = {"chave": "Chave Acesso NF-e", "filial1": "Filial Fisica (1a parte)", "filial2": "Filial Fisica (2a parte)",
        "data": "Data Emissao", "valor_total": "Valor Total Documento", "frete": "Frete", "item": "Item",
        "oper": "Oper SPM", "nat_oper": "Nat Oper", "qtd": "Qtd Fiscal", "unit": "Valor Unit", "icms": "Valor do ICMS",
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
    if nome.startswith(("icms terceiro", "base calc")):
        return "icms_terceiro"
    if re.fullmatch(r"nat\.? ?oper\.?|natureza (de )?operacao", nome):  # coluna opcional (24/11/2026)
        return "nat_oper"
    if nome == "cd empresa":  # coluna opcional (26/11/2026)
        return "cd_empresa"
    if nome.startswith("aliquota"):  # "ALIQUOTA" (da linha ICMS TERCEIROS do Rateio; era BASE CALC. ICMS ate 05/11/2026)
        return "aliquota_icms"
    return COLUNAS.get(nome)


def formatar_valor(valor):
    """Numero como o SRP quer: ponto decimal, sem milhar (47407.28). Devolve (texto, erro)."""
    if valor is None or valor == "":
        return "", None
    if isinstance(valor, bool):
        return "", "valor inválido"
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and valor.is_integer():
            valor = int(valor)
        texto = format(Decimal(repr(valor)), "f") if isinstance(valor, float) else str(valor)
    else:
        texto = str(valor).strip().replace(" ", "")
        if texto.count(",") == 1 and "." not in texto:
            texto = texto.replace(",", ".")  # 47407,28 digitado como texto
    if not re.fullmatch(r"\d+(?:\.\d+)?", texto):
        return "", f"valor inválido: {valor!r} (use só números, com ponto ou vírgula decimal, sem separador de milhar)"
    try:
        Decimal(texto)
    except InvalidOperation:
        return "", f"valor inválido: {valor!r}"
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
    return "", f"data inválida: {valor!r} (use dd/mm/aaaa)"


def nat_oper_da_chave(chave, oper=OPER_COM_NAT):
    """Nat Oper do item: 1949 se o emitente da nota e de SC (chave comeca com 42), senao 2949. Sem chave ou com outra
    OPER SPM que nao a 542: '' (o robo nao mexe no Nat Oper que o SRP trouxe)."""
    if not chave or oper != OPER_COM_NAT:
        return ""
    return NAT_DENTRO if chave[:2] == UF_SC else NAT_FORA


def formatar_chave(valor):
    if valor is None or valor == "":
        return "", None
    if not isinstance(valor, str):
        # o Excel guarda numero com 15 digitos de precisao: os 44 da chave ja chegam errados
        return "", "a CHAVE NF está como número no Excel e perde dígitos: formate a coluna como Texto e cole a chave de novo"
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


def _cor_rgb(fill):
    """'RRGGBB' do fundo solido em RGB (sem cor do tema), ou None."""
    if fill is None or fill.fill_type != "solid" or fill.fgColor.type != "rgb" or not isinstance(fill.fgColor.rgb, str):
        return None
    return fill.fgColor.rgb[-6:].upper()


def ler_planilha(caminho):
    """Linhas da planilha ja formatadas para o SRP. Devolve (registros, erros)."""
    from openpyxl import load_workbook

    # read_only deixa o arquivo aberto ate o close(): sem ele a planilha ficava travada e o Excel nao conseguia salvar
    # por cima (06/11/2026: "salvar e fechar" no Iniciar travou o Excel numa janela de erro)
    livro = load_workbook(caminho, read_only=True, data_only=True)
    try:
        aba = core.aba_de_dados(livro)
        nome_aba = aba.title
        celulas = list(aba.iter_rows())  # nunca a aba "Lançados" (copia das integradas)
        tema = core.cores_do_tema(getattr(livro, "loaded_theme", None))
    finally:
        livro.close()
    valores = [tuple(c.value for c in linha) for linha in celulas]
    # formula sem o valor guardado (planilha gravada sem o Excel calcular): as celulas viriam vazias e os erros
    # ("CHAVE NF em branco"...) enganariam (24/11/2026)
    faltam = core.formulas_sem_valor(caminho, nome_aba, valores)
    if faltam:
        exemplos = ", ".join(faltam[:4]) + (" ..." if len(faltam) > 4 else "")
        raise ValueError(f"A planilha tem {core.qtd(len(faltam), 'célula', 'células')} com fórmula sem o valor calculado "
                         f"({exemplos}). Abra a planilha no Excel, salve (Ctrl+S) e carregue de novo.")
    linhas = iter(valores)
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
                         f"{', '.join(repetidas)}): deixe só a {TITULOS['icms_terceiro']}")
    faltando = [TITULOS[c] for c in TITULOS if c not in indices]
    if faltando:
        raise ValueError(f"Colunas ausentes no cabeçalho: {', '.join(faltando)}")
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
        # nota lancada com "." na Serie: o robo pinta com VERDE_SERIE_PONTO (o app mostra outro verde na tabela)
        r["serie_ponto"] = r["verde"] and any(_cor_rgb(getattr(c, "fill", None)) == core.VERDE_SERIE_PONTO
                                              for c in preenchidas)
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
            problemas.append(f"FRETE inválido: {bruto['frete']!r} (use CIF ou FOB)")
        for campo in VALORES:
            r[campo], erro = formatar_valor(bruto[campo])
            problemas.append(f"{TITULOS[campo]}: {erro}" if erro else None)
        r["nat_oper_planilha"] = core.texto(bruto.get("nat_oper"))
        for campo in ("item", "oper"):
            r[campo] = core.texto(bruto[campo])
            if r[campo] and not re.fullmatch(r"[0-9A-Za-z.\-]+", r[campo]):
                problemas.append(f"{TITULOS[campo]} com caracteres invalidos: {r[campo]!r}")
        empresa = core.texto(bruto.get("cd_empresa"))
        if empresa:  # CD EMPRESA preenchido manda na OPER SPM (a da planilha e ignorada)
            r["oper"] = OPER_POR_EMPRESA.get(empresa, "")
            if not r["oper"]:
                problemas.append(f"{TITULO_CD_EMPRESA} {empresa!r} sem OPER SPM definida (30 = 542, 36 = 536)")
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
            erros.append(f"linha {r['linha']}: a CHAVE NF ...{r['chave'][-6:]} já apareceu antes (linha {vistas[r['chave']]}); "
                         "deixe os itens da mesma nota em linhas seguidas")
        if not r["chave"] or not notas or r["chave"] != notas[-1]["chave"]:
            if not r["chave"] and notas:
                if any(r[c] for c in CABECALHO_SEM_CHAVE):  # cabecalho de outra nota sem a chave: nao adivinhar
                    erros.append(f"linha {r['linha']}: CHAVE NF em branco com o cabeçalho preenchido (para mais um item "
                                 "da nota de cima, deixe a chave e o cabeçalho em branco ou repita a chave)")
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
    for nota in notas:  # o Nat Oper sai da chave da nota (item com a chave em branco = o da nota de cima)
        for r in nota["itens"]:
            r["nat_oper"] = nat_oper_da_chave(nota["chave"], r["oper"])
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
            planilha = r.get("nat_oper_planilha")
            if planilha and r["nat_oper"] and planilha != r["nat_oper"]:  # so na OPER SPM 542
                erros.append(f"linha {r['linha']}: {TITULO_NAT_OPER} {planilha} não confere com a CHAVE NF (emitente "
                             f"{'de SC' if r['nat_oper'] == NAT_DENTRO else 'fora de SC'}: {r['nat_oper']}; dentro de SC "
                             f"{NAT_DENTRO}, fora {NAT_FORA})")
            terceiro, aliquota = (Decimal(r.get(c) or "0") for c in ("icms_terceiro", "aliquota_icms"))
            if (terceiro != 0) != (aliquota != 0):
                erros.append(f"linha {r['linha']}: {TITULOS['icms_terceiro']} e {TITULOS['aliquota_icms']} precisam ser os dois 0 "
                             "(sem imposto) ou os dois diferentes de 0 (com imposto)")
            elif 0 < aliquota < 1:  # celula em porcentagem no Excel: 12% chega como 0,12 e o SRP receberia 0.12
                erros.append(f"linha {r['linha']}: ALÍQUOTA {r['aliquota_icms']} parece porcentagem do Excel (12% vira 0,12): "
                             "digite o número da alíquota, ex.: 12")
            elif aliquota < 0 or aliquota > 100:
                erros.append(f"linha {r['linha']}: ALÍQUOTA inválida: {r['aliquota_icms']}")
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
        teclas = [("(Nat Oper sem mexer)" if c == "nat_oper" and not r[c] else f"'{r[c]}'") + f" {n}xENTER"
                  for c, n in SEQ_ITEM]
        partes.append(f"| linha {r['linha']}: " + " ".join(teclas)
                      + " ENTER até ICMS TERCEIROS no Rateio: "
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
        # "Nat" em cima, "Oper" na linha de baixo; o "DM" ao lado pode vir junto ("oper dm"; print de 24/11/2026)
        "nat_oper": lambda t: (t.split()[:1] == ["oper"] and "spm" not in t) or t.startswith("nat oper"),
        "qtd": lambda t: "fiscal" in t and "td" in t and "unit" not in t,
        "unit": lambda t: "unitario" in t,
    }
    for campo, pred in grade.items():
        caixa = next((c for texto, c in lista if pred(texto) and na_linha_da_grade(c)), None)
        if caixa:
            achados[campo] = caixa
    # "Oper" sozinho so vale como Nat Oper entre o Oper Spm e o Qtd.Fiscal (outra palavra "Oper" na mesma altura nao
    # pode virar o rotulo; sem ele o robo so conta os Enter, e o Qtd.Fiscal continua conferido)
    nat, oper, qtd = (achados.get(c) for c in ("nat_oper", "oper", "qtd"))
    if nat and ((oper and nat[0] <= oper[2]) or (qtd and nat[2] >= qtd[0])):
        del achados["nat_oper"]
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
    # proporcional a tela (1366x768: ~107 px); a selecao do texto da Data deixava ~170 px sem ciano (4K)
    folga = max(300 * imagem.width // 3840, 8 * (base - topo))
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
    if campo in ("item", "oper", "nat_oper", "qtd"):  # colunas da grade visiveis no inicio (depois do Valor Unit a grade rola)
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
# linha da grade de tributos do Rateio em que o robo digita (`valores_rateio`): ICMS TERCEIROS (cod. 114). Na OPER SPM
# 542 a grade so tem a linha ICMS S/ TRANSFERENCIA (cod. 101, Ident 2 e Base preenchidos pelo SRP; print do usuario,
# 24/11/2026): os valores vao nela. Nos dois casos a chave do rotulo em `localizar_rateio` continua "terceiros".
LINHA_RATEIO = {"terceiros": "ICMS TERCEIROS", "transferencia": "ICMS S/ TRANSFERENCIA"}
TESTE_LINHA_RATEIO = {
    # em 1366x768 o OCR leu "11s icms erceiros" e "is terceiros" (05/11/2026)
    "terceiros": lambda t: "erceiro" in t or "icms tercei" in t,
    "transferencia": lambda t: "ransfer" in t,
}


def linha_rateio(r):
    """Linha do Rateio do item: "transferencia" na OPER SPM 542, senao "terceiros"."""
    return "transferencia" if r.get("oper") == OPER_COM_NAT else "terceiros"


NOME_RATEIO = {"ident": "Ident Tributo", "base": "Base Calculo", "aliquota": "Aliquota", "lancamento": "Lancamento"}


def arquivo_posicoes():
    """Posicoes do Rateio lidas no OCR, guardadas entre execucoes (ao lado do exe; fora do git)."""
    return core.pasta_base() / "posicoes_0236.json"


def _chave_posicoes(linha):
    """Chave no arquivo: "rateio" (grade do ICMS TERCEIROS, o nome de antes) ou "rateio_<linha>" (ex.: a da 542)."""
    return "rateio" if linha == "terceiros" else f"rateio_{linha}"


def ler_posicoes(tamanho):
    """Posicoes do Rateio salvas para uma janela do Citrix deste tamanho, por linha da grade de tributos
    ({"terceiros": {...}, "transferencia": {...}}); outro tamanho/monitor ou nada salvo = {}."""
    try:
        dados = json.loads(arquivo_posicoes().read_text(encoding="utf-8"))
        if tuple(dados.get("janela", ())) == tuple(tamanho):
            return {linha: {chave: tuple(caixa) for chave, caixa in dados[_chave_posicoes(linha)].items()}
                    for linha in LINHA_RATEIO if _chave_posicoes(linha) in dados}
    except (OSError, ValueError, KeyError, TypeError, AttributeError):
        pass
    return {}


def _gravar_no_arquivo_posicoes(tamanho, chave, valor):
    """Grava `chave` no `posicoes_0236.json`, mantendo o resto salvo para a mesma janela (outro tamanho = comeca de novo)."""
    try:
        try:
            dados = json.loads(arquivo_posicoes().read_text(encoding="utf-8"))
            if tuple(dados.get("janela", ())) != tuple(tamanho):
                dados = {}
        except (OSError, ValueError):
            dados = {}
        dados["janela"] = list(tamanho)
        dados[chave] = valor
        core.gravar_json(arquivo_posicoes(), dados)
    except (OSError, AttributeError, TypeError):
        pass


def gravar_posicoes(tamanho, rateio, linha="terceiros"):
    """Grava as posicoes da grade da `linha`, mantendo as das outras linhas salvas para a mesma janela."""
    _gravar_no_arquivo_posicoes(tamanho, _chave_posicoes(linha), {c: list(v) for c, v in rateio.items()})


def ler_sem_cor(tamanho):
    """(campo, Enter) em que o foco ficou sem cor nas execucoes anteriores, com a janela do Citrix deste tamanho."""
    try:
        dados = json.loads(arquivo_posicoes().read_text(encoding="utf-8"))
        if tuple(dados.get("janela", ())) == tuple(tamanho):
            return {(str(campo), int(vez)) for campo, vez in dados.get("sem_cor", [])}
    except (OSError, ValueError, TypeError, AttributeError):
        pass
    return set()


def ler_ok_referencia(tamanho):
    """{area da janela NF de referencia: botao OK (coordenadas da area)} salvos para a janela do Citrix deste tamanho."""
    try:
        dados = json.loads(arquivo_posicoes().read_text(encoding="utf-8"))
        if tuple(dados.get("janela", ())) == tuple(tamanho):
            return {tuple(int(v) for v in x[:4]): tuple(int(v) for v in x[4:8]) for x in dados.get("ok_referencia", [])}
    except (OSError, ValueError, TypeError, AttributeError):
        pass
    return {}


def gravar_ok_referencia(tamanho, salvos):
    _gravar_no_arquivo_posicoes(tamanho, "ok_referencia", [list(area) + list(botao) for area, botao in salvos.items()])


def gravar_sem_cor(tamanho, vistos):
    _gravar_no_arquivo_posicoes(tamanho, "sem_cor", sorted([campo, vez] for campo, vez in vistos))


def tem_imposto(r):
    """Os dois valores da planilha diferentes de 0 = nota com imposto (vazio conta como 0)."""
    return Decimal(r.get("icms_terceiro") or "0") != 0 and Decimal(r.get("aliquota_icms") or "0") != 0


def valores_rateio(r):
    """O que digitar na linha ICMS TERCEIROS, por coluna (na ordem da tela)."""
    if tem_imposto(r):
        return {"ident": "1", "base": r["icms_terceiro"], "aliquota": r["aliquota_icms"]}
    return {"ident": "0", "base": "0", "aliquota": "0"}


def localizar_rateio(palavras, linha="terceiros"):
    """Rotulos do Rateio lidos no OCR: a linha do item (chave "terceiros": ICMS TERCEIROS ou, com `linha`
    "transferencia", ICMS S/ TRANSFERENCIA), colunas da grade de tributos e botao Tributos OK."""
    lista = frases(palavras)
    testes = {
        "terceiros": TESTE_LINHA_RATEIO[linha],
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


def ler_rateio(imagem, tesseract, linha="terceiros"):
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
            for chave, caixa in localizar_rateio(core.palavras_tela(img, tesseract, escala=escala), linha).items():
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
# Espera pela 1a janela depois do 1o Integrar antes do 1o Enter as cegas (usuario, 24/11/2026: "a tela que aparece quando
# clica em Integrar nao esta dando tempo de aparecer"; nas notas da 542 o robo leu a tela 1,7 s depois do clique, nao viu
# janela e deu ~6 Enter em 8 s). Nos logs de novembro so 5 de 242 notas precisaram do 2o Integrar: quase sempre uma janela
# aparece, entao a espera acaba assim que a barra azul surge e so custa o tempo todo quando nada aparece.
ESPERA_JANELA_INTEGRAR = 20
LER_JANELA_A_CADA = 3  # nessa espera, sem barra azul, le a tela inteira (~1,3 s em 1366x768) a cada 3 s


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


def natureza_operacao(texto):
    """Janela "Natureza de Operacao" (lista de valores do SRP: Find "1949%", Cod./Descricao, [Find] [OK] [Cancel]) que
    abre no Enter depois do Nat Oper digitado, com o cursor parado no Nat Oper por tras (print de 24/11/2026: dois 1949
    "OUTRAS ENTRADAS NAO ESPECIFICADAS"). O robo clica em OK (fica a linha ja selecionada) e segue (usuario, 24/11/2026)."""
    return "natureza" in texto or "outras entradas" in texto or ("find" in texto and "descri" in texto and "cancel" in texto)


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
            arquivo = open(_LOG["caminho"], "a", encoding="utf-8")  # noqa: SIM115 (fechado no fim da funcao; nem sempre e aberto aqui)
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
    esses): {sufixo: {"sucesso", "a_mao", "sem_confirmacao", "serie_ponto": bool, "quando": 'dd/mm hh:mm', "em": datetime
    do log (vai para a coluna LANCADO EM da aba Lançados)}}.
    sucesso = o robo viu "Integrado com Sucesso"; a_mao = o usuario terminou a nota numa pausa; sem_confirmacao = o
    robo integrou mas nao viu a mensagem de sucesso (pausou para conferir). Serve para conferir a planilha (usuario,
    09/11/2026). Os logs ficam 30 dias."""
    notas = {}
    # os logs com mais de 2 dias ficam em logs/arquivo-AAAAMM.zip (core.compactar_logs): le soltos e zipados
    for nome, texto in core.ler_logs("*-snac0236.log", pasta):
        quando = f"{nome[6:8]}/{nome[4:6]} {nome[9:11]}:{nome[11:13]}"
        try:
            em = datetime.strptime(nome[:13], "%Y%m%d-%H%M")
        except ValueError:
            em = None
        for linha in core.sem_acentos(texto).splitlines():  # logs antigos (sem acento) e novos (com)
            achado = _OK_NOTA.search(linha)
            if achado:
                nota = notas.setdefault(achado.group(1), {"sucesso": False, "a_mao": False, "sem_confirmacao": False,
                                                          "serie_ponto": False, "quando": quando, "em": em})
                resto = achado.group(2)
                chave = ("sucesso" if "integrada com sucesso" in resto else "a_mao" if "terminada a mao" in resto
                         else "sem_confirmacao")
                nota[chave], nota["quando"], nota["em"] = True, quando, em
                continue
            achado = _JA_EXISTE.search(linha)
            if achado:
                notas.setdefault(achado.group(1), {"sucesso": False, "a_mao": False, "sem_confirmacao": False,
                                                   "serie_ponto": False, "quando": quando, "em": em})["serie_ponto"] = True
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


class Execucao0236:
    """Uma execucao do `executar` (v1.14.0: antes uma funcao de 881 linhas com 33 funcoes internas).
    Os parametros e o estado que as etapas dividem ficam em self; cada etapa e um metodo; `rodar` e o laco."""

    def __init__(self, notas, janela_trecho, pausa, render, espera, parar, aviso, estado=None, retomar=None, intervalo=0,
             serie_ponto=None, pausar=None):
        self.notas = notas
        self.janela_trecho = janela_trecho
        self.pausa = pausa
        self.render = render
        self.espera = espera
        self.parar = parar
        self.aviso = aviso
        self.estado = estado
        self.retomar = retomar
        self.intervalo = intervalo
        self.serie_ponto = serie_ponto
        self.pausar = pausar

    def registrar(self, texto, mostrar=True):
        """Linha no log com hora e o tempo desde a linha anterior (+s), para achar o que demora."""
        with _TRAVA_LOG:
            if not self.log.closed:
                _escrever_log(texto)
        if mostrar:
            self.aviso(texto)

    def marcar(self, linhas, nome):
        for linha in linhas:
            if self.estado:
                self.estado(linha, nome)
        self.atual[:] = list(linhas) if nome == "executando" else [n for n in self.atual if n not in linhas]

    def checar_parada(self):
        if self.parar.is_set() or core.mouse_no_canto():
            raise KeyboardInterrupt

    def devagar(self, texto):
        """Modo passo a passo: anuncia a proxima acao e espera `intervalo` segundos (F8/Parar/mouse no canto valem)."""
        if self.intervalo <= 0:
            return
        self.aviso(f"Passo a passo: {texto}")
        self.registrar(f"  passo a passo: {texto}", mostrar=False)
        fim = time.time() + self.intervalo
        while time.time() < fim:
            self.checar_parada()
            time.sleep(0.1)

    def digitar(self, texto, campo):
        self.checar_parada()
        self.devagar(f"digitar '{texto}' em {NOME[campo]}" if texto else f"{NOME[campo]} fica em branco")
        if texto:
            core.send_keys(texto.replace(",", "{,}"), pause=0.03)
            time.sleep(min(self.pausa, 0.05))

    def limpar_campo(self, campo, quantos=10):
        """Apaga o conteudo do campo em foco: End + Backspace (serve com o texto selecionado ou com o cursor no meio)."""
        self.checar_parada()
        self.devagar(f"apagar o que houver em {NOME[campo]}")
        core.send_keys(f"{{END}}{{BACKSPACE {quantos}}}", pause=0.03)

    def enters(self, campo, quantos, foco):
        """Enter `quantos` vezes; cada um precisa mover o foco ou, na grade rolada, rolar a grade (senao: aviso do SRP
        ou campo preso)."""
        for vez in range(quantos):
            self.checar_parada()
            self.devagar(f"Enter {vez + 1} de {quantos} depois do {NOME[campo]}")
            antes = core.capturar(self.janela) if campo in GRADE_ROLADA and foco is not None else None
            core.send_keys("{ENTER}", pause=0.03)
            foco = self.depois_do_enter(campo, vez, foco, antes)
        return foco

    def depois_do_enter(self, campo, vez, foco, antes=None):
        """Nova caixa do foco (None = campo sem cor, ex.: lista Lemg); aviso ou foco parado = Travou. Onde o foco ja
        ficou sem cor nesta execucao (mesmo campo e mesmo Enter), espera so SEM_COR_CONHECIDO s sem ciano."""
        curto = (campo, vez) in self.sem_cor_visto
        novo = aguardar(self.janela, foco, self.render, self.fixos, sem_cor=SEM_COR_CONHECIDO if curto else SEM_COR, antes=antes)
        if novo in (None, "aviso") and campo in ("chave", "serie"):
            texto = self.janela_ja_existe()
            if texto and campo == "chave" and not self.serie_com_ponto_feita:
                raise DocumentoExiste(campo, texto)
            if texto:
                raise Travou(campo, f"o SRP diz que o Documento Recebimento já existe mesmo com '.' na Série ('{texto[:90]}')")
        if novo is None:
            # 09/11/2026 13:04: o Enter depois do Valor do ICMS ficou 4,5 s parado (nas outras notas, ate 0,7 s) e o robo
            # pausou sem saber por que. Antes de pausar espera mais uma vez (SRP lento), sem apertar nada.
            novo = aguardar(self.janela, foco, self.render, self.fixos, antes=antes)
            if novo is not None:
                self.registrar(f"  Enter {vez + 1} depois do {NOME[campo]}: o SRP demorou mais de {self.render:.0f} s", mostrar=False)
        if novo in (None, "aviso"):
            motivo = "apareceu um aviso do SRP" if novo == "aviso" else "o cursor não saiu do lugar"
            raise Travou(campo, f"{motivo} no Enter {vez + 1} depois do {NOME[campo]}{self.diagnostico(campo)}")
        if novo == "sem_cor" and not curto:
            self.sem_cor_visto.add((campo, vez))
            # lembrado tambem para a proxima execucao (24/11/2026: com 1 nota por execucao, os 4 campos sem cor da nota
            # (Enter 3 e 5 depois da Data, 9 e 10 depois do Valor do ICMS) custavam 1,2 s cada em vez de 0,4 s)
            gravar_sem_cor(self.tamanho, self.sem_cor_visto)
        novo = None if novo == "sem_cor" else novo
        self.registrar(f"  Enter {vez + 1} depois do {NOME[campo]}: foco={novo} perto de {onde_esta(novo, self.lidas)!r}",
                  mostrar=False)
        return novo

    def diagnostico(self, campo):
        """Quando o cursor prende: texto da janela do SRP por cima (barra azul), se houver, e um print da janela do
        Citrix em logs\\telas (fica so neste computador; apagado com os logs, em 30 dias). Texto para a mensagem."""
        try:
            imagem = core.capturar(self.janela)
        except Exception:
            return ""
        nome = f"{datetime.now():%Y%m%d-%H%M%S}-{campo}.png"
        try:
            pasta = core.pasta_logs() / "telas"
            pasta.mkdir(exist_ok=True)
            imagem.save(pasta / nome)
            self.registrar(f"  print da tela: logs\\telas\\{nome}", mostrar=False)
        except Exception as erro:
            self.registrar(f"  print da tela falhou: {erro}", mostrar=False)
        texto = ""
        if self.tesseract:
            try:
                barras = barras_dialogo(imagem)
                if barras:
                    texto = normalizar(" ".join(p[0] for p in core.palavras_tela(
                        imagem.crop(area_dialogo(imagem, barras)), self.tesseract, escala=4 if imagem.width < 2000 else 2)))
            except Exception as erro:
                self.registrar(f"  leitura da janela por cima falhou: {erro}", mostrar=False)
        self.registrar(f"  janela por cima: {texto[:200]!r}" if texto else "  nenhuma janela do SRP por cima", mostrar=False)
        return f" (janela do SRP na tela: '{texto[:100]}')" if texto else ""

    def local(self, caixa):
        if caixa is None:
            return " (o cursor está num campo sem cor, como uma lista)"
        texto = onde_esta(caixa, self.lidas)
        return f" (o cursor está em '{texto}')" if texto else ""

    def ir_ate(self, campo, alvo, quantos, foco, anterior):
        """Depois de digitar `campo`, Enter ate o foco chegar em `alvo` (conferido pelo rotulo lido no inicio), no maximo
        `quantos` + 4 vezes: o SRP pula ou nao campos que ele mesmo preenche (CNPJ, Filial Fiscal...). Sem rotulo para
        conferir, so conta os `quantos` Enter combinados."""
        if posicao_ok(alvo, foco, self.rotulos, anterior, self.lidas) is None:
            return self.enters(campo, quantos, foco)
        limite = quantos + 4
        for vez in range(limite):
            self.checar_parada()
            self.devagar(f"Enter {vez + 1} depois do {NOME[campo]} (procurando {NOME[alvo]}; previstos {quantos})")
            core.send_keys("{ENTER}", pause=0.03)
            foco = self.depois_do_enter(campo, vez, foco)
            if posicao_ok(alvo, foco, self.rotulos, anterior, self.lidas):
                if vez + 1 != quantos:
                    self.registrar(f"  {NOME[alvo]}: cheguei com {vez + 1} Enter (previstos {quantos})", mostrar=False)
                return foco
        raise Travou(alvo, f"não cheguei no campo {NOME[alvo]} com {limite} Enter depois do {NOME[campo]}{self.local(foco)}")

    def sair_do_nat_oper(self, foco, quantos):
        """Enter depois do Nat Oper digitado (so na OPER SPM 542). O SRP pode abrir a janela "Natureza de Operacao" (lista
        de valores) e o ciano fica parado no Nat Oper por tras dela: por isso aqui se olha a barra azul de janela a cada
        captura, em vez do `aguardar` (que so esperaria o foco andar). Janela: `fechar_natureza` (OK). Depois, se o foco
        ficou no Nat Oper, Enter ate o Qtd.Fiscal (`ir_ate`); se andou, o laco dos itens confere a coluna."""
        self.checar_parada()
        self.devagar(f"Enter depois do {NOME['nat_oper']} (pode abrir a janela 'Natureza de Operação')")
        core.send_keys("{ENTER}", pause=0.03)
        prazo = time.time() + self.render
        while True:
            self.checar_parada()
            imagem = core.capturar(self.janela)
            barras = barras_dialogo(imagem)
            if barras:
                time.sleep(0.3)  # a janela termina de desenhar antes do OCR
                imagem = core.capturar(self.janela)
                self.fechar_natureza(imagem, area_dialogo(imagem, barras_dialogo(imagem) or barras))
                break
            atual, aviso = ler_estado(self.janela, self.fixos, imagem)
            if aviso:
                raise Travou("nat_oper", f"apareceu um aviso do SRP no Enter depois do Nat Oper{self.diagnostico('nat_oper')}")
            if atual and (abs(atual[0] - foco[0]) > 5 or abs(atual[1] - foco[1]) > 5):
                self.registrar(f"  Enter depois do Nat Oper: sem janela, foco={atual}", mostrar=False)
                return atual
            if time.time() >= prazo:
                raise Travou("nat_oper", "o cursor não saiu do Nat Oper e nenhuma janela abriu depois do Enter"
                             + self.diagnostico("nat_oper"))
            time.sleep(0.05)
        novo, prazo = None, time.time() + self.render  # janela fechada: onde o foco ficou
        while novo is None and time.time() < prazo:
            self.checar_parada()
            novo = ler_estado(self.janela, self.fixos)[0]
            if novo is None:
                time.sleep(0.05)
        if novo is None:
            raise Travou("nat_oper", "depois do OK na 'Natureza de Operação' o cursor ficou num campo sem cor")
        if abs(novo[0] - foco[0]) <= 5 and abs(novo[1] - foco[1]) <= 5:  # ainda no Nat Oper
            return self.ir_ate("nat_oper", "qtd", quantos, novo, anterior=novo)
        self.registrar(f"  depois do OK na 'Natureza de Operação' o foco andou: {novo}", mostrar=False)
        return novo

    def fechar_natureza(self, imagem, area):
        """Janela "Natureza de Operacao" depois do Nat Oper: confere o texto (outra janela = pausa) e da Enter (na lista
        de valores o Enter escolhe a linha selecionada, como o OK). Espera a janela fechar. Ate a v1.16.5 procurava o
        botao OK pelo OCR (~1 s) para clicar; no SRP real (24/11/2026) o OK nunca foi lido e o Enter sempre fechou."""
        if not self.tesseract:
            raise Travou("nat_oper", "abriu uma janela depois do Nat Oper e sem o OCR não sei qual é")
        texto = normalizar(" ".join(p[0] for p in core.palavras_tela(imagem.crop(area), self.tesseract)))
        if not natureza_operacao(texto):
            raise Travou("nat_oper", f"apareceu uma janela do SRP depois do Nat Oper ('{texto[:80]}')"
                         + self.diagnostico("nat_oper"))
        self.devagar("Enter na 'Natureza de Operação' (OK)")
        core.send_keys("{ENTER}", pause=0.03)
        self.registrar(f"  Natureza de Operação: dei Enter ({texto[:60]!r})", mostrar=False)
        prazo = time.time() + self.render
        while barras_dialogo(core.capturar(self.janela)):
            self.checar_parada()
            if time.time() >= prazo:
                raise Travou("nat_oper", "a janela 'Natureza de Operação' não fechou depois do OK" + self.diagnostico("nat_oper"))
            time.sleep(0.05)

    def fechar_referencia(self, imagem, area):
        """Janela "Nota fiscal de referencia do faturamento": so OK, sem ler a nota (usuario, 05/11/2026). Clique no
        botao OK; sem o botao lido, Enter ate a janela fechar (o Enter atravessa a grade e chega no OK), no maximo 8.
        A posicao do OK fica guardada pela area da janela (`ok_referencia`, tambem no `posicoes_0236.json`): na mesma
        area clica sem o OCR do botao (~1 s, 24/11/2026) e confere que a janela fechou; nao fechou = le o botao."""
        ret = self.janela.rectangle()
        salvo = self.ok_referencia.get(tuple(area))
        if salvo:
            self.devagar("clicar em OK na 'Nota Fiscal de Referência'")
            clicar(ret.left + area[0] + (salvo[0] + salvo[2]) / 2, ret.top + area[1] + (salvo[1] + salvo[3]) / 2)
            prazo = time.time() + self.render
            while time.time() < prazo:
                self.checar_parada()
                agora = core.capturar(self.janela)
                if not barras_dialogo(agora) or tela_mudou(imagem, agora, PIXELS_JANELA):
                    self.registrar("  NF de referência: cliquei em OK (posição guardada)", mostrar=False)
                    return
                time.sleep(0.05)
            self.registrar("  NF de referência: o clique na posição guardada não fechou a janela; leio o botão", mostrar=False)
            imagem = core.capturar(self.janela)
        botao = botao_referencia(imagem, area, self.tesseract)
        if botao:
            self.devagar("clicar em OK na 'Nota Fiscal de Referência'")
            clicar(ret.left + area[0] + (botao[0] + botao[2]) / 2, ret.top + area[1] + (botao[1] + botao[3]) / 2)
            self.registrar("  NF de referência: cliquei em OK", mostrar=False)
            self.ok_referencia[tuple(area)] = tuple(botao)
            gravar_ok_referencia(self.tamanho, self.ok_referencia)
            return
        self.registrar("  NF de referência: não li o botão OK; Enter até a janela fechar", mostrar=False)
        for vez in range(8):
            self.checar_parada()
            self.devagar(f"Enter {vez + 1} na 'Nota Fiscal de Referência' (até o OK)")
            antes = core.capturar(self.janela)
            core.send_keys("{ENTER}", pause=0.03)
            self.esperar_tela(antes, 1.0)
            if not barras_dialogo(core.capturar(self.janela)):
                self.registrar(f"  NF de referência: fechou com {vez + 1} Enter", mostrar=False)
                return
        raise Travou("item", "a janela 'Nota Fiscal de Referência' não fechou (não li o botão OK e 8 Enter não bastaram)")

    def janela_ja_existe(self):
        """Texto (normalizado) da janela "Documento Recebimento ... ja existe" se ela estiver aberta; senao None. So e
        chamada quando o cursor nao anda depois da Chave ou da Serie (le o recorte da barra azul; sem barra, a tela)."""
        if not self.tesseract:
            return None
        imagem = core.capturar(self.janela)
        barras = barras_dialogo(imagem)
        alvo = imagem.crop(area_dialogo(imagem, barras)) if barras else imagem
        for escala in ((2, 4) if barras else (2,)):  # em 1366x768 o texto da janela e pequeno
            texto = normalizar(" ".join(p[0] for p in core.palavras_tela(alvo, self.tesseract, escala=escala)))
            if documento_ja_existe(texto):
                return texto
        return None

    def serie_com_ponto(self, nota, foco, texto):
        """Janela "Documento Recebimento ... ja existe, foi integrado no sistema em ..." depois da Chave (usuario,
        09/11/2026): Enter (OK), apaga a Chave, Enter ate a Serie e poe "." no fim dela ("1" -> "1."); depois segue
        com Enter ate a Filial Fisica, como sempre. Devolve o foco na Filial Fisica."""
        self.registrar(f"  janela 'já existe': {texto[:120]!r}", mostrar=False)
        self.checar_parada()
        self.devagar("Enter no aviso 'Documento Recebimento já existe' (OK)")
        antes = core.capturar(self.janela)
        core.send_keys("{ENTER}", pause=0.03)
        self.esperar_tela(antes, self.render)
        if self.janela_ja_existe():
            raise Travou("chave", "a janela 'Documento Recebimento já existe' não fechou com Enter")
        foco = foco_tela(self.janela) or foco
        if posicao_ok("chave", foco, self.rotulos, None, self.lidas) is False:
            raise Travou("chave", f"depois do OK na janela 'já existe' o cursor não voltou para a Chave{self.local(foco)}")
        self.limpar_campo("chave", 60)
        self.serie_com_ponto_feita = True
        rotulo = next((c for t, c in self.lidas if t.startswith("serie")), None)
        if not rotulo:
            raise Travou("serie", "o SRP disse que o Documento Recebimento já existe, mas não li o rótulo Série na tela")
        limite = ENTERS_CHAVE_SERIE + 4
        for vez in range(limite):
            self.checar_parada()
            self.devagar(f"Enter {vez + 1} depois de apagar a Chave (procurando a Série)")
            core.send_keys("{ENTER}", pause=0.03)
            foco = self.depois_do_enter("chave", vez, foco)
            if foco and ao_lado(foco, rotulo, self.lidas):
                break
        else:
            raise Travou("serie", f"não cheguei na Série com {limite} Enter depois de apagar a Chave{self.local(foco)}")
        self.checar_parada()
        self.devagar("ir para o fim da Série e digitar '.'")
        core.send_keys("{END}.", pause=0.03)
        time.sleep(min(self.pausa, 0.05))
        self.registrar(f"Nota ...{nota['chave'][-6:]}: o SRP disse que o Documento Recebimento já existe; apaguei a Chave e "
                  "pus '.' no fim da Série.")
        return self.ir_ate("serie", "filial1", 2, foco, anterior=foco)

    def janela_no_caminho(self, foco):
        """Depois de cada Enter do Frete ao Item: janela por cima (barra de titulo azul, ~0,07 s)? A "Nota fiscal de
        referencia do faturamento" abre nesse caminho (usuario, 05/11/2026): `fechar_referencia` e segue. Outra janela
        qualquer: pausa (nao da Enter as cegas)."""
        if not self.tesseract:
            return foco
        imagem = core.capturar(self.janela)
        barras = barras_dialogo(imagem)
        if not barras:
            return foco
        area = area_dialogo(imagem, barras)
        texto = normalizar(" ".join(p[0] for p in core.palavras_tela(imagem.crop(area), self.tesseract)))
        if not nf_referencia(texto):
            raise Travou("item", f"apareceu uma janela do SRP entre o Frete e o Item ('{texto[:60]}')")
        self.fechar_referencia(imagem, area)
        novo = aguardar(self.janela, foco, self.render, self.fixos)
        if novo in (None, "aviso"):
            raise Travou("item", "depois do OK na 'Nota Fiscal de Referência' "
                         + ("apareceu um aviso do SRP" if novo == "aviso" else "a tela não mudou"))
        self.registrar(f"  NF de referência fechada entre o Frete e o Item: foco={novo}", mostrar=False)
        return novo if isinstance(novo, tuple) else None

    def ir_para_item(self, foco):
        """Depois do Frete: Enter ate o Item da 1a linha da grade, no maximo ENTERS_FRETE_ITEM + 4 (usuario, 05/11/2026).
        De 04/11 a 05/11 o robo clicava no Item, mas o clique pulava os campos em que o SRP abre a janela "Nota fiscal de
        referencia do faturamento", que entao so aparecia depois (raramente). Depois de cada Enter, `janela_no_caminho`.
        Sem o rotulo Item lido: so a contagem."""
        conferir = posicao_ok("item", foco, self.rotulos, foco, self.lidas) is not None
        limite = ENTERS_FRETE_ITEM + 4 if conferir else ENTERS_FRETE_ITEM
        for vez in range(limite):
            self.checar_parada()
            self.devagar(f"Enter {vez + 1} depois do Frete (procurando o Item da grade)")
            core.send_keys("{ENTER}", pause=0.03)
            foco = self.janela_no_caminho(self.depois_do_enter("frete", vez, foco))
            if conferir and posicao_ok("item", foco, self.rotulos, None, self.lidas):
                self.registrar(f"  Item: cheguei com {vez + 1} Enter depois do Frete", mostrar=False)
                return foco
        if not conferir:
            return foco
        raise Travou("item", f"não cheguei no Item da grade com {limite} Enter depois do Frete{self.local(foco)}")

    def descer_para_item(self, foco, anterior):
        """Item seguinte da mesma nota. Depois do Tributos OK o SRP deixa o foco no Item da MESMA linha (05/11/2026 15:20:
        o robo digitou o 2o item por cima do 1o): seta para baixo leva ao Item da linha de baixo (sugestao do usuario).
        `anterior` = caixa do Item da linha do item anterior."""
        meio = lambda c: (c[1] + c[3]) / 2  # noqa: E731
        if (isinstance(foco, tuple) and posicao_ok("item", foco, self.rotulos, None, self.lidas) is not False
                and meio(foco) > anterior[3]):
            self.registrar(f"  depois do Tributos OK o foco já está no Item da linha de baixo: {foco}", mostrar=False)
            return foco
        if not (isinstance(foco, tuple) and posicao_ok("item", foco, self.rotulos, None, self.lidas) is not False
                and anterior[1] <= meio(foco) <= anterior[3]):
            raise Travou("item", f"depois do Tributos OK o cursor não ficou no Item da linha do item anterior{self.local(foco)}; "
                                 "não sei em que linha digitar o próximo item")
        self.checar_parada()
        self.devagar("seta para baixo: Item da próxima linha da grade")
        core.send_keys("{DOWN}", pause=0.03)
        novo = aguardar(self.janela, foco, self.render, self.fixos)
        if novo == "aviso":
            raise Travou("item", "apareceu um aviso do SRP depois da seta para baixo na grade")
        if (isinstance(novo, tuple) and posicao_ok("item", novo, self.rotulos, None, self.lidas) is not False
                and meio(novo) > anterior[3]):
            self.registrar(f"  seta para baixo: Item da linha de baixo, foco={novo}", mostrar=False)
            return novo
        raise Travou("item", f"a seta para baixo não levou ao Item da linha de baixo{self.local(novo if isinstance(novo, tuple) else None)}")

    def entrar_no_valor_unit(self, foco):
        """2o Enter depois do Qtd Fiscal: sai do Qtd.Fisica e entra no Valor Unit, que fica AMARELO (sem ciano) e faz a
        grade rolar para a direita. Conta como entrou se o ciano andar ou sumir, ou se a faixa acima do campo mudar
        (rolagem); nada disso em `render` s = Travou. Devolve o ciano que sobrar na tela (ou None)."""
        self.checar_parada()
        self.devagar("Enter 2 de 2 depois do Qtd Fiscal (entra no Valor Unit, que fica amarelo)")
        antes = core.capturar(self.janela)
        core.send_keys("{ENTER}", pause=0.03)
        prazo, sumiu = time.time() + self.render, None
        while time.time() < prazo:
            atual, tem_aviso = ler_estado(self.janela, self.fixos)
            if tem_aviso:
                raise Travou("unit", "apareceu um aviso do SRP no Enter 2 depois do Qtd Fiscal")
            andou = atual is not None and foco is not None and (abs(atual[0] - foco[0]) > 5 or abs(atual[1] - foco[1]) > 5)
            sumiu = (sumiu or time.time()) if atual is None else None
            rolou = foco is not None and mudou_acima(antes, core.capturar(self.janela), foco)
            if andou or rolou or (sumiu and time.time() - sumiu >= min(0.3, self.render / 2)):
                time.sleep(0.3)  # a grade termina de rolar
                if andou or rolou:  # sem ciano (o normal: Valor Unit amarelo) nao precisa reler
                    atual, _ = ler_estado(self.janela, self.fixos)
                self.registrar(f"  Enter 2 depois do Qtd Fiscal: no Valor Unit (andou={andou}, rolou={rolou}, "
                          f"sem ciano={bool(sumiu)}); ciano agora={atual}", mostrar=False)
                return atual
            time.sleep(0.05)
        raise Travou("unit", "o cursor não saiu do Qtd.Fisica no Enter 2 depois do Qtd Fiscal (a grade não rolou)")

    def enter_rateio(self, campo, vez, foco):
        self.checar_parada()
        antes = core.capturar(self.janela) if foco is not None else None
        core.send_keys("{ENTER}", pause=0.03)
        return self.depois_do_enter(campo, vez, foco, antes)

    def achar_botao(self, chave, antes):
        """Rele o botao (`chave` de localizar_rateio) na faixa em volta de onde estava; None se nao estiver a vista."""
        tela_agora = core.capturar(self.janela)
        esq, topo = max(antes[0] - 300, 0), max(antes[1] - 80, 0)
        recorte = tela_agora.crop((esq, topo, min(antes[2] + 300, tela_agora.width), min(antes[3] + 80, tela_agora.height)))
        caixa = localizar_rateio(core.palavras_tela(binarizar(recorte), self.tesseract)).get(chave)
        return (caixa[0] + esq, caixa[1] + topo, caixa[2] + esq, caixa[3] + topo) if caixa else None

    def confirma_terceiros(self, rateio, linha="terceiros"):
        """Le so o recorte em volta do rotulo ICMS TERCEIROS (~0,5 s, contra ~5 s da tela inteira)."""
        c = rateio["terceiros"]
        h = c[3] - c[1]
        imagem = core.capturar(self.janela)
        recorte = imagem.crop((max(c[0] - 2 * h, 0), max(c[1] - h, 0), min(c[2] + 2 * h, imagem.width),
                               min(c[3] + h, imagem.height)))
        escala = 2 if h >= 20 else 4  # em 1366x768 o rotulo tem ~8 px
        for img in (recorte, binarizar(recorte)) + ((binarizar(recorte, 120),) if escala > 2 else ()):
            texto = normalizar(" ".join(p[0] for p in core.palavras_tela(img, self.tesseract, escala=escala)))
            if TESTE_LINHA_RATEIO[linha](texto):
                return True
        self.registrar(f"  recorte do {LINHA_RATEIO[linha]} não confere: {texto!r}", mostrar=False)
        return False

    def clicar_ident_terceiros(self, pos, foco, vez):
        """Com o Rateio aberto e o foco no Ident Tributo de uma linha acima (199/198): clique no Ident Tributo da linha
        ICMS TERCEIROS, em vez dos ~6 Enter pelas linhas de cima (usuario, 04/11/2026). Foco que nao anda: segue com
        Enter, como antes."""
        t = pos["terceiros"]
        x, y = (foco[0] + foco[2]) / 2, (t[1] + t[3]) / 2
        ret = self.janela.rectangle()
        self.checar_parada()
        self.devagar("clicar no Ident Tributo da linha ICMS TERCEIROS")
        clicar(ret.left + x, ret.top + y)
        novo = aguardar(self.janela, foco, self.render, self.fixos)
        if novo == "aviso":
            raise Travou("rateio", "apareceu um aviso do SRP depois do clique no Ident Tributo do ICMS TERCEIROS")
        if isinstance(novo, tuple):
            self.registrar(f"  cliquei no Ident Tributo do ICMS TERCEIROS em ({x:.0f}, {y:.0f}) no Enter {vez + 1}: foco={novo}",
                      mostrar=False)
            return novo
        self.registrar(f"  cliquei no Ident Tributo do ICMS TERCEIROS, mas o foco não andou ({novo}): sigo com Enter", mostrar=False)
        return None if novo == "sem_cor" else foco

    def linha_dif(self, pos, foco):
        """Foco numa celula de uma linha acima da linha alvo: e a DIF. ALIQUOTAS? Le o recorte do rotulo da linha (a
        esquerda do Ident), so na altura do foco para nao pegar o rotulo das linhas vizinhas."""
        h = foco[3] - foco[1]
        imagem = core.capturar(self.janela)
        recorte = imagem.crop((max(pos["terceiros"][0] - 2 * h, 0), max(foco[1] - h // 4, 0), pos["ident"][0],
                               min(foco[3] + h // 4, imagem.height)))
        escala = 2 if h >= 20 else 4
        texto = ""
        for img in (recorte, binarizar(recorte)):
            texto = normalizar(" ".join(p[0] for p in core.palavras_tela(img, self.tesseract, escala=escala)))
            if "dif" in texto or "aliquotas" in texto:
                return True
        self.registrar(f"  linha acima da alvo no Rateio não é a DIF. ALIQUOTAS: {texto!r}", mostrar=False)
        return False

    def zerar_dif(self, r, pos, foco, vistas):
        """Sem imposto na planilha, a linha 108 DIF. ALIQUOTAS (acima da linha alvo, Ident 2, Base 300.00, Aliquota 5.0000
        vindos do SRP; usuario, 26/11/2026) precisa de Ident, Base e Aliquota 0. Devolve o novo foco ou False."""
        if tem_imposto(r) or not pos or foco is None or not (pos["ident"][3] < foco[1] and foco[3] < pos["terceiros"][1]):
            return False
        coluna = coluna_rateio(foco, pos)
        if coluna not in ("ident", "base", "aliquota") or foco[2] - foco[0] > 2 * (pos[coluna][2] - pos[coluna][0]):
            return False  # na grade principal o campo nessa altura e largo
        if foco[1] in vistas:
            return False
        vistas.add(foco[1])
        if not self.linha_dif(pos, foco):
            return False
        self.registrar(f"  linha {r['linha']}: DIF. ALIQUOTAS no Rateio (foco em {NOME_RATEIO[coluna]}): zero a linha",
                       mostrar=False)
        return self.preencher_linha(foco, pos, foco, valores_rateio(r), "DIF. ALIQUOTAS", coluna)

    def preencher_linha(self, foco, rateio, rotulo, valores, nome_linha, primeira="ident"):
        """Digita `valores` na linha do Rateio na altura de `rotulo`, do campo em foco (`primeira`) em diante, com Enter
        entre os campos. Devolve o foco no ultimo campo digitado."""
        pendentes = [c for c in COLUNAS_RATEIO if c in valores]
        pendentes = pendentes[pendentes.index(primeira):]
        # O Ident do ICMS TERCEIROS ja vem com 1 no SRP (05/11/2026; o do ICMS S/ TRANSFERENCIA com 2, 24/11/2026).
        # Chegando pelo clique, o cursor fica onde clicou e o texto nao e selecionado: digitar "1" virava "11" e dava
        # erro. Pelo Enter o texto vem selecionado. Apaga antes.
        # o SRP pode trazer valor no campo (Base 250.00 e Aliquota .0000 no ICMS S/ TRANSFERENCIA, 24/11/2026)
        self.limpar_campo(primeira, 4 if primeira == "ident" else 12)  # Ident: 1 digito; 12 cobre 1234567.89
        self.digitar(valores[primeira], primeira)
        pendentes.remove(primeira)
        for vez in range(len(COLUNAS_RATEIO) + 2):
            if not pendentes:
                break
            self.devagar(f"Enter {vez + 1} na linha {nome_linha} (procurando {NOME_RATEIO[pendentes[0]]})")
            foco = self.enter_rateio("rateio", vez, foco)
            if not na_linha(foco, rotulo):
                raise Travou("rateio", f"o cursor saiu da linha {nome_linha} antes do campo {NOME_RATEIO[pendentes[0]]}"
                                       f"{self.local(foco)}; confira os campos dessa linha")
            coluna = coluna_rateio(foco, rateio)
            self.registrar(f"  Enter {vez + 1} na linha {nome_linha}: foco={foco} coluna={coluna}", mostrar=False)
            if coluna in pendentes:
                if coluna != pendentes[0]:
                    raise Travou("rateio", f"o cursor pulou o campo {NOME_RATEIO[pendentes[0]]} da linha {nome_linha}")
                self.limpar_campo(coluna, 12)  # 20 Backspace a 0,03 s = 0,6 s por campo
                self.digitar(valores[coluna], coluna)
                pendentes.remove(coluna)
        if pendentes:
            raise Travou("rateio", f"não cheguei no campo {NOME_RATEIO[pendentes[0]]} da linha {nome_linha}")
        return foco

    def preencher_rateio(self, r, foco):
        """Depois do Valor do ICMS: Enter ate o Ident Tributo da linha ICMS TERCEIROS (lida no OCR quando o Rateio abre),
        digita os valores da linha (`valores_rateio`), clica em Tributos OK e espera o Rateio fechar."""
        if not self.tesseract:
            raise Travou("rateio", "sem o Tesseract não consigo achar a linha ICMS TERCEIROS no Rateio")
        linha_alvo = linha_rateio(r)
        nome_linha = LINHA_RATEIO[linha_alvo]
        rateio = None
        # posicoes lidas numa nota anterior (a janela do Rateio abre sempre no mesmo lugar), separadas por grade: a da 542
        # so tem a linha ICMS S/ TRANSFERENCIA, na altura da 1a linha (24/11/2026)
        salvo = self.rateio_salvo.get(linha_alvo)
        clicou = linha_alvo != "terceiros"  # o clique pula as linhas de cima da grade do ICMS TERCEIROS; na 542 nao ha
        vistas = set()  # linhas acima da alvo ja conferidas (DIF. ALIQUOTAS?)
        for vez in range(MAX_ATE_RATEIO):
            self.devagar(f"Enter {vez + 1} depois do Valor do ICMS (procurando a linha {nome_linha} do Rateio)")
            foco = self.enter_rateio("icms", vez, foco)
            pos = rateio or salvo
            novo = self.zerar_dif(r, pos, foco, vistas)
            if novo is not False:
                foco = novo
                continue
            # foco estreito na coluna Ident, abaixo do titulo e acima da linha ICMS TERCEIROS = Rateio aberto numa linha
            # de cima (na grade principal o campo nessa altura e largo: 635 px contra 107 px da celula Ident, 4K)
            if (pos and not clicou and celula_ident(foco, pos)
                    and foco[1] > pos["ident"][3] and foco[3] < pos["terceiros"][1]):
                clicou = True
                foco = self.clicar_ident_terceiros(pos, foco, vez)
            if rateio is None and salvo and na_linha(foco, salvo["terceiros"]) and celula_ident(foco, salvo):
                if self.confirma_terceiros(salvo, linha_alvo):
                    rateio = salvo
                    self.registrar(f"  Rateio: posições da nota anterior; {nome_linha} conferido no recorte (Enter {vez + 1})",
                              mostrar=False)
                else:
                    salvo = None  # le a tela inteira de novo
            if salvo and rateio is None and vez + 1 >= LER_RATEIO_DESDE + 10:
                salvo = None  # nao bateu com as posicoes salvas: le a tela inteira
            if rateio is None and not salvo and vez + 1 >= LER_RATEIO_DESDE:
                imagem = core.capturar(self.janela)
                # sem a barra azul da janela "Rateios" o Rateio ainda nao abriu: nada de ler a tela inteira (no SRP real,
                # 24/11/2026, as leituras do 8o e 9o Enter levaram 8,7 s cada e nao acharam nada; a do 10o, 1,9 s)
                if not barras_dialogo(imagem) and vez + 1 < LER_RATEIO_DESDE + 3:
                    self.registrar(f"  Enter {vez + 1}: sem a janela do Rateio na tela; não leio ainda", mostrar=False)
                    continue
                t_leitura = time.time()
                lido = ler_rateio(imagem, self.tesseract, linha_alvo)
                self.registrar(f"  leitura da tela procurando o Rateio: {time.time() - t_leitura:.1f}s", mostrar=False)
                if "terceiros" in lido and "tributos_ok" in lido:
                    rateio = lido
                    self.registrar(f"  Rateio lido no Enter {vez + 1}: {rateio}", mostrar=False)
                    faltam = [NOME_RATEIO[c] for c in COLUNAS_RATEIO if (c not in rateio and c in valores_rateio(r)) or
                              (c == "ident" and c not in rateio)]
                    if faltam:
                        raise Travou("rateio", f"não li no Rateio a(s) coluna(s) {', '.join(faltam)}")
                    self.rateio_salvo[linha_alvo] = rateio
                    gravar_posicoes(self.tamanho, rateio, linha_alvo)
                    novo = self.zerar_dif(r, rateio, foco, vistas)  # Rateio lido com o foco ja na DIF. ALIQUOTAS
                    if novo is not False:
                        foco = novo
                        continue
            if rateio and na_linha(foco, rateio["terceiros"]) and celula_ident(foco, rateio):
                self.registrar(f"  {nome_linha}: cheguei no Ident Tributo com {vez + 1} Enter depois do Valor do ICMS", mostrar=False)
                break
            linha = (rateio or salvo or {}).get("terceiros")
            if linha and foco is not None and foco[1] > linha[3] + (foco[3] - foco[1]):
                raise Travou("rateio", f"passei da linha {nome_linha} no Rateio sem parar no Ident Tributo")
        else:
            raise Travou("rateio", f"não cheguei no Ident Tributo da linha {nome_linha} com {MAX_ATE_RATEIO} Enter depois "
                         "do Valor do ICMS" + ("" if rateio else " (não achei o Rateio na tela)"))
        valores = valores_rateio(r)
        self.registrar(f"  linha {r['linha']}: {'com' if tem_imposto(r) else 'sem'} imposto na {nome_linha}: {valores}",
                       mostrar=False)
        self.preencher_linha(foco, rateio, rateio["terceiros"], valores, nome_linha)
        botao = self.achar_botao("tributos_ok", rateio["tributos_ok"])
        if not botao:
            raise Travou("rateio", "não achei o botão Tributos OK na tela")
        ret = self.janela.rectangle()
        self.checar_parada()
        self.devagar("clicar em Tributos OK")
        # com a barra azul da janela Rateios na tela, o fechamento e conferido por ela (~0,07 s por consulta) em vez de
        # procurar o botao pelo OCR a cada 0,3 s (24/11/2026: ~1,2 s depois do clique)
        com_barra = bool(barras_dialogo(core.capturar(self.janela)))
        clicar(ret.left + (botao[0] + botao[2]) / 2, ret.top + (botao[1] + botao[3]) / 2)
        self.registrar(f"  linha {r['linha']}: clique em Tributos OK", mostrar=False)
        prazo = time.time() + 2 * self.render
        while time.time() < prazo:
            self.checar_parada()
            if com_barra:
                time.sleep(0.1)
                if not barras_dialogo(core.capturar(self.janela)):
                    time.sleep(0.1)
                    return foco_tela(self.janela)
                continue
            time.sleep(0.3)
            if not self.achar_botao("tributos_ok", rateio["tributos_ok"]):
                time.sleep(0.3)
                return foco_tela(self.janela)
        raise Travou("rateio", "o Rateio não fechou depois do Tributos OK (aviso do SRP?)")

    def esperar_continuar(self):
        self.retomar.clear()
        while not self.retomar.is_set():
            if self.parar.is_set():
                raise KeyboardInterrupt
            time.sleep(0.2)

    def foco_depois_de_continuar(self):
        self.janela.set_focus()
        self.aviso("Continuando em 2 s...")
        for _ in range(20):
            self.checar_parada()
            time.sleep(0.1)
        return foco_tela(self.janela)

    def esperar_na_chave(self, texto, linhas):
        """Pausa ate o cursor estar na Chave Acesso NF-e (proxima nota) e o usuario clicar em Continuar."""
        self.registrar(texto)
        self.marcar(linhas, "pausado")
        while True:
            self.esperar_continuar()
            caixa = self.foco_depois_de_continuar()
            if caixa and posicao_ok("chave", caixa, self.rotulos, None, self.lidas) is not False:
                return caixa
            self.registrar("Ainda pausado: deixe o cursor no campo Chave Acesso NF-e (vazio) e clique em Continuar.")
            self.marcar(linhas, "pausado")

    def esperar_no_item(self, r):
        self.registrar(f"Linha {r['linha']}: o cursor não está no Item da grade. PAUSADO. Deixe o cursor no Item da próxima "
                  "linha livre da grade e clique em Continuar.")
        self.marcar([r["linha"]], "pausado")
        while True:
            self.esperar_continuar()
            caixa = self.foco_depois_de_continuar()
            if caixa and posicao_ok("item", caixa, self.rotulos, None, self.lidas) is not False:
                self.marcar([r["linha"]], "executando")
                return caixa
            self.registrar("Ainda pausado: deixe o cursor no Item da próxima linha livre da grade e clique em Continuar.")
            self.marcar([r["linha"]], "pausado")

    def localizar_integrar(self):
        if not self.tesseract:
            return None
        try:
            # le so a faixa em volta de onde o botao estava no inicio (mais rapido); sem posicao inicial, a tela toda
            tela_agora = core.capturar(self.janela)
            antes = self.rotulos.get("integrar")
            if antes:
                esq, topo = max(antes[0] - 300, 0), max(antes[1] - 80, 0)
                recorte = tela_agora.crop((esq, topo, min(antes[2] + 300, tela_agora.width), min(antes[3] + 80, tela_agora.height)))
                caixa = localizar_rotulos(core.palavras_tela(binarizar(recorte), self.tesseract)).get("integrar")
                achado = (caixa[0] + esq, caixa[1] + topo, caixa[2] + esq, caixa[3] + topo) if caixa else None
            else:
                achado = ler_rotulos(tela_agora, self.tesseract).get("integrar")
        except Exception as erro:
            self.registrar(f"  leitura do botão Integrar falhou: {erro}", mostrar=False)
            return None
        return achado

    def clicar_integrar(self, nota, vez):
        """Clica em Integrar. No 1o clique, com o botao lido no inicio e nenhuma janela por cima (barra azul), clica ali
        sem novo OCR (~0,7 s): nos logs de novembro, 242 cliques, o botao achado de novo sempre estava no mesmo lugar do
        inicio. O 2o clique (depois dos 7 Enter) sempre confere pelo OCR: uma janela sem barra azul lida (outra
        resolucao) pode ter ficado aberta, e o clique cego cairia nela. True = clicou."""
        botao = self.rotulos.get("integrar")
        if vez > 1 or not botao or barras_dialogo(core.capturar(self.janela)):
            botao = self.localizar_integrar()
        if not botao:
            return False
        ret = self.janela.rectangle()
        x, y = ret.left + (botao[0] + botao[2]) / 2, ret.top + (botao[1] + botao[3]) / 2
        self.checar_parada()
        self.devagar("clicar em Integrar" + (" de novo" if vez > 1 else ""))
        antes = core.capturar(self.janela)
        clicar(x, y)
        self.registrar(f"  nota ...{nota['chave'][-6:]}: clique {vez} em Integrar em ({x:.0f}, {y:.0f})", mostrar=False)
        prazo = time.time() + self.render  # espera a tela reagir ao clique antes do proximo passo
        while time.time() < prazo and not tela_mudou(antes, core.capturar(self.janela)):
            self.checar_parada()
            time.sleep(0.1)
        time.sleep(0.3)
        return True

    def esperar_janela(self, antes, limite):
        """Depois do 1o Integrar sem janela na tela: espera ate `limite` s uma barra azul de janela (`barras_dialogo`,
        ~0,07 s por consulta) ou uma janela conhecida lida na tela inteira (sem barra azul, outra resolucao): a cada
        `LER_JANELA_A_CADA` s, logo que a tela muda e no fim do prazo. A tela mudar sem janela nao basta (24/11/2026
        15:14: mudou 1,3 s depois do clique, sem janela; o robo voltou aos 7 Enter e nao viu o "Integrado com Sucesso").
        True = janela para ler; False = nada no prazo."""
        inicio = ultima = time.time()
        while True:
            self.checar_parada()
            agora = core.capturar(self.janela)
            passou = time.time() - inicio
            if barras_dialogo(agora):
                self.registrar(f"  janela apareceu {passou:.1f}s depois do clique no Integrar", mostrar=False)
                return True
            mudou = tela_mudou(antes, agora, PIXELS_JANELA)
            if mudou:
                antes = agora  # a proxima leitura so quando mudar de novo (ou no tempo)
                self.registrar(f"  a tela mudou {passou:.1f}s depois do clique no Integrar, sem barra azul: leio", mostrar=False)
            if mudou or time.time() - ultima >= min(LER_JANELA_A_CADA, limite) or passou >= limite:
                ultima = time.time()
                if janelas_no_texto(normalizar(" ".join(p[0] for p in core.palavras_tela(agora, self.tesseract)))):
                    self.registrar(f"  janela lida na tela inteira {passou:.1f}s depois do clique (sem barra azul)", mostrar=False)
                    return True
            if passou >= limite:
                self.registrar(f"  nenhuma janela em {limite}s depois do Integrar: sigo com os Enter", mostrar=False)
                return False
            time.sleep(0.2)

    def esperar_tela(self, antes, limite):
        """Espera a tela mudar em relacao a `antes` e parar de mudar (janela terminou de abrir/fechar), no maximo
        `limite` s. Evita ler a tela de novo (~2,5 s em 4K) antes de ela mudar."""
        prazo, anterior = time.time() + limite, None
        while time.time() < prazo:
            self.checar_parada()
            agora = core.capturar(self.janela)
            if anterior is not None and tela_mudou(antes, agora, PIXELS_JANELA) and not tela_mudou(anterior, agora, PIXELS_JANELA):
                return True
            anterior = agora
            time.sleep(0.15)
        return False

    def ajustar_acessoria(self, imagem, area):
        """Janela "Valor Calculado para Documento": le os 3 totais (varias leituras do recorte; em 1366x768 o texto e
        pequeno) e, com despesa acessoria de ate LIMITE_ACESSORIA, clica em Ajustar. Devolve None (clicou) ou o motivo
        para nao clicar."""
        recorte = imagem.crop(area)
        valores, botao = None, None
        # no print real (06/11/2026) a leitura 4x leu os 3 totais e a binarizada nao leu nada: fica por ultimo
        for img, escala in ((recorte, 2), (recorte, 4), (binarizar(recorte), 2)):
            palavras = core.palavras_tela(img, self.tesseract, escala=escala)
            valores = valores or valores_acessoria(normalizar(" ".join(p[0] for p in palavras)))
            botao = botao or botao_ajustar(frases(palavras))
            if valores and botao:
                break
        self.registrar(f"  despesa acessória: valores lidos={valores}, botão Ajustar={botao}", mostrar=False)
        if not valores:
            return "não consegui ler os totais (Geral, Documento e Acessória conferindo entre si)"
        geral, documento, acessoria = valores
        if abs(acessoria) > LIMITE_ACESSORIA + 1e-9:
            return f"despesa acessoria de {acessoria:.2f} (Geral {geral:.2f}, Documento {documento:.2f}), acima de 0,05"
        if not botao:
            return f"despesa acessória de {acessoria:.2f}, mas não achei o botão Ajustar"
        self.devagar(f"clicar em Ajustar (despesa acessória de {acessoria:.2f})")
        ret = self.janela.rectangle()
        self.checar_parada()
        clicar(ret.left + area[0] + (botao[0] + botao[2]) / 2, ret.top + area[1] + (botao[1] + botao[3]) / 2)
        self.registrar(f"  despesa acessória de {acessoria:.2f} (até 0,05): cliquei em Ajustar", mostrar=False)
        return None

    def integrar(self, nota):
        """Integrar e o que vem depois (usuario, 04/11/2026), lendo a tela antes de cada passo:
        aviso do frete alterado para CIF -> Enter (OK); "Resumo dos Tributos do Documento" -> Enter (Confirmar);
        janela "Centro de Custo e Conta Contabil para Frete LTCC" -> CLIQUE em "Frete LTCC" (Enter ali roda sem fim);
        "Valor Calculado para Documento" com despesa acessoria de ate 0,05 -> CLIQUE em Ajustar (acima: pausa);
        "Documento Integrado com Sucesso." -> Enter (OK) e acaba. Sem nenhuma dessas janelas: espera ate
        ESPERA_JANELA_INTEGRAR s a 1a janela aparecer (`esperar_janela`, 24/11/2026) e so entao Enter, ate
        ENTERS_DEPOIS_INTEGRAR, e depois Integrar de novo. Janela ja tratada que continua aberta: espera, sem Enter.
        Devolve 0 (nao achei o Integrar), 1 (nao achei o 2o Integrar), 2 (sem "Integrado com Sucesso") ou 3 (sucesso)."""
        if not self.clicar_integrar(nota, 1):
            return 0
        if not self.tesseract:
            return 1
        etiqueta = f"nota ...{nota['chave'][-6:]}"
        feitos, enters, cliques, esperas, desde_clique = [], 0, 1, 0, 0
        esperou_janela = False
        for _ in range(MAX_PASSOS_INTEGRAR):
            self.checar_parada()
            imagem = core.capturar(self.janela)
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
                janelas = janelas_no_texto(normalizar(" ".join(p[0] for p in core.palavras_tela(imagem.crop(area), self.tesseract))))
            # logo depois do 1o clique, sem barra: nada de ler a tela inteira ja (1,3 s; a janela costuma vir depois):
            # `esperar_janela` olha a barra e le a tela inteira de vez em quando
            primeira = desde_clique == 0 and (cliques > 1 or esperou_janela)
            if not janelas and (barras or not self.confia_barra or primeira or esperas >= LEITURAS_AVISO - 1):
                area = (0, 0, imagem.width, imagem.height)
                janelas = janelas_no_texto(normalizar(" ".join(p[0] for p in core.palavras_tela(imagem, self.tesseract))))
                if janelas and not barras and self.confia_barra:
                    self.confia_barra = False
                    self.registrar(f"  janela {janelas} sem barra de título azul: daqui em diante leio sempre a tela inteira",
                              mostrar=False)
            desde_clique += 1
            onde = "não li (nenhuma barra de janela)" if area is None else (
                "tela inteira" if area == (0, 0, imagem.width, imagem.height) else f"recorte {area}")
            self.registrar(f"  leitura depois do Integrar: {time.time() - t_leitura:.1f}s, {onde}, janelas={janelas}", mostrar=False)
            nova = next((nome for nome in janelas if nome not in feitos), None)
            if nova == "referencia":
                try:
                    self.fechar_referencia(imagem, area)
                except Travou as erro:
                    self.registrar(f"Nota ...{nota['chave'][-6:]}: {erro.motivo}. Termine a integração no SRP.")
                    return 2
                # no log de 05/11 16:30 a janela fechou no 7o Enter e o 2o Integrar seguiu com frete, Resumo, LTCC e
                # sucesso: depois do OK vai direto ao 2o Integrar, sem Enter as cegas
                enters = ENTERS_DEPOIS_INTEGRAR
            elif nova == "acessoria":
                problema = self.ajustar_acessoria(imagem, area)
                if problema:
                    self.registrar(f"Nota ...{nota['chave'][-6:]}: janela 'Valor Calculado para Documento': {problema}. "
                              "Confira e clique em Ajustar ou Cancelar no SRP.")
                    return 2
            elif nova == "ltcc":
                botao = botao_frete_ltcc(frases(core.palavras_tela(binarizar(imagem.crop(area)), self.tesseract)))
                if not botao:
                    self.registrar(f"Nota ...{nota['chave'][-6:]}: não achei o botão Frete LTCC. Clique em Frete LTCC no SRP.")
                    return 2
                self.devagar("clicar em Frete LTCC")
                ret = self.janela.rectangle()
                clicar(ret.left + area[0] + (botao[0] + botao[2]) / 2, ret.top + area[1] + (botao[1] + botao[3]) / 2)
            elif nova:
                self.devagar({"frete": "Enter no aviso 'frete alterado para CIF' (OK)",
                         "resumo": "Enter no 'Resumo dos Tributos do Documento' (Confirmar)",
                         "sucesso": "Enter em 'Documento Integrado com Sucesso' (OK)"}[nova])
                core.send_keys("{ENTER}", pause=0.03)
            if nova:
                self.registrar(f"  {etiqueta}: {nova} tratado", mostrar=False)
                feitos.append(nova)
                esperas = 0
                if nova == "sucesso":
                    self.registrar(f"Nota ...{nota['chave'][-6:]}: o SRP disse 'Documento Integrado com Sucesso'.")
                    self.esperar_tela(imagem, 1.5)
                    return 3
                self.esperar_tela(imagem, self.render)
                continue
            # a NF de referencia vem antes do 2o Integrar: depois dela, sem janela, segue para o 2o Integrar
            if janelas or any(f != "referencia" for f in feitos) or (cliques == 2 and enters >= ENTERS_DEPOIS_INTEGRAR):
                # janela ja tratada ainda aberta, entre uma janela e outra, ou ja integrei 2 vezes: so espera a tela
                # mudar (Enter aqui pode cair na janela seguinte; na do Frete LTCC roda sem fim)
                esperas += 1
                if esperas >= LEITURAS_AVISO:
                    break
                self.esperar_tela(imagem, self.render)
                continue
            if cliques == 1 and enters == 0 and not esperou_janela:
                # o SRP pode demorar para abrir a 1a janela (Resumo, Referencia...): Enter antes dela cai no formulario
                esperou_janela = True
                if self.esperar_janela(imagem, ESPERA_JANELA_INTEGRAR):
                    desde_clique = 0  # a proxima leitura e da tela inteira, com ou sem barra azul
                    continue
            if enters < ENTERS_DEPOIS_INTEGRAR:
                enters += 1
                self.devagar(f"Enter {enters} de {ENTERS_DEPOIS_INTEGRAR} depois do Integrar")
                core.send_keys("{ENTER}", pause=0.03)
                self.esperar_tela(imagem, 1.0)
                continue
            if not self.clicar_integrar(nota, 2):
                self.registrar(f"Nota ...{nota['chave'][-6:]}: dei {ENTERS_DEPOIS_INTEGRAR} Enter depois do Integrar, mas não "
                          "achei o Integrar de novo. Termine a integração no SRP.")
                return 1
            cliques, desde_clique = 2, 0
        self.registrar(f"Nota ...{nota['chave'][-6:]}: não vi a mensagem 'Documento Integrado com Sucesso'"
                  + (f" (janela ainda aberta: {', '.join(janelas)})" if janelas else "") + ". Confira no SRP.")
        return 2

    def ir_para_chave(self):
        """Depois do "Integrado com Sucesso" o SRP volta para a tela vazia com o foco no CNPJ Fornecedor: clica no campo
        Chave Acesso NF-e (linha de cima, a direita do rotulo lido no inicio) e confere que o foco foi para la (usuario,
        04/11/2026). Devolve o foco na Chave ou None (o robo pausa como antes)."""
        rotulo = self.rotulos.get("chave")
        if not rotulo:
            return None
        prazo, foco = time.time() + self.render, None
        while time.time() < prazo and foco is None:  # espera a tela voltar (foco no CNPJ)
            self.checar_parada()
            foco = foco_tela(self.janela)
            if foco is None:
                time.sleep(0.1)
        altura = rotulo[3] - rotulo[1]
        ret = self.janela.rectangle()
        x, y = rotulo[2] + max(40, 3 * altura), (rotulo[1] + rotulo[3]) / 2
        self.checar_parada()
        self.devagar("clicar no campo Chave Acesso NF-e (próxima nota)")
        clicar(ret.left + x, ret.top + y)
        novo = aguardar(self.janela, foco, self.render, self.fixos)
        if isinstance(novo, tuple) and posicao_ok("chave", novo, self.rotulos, None, self.lidas):
            self.registrar(f"  cliquei na Chave Acesso NF-e em ({ret.left + x:.0f}, {ret.top + y:.0f}): foco={novo}", mostrar=False)
            return novo
        self.registrar(f"  cliquei na Chave Acesso NF-e, mas o foco não foi para lá (foco={novo})", mostrar=False)
        return None

    def rodar(self):
        core.limpar_logs()
        caminho_log = core.pasta_logs() / f"{datetime.now():%Y%m%d-%H%M%S}-snac0236.log"
        self.log = open(caminho_log, "w", encoding="utf-8")  # noqa: SIM115 (aberto a execucao inteira; fechado no fim)
        with _TRAVA_LOG:
            _LOG.update(caminho=caminho_log, arquivo=self.log, ultimo=time.time())
        self.rateio_salvo = {}  # posicoes do Rateio da nota anterior (ou da execucao anterior), por linha da grade
        self.confia_barra = True  # a barra azul achou todas as janelas ate agora (senao le sempre a tela inteira)
        self.atual = []  # linhas marcadas 'executando'
        self.serie_com_ponto_feita = False  # nesta nota o robo ja pos "." na Serie (janela "ja existe")
        self.sem_cor_visto = set()  # (campo, Enter) em que o foco ficou sem cor (tambem nas anteriores): espera curta
        self.ok_referencia = {}  # area da janela NF de referencia -> botao OK (lido uma vez)


        try:
            self.tesseract = core.caminho_tesseract()
            try:
                self.janela = core.janela_citrix(self.janela_trecho)
                self.janela.set_focus()
            except LookupError:
                # SNAC0236 aberta fora do Citrix: usa a janela em que o usuario clicar durante a contagem
                self.aviso(f"Citrix não aberto: clique no campo Chave Acesso NF-e da SNAC0236. Iniciando em {self.espera}s "
                           "(mouse no canto superior esquerdo aborta).")
                inicio = time.time()
                while time.time() - inicio < self.espera:
                    if self.parar.is_set():
                        self.registrar("Interrompido antes de iniciar.")
                        return 1
                    time.sleep(0.1)
                self.janela = core.janela_ativa()
                if self.janela is None:
                    self.registrar("Não iniciei: não achei a janela da SNAC0236 em primeiro plano. Clique no campo Chave Acesso "
                                   "NF-e e inicie de novo. Nada foi digitado.")
                    return 1
                self.janela_trecho = self.janela.window_text()
                self.espera = 0  # a contagem ja passou
                self.registrar(f"  janela fora do Citrix: '{self.janela_trecho}'", mostrar=False)
            ret0 = self.janela.rectangle()
            self.tamanho = (ret0.right - ret0.left, ret0.bottom - ret0.top)
            self.rateio_salvo = ler_posicoes(self.tamanho)  # da execucao anterior: a 1a nota tambem nao le a tela inteira
            if self.rateio_salvo:
                self.registrar("  posições do Rateio da execução anterior (mesmo tamanho de janela)", mostrar=False)
            self.sem_cor_visto = ler_sem_cor(self.tamanho)
            self.ok_referencia = ler_ok_referencia(self.tamanho)
            if self.sem_cor_visto:
                self.registrar(f"  campos sem cor das execuções anteriores: {sorted(self.sem_cor_visto)}", mostrar=False)
            self.aviso(f"Clique no campo Chave Acesso NF-e. Iniciando em {self.espera}s (mouse no canto superior esquerdo aborta).")
            inicio = time.time()
            tela = core.capturar(self.janela)
            ret = self.janela.rectangle()
            problema = core.problema_tela((ret.left, ret.top, ret.right, ret.bottom), tela)
            if problema:
                self.registrar(f"Não iniciei: {problema.replace('no campo Item', 'no campo Chave Acesso NF-e')}. Nada foi digitado.")
                return 1
            try:
                self.fixos = core.caixas_cor(tela, "vermelho")
            except Exception:
                self.fixos = []
            self.rotulos, self.lidas = {}, []
            if self.tesseract:
                try:
                    self.rotulos = ler_rotulos(tela, self.tesseract, self.lidas)
                except Exception as erro:
                    self.registrar(f"  leitura dos rótulos falhou: {erro}", mostrar=False)
            self.registrar(f"  rótulos lidos: {self.rotulos}", mostrar=False)
            faltam = [NOME.get(c, c) for c in ("chave", "filial", "data", "valor_total", "frete", "item", "oper", "qtd", "unit")
                      if c not in self.rotulos]
            if faltam:
                self.aviso(f"ATENÇÃO: não li na tela os rótulos {', '.join(faltam)}: nesses campos o robô só conta os Enter, sem conferir a posição.")
            if "integrar" not in self.rotulos:
                self.aviso("ATENÇÃO: não achei o botão Integrar na tela: no fim de cada nota o robô pausa para você clicar.")
            while time.time() - inicio < self.espera:
                if self.parar.is_set():
                    self.registrar("Interrompido antes de iniciar.")
                    return 1
                time.sleep(0.1)
            foco = foco_tela(self.janela)
            if not foco:
                self.registrar("Não iniciei: nenhum campo do SRP em foco (ciano). Clique no campo Chave Acesso NF-e e inicie de novo. "
                          "Nada foi digitado.")
                return 1
            if posicao_ok("chave", foco, self.rotulos, None, self.lidas) is False:
                self.registrar("Não iniciei: o cursor não está no campo Chave Acesso NF-e. Clique nele e inicie de novo. Nada foi digitado.")
                return 1

            for numero, nota in enumerate(self.notas, start=1):
                linhas = [r["linha"] for r in nota["itens"]]
                ultima = numero == len(self.notas)
                if core.titulo_janela_ativa() and self.janela_trecho.lower() not in core.titulo_janela_ativa().lower():
                    self.registrar(f"Foco saiu do Citrix antes da nota da linha {linhas[0]}. Interrompido.")
                    return 1
                t0 = time.time()
                etiqueta = f"Nota ...{nota['chave'][-6:]} (linha{'s' if len(linhas) > 1 else ''} {linhas[0]}" + \
                           (f"-{linhas[-1]})" if len(linhas) > 1 else ")")
                self.serie_com_ponto_feita = False
                try:
                    self.marcar([linhas[0]], "executando")
                    if posicao_ok("chave", foco, self.rotulos, None, self.lidas) is False:
                        raise Travou("chave", f"o cursor não está no campo {NOME['chave']}{self.local(foco)}")
                    for indice, (campo, quantos) in enumerate(SEQ_CABECALHO):
                        if foco is None:
                            raise Travou(campo, f"não sei em que campo o cursor está antes do {NOME[campo]}{self.local(foco)}")
                        self.digitar(nota[campo], campo)
                        if campo == "frete":  # ultimo do cabecalho: Enter ate o Item da grade (janela de NF de referencia no caminho)
                            foco = self.ir_para_item(foco)
                            continue
                        try:
                            foco = self.ir_ate(campo, SEQ_CABECALHO[indice + 1][0], quantos, foco, anterior=foco)
                        except DocumentoExiste as existe:  # so no Enter depois da Chave
                            foco = self.serie_com_ponto(nota, foco, existe.motivo)
                    caixa_item = None  # Item da linha do item anterior
                    topos_linha1 = {}  # campo -> topo dele no 1o item (1a linha da grade): o titulo da coluna fica acima
                    for indice, r in enumerate(nota["itens"]):
                        self.marcar([r["linha"]], "executando")
                        if caixa_item is not None:
                            foco = self.descer_para_item(foco, caixa_item)
                        if posicao_ok("item", foco, self.rotulos, None, self.lidas) is False:
                            if self.retomar is None:
                                raise Travou("item", f"o cursor não está no Item da grade{self.local(foco)}")
                            foco = self.esperar_no_item(r)
                        caixa_item = foco
                        for posicao, (campo, quantos) in enumerate(SEQ_ITEM):
                            if foco is None and campo != "unit":  # o Valor Unit fica amarelo (conferido na entrada)
                                raise Travou(campo, f"o cursor está num campo sem cor antes do {NOME[campo]}; não digito sem ver o campo")
                            if campo != "unit" and posicao_ok(campo, foco, self.rotulos, None, self.lidas) is False:
                                raise Travou(campo, f"o cursor não está na coluna {NOME[campo]}{self.local(foco)}")
                            if campo in TITULO_COLUNA and self.tesseract:
                                if indice == 0:
                                    topos_linha1[campo] = foco[1]
                                titulo = titulo_acima(core.capturar(self.janela), foco, self.tesseract, topos_linha1.get(campo))
                                self.registrar(f"  título acima do {NOME[campo]}: {titulo!r}", mostrar=False)
                                if not any(t in titulo for t in TITULO_COLUNA[campo]):
                                    raise Travou(campo, f"o cursor não está na coluna {NOME[campo]} (título acima: '{titulo}')")
                            if campo != "nat_oper":
                                self.digitar(r[campo], campo)
                            elif r[campo]:  # OPER SPM 542: o SRP ja traz um Nat Oper; apagar e digitar o da chave
                                self.limpar_campo(campo)
                                self.digitar(r[campo], campo)
                            else:  # outra operacao: segue para o Qtd.Fiscal sem mexer no Nat Oper
                                self.devagar(f"{NOME[campo]} fica como o SRP trouxe (OPER SPM {r['oper']})")
                            alvo = SEQ_ITEM[posicao + 1][0] if posicao + 1 < len(SEQ_ITEM) else None
                            if campo == "nat_oper" and r[campo]:  # digitado: pode abrir a janela "Natureza de Operacao"
                                foco = self.sair_do_nat_oper(foco, quantos)
                            elif alvo in ("oper", "nat_oper", "qtd"):  # colunas visiveis no inicio: Enter ate a coluna certa
                                foco = self.ir_ate(campo, alvo, quantos, foco, anterior=foco)
                            elif alvo is None:  # Valor do ICMS: Enter ate a linha ICMS TERCEIROS do Rateio
                                foco = self.preencher_rateio(r, foco)
                            elif alvo == "unit":  # Enter no Qtd.Fisica (ciano), Enter no Valor Unit (amarelo, grade rola)
                                foco = self.entrar_no_valor_unit(self.enters(campo, quantos - 1, foco))
                            else:  # depois do Valor Unit a grade rola para a direita: so a contagem combinada
                                foco = self.enters(campo, quantos, foco)
                        self.registrar(f"  linha {r['linha']}: item {r['item']} digitado", mostrar=False)
                        if indice < len(nota["itens"]) - 1:
                            self.marcar([r["linha"]], "ok")
                except Travou as travou:
                    if self.retomar is None:
                        self.marcar(linhas, "erro")
                        self.registrar(f"{etiqueta}: {travou.motivo}. Interrompido. A nota ficou pela metade no SRP: termine ou "
                                  "cancele a mão.")
                        return 1
                    if ultima:
                        self.registrar(f"{etiqueta}: {travou.motivo}. PAUSADO. Termine esta nota a mão no SRP (inclusive Integrar) "
                                  "e clique em Continuar.")
                        self.marcar(linhas, "pausado")
                        self.esperar_continuar()
                    else:
                        foco = self.esperar_na_chave(
                            f"{etiqueta}: {travou.motivo}. PAUSADO. Termine esta nota a mão no SRP (inclusive Integrar), "
                            "deixe o cursor na Chave Acesso NF-e da próxima nota e clique em Continuar.", linhas)
                    self.registrar(f"OK {etiqueta}: terminada a mão.")
                    if self.serie_com_ponto_feita and self.serie_ponto:
                        self.serie_ponto(linhas)
                    self.marcar(linhas, "ok")
                    continue
                cliques = self.integrar(nota)
                clicou = cliques >= 2
                self.registrar(f"OK {etiqueta}: {len(linhas)} item(ns) em {time.time() - t0:.1f}s" +
                          ("; integrada com sucesso." if cliques == 3 else "."))
                if cliques == 0:
                    self.registrar(f"{etiqueta}: não achei o botão Integrar na tela. Clique em Integrar no SRP.")
                if self.serie_com_ponto_feita and self.serie_ponto:  # antes do "ok": a pintura da planilha usa a outra cor
                    self.serie_ponto(linhas)
                self.marcar(linhas, "ok")
                if cliques == 3:  # integrada com sucesso: clica na Chave (tambem na ultima nota, para a proxima planilha)
                    foco_chave = self.ir_para_chave()
                if ultima:
                    if not clicou and self.retomar is not None:
                        self.registrar("Depois de integrar, clique em Continuar para encerrar.")
                        self.marcar(linhas, "pausado")
                        self.esperar_continuar()
                        self.marcar(linhas, "ok")
                    break
                if cliques == 3 and foco_chave:  # integrada com sucesso e cursor na Chave: proxima nota, sem pausa
                    if self.pausar is not None and self.pausar.is_set() and self.retomar is not None:  # F9: pausa aqui, entre as notas
                        self.pausar.clear()
                        foco = self.esperar_na_chave(
                            "F9: PAUSADO entre as notas. Para seguir, deixe o cursor na Chave Acesso NF-e e clique em "
                            "Continuar.", [self.notas[numero]["itens"][0]["linha"]])
                        continue
                    foco = foco_chave
                    continue
                if self.pausar is not None:
                    self.pausar.clear()  # a pausa abaixo ja atende o F9
                if self.retomar is None:
                    self.registrar("Sem o botão Continuar o robô para aqui: confira a integração e inicie de novo pela próxima nota.")
                    return 1
                foco = self.esperar_na_chave(
                    "PAUSADO: confira a integração no SRP. Para a próxima nota, deixe o cursor na Chave Acesso NF-e e clique "
                    "em Continuar.", [self.notas[numero]["itens"][0]["linha"]])
            self.registrar("Concluído. Confira a integração da última nota no SRP.")
            return 0
        except KeyboardInterrupt:
            self.registrar("Interrompido pelo usuário.")
            return 1
        except Exception as erro:
            self.registrar(f"Erro: {erro}")
            return 1
        finally:
            if self.atual and self.estado:
                for linha in self.atual:
                    self.estado(linha, "erro")
            with _TRAVA_LOG:
                self.log.close()


def executar(notas, janela_trecho, pausa, render, espera, parar, aviso, estado=None, retomar=None, intervalo=0,
             serie_ponto=None, pausar=None):
    """Preenche as notas no SRP. `serie_ponto(linhas)`: chamado antes de marcar 'ok' a nota em que o robo pos "." na
    Serie (janela "ja existe"). `intervalo` (s): espera extra antes de cada acao (digitar, Enter, Integrar) e narra a
    acao no console, para o usuario acompanhar o percurso nos testes (0 = velocidade normal). `estado(linha, nome)` por
    linha da planilha; `retomar` (Event, botao Continuar) liga
    as pausas: depois de cada Integrar e quando algo nao confere. Sem ele, qualquer problema interrompe. `pausar` (Event,
    F9; v1.14.0): pausa entre uma nota integrada e a proxima, com o cursor na Chave. Devolve 0 se concluiu."""
    return Execucao0236(notas=notas, janela_trecho=janela_trecho, pausa=pausa, render=render, espera=espera,
                        parar=parar, aviso=aviso, estado=estado, retomar=retomar, intervalo=intervalo,
                        serie_ponto=serie_ponto, pausar=pausar).rodar()
