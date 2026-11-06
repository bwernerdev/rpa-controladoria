"""Gera planilhas/BASE PADRAO SNAC0236.xlsx: modelo da planilha da aba SNAC0236 (Digitacao Documento Recebimento).

Aba "Notas" (a ativa, a que o robo le): so o cabecalho, com as colunas ja no formato certo (CHAVE NF e FILIAL FISCAL
como texto, data dd/mm/aaaa, valores com casas decimais, FRETE com lista CIF/FOB). Aba "Instrucoes": regras e um
exemplo com dados FICTICIOS (repositorio publico: nada da empresa).

Rodar na pasta RPA: .\\.venv\\Scripts\\python.exe scripts\\gerar_base_0236.py
"""
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "src"))
import snac0236  # noqa: E402

DESTINO = RAIZ / "planilhas" / "BASE PADRAO SNAC0236.xlsx"
ORDEM = ("chave", "filial", "data", "valor_total", "frete", "item", "oper", "qtd", "unit", "icms", "icms_terceiro",
         "aliquota_icms")
FORMATOS = {"chave": "@", "filial": "@", "data": "dd/mm/yyyy", "valor_total": "0.00", "frete": "@", "item": "@",
            "oper": "@", "qtd": "0", "unit": "0.0000", "icms": "0.00", "icms_terceiro": "0.00", "aliquota_icms": "0"}
LARGURAS = {"chave": 50, "filial": 14, "data": 14, "valor_total": 24, "frete": 9, "item": 11, "oper": 11, "qtd": 12,
            "unit": 12, "icms": 15, "icms_terceiro": 17, "aliquota_icms": 11}
LINHAS_FORMATADAS = 1000

FUNDO_CAB = PatternFill("solid", fgColor="1F4E78")
FONTE_CAB = Font(bold=True, color="FFFFFF")
FINA = Side(style="thin", color="BFBFBF")

# exemplo FICTICIO: nota 1 com 1 item (com imposto); nota 2 com 2 itens (o 2o sem imposto, chave e cabecalho em branco)
EXEMPLO = [
    ("1" * 44, "10 20", date(2026, 1, 15), 10500.00, "CIF", "900135", "100", 1000, 10.5, 1260.00, 10500.00, 12),
    ("2" * 44, "10 20", date(2026, 1, 16), 6900.00, "CIF", "900135", "100", 500, 10.5, 630.00, 5250.00, 12),
    (None, None, None, None, None, "900266", "100", 200, 8.25, 0, 0, 0),
]

REGRAS = [
    "Preencha a aba Notas (o robô lê a aba que estiver aberta ao salvar: salve com a aba Notas selecionada).",
    "Linha 1 = cabeçalho. Não mude os nomes nem tire colunas.",
    "CHAVE NF: 44 dígitos, como TEXTO (a coluna já está assim). Colada como número, o Excel perde dígitos.",
    "FILIAL FISCAL: as 2 partes separadas por espaço (ex.: 10 20).",
    "DATA EMISSÃO: data (dd/mm/aaaa).",
    "FRETE: CIF ou FOB.",
    "Valores com ponto ou vírgula decimal, sem separador de milhar.",
    "Vários itens na mesma nota: um item por linha, em linhas seguidas. Nas linhas seguintes da nota, deixe a "
    "CHAVE NF e o cabeçalho (FILIAL FISCAL, DATA EMISSÃO, VALOR TOTAL DOCUMENTO, FRETE) em branco ou repita os "
    "mesmos valores.",
    "Com imposto: BASE CALC. ICMS e ALIQUOTA diferentes de 0 (Rateio: Ident Tributo 1, Base e Alíquota da planilha).",
    "Sem imposto: BASE CALC. ICMS e ALIQUOTA iguais a 0 (Rateio: Ident Tributo 0, Base 0 e Alíquota 0). "
    "Só uma das duas zerada é erro.",
    "ALIQUOTA como número (12), não como porcentagem do Excel (12% vira 0,12 e é recusado).",
    "Os dados do exemplo abaixo são fictícios.",
]


def cabecalho(ws, linha=1):
    for coluna, campo in enumerate(ORDEM, start=1):
        celula = ws.cell(row=linha, column=coluna, value=snac0236.TITULOS[campo])
        celula.fill, celula.font = FUNDO_CAB, FONTE_CAB
        celula.alignment = Alignment(horizontal="center", vertical="center")
        celula.border = Border(top=FINA, bottom=FINA, left=FINA, right=FINA)


def larguras(ws):
    for coluna, campo in enumerate(ORDEM, start=1):
        ws.column_dimensions[ws.cell(row=1, column=coluna).column_letter].width = LARGURAS[campo]


def main():
    wb = Workbook()
    notas = wb.active
    notas.title = "Notas"
    cabecalho(notas)
    larguras(notas)
    notas.freeze_panes = "A2"
    notas.row_dimensions[1].height = 22
    # formato nas celulas das linhas de dados (o Excel guarda o formato por celula; colar valores mantem o formato)
    for linha in range(2, LINHAS_FORMATADAS + 2):
        for coluna, campo in enumerate(ORDEM, start=1):
            notas.cell(row=linha, column=coluna).number_format = FORMATOS[campo]
    frete = DataValidation(type="list", formula1='"CIF,FOB"', allow_blank=True, showErrorMessage=True,
                           errorTitle="FRETE", error="Use CIF ou FOB.")
    letra = notas.cell(row=1, column=ORDEM.index("frete") + 1).column_letter
    frete.add(f"{letra}2:{letra}{LINHAS_FORMATADAS + 1}")
    notas.add_data_validation(frete)

    ajuda = wb.create_sheet("Instrucoes")
    larguras(ajuda)
    ultima = notas.cell(row=1, column=len(ORDEM)).column_letter
    ajuda.cell(row=1, column=1, value="Planilha da SNAC0236 (Digitação Documento Recebimento)").font = Font(bold=True, size=13)
    for i, texto in enumerate(REGRAS, start=3):  # cada regra numa faixa A:L (as colunas tem a largura da aba Notas)
        ajuda.merge_cells(f"A{i}:{ultima}{i}")
        ajuda.cell(row=i, column=1, value=f"•  {texto}").alignment = Alignment(wrap_text=True, vertical="top")
        ajuda.row_dimensions[i].height = 30 if len(texto) > 180 else 15
    inicio = len(REGRAS) + 5
    ajuda.cell(row=inicio - 1, column=1, value="Exemplo (fictício): nota 1 com 1 item; nota 2 com 2 itens, o 2º sem "
               "imposto (chave e cabeçalho em branco na linha do 2º item)").font = Font(bold=True)
    cabecalho(ajuda, inicio)
    for deslocamento, valores in enumerate(EXEMPLO, start=1):
        for coluna, campo in enumerate(ORDEM, start=1):
            c = ajuda.cell(row=inicio + deslocamento, column=coluna, value=valores[coluna - 1])
            c.number_format = FORMATOS[campo]
    wb.active = 0
    DESTINO.parent.mkdir(exist_ok=True)
    wb.save(DESTINO)
    print(DESTINO)


if __name__ == "__main__":
    main()
