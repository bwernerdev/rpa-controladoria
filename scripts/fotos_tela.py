"""Fotografa a janela do app nas etapas principais, nos dois temas, para conferir o visual depois de mexer na tela.

Monta o App com config, progresso, contas aprendidas e Citrix FALSOS (nada toca no config.json, no progresso.json,
nos logs nem no aprendizado.json reais) e com planilhas FICTICIAS (item 900135, C/Custo 7300, conta 8712; chaves
111...1 e 222...2). Cada foto e pedida a propria janela (PrintWindow): nada do resto da tela entra.

Rodar na pasta RPA:  .\\.venv\\Scripts\\python.exe scripts\\fotos_tela.py [pasta]   (padrao: build\\fotos, fora do git)
"""
import ctypes
import sys
import tempfile
import time
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "src"))

from openpyxl import Workbook  # noqa: E402
from PIL import Image  # noqa: E402

import gui  # noqa: E402
import snac0236  # noqa: E402

ORDEM_0236 = ("chave", "filial", "data", "valor_total", "frete", "item", "oper", "qtd", "unit", "icms", "icms_terceiro",
              "aliquota_icms")


def planilha_0234(pasta):
    livro = Workbook()
    aba = livro.active
    aba.append(["Item", "C/Custo", "Conta", "Quantidade"])
    for i in range(12):
        aba.append([str(900135 + i), "7300", "8712", 2 + i])
    caminho = pasta / "REQUISICAO FICTICIA.xlsx"
    livro.save(caminho)
    return str(caminho)


def planilha_0236(pasta):
    livro = Workbook()
    aba = livro.active
    aba.append([snac0236.TITULOS[c] for c in ORDEM_0236])
    linhas = [("1" * 44, "10 20", date(2026, 1, 15), 10500.00, "CIF", "900135", "100", 1000, 10.5, 1260.00, 10500.00, 12),
              ("2" * 44, "10 20", date(2026, 1, 16), 6900.00, "CIF", "900135", "100", 500, 10.5, 630.00, 5250.00, 12),
              (None, None, None, None, None, "900266", "100", 200, 8.25, 0, 0, 0)]
    for linha in linhas:
        aba.append(list(linha))
    for celula in aba["A"][1:]:
        celula.number_format = "@"
    caminho = pasta / "NOTAS FICTICIAS.xlsx"
    livro.save(caminho)
    return str(caminho)


def falsos(config):
    """Troca o que grava ou le arquivos reais por falsos; devolve a funcao que desfaz."""
    trocas = {
        (gui, "ler_config"): lambda: dict(config),
        (gui, "gravar_config"): lambda dados: None,
        (gui, "gravar_progresso"): lambda *a, **k: None,
        (gui, "ler_progresso"): lambda *a, **k: (set(), None),
        (gui, "contas_suspeitas"): lambda *a, **k: [],
        (gui, "citrix_aberto"): lambda: True,
    }
    originais = {chave: getattr(*chave) for chave in trocas}
    for (alvo, nome), valor in trocas.items():
        setattr(alvo, nome, valor)
    return lambda: [setattr(alvo, nome, valor) for (alvo, nome), valor in originais.items()]


def esperar(app, segundos=0.6):
    fim = time.time() + segundos
    while time.time() < fim:
        app.update()
        time.sleep(0.02)


def foto(janela, destino):
    """Imagem SO da janela, pedida a ela mesma (PrintWindow, area interna): o que estiver por cima na tela (outra
    janela, previa da barra de tarefas, notificacao) nao entra. Recortar a tela pegou uma previa de outro app."""
    import win32gui
    import win32ui

    janela.update()
    hwnd = ctypes.windll.user32.GetParent(janela.winfo_id()) or janela.winfo_id()
    esq, topo, dir_, base = win32gui.GetClientRect(hwnd)
    largura, altura = dir_ - esq, base - topo
    dc_janela = win32gui.GetWindowDC(hwnd)
    dc = win32ui.CreateDCFromHandle(dc_janela)
    memoria = dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(dc, largura, altura)
    memoria.SelectObject(bitmap)
    try:
        if not ctypes.windll.user32.PrintWindow(hwnd, memoria.GetSafeHdc(), 1 | 2):  # CLIENTONLY | RENDERFULLCONTENT
            raise RuntimeError("PrintWindow falhou")
        info = bitmap.GetInfo()
        Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bitmap.GetBitmapBits(True), "raw", "BGRX", 0,
                         1).save(destino)
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memoria.DeleteDC()
        dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, dc_janela)
    print(destino)


def fotografar(tema, pasta, planilhas):
    desfazer = falsos({"tema": tema, "aba": "snac0234"})
    try:
        app = gui.App()  # o App chama ativar_dpi() antes do Tk: as coordenadas do ImageGrab ficam certas
        app.attributes("-topmost", True)
        esperar(app, 1.0)
        foto(app, pasta / f"{tema}-0234-vazia.png")
        app.carregar(planilhas["0234"])
        esperar(app)
        foto(app, pasta / f"{tema}-0234-carregada.png")
        app.simular()
        esperar(app)
        foto(app, pasta / f"{tema}-0234-simular.png")
        app.abrir_opcoes()
        esperar(app)
        opcoes = [w for w in app.winfo_children() if isinstance(w, gui.tk.Toplevel)][-1]
        opcoes.attributes("-topmost", True)
        esperar(app, 0.3)
        foto(opcoes, pasta / f"{tema}-opcoes.png")
        opcoes.event_generate("<Escape>")
        app.trocar_aba("snac0236")
        esperar(app)
        foto(app, pasta / f"{tema}-0236-vazia.png")
        app.p0236.carregar(planilhas["0236"])
        esperar(app)
        foto(app, pasta / f"{tema}-0236-carregada.png")
        app.destroy()
    finally:
        desfazer()


def main():
    pasta = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "build" / "fotos"
    pasta.mkdir(parents=True, exist_ok=True)
    temporaria = Path(tempfile.mkdtemp())
    planilhas = {"0234": planilha_0234(temporaria), "0236": planilha_0236(temporaria)}
    for tema in ("claro", "escuro"):
        fotografar(tema, pasta, planilhas)


if __name__ == "__main__":
    main()
