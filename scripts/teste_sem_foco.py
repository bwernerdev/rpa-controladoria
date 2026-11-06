"""Teste: o robo conseguiria trabalhar com o Citrix SEM foco (coberto por outras janelas)?

App separado, nao mexe no robo. Dois testes:
1. Captura: fotografa a janela do Citrix coberta por outra janela, de 3 jeitos (tela como o robo faz hoje;
   PrintWindow simples; PrintWindow no modo DirectX) e mostra se sai a imagem do SRP ou preto/outra janela.
2. Teclas: com o cursor num campo de texto VAZIO do SRP e o foco em OUTRO programa, manda um digito diferente para
   cada combinacao (jeito de envio x janela do Citrix). Depois voce olha o campo: os digitos que apareceram dizem
   quais combinacoes funcionam. Apague o campo depois (nada e gravado no SRP sem voce).

Imagens e resultado ficam em logs/teste_sem_foco (fora do git: podem ter dados da empresa).
Rodar na pasta RPA: .\\.venv\\Scripts\\python.exe scripts\\teste_sem_foco.py
"""
import ctypes
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk

import win32api
import win32con
import win32gui
import win32ui
from PIL import Image, ImageGrab, ImageStat, ImageTk

PASTA = Path(__file__).resolve().parent.parent / "logs" / "teste_sem_foco"
TITULO_CITRIX = "Desktop Viewer"
ESPERA = 5  # segundos para cobrir o Citrix / trocar o foco
# janelas do Citrix que podem receber as teclas (classe -> nome)
ALVOS = {"CtxICADisp": "tela da sessao (CtxICADisp)", "GfxRender1": "desenho (GfxRender1)", None: "janela principal"}
JEITOS = ("WM_CHAR com PostMessage", "WM_KEYDOWN/UP com PostMessage", "WM_CHAR com SendMessage")


def ativar_dpi():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # coordenadas fisicas, como o robo
    except OSError:
        pass


def janela_citrix():
    achadas = []
    win32gui.EnumWindows(lambda h, a: a.append(h) if TITULO_CITRIX in win32gui.GetWindowText(h)
                         and win32gui.IsWindowVisible(h) else None, achadas)
    return achadas[0] if achadas else None


def alvos(principal):
    """{nome: hwnd} das janelas do Citrix que podem receber teclas."""
    filhos = []
    win32gui.EnumChildWindows(principal, lambda h, a: a.append(h) or True, filhos)
    saida = {ALVOS[None]: principal}
    for classe, nome in ALVOS.items():
        h = next((f for f in filhos if classe and win32gui.GetClassName(f) == classe), None)
        if h:
            saida[nome] = h
    return saida


def capturar_printwindow(hwnd, modo):
    esq, topo, dir_, base = win32gui.GetWindowRect(hwnd)
    largura, altura = dir_ - esq, base - topo
    dc_janela = win32gui.GetWindowDC(hwnd)
    dc = win32ui.CreateDCFromHandle(dc_janela)
    dc_mem = dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(dc, largura, altura)
    dc_mem.SelectObject(bitmap)
    ok = ctypes.windll.user32.PrintWindow(hwnd, dc_mem.GetSafeHdc(), modo)
    info, bits = bitmap.GetInfo(), bitmap.GetBitmapBits(True)
    imagem = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bits, "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bitmap.GetHandle())
    dc_mem.DeleteDC()
    dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, dc_janela)
    return imagem, bool(ok)


def capturar_tela(hwnd):
    return ImageGrab.grab(win32gui.GetWindowRect(hwnd), all_screens=True), True


def resumo(imagem):
    """Brilho medio e % de pixels quase pretos (captura que falhou costuma sair preta)."""
    cinza = imagem.convert("L")
    pretos = sum(cinza.histogram()[:12]) / (cinza.width * cinza.height)
    return f"brilho {ImageStat.Stat(cinza).mean[0]:.0f}, {pretos:.0%} preto"


def enviar(hwnd, jeito, caractere):
    if jeito == JEITOS[0]:
        win32api.PostMessage(hwnd, win32con.WM_CHAR, ord(caractere), 1)
    elif jeito == JEITOS[1]:
        vk = win32api.VkKeyScan(caractere) & 0xFF
        scan = win32api.MapVirtualKey(vk, 0)
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 1 | (scan << 16))
        time.sleep(0.03)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk, 1 | (scan << 16) | (1 << 30) | (1 << 31))
    else:
        win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(caractere), 1)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Teste: Citrix sem foco")
        self.geometry("980x640")
        self.fotos = []
        PASTA.mkdir(parents=True, exist_ok=True)
        self.arquivo = PASTA / f"resultado-{datetime.now():%Y%m%d-%H%M%S}.txt"
        topo = ttk.Frame(self, padding=10)
        topo.pack(fill="x")
        ttk.Button(topo, text=f"1. Testar captura (cubra o Citrix em {ESPERA} s)", command=self.teste_captura).pack(side="left")
        ttk.Button(topo, text=f"2. Testar teclas (troque o foco em {ESPERA} s)", command=self.teste_teclas).pack(side="left", padx=8)
        self.status = ttk.Label(self, text="Abra a SNAC0236 no Citrix.", padding=(10, 0), font=("Segoe UI", 11, "bold"))
        self.status.pack(fill="x")
        self.fotos_area = ttk.Frame(self, padding=10)
        self.fotos_area.pack(fill="x")
        self.saida = tk.Text(self, height=14, wrap="word", font=("Consolas", 9))
        self.saida.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        principal = janela_citrix()
        if not principal:
            self.escrever("Nao achei a janela do Citrix (Desktop Viewer). Abra o Citrix e reabra este teste.")
        else:
            self.escrever(f"Citrix: {win32gui.GetWindowText(principal)!r}, {win32gui.GetWindowRect(principal)}")
            for nome, h in alvos(principal).items():
                self.escrever(f"  janela: {nome} = {h} ({win32gui.GetClassName(h)})")
        self.escrever(f"Resultado gravado em {self.arquivo}")

    def escrever(self, texto):
        self.saida.insert("end", texto + "\n")
        self.saida.see("end")
        with open(self.arquivo, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%H:%M:%S} {texto}\n")

    def contar(self, texto, depois, falta=ESPERA):
        if falta:
            self.status.configure(text=f"{texto} ({falta} s)")
            self.after(1000, lambda: self.contar(texto, depois, falta - 1))
        else:
            self.status.configure(text="...")
            self.update()
            depois()

    def foco_atual(self):
        h = win32gui.GetForegroundWindow()
        return h, win32gui.GetWindowText(h)

    # ---------- 1. captura ----------
    def teste_captura(self):
        self.contar("Coloque OUTRA janela por cima do Citrix (sem minimizar o Citrix)", self._capturar)

    def _capturar(self):
        principal = janela_citrix()
        if not principal:
            self.escrever("Captura: Citrix nao encontrado.")
            return
        if win32gui.IsIconic(principal):
            self.escrever("Captura: o Citrix esta MINIMIZADO; restaure e cubra com outra janela.")
            return
        _, titulo = self.foco_atual()
        self.escrever(f"--- Captura (janela em foco: {titulo!r}) ---")
        for widget in self.fotos_area.winfo_children():
            widget.destroy()
        self.fotos = []
        hora = datetime.now().strftime("%H%M%S")
        for nome, funcao in (("tela (como o robo hoje)", lambda: capturar_tela(principal)),
                             ("PrintWindow simples", lambda: capturar_printwindow(principal, 0)),
                             ("PrintWindow modo DirectX", lambda: capturar_printwindow(principal, 2))):
            try:
                imagem, ok = funcao()
            except Exception as erro:
                self.escrever(f"  {nome}: ERRO {erro}")
                continue
            arquivo = PASTA / f"{hora}-{nome.split()[0].lower()}-{'directx' if 'DirectX' in nome else 'normal'}.png"
            imagem.save(arquivo)
            self.escrever(f"  {nome}: {imagem.size}, {resumo(imagem)}{'' if ok else ', PrintWindow devolveu falha'} -> {arquivo.name}")
            quadro = ttk.Frame(self.fotos_area)
            quadro.pack(side="left", padx=6)
            miniatura = imagem.copy()
            miniatura.thumbnail((300, 200))
            self.fotos.append(ImageTk.PhotoImage(miniatura))
            ttk.Label(quadro, image=self.fotos[-1]).pack()
            ttk.Label(quadro, text=nome).pack()
        self.status.configure(text="Captura feita: qual miniatura mostra o SRP (e nao a janela de cima ou preto)?")

    # ---------- 2. teclas ----------
    def teste_teclas(self):
        self.contar("Clique num campo de texto VAZIO do SRP e depois clique em OUTRO programa (ex.: Bloco de Notas)",
                    self._teclas)

    def _teclas(self):
        principal = janela_citrix()
        if not principal:
            self.escrever("Teclas: Citrix nao encontrado.")
            return
        h_foco, titulo = self.foco_atual()
        if win32gui.GetAncestor(h_foco, win32con.GA_ROOT) == principal:
            self.escrever("Teclas: o Citrix ainda estava em foco; o teste nao vale. Troque o foco para outro programa.")
            self.status.configure(text="O Citrix estava em foco: repita trocando para outro programa.")
            return
        self.escrever(f"--- Teclas (janela em foco: {titulo!r}) ---")
        digito = 1
        for nome, hwnd in alvos(principal).items():
            for jeito in JEITOS:
                try:
                    enviar(hwnd, jeito, str(digito))
                    self.escrever(f"  digito {digito}: {jeito} -> {nome}")
                except Exception as erro:
                    self.escrever(f"  digito {digito}: {jeito} -> {nome}: ERRO {erro}")
                digito += 1
                time.sleep(0.8)
        self.status.configure(text="Pronto. Olhe o campo do SRP: quais digitos apareceram? (depois apague o campo)")
        self.escrever("Olhe o campo do SRP e anote os digitos que apareceram; cada digito = uma combinacao acima.")


if __name__ == "__main__":
    ativar_dpi()
    App().mainloop()
