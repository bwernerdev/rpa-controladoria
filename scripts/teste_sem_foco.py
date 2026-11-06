"""Teste: o robo conseguiria trabalhar com o Citrix SEM foco (coberto por outras janelas)?

App separado, nao mexe no robo. Dois testes:
1. Captura: fotografa a janela do Citrix coberta por outra janela, de 3 jeitos (tela como o robo faz hoje;
   PrintWindow simples; PrintWindow no modo DirectX) e mostra se sai a imagem do SRP ou preto/outra janela.
2. Teclas: com o cursor num campo de texto VAZIO do SRP e o foco em OUTRO programa, manda um digito diferente para
   cada combinacao (jeito de envio x janela do Citrix). Depois voce olha o campo: os digitos que apareceram dizem
   quais combinacoes funcionam. Apague o campo depois (nada e gravado no SRP sem voce).

3. Enter, Backspace e clique: com a SNAC0236 vazia e o foco em OUTRO programa, clica em "Outras Observacoes" sem
   mexer o mouse, digita 9, apaga (End + Backspace), da Enter e clica de volta na Chave Acesso. Confere cada passo
   sozinho pela captura (campo ciano e texto do campo). Nao grava nada no SRP.

Imagens e resultado ficam em logs/teste_sem_foco (fora do git: podem ter dados da empresa).
Rodar na pasta RPA: .\\.venv\\Scripts\\python.exe scripts\\teste_sem_foco.py
"""
import ctypes
import sys
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

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "src"))
PASTA = RAIZ / "logs" / "teste_sem_foco"
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


def tecla(hwnd, vk):
    """Tecla especial (Enter, Backspace, End...) por WM_KEYDOWN/UP com PostMessage (o jeito que funcionou no teste 2)."""
    scan = win32api.MapVirtualKey(vk, 0)
    estendida = (1 << 24) if vk in (win32con.VK_END, win32con.VK_HOME, win32con.VK_DOWN, win32con.VK_UP) else 0
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 1 | (scan << 16) | estendida)
    time.sleep(0.03)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk, 1 | (scan << 16) | estendida | (1 << 30) | (1 << 31))


def clique(hwnd, principal, x, y):
    """Clique em (x, y) da imagem da janela do Citrix, por mensagem: o mouse de verdade nao se mexe."""
    esq, topo, _, _ = win32gui.GetWindowRect(principal)
    cx, cy = win32gui.ScreenToClient(hwnd, (esq + x, topo + y))
    posicao = win32api.MAKELONG(cx, cy)
    win32api.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, posicao)
    time.sleep(0.05)
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, posicao)
    time.sleep(0.05)
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, posicao)


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
        ttk.Button(topo, text=f"3. Enter, Backspace e clique (troque o foco em {ESPERA} s)",
                   command=self.teste_navegar).pack(side="left")
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

    # ---------- 3. Enter, Backspace e clique ----------
    def teste_navegar(self):
        self.contar("Deixe a SNAC0236 VAZIA e clique em OUTRO programa (ex.: Bloco de Notas)", self._navegar)

    def _navegar(self):
        import core
        import snac0236

        principal = janela_citrix()
        if not principal:
            self.escrever("Navegar: Citrix nao encontrado.")
            return
        h_foco, titulo = self.foco_atual()
        if win32gui.GetAncestor(h_foco, win32con.GA_ROOT) == principal:
            self.escrever("Navegar: o Citrix ainda estava em foco; o teste nao vale. Troque o foco para outro programa.")
            self.status.configure(text="O Citrix estava em foco: repita trocando para outro programa.")
            return
        tela = alvos(principal).get(ALVOS["CtxICADisp"])
        if not tela:
            self.escrever("Navegar: nao achei a janela CtxICADisp do Citrix.")
            return
        self.escrever(f"--- Enter, Backspace e clique (janela em foco: {titulo!r}) ---")
        hora = datetime.now().strftime("%H%M%S")

        def foto(nome):
            imagem = capturar_printwindow(principal, 2)[0]
            imagem.save(PASTA / f"{hora}-navegar-{nome}.png")
            return imagem, core.maior_caixa_ciano(imagem)

        def esperar():
            time.sleep(1.2)

        def mudou(antes, depois, caixa):
            """% de pixels diferentes dentro do campo (texto digitado ou apagado)."""
            from PIL import ImageChops
            dif = ImageChops.difference(antes.crop(caixa).convert("L"), depois.crop(caixa).convert("L"))
            return sum(dif.histogram()[40:]) / max(dif.width * dif.height, 1)

        resultados = []

        def passo(nome, ok, detalhe):
            resultados.append(ok)
            self.escrever(f"  {'OK    ' if ok else 'FALHOU'} {nome}: {detalhe}")
            self.update()

        imagem, foco0 = foto("0-inicio")
        tesseract = core.caminho_tesseract()
        obs = chave = None
        for escala in (2, 3):  # em 1366x768 o texto e pequeno: 2a leitura mais ampliada se faltar rotulo
            rotulos = snac0236.frases(core.palavras_tela(imagem, tesseract, escala=escala)) if tesseract else []
            obs = obs or next((c for t, c in rotulos if t.startswith("outras obs")), None)
            chave = chave or next((c for t, c in rotulos if t.startswith("chave")), None)
            if obs and chave:
                break
        if not obs or not chave:
            self.escrever(f"  Nao li os rotulos 'Outras Observacoes'/'Chave Acesso' na tela (lidos: {len(rotulos)}). "
                          "A SNAC0236 esta aberta e vazia?")
            return
        alvo_obs = (obs[2] + 40, (obs[1] + obs[3]) // 2)
        alvo_chave = (chave[2] + 60, (chave[1] + chave[3]) // 2)  # comeco do campo (o meio pode ter texto)
        self.escrever(f"  inicio: foco ciano em {foco0}; Outras Observacoes em {obs}, Chave em {chave}")

        # 1. clique (sem mexer o mouse) no campo Outras Observacoes
        clique(tela, principal, *alvo_obs)
        esperar()
        antes, foco1 = foto("1-clique-obs")
        passo("clique em Outras Observacoes", bool(foco1) and foco1[0] <= alvo_obs[0] <= foco1[2]
              and foco1[1] - 4 <= alvo_obs[1] <= foco1[3] + 4, f"foco ciano em {foco1}")
        if not resultados[-1]:
            self.escrever("  Sem o clique no lugar, paro aqui (nao digito em outro campo).")
            self.status.configure(text="O clique sem mover o mouse NAO funcionou (veja o resultado).")
            return
        # 2. digita 9
        enviar(tela, JEITOS[1], "9")
        esperar()
        depois, _ = foto("2-digitou-9")
        passo("digitar 9", mudou(antes, depois, foco1) > 0.01, f"{mudou(antes, depois, foco1):.1%} do campo mudou")
        # 3. End + Backspace: o campo volta a ficar vazio
        tecla(tela, win32con.VK_END)
        tecla(tela, win32con.VK_BACK)
        esperar()
        apagado, _ = foto("3-apagou")
        passo("End + Backspace", mudou(antes, apagado, foco1) < 0.01, f"{mudou(antes, apagado, foco1):.1%} diferente do campo vazio")
        # 4. Enter: o foco sai do campo
        tecla(tela, win32con.VK_RETURN)
        esperar()
        _, foco4 = foto("4-enter")
        passo("Enter", bool(foco4) and foco4 != foco1, f"foco ciano foi de {foco1} para {foco4}")
        # 5. clique de volta na Chave Acesso
        clique(tela, principal, *alvo_chave)
        esperar()
        _, foco5 = foto("5-clique-chave")
        passo("clique na Chave Acesso", bool(foco5) and foco5[0] <= alvo_chave[0] <= foco5[2]
              and foco5[1] - 4 <= alvo_chave[1] <= foco5[3] + 4, f"foco ciano em {foco5}")
        h_foco, titulo = self.foco_atual()
        self.escrever(f"  janela em foco no fim: {titulo!r} (o Citrix nao pode ter pego o foco)")
        texto = "TUDO FUNCIONOU" if all(resultados) else f"{resultados.count(False)} passo(s) falharam"
        self.status.configure(text=f"Teste 3: {texto}. Imagens em logs\\teste_sem_foco.")
        self.escrever(f"  Teste 3: {texto}.")


if __name__ == "__main__":
    ativar_dpi()
    App().mainloop()
