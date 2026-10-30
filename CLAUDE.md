# RPA — robo do SRP
Preenche o formulario "Item Requisicao Suprimento Snac" do SRP com as linhas de uma planilha Excel. Idioma da UI, logs e commits: pt-BR.

## Contexto que importa
- O SRP roda dentro do **Citrix** (janela "Corporativo - Desktop Viewer"): so existe imagem, `pywinauto` NAO enxerga os campos. O robo envia teclas e le a tela por pixels.
- Campo em foco = bloco ciano puro `(0,255,255)`. Conta editavel se o novo foco comeca logo apos o C/Custo (vao < 1,5 largura); se pular para longe, a Conta esta travada e o robo vai direto para Quantidade.
- Sequencia por linha: Item, Tab, C/Custo, Tab, [Conta, Tab], Quantidade, Tab (ja desce para a proxima linha).
- Planilha: cabecalho exato `Item`, `C/Custo`, `Conta`, `Quantidade`. Valores so digitaveis; zeros a esquerda preservados se a celula for texto.

## Mapa
- `src/` — codigo (ver `src/CLAUDE.md`). `planilhas/` — planilhas de trabalho (nao vao ao git). `logs/` — um log por execucao (ignorado). `dist/`, `build/` — saida do PyInstaller (ignorados).
- `assets/` — `icone.ico` (gerado). `scripts/gerar_versao.py` — gera `build/version.txt`. `vendor/` — Tesseract portatil (ignorado). `build.ps1` — gera `dist\RPA-SRP.exe` com icone, versao e a pasta `tesseract` ao lado (fechar o app antes). `requirements.txt` — pywinauto, openpyxl, pillow, tkinterdnd2, pyinstaller.
- Teste de exe sem matar processos: `timeout 8 ./dist/RPA-SRP.exe` (saida 124 = seguiu rodando). Nao encerrar processos por nome (bloqueado e arriscado).

## Comandos (PowerShell, na pasta RPA)
- Janela: `.\.venv\Scripts\python.exe src\gui.py`
- CLI: `.\.venv\Scripts\python.exe src\preencher.py planilhas\ARQUIVO.xlsx [--dry-run] [--limite N] [--pausa S] [--render S] [--inicio LINHA]`
- Exe: `.\build.ps1`

## Regras
- Nunca digitar credenciais; sem login automatizado.
- Valida a planilha inteira antes de enviar; qualquer duvida (foco nao detectado, Conta editavel em branco, foco fora do Citrix) interrompe sem digitar.
- Abortar: botao Parar ou mouse no canto superior esquerdo da tela.
- `pyautogui` nao compila no Python 3.14: usar `pywinauto.keyboard`.
- Dados de planilhas e logs sao da empresa: nao versionar.
- Ignorar: `.venv/`, `build/`, `dist/`, `logs/`, `planilhas/*`.
