# RPA — robo do SRP
Preenche o formulario "Item Requisicao Suprimento Snac" do SRP com as linhas de uma planilha Excel. Idioma da UI, logs e commits: pt-BR.

## Contexto que importa
- O SRP roda dentro do **Citrix** (janela "Corporativo - Desktop Viewer"): so existe imagem, `pywinauto` NAO enxerga os campos. O robo envia teclas e le a tela por pixels.
- Campo em foco = bloco ciano puro `(0,255,255)`. Conta editavel se o novo foco comeca logo apos o C/Custo (vao < 1,5 largura); se pular para longe, a Conta esta travada e o robo vai direto para Quantidade.
- Sequencia por linha: Item, Tab, [C/Custo, Tab], [Conta, Tab], Quantidade, Tab (ja desce para a proxima linha). C/Custo e Conta podem vir travados: o foco pula o campo.
- Qual campo esta em foco: no inicio o Tesseract le o cabecalho da grade (`localizar_colunas`) e cada caixa ciano e classificada pela coluna mais proxima (`qual_coluna`). Sem cabecalho: C/Custo sempre digitado e Conta pela regra do vao (`conta_editavel`).
- Aviso do SRP (ex.: "Conta Contabil invalida ... Conta correta=8712") trava o foco: o robo PAUSA a linha (estado 'pausado'), o usuario corrige no SRP, da Tab e clica em Continuar; o robo retoma do campo onde o cursor estiver (`esperar_correcao`). Cursor no Item = linha completada a mao. Sem cabecalho ou na linha de comando: interrompe.
- Na pausa por aviso, o robo le o texto do aviso (`ler_aviso`, localizado pelo icone vermelho, 5 leituras com voto) e, se houver "Conta correta=N" do mesmo item, guarda em `aprendizado.json` (item|C/Custo -> conta); ao carregar a planilha, avisa as linhas com conta ja recusada (`contas_suspeitas`). Nada e corrigido sozinho.
- Quantidade lida diferente da planilha: pausa (Continuar) em vez de interromper; o usuario confere a linha no SRP.
- O SRP soma itens repetidos na mesma linha: a conferencia da quantidade da alarme falso nesses itens (o app avisa ao carregar).
- Planilha: cabecalho exato `Item`, `C/Custo`, `Conta`, `Quantidade`. Valores so digitaveis; zeros a esquerda preservados se a celula for texto.

## Mapa
- `src/` — codigo (ver `src/CLAUDE.md`). `tests/` — testes automatizados com SRP simulado (ver `tests/CLAUDE.md`). `planilhas/` — planilhas de trabalho (nao vao ao git). `logs/` — um log por execucao (ignorado). `build/` — trabalho do PyInstaller; `RPA-SRP.exe` na raiz (ignorados).
- `assets/` — `icone.ico` (gerado). `scripts/gerar_versao.py` — gera `build/version.txt`. `vendor/` — Tesseract portatil (ignorado). `build.ps1` — gera `RPA-SRP.exe` na raiz do projeto, com icone e versao (fechar o app antes). O exe acha o Tesseract em `vendor\tesseract`; para outra maquina, levar o exe + a pasta como `tesseract\` ao lado. Logs e `config.json` ficam ao lado do exe. `requirements.txt` — pywinauto, openpyxl, pillow, tkinterdnd2, pyinstaller.
- Teste de exe sem matar processos: `timeout 8 ./RPA-SRP.exe` (saida 124 = seguiu rodando). Nao encerrar processos por nome (bloqueado e arriscado).

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
- Ignorar: `.venv/`, `build/`, `logs/`, `planilhas/*`, `vendor/tesseract/`, `RPA-SRP.exe`, `config.json`, `aprendizado.json`.
