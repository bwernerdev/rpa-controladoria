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
- `src/` — codigo (ver `src/CLAUDE.md`). `tests/` — testes automatizados com SRP simulado (ver `tests/CLAUDE.md`). `planilhas/` — planilhas de trabalho (nao vao ao git, exceto `BASE PADRAO.xlsx`). `logs/` — um log por execucao (versionado; `core.limpar_logs` apaga os de mais de 30 dias a cada execucao). `build/` — trabalho do PyInstaller (ignorado); `RPA-SRP.exe` na raiz e `release/*.zip` (ignorados: vao como anexo no GitHub Releases). `progresso.json` — linhas concluidas da ultima planilha, para o app reabrir de onde parou (ignorado). `README.md` — guia para pessoas.
- `assets/` — `icone.ico` (gerado). `scripts/gerar_versao.py` — gera `build/version.txt`. `vendor/` — Tesseract portatil (versionado). `build.ps1` — gera `RPA-SRP.exe` na raiz do projeto, com icone e versao (fechar o app antes), e `release\RPA-SRP-vX.Y.Z.zip` (exe + `tesseract\` + base padrao) para outra maquina. O exe acha o Tesseract em `vendor\tesseract` ou em `tesseract\` ao lado dele. `scripts/enxugar_dll.py` e `scripts/dlls_usadas.py` — enxugam o Tesseract (ver `vendor/CLAUDE.md`). Logs e `config.json` ficam ao lado do exe. `requirements.txt` — pywinauto, openpyxl, pillow, tkinterdnd2, pyinstaller.
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
- Repositorio privado `github.com/bwernerdev/rpa-controladoria` com tudo (Tesseract, logs, `config.json`, `aprendizado.json`), exceto planilhas de trabalho, exe e zip. Cada versao: subir `VERSAO`, `.\build.ps1`, commit, push e anexar o zip num release `vX.Y.Z` (exe no git incharia o historico ~25 MB por build).
- Ignorar: `.venv/`, `__pycache__/`, `.env`, `build/`, `dist/`, `release/`, `RPA-SRP.exe`, `progresso.json`, `planilhas/*` (menos `CLAUDE.md` e `BASE PADRAO.xlsx`).
