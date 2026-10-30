# src
- `core.py` — toda a logica: `ler_planilha`, `validar`, `passos`/`descrever` (teclas por linha), `maior_caixa_ciano`/`caixa_foco`/`aguardar_mudanca`/`conta_editavel` (leitura da tela), `executar(registros, janela, pausa, render, espera, parar, aviso)` (loop principal; `parar` e `threading.Event`, `aviso` recebe o andamento), `pasta_logs` (ao lado do exe quando empacotado).
- `gui.py` — janela tkinter com tema proprio (`COR`, `_estilos`): arrastar/escolher planilha (`tkinterdnd2`, opcional), chips de resumo, tabela com status por linha (Pendente/Em andamento/Concluida/Erro), barra de progresso, Simular, Iniciar, Parar, abrir pasta de logs, console. A execucao roda em thread; mensagens (`str`) e estados por linha (`(linha, nome)`) voltam por `self.fila`.
- `core.executar(..., aviso, estado=None)` — `estado(linha, 'executando'|'ok'|'erro')` alimenta a tabela. `core.py` define `sys.coinit_flags = 2` (COM STA) antes de importar pywinauto: sem isso o tkdnd falha com "unable to initialize OLE2".
- `preencher.py` — CLI fina sobre `core` (`--dry-run`, `--limite`, `--pausa`, `--render`, `--espera`, `--inicio`, `--janela`).
- `sondar.py` — lista controles UIA de uma janela (so leitura). Inutil dentro do Citrix; serve para apps Windows nativos.

## Regras
- Logica nova vai em `core.py`; `gui.py` e `preencher.py` so chamam.
- Ao mudar a deteccao de foco, testar `maior_caixa_ciano` com um print do SRP e conferir o log (`foco C/Custo=... depois=...`).
- Strings com `\n` em Python via heredoc do shell quebram: edite com a ferramenta de edicao, nao com `python - <<EOF`.
