# src
- `core.py` — toda a logica: `ler_planilha`, `validar`, `passos`/`descrever` (teclas por linha), `maior_caixa_ciano`/`caixa_foco`/`aguardar_mudanca`/`conta_editavel` (leitura da tela), `executar(registros, janela, pausa, render, espera, parar, aviso)` (loop principal; `parar` e `threading.Event`, `aviso` recebe o andamento), `pasta_logs` (ao lado do exe quando empacotado).
- `gui.py` — janela tkinter (escolher planilha, tabela de previa, Simular, Iniciar, Parar, saida). A execucao roda em thread; mensagens voltam por fila (`self.fila`).
- `preencher.py` — CLI fina sobre `core` (`--dry-run`, `--limite`, `--pausa`, `--render`, `--espera`, `--inicio`, `--janela`).
- `sondar.py` — lista controles UIA de uma janela (so leitura). Inutil dentro do Citrix; serve para apps Windows nativos.

## Regras
- Logica nova vai em `core.py`; `gui.py` e `preencher.py` so chamam.
- Ao mudar a deteccao de foco, testar `maior_caixa_ciano` com um print do SRP e conferir o log (`foco C/Custo=... depois=...`).
- Strings com `\n` em Python via heredoc do shell quebram: edite com a ferramenta de edicao, nao com `python - <<EOF`.
