# Robô SRP
Preenche o formulário **Item Requisição Suprimento Snac** do SRP (dentro do Citrix) com as linhas de uma planilha Excel (Item, C/Custo, Conta, Quantidade), como se fosse digitado à mão. Na aba **SNAC0236**, lança e integra notas na **Digitação Documento Recebimento**.
Só digita e confere: nunca faz login nem corrige dado sozinho. Se o SRP recusa algo ou a quantidade não bate, **pausa** e espera você.

## Instalar
1. Em [Releases](../../releases), baixe o `RPA-SRP-vX.Y.Z.zip` mais recente e extraia.
2. Duplo clique em **`instalar.cmd`**: copia para `%LOCALAPPDATA%\Programs\RPA-SRP` (sem administrador) e cria o atalho **Robô SRP** (desktop e menu Iniciar). A pasta extraída pode ser apagada.
3. Sem instalar: `RPA-SRP.exe` roda da pasta extraída, com `tesseract\` ao lado.

**Atualização** (desde a v1.9.0): ao abrir e a cada hora, o app confere o GitHub e pergunta se quer atualizar (baixa, confere SHA-256, troca os arquivos e reabre). Mantém config, contas aprendidas, andamento e logs. Durante um preenchimento, só pergunta ao terminar. Conferir na hora: clique na versão no rodapé. Resultado em Detalhes e `logs\atualizacao.log`. Até a v1.12.18 não há fallback de proxy: em rede com proxy automático, atualizar à mão uma vez (instalar com o zip novo, app fechado).

## Preparar a planilha
Modelo: `planilhas\BASE PADRAO.xlsx`. Linha 1 exata: `Item | C/Custo | Conta | Quantidade`.
- Uma linha por item; linhas vazias são ignoradas.
- **Conta** pode ficar em branco quando o SRP a preenche (campo travado).
- Item, C/Custo e Conta são texto no modelo: zeros à esquerda ficam.

## Usar
1. Abra o SRP no Citrix, na tela de itens da requisição.
2. Arraste a planilha para o Robô SRP. Ele valida tudo e avisa: erros, itens repetidos (o SRP soma) e contas já recusadas para o mesmo item e C/Custo.
3. **Simular** mostra o que seria digitado, sem enviar.
4. **Iniciar**: depois do OK, 5 s para clicar no campo **Item** da primeira linha livre.
5. Atalhos (valem com o foco no Citrix): **F9** pausa ao fim da linha (de novo, continua); **F8** ou **Parar** ou mouse no canto superior esquerdo: para.

**Quando pausa:**
- *Aviso do SRP* (ex.: "Conta Contábil inválida… Conta correta=8712"): corrija no SRP, Tab, **Continuar**. Ele guarda a conta correta e avisa nas próximas planilhas.
- *Quantidade diferente*: confira a linha no SRP, **Continuar**.

**Se fechar no meio:** ao reabrir a mesma planilha (sem mudanças), oferece **Continuar da linha N**. Confira antes se a última linha ficou pela metade.

## Aba SNAC0236 (Digitação Documento Recebimento)
Modelo: `planilhas\BASE PADRAO SNAC0236.xlsx` (aba **Notas** com colunas formatadas; aba **Instrucoes** com regras e exemplo fictício). Salve com a aba Notas selecionada: o robô lê a aba aberta.

Colunas: CHAVE NF, FILIAL FISCAL, DATA EMISSÃO, VALOR TOTAL DOCUMENTO, FRETE (CIF/FOB), ITEM, OPER SPM, QTD FISCAL, VALOR UNIT, VALOR DO ICMS, BASE CALC. ICMS (o nome antigo `ICMS TERCEIROS (BASE DE CALCULO NF)` também serve), ALIQUOTA; NAT OPER opcional.
- **CHAVE NF** em texto (44 dígitos). Linhas seguidas com a mesma chave = itens da mesma nota; nos itens seguintes, chave e cabeçalho podem ficar em branco (chave em branco com cabeçalho preenchido é erro).
- Valores com ponto decimal (`47407.28`). ALIQUOTA como número (12), não como % do Excel (12% vira 0,12 e é recusado).
- BASE CALC. ICMS e ALIQUOTA: ambas 0 = **sem imposto** (Ident Tributo 0, Base 0, Alíquota 0); ambas diferentes de 0 = **com imposto** (Ident Tributo 1). Só uma zerada é erro.
- **Nat Oper:** só na OPER SPM 542 o robô apaga o campo e digita 1949 (emitente de SC: chave começando com 42) ou 2949 (fora de SC). Nas outras operações não mexe. A janela "Natureza de Operação" recebe OK; outra janela pausa. Se a planilha traz NAT OPER num item da 542, tem que bater com a chave (senão, erro ao carregar). No Rateio da 542 ele preenche a linha ICMS S/ TRANSFERENCIA (não a ICMS TERCEIROS).

**Uso:** cursor no campo **Chave Acesso NF-e** e Iniciar. O robô:
1. Preenche cabeçalho e grade. Se o SRP disser "Documento Recebimento … já existe", dá OK, apaga a Chave e põe um ponto no fim da Série (`1` vira `1.`); recusando de novo, pausa.
2. Preenche a linha ICMS TERCEIROS do Rateio e clica em Tributos OK.
3. Integra e trata as janelas: despesa acessória até 0,05 (Ajustar; acima, pausa), frete alterado para CIF, Resumo dos Tributos, Frete LTCC, Integrado com Sucesso.
4. Volta à Chave da próxima nota.

**Verde = concluída.** O robô pinta as linhas da nota de verde no Excel (no app também). Com a planilha aberta no Excel, ao Iniciar ele salva e fecha **só ela**; se ela tinha mudanças nos dados, recarrega e pede para você conferir e iniciar de novo. Se não puder fechar (ex.: célula em edição), pinta na janela do Excel e você salva; se nem isso, avisa nos detalhes e segue. Nota lançada com ponto na Série fica num **verde mais escuro**, com comentário na célula da CHAVE NF. Mensagens da pintura também vão para o log.
- Ao carregar, notas **já verdes** (por você ou pelo robô, qualquer tom, em todas as células preenchidas) contam como feitas e ele começa pela primeira não verde. Nota com só parte das linhas verdes é digitada inteira, com aviso.
- **NFs a lançar** (ao lado de "Passo a passo"): com 10, lança as 10 primeiras pendentes e o próximo Iniciar segue com as 10 seguintes. Vazio ou 0 = todas.
- **Conferir com os logs** (rodapé da aba): compara a planilha com os logs dos últimos 30 dias (guardam só os 6 últimos dígitos da chave) e mostra: integradas e não verdes (oferece pintar), verdes sem confirmação (confira no SRP), verdes ausentes dos logs e as com ponto na Série. No Iniciar, notas já integradas segundo os logs e não verdes: pergunta se pinta e pula (relançar duplicaria o documento).
- **Tempo previsto** após importar sem erros (e durante a execução, com horário de término): ~40 s por nota + ~25 s por item extra antes da 1ª nota; depois, o ritmo medido. Tempo pausado não conta.
- Sem planilha: área de arrastar, **Escolher planilha**, **Abrir a última** e as colunas exigidas. O cartão mostra concluídas e pendentes. Detalhes ficam recolhidos (**Detalhes ▸**) e abrem sozinhos com erro ou aviso. Durante o preenchimento: só **Parar** e, na pausa, **Continuar**.

Se algo não confere, pausa: termine a nota à mão no SRP, deixe o cursor na Chave da próxima e clique em **Continuar**.

## Arquivos ao lado do exe
| Arquivo | O que é |
|---|---|
| `logs\` | um registro por execução; após 2 dias vão para `arquivo-AAAAMM.zip`; mais de 30 dias são apagados; prints de `logs\telas` ficam 7 dias |
| `aprendizado.json` | contas corretas informadas pelo SRP (item + C/Custo) |
| `progresso.json` | linhas concluídas da última planilha |
| `posicoes_0236.json` | posições do Rateio da SNAC0236 na tela |
| `config.json` | preferências (tema) |

## Para quem mexe no código
Python 3.14 no Windows, na pasta do projeto:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe src\gui.py                       # janela
.\.venv\Scripts\python.exe -m unittest discover -s tests    # testes (SRP simulado)
.\build.ps1                                                 # build\dist\RPA-SRP.exe + release\RPA-SRP-vX.Y.Z.zip
```

Regras e funcionamento: `CLAUDE.md` de cada pasta. Tesseract em `vendor\tesseract` (UB-Mannheim 5.4.0, enxugado; ver `vendor/CLAUDE.md`).
