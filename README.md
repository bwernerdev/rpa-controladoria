# Robô SRP

Preenche o formulário **Item Requisição Suprimento Snac** do SRP (que roda dentro do Citrix) com as linhas de uma planilha Excel: Item, C/Custo, Conta e Quantidade, linha a linha, como se fosse digitado à mão. Na aba **SNAC0236**, lança e integra notas na **Digitação Documento Recebimento** (ver abaixo).

O robô só digita e confere. Ele nunca faz login nem corrige dado sozinho: quando o SRP recusa algo ou a quantidade na tela não bate com a planilha, ele **pausa** e espera você resolver.

## Instalar

1. Em [Releases](../../releases), baixe o `RPA-SRP-vX.Y.Z.zip` mais recente.
2. Extraia o zip (botão direito → Extrair tudo).
3. Dê duplo clique em **`instalar.cmd`**. Ele copia o app para `%LOCALAPPDATA%\Programs\RPA-SRP` (sem precisar de administrador) e cria o atalho **Robô SRP** na área de trabalho e no menu Iniciar.
4. Abra pelo atalho. A pasta extraída pode ser apagada.

**Atualizações** (a partir da v1.9.0): ao abrir, o app confere se há versão nova aqui no GitHub e pergunta se quer atualizar. Ele baixa o pacote, confere a integridade (SHA-256), fecha, troca os arquivos e abre de novo sozinho. Configurações, contas aprendidas, andamento e logs são mantidos. Durante um preenchimento ele só pergunta ao terminar. Sem internet, nada acontece.

Para atualizar à mão, faça o mesmo que na instalação com o zip novo (feche o app antes).

Sem instalar: o `RPA-SRP.exe` também roda direto da pasta extraída, desde que a pasta `tesseract\` (leitura da tela) fique ao lado dele.

Na primeira vez o Windows pode levar alguns segundos a mais para abrir (verificação do antivírus).

## Preparar a planilha

Use `planilhas\BASE PADRAO.xlsx` como modelo. A linha 1 tem que ter exatamente:

| Item | C/Custo | Conta | Quantidade |
|------|---------|-------|------------|

- Uma linha por item; linhas vazias são ignoradas.
- **Conta** pode ficar em branco quando o SRP a preenche sozinho (campo travado).
- Item, C/Custo e Conta já estão como texto no modelo, então zeros à esquerda não somem.

## Usar

1. Abra o SRP no Citrix, na tela de itens da requisição.
2. No Robô SRP, arraste a planilha para a janela (ou clique para escolher). Ele valida tudo antes e avisa:
   - erros na planilha;
   - itens repetidos (o SRP soma as quantidades de itens iguais);
   - contas que o SRP já recusou antes para o mesmo item e C/Custo.
3. **Simular** mostra o que seria digitado, sem enviar nada.
4. **Iniciar**: depois do OK você tem 5 segundos para clicar no campo **Item** da primeira linha livre do SRP.
5. Atalhos enquanto o robô roda, mesmo com o foco no Citrix:
   - **F9** pausa ao terminar a linha atual. Com ele pausado, F9 de novo continua, igual ao botão **Continuar**.
   - **F8** para.
   - Também param: o botão **Parar**, ou levar o mouse ao canto superior esquerdo da tela.

### Quando ele pausa

- **Aviso do SRP** (ex.: "Conta Contábil inválida… Conta correta=8712"): corrija o campo no SRP, dê Tab e clique em **Continuar**. Ele lê o aviso, guarda a conta correta e avisa nas próximas planilhas.
- **Quantidade diferente na tela**: confira a linha no SRP e clique em **Continuar**.

### Se fechar no meio

O app guarda as linhas já concluídas. Ao abrir de novo a mesma planilha (sem mudanças), ele oferece **Continuar da linha N**. Antes de continuar, confira no SRP se a última linha ficou pela metade.

## Aba SNAC0236 (Digitação Documento Recebimento)

Use `planilhas\BASE PADRAO SNAC0236.xlsx` como modelo: a aba **Notas** tem o cabeçalho e as colunas já formatadas (CHAVE NF como texto, data, valores, lista CIF/FOB), e a aba **Instrucoes** tem as regras e um exemplo fictício. Salve com a aba Notas selecionada: o robô lê a aba aberta.

Lança notas de recebimento a partir de uma planilha com as colunas: CHAVE NF, FILIAL FISCAL, DATA EMISSÃO, VALOR TOTAL DOCUMENTO, FRETE (CIF ou FOB), ITEM, OPER SPM, QTD FISCAL, VALOR UNIT, VALOR DO ICMS, BASE CALC. ICMS (base de cálculo da linha ICMS TERCEIROS; o nome antigo ICMS TERCEIROS (BASE DE CALCULO NF) também serve) e ALIQUOTA.

- **CHAVE NF** como texto (44 dígitos). Linhas seguidas com a mesma chave são itens da mesma nota. Nos itens seguintes da nota, a chave e o cabeçalho (FILIAL FISCAL, DATA EMISSÃO, VALOR TOTAL DOCUMENTO, FRETE) podem ficar em branco; chave em branco com o cabeçalho preenchido é erro.
- Valores com ponto decimal (ex.: `47407.28`).
- BASE CALC. ICMS e ALIQUOTA iguais a 0 = **sem imposto** (Rateio com Ident Tributo 0, Base 0 e Alíquota 0). As duas diferentes de 0 = **com imposto** (Ident Tributo 1, Base Cálculo = BASE CALC. ICMS e Alíquota = ALIQUOTA). Só uma delas zerada é erro na planilha.
- ALIQUOTA como número (12), não como porcentagem do Excel (12% chega como 0,12 e é recusado).

Para usar, deixe o cursor no campo **Chave Acesso NF-e** e clique em Iniciar. O robô:
1. preenche o cabeçalho e a grade (se o SRP disser "Documento Recebimento … já existe, foi integrado no sistema em …", dá OK, apaga a Chave e põe um ponto no fim da Série, ex.: `1` vira `1.`; se o SRP recusar de novo, pausa);
2. preenche a linha ICMS TERCEIROS do Rateio e clica em Tributos OK;
3. integra e trata as janelas do SRP: despesa acessória de até 0,05 (clica em Ajustar; acima disso, pausa para você conferir), frete alterado para CIF, Resumo dos Tributos, Frete LTCC e Integrado com Sucesso;
4. volta para a Chave da próxima nota.

Cada nota concluída fica **verde no app e na planilha**: o robô pinta de verde as linhas dela no Excel. Com a planilha fechada, ele grava o arquivo. Ao clicar em Iniciar, se a planilha carregada estiver aberta no Excel, o robô salva e fecha **só ela** (as outras planilhas e o Excel continuam abertos). Se ela tinha alterações que mudam os dados, o robô recarrega e pede para você conferir e clicar em Iniciar de novo. Se não der para fechar (ex.: célula em edição), ele avisa e pinta na janela do Excel; aí você salva. Se não conseguir pintar, só avisa nos detalhes e segue.

Nota lançada com ponto na Série (o SRP disse que o documento já existia) fica num **verde mais escuro**, com um comentário na célula da CHAVE NF explicando. As mensagens da pintura também vão para o log da execução.

**Conferir com os logs** (link no rodapé da aba): compara a planilha carregada com os logs dos últimos 30 dias e mostra nos detalhes: notas que o robô integrou ("Integrado com Sucesso") e não estão verdes, e oferece pintá-las; notas verdes sem essa confirmação (terminadas à mão ou sem a mensagem: confira no SRP); notas verdes que não aparecem nos logs; e as lançadas com ponto na Série. Os logs guardam só os 6 últimos dígitos da chave. No **Iniciar**, se alguma nota do lote já foi integrada segundo os logs e não está verde, o robô pergunta se pinta e pula essas (lançar de novo duplicaria o documento com ponto na Série).

Ao carregar a planilha, as notas que **já estão verdes** (pintadas pelo robô ou por você, em qualquer tom de verde, em todas as células preenchidas da linha) contam como feitas: aparecem como Concluídas e o robô começa pela primeira nota que não está verde ("Continuar da linha N"). Nota com só parte das linhas em verde é digitada inteira, com um aviso.

Para lançar só algumas notas, preencha **NFs a lançar** (ao lado de "Passo a passo"; só números) antes de clicar em Iniciar: com 10, o robô lança as 10 primeiras notas pendentes e para. No Iniciar seguinte ele continua pelas próximas 10. Em branco (ou 0), lança todas. O tempo previsto mostra quantas serão lançadas ("para 10 de 37 nota(s)").

Depois de importar a planilha (sem erros), abaixo da barra de progresso aparece o **tempo previsto** para as notas pendentes e o horário de término se iniciar agora. Enquanto roda, no mesmo lugar aparece a nota atual, o tempo decorrido e o **tempo previsto para a conclusão** (com o horário de término). Antes da 1ª nota a previsão usa ~40 s por nota e ~25 s por item a mais; depois, o ritmo medido. O tempo pausado não entra na conta. A confirmação do Iniciar também mostra o tempo previsto.

Se algo não confere, ele pausa: termine a nota à mão no SRP, deixe o cursor na Chave Acesso NF-e da próxima nota e clique em **Continuar**.

## Arquivos ao lado do exe

| Arquivo | O que é |
|---|---|
| `logs\` | um registro por execução; os de mais de 30 dias são apagados sozinhos |
| `aprendizado.json` | contas corretas informadas pelo SRP (item + C/Custo) |
| `progresso.json` | linhas concluídas da última planilha |
| `posicoes_0236.json` | posições do Rateio da SNAC0236 na tela (para a primeira nota também ser rápida) |
| `config.json` | preferências (tema claro/escuro) |

## Para quem mexe no código

Python 3.14 no Windows. Na pasta do projeto:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe src\gui.py                            # abre a janela
.\.venv\Scripts\python.exe -m unittest discover -s tests         # testes (SRP simulado, sem Citrix)
.\build.ps1                                                      # gera build\dist\RPA-SRP.exe e release\RPA-SRP-vX.Y.Z.zip
```

Detalhes de funcionamento e regras estão nos `CLAUDE.md` de cada pasta. O Tesseract em `vendor\tesseract` vem do instalador UB-Mannheim 5.4.0, enxugado (ver `vendor/CLAUDE.md`).
