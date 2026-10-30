# Robô SRP

Preenche o formulário **Item Requisição Suprimento Snac** do SRP (que roda dentro do Citrix) com as linhas de uma planilha Excel: Item, C/Custo, Conta e Quantidade, linha a linha, como se fosse digitado à mão.

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

## Arquivos ao lado do exe

| Arquivo | O que é |
|---|---|
| `logs\` | um registro por execução; os de mais de 30 dias são apagados sozinhos |
| `aprendizado.json` | contas corretas informadas pelo SRP (item + C/Custo) |
| `progresso.json` | linhas concluídas da última planilha |
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
