# SLC Agrícola – Desafio Prático: Analista de Suprimentos com foco em IA

Aplicação local em Python que realiza alterações na planilha de Suprimentos
(`base_cadastro_e_requisições*.xlsx`), composta pelas abas **BASE_CADASTRO_ITENS**
(cadastro de itens) e **BASE_RCs** (requisições de compra em aberto).

A interface gráfica é construída em **Tkinter** com o tema moderno **ttkbootstrap**
(biblioteca open source, MIT), permitindo visualizar e selecionar dados das
planilhas diretamente no aplicativo — sem precisar abrir o Excel para consultar
números de RCs ou códigos de itens.

## O que a aplicação faz

### 1) Redirecionamento de RC por exceção

O fluxo foi pensado para não precisar sair da tela:

1. Na aba **BASE_RCs** da própria aplicação, filtre pela busca (ex.: digite parte
   da RC ou da descrição) e selecione uma ou **várias linhas** (Ctrl/Shift clic,
   como em qualquer lista do Windows);
2. Clique em **“＋ Adicionar RCs selecionadas”** — as RCs entram na lista de
   seleção da operação 1, exibida como *chips* removíveis (clique no **×** do chip
   para retirar uma RC, ou em **Limpar** para zerar a seleção);
   - *Duplo clique* em uma linha também adiciona a RC;
   - Ou use o campo **“Adicionar manualmente”** digitando uma RC ou uma lista
     (separada por vírgula, ponto e vírgula ou espaço) e clicando em **+ Adicionar**;
3. Digite o **Novo Grupo Comprador** e clique em **Executar redirecionamento**.

Ao executar, a aplicação:

- altera o Grupo Comprador das RCs informadas na aba **BASE_RCs**;
- **não altera o cadastro dos itens** (é um caso de delegação por exceção);
- registra a exceção na aba **EXCECOES_RCs** (RC, grupo de exceção e data/hora),
  garantindo que a informação permaneça salva após fechar o aplicativo;
- exibe mensagem com o resultado e remove da seleção as RCs já processadas.

### 2) Correção do cadastro de um item

1. Na aba **BASE_CADASTRO_ITENS**, selecione o item (busca disponível) e clique
   em **“Usar item selecionado”** — o código do item preenche o campo da operação
   2 (*duplo clique* também funciona); ou digite o código manualmente;
2. Digite o **Novo Grupo Comprador** e clique em **Executar correção**.

Ao executar, a aplicação:

- altera o Grupo Comprador do item na aba **BASE_CADASTRO_ITENS**;
- atualiza todas as linhas da aba **BASE_RCs** que usam esse item;
- **não altera RCs que possuem redirecionamento por exceção**;
- exibe mensagem com o resultado (incluindo as linhas preservadas por exceção).

### Regra mais importante
> **Uma exceção cadastrada para uma RC sempre prevalece sobre o cadastro do item.**

Exemplo: se a RC 10057937 foi redirecionada por exceção para o grupo 150 e depois o
cadastro do item 7529896 (usado nessa RC) for alterado para o grupo 120, o cadastro
do item passa a ser 120, mas a linha da RC 10057937 permanece no grupo 150.

Além disso, sempre que o aplicativo carrega a planilha ele **reaplica automaticamente**
as exceções registradas às RCs — assim, mesmo que a planilha base seja atualizada
externamente, as RCs com exceção continuam com o grupo comprador delegado. A aba
**EXCECOES_RCs** (visível na aplicação) permite conferir todas as exceções vigentes.

### Validações tratadas
- RC não encontrada na aba BASE_RCs;
- Item não encontrado na aba BASE_CADASTRO_ITENS;
- Grupo Comprador não informado (ou valor inválido);
- RC que já possui exceção (o aplicativo pergunta se deseja substituir ou manter);
- Valores inválidos na lista de RCs informada manualmente;
- Planilha aberta no Excel no momento do salvamento (orienta fechar o arquivo).

### Recursos da interface
- Abas com as tabelas **BASE_RCs**, **BASE_CADASTRO_ITENS**, **EXCECOES_RCs** e
  **Histórico** (registro das operações da sessão);
- **Busca com filtragem instantânea** em cada tabela (ignora acentos e maiúsculas);
- **Ordenação** por qualquer coluna (clique no cabeçalho; clique novamente inverte);
- Seleção múltipla de RCs com Ctrl/Shift, igual em qualquer lista do Windows;
- Barra de status com arquivo carregado e contagens (itens, linhas de RC, exceções).

## Requisitos

- Python 3.8 ou superior;
- Bibliotecas open source: `openpyxl` (Excel) e `ttkbootstrap` (tema moderno do
  Tkinter). A interface usa o Tkinter, que já acompanha o Python.

```bash
pip install -r requirements.txt
# ou: pip install openpyxl ttkbootstrap
```

## Como executar

1. Mantenha a planilha base (ex.: `base_cadastro_e_requisições_1.xlsx`) na **mesma
   pasta do aplicativo** — o aplicativo localiza o arquivo automaticamente ao abrir.
   Também é possível selecionar outro arquivo pelo botão **Procurar…** ou colar o
   caminho no campo **Arquivo Excel** e clicar em **Carregar / Recarregar**.
2. Execute:

```bash
python app_suprimentos.py
```

> **Dica:** feche a planilha no Microsoft Excel antes de executar operações no
> aplicativo, caso contrário o Windows bloqueia o salvamento.

## Arquivos do projeto

| Arquivo | Descrição |
|---|---|
| `app_suprimentos.py` | Código-fonte da aplicação (Tkinter + ttkbootstrap + openpyxl) |
| `requirements.txt` | Dependências (`openpyxl`, `ttkbootstrap`) |
| `base_cadastro_e_requisições_1.xlsx` | Planilha base (com as alterações realizadas em testes) |
| `README.md` | Este arquivo de instruções |

## Próxima etapa (fora desta versão)

- Geração do **executável** com a biblioteca **PyInstaller**, conforme solicitado na entrega.
