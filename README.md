===============================================================================
SLC AGRÍCOLA - DESAFIO PRÁTICO
Analista de Suprimentos com foco em IA
===============================================================================

Aplicação local em Python que realiza alterações na planilha de Suprimentos
(base_cadastro_e_requisições*.xlsx), composta pelas abas BASE_CADASTRO_ITENS
(cadastro de itens) e BASE_RCs (requisições de compra em aberto).

A interface gráfica é construída em Tkinter com o tema moderno ttkbootstrap
(biblioteca open source, MIT), permitindo visualizar e selecionar dados das
planilhas diretamente no aplicativo - sem precisar abrir o Excel para
consultar números de RCs ou códigos de itens.

Este documento explica, nesta ordem:

   1. Como rodar o executável (a entrega principal);
   2. O que o aplicativo faz, no geral;
   3. Como rodar pelo Python (código-fonte);
   4. Como o aplicativo foi construído;
   5. Arquivos do projeto.


-------------------------------------------------------------------------------
1) COMO RODAR O EXECUTÁVEL (ENTREGA PRINCIPAL)
-------------------------------------------------------------------------------

O executável já vem pronto e não precisa de Python instalado, nem de nenhuma
outra dependência. É o arquivo SuprimentosSLC.exe (ícone da SLC).

Passo a passo:

   1. Copie o arquivo SuprimentosSLC.exe para uma pasta no computador;
   2. Coloque na MESMA pasta a planilha base (ex.:
      base_cadastro_e_requisições_1.xlsx);
   3. Dê um duplo clique no SuprimentosSLC.exe. Pronto.

Como o aplicativo encontra a planilha:

   - Ao abrir, ele procura automaticamente na própria pasta do executável um
     arquivo .xlsx/.xlsm cujo nome comece com "base" e contenha "cadastro"
     (ex.: base_cadastro_e_requisições_1.xlsx). Se houver mais de um, usa o
     primeiro em ordem alfabética;
   - Se não encontrar nenhum, a barra de status mostra "Nenhum arquivo
     carregado" e a planilha pode ser selecionada pelo botão "Procurar…" ou
     colando o caminho no campo "Arquivo Excel" e clicando em "Carregar /
     Recarregar";
   - A planilha é aberta pelo próprio aplicativo e todas as alterações são
     salvas diretamente no arquivo, na hora em que cada operação é executada.

Observações importantes:

   - Feche a planilha no Microsoft Excel antes de executar operações no
     aplicativo. Se o arquivo estiver aberto no Excel, o Windows bloqueia o
     salvamento e o aplicativo avisa (nenhuma alteração é gravada);
   - O ícone da SLC aparece no arquivo .exe (Windows Explorer), na barra de
     título da janela, no Alt+Tab e na barra de tarefas;
   - Por ser do tipo "onefile", na primeira abertura o executável leva alguns
     segundos a mais (extração temporária); nas aberturas seguintes costuma
     abrir mais rápido.


-------------------------------------------------------------------------------
2) O QUE O APLICATIVO FAZ (VISÃO GERAL)
-------------------------------------------------------------------------------

2.1) Redirecionamento de RC por exceção

O fluxo foi pensado para não precisar sair da tela:

   1. Na aba BASE_RCs da própria aplicação, filtre pela busca (ex.: digite
      parte da RC ou da descrição) e selecione uma ou VÁRIAS linhas
      (Ctrl/Shift clique, como em qualquer lista do Windows);
   2. Clique em "Adicionar RCs selecionadas" - as RCs entram na lista de
      seleção da operação 1, exibida como chips removíveis (clique no "×"
      do chip para retirar uma RC, ou em "Limpar" para zerar a seleção);
      - Duplo clique em uma linha também adiciona a RC;
      - Ou use o campo "Adicionar manualmente" digitando uma RC ou uma lista
        (separada por vírgula, ponto e vírgula ou espaço) e clicando em
        "+ Adicionar";
   3. Digite o Novo Grupo Comprador e clique em "Executar redirecionamento".

Ao executar, a aplicação:

   - altera o Grupo Comprador das RCs informadas na aba BASE_RCs;
   - NÃO altera o cadastro dos itens (é um caso de delegação por exceção);
   - registra a exceção na aba EXCECOES_RCs (RC, grupo de exceção e data/hora),
     garantindo que a informação permaneça salva após fechar o aplicativo;
   - exibe mensagem com o resultado e remove da seleção as RCs já processadas.

2.2) Correção do cadastro de um item

   1. Na aba BASE_CADASTRO_ITENS, selecione o item (busca disponível) e clique
      em "Usar item selecionado" - o código do item preenche o campo da
      operação 2 (duplo clique também funciona); ou digite o código
      manualmente;
   2. Digite o Novo Grupo Comprador e clique em "Executar correção".

Ao executar, a aplicação:

   - altera o Grupo Comprador do item na aba BASE_CADASTRO_ITENS;
   - atualiza todas as linhas da aba BASE_RCs que usam esse item;
   - NÃO altera RCs que possuem redirecionamento por exceção;
   - exibe mensagem com o resultado (incluindo as linhas preservadas por
     exceção).

2.3) Regra mais importante

   UMA EXCEÇÃO CADASTRADA PARA UMA RC SEMPRE PREVALECE SOBRE O CADASTRO DO
   ITEM.

Exemplo: se a RC 10057937 foi redirecionada por exceção para o grupo 150 e
depois o cadastro do item 7529896 (usado nessa RC) for alterado para o grupo
120, o cadastro do item passa a ser 120, mas a linha da RC 10057937 permanece
no grupo 150.

Além disso, sempre que o aplicativo carrega a planilha ele reaplica
automaticamente as exceções registradas às RCs - assim, mesmo que a planilha
base seja atualizada externamente, as RCs com exceção continuam com o grupo
comprador delegado. A aba EXCECOES_RCs (visível na aplicação) permite conferir
todas as exceções vigentes.

2.4) Validações tratadas

   - RC não encontrada na aba BASE_RCs;
   - Item não encontrado na aba BASE_CADASTRO_ITENS;
   - Grupo Comprador não informado (ou valor inválido);
   - RC que já possui exceção (o aplicativo pergunta se deseja substituir ou
     manter);
   - Valores inválidos na lista de RCs informada manualmente;
   - Planilha aberta no Excel no momento do salvamento (orienta fechar o
     arquivo).

2.5) Recursos da interface

   - Abas com as tabelas BASE_RCs, BASE_CADASTRO_ITENS, EXCECOES_RCs e
     Histórico (registro das operações da sessão);
   - Busca com filtragem instantânea em cada tabela (ignora acentos e
     maiúsculas);
   - Ordenação por qualquer coluna (clique no cabeçalho; clique novamente para
     inverter);
   - Seleção múltipla de RCs com Ctrl/Shift, igual em qualquer lista do
     Windows;
   - Barra de status com arquivo carregado e contagens (itens, linhas de RC,
     exceções);
   - Ícone da SLC na janela (barra de título, Alt+Tab e barra de tarefas).


-------------------------------------------------------------------------------
3) COMO RODAR PELO PYTHON (CÓDIGO-FONTE)
-------------------------------------------------------------------------------

Requisitos:

   - Python 3.8 ou superior;
   - Bibliotecas open source: openpyxl (Excel), ttkbootstrap (tema moderno do
     Tkinter) e Pillow (logotipo e ícone da janela). A interface usa o
     Tkinter, que já acompanha o Python.

Instalação das dependências e execução:

   pip install -r requirements.txt
   python app_suprimentos.py

Da mesma forma que no executável, mantenha a planilha base na mesma pasta do
arquivo app_suprimentos.py para o carregamento automático (ou selecione o
arquivo pelo botão "Procurar…").

Observação: feche a planilha no Microsoft Excel antes de executar operações
no aplicativo, caso contrário o Windows bloqueia o salvamento.


-------------------------------------------------------------------------------
4) COMO O APLICATIVO FOI CONSTRUÍDO
-------------------------------------------------------------------------------

Tecnologias:

   - Python 3 + Tkinter, com o tema ttkbootstrap (interface nativa de
     desktop, sem necessidade de navegador);
   - openpyxl para leitura e gravação da planilha Excel;
   - Pillow para exibir o logotipo no cabeçalho e o ícone na janela.

Organização do código (app_suprimentos.py):

   - ExcelService: camada de dados. Encapsula o acesso à planilha e as duas
     operações de negócio (redirecionar_rcs e corrigir_cadastro_item),
     incluindo o registro e a reaplicação automática das exceções;
   - Tabela: componente reutilizável de tabela (Treeview) com busca,
     ordenação por coluna e contagem de linhas;
   - ListaRcs: lista visual de "chips" removíveis das RCs selecionadas;
   - App: janela principal (montagem da interface, diálogos e ações).

Geração do executável (PyInstaller):

   - O projeto inclui o arquivo SuprimentosSLC.spec, já configurado;
   - Para gerar o executável novamente:

        pip install pyinstaller
        pyinstaller SuprimentosSLC.spec --clean --noconfirm

   - O resultado é o arquivo dist/SuprimentosSLC.exe (versão "onefile", sem
     janela de console);
   - O ícone assets/slc.ico é usado em tudo: no arquivo .exe (embutido pelo
     PyInstaller) e na janela do aplicativo, pois a pasta assets é empacotada
     dentro do executável.


-------------------------------------------------------------------------------
5) ARQUIVOS DO PROJETO
-------------------------------------------------------------------------------

   app_suprimentos.py ............... Código-fonte da aplicação
   requirements.txt ................. Dependências (openpyxl, ttkbootstrap,
                                      Pillow)
   assets/slc.ico ................... Ícone da aplicação (janela e executável)
   assets/slc_brand_name.png ........ Logotipo exibido no cabeçalho da
                                      interface
   SuprimentosSLC.spec .............. Configuração do PyInstaller
   SuprimentosSLC.exe ............... Executável (entrega principal)
   base_cadastro_e_requisições_1.xlsx  Planilha base (com as alterações
                                      realizadas em testes)
   README.md ........................ Este arquivo de instruções
