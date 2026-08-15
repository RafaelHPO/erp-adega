# ERP VBA + MariaDB - Sistema de Gestao Comercial

Versao publica e sanitizada de um sistema ERP desenvolvido em Microsoft Excel, VBA, UserForms, Power Query e MariaDB. O projeto nasceu para uma operacao de adega, mas a arquitetura pode ser adaptada para outros pequenos negocios que precisam controlar vendas, estoque, compras, caixa, financeiro, usuarios, fornecedores e indicadores.

Esta pasta foi preparada para GitHub/LinkedIn: nao inclui dados reais, backups brutos, planilhas originais, instaladores, credenciais, IPs privados ou arquivos binarios dos UserForms.

## Visao Geral

O sistema cobre o ciclo operacional principal de uma loja:

- autenticacao de usuarios;
- abertura, sangria e fechamento de caixa;
- cadastro de produtos, fornecedores, categorias, grupos, combos e usuarios;
- registro de compras e entradas de produtos;
- conferencia de itens importados de XML/NFC-e;
- vendas, itens de venda, pagamentos, pendencias e cancelamentos;
- contas a pagar, contas a receber e fluxo de caixa;
- movimentacao e acerto de estoque;
- logs operacionais;
- dashboard gerencial com Power Query e modelo de dados;
- ETL com Power Query e rotinas VBA para processar arquivos XML/NFC-e.

## Arquitetura

Usuario
  |
  v
Excel + VBA UserForms
  |-- formularios operacionais
  |-- modulos VBA de apoio
  |-- rotinas ETL
  |
  v
ADODB + MySQL ODBC
  |
  v
MariaDB
  |-- tabelas relacionais
  |-- stored procedures
  |-- regras de estoque, venda, caixa, financeiro e licenca
  |
  v
Dashboard / Power Query / Modelo de Dados


## Tecnologias

- Microsoft Excel com macros (`.xlsm`) no ambiente privado original
- VBA
- VBA UserForms
- ADODB
- MySQL ODBC 9.7 Unicode Driver
- MariaDB
- SQL e stored procedures
- Power Query
- Modelo de Dados do Excel
- Google Drive para sincronizacao/backup no ambiente original
- Tailscale para acesso remoto ao banco em alguns cenarios
- RustDesk para suporte remoto em implantacoes assistidas

## Estrutura Publica

PUBLICAR_GITHUB/
|-- README.md
|-- .gitignore
|-- docs/
|   |-- PUBLICACAO.md
|-- sql/
|   |-- schema.sql
|   |-- procedures.sql
|-- vba/
|   |-- forms/
|   |   |-- frmLogin.frm
|   |   |-- frmPrincipal.frm
|   |   |-- frmVendas.frm
|   |   |-- frmPedido.frm
|   |   |-- demais UserForms em texto
|   |-- modules/
|       |-- ModConexao.bas
|       |-- modSistema.bas
|       |-- ModComboBox.bas
|       |-- ModETL*.bas
|-- screenshots/

## O Que Foi Sanitizado

- Os arquivos `.xlsm` originais nao foram incluidos.
- Os arquivos `.frx` foram omitidos por conterem recursos binarios/imagens.
- Backups SQL completos com dados reais nao foram incluidos.
- Os blocos `INSERT INTO` foram removidos dos scripts SQL publicos.
- Instaladores `.exe` e `.msi` nao foram incluidos.
- Senhas, IPs, usuarios reais e dados de conexao foram mascarados.
- Arquivos reais de XML/NFC-e nao foram incluidos.

## Banco de Dados

A pasta `sql/` contem uma versao publica da estrutura MariaDB:

- `schema.sql`: estrutura das tabelas, chaves e relacionamentos, sem dados reais.
- `procedures.sql`: stored procedures usadas pelo VBA, sem definers reais e sem carga de dados.

### Tabelas Principais

| Tabela | Finalidade |
|---|---|
| `tab_caixa` | Controle de abertura, fundo, sangria, fechamento e status do caixa |
| `tab_categorias` | Categorias de produtos e markup |
| `tab_compras` | Cabecalho das compras/notas de entrada |
| `tab_itenscompra` | Itens vinculados a compras |
| `tab_fornecedores` | Cadastro de fornecedores |
| `tab_produtos` | Cadastro principal de produtos, codigos, precos, custos, estoque e status |
| `tab_movimentoestoque` | Historico de entradas, saidas, acertos e movimentacoes |
| `tab_vendas` | Cabecalho das vendas |
| `tab_itensvenda` | Itens vendidos por venda |
| `tab_pagamentos` | Pagamentos recebidos por forma de pagamento |
| `tab_contasapagar` | Contas a pagar |
| `tab_contasareceber` | Contas a receber |
| `tab_grupos` | Grupos de produtos |
| `tab_produtosgrupo` | Produtos vinculados a grupos |
| `tab_gruposcombo` | Grupos vinculados a combos |
| `tab_itenscombo` | Itens que compoem combos |
| `tab_usuarios` | Usuarios do sistema |
| `tab_config` | Configuracoes gerais, versao e parametros |
| `tab_logeventos` | Registro de eventos operacionais |
| `temp_produtos` | Estrutura temporaria/apoio para carga de produtos |

### Stored Procedures

Procedures documentadas na versao publica:

PROC_ABRIRCAIXA
PROC_ABRIRVENDA
PROC_ACERTOESTOQUE
PROC_ADCITEMGRUPO
PROC_ADCITENSCOMBO
PROC_ADCITENSVENDA
PROC_ATIVARLICENCA
PROC_BAIXARPAGAMENTOS
PROC_CADASTRARCOMBO
PROC_CADASTRARFORNECEDOR
PROC_CADASTRARGRUPO
PROC_CADASTRARPRODUTO
PROC_CADASTRARUSUARIO
PROC_CALCULAPRECOVENDA
PROC_CALCULARESTOQUEMINIMO
PROC_CANCELARENTRADA
PROC_CANCELARVENDA
PROC_EDITARCOMBO
PROC_ENTRADAPRODUTOS
PROC_ESTORNAR
PROC_EXCLUIRGRUPO
PROC_FECHARCAIXA
PROC_FECHARENTRADA
PROC_FECHARVENDA
PROC_LOGIN
PROC_REGISTRARCOMPRA
PROC_REMOVEITEMCOMBO
PROC_REMOVEITEMGRUPO
PROC_RESUMOCAIXA
PROC_RETIRARITENSCOMPRA
PROC_RETIRARITENSVENDA
PROC_SALVARCOMBO
PROC_SANGRIACAIXA
PROC_VERIFICARLICENCA

### Areas Cobertas Pelas Procedures

- autenticacao e licenca: `PROC_LOGIN`, `PROC_VERIFICARLICENCA`, `PROC_ATIVARLICENCA`;
- caixa: `PROC_ABRIRCAIXA`, `PROC_SANGRIACAIXA`, `PROC_FECHARCAIXA`, `PROC_RESUMOCAIXA`;
- vendas: `PROC_ABRIRVENDA`, `PROC_ADCITENSVENDA`, `PROC_RETIRARITENSVENDA`, `PROC_FECHARVENDA`, `PROC_CANCELARVENDA`;
- pagamentos: `PROC_BAIXARPAGAMENTOS`, `PROC_ESTORNAR`;
- compras e entradas: `PROC_REGISTRARCOMPRA`, `PROC_ENTRADAPRODUTOS`, `PROC_RETIRARITENSCOMPRA`, `PROC_FECHARENTRADA`, `PROC_CANCELARENTRADA`;
- produtos e estoque: `PROC_CADASTRARPRODUTO`, `PROC_ACERTOESTOQUE`, `PROC_CALCULAPRECOVENDA`, `PROC_CALCULARESTOQUEMINIMO`;
- fornecedores, usuarios, grupos e combos: procedures de cadastro, edicao, exclusao e vinculacao.

## Modulos VBA

| Modulo | Responsabilidade |
|---|---|
| `ModConexao.bas` | Cria conexao ADODB com MariaDB via MySQL ODBC, testa conexao e verifica licenca |
| `modSistema.bas` | Inicializacao, configuracoes, formatacoes, modo admin, fechamento do sistema, senha de caixa e resumo |
| `ModUsuarioLogado.bas` | Variaveis globais do usuario autenticado |
| `ModCaixaAberto.bas` | Consulta e validacao de caixa aberto |
| `ModComboBox.bas` | Carregamento de combos de produtos, categorias, unidades, setores, fornecedores e status |
| `modNz.bas` | Funcoes auxiliares para nulos, numeros, datas e montagem segura de valores SQL |
| `modGlob.bas` | Utilitarios globais, como verificacao de campo em recordset |
| `ModETLbase.bas` | Abertura, fechamento, carga e exibicao de arquivos usados no ETL |
| `ModETLestrutura.bas` | Criacao da aba de apoio para ETL |
| `ModETLcabecalho.bas` | Extracao de cabecalho de nota: CNPJ, numero, serie, emissao, valor, pagamento e EAN |
| `ModETLitens.bas` | Extracao e normalizacao de itens, quantidades, unidades, codigos, descricoes e valores |
| `ModETLprocess.bas` | Orquestracao principal do processamento de NFC-e |

## Interface VBA

O sistema possui 27 UserForms exportados em `.frm` nesta versao publica. Os arquivos `.frm` guardam estrutura, propriedades e codigo VBA. Os `.frx` originais foram omitidos porque podem conter imagens, logos ou outros recursos privados.

| Formulario | Funcao |
|---|---|
| `frmLogin` | Login do usuario e abertura do sistema |
| `frmPrincipal` | Menu principal, KPIs, atalhos operacionais, backup automatico, abertura do dashboard e controle de caixa |
| `frmLicenca` | Ativacao e validacao de licenca |
| `frmCaixa` | Abertura, sangria e fechamento de caixa |
| `frmVendas` | Lista, filtro, abertura, edicao e cancelamento de vendas |
| `frmPedido` | Operacao da venda: produtos, quantidades, desconto, pagamento e fechamento |
| `frmPagamento` | Registro de pagamento em dinheiro, pix, cartao, vale/refeicao ou multiplas formas |
| `frmPagamentos` | Consulta e estorno de pagamentos |
| `frmVendaDiaria` | Resumo/consulta de vendas do dia |
| `frmProdutos` | Consulta e manutencao de produtos, combos, preco de venda e acerto |
| `frmCadProduto` | Cadastro e edicao de produto |
| `frmAcerto` | Acerto manual de estoque |
| `frmMovimentoEstoque` | Consulta de movimentacoes de estoque por produto, tipo, periodo e usuario |
| `frmCategorias` | Cadastro e manutencao de categorias |
| `frmGrupos` | Consulta e manutencao de grupos |
| `frmCadGrupo` | Cadastro de grupo e vinculacao de itens |
| `frmCadCombo` | Cadastro, edicao e composicao de combos |
| `frmFornecedores` | Consulta e manutencao de fornecedores |
| `frmCadFornecedor` | Cadastro e edicao de fornecedor |
| `frmCompra` | Importacao/registro de compras e notas |
| `frmEntradas` | Lista de compras/entradas e acesso a conferencia |
| `frmItensCompra` | Inclusao, remocao e fechamento de itens de compra |
| `frmConferencia` | Conferencia entre XML/importacao e produtos cadastrados |
| `frmConsulta` | Consulta rapida de produto por EAN/produto |
| `frmFluxoCaixa` | Fluxo financeiro, contas a pagar e contas a receber |
| `frmContas` | Cadastro, edicao, baixa e exclusao de contas |
| `frmLogEventos` | Consulta de logs por evento, data, usuario e relacionamento |
| `frmUsuarios` | Consulta e manutencao de usuarios |
| `frmCadUsuario` | Cadastro e edicao de usuarios |
| `frmSenha` | Confirmacao de senha para operacoes protegidas |

## Fluxos Principais

### Login e Inicializacao

1. A abertura da pasta de trabalho inicializa o sistema.
2. `ModConexao.ConectarBanco` abre a conexao ADODB usando o driver MySQL ODBC.
3. A licenca e validada por `PROC_VERIFICARLICENCA`.
4. `frmLogin` chama `PROC_LOGIN` com usuario e senha.
5. Com login aprovado, o sistema grava `UsuarioLogado` e `IDUsuarioLogado`.
6. A tela principal `frmPrincipal` e aberta para o operador.

### Tela Principal

`frmPrincipal` centraliza a navegacao e o resumo operacional:

- status do caixa aberto ou fechado;
- atalhos de cadastro, operacao, estoque, financeiro e configuracao;
- abertura do dashboard gerencial;
- KPIs de faturamento mensal e diario;
- quantidade de vendas e ticket medio;
- proximos vencimentos de contas a pagar;
- execucao manual e automatica de backup no ambiente privado.

### Caixa

O caixa operacional depende de senha ou confirmacao para acoes sensiveis:

1. O operador solicita abertura de caixa.
2. O sistema registra fundo inicial por `PROC_ABRIRCAIXA`.
3. Durante a operacao, vendas finalizadas sao vinculadas ao caixa aberto.
4. Sangrias sao registradas por `PROC_SANGRIACAIXA`.
5. O fechamento e executado por `PROC_FECHARCAIXA`.
6. O resumo financeiro do caixa pode ser consultado por `PROC_RESUMOCAIXA`.

### Vendas

1. `frmVendas` abre ou seleciona uma venda.
2. `PROC_ABRIRVENDA` cria a venda em status aberto.
3. `frmPedido` adiciona itens por produto ou codigo EAN.
4. `PROC_ADCITENSVENDA` grava o item, calcula subtotal e aplica a regra de estoque.
5. `frmPagamento` registra os pagamentos por `PROC_BAIXARPAGAMENTOS`.
6. O sistema valida se existem itens e se o total pago cobre o total da venda.
7. `PROC_FECHARVENDA` fecha a venda vinculando-a ao caixa aberto.
8. Cancelamentos usam `PROC_CANCELARVENDA`.
9. Remocao de item usa `PROC_RETIRARITENSVENDA`.
10. Estornos de pagamento usam `PROC_ESTORNAR`.

### Compras, XML e Entrada de Produtos

O fluxo de compras combina VBA, Power Query e banco:

1. Arquivos XML/NFC-e sao tratados por consultas e rotinas de ETL.
2. Power Query consolida cabecalhos e produtos importados.
3. `ModETLprocess.ProcessarNFCe` abre o arquivo de origem, carrega dados, cria aba de ETL, extrai cabecalho e itens.
4. `frmCompra` registra ou importa a compra.
5. `frmEntradas` lista compras/entradas e abre a conferencia.
6. `frmConferencia` compara itens importados com produtos cadastrados.
7. `frmItensCompra` permite incluir, remover e fechar itens de compra.
8. `PROC_REGISTRARCOMPRA`, `PROC_ENTRADAPRODUTOS`, `PROC_FECHARENTRADA`, `PROC_CANCELARENTRADA` e `PROC_RETIRARITENSCOMPRA` persistem e controlam o fluxo.

### Estoque e Produtos

O cadastro de produtos controla codigo de barras, fornecedor, medidas de compra/venda, categoria, setor, tipo, marca, estoque atual, estoque minimo, custo, custo medio, preco de venda e status.

Rotinas relacionadas:

- `frmProdutos` para consulta e manutencao;
- `frmCadProduto` para cadastro e edicao;
- `frmAcerto` para ajustes manuais;
- `frmMovimentoEstoque` para historico;
- `PROC_CADASTRARPRODUTO`;
- `PROC_ACERTOESTOQUE`;
- `PROC_CALCULAPRECOVENDA`;
- `PROC_CALCULARESTOQUEMINIMO`.

### Financeiro

O financeiro contempla:

- contas a pagar;
- contas a receber;
- fluxo de caixa;
- baixas;
- vencimentos;
- pagamentos e estornos de venda.

As telas principais sao `frmFluxoCaixa`, `frmContas`, `frmPagamento` e `frmPagamentos`.

### Dashboard e Indicadores

O dashboard foi separado da operacao principal e utiliza consultas/modelo de dados para acompanhar indicadores como:

- faturamento diario;
- faturamento mensal;
- quantidade de vendas;
- ticket medio;
- venda por produto;
- desempenho por periodo;
- acompanhamento financeiro e operacional.

## Configuracao de Conexao

A connection string foi mascarada nesta versao publica:

Driver={MySQL ODBC 9.7 Unicode Driver}
Server=<host-do-banco>
Port=<porta>
Database=<nome-do-banco>
User=<usuario>
Password=<senha>
Option=3

Em uma implantacao real, o host pode ser local ou remoto. Em cenarios com acesso remoto, o endereco do banco pode ser o IP privado da maquina na rede segura usada pelo cliente.

## Adaptacao Para Outros Clientes

Embora tenha sido criado para uma adega, o sistema pode ser adaptado para outros segmentos comerciais com ajustes em:

- categorias, setores e tipos de produto;
- regras de estoque;
- regras de preco e markup;
- formatos de entrada XML/NFC-e;
- dashboard e indicadores;
- parametros de banco;
- identidade visual dos UserForms;
- regras fiscais e operacionais especificas do cliente.

## Status

Versao documentada: `v1.0`

Estado da versao publica:

- codigo VBA exportado em texto;
- forms exportados em `.frm`;
- SQL estrutural sanitizado;
- procedures sanitizadas;
- documentacao pronta para portfolio;
- dados reais e binarios privados removidos.

## Possiveis Evolucoes

- externalizar configuracoes de conexao;
- separar scripts SQL por tabela/procedure;
- criar massa de dados ficticia para demonstracao;
- adicionar instalador/configurador do ambiente;
- migrar parte das regras para uma API;
- recriar interface em Python, web ou desktop moderno;
- evoluir o dashboard para Power BI;
- adicionar testes de integridade para procedures criticas;
- documentar diagrama ER do banco.

## Sobre o Projeto

Este projeto representa uma solucao real de automacao e gestao criada com Excel, VBA e MariaDB. Alem de resolver processos operacionais de uma loja, ele demonstra modelagem de banco, integracao com ODBC, automacao de interface, rotinas de ETL e construcao de indicadores gerenciais.
