-- Schema publico sanitizado. Sem dados reais.

DROP TABLE IF EXISTS `tab_caixa`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_caixa` (
  `IDCAIXA` int(11) NOT NULL AUTO_INCREMENT,
  `FUNDO` decimal(10,2) DEFAULT 0.00,
  `VALORCAIXA` decimal(10,2) DEFAULT 0.00,
  `SANGRIA` decimal(10,2) DEFAULT 0.00,
  `DATAABERTURA` datetime DEFAULT current_timestamp(),
  `IDUSUARIO` int(11) NOT NULL,
  `STATUS` varchar(15) NOT NULL,
  `DATAFECHAMENTO` datetime DEFAULT NULL,
  PRIMARY KEY (`IDCAIXA`),
  KEY `FK_CAIXA_USUARIO` (`IDUSUARIO`),
  CONSTRAINT `FK_CAIXA_USUARIO` FOREIGN KEY (`IDUSUARIO`) REFERENCES `tab_usuarios` (`IDUSUARIO`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_categorias`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_categorias` (
  `IDCATEGORIA` int(11) NOT NULL AUTO_INCREMENT,
  `DESCRICAO` varchar(50) NOT NULL,
  `MARKUP` decimal(10,2) NOT NULL DEFAULT 0.00,
  PRIMARY KEY (`IDCATEGORIA`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_compras`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_compras` (
  `IDCOMPRA` int(11) NOT NULL AUTO_INCREMENT,
  `IDFORNECEDOR` int(11) NOT NULL,
  `CHAVENF` varchar(44) DEFAULT NULL,
  `NUMERONF` int(11) DEFAULT NULL,
  `DATACOMPRA` datetime DEFAULT current_timestamp(),
  `VALORTOTAL` decimal(10,2) DEFAULT NULL,
  `STATUS` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`IDCOMPRA`),
  UNIQUE KEY `CHAVENF` (`CHAVENF`),
  KEY `FK_COMPRAS_FORNECEDOR` (`IDFORNECEDOR`),
  CONSTRAINT `FK_COMPRAS_FORNECEDOR` FOREIGN KEY (`IDFORNECEDOR`) REFERENCES `tab_fornecedores` (`IDFORNECEDOR`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_config`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_config` (
  `IDCONFIG` int(11) NOT NULL AUTO_INCREMENT,
  `DATA_ALTERACAO` datetime DEFAULT current_timestamp(),
  `CHAVE` varchar(100) DEFAULT NULL,
  `VALOR` varchar(300) DEFAULT NULL,
  PRIMARY KEY (`IDCONFIG`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_contasapagar`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_contasapagar` (
  `IDAPAGAR` int(11) NOT NULL AUTO_INCREMENT,
  `IDFORNECEDOR` int(11) DEFAULT NULL,
  `DESCRICAO` varchar(80) DEFAULT NULL,
  `VALOR` decimal(10,2) DEFAULT NULL,
  `VENCIMENTO` date NOT NULL,
  `STATUS` varchar(15) DEFAULT NULL,
  `DATALANCAMENTO` datetime DEFAULT current_timestamp(),
  `DATAPAGAMENTO` date DEFAULT NULL,
  PRIMARY KEY (`IDAPAGAR`),
  KEY `FK_APAGAR_FORNECEDOR` (`IDFORNECEDOR`),
  CONSTRAINT `FK_APAGAR_FORNECEDOR` FOREIGN KEY (`IDFORNECEDOR`) REFERENCES `tab_fornecedores` (`IDFORNECEDOR`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_contasareceber`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_contasareceber` (
  `IDARECEBER` int(11) NOT NULL AUTO_INCREMENT,
  `DESCRICAO` varchar(80) DEFAULT NULL,
  `VALOR` decimal(6,2) DEFAULT NULL,
  `VENCIMENTO` date NOT NULL,
  `STATUS` varchar(15) DEFAULT NULL,
  `DATARECEBIMENTO` date DEFAULT NULL,
  `OBSERVACAO` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`IDARECEBER`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_fornecedores`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_fornecedores` (
  `IDFORNECEDOR` int(11) NOT NULL AUTO_INCREMENT,
  `NOME` varchar(80) DEFAULT NULL,
  `CNPJ` varchar(24) DEFAULT NULL,
  `CONTATO` varchar(13) DEFAULT NULL,
  `ENDERECO` varchar(100) DEFAULT NULL,
  `NUMERO` varchar(10) DEFAULT NULL,
  `BAIRRO` varchar(50) DEFAULT NULL,
  `STATUS` varchar(10) DEFAULT NULL,
  `DATACADASTRO` datetime DEFAULT current_timestamp(),
  PRIMARY KEY (`IDFORNECEDOR`),
  UNIQUE KEY `CNPJ` (`CNPJ`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_grupos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_grupos` (
  `IDGRUPO` int(11) NOT NULL AUTO_INCREMENT,
  `DESCRICAO` varchar(30) NOT NULL,
  PRIMARY KEY (`IDGRUPO`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_gruposcombo`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_gruposcombo` (
  `IDCOMBO` int(11) NOT NULL,
  `IDGRUPO` int(11) NOT NULL,
  `QUANTIDADE` decimal(10,2) NOT NULL,
  PRIMARY KEY (`IDCOMBO`,`IDGRUPO`),
  KEY `IDGRUPO` (`IDGRUPO`),
  CONSTRAINT `1` FOREIGN KEY (`IDCOMBO`) REFERENCES `tab_produtos` (`IDPRODUTO`),
  CONSTRAINT `2` FOREIGN KEY (`IDGRUPO`) REFERENCES `tab_grupos` (`IDGRUPO`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_itenscombo`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_itenscombo` (
  `IDITEMCOMBO` int(11) NOT NULL AUTO_INCREMENT,
  `IDCOMBO` int(11) NOT NULL,
  `IDPRODUTO` int(11) DEFAULT NULL,
  `QUANTIDADE` decimal(10,2) DEFAULT NULL,
  `CUSTO` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`IDITEMCOMBO`),
  KEY `FK_IDCOMBO` (`IDCOMBO`),
  KEY `FK_IDPRODUTO` (`IDPRODUTO`),
  CONSTRAINT `FK_IDCOMBO` FOREIGN KEY (`IDCOMBO`) REFERENCES `tab_produtos` (`IDPRODUTO`),
  CONSTRAINT `FK_IDPRODUTO` FOREIGN KEY (`IDPRODUTO`) REFERENCES `tab_produtos` (`IDPRODUTO`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_itenscompra`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_itenscompra` (
  `IDITEMCOMPRA` int(11) NOT NULL AUTO_INCREMENT,
  `IDCOMPRA` int(11) NOT NULL,
  `IDPRODUTO` int(11) NOT NULL,
  `MEDIDACOMPRA` varchar(5) DEFAULT NULL,
  `QUANTIDADE` int(11) NOT NULL,
  `CUSTOUNITARIO` decimal(10,2) DEFAULT NULL,
  `SUBTOTAL` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`IDITEMCOMPRA`),
  KEY `FK_ITENSCOMPRA_COMPRAS` (`IDCOMPRA`),
  CONSTRAINT `FK_ITENSCOMPRA_COMPRAS` FOREIGN KEY (`IDCOMPRA`) REFERENCES `tab_compras` (`IDCOMPRA`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_itensvenda`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_itensvenda` (
  `IDITEMVENDA` int(11) NOT NULL AUTO_INCREMENT,
  `IDVENDA` int(11) NOT NULL,
  `IDPRODUTO` int(11) NOT NULL,
  `QUANTIDADE` int(11) NOT NULL,
  `PRECOUNITARIO` decimal(10,2) DEFAULT NULL,
  `CUSTOUNITARIO` decimal(10,2) DEFAULT NULL,
  `SUBTOTAL` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`IDITEMVENDA`),
  KEY `FK_ITENSVENDA_VENDAS` (`IDVENDA`),
  KEY `FK_ITENSVENDA_PRODUTOS` (`IDPRODUTO`),
  CONSTRAINT `FK_ITENSVENDA_PRODUTOS` FOREIGN KEY (`IDPRODUTO`) REFERENCES `tab_produtos` (`IDPRODUTO`),
  CONSTRAINT `FK_ITENSVENDA_VENDAS` FOREIGN KEY (`IDVENDA`) REFERENCES `tab_vendas` (`IDVENDA`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_logeventos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_logeventos` (
  `IDEVENTO` int(11) NOT NULL AUTO_INCREMENT,
  `TIPOEVENTO` varchar(30) NOT NULL,
  `DESCRICAO` varchar(80) NOT NULL,
  `IDRELACIONADO` int(11) DEFAULT NULL,
  `DATAEVENTO` datetime DEFAULT current_timestamp(),
  `IDUSUARIO` int(11) NOT NULL,
  PRIMARY KEY (`IDEVENTO`),
  KEY `FK_EVENTO_USUARIO` (`IDUSUARIO`),
  CONSTRAINT `FK_EVENTO_USUARIO` FOREIGN KEY (`IDUSUARIO`) REFERENCES `tab_usuarios` (`IDUSUARIO`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_movimentoestoque`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_movimentoestoque` (
  `IDMOVIMENTO` int(11) NOT NULL AUTO_INCREMENT,
  `IDPRODUTO` int(11) NOT NULL,
  `TIPOMOVIMENTO` varchar(20) NOT NULL,
  `QUANTIDADE` decimal(10,2) NOT NULL DEFAULT 0.00,
  `CUSTOUNITARIO` decimal(10,2) NOT NULL,
  `ESTOQUEANTERIOR` int(11) DEFAULT NULL,
  `ESTOQUEPOSTERIOR` int(11) DEFAULT NULL,
  `DATAMOVIMENTO` datetime DEFAULT current_timestamp(),
  `IDUSUARIO` int(11) DEFAULT NULL,
  `OBSERVACAO` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`IDMOVIMENTO`),
  KEY `FK_MOVESTOQUE_PRODUTO` (`IDPRODUTO`),
  KEY `FK_MOVESTOQUE_USUARIO` (`IDUSUARIO`),
  CONSTRAINT `FK_MOVESTOQUE_PRODUTO` FOREIGN KEY (`IDPRODUTO`) REFERENCES `tab_produtos` (`IDPRODUTO`),
  CONSTRAINT `FK_MOVESTOQUE_USUARIO` FOREIGN KEY (`IDUSUARIO`) REFERENCES `tab_usuarios` (`IDUSUARIO`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_pagamentos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_pagamentos` (
  `IDPAGAMENTO` int(11) NOT NULL AUTO_INCREMENT,
  `IDVENDA` int(11) DEFAULT NULL,
  `FORMAPAGAMENTO` varchar(30) DEFAULT NULL,
  `VALORRECEBIDO` decimal(10,2) DEFAULT 0.00,
  `VALORPAGO` decimal(10,2) DEFAULT 0.00,
  `DATAPAGAMENTO` date DEFAULT curdate(),
  `HORAPAGAMENTO` time DEFAULT curtime(),
  PRIMARY KEY (`IDPAGAMENTO`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_produtos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_produtos` (
  `IDPRODUTO` int(11) NOT NULL AUTO_INCREMENT,
  `CODIGOFORNECEDOR` int(11) DEFAULT NULL,
  `NOME` varchar(80) NOT NULL,
  `CODIGOBARRASCX` varchar(15) DEFAULT NULL,
  `CODIGOBARRAS` varchar(15) DEFAULT NULL,
  `MEDIDACOMPRA` varchar(5) DEFAULT NULL,
  `QUANTIDADECOMPRA` decimal(10,2) DEFAULT NULL,
  `QUANTIDADEEMBALAGEM` decimal(10,2) DEFAULT NULL,
  `MEDIDAVENDA` varchar(5) DEFAULT NULL,
  `CATEGORIA` varchar(30) DEFAULT NULL,
  `SETOR` varchar(30) DEFAULT NULL,
  `TIPO` varchar(10) DEFAULT NULL,
  `MARCA` varchar(40) DEFAULT NULL,
  `ESTOQUEATUAL` decimal(10,2) DEFAULT 0.00,
  `QUANTIDADEUSO` decimal(10,2) DEFAULT 0.00,
  `ESTOQUEMINIMO` decimal(10,2) DEFAULT NULL,
  `CUSTO` decimal(10,2) DEFAULT 0.00,
  `CUSTOUNITARIO` decimal(10,2) DEFAULT 0.00,
  `CUSTOMEDIO` decimal(10,2) DEFAULT 0.00,
  `PRECOVENDA` decimal(10,2) DEFAULT 0.00,
  `STATUS` varchar(10) DEFAULT NULL,
  `DATACADASTRO` datetime DEFAULT current_timestamp(),
  `IDCATEGORIA` int(11) DEFAULT NULL,
  PRIMARY KEY (`IDPRODUTO`),
  UNIQUE KEY `CODIGOFORNECEDOR` (`CODIGOFORNECEDOR`),
  KEY `FK_PRODUTO_CATEGORIA` (`IDCATEGORIA`),
  CONSTRAINT `FK_PRODUTO_CATEGORIA` FOREIGN KEY (`IDCATEGORIA`) REFERENCES `tab_categorias` (`IDCATEGORIA`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_produtosgrupo`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_produtosgrupo` (
  `IDGRUPO` int(11) NOT NULL,
  `IDPRODUTO` int(11) NOT NULL,
  `QUANTIDADE` decimal(10,2) NOT NULL,
  PRIMARY KEY (`IDGRUPO`,`IDPRODUTO`),
  KEY `IDPRODUTO` (`IDPRODUTO`),
  CONSTRAINT `1` FOREIGN KEY (`IDGRUPO`) REFERENCES `tab_grupos` (`IDGRUPO`),
  CONSTRAINT `2` FOREIGN KEY (`IDPRODUTO`) REFERENCES `tab_produtos` (`IDPRODUTO`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_usuarios`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_usuarios` (
  `IDUSUARIO` int(11) NOT NULL AUTO_INCREMENT,
  `NOME` varchar(80) NOT NULL,
  `USUARIO` varchar(20) NOT NULL,
  `SENHA` varchar(64) NOT NULL,
  `DATACADASTRO` datetime DEFAULT current_timestamp(),
  PRIMARY KEY (`IDUSUARIO`),
  UNIQUE KEY `LOGIN` (`USUARIO`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `tab_vendas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_vendas` (
  `IDVENDA` int(11) NOT NULL AUTO_INCREMENT,
  `IDCAIXA` int(11) DEFAULT NULL,
  `REFERENCIA` varchar(100) DEFAULT NULL,
  `DATAVENDA` date DEFAULT curdate(),
  `HORAVENDA` time DEFAULT curtime(),
  `VALORTOTAL` decimal(10,2) DEFAULT NULL,
  `DESCONTO` decimal(10,2) DEFAULT NULL,
  `VALORFINAL` decimal(10,2) DEFAULT NULL,
  `FORMAPAGAMENTO` varchar(15) DEFAULT NULL,
  `STATUS` varchar(30) DEFAULT NULL,
  `IDUSUARIO` int(11) NOT NULL,
  `VALOR_PENDENTE` decimal(10,2) DEFAULT 0.00,
  PRIMARY KEY (`IDVENDA`),
  KEY `FK_VENDAS_USUARIO` (`IDUSUARIO`),
  CONSTRAINT `FK_VENDAS_USUARIO` FOREIGN KEY (`IDUSUARIO`) REFERENCES `tab_usuarios` (`IDUSUARIO`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

DROP TABLE IF EXISTS `temp_produtos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `temp_produtos` (
  `IDPRODUTO` int(11) DEFAULT NULL,
  `CODIGOFORNECEDOR` int(11) DEFAULT NULL,
  `NOME` varchar(80) DEFAULT NULL,
  `CODIGOBARRASCX` varchar(15) DEFAULT NULL,
  `CODIGOBARRAS` varchar(13) DEFAULT NULL,
  `MEDIDACOMPRA` varchar(5) DEFAULT NULL,
  `QUANTIDADECOMPRA` int(11) DEFAULT NULL,
  `MEDIDAVENDA` varchar(5) DEFAULT NULL,
  `QUANTIDADEVENDA` int(11) DEFAULT NULL,
  `CATEGORIA` varchar(30) DEFAULT NULL,
  `TIPO` varchar(10) DEFAULT NULL,
  `MARCA` varchar(40) DEFAULT NULL,
  `ESTOQUEATUAL` decimal(10,2) DEFAULT NULL,
  `QUANTIDADEUSO` decimal(10,2) DEFAULT NULL,
  `ESTOQUEMINIMO` int(11) DEFAULT NULL,
  `CUSTO` decimal(10,2) DEFAULT NULL,
  `CUSTOMEDIO` decimal(10,2) DEFAULT NULL,
  `PRECOVENDA` decimal(10,2) DEFAULT NULL,
  `STATUS` varchar(10) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;
