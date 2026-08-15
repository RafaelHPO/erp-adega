/*M!999999\- enable the sandbox mode */ 
-- MariaDB dump 10.19-12.3.2-MariaDB, for Win64 (AMD64)
--
-- Host: demo    Database: adega_demo
-- ------------------------------------------------------
-- Server version	12.3.2-MariaDB

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*M!100616 SET @OLD_NOTE_VERBOSITY=@@NOTE_VERBOSITY, NOTE_VERBOSITY=0 */;

--
-- Table structure for table `tab_caixa`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_caixa`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_caixa` WRITE;
/*!40000 ALTER TABLE `tab_caixa` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_caixa` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_categorias`
--

DROP TABLE IF EXISTS `tab_categorias`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_categorias` (
  `IDCATEGORIA` int(11) NOT NULL AUTO_INCREMENT,
  `DESCRICAO` varchar(50) NOT NULL,
  `MARKUP` decimal(10,2) NOT NULL DEFAULT 0.00,
  PRIMARY KEY (`IDCATEGORIA`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_categorias`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_categorias` WRITE;
/*!40000 ALTER TABLE `tab_categorias` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_categorias` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_compras`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_compras`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_compras` WRITE;
/*!40000 ALTER TABLE `tab_compras` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_compras` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_config`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_config`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_config` WRITE;
/*!40000 ALTER TABLE `tab_config` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_config` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_contasapagar`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_contasapagar`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_contasapagar` WRITE;
/*!40000 ALTER TABLE `tab_contasapagar` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_contasapagar` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_contasareceber`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_contasareceber`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_contasareceber` WRITE;
/*!40000 ALTER TABLE `tab_contasareceber` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_contasareceber` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_fornecedores`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_fornecedores`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_fornecedores` WRITE;
/*!40000 ALTER TABLE `tab_fornecedores` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_fornecedores` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_grupos`
--

DROP TABLE IF EXISTS `tab_grupos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8mb4 */;
CREATE TABLE `tab_grupos` (
  `IDGRUPO` int(11) NOT NULL AUTO_INCREMENT,
  `DESCRICAO` varchar(30) NOT NULL,
  PRIMARY KEY (`IDGRUPO`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_grupos`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_grupos` WRITE;
/*!40000 ALTER TABLE `tab_grupos` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_grupos` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_gruposcombo`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_gruposcombo`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_gruposcombo` WRITE;
/*!40000 ALTER TABLE `tab_gruposcombo` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_gruposcombo` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_itenscombo`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_itenscombo`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_itenscombo` WRITE;
/*!40000 ALTER TABLE `tab_itenscombo` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_itenscombo` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_itenscompra`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_itenscompra`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_itenscompra` WRITE;
/*!40000 ALTER TABLE `tab_itenscompra` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_itenscompra` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_itensvenda`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_itensvenda`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_itensvenda` WRITE;
/*!40000 ALTER TABLE `tab_itensvenda` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_itensvenda` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_logeventos`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_logeventos`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_logeventos` WRITE;
/*!40000 ALTER TABLE `tab_logeventos` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_logeventos` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_movimentoestoque`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_movimentoestoque`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_movimentoestoque` WRITE;
/*!40000 ALTER TABLE `tab_movimentoestoque` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_movimentoestoque` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_pagamentos`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_pagamentos`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_pagamentos` WRITE;
/*!40000 ALTER TABLE `tab_pagamentos` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_pagamentos` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_produtos`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_produtos`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_produtos` WRITE;
/*!40000 ALTER TABLE `tab_produtos` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_produtos` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_produtosgrupo`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_produtosgrupo`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_produtosgrupo` WRITE;
/*!40000 ALTER TABLE `tab_produtosgrupo` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_produtosgrupo` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_usuarios`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_usuarios`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_usuarios` WRITE;
/*!40000 ALTER TABLE `tab_usuarios` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_usuarios` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `tab_vendas`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tab_vendas`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `tab_vendas` WRITE;
/*!40000 ALTER TABLE `tab_vendas` DISABLE KEYS */;
/*!40000 ALTER TABLE `tab_vendas` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Table structure for table `temp_produtos`
--

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
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `temp_produtos`
--

SET @OLD_AUTOCOMMIT=@@AUTOCOMMIT, @@AUTOCOMMIT=0;
LOCK TABLES `temp_produtos` WRITE;
/*!40000 ALTER TABLE `temp_produtos` DISABLE KEYS */;
/*!40000 ALTER TABLE `temp_produtos` ENABLE KEYS */;
UNLOCK TABLES;
COMMIT;
SET AUTOCOMMIT=@OLD_AUTOCOMMIT;

--
-- Dumping events for database 'adega'
--


-- ------------------------------------------------------
-- Dados ficticios para demonstracao e screenshots
-- Massa gerada sem dados reais do cliente original.
-- ------------------------------------------------------

INSERT INTO `tab_usuarios` (`IDUSUARIO`, `NOME`, `USUARIO`, `SENHA`, `DATACADASTRO`) VALUES
(1,'Administrador Demo','admin.demo','9F86D081884C7D659A2FEAA0C55AD015A3BF4F1B2B0B822CD15D6C15B0F00A08','2026-08-01 09:00:00'),
(2,'Operador Caixa','caixa.demo','9F86D081884C7D659A2FEAA0C55AD015A3BF4F1B2B0B822CD15D6C15B0F00A08','2026-08-01 09:15:00'),
(3,'Gestor Comercial','gestor.demo','9F86D081884C7D659A2FEAA0C55AD015A3BF4F1B2B0B822CD15D6C15B0F00A08','2026-08-01 09:30:00');

INSERT INTO `tab_config` (`IDCONFIG`, `DATA_ALTERACAO`, `CHAVE`, `VALOR`) VALUES
(1,'2026-08-01 08:00:00','VERSAO_SISTEMA','1.0-DEMO'),
(2,'2026-08-01 08:00:00','DATA_VERSAO','15/08/2026'),
(3,'2026-08-01 08:00:00','LICENCA_CHAVE','DEMO-ERP-2026'),
(4,'2026-08-01 08:00:00','LICENCA_VALIDADE','2027-12-31'),
(5,'2026-08-01 08:00:00','SENHA_ADMIN','9F86D081884C7D659A2FEAA0C55AD015A3BF4F1B2B0B822CD15D6C15B0F00A08');

INSERT INTO `tab_categorias` (`IDCATEGORIA`, `DESCRICAO`, `MARKUP`) VALUES
(1,'Bebidas',55.0),
(2,'Mercearia',40.0),
(3,'Conveniencia',45.0),
(4,'Combos',35.0),
(5,'Sem alcool',38.0);

INSERT INTO `tab_fornecedores` (`IDFORNECEDOR`, `NOME`, `CNPJ`, `CONTATO`, `ENDERECO`, `NUMERO`, `BAIRRO`, `STATUS`, `DATACADASTRO`) VALUES
(1,'Distribuidora Alfa Demo','00.000.000/0001-01','(11) 90000-0001','Rua Exemplo','100','Centro','ATIVO','2026-08-01 10:00:00'),
(2,'Atacado Beta Demo','00.000.000/0001-02','(11) 90000-0002','Avenida Modelo','250','Comercial','ATIVO','2026-08-01 10:10:00'),
(3,'Fornecedor Gama Demo','00.000.000/0001-03','(11) 90000-0003','Rua Teste','45','Industrial','ATIVO','2026-08-01 10:20:00');

INSERT INTO `tab_produtos` (`IDPRODUTO`, `CODIGOFORNECEDOR`, `NOME`, `CODIGOBARRASCX`, `CODIGOBARRAS`, `MEDIDACOMPRA`, `QUANTIDADECOMPRA`, `QUANTIDADEEMBALAGEM`, `MEDIDAVENDA`, `CATEGORIA`, `SETOR`, `TIPO`, `MARCA`, `ESTOQUEATUAL`, `QUANTIDADEUSO`, `ESTOQUEMINIMO`, `CUSTO`, `CUSTOMEDIO`, `PRECOVENDA`, `STATUS`, `DATACADASTRO`, `IDGRUPO`, `IDCOMBO`) VALUES
(1,1,'Cerveja Pilsen 350ml','7890000000010','7890000001017','CX',12,12,'UN','Bebidas','Bebidas','PRODUTO','Marca Demo',84.0,0.0,24,3.2,3.3,5.99,'ATIVO','2026-08-01 11:00:00',NULL,NULL),
(2,1,'Vinho Tinto Seco 750ml','7890000000027','7890000001024','CX',6,6,'UN','Bebidas','Vinhos','PRODUTO','Vale Demo',30.0,0.0,8,24.9,25.5,42.9,'ATIVO','2026-08-01 11:05:00',NULL,NULL),
(3,2,'Agua Mineral 500ml','7890000000034','7890000001031','FD',12,12,'UN','Sem alcool','Bebidas','PRODUTO','Fonte Demo',120.0,0.0,36,1.1,1.2,2.5,'ATIVO','2026-08-01 11:10:00',NULL,NULL),
(4,2,'Refrigerante Cola 2L','7890000000041','7890000001048','FD',6,6,'UN','Sem alcool','Bebidas','PRODUTO','Refri Demo',42.0,0.0,18,5.8,6.0,9.99,'ATIVO','2026-08-01 11:15:00',NULL,NULL),
(5,3,'Carvao 3kg','7890000000058','7890000001055','PC',1,1,'UN','Conveniencia','Churrasco','PRODUTO','Brasa Demo',18.0,0.0,6,10.5,10.9,18.9,'ATIVO','2026-08-01 11:20:00',NULL,NULL),
(6,3,'Gelo 5kg','7890000000065','7890000001062','PC',1,1,'UN','Conveniencia','Gelo','PRODUTO','Ice Demo',25.0,0.0,10,6.0,6.2,12.0,'ATIVO','2026-08-01 11:25:00',NULL,NULL),
(7,1,'Combo Churrasco Demo','7890000000072','7890000001079','UN',1,1,'UN','Combos','Combos','COMBO','ERP Demo',10.0,0.0,3,28.0,28.0,49.9,'ATIVO','2026-08-01 11:30:00',NULL,NULL),
(8,2,'Suco Integral Uva 1L','7890000000089','7890000001086','CX',6,6,'UN','Sem alcool','Bebidas','PRODUTO','Natural Demo',36.0,0.0,12,7.5,7.8,13.9,'ATIVO','2026-08-01 11:35:00',NULL,NULL);

INSERT INTO `tab_grupos` (`IDGRUPO`, `DESCRICAO`) VALUES
(1,'Bebidas geladas'),
(2,'Churrasco');

INSERT INTO `tab_produtosgrupo` (`IDGRUPO`, `IDPRODUTO`, `QUANTIDADE`) VALUES
(1,1,1.0),
(1,3,1.0),
(1,4,1.0),
(2,5,1.0),
(2,6,1.0);

INSERT INTO `tab_gruposcombo` (`IDCOMBO`, `IDGRUPO`, `QUANTIDADE`) VALUES
(7,1,2.0),
(7,2,1.0);

INSERT INTO `tab_itenscombo` (`IDITEMCOMBO`, `IDCOMBO`, `IDPRODUTO`, `QUANTIDADE`, `CUSTO`) VALUES
(1,7,1,6.0,3.3),
(2,7,5,1.0,10.9),
(3,7,6,1.0,6.2);

INSERT INTO `tab_caixa` (`IDCAIXA`, `FUNDO`, `VALORCAIXA`, `SANGRIA`, `DATAABERTURA`, `IDUSUARIO`, `STATUS`, `DATAFECHAMENTO`) VALUES
(1,100.0,486.4,50.0,'2026-08-12 09:00:00',2,'FECHADO','2026-08-12 22:10:00'),
(2,120.0,612.2,80.0,'2026-08-13 09:05:00',2,'FECHADO','2026-08-13 22:15:00'),
(3,100.0,428.7,0.0,'2026-08-14 09:10:00',2,'FECHADO','2026-08-14 22:05:00'),
(4,150.0,0.0,0.0,'2026-08-15 09:00:00',2,'ABERTO',NULL);

INSERT INTO `tab_vendas` (`IDVENDA`, `IDCAIXA`, `REFERENCIA`, `DATAVENDA`, `HORAVENDA`, `VALORTOTAL`, `DESCONTO`, `VALORFINAL`, `FORMAPAGAMENTO`, `STATUS`, `IDUSUARIO`, `VALOR_PENDENTE`) VALUES
(1,1,'BALCAO 001','2026-08-12','10:15:00',35.94,0.0,35.94,'PIX','FECHADO',2,0.0),
(2,1,'BALCAO 002','2026-08-12','14:40:00',61.8,5.0,56.8,'CARTAO','FECHADO',2,0.0),
(3,2,'DELIVERY 010','2026-08-13','18:20:00',92.7,0.0,92.7,'MULTI','FECHADO',2,0.0),
(4,3,'BALCAO 017','2026-08-14','20:05:00',49.9,0.0,49.9,'DINHEIRO','FECHADO',2,0.0),
(5,4,'MESA DEMO','2026-08-15','12:30:00',72.79,0.0,72.79,'PIX','FECHADO',2,0.0),
(6,4,'PEDIDO ABERTO','2026-08-15','15:45:00',23.98,0.0,23.98,NULL,'ABERTO',2,23.98);

INSERT INTO `tab_itensvenda` (`IDITEMVENDA`, `IDVENDA`, `IDPRODUTO`, `QUANTIDADE`, `PRECOUNITARIO`, `CUSTOUNITARIO`, `SUBTOTAL`) VALUES
(1,1,1,6.0,5.99,3.3,35.94),
(2,2,2,1.0,42.9,25.5,42.9),
(3,2,6,1.0,12.0,6.2,12.0),
(4,2,3,2.0,2.95,1.2,5.9),
(5,3,1,6.0,5.99,3.3,35.94),
(6,3,4,2.0,9.99,6.0,19.98),
(7,3,5,1.0,18.9,10.9,18.9),
(8,3,6,1.0,12.0,6.2,12.0),
(9,3,3,2.0,2.94,1.2,5.88),
(10,4,7,1.0,49.9,28.0,49.9),
(11,5,2,1.0,42.9,25.5,42.9),
(12,5,4,2.0,9.99,6.0,19.98),
(13,5,3,4.0,2.48,1.2,9.91),
(14,6,1,4.0,5.99,3.3,23.98);

INSERT INTO `tab_pagamentos` (`IDPAGAMENTO`, `IDVENDA`, `FORMAPAGAMENTO`, `VALORRECEBIDO`, `VALORPAGO`, `DATAPAGAMENTO`, `HORAPAGAMENTO`) VALUES
(1,1,'PIX',35.94,35.94,'2026-08-12','10:17:00'),
(2,2,'CARTAO',56.8,56.8,'2026-08-12','14:43:00'),
(3,3,'PIX',50.0,50.0,'2026-08-13','18:25:00'),
(4,3,'CARTAO',42.7,42.7,'2026-08-13','18:26:00'),
(5,4,'DINHEIRO',50.0,49.9,'2026-08-14','20:08:00'),
(6,5,'PIX',72.79,72.79,'2026-08-15','12:35:00');

INSERT INTO `tab_compras` (`IDCOMPRA`, `IDFORNECEDOR`, `CHAVENF`, `NUMERONF`, `DATACOMPRA`, `VALORTOTAL`, `STATUS`) VALUES
(1,1,'DEMO000000000000000000000000000000000000000001',1001,'2026-08-10 09:30:00',684.0,'FECHADO'),
(2,2,'DEMO000000000000000000000000000000000000000002',1002,'2026-08-13 10:10:00',428.4,'FECHADO'),
(3,3,'DEMO000000000000000000000000000000000000000003',1003,'2026-08-15 11:00:00',183.0,'ABERTO');

INSERT INTO `tab_itenscompra` (`IDITEMCOMPRA`, `IDCOMPRA`, `IDPRODUTO`, `MEDIDACOMPRA`, `QUANTIDADE`, `CUSTOUNITARIO`, `SUBTOTAL`) VALUES
(1,1,1,'CX',5.0,38.4,192.0),
(2,1,2,'CX',4.0,149.4,597.6),
(3,2,3,'FD',6.0,13.2,79.2),
(4,2,4,'FD',5.0,34.8,174.0),
(5,2,8,'CX',3.0,45.0,135.0),
(6,3,5,'PC',8.0,10.5,84.0),
(7,3,6,'PC',10.0,6.0,60.0);

INSERT INTO `tab_movimentoestoque` (`IDMOVIMENTO`, `IDPRODUTO`, `TIPOMOVIMENTO`, `QUANTIDADE`, `CUSTOUNITARIO`, `ESTOQUEANTERIOR`, `ESTOQUEPOSTERIOR`, `DATAMOVIMENTO`, `IDUSUARIO`, `OBSERVACAO`) VALUES
(1,1,'ENTRADA',60.0,3.2,30.0,90.0,'2026-08-10 09:45:00',2,'Compra demo 1001'),
(2,2,'ENTRADA',24.0,24.9,12.0,36.0,'2026-08-10 09:50:00',2,'Compra demo 1001'),
(3,1,'SAIDA',6.0,3.3,90.0,84.0,'2026-08-12 10:17:00',2,'Venda demo 1'),
(4,2,'SAIDA',1.0,25.5,36.0,35.0,'2026-08-12 14:43:00',2,'Venda demo 2'),
(5,3,'ENTRADA',72.0,1.1,60.0,132.0,'2026-08-13 10:20:00',2,'Compra demo 1002'),
(6,4,'ENTRADA',30.0,5.8,20.0,50.0,'2026-08-13 10:25:00',2,'Compra demo 1002'),
(7,7,'SAIDA',1.0,28.0,11.0,10.0,'2026-08-14 20:08:00',2,'Venda demo 4'),
(8,5,'ACERTO',2.0,10.9,16.0,18.0,'2026-08-15 08:30:00',3,'Ajuste inventario demo');

INSERT INTO `tab_contasapagar` (`IDAPAGAR`, `IDFORNECEDOR`, `DESCRICAO`, `VALOR`, `VENCIMENTO`, `STATUS`, `DATALANCAMENTO`, `DATAPAGAMENTO`) VALUES
(1,1,'Boleto compra demo 1001',684.0,'2026-08-20','PENDENTE','2026-08-10 10:00:00',NULL),
(2,2,'Boleto compra demo 1002',428.4,'2026-08-25','PENDENTE','2026-08-13 10:30:00',NULL),
(3,3,'Despesa operacional demo',180.0,'2026-08-14','PAGO','2026-08-01 09:00:00','2026-08-14');

INSERT INTO `tab_contasareceber` (`IDARECEBER`, `DESCRICAO`, `VALOR`, `VENCIMENTO`, `STATUS`, `DATARECEBIMENTO`, `OBSERVACAO`) VALUES
(1,'Recebimento delivery demo',92.7,'2026-08-13','RECEBIDO','2026-08-13','Pedido demo'),
(2,'Venda corporativa demo',320.0,'2026-08-22','PENDENTE',NULL,'Cliente ficticio');

INSERT INTO `tab_logeventos` (`IDEVENTO`, `TIPOEVENTO`, `DESCRICAO`, `IDRELACIONADO`, `DATAEVENTO`, `IDUSUARIO`) VALUES
(1,'LOGIN','Login efetuado no ambiente demo',2,'2026-08-15 09:00:00',2),
(2,'CAIXA','Caixa demo aberto',4,'2026-08-15 09:01:00',2),
(3,'VENDA','Venda demo finalizada',5,'2026-08-15 12:35:00',2),
(4,'ESTOQUE','Acerto de estoque demo',5,'2026-08-15 08:30:00',3);

INSERT INTO `temp_produtos` (`IDPRODUTO`, `CODIGOFORNECEDOR`, `NOME`, `CODIGOBARRASCX`, `CODIGOBARRAS`, `MEDIDACOMPRA`, `QUANTIDADECOMPRA`, `MEDIDAVENDA`, `QUANTIDADEVENDA`, `CATEGORIA`, `TIPO`, `MARCA`, `ESTOQUEATUAL`, `QUANTIDADEUSO`, `ESTOQUEMINIMO`, `CUSTO`, `CUSTOMEDIO`, `PRECOVENDA`, `STATUS`) VALUES
(101,1,'Produto Importado Demo','7899999999010','7899999999027','CX',12,'UN',1,'Bebidas','PRODUTO','Demo',0.0,0.0,5,4.0,4.0,7.9,'PENDENTE');

--
-- Dumping routines for database 'adega'
--
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ABRIRCAIXA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ABRIRCAIXA`(
    IN P_IDUSUARIO INT,
    IN P_FUNDO DECIMAL(10,2)
)
BEGIN

DECLARE V_IDCAIXA INT DEFAULT 0;
DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN
    GET DIAGNOSTICS CONDITION 1 V_SQLSTATE = RETURNED_SQLSTATE, V_MSG = MESSAGE_TEXT;
    ROLLBACK;
    SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;
END;

IF EXISTS(SELECT 1 FROM TAB_CAIXA WHERE STATUS='ABERTO') THEN

    SELECT 'CAIXA JÁ ABERTO!' AS RETORNO;

ELSEIF P_FUNDO < 0 THEN

    SELECT 'VALOR DE FUNDO INVÁLIDO!' AS RETORNO;

ELSEIF P_FUNDO > 100 THEN

    SELECT 'VALOR DE FUNDO ACIMA DO LIMITE!' AS RETORNO;

ELSE

    START TRANSACTION;

    INSERT INTO TAB_CAIXA(FUNDO,VALORCAIXA,SANGRIA,IDUSUARIO,STATUS)
    VALUES(P_FUNDO,P_FUNDO,0,P_IDUSUARIO,'ABERTO');

    SET V_IDCAIXA = LAST_INSERT_ID();

    INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO,DESCRICAO,IDRELACIONADO,IDUSUARIO)
    VALUES('CAIXA',CONCAT('ABERTURA DE CAIXA - FUNDO: ',P_FUNDO),V_IDCAIXA,P_IDUSUARIO);

    COMMIT;

    SELECT 'CAIXA ABERTO' AS STATUS, V_IDCAIXA AS IDCAIXA, P_FUNDO AS FUNDO;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ABRIRVENDA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ABRIRVENDA`(
	IN `P_IDUSUARIO` INT
)
BEGIN 

DECLARE V_IDVENDA INT;

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

IF NOT EXISTS ( SELECT 1 FROM TAB_CAIXA WHERE STATUS = 'ABERTO') THEN
SELECT 'NENHUM CAIXA ABERTO!' AS ERRO;

ELSE 

START TRANSACTION ;

INSERT INTO TAB_VENDAS(STATUS,IDUSUARIO) VALUES ('ABERTO',p_idusuario);

SET V_IDVENDA = LAST_INSERT_ID();

INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO,IDUSUARIO)
VALUES ('VENDA','ABERTURA DE VENDA',V_IDVENDA,P_IDUSUARIO);

COMMIT;

END IF;

SELECT V_IDVENDA AS IDVENDA;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ACERTOESTOQUE` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ACERTOESTOQUE`(
    IN P_IDUSUARIO INT,
    IN P_IDPRODUTO INT,
    IN P_QUANTIDADE DECIMAL(10,2),
    IN P_TIPOACERTO VARCHAR(15),
    IN P_MOTIVOACERTO VARCHAR(80)
)
BEGIN

    DECLARE V_ESTOQUEATUAL DECIMAL(10,2) DEFAULT 0;
    DECLARE V_NOVOESTOQUE DECIMAL(10,2) DEFAULT 0;

    DECLARE V_SQLSTATE VARCHAR(5);
    DECLARE V_MSG TEXT;

    DECLARE EXIT HANDLER FOR SQLEXCEPTION
    BEGIN

        GET DIAGNOSTICS CONDITION 1
            V_SQLSTATE = RETURNED_SQLSTATE,
            V_MSG = MESSAGE_TEXT;

        ROLLBACK;

        SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS MSG;

    END;

    START TRANSACTION;

    SELECT IFNULL(ESTOQUEATUAL,0)
      INTO V_ESTOQUEATUAL
      FROM TAB_PRODUTOS
     WHERE IDPRODUTO = P_IDPRODUTO;

    IF V_ESTOQUEATUAL IS NULL THEN
        SET V_ESTOQUEATUAL = 0;
    END IF;

    IF UPPER(P_TIPOACERTO) = 'ENTRADA' THEN

        SET V_NOVOESTOQUE = V_ESTOQUEATUAL + P_QUANTIDADE;

    ELSE

        SET V_NOVOESTOQUE = V_ESTOQUEATUAL - P_QUANTIDADE;

    END IF;

    UPDATE TAB_PRODUTOS
       SET ESTOQUEATUAL = IFNULL(V_NOVOESTOQUE,0)
     WHERE IDPRODUTO = P_IDPRODUTO;

    INSERT INTO TAB_LOGEVENTOS
    (
        TIPOEVENTO,
        DESCRICAO,
        IDRELACIONADO,
        IDUSUARIO
    )
    VALUES
    (
        P_TIPOACERTO,
        P_MOTIVOACERTO,
        P_IDPRODUTO,
        P_IDUSUARIO
    );

    INSERT INTO TAB_MOVIMENTOESTOQUE
    (
        IDPRODUTO,
        TIPOMOVIMENTO,
        QUANTIDADE,
        CUSTOUNITARIO,
        ESTOQUEANTERIOR,
        ESTOQUEPOSTERIOR,
        IDUSUARIO,
        OBSERVACAO
    )
    VALUES
    (
        P_IDPRODUTO,
        P_TIPOACERTO,
        P_QUANTIDADE,
        0,
        V_ESTOQUEATUAL,
        V_NOVOESTOQUE,
        P_IDUSUARIO,
        P_MOTIVOACERTO
    );

    COMMIT;

    SELECT 'LANÇAMENTO REALIZADO COM SUCESSO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ADCITEMGRUPO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ADCITEMGRUPO`(
    IN P_IDGRUPO INT,
    IN P_IDPRODUTO INT,
    IN P_QUANTIDADE DECIMAL(10,2)
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN
    GET DIAGNOSTICS CONDITION 1 V_SQLSTATE = RETURNED_SQLSTATE, V_MSG = MESSAGE_TEXT;
    ROLLBACK;
    SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;
END;

IF NOT EXISTS(
    SELECT 1 FROM TAB_GRUPOS WHERE IDGRUPO=P_IDGRUPO
) THEN

    SELECT 'GRUPO NÃO EXISTE' AS RETORNO;

ELSEIF NOT EXISTS(
    SELECT 1 FROM TAB_PRODUTOS 
    WHERE IDPRODUTO=P_IDPRODUTO
    AND STATUS='ATIVO'
) THEN

    SELECT 'PRODUTO NÃO EXISTE OU INATIVO' AS RETORNO;

ELSEIF P_QUANTIDADE <= 0 THEN

    SELECT 'QUANTIDADE INVÁLIDA' AS RETORNO;

ELSEIF EXISTS(
    SELECT 1 
    FROM TAB_PRODUTOSGRUPO
    WHERE IDGRUPO=P_IDGRUPO
    AND IDPRODUTO=P_IDPRODUTO
) THEN

    SELECT 'PRODUTO JÁ CADASTRADO NO GRUPO' AS RETORNO;

ELSE

    START TRANSACTION;

    INSERT INTO TAB_PRODUTOSGRUPO(IDGRUPO,IDPRODUTO,QUANTIDADE)
    VALUES(P_IDGRUPO,P_IDPRODUTO,P_QUANTIDADE);

    COMMIT;

    SELECT 'ITEM ADICIONADO AO GRUPO' AS RETORNO;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ADCITENSCOMBO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ADCITENSCOMBO`(
	IN `P_IDCOMBO` INT,
	IN `P_IDPRODUTO` INT,
	IN `P_QUANTIDADE` DECIMAL(10,2)
)
BEGIN

DECLARE V_CUSTOPRODUTO DECIMAL(10,2);
DECLARE V_CUSTOFRACIONADO DECIMAL(10,2);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

SELECT CUSTOUNITARIO INTO V_CUSTOPRODUTO
FROM TAB_PRODUTOS WHERE IDPRODUTO = P_IDPRODUTO;

START TRANSACTION;

SET V_CUSTOFRACIONADO = V_CUSTOPRODUTO * P_QUANTIDADE;

INSERT INTO TAB_ITENSCOMBO(IDCOMBO,IDPRODUTO,QUANTIDADE,CUSTO)
VALUES (P_IDCOMBO, P_IDPRODUTO, P_QUANTIDADE, V_CUSTOFRACIONADO);

COMMIT;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ADCITENSVENDA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ADCITENSVENDA`(
	IN `P_IDVENDA` INT,
	IN `P_IDUSUARIO` INT,
	IN `P_IDPRODUTO` INT,
	IN `P_QUANTIDADE` INT,
	IN `P_TIPO` VARCHAR(10),
	IN `P_PRECOVENDA` DECIMAL(10,2)
)
TUDO: BEGIN 

DECLARE V_CONSUMO DECIMAL(10,2);
DECLARE V_SALDOUSO DECIMAL(10,2);
DECLARE V_CUSTOTOTALCOMBO DECIMAL(10,2);
DECLARE V_FIM INT DEFAULT 0;
DECLARE V_IDPRODUTOCOMBO INT;
DECLARE V_QTDEFRACIONADO DECIMAL(10,2);
DECLARE V_CUSTO DECIMAL(10,2);
DECLARE V_SUBTOTAL DECIMAL(10,2);
DECLARE V_QTDEESTOQUE INT;
DECLARE V_STATUSVENDA VARCHAR(20);
DECLARE V_STATUSITEM VARCHAR(30);
DECLARE V_ESTOQUEMINIMO INT;

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE C_ITENSCOMBO CURSOR FOR 
SELECT IDPRODUTO, QUANTIDADE
FROM TAB_ITENSCOMBO WHERE IDCOMBO = P_IDPRODUTO;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

DECLARE CONTINUE HANDLER FOR NOT FOUND 
SET V_FIM = 1;

SELECT STATUS INTO V_STATUSVENDA 
FROM TAB_VENDAS
WHERE IDVENDA = P_IDVENDA; 

SET V_SUBTOTAL = P_QUANTIDADE*P_PRECOVENDA; 

IF V_STATUSVENDA NOT LIKE 'ABERTO' THEN
SELECT 'VENDA CONCLUIDA OU CANCELADA' AS MSG;
LEAVE TUDO;
END IF;

START TRANSACTION ;

IF P_TIPO = 'COMBO' THEN

OPEN C_ITENSCOMBO;

LOOPCANCEL: WHILE V_FIM = 0 DO

    FETCH C_ITENSCOMBO
    INTO V_IDPRODUTOCOMBO, V_QTDEFRACIONADO;

    IF V_FIM = 1 THEN
        LEAVE LOOPCANCEL;
    END IF;

    SET V_CONSUMO = P_QUANTIDADE * V_QTDEFRACIONADO;

    SELECT ESTOQUEATUAL,
           IFNULL(QUANTIDADEUSO,0)
    INTO V_QTDEESTOQUE,
         V_SALDOUSO
    FROM TAB_PRODUTOS
    WHERE IDPRODUTO = V_IDPRODUTOCOMBO;
    
    IF V_QTDEESTOQUE <= 0 AND V_SALDOUSO < V_CONSUMO THEN

        SET V_STATUSITEM = 'PRODUTO SEM ESTOQUE';
        SELECT V_STATUSITEM AS MSG;
        ROLLBACK;
        LEAVE TUDO;

    END IF;

    SET V_STATUSITEM = 'ADICIONADO';

    SELECT CUSTOUNITARIO
    INTO V_CUSTO
    FROM TAB_PRODUTOS
    WHERE IDPRODUTO = V_IDPRODUTOCOMBO;
    
    IF V_SALDOUSO = 0 THEN

        UPDATE TAB_PRODUTOS
        SET ESTOQUEATUAL = ESTOQUEATUAL - 1,
            QUANTIDADEUSO = 1
        WHERE IDPRODUTO = V_IDPRODUTOCOMBO;

        SET V_QTDEESTOQUE = V_QTDEESTOQUE - 1;
        SET V_SALDOUSO = 1;

    END IF;
    
    IF V_SALDOUSO < V_CONSUMO THEN

        UPDATE TAB_PRODUTOS
        SET ESTOQUEATUAL = ESTOQUEATUAL - 1,
            QUANTIDADEUSO = QUANTIDADEUSO + 1
        WHERE IDPRODUTO = V_IDPRODUTOCOMBO;

        SET V_QTDEESTOQUE = V_QTDEESTOQUE - 1;
        SET V_SALDOUSO = V_SALDOUSO + 1;

    END IF;
    
    UPDATE TAB_PRODUTOS
    SET QUANTIDADEUSO = QUANTIDADEUSO - V_CONSUMO
    WHERE IDPRODUTO = V_IDPRODUTOCOMBO;

UPDATE TAB_VENDAS
SET VALOR_PENDENTE = VALOR_PENDENTE + v_SUBTOTAL
WHERE IDVENDA = p_IDVENDA;

    INSERT INTO TAB_MOVIMENTOESTOQUE(
        IDPRODUTO,
        TIPOMOVIMENTO,
        QUANTIDADE,
        CUSTOUNITARIO,
        ESTOQUEANTERIOR,
        ESTOQUEPOSTERIOR,
        IDUSUARIO,
        OBSERVACAO
    )
    VALUES (
        V_IDPRODUTOCOMBO,
        'SAIDA',
        -V_CONSUMO,
        V_CUSTO,
        V_QTDEESTOQUE,
        V_QTDEESTOQUE,
        P_IDUSUARIO,
        'ADICIONA ITEM PEDIDO'
    );

END WHILE LOOPCANCEL;

CLOSE C_ITENSCOMBO;

SELECT CUSTOUNITARIO INTO V_CUSTOTOTALCOMBO FROM TAB_PRODUTOS WHERE IDPRODUTO = P_IDPRODUTO;

INSERT INTO TAB_ITENSVENDA(IDVENDA, IDPRODUTO, QUANTIDADE,
PRECOUNITARIO, CUSTOUNITARIO, SUBTOTAL)
VALUES (P_IDVENDA, P_IDPRODUTO, P_QUANTIDADE, P_PRECOVENDA, V_CUSTOTOTALCOMBO, V_SUBTOTAL);

ELSE

SELECT ESTOQUEATUAL INTO V_QTDEESTOQUE
FROM TAB_PRODUTOS
WHERE IDPRODUTO = P_IDPRODUTO;

IF V_QTDEESTOQUE <=0 THEN
SET V_STATUSITEM ='PRODUTO SEM ESTOQUE';
SELECT V_STATUSITEM AS MSG;
ROLLBACK;
LEAVE TUDO;

ELSEIF P_QUANTIDADE > V_QTDEESTOQUE THEN
SET V_STATUSITEM = 'QUANTIDADE INDISPONIVEL';
SELECT V_STATUSITEM AS MSG, (V_QTDEESTOQUE - P_QUANTIDADE) AS FALTA;
ROLLBACK;
LEAVE TUDO;

ELSE
SET V_STATUSITEM = 'ADICIONADO' ;
END IF;

SELECT CUSTOUNITARIO INTO V_CUSTO FROM TAB_PRODUTOS
WHERE IDPRODUTO = P_IDPRODUTO; 

INSERT INTO TAB_ITENSVENDA(IDVENDA, IDPRODUTO, QUANTIDADE,
PRECOUNITARIO, CUSTOUNITARIO, SUBTOTAL)
VALUES (P_IDVENDA, P_IDPRODUTO, P_QUANTIDADE, P_PRECOVENDA, V_CUSTO, V_SUBTOTAL);

INSERT INTO TAB_MOVIMENTOESTOQUE(IDPRODUTO, TIPOMOVIMENTO, QUANTIDADE, CUSTOUNITARIO,
ESTOQUEANTERIOR, ESTOQUEPOSTERIOR, IDUSUARIO, OBSERVACAO )
VALUES (P_IDPRODUTO,'SAIDA',-P_QUANTIDADE, V_CUSTO,
 V_QTDEESTOQUE, (V_QTDEESTOQUE-P_QUANTIDADE), P_IDUSUARIO, 'ADICIONA ITEM PEDIDO');

UPDATE TAB_PRODUTOS
SET ESTOQUEATUAL = (ESTOQUEATUAL - P_QUANTIDADE),
ESTOQUEMINIMO = V_ESTOQUEMINIMO
WHERE IDPRODUTO = P_IDPRODUTO;

END IF;

INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO,IDUSUARIO)
VALUES ('SAIDA','ADICIONA ITEM',P_IDVENDA,P_IDUSUARIO);

COMMIT;

SELECT V_STATUSITEM AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ATIVARLICENCA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ATIVARLICENCA`(
    IN pChave VARCHAR(255)
)
BEGIN

    DECLARE vHashBanco VARCHAR(255);
    DECLARE vValidade DATE;

    SELECT VALOR
      INTO vHashBanco
      FROM TAB_CONFIG
     WHERE CHAVE = 'LICENCA_CHAVE';


    IF SHA2(pChave,256) = vHashBanco THEN


        SET vValidade = LAST_DAY(CURDATE());


        UPDATE TAB_CONFIG
           SET VALOR = 'ATIVA'
         WHERE CHAVE = 'LICENCA_STATUS';


        UPDATE TAB_CONFIG
           SET VALOR = DATE_FORMAT(vValidade,'%Y-%m-%d')
         WHERE CHAVE = 'LICENCA_VALIDADE';


        UPDATE TAB_CONFIG
           SET VALOR = DATE_FORMAT(NOW(),'%Y-%m-%d %H:%i:%s')
         WHERE CHAVE = 'LICENCA_ULTIMO_ACESSO';


        SELECT
            1 AS SUCESSO,
            'LICENCA_ATIVADA' AS RETORNO,
            DATE_FORMAT(vValidade,'%d/%m/%Y') AS VALIDADE,
            'Licença ativada com sucesso!' AS MSG;


    ELSE


        SELECT
            0 AS SUCESSO,
            'CHAVE_INVALIDA' AS RETORNO,
            NULL AS VALIDADE,
            'Chave de ativação inválida.' AS MSG;


    END IF;


END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_BAIXARPAGAMENTOS` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_BAIXARPAGAMENTOS`(
    IN P_IDUSUARIO INT,
    IN P_IDVENDA INT,
    IN P_FORMA VARCHAR(30),
    IN P_VALORRECEBIDO DECIMAL(10,2),
    IN P_VALORPAGO DECIMAL(10,2)
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE V_TOTALVENDA DECIMAL(10,2);
DECLARE V_TOTALPAGO DECIMAL(10,2);
DECLARE V_PENDENTE DECIMAL(10,2);

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    ROLLBACK;

    SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;

END;

IF NOT EXISTS(
    SELECT 1
    FROM TAB_VENDAS
    WHERE IDVENDA = P_IDVENDA
) THEN

    SELECT 'VENDA NÃO ENCONTRADA' AS RETORNO;

ELSEIF P_VALORPAGO <= 0 THEN

    SELECT 'VALOR INVÁLIDO' AS RETORNO;

ELSE

    SELECT IFNULL(SUM(SUBTOTAL),0)
    INTO V_TOTALVENDA
    FROM TAB_ITENSVENDA
    WHERE IDVENDA = P_IDVENDA;

    SELECT IFNULL(SUM(VALORPAGO),0)
    INTO V_TOTALPAGO
    FROM TAB_PAGAMENTOS
    WHERE IDVENDA = P_IDVENDA;

    SET V_PENDENTE = V_TOTALVENDA - V_TOTALPAGO;

    IF V_PENDENTE <= 0 THEN

        SELECT 'VENDA JÁ QUITADA' AS RETORNO;

    ELSE

        IF P_VALORPAGO > V_PENDENTE THEN
            SET P_VALORPAGO = V_PENDENTE;
        END IF;

        START TRANSACTION;

        INSERT INTO TAB_PAGAMENTOS(
            IDVENDA,
            FORMAPAGAMENTO,
            VALORRECEBIDO,
            VALORPAGO
        )
        VALUES(
            P_IDVENDA,
            P_FORMA,
            P_VALORRECEBIDO,
            P_VALORPAGO
        );

        UPDATE TAB_VENDAS
        SET VALOR_PENDENTE =
            GREATEST(0, IFNULL(VALOR_PENDENTE,0) - P_VALORPAGO)
        WHERE IDVENDA = P_IDVENDA;

        IF P_FORMA = 'DINHEIRO' THEN

            IF NOT EXISTS(
                SELECT 1
                FROM TAB_CAIXA
                WHERE STATUS = 'ABERTO'
            ) THEN

                SIGNAL SQLSTATE '45000'
                SET MESSAGE_TEXT = 'NÃO EXISTE CAIXA ABERTO';

            END IF;

            UPDATE TAB_CAIXA
            SET VALORCAIXA = IFNULL(VALORCAIXA,0) + P_VALORPAGO
            WHERE STATUS = 'ABERTO';

        END IF;

        INSERT INTO TAB_LOGEVENTOS(
            TIPOEVENTO,
            DESCRICAO,
            IDRELACIONADO,
            IDUSUARIO
        )
        VALUES(
            'PAGAMENTO',
            CONCAT(P_FORMA,' R$ ',FORMAT(P_VALORPAGO,2)),
            P_IDVENDA,
            P_IDUSUARIO
        );

        COMMIT;

        SELECT 'PAGAMENTO REGISTRADO' AS RETORNO;

    END IF;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CADASTRARCOMBO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CADASTRARCOMBO`(
	IN `P_NOME` VARCHAR(80),
	IN `P_MEDIDAVENDA` VARCHAR(5),
	IN `P_CATEGORIA` VARCHAR(40),
	IN `P_MARCA` VARCHAR(40),
	IN `P_PRECOVENDA` DECIMAL(10,2)
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

START TRANSACTION;

INSERT INTO TAB_PRODUTOS
(
    NOME,
    MEDIDAVENDA,
    CATEGORIA,
    TIPO,
    MARCA,
    STATUS,
    PRECOVENDA
)
VALUES
(
    P_NOME,
    P_MEDIDAVENDA,
    P_CATEGORIA,
    'COMBO',
    P_MARCA,
    'ATIVO',
    P_PRECOVENDA
);

COMMIT ;

SELECT LAST_INSERT_ID() AS IDCOMBO;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CADASTRARFORNECEDOR` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CADASTRARFORNECEDOR`(
	IN `P_NOME` VARCHAR(80),
	IN `P_CNPJ` VARCHAR(24),
	IN `P_CONTATO` VARCHAR(13),
	IN `P_ENDERECO` VARCHAR(100),
	IN `P_NUMERO` VARCHAR(10),
	IN `P_BAIRRO` VARCHAR (50),
	IN `P_STATUS` VARCHAR(15),
	IN `P_IDUSUARIO` INT
)
BEGIN

DECLARE V_NOVOFORNECEDOR INT;


DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

START TRANSACTION;

INSERT INTO TAB_FORNECEDORES(NOME,CNPJ,CONTATO,ENDERECO,NUMERO,BAIRRO,STATUS)
VALUES (P_NOME,P_CNPJ,P_CONTATO,P_ENDERECO,P_NUMERO,P_BAIRRO,P_STATUS);

SET V_NOVOFORNECEDOR = LAST_INSERT_ID(); 

INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO, IDUSUARIO)
VALUES ('CADASTRO', 'CADASTRO DE PRODUTO', V_NOVOFORNECEDOR, P_IDUSUARIO);

COMMIT;

SELECT 'CADASTRADO COM SUCESSO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CADASTRARGRUPO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CADASTRARGRUPO`(
    IN P_DESCRICAO VARCHAR(30)
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    ROLLBACK;

    SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;

END;


IF TRIM(P_DESCRICAO) = '' THEN

    SELECT 'DESCRIÇÃO OBRIGATÓRIA' AS RETORNO;

ELSEIF EXISTS(
    SELECT 1 
    FROM TAB_GRUPOS 
    WHERE DESCRICAO = P_DESCRICAO
) THEN

    SELECT 'GRUPO JÁ EXISTE' AS RETORNO;

ELSE

    START TRANSACTION;

    INSERT INTO TAB_GRUPOS(DESCRICAO)
    VALUES(P_DESCRICAO);

    COMMIT;

    SELECT 
        'GRUPO CADASTRADO' AS RETORNO,
        LAST_INSERT_ID() AS IDGRUPO;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CADASTRARPRODUTO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CADASTRARPRODUTO`(
	IN `P_CODIGOFORNECEDOR` INT,
	IN `P_NOME` VARCHAR(80),
	IN `P_CODIGOBARRASCX` VARCHAR(15),
	IN `P_CODIGOBARRAS` VARCHAR(15),
	IN `P_MEDIDACOMPRA` VARCHAR(5),
	IN `P_QUANTIDADECOMPRA` INT,
	IN `P_QUANTIDADEEMBALAGEM` INT,
	IN `P_MEDIDAVENDA` VARCHAR(5),
	IN `P_CATEGORIA` INT,
	IN `P_SETOR` VARCHAR(30),
	IN `P_MARCA` VARCHAR(40),
	IN `P_CUSTO` DECIMAL(10,2),
	IN `P_PRECO` DECIMAL(10,2),
	IN `P_STATUS` VARCHAR(10),
	IN `P_IDUSUARIO` INT
)
BEGIN

DECLARE V_IDNOVOPRODUTO INT;

   DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

    START TRANSACTION;

INSERT INTO TAB_PRODUTOS (
CODIGOFORNECEDOR,
        NOME,
CODIGOBARRASCX,
        CODIGOBARRAS,
        MEDIDACOMPRA,
        QUANTIDADECOMPRA,
        QUANTIDADEEMBALAGEM,
        MEDIDAVENDA,
        IDCATEGORIA,
        SETOR,
TIPO,
        MARCA,
        CUSTOUNITARIO,
CUSTOMEDIO,
PRECOVENDA,
        STATUS

    )
VALUES (

        P_CODIGOFORNECEDOR,
        P_NOME,
P_CODIGOBARRASCX,
        P_CODIGOBARRAS,
        P_MEDIDACOMPRA,
        P_QUANTIDADECOMPRA,
        P_QUANTIDADEEMBALAGEM,
        P_MEDIDAVENDA,
        P_CATEGORIA,
        P_SETOR,
        'UNIT',
        P_MARCA,
        P_CUSTO,
P_CUSTO,
P_PRECO,
        P_STATUS

    );

    SET V_IDNOVOPRODUTO = LAST_INSERT_ID();

    INSERT INTO TAB_LOGEVENTOS (
    TIPOEVENTO,
    DESCRICAO,
    IDRELACIONADO,
    IDUSUARIO
)
VALUES (
    'CADASTRO',
    'CADASTRO DE PRODUTO',
    V_IDNOVOPRODUTO,
    P_IDUSUARIO
);

COMMIT;

SELECT V_IDNOVOPRODUTO AS IDPRODUTO;

SELECT 'PRODUTO CADASTRADO COM SUCESSO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CADASTRARUSUARIO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CADASTRARUSUARIO`(
	IN `P_NOME` VARCHAR(40),
	IN `P_USUARIO` VARCHAR(20),
	IN `P_SENHA` VARCHAR(30)
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

START TRANSACTION;

INSERT INTO TAB_USUARIOS(NOME,USUARIO,SENHA)
VALUES (P_NOME,UPPER(P_USUARIO),SHA2(P_SENHA,256));

COMMIT;

SELECT 'USUARIO CADASTRADO COM SUCESSO' AS CONCLUIDO;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CALCULAPRECOVENDA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CALCULAPRECOVENDA`()
BEGIN


DECLARE V_IDPRODUTO INT;
DECLARE V_FIM INT DEFAULT 0;
DECLARE V_PRECOVENDA DECIMAL(10,2);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE C_CALCULAPRECOVENDA CURSOR FOR
	SELECT IDPRODUTO, 
			CUSTOMEDIO
	FROM TAB_PRODUTOS;
	
DECLARE CONTINUE HANDLER FOR NOT FOUND
SET V_FIM = 1;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

START TRANSACTION;


OPEN C_CALCULAPRECOVENDA;

LOOP_CANCEL: WHILE V_FIM = 0 DO

FETCH C_CALCULAPRECOVENDA 
INTO V_IDPRODUTO, V_PRECOVENDA;

IF V_FIM = 1 THEN
    LEAVE LOOP_CANCEL;
END IF;


UPDATE TAB_PRODUTOS
SET PRECOVENDA = V_PRECOVENDA*1.32
WHERE IDPRODUTO = V_IDPRODUTO;


COMMIT;

END WHILE;

CLOSE C_CALCULAPRECOVENDA;

SELECT 'PREÇO DE VENDA ATUALIZADO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CALCULARESTOQUEMINIMO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CALCULARESTOQUEMINIMO`()
TUDO: BEGIN


DECLARE V_IDPRODUTO INT;
DECLARE V_FIM INT DEFAULT 0;
DECLARE V_QTDEVENDIDA INT;
DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE C_CALCULAESTMINIMO CURSOR FOR
	SELECT IDPRODUTO, IFNULL(AVG(QUANTIDADE)*1.5,0)
	FROM TAB_ITENSVENDA IV
	JOIN TAB_VENDAS V ON V.IDVENDA = IV.IDVENDA 
	WHERE MONTH(V.DATAVENDA) = MONTH(CURRENT_DATE)
	GROUP BY IDPRODUTO;


DECLARE CONTINUE HANDLER FOR NOT FOUND
SET V_FIM = 1;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

START TRANSACTION;


OPEN C_CALCULAESTMINIMO;

LOOP_CANCEL: WHILE V_FIM = 0 DO

FETCH C_CALCULAESTMINIMO 
INTO V_IDPRODUTO, V_QTDEVENDIDA;

IF V_FIM = 1 THEN
    LEAVE LOOP_CANCEL;
END IF;


IF NOT EXISTS (SELECT 1 FROM tab_itensvenda IV
				JOIN tab_vendas V 
					ON V.IDVENDA = IV.IDVENDA
				WHERE IDPRODUTO = V_IDPRODUTO
				AND MONTH(V.DATAVENDA) = MONTH(CURRENT_DATE))
THEN
ROLLBACK;
LEAVE TUDO;
ELSE


UPDATE TAB_PRODUTOS
SET ESTOQUEMINIMO = V_QTDEVENDIDA
WHERE IDPRODUTO = V_IDPRODUTO;
END IF;

COMMIT;

END WHILE;

CLOSE C_CALCULAESTMINIMO;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CANCELARENTRADA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CANCELARENTRADA`(
	IN `P_IDCOMPRA` INT,
	IN `P_IDUSUARIO` INT
)
BEGIN

DECLARE V_FIM INT DEFAULT 0;
DECLARE V_IDPRODUTO INT;
DECLARE V_QUANTIDADE INT;
DECLARE V_STATUSCOMPRA VARCHAR(20);
DECLARE V_QTDEESTOQUE INT;

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE C_RETORNARESTOQUE CURSOR FOR
SELECT IDPRODUTO, QUANTIDADE 
FROM TAB_ITENSCOMPRA
WHERE IDCOMPRA = P_IDCOMPRA;

DECLARE CONTINUE HANDLER FOR NOT FOUND
SET V_FIM = 1;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;
SELECT STATUS INTO V_STATUSCOMPRA
FROM TAB_COMPRAS
WHERE IDCOMPRA = P_IDCOMPRA;

IF V_STATUSCOMPRA = 'CANCELADO' THEN

SELECT 'ENTRADA JA CANCELADA' AS ERRO;

ELSE 

START TRANSACTION;


OPEN C_RETORNARESTOQUE;

LOOP_CANCEL: WHILE TRUE DO

FETCH C_RETORNARESTOQUE INTO V_IDPRODUTO, V_QUANTIDADE;

IF V_FIM = 1 THEN 
    LEAVE LOOP_CANCEL;
END IF;

SELECT ESTOQUEATUAL 
INTO V_QTDEESTOQUE
FROM TAB_PRODUTOS
WHERE IDPRODUTO = V_IDPRODUTO;

UPDATE TAB_PRODUTOS 
SET ESTOQUEATUAL = ESTOQUEATUAL - V_QUANTIDADE
WHERE IDPRODUTO = V_IDPRODUTO;

INSERT INTO TAB_MOVIMENTOESTOQUE(
    IDPRODUTO, TIPOMOVIMENTO, QUANTIDADE, CUSTOUNITARIO,
    ESTOQUEANTERIOR, ESTOQUEPOSTERIOR, IDUSUARIO, OBSERVACAO
)
VALUES (
    V_IDPRODUTO,'SAIDA',V_QUANTIDADE, 0,
    V_QTDEESTOQUE, (V_QTDEESTOQUE - V_QUANTIDADE),
    P_IDUSUARIO, 'ENTRADA CANCELADA'
);

END WHILE;

CLOSE C_RETORNARESTOQUE;

INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO,IDUSUARIO)
VALUES ('SAIDA','ENTRADA CANCELADA',P_IDCOMPRA,P_IDUSUARIO);

DELETE FROM TAB_ITENSCOMPRA 
WHERE IDCOMPRA = P_IDCOMPRA;

UPDATE TAB_COMPRAS
SET STATUS = 'CANCELADO'
WHERE IDCOMPRA = P_IDCOMPRA;

COMMIT;

END IF;

SELECT 'ENTRADA CANCELADA COM SUCESSO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_CANCELARVENDA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_CANCELARVENDA`(
	IN `P_IDVENDA` INT,
	IN `P_IDUSUARIO` INT
)
TUDO: BEGIN

DECLARE V_FIM INT DEFAULT 0;
DECLARE V_IDPRODUTO INT;
DECLARE V_QUANTIDADE DECIMAL(10,3);
DECLARE V_STATUSVENDA VARCHAR(20);
DECLARE V_TIPO VARCHAR(10);
DECLARE V_CAIXA DECIMAL(10,2) DEFAULT 0;
DECLARE V_QTDEESTOQUE DECIMAL(10,3);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE C_ITENS CURSOR FOR
SELECT IDPRODUTO, QUANTIDADE
FROM TAB_ITENSVENDA
WHERE IDVENDA = P_IDVENDA;

DECLARE CONTINUE HANDLER FOR NOT FOUND
SET V_FIM = 1;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

SELECT STATUS INTO V_STATUSVENDA
FROM TAB_VENDAS
WHERE IDVENDA = P_IDVENDA;

IF V_STATUSVENDA = 'CANCELADO' THEN
	SELECT 'VENDA JA CANCELADA' AS MSG;
	LEAVE TUDO;
END IF;

SELECT IFNULL(SUM(VALORPAGO),0)
INTO V_CAIXA
FROM TAB_PAGAMENTOS
WHERE IDVENDA = P_IDVENDA
AND FORMAPAGAMENTO = 'DINHEIRO';

START TRANSACTION;

OPEN C_ITENS;

LOOP_CANCEL: WHILE V_FIM = 0 DO

	FETCH C_ITENS
	INTO V_IDPRODUTO, V_QUANTIDADE;

	IF V_FIM = 1 THEN
		LEAVE LOOP_CANCEL;
	END IF;

	SELECT TIPO INTO V_TIPO
	FROM TAB_PRODUTOS
	WHERE IDPRODUTO = V_IDPRODUTO;

	IF V_TIPO = 'COMBO' THEN

		UPDATE TAB_PRODUTOS P
		JOIN TAB_ITENSCOMBO IC
			ON IC.IDPRODUTO = P.IDPRODUTO
		SET P.ESTOQUEATUAL =
			P.ESTOQUEATUAL + (V_QUANTIDADE * IC.QUANTIDADE)
		WHERE IC.IDCOMBO = V_IDPRODUTO;

		INSERT INTO TAB_MOVIMENTOESTOQUE(
		IDPRODUTO,TIPOMOVIMENTO,QUANTIDADE,CUSTOUNITARIO,
		ESTOQUEANTERIOR,ESTOQUEPOSTERIOR,IDUSUARIO,OBSERVACAO)

		SELECT
		P.IDPRODUTO,
		'ENTRADA',
		(V_QUANTIDADE * IC.QUANTIDADE),
		0,
		(P.ESTOQUEATUAL - (V_QUANTIDADE * IC.QUANTIDADE)),
		P.ESTOQUEATUAL,
		P_IDUSUARIO,
		'VENDA CANCELADA'

		FROM TAB_PRODUTOS P
		JOIN TAB_ITENSCOMBO IC
			ON IC.IDPRODUTO = P.IDPRODUTO
		WHERE IC.IDCOMBO = V_IDPRODUTO;

	ELSE

		SELECT ESTOQUEATUAL
		INTO V_QTDEESTOQUE
		FROM TAB_PRODUTOS
		WHERE IDPRODUTO = V_IDPRODUTO;

		UPDATE TAB_PRODUTOS
		SET ESTOQUEATUAL = ESTOQUEATUAL + V_QUANTIDADE
		WHERE IDPRODUTO = V_IDPRODUTO;

		INSERT INTO TAB_MOVIMENTOESTOQUE(
		IDPRODUTO,TIPOMOVIMENTO,QUANTIDADE,CUSTOUNITARIO,
		ESTOQUEANTERIOR,ESTOQUEPOSTERIOR,IDUSUARIO,OBSERVACAO)
		VALUES(
		V_IDPRODUTO,
		'ENTRADA',
		V_QUANTIDADE,
		0,
		V_QTDEESTOQUE,
		(V_QTDEESTOQUE + V_QUANTIDADE),
		P_IDUSUARIO,
		'VENDA CANCELADA');

	END IF;

END WHILE;

CLOSE C_ITENS;

DELETE FROM TAB_ITENSVENDA
WHERE IDVENDA = P_IDVENDA;

UPDATE TAB_VENDAS
SET STATUS = 'CANCELADO'
WHERE IDVENDA = P_IDVENDA;

IF V_CAIXA > 0 THEN

    UPDATE TAB_CAIXA
    SET VALORCAIXA = VALORCAIXA - V_CAIXA
    WHERE STATUS = 'ABERTO';


    INSERT INTO TAB_LOGEVENTOS(
        TIPOEVENTO,
        DESCRICAO,
        IDRELACIONADO,
        IDUSUARIO
    )
    VALUES(
        'ESTORNO CAIXA',
        CONCAT('CANCELAMENTO VENDA - DEVOLVIDO DINHEIRO: ',V_CAIXA),
        P_IDVENDA,
        P_IDUSUARIO
    );

END IF;

INSERT INTO TAB_LOGEVENTOS(
TIPOEVENTO,DESCRICAO,IDRELACIONADO,IDUSUARIO)
VALUES(
'ENTRADA',
'VENDA CANCELADA',
P_IDVENDA,
P_IDUSUARIO);

COMMIT;

SELECT 'VENDA CANCELADA COM SUCESSO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_EDITARCOMBO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_EDITARCOMBO`(
IN P_IDCOMBO INT,
	IN `P_NOME` VARCHAR(80),
	IN `P_MEDIDAVENDA` VARCHAR(5),
	IN `P_CATEGORIA` VARCHAR(40),
	IN `P_MARCA` VARCHAR(40),
	IN `P_PRECOVENDA` DECIMAL(10,2)
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

START TRANSACTION;

UPDATE TAB_PRODUTOS
SET
    NOME = P_NOME,
    CATEGORIA = P_CATEGORIA,
    MEDIDAVENDA = P_MEDIDAVENDA,
    MARCA = P_MARCA,
    PRECOVENDA = P_PRECOVENDA
WHERE IDPRODUTO = P_IDCOMBO;

COMMIT;

SELECT 'Combo atualizado com sucesso.' AS RETORNO;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ENTRADAPRODUTOS` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ENTRADAPRODUTOS`(
	IN `P_IDCOMPRA` INT,
	IN `P_IDUSUARIO` INT,
	IN `P_IDPRODUTO` INT,
	IN `P_CODIGOFORNECEDOR` INT,
	IN `P_EAN` VARCHAR(15),
	IN `P_CODIGOBARRASCX` VARCHAR(15),
	IN `P_MEDIDACOMPRA` VARCHAR(5),
	IN `P_QUANTIDADECOMPRA` INT,
	IN `P_QUANTIDADEEMBALAGEM` INT,
	IN `P_QUANTIDADECOMPRADA` INT,
	IN `P_OBS` VARCHAR(80),
	IN `P_SUBTOTAL` DECIMAL(10,2)
)
BEGIN

    DECLARE V_IDPRODUTOLOCAL INT DEFAULT 0;
    DECLARE V_QUANTIDADEENTRADA INT;
    DECLARE V_CUSTOMEDIOFINAL DECIMAL(10,2);
    DECLARE V_VALORESTOQUEATUAL DECIMAL(10,2);
    DECLARE V_ESTOQUEATUAL INT;
    DECLARE V_NOVOESTOQUE INT;
DECLARE V_CUSTOUNITARIO DECIMAL(10,2);

    DECLARE V_SQLSTATE VARCHAR(5);
    DECLARE V_MSG TEXT;

    DECLARE EXIT HANDLER FOR SQLEXCEPTION
    BEGIN
        GET DIAGNOSTICS CONDITION 1
            V_SQLSTATE = RETURNED_SQLSTATE,
            V_MSG = MESSAGE_TEXT;

        ROLLBACK;

        SELECT CONCAT('ERRO [', V_SQLSTATE, '] ', V_MSG) AS MSG;
    END;
    
    IF P_IDPRODUTO IS NOT NULL AND P_IDPRODUTO > 0 THEN

        SET V_IDPRODUTOLOCAL = P_IDPRODUTO;

    ELSEIF P_EAN IS NOT NULL AND P_EAN <> '' THEN

        SELECT IDPRODUTO
        INTO V_IDPRODUTOLOCAL
        FROM TAB_PRODUTOS
        WHERE CODIGOBARRAS = P_EAN
        LIMIT 1;

    ELSEIF P_CODIGOFORNECEDOR IS NOT NULL AND P_CODIGOFORNECEDOR <> '' THEN

        SELECT IDPRODUTO
        INTO V_IDPRODUTOLOCAL
        FROM TAB_PRODUTOS
        WHERE CODIGOFORNECEDOR = P_CODIGOFORNECEDOR
        LIMIT 1;

    END IF;
    
    IF V_IDPRODUTOLOCAL IS NULL OR V_IDPRODUTOLOCAL = 0 THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'PRODUTO NAO ENCONTRADO';
    END IF;

    IF (SELECT STATUS FROM TAB_PRODUTOS WHERE IDPRODUTO = V_IDPRODUTOLOCAL) = 'INATIVO' THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'PRODUTO INATIVO';
    END IF;

    IF P_QUANTIDADECOMPRA <= 0 THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'QUANTIDADE INVALIDA';
    END IF;

    IF P_SUBTOTAL <= 0 THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'CUSTO INVALIDO';
    END IF;
    
UPDATE TAB_PRODUTOS
SET MEDIDACOMPRA = P_MEDIDACOMPRA,
QUANTIDADEEMBALAGEM = P_QUANTIDADEEMBALAGEM,
QUANTIDADECOMPRA = P_QUANTIDADECOMPRA
WHERE IDPRODUTO = P_IDPRODUTO;

    SELECT
        (IFNULL(QUANTIDADECOMPRA * QUANTIDADEEMBALAGEM, 0) * P_QUANTIDADECOMPRADA)
    INTO V_QUANTIDADEENTRADA
    FROM TAB_PRODUTOS
    WHERE IDPRODUTO = V_IDPRODUTOLOCAL;

SET V_CUSTOUNITARIO = P_SUBTOTAL/V_QUANTIDADEENTRADA;

    SELECT
        IFNULL(CUSTOMEDIO,0) * IFNULL(ESTOQUEATUAL,0),
        IFNULL(ESTOQUEATUAL,0)
    INTO V_VALORESTOQUEATUAL, V_ESTOQUEATUAL
    FROM TAB_PRODUTOS
    WHERE IDPRODUTO = V_IDPRODUTOLOCAL;

    SET V_NOVOESTOQUE = V_ESTOQUEATUAL + V_QUANTIDADEENTRADA;

    SET V_CUSTOMEDIOFINAL =
        (P_SUBTOTAL + V_VALORESTOQUEATUAL) / V_NOVOESTOQUE;

    START TRANSACTION;

    UPDATE TAB_PRODUTOS
    SET CUSTOUNITARIO = V_CUSTOUNITARIO,
        CUSTOMEDIO = V_CUSTOMEDIOFINAL,
        ESTOQUEATUAL = V_NOVOESTOQUE,
        CODIGOFORNECEDOR = P_CODIGOFORNECEDOR,
        CODIGOBARRASCX = P_CODIGOBARRASCX,
		CUSTOUNITARIO = V_CUSTOUNITARIO
    WHERE IDPRODUTO = V_IDPRODUTOLOCAL;

    INSERT INTO TAB_LOGEVENTOS
        (TIPOEVENTO, DESCRICAO, IDRELACIONADO, IDUSUARIO)
    VALUES
        ('ENTRADA', 'COMPRA DE MERCADORIA', P_IDCOMPRA, P_IDUSUARIO);

    INSERT INTO TAB_MOVIMENTOESTOQUE
        (IDPRODUTO, TIPOMOVIMENTO, QUANTIDADE, CUSTOUNITARIO,
         ESTOQUEANTERIOR, ESTOQUEPOSTERIOR, IDUSUARIO, OBSERVACAO)
    VALUES
        (V_IDPRODUTOLOCAL, 'ENTRADA', V_QUANTIDADEENTRADA, V_CUSTOUNITARIO,
         V_ESTOQUEATUAL, V_NOVOESTOQUE, P_IDUSUARIO, P_OBS);

    INSERT INTO TAB_ITENSCOMPRA
        (IDCOMPRA, IDPRODUTO, MEDIDACOMPRA, QUANTIDADE, CUSTOUNITARIO, SUBTOTAL)
    VALUES
        (P_IDCOMPRA, V_IDPRODUTOLOCAL, P_MEDIDACOMPRA,
         V_QUANTIDADEENTRADA, V_CUSTOUNITARIO, P_SUBTOTAL);

    COMMIT;

    SELECT
        'ADICIONADO' AS MSG,
        V_IDPRODUTOLOCAL AS IDPRODUTO;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_ESTORNAR` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_ESTORNAR`(
    IN P_IDUSUARIO INT,
    IN P_IDPAGAMENTO INT,
    IN P_VALORESTORNO DECIMAL(10,2),
    IN P_MOTIVO VARCHAR(90)
)
BEGIN

DECLARE V_IDVENDA INT;
DECLARE V_FORMAPAGAMENTO VARCHAR(30);
DECLARE V_VALORPAGAMENTO DECIMAL(10,2);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN
    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    ROLLBACK;

    SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;
END;

START TRANSACTION;

SELECT
    IDVENDA,
    FORMAPAGAMENTO,
    VALORPAGO
INTO
    V_IDVENDA,
    V_FORMAPAGAMENTO,
    V_VALORPAGAMENTO
FROM TAB_PAGAMENTOS
WHERE IDPAGAMENTO = P_IDPAGAMENTO;

IF V_VALORPAGAMENTO IS NULL THEN

    ROLLBACK;
    SELECT 'PAGAMENTO NÃO ENCONTRADO.' AS RETORNO;

ELSEIF P_VALORESTORNO > V_VALORPAGAMENTO THEN

    ROLLBACK;
    SELECT 'VALOR DE ESTORNO MAIOR QUE O VALOR PAGO.' AS RETORNO;

ELSE

    UPDATE TAB_PAGAMENTOS
       SET VALORPAGO = VALORPAGO - P_VALORESTORNO
     WHERE IDPAGAMENTO = P_IDPAGAMENTO;

    IF V_FORMAPAGAMENTO = 'DINHEIRO' THEN

        UPDATE TAB_CAIXA
           SET VALORCAIXA = VALORCAIXA - P_VALORESTORNO
         WHERE STATUS = 'ABERTO';

    END IF;

    INSERT INTO TAB_LOGEVENTOS
    (
        TIPOEVENTO,
        DESCRICAO,
        IDRELACIONADO,
        IDUSUARIO
    )
    VALUES
    (
        'ESTORNO',
        CONCAT(P_MOTIVO,' | Valor: R$ ',FORMAT(P_VALORESTORNO,2)),
        V_IDVENDA,
        P_IDUSUARIO
    );

    COMMIT;

    SELECT 'OK' AS RETORNO;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_EXCLUIRGRUPO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_EXCLUIRGRUPO`(
    IN P_IDGRUPO INT
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

GET DIAGNOSTICS CONDITION 1
V_SQLSTATE = RETURNED_SQLSTATE,
V_MSG = MESSAGE_TEXT;

ROLLBACK;

SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;

END;


START TRANSACTION;


DELETE FROM TAB_PRODUTOSGRUPO
WHERE IDGRUPO = P_IDGRUPO;


DELETE FROM TAB_GRUPOS
WHERE IDGRUPO = P_IDGRUPO;


COMMIT;


SELECT 'GRUPO EXCLUIDO' AS RETORNO;


END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_FECHARCAIXA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`%` PROCEDURE `PROC_FECHARCAIXA`(
	IN `P_IDCAIXA` INT,
	IN `P_IDUSUARIO` INT
)
BEGIN 

DECLARE V_FUNDO DECIMAL(10,2);
DECLARE V_SANGRIASTOTAL DECIMAL(10,2);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

SELECT FUNDO INTO V_FUNDO
FROM TAB_CAIXA 
WHERE IDCAIXA = P_IDCAIXA
AND STATUS = 'ABERTO';

SELECT SANGRIA INTO V_SANGRIASTOTAL
FROM TAB_CAIXA
WHERE IDCAIXA = P_IDCAIXA
AND STATUS = 'ABERTO';

IF NOT EXISTS ( SELECT 1 FROM TAB_CAIXA WHERE STATUS = 'ABERTO') THEN
SELECT 'NENHUM CAIXA ABERTO!' AS ERRO;

ELSEIF V_FUNDO >100 THEN
SELECT 'NECESSARIO REALIZAR SANGRIA ANTES DO FECHAMENTO' AS ERRO;

ELSE

START TRANSACTION;

UPDATE TAB_CAIXA 
SET STATUS = 'FECHADO',
	DATAFECHAMENTO = CURRENT_TIMESTAMP
WHERE IDCAIXA = P_IDCAIXA;

INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO,DESCRICAO,IDRELACIONADO,IDUSUARIO)
VALUES ('FECHAMENTO','FECHAMENTO DE CAIXA',P_IDCAIXA, P_IDUSUARIO);

COMMIT;

SELECT 'CAIXA FECHADO' AS MSG, 
	   V_SANGRIASTOTAL AS `VALOR TOTAL`,
	   V_FUNDO AS SOBRA;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_FECHARENTRADA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_FECHARENTRADA`(
	IN `P_IDCOMPRA` INT,
	IN `P_IDUSUARIO` INT,
	IN `P_VENCIMENTO` DATE,
	IN `P_STATUSPAGAMENTO` VARCHAR(20)
)
BEGIN

DECLARE V_VALORTOTAL DECIMAL(10,2);
DECLARE V_STATUSCOMPRA VARCHAR(20);
DECLARE V_FORNECEDOR INT;

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

SELECT STATUS INTO V_STATUSCOMPRA
FROM TAB_COMPRAS
WHERE IDCOMPRA = P_IDCOMPRA;

IF V_STATUSCOMPRA NOT LIKE 'ABERTO' THEN
SELECT 'ENTRADA CONCLUIDA OU CANCELADA' AS MSG;

ELSEIF NOT EXISTS (SELECT 1 FROM TAB_ITENSCOMPRA WHERE IDCOMPRA = P_IDCOMPRA) THEN
SELECT 'ENTRADA NÃO POSSUI PRODUTOS' AS MSG;

ELSE

SELECT SUM(SUBTOTAL) INTO V_VALORTOTAL
FROM TAB_ITENSCOMPRA WHERE IDCOMPRA = P_IDCOMPRA;



SELECT IDFORNECEDOR INTO V_FORNECEDOR
FROM TAB_COMPRAS
	WHERE IDCOMPRA = P_IDCOMPRA;

START TRANSACTION;


UPDATE TAB_COMPRAS
	SET STATUS = 'CONCLUIDO' 
WHERE IDCOMPRA = P_IDCOMPRA;


INSERT INTO TAB_CONTASAPAGAR(
		IDFORNECEDOR,
		DESCRICAO,
		VALOR,
		VENCIMENTO,
		STATUS
)
VALUES(
	   V_FORNECEDOR,
	   'COMPRA DE MERCADORIA',
	   V_VALORTOTAL,
	   P_VENCIMENTO,
	   P_STATUSPAGAMENTO
);


INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO,IDUSUARIO)
VALUES ('ENTRADA','CONCLUSÃO DE ENTRADA',P_IDCOMPRA,P_IDUSUARIO);

COMMIT;

SELECT 'ENTRADA PROCESSADA COM SUCESSO' AS MSG;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_FECHARVENDA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`%` PROCEDURE `PROC_FECHARVENDA`(
    IN P_IDVENDA INT,
    IN P_IDCAIXA INT,
    IN P_IDUSUARIO INT,
    IN P_DESCONTO DECIMAL(10,2)
)
BEGIN

DECLARE V_VALORTOTAL DECIMAL(10,2);
DECLARE V_VALORPAGO DECIMAL(10,2);
DECLARE V_STATUSVENDA VARCHAR(20);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN
    ROLLBACK;

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT('ERRO [', V_SQLSTATE, '] ', V_MSG) AS RETORNO;
END;

-- STATUS DA VENDA
SELECT STATUS INTO V_STATUSVENDA
FROM TAB_VENDAS
WHERE IDVENDA = P_IDVENDA;

IF V_STATUSVENDA NOT LIKE 'ABERTO' THEN

    SELECT 'VENDA CONCLUIDA OU CANCELADA' AS ERRO;

ELSEIF NOT EXISTS (
    SELECT 1 FROM TAB_ITENSVENDA WHERE IDVENDA = P_IDVENDA
) THEN

    SELECT 'VENDA NÃO POSSUI PRODUTOS' AS ERRO;

ELSE

    -- TOTAL DA VENDA
    SELECT IFNULL(SUM(IV.QUANTIDADE * IV.PRECOUNITARIO),0)
    INTO V_VALORTOTAL
    FROM TAB_ITENSVENDA IV
    WHERE IV.IDVENDA = P_IDVENDA;

    -- TOTAL PAGO
    SELECT IFNULL(SUM(P.VALORPAGO),0)
    INTO V_VALORPAGO
    FROM TAB_PAGAMENTOS P
    WHERE P.IDVENDA = P_IDVENDA;

    -- VALIDACAO FINAL
    IF V_VALORPAGO < (V_VALORTOTAL - P_DESCONTO) THEN

        SELECT 'VENDA AINDA NÃO FOI PAGA TOTALMENTE' AS ERRO;

    ELSE

        START TRANSACTION;

        UPDATE TAB_VENDAS
        SET STATUS = 'CONCLUIDO',
            IDCAIXA = P_IDCAIXA,
            VALORTOTAL = V_VALORTOTAL,
            DESCONTO = P_DESCONTO,
            VALORFINAL = (V_VALORTOTAL - P_DESCONTO)
        WHERE IDVENDA = P_IDVENDA;

        INSERT INTO TAB_LOGEVENTOS
        (TIPOEVENTO, DESCRICAO, IDRELACIONADO, IDUSUARIO)
        VALUES
        ('VENDA', 'VENDA CONCLUIDA', P_IDVENDA, P_IDUSUARIO);

        COMMIT;

        CALL PROC_CALCULARESTOQUEMINIMO;

        SELECT 'VENDA CONCLUIDA' AS MSG;

    END IF;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_LOGIN` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_LOGIN`(
	IN `P_USUARIO` VARCHAR(20),
	IN `P_SENHA` VARCHAR(30)
)
BEGIN
DECLARE V_LIBERAUSUARIO CHAR(1) DEFAULT 0;
DECLARE V_LIBERASENHA CHAR(1) DEFAULT 0;
DECLARE V_LIBERALOGIN CHAR(1) DEFAULT 0;
DECLARE V_USUARIO VARCHAR(20);
DECLARE V_ID INT;
DECLARE V_SENHA VARCHAR(64);
DECLARE V_NOME VARCHAR(40);
DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

SELECT IDUSUARIO,USUARIO,NOME INTO V_ID,V_USUARIO,V_NOME
FROM TAB_USUARIOS
WHERE USUARIO = P_USUARIO;

SELECT SENHA INTO V_SENHA
FROM TAB_USUARIOS
WHERE USUARIO = V_USUARIO 
AND SENHA = SHA2(P_SENHA,256);


IF V_USUARIO IS NOT NULL THEN

SET V_LIBERAUSUARIO = 1;

END IF;

IF V_SENHA IS NOT NULL THEN

SET V_LIBERASENHA = 1;

END IF;

IF V_LIBERAUSUARIO = 0 AND V_LIBERASENHA = 0 THEN
SELECT 'ERRO! USUARIO OU SENHA INCORRETOS' AS MSG,V_ID AS IDUSUARIO, NULL AS USUARIO, V_LIBERALOGIN AS STATUS;

ELSEIF V_LIBERAUSUARIO = 0 AND V_LIBERASENHA = 1 THEN 
SELECT 'ERRO! USUARIO OU SENHA INCORRETOS' AS MSG,V_ID AS IDUSUARIO, NULL AS USUARIO, V_LIBERALOGIN AS STATUS;

ELSEIF V_LIBERAUSUARIO = 1 AND V_LIBERASENHA = 0 THEN
SELECT 'ERRO! SENHA INCORRETA' AS MSG,V_ID AS IDUSUARIO,V_USUARIO AS USUARIO, V_LIBERALOGIN AS STATUS;

ELSE 

START TRANSACTION;


INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO, IDUSUARIO)
VALUES ('LOGIN', 'ACESSO AO SISTEMA', V_ID, V_ID);

COMMIT;

SET V_LIBERALOGIN = 1;
SELECT 'Logado com Sucesso!' AS MSG,V_ID AS IDUSUARIO, V_USUARIO AS USUARIO, V_NOME AS NOME, V_LIBERALOGIN AS STATUS;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_REGISTRARCOMPRA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_REGISTRARCOMPRA`(
	IN `P_IDFORNECEDOR` INT,
	IN `P_CHAVENF` VARCHAR(44),
	IN `P_NNF` INT,
	IN `P_IDUSUARIO` INT,
	IN `P_VALORTOTAL` DECIMAL(10,2)
)
BEGIN

DECLARE V_IDCOMPRAGERADO INT DEFAULT 0;

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

IF EXISTS( SELECT 1 FROM TAB_COMPRAS
			WHERE CHAVENF = P_CHAVENF) THEN
SELECT 'CHAVE NFE JA REGISTRADA' AS MSG;

ELSE

START TRANSACTION;

INSERT INTO TAB_COMPRAS(IDFORNECEDOR, CHAVENF, NUMERONF, VALORTOTAL,STATUS)
VALUES (P_IDFORNECEDOR, P_CHAVENF, P_NNF, P_VALORTOTAL,'ABERTO');

SET V_IDCOMPRAGERADO = LAST_INSERT_ID();

INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO,IDUSUARIO)
VALUES ('ENTRADA','COMPRA DE MERCADORIA',V_IDCOMPRAGERADO,P_IDUSUARIO);

SELECT V_IDCOMPRAGERADO AS IDCOMPRA;

COMMIT;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_REMOVEITEMCOMBO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_REMOVEITEMCOMBO`(
    IN P_IDITEMCOMBO INT
)
BEGIN

    DELETE FROM TAB_ITENSCOMBO
    WHERE IDITEMCOMBO = P_IDITEMCOMBO;

COMMIT;

    SELECT 'REMOVIDO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_REMOVEITEMGRUPO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_REMOVEITEMGRUPO`(
    IN P_IDGRUPO INT,
    IN P_IDPRODUTO INT
)
BEGIN

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN
    GET DIAGNOSTICS CONDITION 1 
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    ROLLBACK;
    SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;
END;


IF NOT EXISTS(
    SELECT 1 
    FROM TAB_PRODUTOSGRUPO
    WHERE IDGRUPO = P_IDGRUPO
    AND IDPRODUTO = P_IDPRODUTO
) THEN

    SELECT 'ITEM NÃO ENCONTRADO' AS RETORNO;

ELSE

    START TRANSACTION;

    DELETE FROM TAB_PRODUTOSGRUPO
    WHERE IDGRUPO = P_IDGRUPO
    AND IDPRODUTO = P_IDPRODUTO;

    COMMIT;

    SELECT 'ITEM REMOVIDO DO GRUPO' AS RETORNO;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_RESUMOCAIXA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`%` PROCEDURE `PROC_RESUMOCAIXA`()
BEGIN

DECLARE V_FATURADOPIX DECIMAL(10,2);
DECLARE V_FATURADODINHEIRO DECIMAL(10,2);
DECLARE V_FATURADOCARTAO DECIMAL(10,2);
DECLARE V_VALORCAIXA DECIMAL(10,2);
DECLARE V_SANGRIA DECIMAL(10,2);
DECLARE V_DATACAIXA DATE;
DECLARE V_FUNDO DECIMAL(10,2);

IF NOT EXISTS ( SELECT 1 FROM TAB_CAIXA WHERE STATUS = 'ABERTO') THEN
SELECT 'NENHUM CAIXA ABERTO!' AS ERRO;

ELSE

SELECT DATAABERTURA INTO V_DATACAIXA FROM TAB_CAIXA WHERE STATUS = 'ABERTO';

SELECT COALESCE(SUM(VALORTOTAL),0) INTO V_FATURADOPIX FROM TAB_VENDAS 
WHERE FORMAPAGAMENTO = 'PIX' AND DATAVENDA = V_DATACAIXA;

SELECT COALESCE(SUM(VALORTOTAL),0) INTO V_FATURADODINHEIRO FROM TAB_VENDAS 
WHERE FORMAPAGAMENTO = 'DINHEIRO' AND DATAVENDA = V_DATACAIXA;

SELECT COALESCE(SUM(VALORTOTAL),0) INTO V_FATURADOCARTAO FROM TAB_VENDAS 
WHERE FORMAPAGAMENTO = 'CARTAO' AND DATAVENDA = V_DATACAIXA;

SELECT VALORCAIXA,SANGRIA,FUNDO INTO V_VALORCAIXA,V_SANGRIA,V_FUNDO FROM TAB_CAIXA WHERE STATUS = 'ABERTO';

SELECT IFNULL(V_FATURADOPIX,0) AS PIX,
	  IFNULL(V_FATURADODINHEIRO,0) AS DINHEIRO,
	   IFNULL(V_FATURADOCARTAO,0) AS CARTAO,
	   IFNULL(V_VALORCAIXA,0) AS CAIXA,
	   IFNULL(V_SANGRIA,0) AS SANGRIA,
	   IFNULL(V_FUNDO,0) AS FUNDO,
	  IFNULL(V_FUNDO + V_VALORCAIXA,0) AS `DINHEIRO ESPERADO`;
	END IF;   
END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_RETIRARITENSCOMPRA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_RETIRARITENSCOMPRA`(
	IN `P_IDCOMPRA` INT,
	IN `P_IDUSUARIO` INT,
	IN `P_IDPRODUTO` INT,
	IN `P_QUANTIDADE` INT
)
BEGIN

DECLARE V_STATUSCOMPRA VARCHAR(20);
DECLARE V_ITENSCOMPRA INT;
DECLARE V_QTDEESTOQUE INT;

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;

IF NOT EXISTS (SELECT 1 FROM TAB_ITENSCOMPRA WHERE IDCOMPRA = P_IDCOMPRA) THEN
SELECT 'ENTRADA AINDA NÃO POSSUI PRODUTOS' AS ERRO;
ELSE

SELECT STATUS INTO V_STATUSCOMPRA 
FROM TAB_COMPRAS
WHERE IDCOMPRA = P_IDCOMPRA;

IF V_STATUSCOMPRA NOT LIKE 'ABERTO' THEN
SELECT 'ENTRADA CONCLUIDA OU CANCELADA' AS ERRO;
ELSE

START TRANSACTION;

SELECT ESTOQUEATUAL INTO V_QTDEESTOQUE
FROM TAB_PRODUTOS
WHERE IDPRODUTO = P_IDPRODUTO;

DELETE FROM TAB_ITENSCOMPRA
WHERE IDCOMPRA = P_IDCOMPRA
AND IDPRODUTO = P_IDPRODUTO;

UPDATE TAB_PRODUTOS 
SET ESTOQUEATUAL = (ESTOQUEATUAL - P_QUANTIDADE)
WHERE IDPRODUTO = P_IDPRODUTO;

INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO, DESCRICAO, IDRELACIONADO,IDUSUARIO)
VALUES ('ENTRADA','RETIRA ITEM DA ENTRADA',P_IDCOMPRA,P_IDUSUARIO);

INSERT INTO TAB_MOVIMENTOESTOQUE(IDPRODUTO, TIPOMOVIMENTO, QUANTIDADE, CUSTOUNITARIO,
ESTOQUEANTERIOR, ESTOQUEPOSTERIOR, IDUSUARIO)
VALUES (P_IDPRODUTO,'ENTRADA',P_QUANTIDADE, 0,
 V_QTDEESTOQUE, (V_QTDEESTOQUE-P_QUANTIDADE), P_IDUSUARIO);

COMMIT;

SELECT 'ITEM RETIRADO COM SUCESSO' AS MSG;

END IF;
END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_RETIRARITENSVENDA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_RETIRARITENSVENDA`(
	IN `P_IDVENDA` INT,
	IN `P_IDUSUARIO` INT,
	IN `P_IDPRODUTO` INT,
	IN `P_QUANTIDADE` INT
)
TUDO: BEGIN



DECLARE V_STATUSVENDA VARCHAR(20);
DECLARE V_QTDEESTOQUE DECIMAL(10,2);
DECLARE V_TIPO VARCHAR(10);

DECLARE V_FIM INT DEFAULT 0;
DECLARE V_IDPRODUTOCOMBO INT;
DECLARE V_QTDEFRACIONADO DECIMAL(10,2);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE C_ITENSCOMBO CURSOR FOR
SELECT IDPRODUTO, QUANTIDADE
FROM TAB_ITENSCOMBO
WHERE IDCOMBO = P_IDPRODUTO;

DECLARE CONTINUE HANDLER FOR NOT FOUND
SET V_FIM = 1;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;



IF NOT EXISTS (
SELECT 1
FROM TAB_ITENSVENDA
WHERE IDVENDA = P_IDVENDA
) THEN

SELECT 'VENDA NÃO POSSUI PRODUTOS' AS MSG;
LEAVE TUDO;

END IF;



SELECT STATUS
INTO V_STATUSVENDA
FROM TAB_VENDAS
WHERE IDVENDA = P_IDVENDA;



IF V_STATUSVENDA <> 'ABERTO' THEN

SELECT 'VENDA CONCLUIDA OU CANCELADA' AS MSG;
LEAVE TUDO;

END IF;



SELECT TIPO
INTO V_TIPO
FROM TAB_PRODUTOS
WHERE IDPRODUTO = P_IDPRODUTO;

START TRANSACTION;



IF V_TIPO = 'COMBO' THEN



OPEN C_ITENSCOMBO;

LOOP_COMBO: WHILE V_FIM = 0 DO

    

    FETCH C_ITENSCOMBO
    INTO V_IDPRODUTOCOMBO, V_QTDEFRACIONADO;

    IF V_FIM = 1 THEN
        LEAVE LOOP_COMBO;
    END IF;

    

    SELECT ESTOQUEATUAL
    INTO V_QTDEESTOQUE
    FROM TAB_PRODUTOS
    WHERE IDPRODUTO = V_IDPRODUTOCOMBO;

    

    UPDATE TAB_PRODUTOS
    SET ESTOQUEATUAL =
        ESTOQUEATUAL + (P_QUANTIDADE * V_QTDEFRACIONADO)
    WHERE IDPRODUTO = V_IDPRODUTOCOMBO;

    

    INSERT INTO TAB_MOVIMENTOESTOQUE(
        IDPRODUTO,
        TIPOMOVIMENTO,
        QUANTIDADE,
        CUSTOUNITARIO,
        ESTOQUEANTERIOR,
        ESTOQUEPOSTERIOR,
        IDUSUARIO
    )
    VALUES (
        V_IDPRODUTOCOMBO,
        'ENTRADA',
        (P_QUANTIDADE * V_QTDEFRACIONADO),
        0,
        V_QTDEESTOQUE,
        (V_QTDEESTOQUE + (P_QUANTIDADE * V_QTDEFRACIONADO)),
        P_IDUSUARIO
    );

END WHILE;



CLOSE C_ITENSCOMBO;

ELSE



SELECT ESTOQUEATUAL
INTO V_QTDEESTOQUE
FROM TAB_PRODUTOS
WHERE IDPRODUTO = P_IDPRODUTO;



UPDATE TAB_PRODUTOS
SET ESTOQUEATUAL = ESTOQUEATUAL + P_QUANTIDADE
WHERE IDPRODUTO = P_IDPRODUTO;



INSERT INTO TAB_MOVIMENTOESTOQUE(
    IDPRODUTO,
    TIPOMOVIMENTO,
    QUANTIDADE,
    CUSTOUNITARIO,
    ESTOQUEANTERIOR,
    ESTOQUEPOSTERIOR,
    IDUSUARIO
)
VALUES (
    P_IDPRODUTO,
    'ENTRADA',
    P_QUANTIDADE,
    0,
    V_QTDEESTOQUE,
    (V_QTDEESTOQUE + P_QUANTIDADE),
    P_IDUSUARIO
);

END IF;



DELETE FROM TAB_ITENSVENDA
WHERE IDVENDA = P_IDVENDA
AND IDPRODUTO = P_IDPRODUTO;



INSERT INTO TAB_LOGEVENTOS(
TIPOEVENTO,
DESCRICAO,
IDRELACIONADO,
IDUSUARIO
)
VALUES (
'ENTRADA',
'RETIRA ITEM DO PEDIDO',
P_IDVENDA,
P_IDUSUARIO
);

COMMIT;

SELECT 'ITEM REMOVIDO COM SUCESSO' AS MSG;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_SALVARCOMBO` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_SALVARCOMBO`(
	IN `P_IDPRODUTO` INT
)
BEGIN

DECLARE V_CUSTOTOTALCOMBO DECIMAL(10,2);

DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN

    GET DIAGNOSTICS CONDITION 1
        V_SQLSTATE = RETURNED_SQLSTATE,
        V_MSG = MESSAGE_TEXT;

    SELECT CONCAT(
        'ERRO [',
        V_SQLSTATE,
        '] ',
        V_MSG
    ) AS RETORNO;

    ROLLBACK;

END;
SELECT SUM(CUSTO) INTO V_CUSTOTOTALCOMBO
FROM TAB_ITENSCOMBO
WHERE IDCOMBO = P_IDPRODUTO;

START TRANSACTION;

UPDATE TAB_PRODUTOS
SET CUSTO = V_CUSTOTOTALCOMBO,
	 customedio = V_CUSTOTOTALCOMBO
WHERE IDPRODUTO = P_IDPRODUTO;

COMMIT;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_SANGRIACAIXA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_SANGRIACAIXA`(
    IN P_IDCAIXA INT,
    IN P_VALOR DECIMAL(10,2),
    IN P_IDUSUARIO INT
)
BEGIN

DECLARE V_VALORCAIXA DECIMAL(10,2);
DECLARE V_SQLSTATE VARCHAR(5);
DECLARE V_MSG TEXT;

DECLARE EXIT HANDLER FOR SQLEXCEPTION
BEGIN
    GET DIAGNOSTICS CONDITION 1 V_SQLSTATE = RETURNED_SQLSTATE, V_MSG = MESSAGE_TEXT;
    ROLLBACK;
    SELECT CONCAT('ERRO [',V_SQLSTATE,'] ',V_MSG) AS RETORNO;
END;


SELECT VALORCAIXA INTO V_VALORCAIXA
FROM TAB_CAIXA
WHERE IDCAIXA = P_IDCAIXA
AND STATUS = 'ABERTO';


IF V_VALORCAIXA IS NULL THEN

    SELECT 'NENHUM CAIXA ABERTO!' AS RETORNO;

ELSEIF P_VALOR <= 0 THEN

    SELECT 'VALOR DE SANGRIA INVÁLIDO!' AS RETORNO;

ELSEIF P_VALOR > V_VALORCAIXA THEN

    SELECT 'VALOR MAIOR QUE O CAIXA DISPONÍVEL!' AS RETORNO;

ELSE

    START TRANSACTION;

    UPDATE TAB_CAIXA
    SET VALORCAIXA = VALORCAIXA - P_VALOR,
        SANGRIA = SANGRIA + P_VALOR
    WHERE IDCAIXA = P_IDCAIXA;

    INSERT INTO TAB_LOGEVENTOS(TIPOEVENTO,DESCRICAO,IDRELACIONADO,IDUSUARIO)
    VALUES('SANGRIA',CONCAT('SANGRIA DE CAIXA - VALOR: ',P_VALOR),P_IDCAIXA,P_IDUSUARIO);

    COMMIT;

    SELECT 'SANGRIA REALIZADA' AS MSG, P_VALOR AS VALORRETIRADO;

END IF;

END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!50003 SET @saved_sql_mode       = @@sql_mode */ ;
/*!50003 SET sql_mode              = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ ;
/*!50003 DROP PROCEDURE IF EXISTS `PROC_VERIFICARLICENCA` */;
/*!50003 SET @saved_cs_client      = @@character_set_client */ ;
/*!50003 SET @saved_cs_results     = @@character_set_results */ ;
/*!50003 SET @saved_col_connection = @@collation_connection */ ;
/*!50003 SET character_set_client  = utf8mb4 */ ;
/*!50003 SET character_set_results = utf8mb4 */ ;
/*!50003 SET collation_connection  = utf8mb4_uca1400_ai_ci */ ;
DELIMITER ;;
CREATE DEFINER=`root`@`localhost` PROCEDURE `PROC_VERIFICARLICENCA`()
BEGIN

    DECLARE vStatus VARCHAR(20);
    DECLARE vValidade DATE;


    SELECT VALOR
      INTO vStatus
      FROM TAB_CONFIG
     WHERE CHAVE = 'LICENCA_STATUS';


    SELECT CASE
               WHEN VALOR = '' THEN NULL
               ELSE STR_TO_DATE(VALOR,'%Y-%m-%d')
           END
      INTO vValidade
      FROM TAB_CONFIG
     WHERE CHAVE = 'LICENCA_VALIDADE';



    -- Primeira ativação
    IF vStatus = 'PENDENTE' THEN


        SELECT
            0 AS SUCESSO,
            'LICENCA_PENDENTE' AS RETORNO,
            vValidade AS VALIDADE,
            'Este sistema ainda não foi ativado. Digite sua chave de ativação.' AS MSG;



    -- Bloqueio manual
    ELSEIF vStatus = 'BLOQUEADA' THEN


        SELECT
            0 AS SUCESSO,
            'LICENCA_BLOQUEADA' AS RETORNO,
            vValidade AS VALIDADE,
            'Esta licença está bloqueada. Entre em contato com o suporte.' AS MSG;



    -- Licença vencida
    ELSEIF vValidade IS NULL OR CURDATE() > vValidade THEN


        SELECT
            0 AS SUCESSO,
            'LICENCA_EXPIRADA' AS RETORNO,
            vValidade AS VALIDADE,
            CONCAT(
                'Sua licença expirou em ',
                DATE_FORMAT(vValidade,'%d/%m/%Y'),
                '. Digite uma nova chave de ativação.'
            ) AS MSG;



    -- Licença válida
    ELSE


        UPDATE TAB_CONFIG
           SET VALOR = DATE_FORMAT(NOW(),'%Y-%m-%d %H:%i:%s')
         WHERE CHAVE = 'LICENCA_ULTIMO_ACESSO';



        SELECT
            1 AS SUCESSO,
            'OK' AS RETORNO,
            vValidade AS VALIDADE,
            'Licença válida.' AS MSG;



    END IF;


END
;;
DELIMITER ;
/*!50003 SET sql_mode              = @saved_sql_mode */ ;
/*!50003 SET character_set_client  = @saved_cs_client */ ;
/*!50003 SET character_set_results = @saved_cs_results */ ;
/*!50003 SET collation_connection  = @saved_col_connection */ ;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*M!100616 SET NOTE_VERBOSITY=@OLD_NOTE_VERBOSITY */;

-- Dump completed on 2026-08-15 16:27:37
