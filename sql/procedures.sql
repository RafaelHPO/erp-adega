-- Procedures publicas sanitizadas. Sem dados reais.


/*M!999999\- enable the sandbox mode */ 
-- MariaDB dump 10.19-12.3.2-MariaDB, for Win64 (AMD64)
--
-- Host: <removido>    Database: <nome-do-banco>
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
CREATE PROCEDURE `PROC_ABRIRCAIXA`(
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


    SET V_IDCAIXA = LAST_INSERT_ID();


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
CREATE PROCEDURE `PROC_ABRIRVENDA`(
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


SET V_IDVENDA = LAST_INSERT_ID();


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
CREATE PROCEDURE `PROC_ACERTOESTOQUE`(
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
CREATE PROCEDURE `PROC_ADCITEMGRUPO`(
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
CREATE PROCEDURE `PROC_ADCITENSCOMBO`(
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
CREATE PROCEDURE `PROC_ADCITENSVENDA`(
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


END WHILE LOOPCANCEL;

CLOSE C_ITENSCOMBO;

SELECT CUSTOUNITARIO INTO V_CUSTOTOTALCOMBO FROM TAB_PRODUTOS WHERE IDPRODUTO = P_IDPRODUTO;


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



UPDATE TAB_PRODUTOS
SET ESTOQUEATUAL = (ESTOQUEATUAL - P_QUANTIDADE),
ESTOQUEMINIMO = V_ESTOQUEMINIMO
WHERE IDPRODUTO = P_IDPRODUTO;

END IF;


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
CREATE PROCEDURE `PROC_ATIVARLICENCA`(
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
CREATE PROCEDURE `PROC_BAIXARPAGAMENTOS`(
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
CREATE PROCEDURE `PROC_CADASTRARCOMBO`(
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
CREATE PROCEDURE `PROC_CADASTRARFORNECEDOR`(
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


SET V_NOVOFORNECEDOR = LAST_INSERT_ID(); 


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
CREATE PROCEDURE `PROC_CADASTRARGRUPO`(
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
CREATE PROCEDURE `PROC_CADASTRARPRODUTO`(
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


    SET V_IDNOVOPRODUTO = LAST_INSERT_ID();


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
CREATE PROCEDURE `PROC_CADASTRARUSUARIO`(
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
CREATE PROCEDURE `PROC_CALCULAPRECOVENDA`()
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
CREATE PROCEDURE `PROC_CALCULARESTOQUEMINIMO`()
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
CREATE PROCEDURE `PROC_CANCELARENTRADA`(
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


END WHILE;

CLOSE C_RETORNARESTOQUE;


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
CREATE PROCEDURE `PROC_CANCELARVENDA`(
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


	ELSE

		SELECT ESTOQUEATUAL
		INTO V_QTDEESTOQUE
		FROM TAB_PRODUTOS
		WHERE IDPRODUTO = V_IDPRODUTO;

		UPDATE TAB_PRODUTOS
		SET ESTOQUEATUAL = ESTOQUEATUAL + V_QUANTIDADE
		WHERE IDPRODUTO = V_IDPRODUTO;


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



END IF;


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
CREATE PROCEDURE `PROC_EDITARCOMBO`(
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
CREATE PROCEDURE `PROC_ENTRADAPRODUTOS`(
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
CREATE PROCEDURE `PROC_ESTORNAR`(
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
CREATE PROCEDURE `PROC_EXCLUIRGRUPO`(
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
CREATE PROCEDURE `PROC_FECHARCAIXA`(
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
CREATE PROCEDURE `PROC_FECHARENTRADA`(
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
CREATE PROCEDURE `PROC_FECHARVENDA`(
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
CREATE PROCEDURE `PROC_LOGIN`(
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
CREATE PROCEDURE `PROC_REGISTRARCOMPRA`(
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


SET V_IDCOMPRAGERADO = LAST_INSERT_ID();


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
CREATE PROCEDURE `PROC_REMOVEITEMCOMBO`(
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
CREATE PROCEDURE `PROC_REMOVEITEMGRUPO`(
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
CREATE PROCEDURE `PROC_RESUMOCAIXA`()
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
CREATE PROCEDURE `PROC_RETIRARITENSCOMPRA`(
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
CREATE PROCEDURE `PROC_RETIRARITENSVENDA`(
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




END IF;



DELETE FROM TAB_ITENSVENDA
WHERE IDVENDA = P_IDVENDA
AND IDPRODUTO = P_IDPRODUTO;




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
CREATE PROCEDURE `PROC_SALVARCOMBO`(
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
CREATE PROCEDURE `PROC_SANGRIACAIXA`(
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
CREATE PROCEDURE `PROC_VERIFICARLICENCA`()
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
