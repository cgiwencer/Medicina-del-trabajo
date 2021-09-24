/*
MySQL Backup
Source Server Version: 5.1.30
Source Database: medtrabajo
Date: 22/03/2021 00:29:55
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
--  Table structure for `mensual`
-- ----------------------------
DROP TABLE IF EXISTS `mensual`;
CREATE TABLE `mensual` (
  `MeEmPrT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeEMPr` int(3) DEFAULT NULL,
  `MeEmPOT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeEMPO` int(3) DEFAULT NULL,
  `MeLaPrT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeLaPr` int(3) DEFAULT NULL,
  `MeLaPOT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeLaPO` int(3) DEFAULT NULL,
  `MeLaPrST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeLaPrS` int(3) DEFAULT NULL,
  `MeLaPOST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeLaPOS` int(3) DEFAULT NULL,
  `MeRxPrT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRxPr` int(3) DEFAULT NULL,
  `MeRxPOT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRxPO` int(3) DEFAULT NULL,
  `MeRxPrST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRxPrS` int(3) DEFAULT NULL,
  `MeRxPOST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRxPOS` int(3) DEFAULT NULL,
  `MeRLaPrT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRLaPr` int(3) DEFAULT NULL,
  `MeRLaPOT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRLaPO` int(3) DEFAULT NULL,
  `MeRLaPrST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRLaPrS` int(3) DEFAULT NULL,
  `MeRLaPOST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRLaPOS` int(3) DEFAULT NULL,
  `MeRRxPrT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRRxPr` int(3) DEFAULT NULL,
  `MeRRxPOT` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRRxPO` int(3) DEFAULT NULL,
  `MeRRxPrST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRRxPrS` int(3) DEFAULT NULL,
  `MeRRxPOST` varchar(200) COLLATE utf8_spanish_ci DEFAULT NULL,
  `MeRRxPOS` int(3) DEFAULT NULL,
  `MeMes` varchar(20) COLLATE utf8_spanish_ci DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_spanish_ci;

-- ----------------------------
--  Records 
-- ----------------------------
